import test from "node:test";
import assert from "node:assert/strict";
import {
  buildMessagesUrl,
  buildTokenRequest,
  messageToTask,
  messagesToTasks,
  parseTokenResponse,
  poll,
  stateKeyForChannel,
  stripHtml,
  triggers,
} from "../index.js";

test("triggers map registers poll only", () => {
  assert.deepEqual(Object.keys(triggers), ["poll"]);
  assert.equal(typeof triggers.poll, "function");
});

// ---------------------------------------------------------------------------
// Auth
// ---------------------------------------------------------------------------

test("buildTokenRequest targets the tenant token endpoint with client credentials", () => {
  const { url, options } = buildTokenRequest({ tenantId: "tid", clientId: "cid", clientSecret: "sekret" });
  assert.equal(url, "https://login.microsoftonline.com/tid/oauth2/v2.0/token");
  assert.equal(options.method, "POST");
  const form = new URLSearchParams(options.body);
  assert.equal(form.get("client_id"), "cid");
  assert.equal(form.get("client_secret"), "sekret");
  assert.equal(form.get("grant_type"), "client_credentials");
  assert.equal(form.get("scope"), "https://graph.microsoft.com/.default");
});

test("parseTokenResponse returns the token and refuses failures loudly", () => {
  assert.equal(parseTokenResponse(200, { access_token: "eyJ…" }), "eyJ…");
  assert.throws(() => parseTokenResponse(401, { error: "invalid_client", error_description: "bad secret" }), /bad secret/);
  assert.throws(() => parseTokenResponse(200, {}), /token request failed/);
});

// ---------------------------------------------------------------------------
// URL and text helpers
// ---------------------------------------------------------------------------

test("buildMessagesUrl continues a cursor verbatim and otherwise starts a delta query", () => {
  const cursor = "https://graph.microsoft.com/v1.0/teams/t/channels/c/messages/delta?$deltatoken=abc";
  assert.equal(buildMessagesUrl("t", "c", cursor), cursor);
  const fresh = buildMessagesUrl("team 1", "19:ch@thread.tacv2", undefined);
  assert.equal(fresh, "https://graph.microsoft.com/v1.0/teams/team%201/channels/19%3Ach%40thread.tacv2/messages/delta?$top=50");
});

test("stripHtml flattens Teams markup to readable text", () => {
  assert.equal(stripHtml("<p>Hello <b>world</b></p><p>Second&nbsp;line &amp; more</p>"), "Hello world\nSecond line & more");
  assert.equal(stripHtml(""), "");
  assert.equal(stripHtml(undefined), "");
});

// ---------------------------------------------------------------------------
// The transform + dedup
// ---------------------------------------------------------------------------

const human = (id, text, extra = {}) => ({
  id,
  messageType: "message",
  createdDateTime: "2026-10-24T10:00:00Z",
  from: { user: { displayName: "Amina" } },
  body: { contentType: "html", content: `<p>${text}</p>` },
  webUrl: "https://teams.microsoft.com/l/message/x",
  ...extra,
});

test("messageToTask builds a task from a human message and skips everything else", () => {
  const task = messageToTask("C1", human("m1", "Ship the release"));
  assert.equal(task.title, "Ship the release");
  assert.ok(task.description.includes("Amina"));
  assert.ok(task.description.includes("Open in Teams"));

  assert.equal(messageToTask("C1", human("m2", "x", { messageType: "systemEventMessage" })), null, "system events");
  assert.equal(messageToTask("C1", human("m3", "x", { deletedDateTime: "2026-10-24T10:01:00Z" })), null, "deleted messages");
  assert.equal(messageToTask("C1", human("m4", "x", { from: { application: { displayName: "Bot" } } })), null, "app posts");
  assert.equal(messageToTask("C1", human("m5", "   ")), null, "empty bodies");
});

test("messagesToTasks dedups by message id", () => {
  const seen = new Set(["m1"]);
  const tasks = messagesToTasks("C1", [human("m1", "dup"), human("m2", "new"), human("m2", "dup again")], seen);
  assert.deepEqual(tasks.map(t => t.title), ["new"]);
  assert.ok(seen.has("m2"));
});

test("long first lines are truncated to a bounded title", () => {
  const task = messageToTask("C1", human("m1", "x".repeat(200) + "<br>rest"));
  assert.ok(task.title.length <= 80);
  assert.ok(task.title.endsWith("…"));
});

// ---------------------------------------------------------------------------
// Poll end to end with a stub fetch
// ---------------------------------------------------------------------------

function stubGraph(pagesByUrl) {
  const calls = [];
  return {
    calls,
    fetch: async (url, options) => {
      calls.push({ url, options });
      if (url.includes("login.microsoftonline.com")) {
        return { status: 200, ok: true, json: async () => ({ access_token: "eyJ-test" }) };
      }
      const page = pagesByUrl[url];
      if (!page) return { status: 404, ok: false, json: async () => ({ error: { message: "not found" } }) };
      return { status: 200, ok: true, json: async () => page };
    },
  };
}

const secrets = {
  MSGRAPH_TENANT_ID: "tid",
  MSGRAPH_CLIENT_ID: "cid",
  MSGRAPH_CLIENT_SECRET: "sekret",
  MSGRAPH_TEAM_ID: "team-1",
  MSGRAPH_CHANNEL_IDS: "chan-1",
};

const startUrl = buildMessagesUrl("team-1", "chan-1", undefined);

test("poll walks pages to the deltaLink, creates tasks, and persists the cursor", async () => {
  const page2 = "https://graph.microsoft.com/v1.0/next/2";
  const deltaLink = "https://graph.microsoft.com/v1.0/delta/abc";
  const { calls, fetch } = stubGraph({
    [startUrl]: { value: [human("m1", "first"), human("m2", "second")], "@odata.nextLink": page2 },
    [page2]: { value: [human("m2", "second"), human("m3", "third")], "@odata.deltaLink": deltaLink },
  });

  const result = await poll({ secrets, state: {}, fetch });

  assert.deepEqual(result.operations.map(op => op.value.title), ["first", "second", "third"], "id-dedup spans pages");
  assert.equal(result.state[stateKeyForChannel("team-1", "chan-1")], deltaLink);
  const graphCalls = calls.filter(c => !c.url.includes("login.microsoftonline.com"));
  assert.equal(graphCalls.length, 2);
  assert.equal(graphCalls[0].options.headers.authorization, "Bearer eyJ-test");
});

test("poll resumes from the stored deltaLink verbatim", async () => {
  const stored = "https://graph.microsoft.com/v1.0/delta/stored";
  const { calls, fetch } = stubGraph({
    [stored]: { value: [human("m9", "fresh")], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/delta/newer" },
  });

  const result = await poll({ secrets, state: { [stateKeyForChannel("team-1", "chan-1")]: stored }, fetch });

  assert.equal(calls.filter(c => !c.url.includes("login.microsoftonline.com"))[0].url, stored);
  assert.deepEqual(result.operations.map(op => op.value.title), ["fresh"]);
  assert.equal(result.state[stateKeyForChannel("team-1", "chan-1")], "https://graph.microsoft.com/v1.0/delta/newer");
});

test("a Graph failure aborts the run so the cursor is not committed", async () => {
  const { fetch } = stubGraph({}); // every Graph URL 404s
  await assert.rejects(() => poll({ secrets, state: {}, fetch }), /Microsoft Graph messages failed/);
});

test("missing configuration fails loudly, missing channels is a no-op", async () => {
  await assert.rejects(() => poll({ secrets: {}, state: {}, fetch: async () => { throw new Error("no"); } }), /MSGRAPH_TENANT_ID/);
  const result = await poll({ secrets: { ...secrets, MSGRAPH_CHANNEL_IDS: " , " }, state: {}, fetch: async () => { throw new Error("must not be called"); } });
  assert.deepEqual(result.operations, []);
});
