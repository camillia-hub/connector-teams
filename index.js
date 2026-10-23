// connector-teams
// Feeds Microsoft Teams channel conversations into Camillia: every new human
// channel message becomes a task. Poll-only: every five minutes the connector
// calls Microsoft Graph with an app-only token (client credentials) and walks
// each configured channel's messages with a delta query. The deltaLink Graph
// returns is the dedup cursor, persisted per channel in the connector state —
// Graph itself guarantees a deltaLink only ever yields each change once.
// Grant requested: create-task (the connector only ever returns create_task
// operations — it never touches existing tasks).

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const MAX_TITLE_LENGTH = 80;
const MAX_OPERATIONS_PER_RUN = 20;
const MAX_PAGES_PER_CHANNEL = 3;

// ---------------------------------------------------------------------------
// Auth (client credentials)
// ---------------------------------------------------------------------------

export function buildTokenRequest({ tenantId, clientId, clientSecret }) {
  return {
    url: `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`,
    options: {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded", accept: "application/json" },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      }).toString(),
    },
  };
}

export function parseTokenResponse(status, data) {
  if (status !== 200 || !data || typeof data.access_token !== "string" || data.access_token === "") {
    const detail = data?.error_description ?? data?.error ?? `HTTP ${status}`;
    throw new Error(`Microsoft Graph token request failed: ${detail}`);
  }
  return data.access_token;
}

// ---------------------------------------------------------------------------
// Reading channel messages (delta queries)
// ---------------------------------------------------------------------------

/// With a cursor, Graph continues exactly there (the cursor is a deltaLink or
/// nextLink); without one, a fresh delta query starts at the current tail.
export function buildMessagesUrl(teamId, channelId, cursor) {
  if (cursor) return cursor;
  return `${GRAPH_ROOT}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/delta?$top=50`;
}

/// Teams bodies are HTML; tasks keep readable plain text.
export function stripHtml(html) {
  return (typeof html === "string" ? html : "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/[ \t]+\n/g, "\n")
    .trim();
}

/// One Teams message as a create_task operation value; null when it is not
/// human conversation (system events, deletions, app posts, empty bodies).
export function messageToTask(channelId, message) {
  if (!message || message.messageType !== "message") return null;
  if (message.deletedDateTime) return null;
  const author = message.from?.user?.displayName;
  if (!author) return null;
  const text = stripHtml(message.body?.content);
  if (!text) return null;
  const firstLine = text.split("\n", 1)[0];
  const title = firstLine.length > MAX_TITLE_LENGTH ? firstLine.slice(0, MAX_TITLE_LENGTH - 1) + "…" : firstLine;
  const description = [
    "Reported from Microsoft Teams (channel poll).",
    "",
    "> " + text.split("\n").join("\n> "),
    "",
    `- Author: ${author}`,
    `- Channel: ${channelId}`,
    message.createdDateTime ? `- Posted: ${message.createdDateTime}` : null,
    message.webUrl ? `- Open in Teams: ${message.webUrl}` : null,
  ].filter(line => line !== null).join("\n");
  return { title, description };
}

/// The new tasks of one delta page; ids dedup within and across pages of a run.
export function messagesToTasks(channelId, messages, seenIds) {
  const tasks = [];
  for (const message of Array.isArray(messages) ? messages : []) {
    const id = message?.id;
    if (id && seenIds.has(id)) continue;
    const task = messageToTask(channelId, message);
    if (task) {
      tasks.push(task);
      if (id) seenIds.add(id);
    }
  }
  return tasks;
}

export function stateKeyForChannel(teamId, channelId) {
  return `graph:${teamId}:${channelId}`;
}

// ---------------------------------------------------------------------------
// The poll trigger
// ---------------------------------------------------------------------------

function requireSecret(secrets, key) {
  const value = (secrets?.[key] ?? "").trim();
  if (!value) throw new Error(`${key} is not configured on this install.`);
  return value;
}

/// Reads every configured channel from its persisted deltaLink, creates one
/// task per new human message, and returns the advanced deltaLinks. The state
/// is committed only when the operations are accepted, so a Graph outage
/// retries the same window next time.
export async function poll({ secrets, state, fetch }) {
  const tenantId = requireSecret(secrets, "MSGRAPH_TENANT_ID");
  const clientId = requireSecret(secrets, "MSGRAPH_CLIENT_ID");
  const clientSecret = requireSecret(secrets, "MSGRAPH_CLIENT_SECRET");
  const teamId = requireSecret(secrets, "MSGRAPH_TEAM_ID");
  const channels = (secrets?.MSGRAPH_CHANNEL_IDS ?? "").split(",").map(id => id.trim()).filter(Boolean);
  const newState = { ...(state ?? {}) };
  if (channels.length === 0) return { operations: [], state: newState };

  const fetchFn = fetch ?? globalThis.fetch;
  const tokenRequest = buildTokenRequest({ tenantId, clientId, clientSecret });
  const tokenResponse = await fetchFn(tokenRequest.url, tokenRequest.options);
  const token = parseTokenResponse(tokenResponse.status, await tokenResponse.json());
  const graphHeaders = { authorization: `Bearer ${token}`, accept: "application/json" };

  const operations = [];
  for (const channelId of channels) {
    const stateKey = stateKeyForChannel(teamId, channelId);
    let url = buildMessagesUrl(teamId, channelId, state?.[stateKey]);
    const seenIds = new Set();
    for (let page = 0; page < MAX_PAGES_PER_CHANNEL && url; page++) {
      const response = await fetchFn(url, { headers: graphHeaders });
      const data = await response.json().catch(() => null);
      if (!response.ok) {
        const detail = data?.error?.message ?? `HTTP ${response.status}`;
        throw new Error(`Microsoft Graph messages failed for ${channelId}: ${detail}`);
      }
      for (const task of messagesToTasks(channelId, data?.value, seenIds)) {
        operations.push({ type: "create_task", value: task });
      }
      // A deltaLink ends the channel (it is the cursor); a nextLink continues paging.
      if (data?.["@odata.deltaLink"]) {
        newState[stateKey] = data["@odata.deltaLink"];
        url = null;
      } else {
        url = data?.["@odata.nextLink"] ?? null;
      }
    }
  }
  return { operations: operations.slice(0, MAX_OPERATIONS_PER_RUN), state: newState };
}

export const triggers = { poll };
