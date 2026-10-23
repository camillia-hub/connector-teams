# Teams Connector

Creates Camillia tasks from Microsoft Teams channel messages — a conversation
in a channel becomes tracked work. Poll-only: every five minutes the
connector reads each configured channel through **Microsoft Graph delta
queries** and creates one task per new human message.

## Grants requested

- `create-task` — the connector only ever returns `create_task` operations. It
  never comments on or edits existing tasks.

## Secrets

All five are required. Values are stored encrypted and never returned by the
API after they are set.

| Key | What |
| --- | --- |
| `MSGRAPH_TENANT_ID` | Microsoft Entra tenant id. |
| `MSGRAPH_CLIENT_ID` | Application (client) id of the app registration. |
| `MSGRAPH_CLIENT_SECRET` | Client secret of that app registration. |
| `MSGRAPH_TEAM_ID` | The team (group) id to read. |
| `MSGRAPH_CHANNEL_IDS` | Comma-separated channel ids, e.g. `19:abc…@thread.tacv2`. |

The app registration needs the **application** permission
`ChannelMessage.Read.All` (admin-consented); the connector authenticates with
client credentials (`oauth2/v2.0/token`, scope
`https://graph.microsoft.com/.default`).

## How dedup works

For each configured channel the connector runs a
`/teams/{team}/channels/{channel}/messages/delta` query. Graph's delta query
is itself a cursor: the `@odata.deltaLink` a run ends with is persisted in the
connector state (`state["graph:<team>:<channel>"] = <deltaLink>`), and the next
run continues exactly there — Graph only ever yields each change once. Within
a run, message ids dedup paging overlaps. The state is committed only when the
reported operations are accepted, so a Graph outage retries the same window
next time.

Messages that never become tasks (and never move the cursor): system event
messages, deleted messages, posts from applications (no `from.user`), and
empty bodies.

## Behaviour details

- Up to 3 delta pages per channel per run, capped at 20 tasks per run.
- Teams bodies are HTML; the task description keeps a readable plain-text
  quote with author, channel, timestamp and a deep link back to Teams.
- A task title is the message's first line, truncated to 80 characters.
- Tasks land in the install's first scoped workspace (scope the install to
  the workspace you want Teams to feed).
- When no agent of your organization is online, a poll occurrence is consumed
  and the next one retries — polls fire slightly late, never twice.

## Development

Requires Node.js 22 or newer.

```sh
node scripts/validate.mjs   # manifest + layout checks
npm test                    # unit tests (node --test), including the dedup logic
```
