# n8n-nodes-agents-sdk

[![npm version](https://img.shields.io/npm/v/n8n-nodes-agents-sdk.svg)](https://www.npmjs.com/package/n8n-nodes-agents-sdk)
[![npm downloads](https://img.shields.io/npm/dm/n8n-nodes-agents-sdk.svg)](https://www.npmjs.com/package/n8n-nodes-agents-sdk)
[![license](https://img.shields.io/npm/l/n8n-nodes-agents-sdk.svg)](./LICENSE)
[![node](https://img.shields.io/node/v/n8n-nodes-agents-sdk.svg)](https://nodejs.org/)

Build **Microsoft Teams**, **M365 Copilot**, **WebChat**, and **Direct Line** bots with [n8n](https://n8n.io) as the orchestration layer. Exposes the [Microsoft 365 Agents SDK](https://learn.microsoft.com/microsoft-365/agents-sdk/) as first-class n8n nodes — JWT validation handled, MSAL tokens cached, no raw HTTP, no hand-assembled Activity JSON.

Built on `@microsoft/agents-hosting` (the successor to the archived `botbuilder` SDK).

---

## Why

The typical Bot Framework integration in n8n today is a Webhook → Code → HTTP Request chain with OAuth2 credentials, hand-built URLs, and no JWT validation (anyone who finds the webhook URL can impersonate Azure Bot Service). This package replaces all of that with four nodes that speak the Activity protocol natively.

| Without this package | With this package |
|---|---|
| Webhook node → Code to parse Activity | `M365 Agent Trigger` — parsed envelope |
| Manual JWT validation (often skipped) | Validated on every POST against Microsoft JWKS |
| OAuth2 credential + scope + token refresh plumbing | `M365 Agent API` credential — App ID + secret + tenant |
| HTTP Request with hand-built URLs (`/v3/conversations/.../activities/...`) | `M365 Send Activity` — pick an operation from a dropdown |
| Manual `;messageid=` thread suffix for Teams | `operation: replyInThread` |

## Nodes

| Node | Purpose |
|---|---|
| **M365 Agent Trigger** | Receives POSTs from Azure Bot Service. Validates JWT. Parses incoming Activity into an envelope (`conversationReference`, `activity`, `parsed`, `raw`). Responds to GET for Azure's endpoint-verification ping. |
| **M365 Text Message** | Builds a text Activity. Preserves envelope state. |
| **M365 Card Template** | Renders an Adaptive Card template with `${field}` placeholders via `adaptivecards-templating`. Preserves envelope state. |
| **M365 Send Activity** | Sends the envelope to Azure Bot Service. Five operations: `reply`, `proactive`, `update`, `delete`, `replyInThread`. Uses MSAL internally for token acquisition and caching. |

All four nodes are marked `usableAsTool: true`, so an n8n AI Agent can use them directly.

## Credential

**M365 Agent API** — the App ID, client secret, and tenant of your Azure Bot registration. Three app types supported:

- **SingleTenant** — production default; token scoped to your Entra tenant.
- **MultiTenant** — deprecated by Microsoft after 2025-07-31 for new bots; kept for legacy.
- **UserAssignedMsi** — Azure managed identity.

## Requirements

- **Node.js** 20 or later
- **n8n** self-hosted (Cloud not supported — this package has six external runtime dependencies that Cloud's sandbox will not load)
- An **Azure Bot** registration (free tier `F0` is fine) with App ID + secret + tenant
- Channel configured in Azure (Teams / Copilot / WebChat / Direct Line)

## Install

### From n8n UI (standalone self-hosted)

**Settings → Community Nodes → Install** → `n8n-nodes-agents-sdk`

### From npm

```bash
npm install n8n-nodes-agents-sdk
```

### Queue mode (app + worker separation)

n8n's UI installer is **not supported in queue mode** — it only installs on the node that handles the UI click, leaving workers out of sync. Install at container boot instead:

```sh
# In your n8n container entrypoint, before exec /docker-entrypoint.sh:
cd /home/node/.n8n/nodes
npm install --no-audit --no-fund --omit=dev --omit=peer --omit=optional \
  n8n-nodes-agents-sdk@<pinned-version>
```

Both `app` and `worker` services need the package installed. With a shared volume you risk install races; use per-service bind mounts.

## Quick start

A working echo bot in three nodes:

```
[M365 Agent Trigger]  →  [M365 Text Message]  →  [M365 Send Activity]
```

1. **Credential.** Create an `M365 Agent API` credential with your Azure Bot App ID, client secret, and tenant.
2. **Trigger.** Drop an **M365 Agent Trigger**, attach the credential. Copy the Production webhook URL from the node UI.
3. **Azure Bot messaging endpoint.** Paste the URL into your Azure Bot registration's *Messaging endpoint* field.
4. **Build the reply.** Add an **M365 Text Message**; set *Text* to `Echo: {{ $json.activity.text }}`.
5. **Send.** Add an **M365 Send Activity**; select *Operation* = `Reply`. Leave `conversationReference` and `activity` at their defaults — they auto-read the envelope.
6. **Activate** the workflow. Send a message to your bot in Teams. You should see the echo within a second.

## Common patterns

### Proactive message (1C, CRM, monitoring, etc. triggering Teams)

```
[Webhook]  →  [Set conversationReference]  →  [M365 Card Template]  →  [M365 Send Activity: proactive]
```

The proactive conversationReference is a plain JSON object pointing at the target channel:

```json
{
  "serviceUrl": "https://smba.trafficmanager.net/emea/<tenant>/",
  "conversation": { "id": "19:...@thread.tacv2" },
  "channelId": "msteams"
}
```

### Interactive card with updates

```
[Trigger] → [Lookup State] → [Code: decide new state] → [M365 Card Template] → [M365 Send Activity: update]
                                                     ↓
                                                    [M365 Text Message: status line]
                                                    [M365 Send Activity: replyInThread]
```

### Adaptive Card templating

Paste a template from the [Adaptive Cards Designer](https://adaptivecards.microsoft.com/designer) into the `Card Template (JSON)` field. Use `${field}` placeholders:

```json
{
  "type": "AdaptiveCard",
  "version": "1.6",
  "speak": "Заказ № ${order_number} принят",
  "body": [
    { "type": "TextBlock", "text": "Заказ № ${order_number}", "weight": "bolder" },
    { "type": "TextBlock", "text": "${car}" }
  ]
}
```

Binding data defaults to `={{ $json }}` — the whole incoming item. The designer's Sample Data Editor accepts the same object for preview.

## Envelope contract

Every node in this package reads from and writes to items of this shape:

```jsonc
{
  "conversationReference": { "serviceUrl", "conversation": {"id", "conversationType"}, "activityId", "channelId", "bot", "user", "locale" },
  "activity": { "type", "text", "attachments", ... },
  "parsed": { "action", "submitData", "userName", "userId", ... },  // Trigger only
  "raw": { /* full incoming Activity */ }                            // Trigger only
}
```

- Trigger emits the full envelope.
- Builders (`M365TextMessage`, `M365CardTemplate`) populate `activity`; preserve `conversationReference` and pass through all ancillary fields (state carried through your workflow).
- Send Activity reads both by default; routes via `conversationReference.serviceUrl` + `conversation.id`, sends `activity`.

## Development

```bash
npm install
npm run build          # compile TS → dist/
npm run build:watch    # tsc --watch
npm run dev            # spins up n8n with the package loaded
npm run lint
npm run lint:fix
npm test               # vitest — unit tests, offline
npm run test:integration   # hits real Azure; needs M365_TEST_* env vars
```

Release is CI-driven. Bump the version, tag, push:

```bash
npm version patch
git push --follow-tags
```

The GitHub Actions workflow (`.github/workflows/publish.yml`) runs lint + test + build + `npm publish` via Trusted Publisher OIDC — no NPM token needed.

## Links

- [npm package](https://www.npmjs.com/package/n8n-nodes-agents-sdk)
- [GitHub repo](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk)
- [Microsoft 365 Agents SDK docs](https://learn.microsoft.com/microsoft-365/agents-sdk/)
- [Azure Bot Service](https://learn.microsoft.com/azure/bot-service/)
- [Adaptive Cards Designer](https://adaptivecards.microsoft.com/designer)

## License

[MIT](./LICENSE) © Ruben Khachaturov
