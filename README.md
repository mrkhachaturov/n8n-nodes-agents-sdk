# n8n-nodes-agents-sdk

[![npm version](https://img.shields.io/npm/v/n8n-nodes-agents-sdk.svg)](https://www.npmjs.com/package/n8n-nodes-agents-sdk)
[![npm downloads](https://img.shields.io/npm/dm/n8n-nodes-agents-sdk.svg)](https://www.npmjs.com/package/n8n-nodes-agents-sdk)
[![license](https://img.shields.io/npm/l/n8n-nodes-agents-sdk.svg)](./LICENSE)
[![node](https://img.shields.io/node/v/n8n-nodes-agents-sdk.svg)](https://nodejs.org/)

Build **Microsoft Teams**, **M365 Copilot**, **WebChat**, and **Direct Line** bots with [n8n](https://n8n.io) as your orchestration layer.

This package wraps the [Microsoft 365 Agents SDK](https://learn.microsoft.com/microsoft-365/agents-sdk/) (the supported successor to `botbuilder`) as first-class n8n nodes — JWT validation, MSAL token management, Bot Connector routing, and Adaptive Card templating all handled for you. No raw HTTP, no hand-assembled Activity JSON, no OAuth plumbing.

---

## What you get

| Without this package | With this package |
|---|---|
| Webhook → Code node to parse Activity JSON | **M365 Agent Trigger** — parsed envelope out of the box |
| Manual JWT validation (often skipped entirely) | Validated on every POST against Microsoft JWKS |
| OAuth2 credential + scope + token-refresh plumbing | **M365 Agent API** credential — App ID + secret + tenant |
| HTTP Request with hand-built `/v3/conversations/...` URLs | **M365 Agent** — pick a resource and an operation from dropdowns |
| Manual `;messageid=` thread suffix for Teams threads | Resource `Message`, operation `Reply in Thread` |
| Hand-written Adaptive Card JSON with string-concat templating | `adaptivecards-templating` + `${field}` bindings built in |

---

## Requirements

- **n8n** 1.54+ (self-hosted — n8n Cloud's sandbox cannot load this package's runtime dependencies)
- **Node.js** 20 or later
- An **Azure Bot** registration (free `F0` tier is fine) with App ID + client secret + tenant
- A channel configured in Azure (Teams, M365 Copilot, WebChat, or Direct Line)

---

## Install

### From the n8n UI (single-instance self-hosted)

**Settings → Community Nodes → Install** → `n8n-nodes-agents-sdk`

### From npm

```bash
npm install n8n-nodes-agents-sdk
```

### Queue mode (separate app + worker)

n8n's UI installer does **not** work in queue mode — it only installs the package on the node that handled the UI click, leaving workers out of sync. Install at container start instead:

```sh
# In your entrypoint, before exec /docker-entrypoint.sh:
cd /home/node/.n8n/nodes
npm install --no-audit --no-fund --omit=dev --omit=peer --omit=optional \
  n8n-nodes-agents-sdk@<pinned-version>
```

Both `app` and `worker` services need the package available. Use per-service bind mounts to avoid install races on shared volumes.

---

## Nodes

### M365 Agent Trigger

Receives POSTs from Azure Bot Service, validates the Bot Framework JWT, and emits a parsed envelope.

- Validates the inbound JWT against the live Microsoft JWKS on every request
- Handles Azure's GET endpoint-verification ping automatically
- Filter by Activity type (`message`, `invoke`, `conversationUpdate`, etc.)
- `Response Mode` parameter: `Immediate` (default — reply 200 right away) or `Wait For Response Node` (keep the HTTP connection open for a downstream **M365 Agent → Invoke Response** node — required for `Action.Execute` invokes and messaging-extension flows)

### M365 Agent

Unified action node. Pick a **Resource** and an **Operation**; the canvas card shows `operation: resource`.

The **Conversation Source** toggle picks between `From Envelope` (default — read the conversation routing from the inbound item, works for every reply/update/delete flow) and `Specify Manually` (fill `serviceUrl`, `conversation.id`, `channelId`, and optional `activityId` yourself — for proactive sends originating outside the bot webhook).

<details>
  <summary><b>Resource: Message</b> — send / reply / update / delete / reply in thread</summary>

  - ✅ **Send** — post a new message to a conversation
  - ✅ **Reply** — reply to a specific activity (auto-threads on Teams)
  - ✅ **Update** — edit a previously-sent activity
  - ✅ **Delete** — delete a previously-sent activity
  - ✅ **Reply in Thread** — reply inside an existing Teams thread via the `;messageid=<parentActivityId>` suffix

  All five operations share a **Text** field (with expression support), a **Conversation Source** selector, and an **Options** collection for optional knobs (`Workflow Footer`, and — in future milestones — mentions and suggested actions).
</details>

<details>
  <summary><b>Resource: Card</b> — send / update Adaptive Cards with templating</summary>

  - ✅ **Send** — render an Adaptive Card template against a data object and send it as a card attachment
  - ✅ **Update** — re-render the template with new binding data and update a card you already posted (requires `activityId` on the conversation reference)

  Author the card in the [Adaptive Cards Designer](https://adaptivecards.microsoft.com/designer), paste the JSON into **Card Template**, and reference fields with `${field}` placeholders. **Binding Data** defaults to the whole inbound item (`={{ $json }}`). **Options → Fallback Text** shows on clients that can't render Adaptive Cards (notifications, mobile lockscreens).
</details>

<details>
  <summary><b>Resource: Invoke Response</b> — respond to Action.Execute / messaging-extension invokes</summary>

  - ✅ **Respond** — write a synchronous response to an invoke activity

  Required when handling `Action.Execute` button callbacks, messaging-extension searches, or any other invoke flow where Teams expects a same-request response. Pair with the Trigger's `Response Mode: Wait For Response Node`. **Response Shape** offers Simple (`{ status, body }` for generic invokes) or Advanced (`{ statusCode, type, value }` for Adaptive Card refreshes and follow-up messages).
</details>

Both nodes use the same credential — either **M365 Agent API** (classic bot) or **M365 Agent 365 API** (Agent 365). Neither is exposed as an AI-agent tool — side effects (sending to Azure Bot Service) and triggers (external webhook) aren't safe for autonomous LLM invocation.

---

## Agent 365 support (v0.3.0)

v0.3.0 adds first-class support for the **Entra Agent Identity Blueprint** alongside the existing classic Azure Bot credential. This lets you register your bot as a Microsoft 365 Agent Identity (preview) and issue tokens through MSAL or an isolated auth sidecar — no changes required for existing classic bot deployments.

### New credential: `M365Agent365Api`

Used when `Authentication Kind = Agent 365`. Replaces the classic App ID + secret fields with blueprint-aware fields (inline MSAL or sidecar URL).

### `authKind` parameter

| Value | Credential required | Description |
|---|---|---|
| `classicBot` | `M365AgentApi` | Existing Azure Bot Service flow — unchanged from v0.2.1 |
| `agent365` | `M365Agent365Api` | Entra Agent Identity Blueprint — inline MSAL or sidecar |

### Identity modes (Agent 365 only)

| Mode | When to use | Transport required |
|---|---|---|
| `autonomous` | Agent acts as its Blueprint app identity | inline or sidecar |
| `agentUser` | Agent acts as its own M365 user (requires separate license) | sidecar only |
| `interactiveOBO` | Delegated calls to Graph/MCP as inbound user | sidecar only (M2-preview, not UI-exposed in M1) |

### Further reading

- **Sidecar deployment**: [examples/sidecar/README.md](examples/sidecar/README.md)
- **Auth guide** (choosing between classic, inline, and sidecar): [docs/auth-guide.md](docs/auth-guide.md)
- **OBO guide** (why OBO is M2-only and what it unlocks): [docs/obo-guide.md](docs/obo-guide.md)
- **Licensing notes** (Frontier preview vs GA cost implications): [docs/licensing-notes.md](docs/licensing-notes.md)

---

## Credential

**M365 Agent API** — your Azure Bot registration's App ID, client secret, and tenant. Three app types:

- **SingleTenant** *(recommended)* — token scoped to a single Entra tenant
- **MultiTenant** — kept for legacy registrations; Microsoft deprecated this path for new bots after 2025-07-31
- **UserAssignedMsi** — Azure managed identity

The built-in credential test hits `login.microsoftonline.com` to verify network reachability. Full token-acquisition validation runs inside the Trigger on every inbound webhook.

---

## Quick start — echo bot in 3 nodes

```
[M365 Agent Trigger]  →  [Set]  →  [M365 Agent — Message: Reply]
```

1. **Credential.** Create an **M365 Agent API** credential with your Azure Bot App ID, client secret, and tenant.
2. **Trigger.** Drop an **M365 Agent Trigger**, attach the credential. Copy the node's Production webhook URL.
3. **Messaging endpoint.** Paste the URL into your Azure Bot registration's *Messaging endpoint* field.
4. **Compose the reply.** Drop a **Set** node (or just write an expression directly in step 5's Text field): `Echo: {{ $json.activity.text }}`.
5. **Reply.** Drop an **M365 Agent** node. Set **Resource** = `Message`, **Operation** = `Reply`. Leave **Conversation Source** at `From Envelope`. Fill **Text** with the expression from step 4.
6. **Activate** the workflow, message your bot in Teams, and the echo comes back within a second.

---

## Envelope contract

Every node in this package reads and writes items of this shape:

```jsonc
{
  "conversationReference": {
    "serviceUrl", "conversation": { "id", "conversationType" },
    "activityId", "channelId", "bot", "user", "locale"
  },
  "activity": { "type", "text", "attachments", ... },
  "parsed": {      // Trigger only — convenience fields
    "action", "submitData", "userName", "userId", ...
  },
  "raw": { ... }    // Trigger only — full incoming Activity
}
```

**Contract:**
- The **Trigger** emits the full envelope on every inbound activity.
- The **M365 Agent** node reads `conversationReference` (via the `From Envelope` default) to route the outbound call. On output it spreads every input field through unchanged and adds an operation-specific result field (`sendResult` / `replyResult` / `updateResult` / `deleteResult` / `replyInThreadResult` / `cardSendResult` / `cardUpdateResult`).

Ancillary state you carry through your workflow (correlation IDs, cached state, upstream webhook payloads, etc.) survives the `M365 Agent` node unchanged.

---

## Adaptive Card templating

Paste a template from the [Adaptive Cards Designer](https://adaptivecards.microsoft.com/designer) into **Card Template**. Use `${field}` for placeholders. `${field}` bindings resolve against whatever object you put in **Binding Data** (defaults to the whole inbound item):

```json
{
  "type": "AdaptiveCard",
  "version": "1.6",
  "body": [
    { "type": "TextBlock", "text": "Order #${orderId}", "weight": "bolder", "size": "large" },
    { "type": "TextBlock", "text": "${status}" }
  ],
  "actions": [
    { "type": "Action.Submit", "title": "Acknowledge", "data": { "action": "ack", "orderId": "${orderId}" } }
  ]
}
```

Any `[Parsing] Unknown property` warnings the designer shows on `Action.Submit.data` custom keys are [false positives](https://github.com/microsoft/AdaptiveCards/blob/main/samples/v1.0/Tests/Feedback.json) — they're part of the Adaptive Cards spec, just not typed in the designer's bundled schema.

---

## Proactive sends (no inbound activity)

To message a channel from a workflow that wasn't triggered by the Trigger (a schedule, an external webhook, a database row, anything), switch **Conversation Source** to `Specify Manually` and fill in the fields. For Teams channels, `conversation.id` looks like `19:xxx@thread.tacv2` and you'll typically have captured it from a prior bot interaction.

```
[Schedule Trigger]  →  [HTTP Request: fetch data]  →  [M365 Agent — Card: Send (manual)]
```

---

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

Release is CI-driven via GitHub Actions Trusted Publishing — no NPM token needed. Bump the version, tag, push:

```bash
npm version minor       # or patch / major
git push --follow-tags
```

---

## Links

- [npm package](https://www.npmjs.com/package/n8n-nodes-agents-sdk)
- [GitHub repository](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk)
- [Microsoft 365 Agents SDK documentation](https://learn.microsoft.com/microsoft-365/agents-sdk/)
- [Azure Bot Service documentation](https://learn.microsoft.com/azure/bot-service/)
- [Adaptive Cards Designer](https://adaptivecards.microsoft.com/designer)

---

## License

[MIT](./LICENSE) © Ruben Khachaturov
