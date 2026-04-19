# n8n-nodes-agents-sdk

[![npm version](https://img.shields.io/npm/v/n8n-nodes-agents-sdk.svg)](https://www.npmjs.com/package/n8n-nodes-agents-sdk)
[![npm downloads](https://img.shields.io/npm/dm/n8n-nodes-agents-sdk.svg)](https://www.npmjs.com/package/n8n-nodes-agents-sdk)
[![license](https://img.shields.io/npm/l/n8n-nodes-agents-sdk.svg)](./LICENSE)
[![node](https://img.shields.io/node/v/n8n-nodes-agents-sdk.svg)](https://nodejs.org/)

Build **Microsoft Teams**, **M365 Copilot**, **WebChat**, and **Direct Line** bots with [n8n](https://n8n.io) as your orchestration layer.

This is a community node for self-hosted n8n. It wraps the [Microsoft 365 Agents SDK](https://learn.microsoft.com/microsoft-365/agents-sdk/) (the supported successor to `botbuilder`) so you get JWT validation, MSAL token management, Bot Connector routing, and the **full CardFactory surface** — Adaptive, Hero, Thumbnail, Animation, Audio, Video, Sign-In, Receipt, O365 Connector, and Raw Attachment — as first-class n8n operations. No raw HTTP, no hand-assembled Activity JSON, no OAuth plumbing.

> **0.4.0 breaking change.** The single `Card` resource was split into ten first-class card resources (Adaptive, Animation, Audio, Hero, O365 Connector, Raw Attachment, Receipt, Sign-In, Thumbnail, Video). Workflows saved on 0.3.x with `resource: "card"` will fail loudly on load — rebuild them on `resource: "adaptiveCard"` (for templated Adaptive Cards) or the appropriate new card resource. See [CHANGELOG.md](./CHANGELOG.md) for the full list.

---

## What you get

| Without this package                                      | With this package                                                                                                  |
| --------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| Webhook → Code node to parse Activity JSON                | **M365 Agent Trigger** — parsed envelope out of the box                                                            |
| Manual JWT validation (often skipped entirely)            | Validated on every POST against Microsoft JWKS                                                                     |
| OAuth2 credential + scope + token-refresh plumbing        | **M365 Agent API** credential — App ID + secret + tenant                                                           |
| HTTP Request with hand-built `/v3/conversations/...` URLs | **M365 Agent** — pick a resource and an operation from dropdowns                                                   |
| Manual `;messageid=` thread suffix for Teams threads      | Resource `Message`, operation `Reply in Thread`                                                                    |
| Hand-written card JSON for every card type                | 10 first-class card resources — builder UI for rich cards, templating for Adaptive, raw JSON for connector/receipt |

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

n8n's UI installer does **not** work in queue mode — it only installs on the node that handled the UI click, leaving workers out of sync. Install at container start instead:

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

Unified action node. Pick a **Resource** and an **Operation**; the canvas subtitle renders as `<operation> <resource>` (for example `send heroCard`, `update adaptiveCard`).

The **Conversation Source** toggle picks between `From Envelope` (default — read the conversation routing from the inbound item, works for every reply/update/delete flow) and `Specify Manually` (fill `serviceUrl`, `conversation.id`, `channelId`, and optional `activityId` yourself — for proactive sends originating outside the bot webhook).

<details>
  <summary><b>Resource: Message</b> — send / reply / update / delete / reply in thread</summary>

- ✅ **Send** — post a new message to a conversation
- ✅ **Reply** — reply to a specific activity (auto-threads on Teams)
- ✅ **Update** — edit a previously-sent activity
- ✅ **Delete** — delete a previously-sent activity
- ✅ **Reply in Thread** — reply inside an existing Teams thread via the `;messageid=<parentActivityId>` suffix

All five operations share a **Text** field (with expression support), a **Conversation Source** selector, and an **Options** collection for optional knobs (`Workflow Footer`, Mentions, Suggested Actions).

**Options (inside every body-carrying operation — Send / Reply / Update / Reply in Thread):**

- **Mentions** — add Teams @-mentions. `User` pings one person (paste the user's Teams/AAD object id and display name from your inbound trigger envelope). `Everyone` pings the whole team (uses Teams' magic ID internally). The node inserts matching `<at>Name</at>` tokens into the text and attaches the entities — no hand-assembly needed.
- **Suggested Actions** — quick-reply chip buttons under the message. All 11 SDK action types supported (`imBack`, `messageBack`, `postBack`, `openUrl`, `call`, `downloadFile`, `openApp`, `playAudio`, `playVideo`, `showImage`, `signin`). Channel support varies — Teams, WebChat, and Direct Line each render a subset.

</details>

<details>
  <summary><b>Resource: Adaptive Card</b> — send / update Adaptive Cards with templating</summary>

- ✅ **Send** — render an Adaptive Card template against a data object and send it as a card attachment
- ✅ **Update** — re-render the template with new binding data and update a card you already posted (requires `activityId` on the conversation reference)

Author the card in the [Adaptive Cards Designer](https://adaptivecards.microsoft.com/designer), paste the JSON into **Card Template**, and reference fields with `${field}` placeholders. **Binding Data** defaults to the whole inbound item (`={{ $json }}`). **Options → Fallback Text** shows on clients that can't render Adaptive Cards (notifications, mobile lockscreens).

</details>

<details>
  <summary><b>Resources: Hero / Thumbnail / Animation / Audio / Video / Sign-In / Receipt / O365 Connector / Raw Attachment</b> — the rest of the CardFactory family</summary>

All card resources support **Send** and **Update** operations and share the same authentication, conversation-source, and fallback-text conventions as the Adaptive Card resource.

- **Hero Card** — title / subtitle / text / images / buttons / tap.
- **Thumbnail Card** — same shape as Hero, different rendering.
- **Animation Card** — GIF / animation card with media URLs + buttons (requires at least one media URL).
- **Audio Card** — audio card with media URLs + buttons.
- **Video Card** — video card with media URLs + buttons.
- **Sign-In Card** — sign-in card with title / URL / body text.
- **Receipt Card** — receipt card via JSON content (title / facts / items / total).
- **O365 Connector Card** — Teams-specific connector card via JSON content (title / sections / potentialAction).
- **Raw Attachment** — power-user escape hatch. Supply `Content Type` + `Content` (JSON) directly; used for card families or attachment types not yet first-class here.

Buttons and tap actions across the builder-family resources use a unified **CardAction** row (Type + Title + Value + optional Channel Data / Value JSON / image / text / displayText), covering all 11 SDK action types.

</details>

<details>
  <summary><b>Resource: Invoke Response</b> — respond to Action.Execute / messaging-extension invokes</summary>

- ✅ **Respond** — write a synchronous response to an invoke activity

Required when handling `Action.Execute` button callbacks, messaging-extension searches, or any other invoke flow where Teams expects a same-request response. Pair with the Trigger's `Response Mode: Wait For Response Node`. **Response Shape** offers Simple (`{ status, body }` for generic invokes) or Advanced (`{ statusCode, type, value }` for Adaptive Card refreshes and follow-up messages).

</details>

Both nodes use the same credential — either **M365 Agent API** (classic bot) or **M365 Agent 365 API** (Agent 365). Neither is exposed as an AI-agent tool — side effects (sending to Azure Bot Service) and external webhook triggers aren't safe for autonomous LLM invocation.

---

## Credential

**M365 Agent API** — your Azure Bot registration's App ID, client secret, and tenant. Three app types:

- **SingleTenant** _(recommended)_ — token scoped to a single Entra tenant
- **MultiTenant** — kept for legacy registrations; Microsoft deprecated this path for new bots after 2025-07-31
- **UserAssignedMsi** — Azure managed identity

The built-in credential test hits `login.microsoftonline.com` to verify network reachability. Full token-acquisition validation runs inside the Trigger on every inbound webhook.

---

## Agent 365 support

v0.3.0 added first-class support for the **Entra Agent Identity Blueprint** alongside the classic Azure Bot credential. Register your bot as a Microsoft 365 Agent Identity (preview) and issue tokens through MSAL or an isolated auth sidecar — no changes required for existing classic bot deployments.

### `authKind` parameter

| Value        | Credential required | Description                                             |
| ------------ | ------------------- | ------------------------------------------------------- |
| `classicBot` | `M365AgentApi`      | Existing Azure Bot Service flow — unchanged from v0.2.1 |
| `agent365`   | `M365Agent365Api`   | Entra Agent Identity Blueprint — inline MSAL or sidecar |

### Identity modes (Agent 365 only)

| Mode             | When to use                                                 | Transport required                                                    |
| ---------------- | ----------------------------------------------------------- | --------------------------------------------------------------------- |
| `autonomous`     | Agent acts as its Blueprint app identity                    | inline or sidecar                                                     |
| `agentUser`      | Agent acts as its own M365 user (requires separate license) | sidecar only                                                          |
| `interactiveOBO` | Delegated calls to Graph / MCP as the inbound user          | sidecar only (router-only; not yet surfaced on any card / message UI) |

### Further reading

- **Sidecar deployment**: [examples/sidecar/README.md](examples/sidecar/README.md)
- **Auth guide** (choosing between classic, inline, and sidecar): [docs/auth-guide.md](docs/auth-guide.md)
- **OBO guide** (what `interactiveOBO` is and why it's router-only today): [docs/obo-guide.md](docs/obo-guide.md)
- **Licensing notes** (Frontier preview vs GA cost implications): [docs/licensing-notes.md](docs/licensing-notes.md)

---

## Quick start — echo bot in 3 nodes

```
[M365 Agent Trigger]  →  [Set]  →  [M365 Agent — Message: Reply]
```

1. **Credential.** Create an **M365 Agent API** credential with your Azure Bot App ID, client secret, and tenant.
2. **Trigger.** Drop an **M365 Agent Trigger**, attach the credential. Copy the node's Production webhook URL.
3. **Messaging endpoint.** Paste the URL into your Azure Bot registration's _Messaging endpoint_ field.
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

## Cards

Ten card resources cover every shape the Bot Framework / Teams renders. Pick the one that matches your UX — they fall into three authoring styles:

### 1. Builder cards — fill in a form, no JSON

**Hero Card**, **Thumbnail Card**, **Animation Card**, **Audio Card**, **Video Card**, **Sign-In Card** all expose a form UI. You fill in title, subtitle, text, images, buttons, media URLs — the node assembles the correct SDK shape for you. This is the simplest path for the common case (an order-update card, a "click here to sign in" card, a notification with quick-reply buttons).

The media-family (Animation / Audio / Video) requires at least one media URL and lets you set `autoloop` / `autostart` / `aspect` / `duration` under **Options**. Every builder card's **Buttons** and **Tap Action** share the same unified **CardAction** row (Type + Title + Value + optional Channel Data / Value JSON / image / text / displayText) across all 11 SDK action types (`openUrl`, `imBack`, `messageBack`, `postBack`, `call`, `downloadFile`, `openApp`, `playAudio`, `playVideo`, `showImage`, `signin`). Channel support varies — Teams, WebChat, and Direct Line each render a subset.

Minimal Hero Card workflow parameters:

```yaml
resource: heroCard
operation: send
title: 'Order #{{$json.orderId}}'
subtitle: 'Ready for pickup'
text: 'Tap to confirm'
buttons:
  button:
    - type: openUrl
      title: 'View order'
      value: '{{$json.orderUrl}}'
options:
  fallbackText: 'Order ready'
  tapAction:
    type: openUrl
    value: '{{$json.orderUrl}}'
```

### 2. Adaptive Card — templated JSON with bindings

Paste a template from the [Adaptive Cards Designer](https://adaptivecards.microsoft.com/designer) into **Card Template**. Use `${field}` for placeholders. `${field}` bindings resolve against whatever object you put in **Binding Data** (defaults to the whole inbound item `={{ $json }}`):

```json
{
	"type": "AdaptiveCard",
	"version": "1.6",
	"body": [
		{ "type": "TextBlock", "text": "Order #${orderId}", "weight": "bolder", "size": "large" },
		{ "type": "TextBlock", "text": "${status}" }
	],
	"actions": [
		{
			"type": "Action.Submit",
			"title": "Acknowledge",
			"data": { "action": "ack", "orderId": "${orderId}" }
		}
	]
}
```

Any `[Parsing] Unknown property` warnings the designer shows on `Action.Submit.data` custom keys are [false positives](https://github.com/microsoft/AdaptiveCards/blob/main/samples/v1.0/Tests/Feedback.json) — they're part of the Adaptive Cards spec, just not typed in the designer's bundled schema.

**State-driven variants.** You don't need a separate "pick-a-template-by-state" node. `adaptivecards-templating` supports this directly:

1. **`${field}`** — scalar binding. Change a value across the card with one variable: `"style": "${statusStyle}"`, `"text": "${statusBadgeText}"`.
2. **`$when`** — conditional rendering. Show/hide whole sections per state: `{ "$when": "${status == 'completed'}", "type": "TextBlock", "text": "Done" }`.
3. **`$data`** — iterate over an array. Render one button per item in `cardButtons`.

A single template handles every state (new / accepted / in_progress / paused / completed, etc.) — no picker node needed. See the upstream [`adaptivecards-templating` docs](https://www.npmjs.com/package/adaptivecards-templating) for the full reference.

### 3. JSON-content cards — paste the whole card

**Receipt Card**, **O365 Connector Card**, and **Raw Attachment** take the full card object as JSON in a single field. Use these when the card shape is fixed (Teams connector cards with their own schema) or when you want an escape hatch for attachment types not yet first-class here.

- **Receipt** — `{ title, facts[], items[], total, tax?, vat?, tap?, buttons[] }`.
- **O365 Connector** — Teams-specific: `{ title, summary, themeColor, sections[], potentialAction[] }`. See the [cards reference](https://learn.microsoft.com/microsoftteams/platform/task-modules-and-cards/cards/cards-reference#office-365-connector-card).
- **Raw Attachment** — you provide `contentType` + `content` directly. Useful for card families not yet first-class in this package.

---

## Proactive sends (no inbound activity)

To message a channel from a workflow that wasn't triggered by the Trigger (a schedule, an external webhook, a database row, anything), switch **Conversation Source** to `Specify Manually` and fill in the fields. For Teams channels, `conversation.id` looks like `19:xxx@thread.tacv2` and you'll typically have captured it from a prior bot interaction.

```
[Schedule Trigger]  →  [HTTP Request: fetch data]  →  [M365 Agent — Adaptive Card: Send (manual)]
```

---

## Issues & contributions

Bug reports, feature requests, and PRs are welcome at [github.com/mrkhachaturov/n8n-nodes-agents-sdk](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk/issues). Please include your n8n version, the resource/operation you're using, and the relevant workflow JSON or error message.

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
