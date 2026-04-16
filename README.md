# n8n-nodes-agents-sdk

n8n community nodes for the Microsoft 365 Agents SDK. Build Teams / Copilot / WebChat bots with n8n as the orchestration layer — no raw HTTP, no hand-built Activity JSON, JWT validation done for you.

## Status

Milestone 0 — four nodes and one credential. See the design spec for the full roadmap.

- `M365AgentTrigger` — JWT-validated webhook, emits a parsed envelope (`conversationReference` + `activity` + `parsed` + `raw`)
- `M365TextMessage` — text Activity builder
- `M365CardTemplate` — Adaptive Card template with `${field}` data binding (via `adaptivecards-templating`)
- `M365SendActivity` — reply / proactive / update / delete / replyInThread, covering the Bot Connector surface
- `M365AgentApi` credential — SingleTenant / MultiTenant / UserAssignedMsi

## Install

```bash
npm install n8n-nodes-agents-sdk
```

Or in the n8n UI: **Settings → Community Nodes → Install** → `n8n-nodes-agents-sdk`.

The package has external runtime dependencies (MSAL via `@microsoft/agents-hosting`, `axios`, `jsonwebtoken`, `jwks-rsa`, `adaptivecards-templating`) and is intended for self-hosted n8n — it will not run in n8n Cloud's sandbox.

## Quick start

1. Create a credential of type **Microsoft 365 Agent API** with your Azure Bot App ID, secret, and tenant.
2. Drop an **M365 Agent Trigger** into a workflow; point the Azure Bot registration's messaging endpoint at the webhook URL n8n produces.
3. Build the reply Activity with **M365 Text Message** or **M365 Card Template**.
4. Send it back with **M365 Send Activity** (operation `reply` for a response, `proactive` for an unsolicited push, `update` to edit a card in-place, `replyInThread` for Teams thread replies).

## License

MIT — Ruben Khachaturov
