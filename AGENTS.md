# AGENTS.md

Working context for AI coding agents inside this package. Portable format — Claude Code reads it via `CLAUDE.md → @AGENTS.md` import, Codex/Cursor/other agents read it directly.

## What this is

`n8n-nodes-m365-agents` — an n8n community-nodes package that exposes the **Microsoft 365 Agents SDK** as first-class n8n nodes. Lets n8n workflows receive/send/update Activities from Azure Bot Service across Teams, M365 Copilot, WebChat, and Direct Line — without raw HTTP or manual JWT validation.

Prototype stage. Currently lives under `.dev/n8n-nodes-m365-agents/` inside the parent `n8n-workflows` repo. Will graduate to a public GitHub repo at `github.com/mrkhachaturov/n8n-nodes-m365-agents` once stable, and may be pulled back in as a git submodule.

## Architecture (where this sits)

```
Teams / M365 Copilot / WebChat / Direct Line  (Azure Bot Service channels)
            │
            ▼
     Azure Bot Service              (routing + identity, free tier)
            │  Activity JSON + JWT
            ▼
  M365AgentTrigger (this package)   ← receives, validates JWT, parses
            │
            ▼
  n8n workflow (your business logic, 1C, AI, etc.)
            │
            ▼
  M365 builders (TextMessage / CardTemplate / ...)  ← populate `activity`; preserve `conversationReference`
            │
            ▼
  M365SendActivity (this package)   ← reply / proactive / update / delete / replyInThread
            │
            ▼
     Azure Bot Service → channel → user
```

### Item envelope (contract between nodes)

Every node in this package reads/writes items of this shape:

```jsonc
{
  "conversationReference": { "serviceUrl", "conversation": {"id", "conversationType"}, "activityId", "bot", "user", "channelId", "locale" },
  "activity": { "type", "text", "attachments", ... },   // populated by builders
  "parsed": { "action", "submitData", ... },             // convenience, Trigger only
  "raw": { /* full incoming Activity */ }                // Trigger only
}
```

**Contract:**
- `M365AgentTrigger` emits the full envelope
- Builder nodes (`M365TextMessage`, `M365CardTemplate`, etc.) populate/replace `activity`; must preserve `conversationReference` unchanged
- `M365SendActivity` reads both fields; routes via `conversationReference`, sends `activity`

Violating this (e.g., a builder that strips `conversationReference`) is a bug.

## Tech stack

- **Language**: TypeScript 6.x, compiled to CommonJS (`dist/`)
- **Build tool**: `@n8n/node-cli` — wraps tsc + eslint + release-it
- **Runtime dep**: `@microsoft/agents-hosting` ^1.4.2 (the new SDK, not archived `botbuilder`)
- **Peer dep**: `n8n-workflow` ^2.13.1
- **Node**: >=20 (Node 22 via parent `mise.toml`)

## Skills available

This project is inside the Claude Code `n8n-workflows` workspace. The following skills (plugins) are enabled:

| Skill | Use when |
|-------|----------|
| `/microsoft-docs` | Looking up Microsoft/Azure concept docs — Agents SDK, Bot Framework, Teams, Adaptive Cards, Entra |
| `/microsoft-code-reference` | Finding working MS SDK code samples, verifying API signatures |
| `/microsoft-skill-creator` | Generating new skills for other Microsoft technologies |
| `/n8n-mcp-tools-expert` | Using `n8n-dev` MCP tools (node discovery, validation, templates) |
| `/n8n-code-javascript` | Writing JS code for n8n Code nodes or node internals |
| `/n8n-node-configuration` | Configuring node properties, `displayOptions`, required fields |
| `/n8n-validation-expert` | Interpreting `validate_node` / `validate_workflow` errors |

**Prefer the `/microsoft-docs` skill over WebFetch** — it queries the Microsoft Learn MCP server directly and returns clean excerpts.

## Local reference repositories

Two upstream Microsoft repos are cloned to the parent workspace `.local/` directory. Read them directly instead of searching the web.

### `../../.local/Agents/` — multi-language umbrella repo (microsoft/Agents)

Authoritative protocol specs and cross-language samples.

| Path | What's there |
|------|--------------|
| `specs/activity/protocol-activity.md` | Activity JSON schema — ground truth for Trigger output shape |
| `specs/activity/protocol-cards.md` | Card protocol |
| `specs/channel-api/ChannelAPI-OpenAPI.yaml` | Bot Connector REST API (sendToConversation, updateActivity, deleteActivity) |
| `specs/channel-api/TokenAPI-OpenAPI.yaml` | Token API for sign-in flows |
| `specs/manifest/` | Teams manifest schema |
| `samples/nodejs/quickstart/` | Minimal bot scaffold |
| `samples/nodejs/cards/` | Card patterns |
| `samples/nodejs/auto-signin/` | OAuth sign-in |
| `samples/nodejs/obo-authorization/` | On-Behalf-Of token exchange |
| `samples/nodejs/multi-turn-prompt/` | Conversation state |
| `samples/nodejs/langchain-multiturn/` | LangChain orchestration |
| `AgentErrorCodesJS.md` | JS error codes reference |

### `../../.local/Agents-for-js/` — the JavaScript SDK (microsoft/Agents-for-js)

Actual TypeScript source for the npm packages we depend on, plus ready-to-run samples.

| Path | What's there |
|------|--------------|
| `packages/agents-hosting/src/` | `CloudAdapter`, `TurnContext`, JWT validation — reference for what we need to replicate in the Trigger |
| `packages/agents-activity/src/` | Activity type definitions |
| `packages/agents-hosting-extensions-teams/` | Teams-specific helpers (mentions, meeting data, channel data) |
| `samples/basic/echo.ts` | Minimal echo bot — the reference pattern |
| `samples/basic/cardsWithInvoke.ts` | Adaptive Card `Action.Submit` handling |
| `samples/basic/proactive.ts` | Bot-initiated conversations |
| `samples/basic/streamingSample.ts` | Streaming responses |
| `samples/basic/longRunning.ts` | Typing indicator during long ops |
| `samples/basic/errorHandling.ts` | Error patterns |
| `samples/teams/cardActions.ts` | Teams button callbacks |
| `samples/teams/taskModuleExample.ts` | Teams popup dialogs |
| `samples/teams/msgExtensionExample.ts` | Messaging extensions |

## Commands

```bash
npm install          # first time
npm run build        # compile to dist/
npm run build:watch  # tsc --watch
npm run dev          # spins up n8n with this package loaded
npm run lint         # n8n-specific eslint rules
npm run lint:fix     # auto-fix
npm test             # vitest run
npm run test:watch   # vitest --watch
npm run release      # release-it (once published)
```

Or via `just` (same tasks, shorter):

```bash
just build              # npm run build
just test               # npm test (unit only, offline)
just test-integration   # hits real Azure via mise-loaded credentials
just lint               # npm run lint
just check              # lint + test + build
just watch              # vitest --watch
just deploy-dev         # rsync dist/ to CephFS + force n8n service update
```

## Integration tests (real Azure)

`tests/integration/azure.integration.test.ts` hits `login.microsoftonline.com`
to verify end-to-end that:
1. `buildAuthConfig` + `MsalTokenProvider` actually acquire a Bot Framework token
2. The token passes our `verifyJwt` validator against the real JWKS endpoint

Credentials come from the parent repo's mise-managed 1Password env
(`~/.op-env/at-m365bot.env`) which exports:

```
M365_TEST_CLIENT_ID
M365_TEST_TENANT_ID
M365_TEST_CLIENT_SECRET
```

`just test-integration` wraps `mise exec --` so the env loads automatically.
Without the env vars, the integration `describe` block is skipped cleanly —
CI and anyone without credentials can still run `npm test` offline.

## Dev deploy loop (Docker Swarm + CephFS)

The n8n stack mounts `${SHARED}/automation/n8n` into `/home/node/.n8n` (see
`Infra/Containers/swarm/stacks/automation/n8n/docker-compose.yml`). n8n reads
custom nodes from `/home/node/.n8n/custom`. `just deploy-dev` rsyncs our
`dist/` into that CephFS-backed folder then forces a service update so the
new code loads.

Override targets per-invocation if your setup differs from the defaults:

```bash
SWARM_HOST=node01.astrateam.net \
CEPHFS_PATH=/mnt/shared/automation/n8n/custom \
N8N_SERVICE=n8n_app \
just deploy-dev
```

## Project layout

```
.
├── credentials/
│   └── M365AgentApi.credentials.ts        — App ID / secret / tenant / appType
├── nodes/                                  — Milestone 0 set
│   ├── M365AgentTrigger/                   — JWT-validated webhook + GET health
│   ├── M365SendActivity/                   — reply / proactive / update / delete / replyInThread
│   ├── M365TextMessage/                    — text Activity builder
│   └── M365CardTemplate/                   — Adaptive Card + data binding
├── shared/
│   ├── types.ts                            — envelope, credential, operation types
│   ├── buildAuthConfig.ts                  — credential → SDK AuthConfiguration
│   ├── verifyJwt.ts                        — standalone JWT validator
│   ├── envelope.ts                         — envelope helpers (ConvRef/parse/merge)
│   └── botConnector.ts                     — MSAL + axios + replyInThread
├── tests/
│   ├── credentials/
│   ├── nodes/
│   └── shared/
├── icons/                                  — m365.svg, m365.dark.svg
├── dist/                                   — build output (gitignored)
├── justfile                                — build, test, lint, deploy-dev
├── eslint.config.mjs                       — configWithoutCloudSupport + ignores
├── vitest.config.ts
├── package.json
├── tsconfig.json
├── CLAUDE.md                               — one line: @AGENTS.md
├── AGENTS.md                                — this file
└── README.md                                — public-facing intro
```

## Conventions

- **Never trust Activity defaults.** Always set all node parameters explicitly. Default values in n8n node configs cause silent runtime failures.
- **Use `@microsoft/agents-hosting`, never `botbuilder`.** The old `botbuilder` SDK is archived (Dec 31, 2025).
- **Single-tenant / UserAssignedMsi over multi-tenant.** Multi-tenant bot creation is deprecated after 2025-07-31.
- **Validate JWT on every incoming webhook.** Use the standalone verifier in `shared/verifyJwt.ts` (SDK pieces — `jsonwebtoken`, `jwks-rsa` — driven by the same logic as the SDK's `authorizeJWT` middleware). The SDK's `CloudAdapter.process()` is Express-coupled and incompatible with n8n's `IWebhookFunctions`; replicating ~40 lines of its JWT logic is the correct path. Never skip validation on POST.
- **Support both POST and GET** on the trigger path — POST for messages, GET for health checks / manifest URL verification. Note: the two webhooks are registered as `name: 'default'` (POST) and `name: 'setup'` (GET). `n8n-workflow` only accepts `default | setup` as `webhooks[].name`; that's why the GET is called "setup" even though the path is `/messages`.
- **No premature abstractions.** Start with 2-3 nodes, add more only when a real workflow needs them.
- **Channel-aware but not channel-specific by default.** Use `Activity.ChannelId` + `ChannelData` for per-channel logic rather than separate nodes per channel.

## Decisions log

- **2026-04-16** — SDK choice: `@microsoft/agents-hosting` (not `botbuilder`). Reason: Bot Framework SDK archived 2025-12-31.
- **2026-04-16** — Package naming: unscoped `n8n-nodes-m365-agents` for now. May move under `@mrkhachaturov/` scope on publish.
- **2026-04-16** — Dependency strategy: pinned current majors (TS 6, ESLint 10, release-it 20). Verified compatible with `@n8n/node-cli@0.23.1`.
- **2026-04-17** — ESLint downgraded to `^9.29.0` during M0 build-out. Reason: `@n8n/node-cli@0.23.1` tests with eslint 9 and `eslint-plugin-n8n-nodes-base` uses `context.getFilename()` which was removed in ESLint 10. Revisit when the n8n CLI updates.
- **2026-04-17** — n8n Cloud support OFF: `eslint.config.mjs` uses `configWithoutCloudSupport` and `package.json` `n8n.strict: false`. Reason: the package has six external runtime deps (`@microsoft/agents-hosting`, `@microsoft/agents-activity`, `axios`, `jsonwebtoken`, `jwks-rsa`, `adaptivecards-templating`) — n8n Cloud's sandbox won't load it. Self-hosted n8n is the target.

## References (external, only fetch when skills can't answer)

- https://learn.microsoft.com/microsoft-365/agents-sdk/
- https://github.com/microsoft/Agents (cloned at `../../.local/Agents/`)
- https://github.com/microsoft/Agents-for-js (cloned at `../../.local/Agents-for-js/`)
- https://docs.n8n.io/integrations/creating-nodes/
