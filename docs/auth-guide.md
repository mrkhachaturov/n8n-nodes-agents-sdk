# Auth Guide

How to choose and configure the authentication setup for `n8n-nodes-agents-sdk`.

---

## Classic Bot (Azure Bot Service) — unchanged from v0.2.1

Use when you already have an Azure Bot registration and do not need Agent 365 identity features.

**Steps:**

1. In the [Azure Portal](https://portal.azure.com), create an **Azure Bot** resource (free `F0` tier is fine).
2. In the Bot's **Configuration** blade, set the **Messaging endpoint** to the Production webhook URL of your `M365AgentTrigger` node.
3. In the Bot's **App Registration** (linked from the Bot resource), note the **Application (client) ID** and **Directory (tenant) ID**.
4. Under **Certificates & secrets**, create a new client secret and copy the value immediately.
5. In n8n, create an **M365 Agent API** credential:
   - **App ID** — the Application (client) ID
   - **Client Secret** — the secret value you copied
   - **Tenant ID** — the Directory (tenant) ID
   - **App Type** — `SingleTenant` (recommended for new registrations)
6. On your `M365AgentTrigger` node, set **Authentication Kind** = `Classic Bot` and attach the credential.

No migration is needed if you are upgrading from v0.2.x — the `M365AgentApi` credential shape is unchanged.

---

## Agent 365 — Inline (in-process MSAL)

Use when you want Agent 365 identity without running a separate sidecar process.

**When to choose inline:**

- Your workload only needs the `autonomous` identity mode (agent acts as its Blueprint app identity).
- You are comfortable storing the Blueprint client secret in the n8n credential store.
- You do not need cert-based auth.

**Setup:**

1. Provision an Entra Agent Identity Blueprint for your agent (see the [Entra Agent ID skill](../.claude/skills/entra-agent-id/) for scripted provisioning).
2. Note the **Blueprint App ID**, **Blueprint Client Secret**, and **Tenant ID**.
3. In n8n, create an **M365 Agent 365 API** credential:
   - **Transport** — `inline`
   - **Inline Credential Kind** — `clientSecret`
   - **Blueprint App ID** — the Blueprint's Application (client) ID
   - **Blueprint Secret** — the client secret
   - **Tenant ID** — your Entra tenant ID
4. On your `M365AgentTrigger` and `M365Agent` nodes, set **Authentication Kind** = `Agent 365` and attach the credential.
5. On outbound operations (Message, Card), set **Identity Mode** = `autonomous`.

**Limitations in M1:**

- Only `autonomous` identity mode is supported with the inline transport.
- Certificate-based auth for the Blueprint app is NOT supported in inline mode in M1. Use the sidecar if cert-based auth is required for production.

---

## Agent 365 — Sidecar (HTTP to Entra SDK sidecar)

Use when you need agent-user tokens, cert-based auth, or want auth logic isolated in a dedicated runtime.

**When to choose sidecar:**

- You need the `agentUser` identity mode (agent acts as its own M365 user).
- You want certificate-based auth for the Blueprint app registration (preferred for production).
- You are planning to use M2 Graph/MCP operations with `interactiveOBO` mode.
- You want to keep token acquisition outside the n8n process for auditability.

**Setup:**

1. Deploy the auth sidecar using the artifacts in [`examples/sidecar/`](../examples/sidecar/README.md). The sidecar image is `mcr.microsoft.com/entra-sdk/auth-sidecar`.
2. Note the sidecar's internal URL (e.g., `http://entra-sidecar:8080` inside your stack network).
3. Provision the Blueprint app registration and configure it in the sidecar's environment (see the sidecar README for env vars).
4. In n8n, create an **M365 Agent 365 API** credential:
   - **Transport** — `sidecar`
   - **Sidecar URL** — the sidecar's internal base URL
   - **Agent Instance App ID** — the runtime app ID the sidecar uses for this agent instance
   - **Blueprint App ID** — the Blueprint's Application (client) ID
   - **Outbound Downstream API** (optional) — target API audience for downstream calls (M2)
5. On your nodes, set **Authentication Kind** = `Agent 365` and attach the credential.
6. On outbound operations, set **Identity Mode** to the appropriate value (`autonomous` or `agentUser`).

---

## Decision matrix

| Need                                                       | Pick              |
| ---------------------------------------------------------- | ----------------- |
| Existing Azure Bot, don't want to migrate                  | Classic Bot       |
| Want Agent 365 identity, simple workload, client secret OK | Agent 365 Inline  |
| Need agent-user tokens (agent as M365 user)                | Agent 365 Sidecar |
| Cert-based auth for production                             | Agent 365 Sidecar |
| M2 Graph/MCP operations coming                             | Agent 365 Sidecar |

---

## See also

- [obo-guide.md](obo-guide.md) — why `interactiveOBO` is M2-only and what it unlocks
- [licensing-notes.md](licensing-notes.md) — Frontier preview vs GA cost implications per identity mode
- [examples/sidecar/README.md](../examples/sidecar/README.md) — sidecar deployment (dev + swarm)
