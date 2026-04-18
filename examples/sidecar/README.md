# Sidecar deployment guide

Deploy the Microsoft Entra Agent ID auth sidecar alongside n8n so n8n-nodes-m365-agents can acquire tokens on behalf of Agent Identities.

## Prerequisites

- Docker 24.0+ (or Docker Swarm 24.0+)
- Entra tenant with an Agent Identity Blueprint ([setup guide](https://learn.microsoft.com/entra/agent-id/))
- One of:
  - **Client secret** (dev/testing) — fastest path
  - **Client certificate** (production) — required for cert-based trust

## Env vars

| Var                       | Required | Description                                             |
| ------------------------- | -------- | ------------------------------------------------------- |
| `TENANT_ID`               | Yes      | Entra tenant GUID                                       |
| `BLUEPRINT_APP_ID`        | Yes      | Agent Identity Blueprint appId (NOT the instance appId) |
| `BLUEPRINT_CLIENT_SECRET` | Dev only | Client secret for the Blueprint app                     |

## Dev (Docker Compose)

```bash
cp .env.example .env
# edit .env with your values
docker compose up -d
curl http://localhost:8080/healthz   # expect 200
```

## Production (Docker Swarm + cert-based auth)

1. Create the cert secret (PFX format, passwordless):
   ```bash
   docker secret create entra_cert_pfx /path/to/cert.pfx
   ```
2. Deploy with both compose files:
   ```bash
   docker stack deploy -c docker-compose.yml -c docker-compose.swarm.yml entra-auth
   ```
3. Verify the service:
   ```bash
   docker service ps entra-auth_entra-auth-sidecar
   ```

## Security

- **NEVER expose the sidecar to the public internet** — it only validates the tenant boundary, not network boundaries.
- Put the sidecar on an internal overlay network shared with n8n only.
- Use Docker secrets (not env vars) for cert material in production.
- `ASPNETCORE_ENVIRONMENT=Development` bypasses the default host-filter (required for overlay-network deployments).

## Troubleshooting

| Symptom                                 | Likely cause                                | Fix                                                             |
| --------------------------------------- | ------------------------------------------- | --------------------------------------------------------------- |
| `sidecar 400: AgentIdentity required`   | Missing AgentIdentity query param           | Check credential's `agentInstanceAppId` is set                  |
| `sidecar 401: unauthorized` on outbound | Wrong Blueprint secret/cert                 | Verify Blueprint app credentials in Entra portal                |
| `sidecar 404: unknown downstream`       | DownstreamApi name not configured           | Check `DownstreamApis__<name>__*` env vars                      |
| `TypeError: Failed to parse URL` in n8n | `sidecarUrl` credential field empty         | Set the sidecar URL in the n8n M365 Agent 365 API credential    |
| Sidecar won't start on distroless       | Health check configured on distroless image | Distroless has no shell — disable healthcheck per swarm compose |

## Scope syntax (Phase 0 verified)

The Messaging Bot API scope uses the bare GUID form:

```
a6c6ce43-d3ed-40ae-a6d2-4f9c028e0355/.default
```

NOT `api://a6c6ce43-.../.default`.

Microsoft Graph uses the full URL form:

```
https://graph.microsoft.com/.default
```

## Port

The sidecar listens on **port 8080** (baked into the distroless image via `ASPNETCORE_HTTP_PORTS=8080`). This is a change from the older documentation that showed port 5000.
