# OBO Guide (M2-preview)

Conceptual explainer for the `interactiveOBO` identity mode. This mode is implemented in the shared auth router but is **not exposed as a UI option** in M1 — it is reserved for M2 Graph/MCP operations.

---

## What is OBO?

On-Behalf-Of (OBO) is an OAuth2 token exchange flow that allows a service to acquire a token on behalf of an authenticated user, preserving the user's identity and delegated permissions downstream.

In the context of this package, OBO means: **the agent calls Microsoft Graph (or a tenant MCP server) as the inbound requesting user**, rather than as the agent's own app identity.

### Comparing identity modes

| Mode             | Who the agent acts as                  | Permissions                                                              | License cost                                                                    |
| ---------------- | -------------------------------------- | ------------------------------------------------------------------------ | ------------------------------------------------------------------------------- |
| `autonomous`     | The agent's own Blueprint app identity | Application permissions granted to the Blueprint                         | 1 Agent 365 license per agent (Frontier) or covered by user's Agent 365/E7 (GA) |
| `agentUser`      | The agent's own M365 user account      | User permissions of the agent's M365 account                             | Real M365 seat required (see [licensing-notes.md](licensing-notes.md))          |
| `interactiveOBO` | The inbound requesting user            | That user's delegated permissions (respects their CA, MFA, scope grants) | No extra cost — rides the requesting user's license                             |

---

## Why Messaging Bot API rejects delegated tokens

The Bot Framework Connector API (`https://smba.trafficmanager.net/...`) is an application-permission API. It authenticates callers using the **Bot Framework token audience** (`a6c6ce43-d3ed-40ae-a6d2-4f9c028e0355`), which is only valid for application-permission tokens issued to the bot's own app registration.

OBO tokens are issued to user audiences — they carry a `scp` (scope) claim instead of `roles`, and their `aud` is the target resource (Graph, a tenant API), not the Bot Framework audience. Presenting an OBO token to the Bot Connector will result in a `401 Unauthorized`.

**Consequence:** OBO is NOT a valid identity mode for Message or Card operations, which write to the Bot Connector. This is why the `interactiveOBO` option is absent from the Message and Card operation UIs in M1. Exposing it there would allow users to construct a configuration that always fails at runtime.

---

## What OBO unlocks in M2

M2 will introduce Graph and MCP operation resources on `M365Agent`. These call Microsoft Graph or tenant-hosted MCP servers rather than the Bot Connector, which means:

- **User-context Graph calls** — read/write files, calendar, mail, or Teams data as the inbound user, subject to their actual permissions and any Conditional Access policies applied to their account.
- **Tenant MCP server calls** — forward a user's bearer token to a tenant-internal MCP server, so the MCP server sees a delegated-permission token and can enforce its own authorization.
- **MFA and CA state respected** — because OBO preserves the inbound authentication context, Conditional Access policies that require MFA or compliant devices are evaluated against the user, not bypassed by the agent.

**Licensing advantage:** The agent does not need its own M365 seat. It rides the requesting user's existing license — if the user has an M365 E5 or E7 plan, the Graph call is covered.

---

## Router infrastructure readiness

The shared auth router (`shared/auth/router.ts`) already supports the `interactiveOBO` transport path. When the sidecar transport is active, the router forwards the inbound bearer token to the sidecar's `/AuthorizationHeader/{api}` endpoint, which performs the OBO exchange and returns an `Authorization: Bearer <obo-token>` header for the downstream API call.

This means:

1. The sidecar is already capable of OBO exchange with no additional changes.
2. M2 work is limited to surfacing the OBO path in the Graph/MCP operation UI and wiring the outbound call to use the returned header.
3. No router or sidecar protocol changes are planned for M2.

---

## Summary

- OBO lets the agent act as the inbound user for Graph/MCP calls — not for messaging.
- The Bot Connector API only accepts application-permission tokens; OBO tokens will always be rejected there.
- The sidecar already performs OBO exchange; M2 exposes it via Graph/MCP operation fields.
- OBO costs the agent nothing extra — the requesting user's license covers it.

---

## See also

- [auth-guide.md](auth-guide.md) — choosing between classic, inline, and sidecar transports
- [licensing-notes.md](licensing-notes.md) — cost breakdown per identity mode
