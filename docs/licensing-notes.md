# Licensing Notes

Cost and license requirements for each Agent 365 identity mode, across the Frontier preview period and GA.

> **Frontier preview period ends 2026-05-01.** Licensing terms may change at GA. Check Microsoft's official pricing page for current list prices.

---

## Frontier preview (until 2026-05-01)

| Identity mode    | License required                                                                      | Notes                                                |
| ---------------- | ------------------------------------------------------------------------------------- | ---------------------------------------------------- |
| `autonomous`     | 1 **Agent 365 Frontier** license per agent instance                                   | 25-agent limit per tenant during preview             |
| `agentUser`      | Frontier license + real M365 seat (E5 + Teams Enterprise) on the agent's user account | The agent account must be a real, licensed M365 user |
| `interactiveOBO` | No extra cost beyond the requesting user's existing license                           | M2-only; not UI-exposed in M1                        |

The 25-agent-per-tenant cap during Frontier applies to `autonomous` instances. `agentUser` accounts consume a full user seat and do not count against the agent cap separately — they are treated as regular M365 users.

---

## GA (from 2026-05-01)

| Identity mode    | License required                                                                                     | Notes                                                                        |
| ---------------- | ---------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| `autonomous`     | Covered under the requesting user's **Agent 365** or **M365 E7** license when acting on their behalf | No separate per-agent license if the agent acts on behalf of a licensed user |
| `agentUser`      | Real M365 seat still required: **E5 + Teams Enterprise + Agent 365** (or **M365 E7**)                | No change from Frontier — a real user account with a full suite license      |
| `interactiveOBO` | No extra cost                                                                                        | User's license covers the Graph/MCP call                                     |

---

## WARNING: Agent User Account is a real license

Choosing `agentUser` identity mode means your agent needs its own M365 user account with full suite licensing. This is not a reduced or agent-only SKU.

**What this entails:**

- The agent gets its own mailbox, its own Teams presence, and appears as a user in your tenant's directory.
- The account requires **Microsoft 365 E5 + Teams Enterprise + Agent 365** (or **Microsoft 365 E7** which bundles all three).
- At GA list prices (as of the 2026-03-09 announcement): approximately **$57/month per agent account**.
- During Frontier preview the Frontier license is also required on top of the M365 seat.

**Use `agentUser` only when the agent genuinely needs to BE a user** — for example:

- The agent participates in Teams chats or channels as a named participant.
- The agent needs its own calendar or mailbox.
- Users need to @mention a named M365 identity in Teams.

For automation that simply sends messages or cards on behalf of a workflow, `autonomous` mode is almost always the right choice.

---

## Safer defaults

| Goal                                                           | Recommended mode                                                                     |
| -------------------------------------------------------------- | ------------------------------------------------------------------------------------ |
| Automate message/card sending without user-context Graph calls | `autonomous`                                                                         |
| Agent needs M365 presence (Teams participant, mailbox)         | `agentUser` — accept the license cost                                                |
| Graph/MCP calls as the inbound user (M2)                       | `interactiveOBO` — wait for M2, no extra cost                                        |
| Unsure                                                         | Start with `autonomous`; switch only when a concrete requirement demands `agentUser` |

---

## See also

- [auth-guide.md](auth-guide.md) — credential and transport setup
- [obo-guide.md](obo-guide.md) — why OBO has no extra license cost and what it unlocks in M2
