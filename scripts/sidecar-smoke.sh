#!/usr/bin/env bash
# Sidecar smoke tests — Phase 0 of Agent 365 support.
# Reads env from ./scripts/.env.smoke (git-ignored).

set -euo pipefail
: "${SIDECAR_URL:=http://localhost:5000}"
: "${BLUEPRINT_APP_ID:?required}"
: "${TEST_AGENT_UPN:=}"
: "${TEST_DELEGATED_USER_TOKEN:=}"
: "${TEST_INBOUND_ACTIVITY_BEARER:=}"
: "${TRANSCRIPT:=tests/fixtures/sidecar-smoke-transcript-phase0.json}"

mkdir -p "$(dirname "$TRANSCRIPT")"
echo "[" > "$TRANSCRIPT"
first=1

record() {
  local name="$1" request="$2" response="$3" status="$4" err="$5"
  [[ $first -eq 1 ]] && first=0 || echo "," >> "$TRANSCRIPT"
  # status is kept as a STRING so 'ERR' on transport failure does not break
  # the transcript write. Success cases still hold numeric strings like "200".
  jq -n --arg n "$name" --arg r "$request" --arg b "$response" --arg s "$status" --arg e "$err" \
    '{test: $n, request: $r, status: $s, transport_error: $e, response_preview: ($b[0:800])}' \
    >> "$TRANSCRIPT"
}

run() {
  local name="$1"; shift
  echo "=== $name ==="
  local status="ERR" body="" err=""
  # -f makes curl exit non-zero on HTTP 4xx/5xx; we want those to land in
  # the transcript too, so use -s -o and -w separately and check exit code.
  if out=$(curl -sS -o /tmp/smoke-body.txt -w "%{http_code}" "$@" 2>/tmp/smoke-err.txt); then
    status="$out"
    body=$(cat /tmp/smoke-body.txt)
  else
    err=$(cat /tmp/smoke-err.txt 2>/dev/null)
    body=$(cat /tmp/smoke-body.txt 2>/dev/null)
  fi
  echo "Status: $status${err:+ (transport error: $err)}"
  echo "Body: $body" | head -c 500; echo
  record "$name" "$*" "$body" "$status" "$err"
}

# 1. /healthz
run "healthz" "$SIDECAR_URL/healthz"

# 2. Autonomous — Messaging Bot API
run "autonomous-msgbot" \
  "$SIDECAR_URL/AuthorizationHeaderUnauthenticated/MessagingBotApi?AgentIdentity=$BLUEPRINT_APP_ID"

# 3. Agent-user — requires UPN
if [[ -n "$TEST_AGENT_UPN" ]]; then
  run "agent-user-msgbot" \
    "$SIDECAR_URL/AuthorizationHeaderUnauthenticated/MessagingBotApi?AgentIdentity=$BLUEPRINT_APP_ID&AgentUsername=$TEST_AGENT_UPN"
else
  echo "Skipping agent-user: TEST_AGENT_UPN not set"
fi

# 4. OBO against Graph — M2-preview validation.
# Uses a DELEGATED USER TOKEN (az CLI / MSAL device-code sign-in to the
# Blueprint's exposed api://<blueprintAppId>/access_agent scope).
# This is NOT the same token class as the inbound activity token
# Microsoft sends to your webhook — keep the two distinct.
if [[ -n "$TEST_DELEGATED_USER_TOKEN" ]]; then
  run "obo-graph" \
    -H "Authorization: Bearer $TEST_DELEGATED_USER_TOKEN" \
    "$SIDECAR_URL/AuthorizationHeader/Graph?AgentIdentity=$BLUEPRINT_APP_ID"
else
  echo "Skipping OBO: TEST_DELEGATED_USER_TOKEN not set"
fi

# 5a. /Validate (success) — capture a REAL inbound bearer from an actual Teams
# webhook request to your trigger, then pass it here. Do NOT use the delegated
# user token from step 4 — that has a different audience and will not exercise
# the inbound-activity-token validation path the trigger actually runs.
if [[ -n "$TEST_INBOUND_ACTIVITY_BEARER" ]]; then
  run "validate-inbound-ok" \
    -H "Authorization: Bearer $TEST_INBOUND_ACTIVITY_BEARER" \
    "$SIDECAR_URL/Validate"
else
  echo "Skipping /Validate (success): TEST_INBOUND_ACTIVITY_BEARER not set"
  echo "  To capture: temporarily log req.headers.authorization in the trigger,"
  echo "  send one message via Teams, copy the bearer out, set it here."
fi

# 5b. /Validate (failure) — spec §9.1 test 5 requires a known-bad bearer returning
# 401 so the router's inbound-validation error path is grounded in a real response
# shape, not inferred. A literal malformed JWT is always rejected by Entra/the
# sidecar validator; no credential material needed.
: "${TEST_BAD_INBOUND_BEARER:=not.a.valid.jwt}"
run "validate-inbound-fail" \
  -H "Authorization: Bearer $TEST_BAD_INBOUND_BEARER" \
  "$SIDECAR_URL/Validate"

# 6. Unknown downstream → 404 shape
run "unknown-downstream" \
  "$SIDECAR_URL/AuthorizationHeaderUnauthenticated/NotConfigured?AgentIdentity=$BLUEPRINT_APP_ID"

# 7. Scope syntax variants — record which works (requires reconfiguring sidecar env between runs)
echo "=== SCOPE-VARIANT NOTE ==="
echo "Reconfigure sidecar with DownstreamApis__MessagingBotApi__Scopes__0 set to each of:"
echo "  - a6c6ce43-d3ed-40ae-a6d2-4f9c028e0355/.default"
echo "  - api://a6c6ce43-d3ed-40ae-a6d2-4f9c028e0355/.default"
echo "  - a6c6ce43-d3ed-40ae-a6d2-4f9c028e0355/AgentData.ReadWrite"
echo "Re-run test 2 for each. Record working form in Phase 0 checklist below."

echo "]" >> "$TRANSCRIPT"
echo
echo "Transcript: $TRANSCRIPT"
