# n8n-nodes-agents-sdk — development and deploy tasks.
#
# Deploy target: the n8n Swarm service reads its custom-nodes directory from
# `/home/node/.n8n/custom`, which is backed by `${SHARED}/automation/n8n/custom`
# on the host via the volume `${SHARED}/automation/n8n:/home/node/.n8n`.
#
# The following defaults assume:
#   - SWARM_HOST — FQDN of a swarm manager node you can SSH into
#   - CEPHFS_PATH — host-side path to `${SHARED}/automation/n8n/custom`
#   - N8N_SERVICE — Docker Swarm service name (stack `n8n`, service `app`
#                   → `n8n_app`, as seen in
#                   `Infra/Containers/swarm/stacks/automation/n8n/docker-compose.yml`)
#
# Override any of them inline:
#   SWARM_HOST=swarm02.astrateam.net just deploy-dev

SWARM_HOST := env_var_or_default('SWARM_HOST', 'swarm01.astrateam.net')
CEPHFS_PATH := env_var_or_default('CEPHFS_PATH', '/mnt/cephfs/shared/automation/n8n/custom')
N8N_SERVICE := env_var_or_default('N8N_SERVICE', 'n8n_app')

default:
	@just --list

# Build TypeScript to dist/.
build:
	npm run build

# Compile + rsync dist/ to CephFS + force-update the n8n service.
deploy-dev: build
	rsync -av --delete dist/ {{SWARM_HOST}}:{{CEPHFS_PATH}}/n8n-nodes-agents-sdk/dist/
	rsync -av package.json {{SWARM_HOST}}:{{CEPHFS_PATH}}/n8n-nodes-agents-sdk/package.json
	ssh {{SWARM_HOST}} "docker service update --force {{N8N_SERVICE}}"

# Run unit tests (offline — default scope, excludes tests/integration).
test:
	npm test

# Run integration tests against real Azure. Uses `mise exec` so the
# parent repo's 1Password-managed env (~/.op-env/at-m365bot.env) is loaded,
# populating:
#   - M365_TEST_CLIENT_ID / _SECRET / _TENANT_ID    (shared auth — required by every integration test)
#   - M365_TEST_CONVERSATION_ID / _SERVICE_URL       (required additionally by card-send integration tests:
#                                                     card-adaptive.integration.test.ts, card-hero.integration.test.ts)
# Tests whose env vars are missing skip cleanly via `describe.skip`.
test-integration:
	mise exec -- npm run test:integration

# Lint.
lint:
	npm run lint

# Full CI-style check.
check: lint test build
	@echo "All checks green."

# Watch tests during development.
watch:
	npm run test:watch

# Live smoke — requires LIVE_SMOKE=1 and sidecar env vars. See scripts/live-smoke.ts.
test-smoke:
	LIVE_SMOKE=1 node --import tsx scripts/live-smoke.ts
