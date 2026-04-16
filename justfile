# n8n-nodes-m365-agents — development and deploy tasks.
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
	rsync -av --delete dist/ {{SWARM_HOST}}:{{CEPHFS_PATH}}/n8n-nodes-m365-agents/dist/
	ssh {{SWARM_HOST}} "docker service update --force {{N8N_SERVICE}}"

# Run unit tests (offline — default scope, excludes tests/integration).
test:
	npm test

# Run integration tests against real Azure. Uses `mise exec` so the
# parent repo's 1Password-managed env (~/.op-env/at-m365bot.env) is loaded,
# populating M365_TEST_CLIENT_ID / _SECRET / _TENANT_ID. Tests skip cleanly
# if those vars are not set.
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
