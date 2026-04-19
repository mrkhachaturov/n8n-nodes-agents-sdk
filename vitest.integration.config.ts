import { defineConfig } from 'vitest/config';

// Integration test config — hits real Azure / network endpoints.
// Gated at the test-file level via `describe.skip` when required env vars
// are missing, so this config is safe to run without credentials — each
// test file decides whether to execute based on its own env-var set.
//
// Env vars recognized:
//   - M365_TEST_CLIENT_ID / _TENANT_ID / _CLIENT_SECRET    (shared by all integration tests)
//   - M365_TEST_CONVERSATION_ID / _SERVICE_URL             (required additionally by card-send tests:
//                                                           card-adaptive.integration.test.ts,
//                                                           card-hero.integration.test.ts)
//
// Missing vars → the relevant file's `describe` block is skipped. The auth-only
// test (azure.integration.test.ts) runs whenever the three auth vars are set,
// independent of the two conversation vars.
export default defineConfig({
	test: {
		globals: true,
		environment: 'node',
		include: ['tests/integration/**/*.test.ts'],
		testTimeout: 30_000,
	},
});
