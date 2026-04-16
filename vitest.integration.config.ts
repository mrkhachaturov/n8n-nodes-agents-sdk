import { defineConfig } from 'vitest/config';

// Integration test config — hits real Azure / network endpoints.
// Gated on M365_TEST_CLIENT_ID / _SECRET / _TENANT_ID at the test-file level
// (tests use `describe.skip` when env vars are missing), so this config is safe
// to run even without credentials — it just produces zero executed assertions.
export default defineConfig({
	test: {
		globals: true,
		environment: 'node',
		include: ['tests/integration/**/*.test.ts'],
		testTimeout: 30_000,
	},
});
