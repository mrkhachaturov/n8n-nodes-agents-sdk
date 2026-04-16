import { defineConfig } from 'vitest/config';

export default defineConfig({
	test: {
		globals: true,
		environment: 'node',
		passWithNoTests: true,
		// Default run excludes `tests/integration/**` so `npm test` stays offline.
		// Use `npm run test:integration` (or `just test-integration`) to run those.
		include: ['tests/**/*.test.ts'],
		exclude: ['**/node_modules/**', 'tests/integration/**'],
		coverage: {
			provider: 'v8',
			reporter: ['text', 'html'],
			include: ['credentials/**', 'nodes/**', 'shared/**'],
			exclude: ['**/*.test.ts', '**/node_modules/**'],
		},
	},
});
