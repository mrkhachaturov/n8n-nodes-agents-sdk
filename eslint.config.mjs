import { config } from '@n8n/node-cli/eslint';
import { globalIgnores } from 'eslint/config';

export default [
	globalIgnores(['vitest.config.ts', 'tests/**', 'coverage/**']),
	...config,
];
