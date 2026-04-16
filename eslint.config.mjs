import { configWithoutCloudSupport } from '@n8n/node-cli/eslint';
import { globalIgnores } from 'eslint/config';

// This package intentionally depends on external runtime packages
// (@microsoft/agents-hosting, axios, jsonwebtoken, jwks-rsa, adaptivecards-templating).
// It cannot run on the n8n Cloud sandbox, so we use the lint config variant
// that drops the cloud-specific no-restricted-imports rule.
export default [
	globalIgnores(['vitest.config.ts', 'tests/**', 'coverage/**']),
	...configWithoutCloudSupport,
];
