import { configWithoutCloudSupport } from '@n8n/node-cli/eslint';
import { globalIgnores } from 'eslint/config';

// This package intentionally depends on external runtime packages
// (@microsoft/agents-hosting, axios, jsonwebtoken, jwks-rsa, adaptivecards-templating).
// It cannot run on the n8n Cloud sandbox, so we use the lint config variant
// that drops the cloud-specific no-restricted-imports rule.
export default [
	globalIgnores(['vitest.config.ts', 'tests/**', 'coverage/**']),
	...configWithoutCloudSupport,
	// n8n-nodes-base/node-filename-against-convention expects every .ts under
	// nodes/XXX/ to be named XXX.node.ts. The M0B refactor (Task 3+) uses the
	// multi-file node pattern established by n8n-nodes-base itself (e.g.
	// Slack/V2 with actions/versionDescription.ts), so scope that rule off
	// for sub-tree files.
	{
		files: ['nodes/*/actions/**/*.ts', 'nodes/*/descriptions/**/*.ts'],
		rules: {
			'n8n-nodes-base/node-filename-against-convention': 'off',
		},
	},
];
