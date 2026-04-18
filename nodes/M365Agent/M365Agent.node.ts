/* eslint-disable @n8n/community-nodes/node-usable-as-tool --
 * This node performs side effects (sends / updates / deletes Activities to
 * Azure Bot Service) and must NOT be invocable from an AI Agent's reasoning
 * loop (manifest §10.7). The lint rule's autofix would add `usableAsTool: true`,
 * which is the wrong default for side-effecting nodes. n8n-workflow's type
 * only accepts `true | UsableAsToolDescription | undefined` so the idiomatic
 * "off" is omission. Suppressing this rule keeps that intent explicit.
 */
import type { IExecuteFunctions, INodeType, INodeTypeDescription } from 'n8n-workflow';
import { router } from './actions/router';
import { versionDescription } from './actions/versionDescription';
import { agent365CredentialTest } from '../../shared/auth/credentialTest';

export class M365Agent implements INodeType {
	// Inline spread so @n8n/community-nodes/icon-validation can read `icon`
	// from an ObjectExpression (the rule doesn't follow identifier references).
	// `icon` here is a re-statement of versionDescription.icon — the same value,
	// required by the lint rule to be statically resolvable.
	description: INodeTypeDescription = {
		...versionDescription,
		icon: { light: 'file:../../icons/m365.svg', dark: 'file:../../icons/m365.dark.svg' },
	};

	methods = {
		credentialTest: { agent365CredentialTest },
	};

	async execute(this: IExecuteFunctions) {
		return await router.call(this);
	}
}
