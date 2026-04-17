import type { IExecuteFunctions, INodeType, INodeTypeDescription } from 'n8n-workflow';
import { router } from './actions/router';
import { versionDescription } from './actions/versionDescription';

export class M365Agent implements INodeType {
	description: INodeTypeDescription = {
		...versionDescription,
		// Repeated here so @n8n/community-nodes/icon-validation can resolve the
		// icon statically from the class's description ObjectExpression. The rule
		// requires description to be an inline object literal — it cannot follow
		// identifier references. versionDescription.icon carries the same value.
		icon: { light: 'file:../../icons/m365.svg', dark: 'file:../../icons/m365.dark.svg' },
		// Explicitly false — this node performs side effects (sends/updates/deletes
		// Activities to Azure Bot Service) and must not be invocable from an AI
		// Agent's reasoning loop (manifest §10.7). Stated explicitly here to satisfy
		// @n8n/community-nodes/node-usable-as-tool (rule requires the property present).
		usableAsTool: false,
	};

	async execute(this: IExecuteFunctions) {
		return await router.call(this);
	}
}
