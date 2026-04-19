/* eslint-disable @n8n/community-nodes/node-usable-as-tool --
 * This node serializes conversation references with optional inbound bearer
 * tokens — sensitive auth material that must never be exposed to an AI Agent's
 * reasoning loop. The lint rule's autofix would add `usableAsTool: true`, which
 * is the wrong default for nodes handling bearer tokens. n8n-workflow's type
 * only accepts `true | UsableAsToolDescription | undefined` so the idiomatic
 * "off" is omission (versionDescription omits the field). Suppressing this
 * rule keeps that intent explicit at the class level.
 */
import type {
	IExecuteFunctions,
	INodeType,
	INodeTypeDescription,
	INodeExecutionData,
} from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { versionDescription } from './actions/versionDescription';
import { executeSerializeSave } from './actions/serialize/save.operation';
import { executeSerializeLoad } from './actions/serialize/load.operation';

export class M365ConversationRef implements INodeType {
	// Inline spread so @n8n/community-nodes/icon-validation can read `icon`
	// from an ObjectExpression (the rule doesn't follow identifier references).
	// `icon` here is a re-statement of versionDescription.icon — the same value,
	// required by the lint rule to be statically resolvable.
	description: INodeTypeDescription = {
		...versionDescription,
		icon: { light: 'file:../../icons/m365.svg', dark: 'file:../../icons/m365.dark.svg' },
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		for (let i = 0; i < items.length; i++) {
			const resource = this.getNodeParameter('resource', i) as string;
			const operation = this.getNodeParameter('operation', i) as string;

			if (resource !== 'serialize') {
				throw new NodeOperationError(this.getNode(), `Unknown resource: ${resource}`, {
					itemIndex: i,
				});
			}

			let output;
			if (operation === 'save') output = await executeSerializeSave.call(this, i);
			else if (operation === 'load') output = await executeSerializeLoad.call(this, i);
			else
				throw new NodeOperationError(this.getNode(), `Unknown operation: ${operation}`, {
					itemIndex: i,
				});

			for (const row of output) returnData.push({ json: row });
		}

		return [returnData];
	}
}
