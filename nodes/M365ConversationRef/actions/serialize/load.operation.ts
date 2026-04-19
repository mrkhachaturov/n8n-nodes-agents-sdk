import type { IExecuteFunctions, IDataObject } from 'n8n-workflow';

/** STUB — real implementation ships in Task 16. */
export async function executeSerializeLoad(
	this: IExecuteFunctions,
	itemIndex: number,
): Promise<IDataObject[]> {
	throw new Error(
		`M365ConversationRef Serialize/load (item ${itemIndex}): not yet implemented (Task 16 scaffolding)`,
	);
}
