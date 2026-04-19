import type { IExecuteFunctions, IDataObject } from 'n8n-workflow';

/** STUB — real implementation ships in Task 15. */
export async function executeSerializeSave(
	this: IExecuteFunctions,
	itemIndex: number,
): Promise<IDataObject[]> {
	throw new Error(
		`M365ConversationRef Serialize/save (item ${itemIndex}): not yet implemented (Task 15 scaffolding)`,
	);
}
