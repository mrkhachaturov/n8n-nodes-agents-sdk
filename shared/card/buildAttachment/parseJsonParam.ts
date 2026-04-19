import type { IExecuteFunctions } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';

/**
 * Parse a JSON-typed node parameter into an object, with consistent handling of
 * n8n's empty-expression sentinel `'={}'` and bare `''` (both → empty object).
 *
 * Used by every JSON-shaped buildAttachment (Receipt, O365 Connector, Raw
 * Attachment). Extracted to keep empty-input semantics identical across the
 * three sibling resources — a prior divergence let `'={}'` parse as JSON and
 * throw in Receipt/O365 while Raw accepted it.
 */
export function parseJsonParam(
	ctx: IExecuteFunctions,
	itemIndex: number,
	fieldName: string,
	raw: unknown,
): unknown {
	if (raw === undefined || raw === null) return {};
	if (typeof raw !== 'string') return raw;
	const trimmed = raw.trim();
	if (trimmed === '' || trimmed === '={}') return {};
	try {
		return JSON.parse(trimmed);
	} catch (err) {
		throw new NodeOperationError(
			ctx.getNode(),
			`Invalid JSON in ${fieldName}: ${(err as Error).message}`,
			{ itemIndex },
		);
	}
}
