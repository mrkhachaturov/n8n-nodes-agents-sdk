import type { INodeProperties } from 'n8n-workflow';

export class RawActivityOverrideError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'RawActivityOverrideError';
	}
}

/**
 * Replace the constructed Activity body with a user-supplied JSON object when
 * `rawActivityOverride` is non-empty. Empty/undefined/null → return constructed
 * unchanged. Invalid JSON or non-object values throw RawActivityOverrideError.
 *
 * Per Charter R2 — this is the Option form of the Raw Activity escape hatch.
 * It applies to Message send-style operations (send/reply/update/replyInThread)
 * and to every Card resource's send and update operations.
 */
export function applyRawActivityOverride(
	constructed: Record<string, unknown>,
	override: string | null | undefined,
): Record<string, unknown> {
	if (!override || override.trim() === '') return constructed;

	let parsed: unknown;
	try {
		parsed = JSON.parse(override);
	} catch (err) {
		const e = err as Error;
		throw new RawActivityOverrideError(
			`rawActivityOverride: invalid JSON — ${e.message}`,
		);
	}
	if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) {
		throw new RawActivityOverrideError(
			`rawActivityOverride: expected JSON object at top level (got ${Array.isArray(parsed) ? 'array' : typeof parsed})`,
		);
	}
	return parsed as Record<string, unknown>;
}

/**
 * Reusable INodeProperties row to drop into any operation's Options collection.
 * Keep name and placeholder stable — the Release Surface Coverage Table
 * references this capability as G023-reshape.
 */
export const rawActivityOverrideField: INodeProperties = {
	displayName: 'Raw Activity Override',
	name: 'rawActivityOverride',
	type: 'string',
	typeOptions: { rows: 4 },
	default: '',
	placeholder: '{"type":"message","text":"…","channelData":{…}}',
	description:
		'Optional: replace the constructed Activity body with raw JSON. Empty = existing behavior. Used for protocol fields this UI does not expose (custom channelData, experimental Activity shapes). Must parse to a JSON object.',
};
