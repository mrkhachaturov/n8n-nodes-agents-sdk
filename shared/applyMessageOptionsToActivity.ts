import type { Activity } from '@microsoft/agents-activity';
import type { IDataObject } from 'n8n-workflow';
import { applyMentionsToActivity, type MentionInput } from './mentions';
import { applySuggestedActionsToActivity, type SuggestedActionInput } from './suggestedActions';

/**
 * Pure transform: take the activity that an operation just constructed and
 * the n8n-resolved `options` object, apply the M365 Agent's Message-resource
 * options (mentions, suggestedActions) to it, and return the resulting
 * activity. Throws plain Error on shape problems — call sites wrap in
 * NodeOperationError so the error carries itemIndex per manifest §11.
 *
 * `suggestedActionsTo` scopes the chips to specific recipients. For replies
 * to a user's button click, pass `[userId]` so Teams channel scope renders
 * the chips for that clicker. Empty (default) = broadcast to everyone.
 *
 * The Suggested Actions option is a `collection` (spec §4/§8) containing:
 *   - `chips`: fixedCollection of chip rows at `options.suggestedActions.chips.values`
 *   - `to`:    optional `string[]` at `options.suggestedActions.to` that wins
 *              over the envelope-derived auto-populate.
 *
 * `to` semantics (proactive-path escape hatch):
 *   - `undefined` (not set) → fall back to `suggestedActionsTo` (auto-populate).
 *   - non-empty array       → wins over `suggestedActionsTo`.
 *   - explicit empty array  → explicit broadcast; auto-populate is skipped.
 *   - anything else         → throws (array of strings is the only valid shape).
 */
export function applyMessageOptionsToActivity(
	activity: Partial<Activity>,
	options: IDataObject,
	suggestedActionsTo: string[] = [],
): Partial<Activity> {
	let next = activity;

	const mentions = (options.mentions as { values?: MentionInput[] } | undefined)?.values;
	if (mentions && mentions.length > 0) {
		next = applyMentionsToActivity(next, mentions);
	}

	const suggested = options.suggestedActions as
		| { chips?: { values?: SuggestedActionInput[] }; to?: unknown }
		| undefined;
	const actions = suggested?.chips?.values;
	if (actions && actions.length > 0) {
		const override = suggested?.to;
		let resolvedTo: string[];
		if (override === undefined) {
			resolvedTo = suggestedActionsTo;
		} else {
			if (!Array.isArray(override)) {
				throw new Error(
					`suggestedActions.to: expected an array of user IDs, got ${typeof override}`,
				);
			}
			for (const [i, entry] of override.entries()) {
				if (typeof entry !== 'string') {
					throw new Error(
						`suggestedActions.to[${i}]: expected a string user ID, got ${typeof entry}`,
					);
				}
			}
			// Empty array = explicit broadcast — skip auto-populate entirely.
			resolvedTo = override as string[];
		}
		next = applySuggestedActionsToActivity(next, actions, resolvedTo);
	}

	return next;
}
