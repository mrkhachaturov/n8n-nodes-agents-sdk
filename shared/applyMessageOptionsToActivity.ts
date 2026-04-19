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
 * G022 — `options.suggestedActionsToOverride` lets the workflow explicitly
 * set the recipient list, winning over the envelope-derived auto-populate.
 * This is the proactive-path escape hatch: there is no inbound clicker to
 * target, so users pipe a resolved user ID (for example from an earlier
 * `M365 Conversation Ref` node) into this field. Semantics:
 *   - `undefined` (not set) → fall back to `suggestedActionsTo` (existing).
 *   - non-empty array      → wins over `suggestedActionsTo`.
 *   - explicit empty array → explicit broadcast; auto-populate is skipped.
 *   - anything else        → throws (array of strings is the only valid shape).
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

	const actions = (options.suggestedActions as { values?: SuggestedActionInput[] } | undefined)
		?.values;
	if (actions && actions.length > 0) {
		const override = options.suggestedActionsToOverride;
		let resolvedTo: string[];
		if (override === undefined) {
			resolvedTo = suggestedActionsTo;
		} else {
			if (!Array.isArray(override)) {
				throw new Error(
					`suggestedActionsToOverride: expected an array of user IDs, got ${typeof override}`,
				);
			}
			for (const [i, entry] of override.entries()) {
				if (typeof entry !== 'string') {
					throw new Error(
						`suggestedActionsToOverride[${i}]: expected a string user ID, got ${typeof entry}`,
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
