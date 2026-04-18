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
 */
export function applyMessageOptionsToActivity(
	activity: Partial<Activity>,
	options: IDataObject,
): Partial<Activity> {
	let next = activity;

	const mentions = (options.mentions as { values?: MentionInput[] } | undefined)?.values;
	if (mentions && mentions.length > 0) {
		next = applyMentionsToActivity(next, mentions);
	}

	const actions = (options.suggestedActions as { values?: SuggestedActionInput[] } | undefined)
		?.values;
	if (actions && actions.length > 0) {
		next = applySuggestedActionsToActivity(next, actions);
	}

	return next;
}
