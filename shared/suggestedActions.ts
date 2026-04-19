import type { Activity } from '@microsoft/agents-activity';

/**
 * All CardAction.type values defined by the M365 Agents SDK / Bot Framework
 * protocol. See `.local/Agents-for-js/packages/agents-activity/src/action/actionTypes.ts`
 * (and the .NET `Microsoft.Agents.Core.Models.ActionTypes` reference). Support
 * varies by channel; the node exposes every value and lets Teams / WebChat /
 * Direct Line render what they render.
 */
export type SuggestedActionKind =
	| 'call'
	| 'downloadFile'
	| 'imBack'
	| 'messageBack'
	| 'openApp'
	| 'openUrl'
	| 'playAudio'
	| 'playVideo'
	| 'postBack'
	| 'showImage'
	| 'signin';

export interface SuggestedActionInput {
	type: SuggestedActionKind;
	/** Chip label — required. */
	title: string;
	/**
	 * Action payload — semantics depend on `type`: message text for
	 * `imBack` / `postBack`, hidden payload for `messageBack`, URL for
	 * `openUrl` / `downloadFile` / `showImage` / `playAudio` / `playVideo`,
	 * `tel:` URI for `call`, app name for `openApp`, OAuth connection name
	 * (or URL) for `signin`.
	 */
	value: string;
	/** Only for `messageBack` — the user-visible text shown in the transcript after the click. Ignored for other action types. */
	displayText?: string;
}

export interface CardActionShape {
	type: SuggestedActionKind;
	title: string;
	value: string;
	displayText?: string;
}

export interface SuggestedActionsShape {
	to: string[];
	actions: CardActionShape[];
}

const VALID_TYPES: readonly SuggestedActionKind[] = [
	'call',
	'downloadFile',
	'imBack',
	'messageBack',
	'openApp',
	'openUrl',
	'playAudio',
	'playVideo',
	'postBack',
	'showImage',
	'signin',
];

function assert(cond: unknown, message: string): asserts cond {
	if (!cond) {
		throw new Error(message);
	}
}

export function buildSuggestedActions(
	inputs: SuggestedActionInput[],
	to: string[] = [],
): SuggestedActionsShape {
	const actions: CardActionShape[] = [];
	for (const [i, a] of inputs.entries()) {
		assert(
			VALID_TYPES.includes(a.type),
			`Suggested action #${i + 1}: unknown type '${a.type}' (expected one of ${VALID_TYPES.join(', ')})`,
		);
		assert(a.title && a.title.trim().length > 0, `Suggested action #${i + 1}: title is required`);
		assert(
			a.value !== undefined && a.value !== null && String(a.value).length > 0,
			`Suggested action #${i + 1}: value is required (action text for imBack/postBack, hidden payload for messageBack, URL for openUrl)`,
		);
		const out: CardActionShape = { type: a.type, title: a.title, value: a.value };
		if (a.type === 'messageBack' && a.displayText) {
			out.displayText = a.displayText;
		}
		actions.push(out);
	}
	return { to, actions };
}

/**
 * Returns a new activity with `suggestedActions` set. Empty inputs → unchanged.
 * `to` scopes who sees the chips — for replies to button clicks, pass the
 * clicking user's id so Teams renders the chips for that person in channel
 * scope. Empty array = broadcast to all recipients (SDK default, fine for
 * proactive sends where there's no specific clicker).
 * Never mutates input.
 */
export function applySuggestedActionsToActivity(
	activity: Partial<Activity>,
	inputs: SuggestedActionInput[],
	to: string[] = [],
): Partial<Activity> {
	if (inputs.length === 0) return activity;
	const suggestedActions = buildSuggestedActions(inputs, to);
	return { ...activity, suggestedActions: suggestedActions as Activity['suggestedActions'] };
}
