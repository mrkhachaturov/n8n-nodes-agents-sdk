import type { Activity } from '@microsoft/agents-activity';

/** The four action kinds useful for quick-reply bot UX. */
export type SuggestedActionKind = 'imBack' | 'messageBack' | 'postBack' | 'openUrl';

export interface SuggestedActionInput {
	type: SuggestedActionKind;
	/** Chip label — required. */
	title: string;
	/** Action payload — message text for imBack/postBack, hidden payload for messageBack, URL for openUrl. */
	value: string;
	/** Only for messageBack — the user-visible text shown in the transcript after the click. */
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

const VALID_TYPES: readonly SuggestedActionKind[] = ['imBack', 'messageBack', 'postBack', 'openUrl'];

function assert(cond: unknown, message: string): asserts cond {
	if (!cond) {
		throw new Error(message);
	}
}

export function buildSuggestedActions(inputs: SuggestedActionInput[]): SuggestedActionsShape {
	const actions: CardActionShape[] = [];
	for (const [i, a] of inputs.entries()) {
		assert(
			VALID_TYPES.includes(a.type),
			`Suggested action #${i + 1}: unknown type '${a.type}' (expected one of ${VALID_TYPES.join(', ')})`,
		);
		assert(a.title && a.title.trim().length > 0, `Suggested action #${i + 1}: title is required`);
		const out: CardActionShape = { type: a.type, title: a.title, value: a.value };
		if (a.type === 'messageBack' && a.displayText) {
			out.displayText = a.displayText;
		}
		actions.push(out);
	}
	return { to: [], actions };
}

/**
 * Returns a new activity with `suggestedActions` set. Empty inputs → unchanged.
 * Never mutates input.
 */
export function applySuggestedActionsToActivity(
	activity: Partial<Activity>,
	inputs: SuggestedActionInput[],
): Partial<Activity> {
	if (inputs.length === 0) return activity;
	const suggestedActions = buildSuggestedActions(inputs);
	return { ...activity, suggestedActions: suggestedActions as Activity['suggestedActions'] };
}
