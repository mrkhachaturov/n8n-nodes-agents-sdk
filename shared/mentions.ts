import type { Activity } from '@microsoft/agents-activity';

/**
 * Minimal mention input vocabulary the M365 Agent node's Mentions option
 * exposes. User covers individual @-mentions (most common); Everyone is the
 * Teams team-wide announcement mention backed by the magic ID `29:allchannel`.
 * Channel / Team mentions are deferred to a later milestone — the magic IDs
 * require the inbound activity's channelData, which belongs in M2's Teams
 * channel-info work.
 */
export type MentionKind = 'user' | 'everyone';

export interface MentionInput {
	type: MentionKind;
	/** Required for `user`; ignored (auto-set to `29:allchannel`) for `everyone`. */
	id?: string;
	/** Display name — required. What shows inside the `<at>...</at>` token. */
	name: string;
}

export interface MentionEntity {
	type: 'mention';
	mentioned: { id: string; name: string };
	text: string;
}

export interface MentionResult {
	entities: MentionEntity[];
	textTokens: string[];
}

const EVERYONE_ID = '29:allchannel';

function assert(cond: unknown, message: string): asserts cond {
	if (!cond) {
		throw new Error(message);
	}
}

export function buildMentions(inputs: MentionInput[]): MentionResult {
	const entities: MentionEntity[] = [];
	const textTokens: string[] = [];

	for (const [i, m] of inputs.entries()) {
		assert(m.name && m.name.trim().length > 0, `Mention #${i + 1}: name is required`);
		const id = m.type === 'everyone' ? EVERYONE_ID : m.id;
		assert(
			id && id.trim().length > 0,
			`Mention #${i + 1}: id is required for type=user (use {{ $json.parsed.userId }} or paste the AAD object id)`,
		);

		const token = `<at>${m.name}</at>`;
		entities.push({ type: 'mention', mentioned: { id, name: m.name }, text: token });
		textTokens.push(token);
	}

	return { entities, textTokens };
}

/**
 * Returns a new activity with mention entities merged in and text tokens
 * prepended. Idempotent on text — if the token already appears anywhere in
 * the existing text, it is not prepended a second time. Never mutates input.
 */
export function applyMentionsToActivity(
	activity: Partial<Activity>,
	inputs: MentionInput[],
): Partial<Activity> {
	if (inputs.length === 0) return activity;

	const { entities, textTokens } = buildMentions(inputs);
	const existingText = activity.text ?? '';

	// Prepend any tokens not already present in the text (idempotent).
	const missingTokens = textTokens.filter((t) => !existingText.includes(t));
	const prefix = missingTokens.length > 0 ? `${missingTokens.join(' ')} ` : '';
	const newText = `${prefix}${existingText}`.trimEnd();

	const existingEntities = Array.isArray(activity.entities) ? activity.entities : [];

	return {
		...activity,
		text: newText.length > 0 ? newText : activity.text,
		entities: [...existingEntities, ...entities] as Activity['entities'],
	};
}
