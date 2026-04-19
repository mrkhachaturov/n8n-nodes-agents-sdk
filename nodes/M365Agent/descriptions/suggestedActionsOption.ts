import type { INodeProperties } from 'n8n-workflow';

/**
 * Suggested Actions property — nested inside the Message resource's Options
 * collection. Produces `activity.suggestedActions` at execute time.
 *
 * Type options are alphabetical per `eslint-plugin-n8n-nodes-base/
 * node-param-options-type-unsorted-items`. Display Text only meaningful for
 * messageBack (manifest §5 displayOptions).
 *
 * G022 — Pair this with `suggestedActionsToOverrideOption` below to override
 * the auto-populated `to` list for proactive sends (no inbound clicker).
 */
export const suggestedActionsOption: INodeProperties = {
	displayName: 'Suggested Actions',
	name: 'suggestedActions',
	type: 'fixedCollection',
	typeOptions: { multipleValues: true, sortable: true },
	placeholder: 'Add suggested action',
	default: {},
	description:
		'Quick-reply chips shown under the message. Up to ten. Teams renders them as tappable buttons.',
	options: [
		{
			name: 'values',
			displayName: 'Suggested Actions',
			values: [
				{
					displayName: 'Type',
					name: 'type',
					type: 'options',
					options: [
						{ name: 'Call (Phone Call)', value: 'call' },
						{ name: 'downloadFile (Download File)', value: 'downloadFile' },
						{ name: 'imBack (Send Message)', value: 'imBack' },
						{ name: 'messageBack (Send With Display Text)', value: 'messageBack' },
						{ name: 'openApp (Open App)', value: 'openApp' },
						{ name: 'openUrl (Open Link)', value: 'openUrl' },
						{ name: 'playAudio (Play Audio)', value: 'playAudio' },
						{ name: 'playVideo (Play Video)', value: 'playVideo' },
						{ name: 'postBack (Send Hidden)', value: 'postBack' },
						{ name: 'showImage (Show Image)', value: 'showImage' },
						{ name: 'Signin (Sign-In)', value: 'signin' },
					],
					default: 'imBack',
				},
				{
					displayName: 'Title',
					name: 'title',
					type: 'string',
					default: '',
					required: true,
					description: 'Label shown on the chip',
				},
				{
					displayName: 'Value',
					name: 'value',
					type: 'string',
					default: '',
					required: true,
					description: 'Action payload — see the Type dropdown for what each action expects',
				},
				{
					displayName: 'Display Text',
					name: 'displayText',
					type: 'string',
					default: '',
					description:
						'User-visible text after click — messageBack only. Leave blank for a silent submit.',
					displayOptions: { show: { type: ['messageBack'] } },
				},
			],
		},
	],
};

/**
 * Suggested Actions — To (Override) — G022.
 *
 * Peer to `suggestedActionsOption` inside the Message Options collection.
 * Overrides the auto-populated recipient list on `activity.suggestedActions.to`.
 *
 * By default the executor auto-populates `to` with the inbound clicker's user
 * ID (so Teams channel scope renders the chips for that person). Proactive
 * flows have no inbound clicker — use this field to pass an explicit user ID
 * (typically via an expression like
 * `={{ $('M365 Conversation Ref').item.json.conversationReference.user.id }}`).
 * Leave empty for broadcast to the whole conversation.
 *
 * Semantics (see `shared/applyMessageOptionsToActivity.ts`):
 *   - Field omitted (default) → auto-populate from envelope.
 *   - Populated array         → wins over auto-populate.
 *   - Explicit empty array    → broadcast; auto-populate skipped.
 */
export const suggestedActionsToOverrideOption: INodeProperties = {
	displayName: 'Suggested Actions — To (Override)',
	name: 'suggestedActionsToOverride',
	type: 'string',
	typeOptions: { multipleValues: true },
	default: [],
	placeholder: 'e.g. 29:1abcDefG...',
	description:
		'Override the recipient list on Suggested Actions chips. By default the chips target the inbound clicker (auto-populated from the envelope); set this to explicit user IDs for proactive sends where there is no clicker. Leave empty to broadcast to the whole conversation.',
};
