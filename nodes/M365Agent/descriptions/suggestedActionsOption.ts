import type { INodeProperties } from 'n8n-workflow';

/**
 * Suggested Actions option — nested inside the Message resource's Options
 * collection. Produces `activity.suggestedActions` at execute time.
 *
 * Shape: the outer node is a `collection` wrapper holding two things —
 *   1. `chips`  → `fixedCollection` of repeating type/title/value/displayText rows
 *                 (the actual tap-chips rendered by the channel).
 *   2. `to`     → optional `string[]` recipient override that lives literally
 *                 under Suggested Actions per spec §4 line 113 / §8 line 220.
 *
 * `to` is the proactive-path escape hatch: when there is no inbound clicker to
 * auto-populate, the workflow supplies explicit user IDs (typically via an
 * expression like
 * `={{ $('M365 Conversation Ref').item.json.conversationReference.user.id }}`).
 * Semantics (see `shared/applyMessageOptionsToActivity.ts`):
 *   - Field omitted (default) → auto-populate from envelope.
 *   - Populated array         → wins over auto-populate.
 *   - Explicit empty array    → broadcast; auto-populate skipped.
 *
 * Chip type options are alphabetical per
 * `eslint-plugin-n8n-nodes-base/node-param-options-type-unsorted-items`.
 * Display Text is only meaningful for `messageBack` (manifest §5 displayOptions).
 */
export const suggestedActionsOption: INodeProperties = {
	displayName: 'Suggested Actions',
	name: 'suggestedActions',
	type: 'collection',
	placeholder: 'Add suggested actions',
	default: {},
	options: [
		{
			displayName: 'Chips',
			name: 'chips',
			type: 'fixedCollection',
			typeOptions: { multipleValues: true, sortable: true },
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
							description:
								'Action payload — see the Type dropdown for what each action expects',
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
		},
		{
			displayName: 'To',
			name: 'to',
			type: 'string',
			typeOptions: { multipleValues: true },
			default: [],
			placeholder: 'e.g. 29:1abcDefG...',
			description:
				'Override the recipient list on the Suggested Actions chips. By default the chips target the inbound clicker (auto-populated from the envelope); set this to explicit user IDs for proactive sends where there is no clicker. Leave empty to broadcast to the whole conversation.',
		},
	],
};
