import type { INodeProperties } from 'n8n-workflow';

/**
 * Suggested Actions property — nested inside the Message resource's Options
 * collection. Produces `activity.suggestedActions` at execute time.
 *
 * Type options are alphabetical per `eslint-plugin-n8n-nodes-base/
 * node-param-options-type-unsorted-items`. Display Text only meaningful for
 * messageBack (manifest §5 displayOptions).
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
						{ name: 'imBack (Send Message)', value: 'imBack' },
						{ name: 'messageBack (Send With Display Text)', value: 'messageBack' },
						{ name: 'openUrl (Open Link)', value: 'openUrl' },
						{ name: 'postBack (Send Hidden)', value: 'postBack' },
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
						'Message text for imBack/postBack, hidden payload for messageBack, URL for openUrl',
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
