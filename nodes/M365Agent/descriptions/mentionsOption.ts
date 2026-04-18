import type { INodeProperties } from 'n8n-workflow';

/**
 * Mentions property — nested inside the Message resource's Options collection.
 * Each row produces one `<at>Name</at>` token + one matching Mention entity at
 * execute time. User covers individual @-mentions; Everyone is the Teams
 * team-wide announce backed by the magic ID `29:allchannel` (auto-set — the
 * user doesn't type it).
 *
 * Manifest §4 (fixedCollection + multipleValues), §5 (displayOptions on id),
 * §6.4 (lives inside Options collection).
 */
export const mentionsOption: INodeProperties = {
	displayName: 'Mentions',
	name: 'mentions',
	type: 'fixedCollection',
	typeOptions: { multipleValues: true },
	placeholder: 'Add mention',
	default: {},
	description:
		'Add one or more Teams @-mentions. The node will insert matching &lt;at&gt;Name&lt;/at&gt; tokens into the message text and attach the required mention entities.',
	options: [
		{
			name: 'values',
			displayName: 'Mentions',
			values: [
				{
					displayName: 'Type',
					name: 'type',
					type: 'options',
					options: [
						{
							name: 'Everyone (Team)',
							value: 'everyone',
							description:
								"Pings every member of the team via Teams' built-in announcement mention",
						},
						{
							name: 'User',
							value: 'user',
							description: 'Pings one specific person',
						},
					],
					default: 'user',
					description: 'Everyone pings every member of the team; User pings one person',
				},
				{
					displayName: 'User ID',
					name: 'id',
					type: 'string',
					default: '',
					required: true,
					placeholder: 'e.g. 29:1abcDefG...',
					description:
						'Teams/AAD object ID of the user to mention. Typically pasted from the trigger envelope: {{ $JSON.parsed.userId }}.',
					displayOptions: { show: { type: ['user'] } },
				},
				{
					displayName: 'Name',
					name: 'name',
					type: 'string',
					default: '',
					required: true,
					placeholder: 'e.g. Alice',
					description:
						'Display name shown inside the &lt;at&gt; token. Must be non-empty. Typically: {{ $JSON.parsed.userName }}.',
				},
			],
		},
	],
};
