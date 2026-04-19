/**
 * Known invoke names routable in 0.5.0.
 * Free-text is permitted on the Trigger UI for unknown/future names;
 * this list powers the searchable dropdown options.
 *
 * `tokens/response` is intentionally absent — deferred to Tier 6 Auth.
 */
export const KNOWN_INVOKE_NAMES = [
	'taskModule/fetch',
	'taskModule/submit',
	'composeExtension/query',
	'composeExtension/queryLink',
	'composeExtension/submitAction',
	'composeExtension/selectItem',
	'composeExtension/setting',
	'composeExtension/querySettingUrl',
	'signin/verifyState',
	'signin/tokenExchange',
	'fileConsent/invoke',
	'actionableMessage/executeAction',
	'adaptiveCard/action',
	'handoff/initiate',
	'payments/paymentResponse',
] as const;

export type KnownInvokeName = (typeof KNOWN_INVOKE_NAMES)[number];

export const InvokeNameLabel: Record<KnownInvokeName, string> = {
	'taskModule/fetch': 'Task Module — Fetch (open dialog)',
	'taskModule/submit': 'Task Module — Submit (dialog saved)',
	'composeExtension/query': 'Compose Extension — Search Query',
	'composeExtension/queryLink': 'Compose Extension — Link Unfurl',
	'composeExtension/submitAction': 'Compose Extension — Action Submit',
	'composeExtension/selectItem': 'Compose Extension — Select Item',
	'composeExtension/setting': 'Compose Extension — Settings',
	'composeExtension/querySettingUrl': 'Compose Extension — Settings URL Query',
	'signin/verifyState': 'Sign-In — Verify State',
	'signin/tokenExchange': 'Sign-In — Token Exchange',
	'fileConsent/invoke': 'File Consent — Upload Permission',
	'actionableMessage/executeAction': 'Actionable Message — Execute Action (Outlook)',
	'adaptiveCard/action': 'Adaptive Card — Action.Execute',
	'handoff/initiate': 'Handoff — Initiate Transfer',
	'payments/paymentResponse': 'Payments — Payment Response',
};
