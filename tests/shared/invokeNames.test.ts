import { describe, it, expect } from 'vitest';
import { KNOWN_INVOKE_NAMES, InvokeNameLabel } from '../../shared/invokeNames';

const EXPECTED_05 = [
	'taskModule/fetch', 'taskModule/submit',
	'composeExtension/query', 'composeExtension/queryLink', 'composeExtension/submitAction',
	'composeExtension/selectItem', 'composeExtension/setting', 'composeExtension/querySettingUrl',
	'signin/verifyState', 'signin/tokenExchange',
	'fileConsent/invoke', 'actionableMessage/executeAction',
	'adaptiveCard/action', 'handoff/initiate',
	'payments/paymentResponse',
];

describe('KNOWN_INVOKE_NAMES — 0.5.0 release', () => {
	it('exports exactly the 15 invoke names scheduled for 0.5.0', () => {
		expect(KNOWN_INVOKE_NAMES.slice().sort()).toEqual(EXPECTED_05.slice().sort());
	});

	it('every invoke name has a human-readable label', () => {
		for (const n of KNOWN_INVOKE_NAMES) {
			expect(typeof InvokeNameLabel[n]).toBe('string');
			expect(InvokeNameLabel[n].length).toBeGreaterThan(0);
		}
	});
});
