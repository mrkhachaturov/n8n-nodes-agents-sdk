import { describe, it, expect } from 'vitest';
import { buildReplyInThreadUrl } from '../../shared/botConnector';

describe('shared/botConnector — URL builders', () => {
	it('builds the Teams thread reply URL with ;messageid= suffix', () => {
		const url = buildReplyInThreadUrl('19:thread@tacv2', 'parent-msg-42');
		expect(url).toBe('v3/conversations/19:thread@tacv2;messageid=parent-msg-42/activities');
	});

	it('url-encodes the conversation id if it contains reserved chars', () => {
		const url = buildReplyInThreadUrl('a/b?c', 'pid');
		expect(url).toBe('v3/conversations/a%2Fb%3Fc;messageid=pid/activities');
	});
});

// NOTE: The higher-level client factory (createConnector) integrates with MSAL
// and axios. We test it end-to-end against the Bot Framework Emulator in Task 15
// rather than duck-typing the SDK here — the SDK owns the integration surface.
