import { describe, it, expect, vi } from 'vitest';
import type { IExecuteFunctions } from 'n8n-workflow';
import { buildSignInAttachment } from '../../../../shared/card/buildAttachment/signIn';

function ctx(params: Record<string, unknown>): IExecuteFunctions {
	return {
		getNode: vi.fn().mockReturnValue({ name: 'T', id: 'id', type: 't', typeVersion: 1 }),
		getNodeParameter: vi.fn().mockImplementation((name: string, _i: unknown, def?: unknown) => {
			return name in params ? params[name] : def;
		}),
	} as unknown as IExecuteFunctions;
}

describe('buildSignInAttachment', () => {
	it('builds a signin card with title + url + text', () => {
		const att = buildSignInAttachment(
			ctx({ title: 'Sign in', url: 'https://auth', text: 'Click to continue' }),
			0,
		);
		expect(att.contentType).toBe('application/vnd.microsoft.card.signin');
		const content = att.content as {
			buttons: Array<{ type: string; title: string; value: string }>;
			text?: string;
		};
		expect(content.buttons[0]).toMatchObject({
			type: 'signin',
			title: 'Sign in',
			value: 'https://auth',
		});
		expect(content.text).toBe('Click to continue');
	});

	it('omits text when empty', () => {
		const att = buildSignInAttachment(ctx({ title: 'Sign in', url: 'https://auth' }), 0);
		const content = att.content as { text?: string };
		expect(content.text).toBeUndefined();
	});
});
