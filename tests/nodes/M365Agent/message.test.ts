import { describe, it, expect } from 'vitest';
import * as message from '../../../nodes/M365Agent/actions/message';

describe('message resource description', () => {
	it('exposes Delete / Reply / Reply in Thread / Send / Typing / Update in alphabetical display order', () => {
		// Assert as-written; manifest §14 requires alphabetical by `name` and
		// sort-before-compare would hide a regression.
		const op = message.description.find((p) => p.name === 'operation');
		const names = (op?.options as { name: string }[] | undefined)?.map((o) => o.name);
		expect(names).toEqual(['Delete', 'Reply', 'Reply in Thread', 'Send', 'Typing', 'Update']);
	});

	it('exposes all six message operations (alphabetical by value)', () => {
		const op = message.description.find((p) => p.name === 'operation');
		const values = (op?.options as { value: string }[] | undefined)?.map((o) => o.value);
		expect(values).toEqual(['delete', 'reply', 'replyInThread', 'send', 'typing', 'update']);
	});

	it('Parent Activity ID is a top-level required field shown only for replyInThread', () => {
		const pa = message.description.find((p) => p.name === 'parentActivityId');
		expect(pa).toBeDefined();
		expect(pa?.required).toBe(true);
		const shown = (pa?.displayOptions?.show as Record<string, string[]> | undefined)?.operation;
		expect(shown).toEqual(['replyInThread']);
	});

	it('each operation has an action string for the subtitle', () => {
		const op = message.description.find((p) => p.name === 'operation');
		const actions = (op?.options as { action?: string }[] | undefined)?.map((o) => o.action);
		expect(actions?.every((a) => typeof a === 'string' && a.length > 0)).toBe(true);
	});

	it('operation selector is noDataExpression', () => {
		const op = message.description.find((p) => p.name === 'operation');
		expect(op?.noDataExpression).toBe(true);
	});

	it('Text field is shown for every body-carrying op (send/reply/update/replyInThread) but not delete', () => {
		const text = message.description.find((p) => p.name === 'text');
		const shown = (text?.displayOptions?.show as Record<string, string[]> | undefined)?.operation;
		expect(shown).toEqual(['send', 'reply', 'update', 'replyInThread']); // exact order asserted
		expect(shown).not.toContain('delete');
	});

	it('Options collection includes workflowFooter, mentions, rawActivityOverride, and suggestedActions (to-override nested inside suggestedActions)', () => {
		const opts = message.description.find((p) => p.name === 'options');
		const inner = (opts?.options as Array<{ name: string }>) ?? [];
		const names = inner.map((o) => o.name);
		// The recipient-override field is NO LONGER a peer — it lives as `to` inside
		// the `suggestedActions` collection per spec §4/§8 (audit A3/D2-A).
		expect(names).toEqual([
			'workflowFooter',
			'mentions',
			'rawActivityOverride',
			'suggestedActions',
		]);
	});

	it('mentions and suggestedActions are gated to body-carrying operations', () => {
		const opts = message.description.find((p) => p.name === 'options');
		const shown = (opts?.displayOptions?.show as Record<string, string[]> | undefined)?.operation;
		// Inherited from Options collection — same gate as the existing workflowFooter.
		expect(shown).toEqual(['send', 'reply', 'update', 'replyInThread']);
	});
});
