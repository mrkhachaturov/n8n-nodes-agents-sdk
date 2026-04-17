import { describe, it, expect } from 'vitest';
import * as message from '../../../nodes/M365Agent/actions/message';

describe('message resource description', () => {
	it('exposes Delete / Reply / Send / Update (Task 6 adds Reply in Thread) in alphabetical display order', () => {
		// Assert as-written; manifest §14 requires alphabetical by `name` and
		// sort-before-compare would hide a regression.
		const op = message.description.find((p) => p.name === 'operation');
		const names = (op?.options as { name: string }[] | undefined)?.map((o) => o.name);
		expect(names).toEqual(['Delete', 'Reply', 'Send', 'Update']);
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
});
