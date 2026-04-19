import { describe, it, expect } from 'vitest';
import { M365ConversationRef } from '../../../nodes/M365ConversationRef/M365ConversationRef.node';

describe('M365ConversationRef — node metadata', () => {
	const node = new M365ConversationRef();

	it('has the expected display name and internal name', () => {
		expect(node.description.displayName).toBe('M365 Conversation Ref');
		expect(node.description.name).toBe('m365ConversationRef');
	});

	it('exposes a Serialize resource with save + load operations only (no delete, no adapter)', () => {
		const resourceProp = node.description.properties.find((p) => p.name === 'resource');
		expect(resourceProp?.options).toHaveLength(1);
		expect((resourceProp?.options as any[])[0].value).toBe('serialize');

		const operationProp = node.description.properties.find((p) => p.name === 'operation');
		const values = (operationProp?.options as any[]).map((o) => o.value).sort();
		expect(values).toEqual(['load', 'save']);
	});

	it('includeBearer toggle defaults to false and is save-only', () => {
		const bearerProp = node.description.properties.find((p) => p.name === 'includeBearer');
		expect(bearerProp?.default).toBe(false);
		expect(bearerProp?.displayOptions?.show).toEqual({ resource: ['serialize'], operation: ['save'] });
	});

	it('is NOT marked usableAsTool (the node handles sensitive bearer material)', () => {
		// `usableAsTool` is intentionally OMITTED from versionDescription —
		// n8n-workflow's type only accepts `true | UsableAsToolDescription |
		// undefined`, so the idiomatic "off" is absence. `toBeFalsy` accepts
		// both `undefined` (current state) and `false` (if the type ever
		// widens), matching the assertion pattern used for M365Agent and
		// M365AgentTrigger.
		expect(node.description.usableAsTool).toBeFalsy();
	});
});
