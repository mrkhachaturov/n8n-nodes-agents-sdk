/**
 * 0.5.1 regression — required-field leaks across n8n's editor validator.
 *
 * Why this test exists
 * --------------------
 * 0.5.1 fixed a bug where every Hero/Thumbnail card's `Options.tapAction`
 * sub-collection had Type/Title/Value marked `required: true` and emitted
 * "Parameter X is required" errors on a fresh node and across resource/
 * operation switches (e.g. Message → Reply in Thread). Workflows could
 * not be activated.
 *
 * Root cause (verified against n8n upstream
 * `packages/workflow/src/node-helpers.ts` lines 1540-1546): the validator's
 * `collection` walker pushes every direct child for issue-checking
 * unconditionally — no `isDisplayed` gate, no `value === undefined` guard.
 * The `fixedCollection` walker (lines 1547-1604) DOES guard with
 * `if (value === undefined) continue;`.
 *
 * Two layers of protection
 * ------------------------
 * 1. **Structural guardrail** — walks each node's full property tree and
 *    asserts no `required: true` field is a direct child of a
 *    `type: 'collection'` ancestor. This is the architectural invariant
 *    that catches the bug class for ANY future field — adding a new
 *    card type with required fields nested under a `collection` Options
 *    will fail this test.
 *
 * 2. **Runtime check** — invokes n8n's actual `getNodeParametersIssues`
 *    validator on the M365Agent node configured for the operation RK
 *    was using when he hit the bug (Message → Reply in Thread with
 *    valid text + parentActivityId). Asserts zero false-positive issues
 *    fire. Pre-fix, this test would have caught Type/Title/Value
 *    leaking from Hero/Thumbnail's Options.tapAction.
 */

import { describe, it, expect } from 'vitest';
import {
	NodeHelpers,
	type INode,
	type INodeProperties,
	type INodeTypeDescription,
} from 'n8n-workflow';
import { versionDescription as m365AgentVersionDescription } from '../../nodes/M365Agent/actions/versionDescription';
import { M365AgentTrigger } from '../../nodes/M365AgentTrigger/M365AgentTrigger.node';
import { versionDescription as m365ConversationRefVersionDescription } from '../../nodes/M365ConversationRef/actions/versionDescription';

const m365AgentTriggerDescription = new M365AgentTrigger().description;

// ── Layer 1: Structural guardrail ────────────────────────────────────────────

interface StructuralLeak {
	path: string;
	displayName: string;
	containerChain: string[];
}

/**
 * Walk a property tree and report any `required: true` field whose nearest
 * structural ancestor is a `type: 'collection'`. Such fields will be walked
 * unconditionally by n8n's validator and emit false-positive "is required"
 * errors on a fresh node.
 *
 * Fields nested under a `fixedCollection` are safe — that walker has a
 * `value === undefined` guard.
 */
function findRequiredLeaks(
	properties: readonly INodeProperties[],
	containerChain: string[] = [],
	path = '',
): StructuralLeak[] {
	const leaks: StructuralLeak[] = [];
	for (const prop of properties) {
		const propPath = path ? `${path}.${prop.name}` : prop.name;
		// A field is leaky if it's required AND its closest container is `collection`.
		// `fixedCollection` containers are safe (validator guards on undefined value).
		if (prop.required === true && containerChain[containerChain.length - 1] === 'collection') {
			leaks.push({
				path: propPath,
				displayName: typeof prop.displayName === 'string' ? prop.displayName : prop.name,
				containerChain: [...containerChain],
			});
		}
		// Recurse into children
		if (prop.type === 'collection' && Array.isArray(prop.options)) {
			leaks.push(
				...findRequiredLeaks(
					prop.options as INodeProperties[],
					[...containerChain, 'collection'],
					propPath,
				),
			);
		} else if (prop.type === 'fixedCollection' && Array.isArray(prop.options)) {
			// fixedCollection options are { name, displayName, values: INodeProperties[] }.
			for (const group of prop.options as Array<{ name: string; values: INodeProperties[] }>) {
				if (Array.isArray(group.values)) {
					leaks.push(
						...findRequiredLeaks(
							group.values,
							[...containerChain, 'fixedCollection'],
							`${propPath}.${group.name}`,
						),
					);
				}
			}
		}
	}
	return leaks;
}

describe('0.5.1 regression — structural guardrail (no required: true under a `collection`)', () => {
	it('M365Agent has zero required-leak fields', () => {
		const leaks = findRequiredLeaks(m365AgentVersionDescription.properties);
		expect(leaks, formatLeaksForAssertion(leaks)).toEqual([]);
	});

	it('M365AgentTrigger has zero required-leak fields', () => {
		const leaks = findRequiredLeaks(m365AgentTriggerDescription.properties);
		expect(leaks, formatLeaksForAssertion(leaks)).toEqual([]);
	});

	it('M365ConversationRef has zero required-leak fields', () => {
		const leaks = findRequiredLeaks(m365ConversationRefVersionDescription.properties);
		expect(leaks, formatLeaksForAssertion(leaks)).toEqual([]);
	});
});

function formatLeaksForAssertion(leaks: StructuralLeak[]): string {
	if (leaks.length === 0) return '';
	const lines = leaks.map(
		(l) => `  - ${l.path} ("${l.displayName}") under ${l.containerChain.join(' > ')}`,
	);
	return [
		`Found ${leaks.length} required-field(s) nested directly under a \`type: 'collection'\` ancestor.`,
		`These will fire false-positive "Parameter X is required" errors on a fresh node and across`,
		`resource/operation switches, because n8n's validator walks \`collection\` children unconditionally`,
		`(packages/workflow/src/node-helpers.ts:1540 — no isDisplayed gate, no value === undefined guard).`,
		``,
		`Fix: wrap the parent in a singular \`fixedCollection\` (multipleValues: false). The`,
		`fixedCollection walker has \`if (value === undefined) continue;\` so required children`,
		`only fire once the user has materialised the inner block.`,
		``,
		`Leaking fields:`,
		...lines,
	].join('\n');
}

// ── Layer 2: Runtime check (the exact failure RK hit) ────────────────────────

function makeNode(parameters: Record<string, unknown>): INode {
	return {
		id: 't',
		name: 'M365 Agent',
		type: 'n8n-nodes-agents-sdk.m365Agent',
		typeVersion: 1,
		position: [0, 0],
		parameters,
	};
}

function makeNodeTypeDescription(): INodeTypeDescription {
	return m365AgentVersionDescription;
}

describe('0.5.1 regression — n8n validator (getNodeParametersIssues) returns no false positives', () => {
	it("Message/Reply in Thread with valid text+parentActivityId — RK's failing case", () => {
		// Pre-fix: emitted Parameter "Type" / "Title" / "Value" is required.
		// Post-fix: zero issues.
		const node = makeNode({
			authKind: 'classicBot',
			resource: 'message',
			operation: 'replyInThread',
			text: 'reply body',
			parentActivityId: '1234567890',
			options: {},
		});
		const issues = NodeHelpers.getNodeParametersIssues(
			m365AgentVersionDescription.properties,
			node,
			makeNodeTypeDescription(),
		);
		expect(issues, formatIssues(issues)).toBeNull();
	});

	it('Message/Send with valid text — fresh node, default conversationSource', () => {
		const node = makeNode({
			authKind: 'classicBot',
			resource: 'message',
			operation: 'send',
			text: 'hello',
			conversationSource: 'envelope',
			options: {},
		});
		const issues = NodeHelpers.getNodeParametersIssues(
			m365AgentVersionDescription.properties,
			node,
			makeNodeTypeDescription(),
		);
		expect(issues, formatIssues(issues)).toBeNull();
	});

	it('Adaptive Card/Send with valid cardTemplate — no Tap-Action leak across resources', () => {
		const node = makeNode({
			authKind: 'classicBot',
			resource: 'adaptiveCard',
			operation: 'send',
			cardTemplate: '{"type":"AdaptiveCard","version":"1.4","body":[]}',
			conversationSource: 'envelope',
			options: {},
		});
		const issues = NodeHelpers.getNodeParametersIssues(
			m365AgentVersionDescription.properties,
			node,
			makeNodeTypeDescription(),
		);
		expect(issues, formatIssues(issues)).toBeNull();
	});

	it('Hero Card/Send with valid title and unopened Tap Action — fixedCollection wrapper gates the walk', () => {
		const node = makeNode({
			authKind: 'classicBot',
			resource: 'heroCard',
			operation: 'send',
			title: 'Welcome',
			text: '',
			subtitle: '',
			conversationSource: 'envelope',
			images: {},
			buttons: {},
			options: {}, // tapAction NOT set — fixedCollection walker should skip its inner walk
		});
		const issues = NodeHelpers.getNodeParametersIssues(
			m365AgentVersionDescription.properties,
			node,
			makeNodeTypeDescription(),
		);
		expect(issues, formatIssues(issues)).toBeNull();
	});

	it('Thumbnail Card/Send — same gate', () => {
		const node = makeNode({
			authKind: 'classicBot',
			resource: 'thumbnailCard',
			operation: 'send',
			title: 'Thumb',
			text: '',
			subtitle: '',
			conversationSource: 'envelope',
			images: {},
			buttons: {},
			options: {},
		});
		const issues = NodeHelpers.getNodeParametersIssues(
			m365AgentVersionDescription.properties,
			node,
			makeNodeTypeDescription(),
		);
		expect(issues, formatIssues(issues)).toBeNull();
	});

	it('Hero Card with Tap Action FILLED via the fixedCollection wrapper — required fields satisfied', () => {
		const node = makeNode({
			authKind: 'classicBot',
			resource: 'heroCard',
			operation: 'send',
			title: 'Welcome',
			text: '',
			subtitle: '',
			conversationSource: 'envelope',
			images: {},
			buttons: {},
			options: {
				tapAction: {
					action: { type: 'openUrl', title: 'Open', value: 'https://example.com' },
				},
			},
		});
		const issues = NodeHelpers.getNodeParametersIssues(
			m365AgentVersionDescription.properties,
			node,
			makeNodeTypeDescription(),
		);
		expect(issues, formatIssues(issues)).toBeNull();
	});

	it('Hero Card with Tap Action wrapper opened, action sub-object empty — validator stays clean (n8n path-tracking quirk)', () => {
		// Documents an n8n validator behavior we depend on for the bug fix.
		//
		// The collection walker at node-helpers.ts:1540-1546 pushes children
		// with the SAME basePath as the parent — it does NOT prepend the
		// collection's own name. So when `tapAction` (fixedCollection) is
		// reached as a child of `options` (collection), the recursion sees
		// it with `path=''`, and its lookup
		// `getParameterValueByPath(nodeValues, 'action', 'tapAction')` looks
		// at `node.parameters.tapAction.action` — which is undefined,
		// because the actual value lives at `node.parameters.options.tapAction.action`.
		// The fixedCollection walker hits its `if (value === undefined) continue;`
		// guard and skips the inner walk entirely.
		//
		// Practical consequence: the n8n editor will never emit
		// "Parameter Type/Title/Value is required" for Tap Action — even
		// when the user has opened the wrapper but left the action empty.
		// Runtime validation in shared/card/buildAttachment/hero.ts catches
		// missing values when the workflow actually runs (buildActions
		// falls back to defaults: type='openUrl', title='', value='').
		const node = makeNode({
			authKind: 'classicBot',
			resource: 'heroCard',
			operation: 'send',
			title: 'Welcome',
			text: '',
			subtitle: '',
			conversationSource: 'envelope',
			images: {},
			buttons: {},
			options: {
				tapAction: { action: {} },
			},
		});
		const issues = NodeHelpers.getNodeParametersIssues(
			m365AgentVersionDescription.properties,
			node,
			makeNodeTypeDescription(),
		);
		expect(issues, formatIssues(issues)).toBeNull();
	});
});

function formatIssues(issues: ReturnType<typeof NodeHelpers.getNodeParametersIssues>): string {
	if (issues === null) return '';
	const lines = collectAllMessages(issues);
	return [
		`Expected zero validator issues, but the n8n editor returned:`,
		...lines.map((m) => `  - ${m}`),
	].join('\n');
}

function collectAllMessages(
	issues: ReturnType<typeof NodeHelpers.getNodeParametersIssues>,
): string[] {
	if (issues === null || !issues.parameters) return [];
	const out: string[] = [];
	for (const messages of Object.values(issues.parameters)) {
		out.push(...messages);
	}
	return out;
}
