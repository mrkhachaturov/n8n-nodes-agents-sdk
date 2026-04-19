import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { Attachment } from '@microsoft/agents-activity';
import { Template } from 'adaptivecards-templating';
import type { BotConnectorBundle } from '../../../../shared/botConnector';
import type { AuthKind, M365ClassicBotCred, M365Agent365Cred } from '../../../../shared/types';
import { identityModeFields } from '../identityModeFields';
import { dispatchSend } from '../../../../shared/card/dispatchSend';

/**
 * T3-card-state-machine — Card State Machine / select operation.
 *
 * Select a card variant by state key, apply optional templating binding, and
 * send the rendered Adaptive Card. Mirrors the "state-driven UI" pattern from
 * the car-service workflow where one node picks from N card variants based on
 * an upstream state string (e.g. order-status: "open" | "closed" | "archived").
 *
 * Reuses the shared dispatchSend pipeline (auth routing, bundle cache, raw
 * activity override) via a custom buildAttachment callback — CardState is just
 * another card resource from dispatchSend's point of view.
 */
export const description: INodeProperties[] = [...identityModeFields('cardState', 'select')];

type Variant = { key: string; cardJson: string; bindingData?: string };

/**
 * buildAttachment for CardState — resolves `stateKey`, matches against
 * `variants[]`, parses the variant's Card JSON, and (optionally) expands
 * `${placeholders}` against Binding Data via adaptivecards-templating.
 *
 * Wraps each parse failure in a NodeOperationError tagged with the specific
 * field name (`variants.cardJson` / `variants.bindingData`) so the n8n UI
 * surfaces a yellow user-config banner on the correct input.
 */
function buildCardStateAttachment(ctx: IExecuteFunctions, itemIndex: number): Attachment {
	const stateKey = ctx.getNodeParameter('stateKey', itemIndex) as string;
	const variantsRaw = ctx.getNodeParameter('variants', itemIndex, {}) as {
		values?: Variant[];
	};
	const variants = variantsRaw.values ?? [];

	const match = variants.find((v) => v.key === stateKey);
	if (!match) {
		const available = variants.map((v) => v.key).join(', ') || '(none)';
		throw new NodeOperationError(
			ctx.getNode(),
			`stateKey "${stateKey}" does not match any variant. Available keys: ${available}`,
			{ itemIndex },
		);
	}

	let cardDefinition: unknown;
	try {
		cardDefinition = JSON.parse(match.cardJson);
	} catch (err) {
		throw new NodeOperationError(
			ctx.getNode(),
			`variants.cardJson (key="${match.key}") invalid JSON: ${(err as Error).message}`,
			{ itemIndex },
		);
	}

	let expanded = cardDefinition;
	const bindingRaw = (match.bindingData ?? '').toString().trim();
	if (bindingRaw !== '') {
		let data: unknown;
		try {
			data = JSON.parse(bindingRaw);
		} catch (err) {
			throw new NodeOperationError(
				ctx.getNode(),
				`variants.bindingData (key="${match.key}") invalid JSON: ${(err as Error).message}`,
				{ itemIndex },
			);
		}
		try {
			const tpl = new Template(cardDefinition as Record<string, unknown>);
			expanded = tpl.expand({ $root: data });
		} catch (err) {
			throw new NodeOperationError(
				ctx.getNode(),
				`variants.cardJson (key="${match.key}") template expansion failed: ${(err as Error).message}`,
				{ itemIndex },
			);
		}
	}

	return {
		contentType: 'application/vnd.microsoft.card.adaptive',
		content: expanded,
	};
}

export async function execute(
	this: IExecuteFunctions,
	itemIndex: number,
	authKind: AuthKind,
	credentials: M365ClassicBotCred | M365Agent365Cred,
	bundles: Map<string, BotConnectorBundle>,
): Promise<IDataObject> {
	return dispatchSend({
		ctx: this,
		itemIndex,
		resourceLabel: 'cardState',
		authKind,
		credentials,
		bundles,
		buildAttachment: buildCardStateAttachment,
	});
}
