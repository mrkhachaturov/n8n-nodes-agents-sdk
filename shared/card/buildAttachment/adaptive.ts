import type { IExecuteFunctions } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { Attachment } from '@microsoft/agents-activity';
import { Template } from 'adaptivecards-templating';

function parseJsonParam(
  ctx: IExecuteFunctions,
  itemIndex: number,
  fieldName: string,
  raw: unknown,
): Record<string, unknown> {
  if (raw === undefined || raw === null) return {};
  if (typeof raw !== 'string') return raw as Record<string, unknown>;
  try {
    return JSON.parse(raw) as Record<string, unknown>;
  } catch (err) {
    throw new NodeOperationError(
      ctx.getNode(),
      `Invalid JSON in ${fieldName}: ${(err as Error).message}`,
      { itemIndex },
    );
  }
}

/**
 * Build an Adaptive Card attachment by running the user's card template
 * through adaptivecards-templating `Template.expand()` with the binding data
 * as `$root`. All failures surface as NodeOperationError (user-config class).
 */
export function buildAdaptiveAttachment(
  ctx: IExecuteFunctions,
  itemIndex: number,
): Attachment {
  const cardTemplate = parseJsonParam(
    ctx,
    itemIndex,
    'Card Template',
    ctx.getNodeParameter('cardTemplate', itemIndex),
  );
  const bindingData = parseJsonParam(
    ctx,
    itemIndex,
    'Binding Data',
    ctx.getNodeParameter('bindingData', itemIndex, {}),
  );

  let rendered: unknown;
  try {
    const tpl = new Template(cardTemplate);
    rendered = tpl.expand({ $root: bindingData });
  } catch (err) {
    throw new NodeOperationError(
      ctx.getNode(),
      `Card Template expansion failed: ${(err as Error).message}`,
      { itemIndex },
    );
  }

  return {
    contentType: 'application/vnd.microsoft.card.adaptive',
    content: rendered,
  };
}
