import type { INodeProperties } from 'n8n-workflow';
import type { CardAction } from '@microsoft/agents-activity';
import { cardActionFields, buildActions } from './cardActionRow';

/**
 * Buttons builder — fixedCollection(multipleValues) whose row shape is a CardAction.
 * Used by Hero, Thumbnail, Animation, Audio, and Video resources.
 * `displayOptions` is intentionally not set here — callers wrap the returned
 * property in their resource-scoped `displayOptions`.
 */
export function buttonsField(): INodeProperties {
  return {
    displayName: 'Buttons',
    name: 'buttons',
    type: 'fixedCollection',
    typeOptions: { multipleValues: true },
    default: {},
    options: [
      {
        displayName: 'Button',
        name: 'button',
        values: cardActionFields(),
      },
    ],
  };
}

interface ButtonsCollection {
  button?: Parameters<typeof buildActions>[0];
}

/** Convert the raw Buttons fixedCollection value to `CardAction[]`. */
export function buildButtons(raw: ButtonsCollection | undefined): CardAction[] {
  return buildActions(raw?.button);
}
