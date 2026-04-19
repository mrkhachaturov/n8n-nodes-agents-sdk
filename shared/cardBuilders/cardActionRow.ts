import type { INodeProperties } from 'n8n-workflow';
import type { CardAction } from '@microsoft/agents-activity';

/**
 * Field generator for a CardAction row. The same shape is reused by:
 *  - Buttons (fixedCollection multipleValues) — top-level builder
 *  - Image row Tap Action (collection) — nested inside images.ts
 * One source of truth per the spec's §2.1.
 */
export function cardActionFields(): INodeProperties[] {
  return [
    {
      displayName: 'Type',
      name: 'type',
      type: 'options',
      required: true,
      noDataExpression: true,
      default: 'openUrl',
      options: [
        { name: 'Call', value: 'call' },
        { name: 'Download File', value: 'downloadFile' },
        { name: 'IM Back', value: 'imBack' },
        { name: 'Message Back', value: 'messageBack' },
        { name: 'Open App', value: 'openApp' },
        { name: 'Open URL', value: 'openUrl' },
        { name: 'Play Audio', value: 'playAudio' },
        { name: 'Play Video', value: 'playVideo' },
        { name: 'Post Back', value: 'postBack' },
        { name: 'Show Image', value: 'showImage' },
        { name: 'Sign In', value: 'signin' },
      ],
      description:
        'Activity kind invoked when the user taps this action. Open URL opens a URL; IM Back sends Value as a user message; Post Back sends Value silently to the bot.',
    },
    {
      displayName: 'Title',
      name: 'title',
      type: 'string',
      required: true,
      default: '',
      description: 'Label shown on the button / action',
    },
    {
      displayName: 'Value',
      name: 'value',
      type: 'string',
      required: true,
      default: '',
      placeholder: 'https://... for Open URL, text for IM Back, payload string for Post Back',
      description:
        'String payload associated with the action. Use Options → Value JSON for object payloads (Message Back, Post Back).',
    },
    {
      displayName: 'Options',
      name: 'options',
      type: 'collection',
      placeholder: 'Add option',
      default: {},
      options: [
        {
          displayName: 'Channel Data',
          name: 'channelData',
          type: 'json',
          default: '={}',
          typeOptions: { rows: 2 },
          description: 'Channel-specific data attached to the action',
        },
        {
          displayName: 'Display Text',
          name: 'displayText',
          type: 'string',
          default: '',
          description: 'Text shown in the conversation after the user clicks (Message Back)',
        },
        {
          displayName: 'Image',
          name: 'image',
          type: 'string',
          default: '',
          description: 'URL of an image displayed on the action',
        },
        {
          displayName: 'Image Alt Text',
          name: 'imageAltText',
          type: 'string',
          default: '',
          description: 'Alt text for the action image',
        },
        {
          displayName: 'Text',
          name: 'text',
          type: 'string',
          default: '',
          description: 'Additional text sent back to the bot (Message Back)',
        },
        {
          displayName: 'Value JSON',
          name: 'valueJson',
          type: 'json',
          default: '={}',
          typeOptions: { rows: 2 },
          description:
            'If set, overrides Value when mapping to CardAction.value. Use for object payloads.',
        },
      ],
    },
  ];
}

interface CardActionRowInput {
  type?: string;
  title?: string;
  value?: string;
  options?: {
    valueJson?: unknown;
    channelData?: unknown;
    image?: string;
    imageAltText?: string;
    text?: string;
    displayText?: string;
  };
}

/**
 * Typed error for shared-builder user-config failures.
 * Dispatchers (Tasks 6, 7) recognize this class and normalize it to
 * `NodeOperationError` instead of the default-to-NodeApiError path.
 * Keeping it here — rather than threading `IExecuteFunctions` through the
 * builder signatures — lets the pure builder modules stay n8n-runtime-free
 * and keeps unit tests trivial to construct.
 */
export class CardBuildError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'CardBuildError';
  }
}

/**
 * Parse a raw JSON-editor value. The SDK's `CardAction.value` is `any` and
 * `CardAction.channelData` is `unknown`, so any JSON scalar / array / object
 * is valid. We intentionally preserve the full payload shape.
 */
function parseJsonOr(raw: unknown, fieldName: string): unknown | undefined {
  if (raw === undefined || raw === null) return undefined;
  if (typeof raw !== 'string') return raw;
  const trimmed = raw.trim();
  if (trimmed === '' || trimmed === '={}' || trimmed === '{}') return undefined;
  try {
    return JSON.parse(trimmed);
  } catch (err) {
    throw new CardBuildError(`Invalid JSON in ${fieldName}: ${(err as Error).message}`);
  }
}

/** Convert an array of CardAction row inputs to `CardAction[]`. */
export function buildActions(rows: CardActionRowInput[] | undefined): CardAction[] {
  if (!rows || rows.length === 0) return [];
  return rows.map((row) => {
    const options = row.options ?? {};
    const valueJson = parseJsonOr(options.valueJson, 'Value JSON');
    const channelData = parseJsonOr(options.channelData, 'Channel Data');
    const action: CardAction = {
      type: row.type ?? 'openUrl',
      title: row.title ?? '',
      value: valueJson !== undefined ? valueJson : (row.value ?? ''),
    };
    if (channelData !== undefined) action.channelData = channelData;
    if (options.image) action.image = options.image;
    if (options.imageAltText) action.imageAltText = options.imageAltText;
    if (options.text) action.text = options.text;
    if (options.displayText) action.displayText = options.displayText;
    return action;
  });
}
