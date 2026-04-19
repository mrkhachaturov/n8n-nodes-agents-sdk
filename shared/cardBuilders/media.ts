import type { INodeProperties } from 'n8n-workflow';

interface MediaUrl {
  url: string;
  profile?: string;
}

/**
 * Media builder — fixedCollection(multipleValues) with url + Options.profile.
 * Used by Animation, Audio, and Video resources.
 */
export function mediaField(): INodeProperties {
  return {
    displayName: 'Media URLs',
    name: 'media',
    type: 'fixedCollection',
    typeOptions: { multipleValues: true },
    default: {},
    options: [
      {
        displayName: 'Media Item',
        name: 'mediaItem',
        values: [
          {
            displayName: 'URL',
            name: 'url',
            type: 'string',
            required: true,
            default: '',
            description: 'URL of the media resource',
          },
          {
            displayName: 'Options',
            name: 'options',
            type: 'collection',
            placeholder: 'Add option',
            default: {},
            options: [
              {
                displayName: 'Profile',
                name: 'profile',
                type: 'string',
                default: '',
                description: 'Quality / format hint for the media URL',
              },
            ],
          },
        ],
      },
    ],
  };
}

interface MediaRow {
  url?: string;
  options?: { profile?: string };
}

interface MediaCollection {
  mediaItem?: MediaRow[];
}

/** Convert the raw Media URLs fixedCollection to `MediaUrl[]`. */
export function buildMedia(raw: MediaCollection | undefined): MediaUrl[] {
  const rows = raw?.mediaItem;
  if (!rows || rows.length === 0) return [];
  return rows.map((row) => {
    const item: MediaUrl = { url: row.url ?? '' };
    if (row.options?.profile) item.profile = row.options.profile;
    return item;
  });
}
