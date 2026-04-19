import { describe, it, expect } from 'vitest';
import { versionDescription } from '../../../nodes/M365Agent/actions/versionDescription';

describe('M365Agent versionDescription', () => {
	it('lists exactly 12 resources, alphabetical by display name', () => {
		// Assert the array as-written — sorting before compare hides misordering.
		const resource = versionDescription.properties.find((p) => p.name === 'resource')!;
		const options = resource.options as Array<{ name: string; value: string }>;
		expect(options.map((o) => o.name)).toEqual([
			'Adaptive Card',
			'Animation Card',
			'Audio Card',
			'Hero Card',
			'Invoke Response',
			'Message',
			'O365 Connector Card',
			'Raw Attachment',
			'Receipt Card',
			'Sign-In Card',
			'Thumbnail Card',
			'Video Card',
		]);
		expect(options.map((o) => o.value)).toEqual([
			'adaptiveCard',
			'animationCard',
			'audioCard',
			'heroCard',
			'invokeResponse',
			'message',
			'o365ConnectorCard',
			'rawAttachment',
			'receiptCard',
			'signInCard',
			'thumbnailCard',
			'videoCard',
		]);
	});

	it('Resource default is "message"', () => {
		const resource = versionDescription.properties.find((p) => p.name === 'resource')!;
		expect(resource.default).toBe('message');
	});

	it('subtitle emits "<operation> <resource>" without punctuation', () => {
		expect(versionDescription.subtitle).toBe(
			'={{$parameter["operation"] + " " + $parameter["resource"]}}',
		);
	});

	it('uses m365AgentApi credential', () => {
		expect(versionDescription.credentials?.[0]?.name).toBe('m365AgentApi');
	});

	it('resource param has noDataExpression', () => {
		const resource = versionDescription.properties.find((p) => p.name === 'resource');
		expect(resource?.noDataExpression).toBe(true);
	});
});
