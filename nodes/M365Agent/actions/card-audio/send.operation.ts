import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import type { BotConnectorBundle } from '../../../../shared/botConnector';
import type { AuthKind, M365ClassicBotCred, M365Agent365Cred } from '../../../../shared/types';
import { identityModeFields } from '../identityModeFields';
import { dispatchSend } from '../../../../shared/card/dispatchSend';
import { buildAudioAttachment } from '../../../../shared/card/buildAttachment/audio';

export const description: INodeProperties[] = [...identityModeFields('audioCard', 'send')];

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
		resourceLabel: 'audioCard',
		authKind,
		credentials,
		bundles,
		buildAttachment: buildAudioAttachment,
	});
}
