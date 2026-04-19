import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import type { BotConnectorBundle } from '../../../../shared/botConnector';
import type { AuthKind, M365ClassicBotCred, M365Agent365Cred } from '../../../../shared/types';
import { identityModeFields } from '../identityModeFields';
import { dispatchSend } from '../../../../shared/card/dispatchSend';
import { buildSignInAttachment } from '../../../../shared/card/buildAttachment/signIn';

export const description: INodeProperties[] = [...identityModeFields('signInCard', 'send')];

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
		resourceLabel: 'signInCard',
		authKind,
		credentials,
		bundles,
		buildAttachment: buildSignInAttachment,
	});
}
