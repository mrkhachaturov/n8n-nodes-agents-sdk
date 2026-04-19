import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import type { BotConnectorBundle } from '../../../../shared/botConnector';
import type { AuthKind, M365ClassicBotCred, M365Agent365Cred } from '../../../../shared/types';
import { identityModeFields } from '../identityModeFields';
import { dispatchUpdate } from '../../../../shared/card/dispatchUpdate';
import { buildSignInAttachment } from '../../../../shared/card/buildAttachment/signIn';

export const description: INodeProperties[] = [...identityModeFields('signInCard', 'update')];

export async function execute(
	this: IExecuteFunctions,
	itemIndex: number,
	authKind: AuthKind,
	credentials: M365ClassicBotCred | M365Agent365Cred,
	bundles: Map<string, BotConnectorBundle>,
): Promise<IDataObject> {
	return dispatchUpdate({
		ctx: this,
		itemIndex,
		resourceLabel: 'signInCard',
		authKind,
		credentials,
		bundles,
		buildAttachment: buildSignInAttachment,
	});
}
