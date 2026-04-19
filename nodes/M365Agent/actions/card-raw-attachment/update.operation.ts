import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import type { BotConnectorBundle } from '../../../../shared/botConnector';
import type { AuthKind, M365ClassicBotCred, M365Agent365Cred } from '../../../../shared/types';
import { identityModeFields } from '../identityModeFields';
import { dispatchUpdate } from '../../../../shared/card/dispatchUpdate';
import { buildRawAttachment } from '../../../../shared/card/buildAttachment/rawAttachment';

export const description: INodeProperties[] = [...identityModeFields('rawAttachment', 'update')];

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
		resourceLabel: 'rawAttachment',
		authKind,
		credentials,
		bundles,
		buildAttachment: buildRawAttachment,
	});
}
