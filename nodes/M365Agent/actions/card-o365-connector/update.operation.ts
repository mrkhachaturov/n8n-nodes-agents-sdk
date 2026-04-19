import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import type { BotConnectorBundle } from '../../../../shared/botConnector';
import type { AuthKind, M365ClassicBotCred, M365Agent365Cred } from '../../../../shared/types';
import { identityModeFields } from '../identityModeFields';
import { dispatchUpdate } from '../../../../shared/card/dispatchUpdate';
import { buildO365ConnectorAttachment } from '../../../../shared/card/buildAttachment/o365Connector';

export const description: INodeProperties[] = [
	...identityModeFields('o365ConnectorCard', 'update'),
];

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
		resourceLabel: 'o365ConnectorCard',
		authKind,
		credentials,
		bundles,
		buildAttachment: buildO365ConnectorAttachment,
	});
}
