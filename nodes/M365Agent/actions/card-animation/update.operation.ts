import type { IDataObject, IExecuteFunctions, INodeProperties } from 'n8n-workflow';
import type { BotConnectorBundle } from '../../../../shared/botConnector';
import type { AuthKind, M365ClassicBotCred, M365Agent365Cred } from '../../../../shared/types';
import { identityModeFields } from '../identityModeFields';
import { dispatchUpdate } from '../../../../shared/card/dispatchUpdate';
import { buildAnimationAttachment } from '../../../../shared/card/buildAttachment/animation';

export const description: INodeProperties[] = [...identityModeFields('animationCard', 'update')];

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
		resourceLabel: 'animationCard',
		authKind,
		credentials,
		bundles,
		buildAttachment: buildAnimationAttachment,
	});
}
