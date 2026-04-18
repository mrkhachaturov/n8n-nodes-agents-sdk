import type { INodeProperties } from 'n8n-workflow';
import { identityModeFields as sharedIdentityModeFields } from '../identityModeFields';

type MessageOperation = 'send' | 'reply' | 'update' | 'delete' | 'replyInThread';

/** Thin wrapper — keeps the single-arg call signature used by all message operations.
 *  Delegates to the shared factory in ../identityModeFields with resource='message'. */
export function identityModeFields(operation: MessageOperation): INodeProperties[] {
	return sharedIdentityModeFields('message', operation);
}
