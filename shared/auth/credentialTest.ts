import type { ICredentialTestFunctions, INodeCredentialTestResult } from 'n8n-workflow';
import type { M365Agent365Cred } from '../types';
import { acquireAgent365InlineToken } from './backends/agent365Inline';
import { acquireAgent365SidecarToken } from './backends/agent365Sidecar';

export async function agent365CredentialTest(
	this: ICredentialTestFunctions,
	credential: { data: M365Agent365Cred },
): Promise<INodeCredentialTestResult> {
	const cred = credential.data;
	try {
		if (cred.transport === 'inline') {
			await acquireAgent365InlineToken(
				cred,
				'autonomous',
				cred.outboundDownstreamApi ?? 'MessagingBotApi',
			);
			return { status: 'OK', message: 'inline MSAL token acquired successfully.' };
		}
		// sidecar
		const base = (cred.sidecarUrl ?? '').replace(/\/$/, '');
		const health = await fetch(`${base}/healthz`);
		if (!health.ok) {
			return { status: 'Error', message: `/healthz returned ${health.status}` };
		}
		await acquireAgent365SidecarToken(
			cred,
			'autonomous',
			cred.outboundDownstreamApi ?? 'MessagingBotApi',
		);
		return { status: 'OK', message: 'Sidecar healthy and token acquired.' };
	} catch (err) {
		return { status: 'Error', message: (err as Error).message };
	}
}
