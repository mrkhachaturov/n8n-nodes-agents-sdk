import { MsalTokenProvider } from '@microsoft/agents-hosting';
import { buildAuthConfig } from '../../buildAuthConfig';
import type { M365ClassicBotCred } from '../../types';

const BOT_FRAMEWORK_SCOPE = 'https://api.botframework.com';

export async function acquireClassicBotToken(
	cred: M365ClassicBotCred,
): Promise<{ authorizationHeader: string }> {
	const authConfig = buildAuthConfig(cred);
	const provider = new MsalTokenProvider();
	const token = await provider.getAccessToken(authConfig, BOT_FRAMEWORK_SCOPE);
	return { authorizationHeader: `Bearer ${token}` };
}
