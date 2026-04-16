import jwt, { type JwtHeader, type JwtPayload, type VerifyOptions } from 'jsonwebtoken';
import jwksRsa from 'jwks-rsa';

/** Signing-key fetcher. Default impl uses jwks-rsa; tests inject a fake. */
export type GetSigningKey = (
	jwksUri: string,
	header: JwtHeader,
) => Promise<string | Buffer>;

export interface VerifyJwtOptions {
	clientId: string;
	tenantId?: string;
	authority?: string;
	/** Override signing-key lookup (primarily for tests). */
	getSigningKey?: GetSigningKey;
	/** Override allowed algorithms (tests use HS256 for simplicity). */
	verifyAlgorithms?: jwt.Algorithm[];
}

/**
 * Verify an inbound Azure Bot Service JWT.
 *
 * Replicates the logic of @microsoft/agents-hosting's authorizeJWT middleware,
 * minus its Express coupling.
 */
export async function verifyJwt(
	raw: string,
	opts: VerifyJwtOptions,
): Promise<JwtPayload> {
	const decoded = jwt.decode(raw);
	if (!decoded || typeof decoded === 'string') {
		throw new Error('verifyJwt: token could not be decoded');
	}
	const payload = decoded as JwtPayload;

	const iss = typeof payload.iss === 'string' ? payload.iss : '';
	const jwksUri = buildJwksUri(iss, opts);

	const getKey = opts.getSigningKey ?? defaultGetSigningKey;
	const algorithms = opts.verifyAlgorithms ?? ['RS256'];

	const keyResolver = (
		header: JwtHeader,
		callback: (err: Error | null, key?: string | Buffer) => void,
	) => {
		getKey(jwksUri, header)
			.then((key) => callback(null, key))
			.catch((err) => callback(err as Error));
	};

	const verifyOptions: VerifyOptions = {
		audience: [opts.clientId, 'https://api.botframework.com'],
		ignoreExpiration: false,
		algorithms,
		clockTolerance: 300,
	};

	return await new Promise<JwtPayload>((resolve, reject) => {
		jwt.verify(raw, keyResolver, verifyOptions, (err, user) => {
			if (err) return reject(err);
			resolve(user as JwtPayload);
		});
	});
}

/**
 * Build the JWKS URI for a token issuer.
 * Bot Framework issuer → static URI.
 * Entra → tenant-scoped URI. Tenant preference:
 *   1. Tenant ID parsed from iss
 *   2. opts.tenantId (configured on credential)
 *   3. "common" as last resort
 */
export function buildJwksUri(iss: string, opts: VerifyJwtOptions): string {
	if (iss === 'https://api.botframework.com') {
		return 'https://login.botframework.com/v1/.well-known/keys';
	}
	const authority = (opts.authority ?? 'https://login.microsoftonline.com').replace(/\/$/, '');
	const tenantFromIss = extractTenantFromIss(iss);
	const tenant = tenantFromIss ?? opts.tenantId ?? 'common';
	return `${authority}/${tenant}/discovery/v2.0/keys`;
}

/** Extract the tenant GUID/segment from common Entra issuer URL shapes. */
export function extractTenantFromIss(iss: string): string | undefined {
	try {
		const u = new URL(iss);
		const segments = u.pathname.split('/').filter(Boolean);
		const first = segments[0];
		if (first && first !== 'common' && first !== 'organizations' && first !== 'consumers') {
			return first;
		}
	} catch {
		/* fall through */
	}
	return undefined;
}

const defaultGetSigningKey: GetSigningKey = async (jwksUri, header) => {
	const client = jwksRsa({ jwksUri });
	return await new Promise((resolve, reject) => {
		client.getSigningKey(header.kid, (err, key) => {
			if (err || !key) return reject(err ?? new Error('getSigningKey returned no key'));
			resolve(key.getPublicKey());
		});
	});
};
