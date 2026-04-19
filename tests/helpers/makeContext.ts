/**
 * Minimal n8n context factories for unit testing.
 *
 * These stubs satisfy the IWebhookFunctions / IExecuteFunctions call-site
 * interfaces without importing from n8n-workflow at runtime (the peer dep is
 * available in tests through node_modules but real contexts are not
 * constructable without the full n8n runtime).
 */
import { vi } from 'vitest';

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

/** Build a mock node ref sufficient for error construction. */
export function makeNodeRef(name = 'Test Node') {
	return { name, id: 'test-node-id', type: 'test', typeVersion: 1 };
}

/** Build a credentials record. */
export function makeCredentials(overrides: Record<string, unknown> = {}) {
	return {
		appType: 'SingleTenant',
		clientId: 'test-client-id',
		clientSecret: 'test-secret',
		tenantId: 'test-tenant-id',
		anonymousAllowed: false,
		...overrides,
	};
}

// ---------------------------------------------------------------------------
// IWebhookFunctions mock factory
// ---------------------------------------------------------------------------

export interface WebhookContextOptions {
	method?: string;
	headers?: Record<string, string>;
	/** Parsed request body (Activity shape or plain object). */
	body?: Record<string, unknown>;
	credentials?: Record<string, unknown>;
	/** Map from parameter name → value. */
	parameters?: Record<string, unknown>;
}

export function makeWebhookContext(opts: WebhookContextOptions = {}) {
	const {
		method = 'POST',
		headers = {},
		body = {},
		credentials = makeCredentials(),
		parameters = {},
	} = opts;

	// Track response calls so assertions can inspect them
	const jsonFn = vi.fn();
	const statusFn = vi.fn().mockReturnValue({ json: jsonFn });
	const getResponseObject = vi.fn().mockReturnValue({ status: statusFn });

	const getRequestObject = vi.fn().mockReturnValue({ method, headers, body });
	const getCredentials = vi.fn().mockResolvedValue(credentials);
	const getBodyData = vi.fn().mockReturnValue(body);
	const getNodeParameter = vi.fn().mockImplementation((name: string, defaultValue?: unknown) => {
		return name in parameters ? parameters[name] : defaultValue;
	});
	const getNode = vi.fn().mockReturnValue(makeNodeRef());

	const returnJsonArray = vi
		.fn()
		.mockImplementation((items: Record<string, unknown>[]) => items.map((json) => ({ json })));

	return {
		getRequestObject,
		getResponseObject,
		getCredentials,
		getBodyData,
		getNodeParameter,
		getNode,
		helpers: { returnJsonArray },
		// Expose raw mocks for assertions
		_statusFn: statusFn,
		_jsonFn: jsonFn,
	};
}

// ---------------------------------------------------------------------------
// IExecuteFunctions mock factory
// ---------------------------------------------------------------------------

export interface ExecuteContextOptions {
	/** Input items — each element becomes { json: <element> }. */
	inputItems?: Record<string, unknown>[];
	credentials?: Record<string, unknown>;
	/**
	 * Parameter map: key is `${paramName}:${itemIndex}` for per-item params,
	 * or just `${paramName}` for params that return the same value for every item.
	 */
	parameters?: Record<string, unknown>;
	/** Value returned by `this.continueOnFail()`. Defaults to false. */
	continueOnFail?: boolean;
}

export function makeExecuteContext(opts: ExecuteContextOptions = {}) {
	const {
		inputItems = [],
		credentials = makeCredentials(),
		parameters = {},
		continueOnFail = false,
	} = opts;

	const getInputData = vi.fn().mockReturnValue(inputItems.map((json) => ({ json })));
	const getCredentials = vi.fn().mockResolvedValue(credentials);
	const getNode = vi.fn().mockReturnValue(makeNodeRef());

	const getNodeParameter = vi
		.fn()
		.mockImplementation((name: string, itemIndex: unknown, defaultValue?: unknown) => {
			// Support indexed lookup first, then name-only fallback
			const indexed = `${name}:${String(itemIndex)}`;
			if (indexed in parameters) return parameters[indexed];
			if (name in parameters) return parameters[name];
			return defaultValue;
		});

	const continueOnFailFn = vi.fn().mockReturnValue(continueOnFail);

	return {
		getInputData,
		getCredentials,
		getNode,
		getNodeParameter,
		continueOnFail: continueOnFailFn,
	};
}
