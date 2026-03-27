import { createGraphClient } from '../../src/auth/graphClient';

describe('createGraphClient', () => {
  it('returns a Graph client when called without App Reg env vars (Managed Identity path)', () => {
    const client = createGraphClient({});
    expect(client).toBeDefined();
    expect(typeof client.api).toBe('function');
  });

  it('returns a Graph client when called with App Reg credentials (local dev path)', () => {
    const client = createGraphClient({
      tenantId: 'fake-tenant',
      clientId: 'fake-client',
      clientSecret: 'fake-secret',
    });
    expect(client).toBeDefined();
    expect(typeof client.api).toBe('function');
  });
});
