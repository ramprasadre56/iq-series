import { Client } from '@microsoft/microsoft-graph-client';
import {
  DefaultAzureCredential,
  ClientSecretCredential,
  TokenCredential,
} from '@azure/identity';
import 'isomorphic-fetch';

interface GraphAuthOptions {
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
}

function getCredential(opts: GraphAuthOptions): TokenCredential {
  if (opts.tenantId && opts.clientId && opts.clientSecret) {
    return new ClientSecretCredential(
      opts.tenantId,
      opts.clientId,
      opts.clientSecret
    );
  }
  return new DefaultAzureCredential();
}

export function createGraphClient(opts: GraphAuthOptions): Client {
  const credential = getCredential(opts);
  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const token = await credential.getToken(
          'https://graph.microsoft.com/.default'
        );
        if (!token) throw new Error('Failed to acquire Graph token');
        return token.token;
      },
    },
  });
}
