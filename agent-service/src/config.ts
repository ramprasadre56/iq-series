export interface Config {
  azure: {
    aiEndpoint: string;
    foundryProjectName: string;
    searchEndpoint: string;
    searchIndexName: string;
    tenantId?: string;
    clientId?: string;
    clientSecret?: string;
  };
  reasoningModel: string;
  port: number;
}

function requireEnv(name: string): string {
  const val = process.env[name];
  if (!val) throw new Error(`Missing required env var: ${name}`);
  return val;
}

export function loadConfig(): Config {
  return {
    azure: {
      aiEndpoint: requireEnv('AZURE_AI_ENDPOINT'),
      foundryProjectName: requireEnv('FOUNDRY_PROJECT_NAME'),
      searchEndpoint: requireEnv('AZURE_SEARCH_ENDPOINT'),
      searchIndexName: requireEnv('AZURE_SEARCH_INDEX_NAME'),
      tenantId: process.env.AZURE_TENANT_ID,
      clientId: process.env.AZURE_CLIENT_ID,
      clientSecret: process.env.AZURE_CLIENT_SECRET,
    },
    reasoningModel: process.env.REASONING_MODEL ?? 'azure/gpt-4o',
    port: (() => {
      const p = parseInt(process.env.PORT ?? '3000', 10);
      if (!Number.isFinite(p)) throw new Error('Invalid PORT value: must be a number');
      return p;
    })(),
  };
}
