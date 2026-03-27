"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.loadConfig = loadConfig;
function requireEnv(name) {
    const val = process.env[name];
    if (!val)
        throw new Error(`Missing required env var: ${name}`);
    return val;
}
function loadConfig() {
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
        port: parseInt(process.env.PORT ?? '3000', 10),
    };
}
