# SharePoint Foundry IQ Agent Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a TypeScript Azure Container Apps service that registers as a native Azure AI Foundry Agent with seven SharePoint tools (read via Foundry IQ + write via Microsoft Graph), configurable reasoning model, and full playground testability.

**Architecture:** An Express HTTP API deployed to ACA exposes tool endpoints that the Foundry Agent calls during its reasoning loop. SharePoint reads flow through both the Foundry IQ AI Search index (documents/pages) and Microsoft Graph (lists). All writes go through Microsoft Graph. Auth uses `DefaultAzureCredential` in production (Managed Identity) and `ClientSecretCredential` locally.

**Tech Stack:** TypeScript 5, Node.js 20, Express 4, `@azure/ai-projects` (Foundry Agent SDK), `@microsoft/microsoft-graph-client`, `@azure/identity`, `@azure/search-documents`, Jest, Docker, Azure Container Apps, Bicep

---

## File Map

```
agent-service/
├── src/
│   ├── server.ts                   # Express app entry point, route wiring
│   ├── config.ts                   # Typed env var loading
│   ├── auth/
│   │   └── graphClient.ts          # Graph client factory (Managed Identity vs App Reg)
│   └── tools/
│       ├── knowledge.ts            # searchKnowledge — AI Search / Foundry IQ
│       ├── lists.ts                # getListItems, createListItem, updateListItem
│       ├── documents.ts            # uploadDocument
│       ├── pages.ts                # createPage
│       └── notifications.ts       # sendNotification (mail/Teams)
├── tests/
│   ├── tools/
│   │   ├── knowledge.test.ts
│   │   ├── lists.test.ts
│   │   ├── documents.test.ts
│   │   ├── pages.test.ts
│   │   └── notifications.test.ts
│   └── auth/
│       └── graphClient.test.ts
├── Dockerfile
├── package.json
├── tsconfig.json
└── .env.example

scripts/
└── register-agent.ts               # Upsert Foundry Agent registration

infra/
└── agent.bicep                     # ACA environment + app + identity + role assignments

.github/
└── workflows/
    └── deploy-agent.yml            # Build → push ACR → update ACA → register agent
```

---

## Task 1: Project Scaffold

**Files:**
- Create: `agent-service/package.json`
- Create: `agent-service/tsconfig.json`
- Create: `agent-service/.env.example`
- Create: `agent-service/src/config.ts`

- [ ] **Step 1: Create `agent-service/package.json`**

```json
{
  "name": "sharepoint-foundry-agent",
  "version": "1.0.0",
  "private": true,
  "scripts": {
    "build": "tsc",
    "start": "node dist/server.js",
    "dev": "ts-node src/server.ts",
    "test": "jest --runInBand"
  },
  "dependencies": {
    "@azure/ai-projects": "^1.0.0",
    "@azure/identity": "^4.4.0",
    "@azure/search-documents": "^12.0.0",
    "@microsoft/microsoft-graph-client": "^3.0.0",
    "express": "^4.19.0",
    "isomorphic-fetch": "^3.0.0"
  },
  "devDependencies": {
    "@types/express": "^4.17.21",
    "@types/jest": "^29.5.12",
    "@types/node": "^20.14.0",
    "jest": "^29.7.0",
    "ts-jest": "^29.2.0",
    "ts-node": "^10.9.2",
    "typescript": "^5.5.0"
  }
}
```

- [ ] **Step 2: Create `agent-service/tsconfig.json`**

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "commonjs",
    "lib": ["ES2022"],
    "outDir": "./dist",
    "rootDir": "./src",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "resolveJsonModule": true,
    "forceConsistentCasingInFileNames": true
  },
  "include": ["src/**/*"],
  "exclude": ["node_modules", "dist", "tests"]
}
```

- [ ] **Step 3: Create `agent-service/.env.example`**

```
# Azure AI Foundry / AI Services
AZURE_AI_ENDPOINT=https://<aiservices-name>.cognitiveservices.azure.com
FOUNDRY_PROJECT_NAME=iqseries-project

# Azure AI Search (Foundry IQ knowledge base)
AZURE_SEARCH_ENDPOINT=https://<search-name>.search.windows.net
AZURE_SEARCH_INDEX_NAME=<your-knowledge-base-index>

# Reasoning model — any Azure AI Gateway model string
REASONING_MODEL=azure/gpt-4o

# Microsoft Graph auth (local dev only — leave blank in production)
AZURE_TENANT_ID=
AZURE_CLIENT_ID=
AZURE_CLIENT_SECRET=

# Server
PORT=3000
```

- [ ] **Step 4: Create `agent-service/src/config.ts`**

```typescript
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

function require(name: string): string {
  const val = process.env[name];
  if (!val) throw new Error(`Missing required env var: ${name}`);
  return val;
}

export function loadConfig(): Config {
  return {
    azure: {
      aiEndpoint: require('AZURE_AI_ENDPOINT'),
      foundryProjectName: require('FOUNDRY_PROJECT_NAME'),
      searchEndpoint: require('AZURE_SEARCH_ENDPOINT'),
      searchIndexName: require('AZURE_SEARCH_INDEX_NAME'),
      tenantId: process.env.AZURE_TENANT_ID,
      clientId: process.env.AZURE_CLIENT_ID,
      clientSecret: process.env.AZURE_CLIENT_SECRET,
    },
    reasoningModel: process.env.REASONING_MODEL ?? 'azure/gpt-4o',
    port: parseInt(process.env.PORT ?? '3000', 10),
  };
}
```

- [ ] **Step 5: Install dependencies**

```bash
cd agent-service
npm install
```

Expected: `node_modules/` created, no errors.

- [ ] **Step 6: Verify TypeScript compiles**

```bash
npm run build
```

Expected: `dist/` created (only `config.js` at this point), no TypeScript errors.

- [ ] **Step 7: Commit**

```bash
git add agent-service/
git commit -m "feat: scaffold agent-service TypeScript project"
```

---

## Task 2: Graph Client Auth Factory

**Files:**
- Create: `agent-service/src/auth/graphClient.ts`
- Create: `agent-service/tests/auth/graphClient.test.ts`

- [ ] **Step 1: Write the failing test**

Create `agent-service/tests/auth/graphClient.test.ts`:

```typescript
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
```

- [ ] **Step 2: Run test to verify it fails**

```bash
cd agent-service
npx jest tests/auth/graphClient.test.ts --no-coverage
```

Expected: FAIL — `Cannot find module '../../src/auth/graphClient'`

- [ ] **Step 3: Create `agent-service/src/auth/graphClient.ts`**

```typescript
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
```

- [ ] **Step 4: Run test to verify it passes**

```bash
npx jest tests/auth/graphClient.test.ts --no-coverage
```

Expected: PASS — 2 tests

- [ ] **Step 5: Commit**

```bash
git add agent-service/src/auth/ agent-service/tests/auth/
git commit -m "feat: add Graph client auth factory with Managed Identity + App Reg support"
```

---

## Task 3: Knowledge Tool (AI Search / Foundry IQ)

**Files:**
- Create: `agent-service/src/tools/knowledge.ts`
- Create: `agent-service/tests/tools/knowledge.test.ts`

- [ ] **Step 1: Write the failing test**

Create `agent-service/tests/tools/knowledge.test.ts`:

```typescript
import { searchKnowledge, KnowledgeResult } from '../../src/tools/knowledge';
import { SearchClient } from '@azure/search-documents';

jest.mock('@azure/search-documents');

const mockSearch = jest.fn();
(SearchClient as jest.MockedClass<typeof SearchClient>).mockImplementation(
  () => ({ search: mockSearch } as unknown as SearchClient<Record<string, unknown>>)
);

describe('searchKnowledge', () => {
  const opts = {
    searchEndpoint: 'https://fake.search.windows.net',
    indexName: 'test-index',
  };

  it('returns ranked results for a query', async () => {
    mockSearch.mockResolvedValueOnce({
      results: (async function* () {
        yield { document: { content: 'Azure AI Foundry overview', sourcefile: 'overview.pdf' }, score: 0.95 };
        yield { document: { content: 'Foundry IQ knowledge layer', sourcefile: 'iq.pdf' }, score: 0.88 };
      })(),
    });

    const results: KnowledgeResult[] = await searchKnowledge('What is Foundry IQ?', opts);

    expect(results).toHaveLength(2);
    expect(results[0].content).toBe('Azure AI Foundry overview');
    expect(results[0].source).toBe('overview.pdf');
    expect(results[0].score).toBe(0.95);
  });

  it('returns empty array when no results found', async () => {
    mockSearch.mockResolvedValueOnce({
      results: (async function* () {})(),
    });

    const results = await searchKnowledge('unknown topic', opts);
    expect(results).toEqual([]);
  });

  it('returns AuthError message on 403', async () => {
    mockSearch.mockRejectedValueOnce({ statusCode: 403, message: 'Forbidden' });

    const results = await searchKnowledge('query', opts);
    expect(results).toEqual([{ error: 'AuthError: insufficient permissions for knowledge base index' }]);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx jest tests/tools/knowledge.test.ts --no-coverage
```

Expected: FAIL — `Cannot find module '../../src/tools/knowledge'`

- [ ] **Step 3: Create `agent-service/src/tools/knowledge.ts`**

```typescript
import { SearchClient, AzureKeyCredential } from '@azure/search-documents';
import { DefaultAzureCredential } from '@azure/identity';

export interface KnowledgeResult {
  content?: string;
  source?: string;
  score?: number;
  error?: string;
}

interface KnowledgeOpts {
  searchEndpoint: string;
  indexName: string;
  apiKey?: string;
}

export async function searchKnowledge(
  query: string,
  opts: KnowledgeOpts
): Promise<KnowledgeResult[]> {
  const credential = opts.apiKey
    ? new AzureKeyCredential(opts.apiKey)
    : new DefaultAzureCredential();

  const client = new SearchClient<Record<string, unknown>>(
    opts.searchEndpoint,
    opts.indexName,
    credential
  );

  try {
    const response = await client.search(query, {
      top: 5,
      queryType: 'semantic',
      semanticSearchOptions: { configurationName: 'default' },
    });

    const results: KnowledgeResult[] = [];
    for await (const result of response.results) {
      const doc = result.document as Record<string, unknown>;
      results.push({
        content: String(doc['content'] ?? ''),
        source: String(doc['sourcefile'] ?? doc['metadata_storage_name'] ?? ''),
        score: result.score ?? undefined,
      });
    }
    return results;
  } catch (err: unknown) {
    const e = err as { statusCode?: number; message?: string };
    if (e.statusCode === 401 || e.statusCode === 403) {
      return [{ error: 'AuthError: insufficient permissions for knowledge base index' }];
    }
    if (e.statusCode === 404) {
      return [{ error: 'NotFound: knowledge base index does not exist' }];
    }
    if (e.statusCode === 429) {
      return [{ error: 'RateLimited: AI Search rate limit hit, retry shortly' }];
    }
    return [{ error: `ServiceError: ${e.message ?? 'AI Search unavailable'}` }];
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

```bash
npx jest tests/tools/knowledge.test.ts --no-coverage
```

Expected: PASS — 3 tests

- [ ] **Step 5: Commit**

```bash
git add agent-service/src/tools/knowledge.ts agent-service/tests/tools/knowledge.test.ts
git commit -m "feat: add searchKnowledge tool with AI Search / Foundry IQ integration"
```

---

## Task 4: Lists Tool (Graph API — Read + Write)

**Files:**
- Create: `agent-service/src/tools/lists.ts`
- Create: `agent-service/tests/tools/lists.test.ts`

- [ ] **Step 1: Write the failing test**

Create `agent-service/tests/tools/lists.test.ts`:

```typescript
import { getListItems, createListItem, updateListItem } from '../../src/tools/lists';
import { Client } from '@microsoft/microsoft-graph-client';

const mockApi = jest.fn();
const mockGet = jest.fn();
const mockPost = jest.fn();
const mockPatch = jest.fn();

const fakeClient = {
  api: mockApi,
} as unknown as Client;

beforeEach(() => {
  jest.clearAllMocks();
  mockApi.mockReturnValue({
    get: mockGet,
    post: mockPost,
    patch: mockPatch,
    filter: jest.fn().mockReturnThis(),
    top: jest.fn().mockReturnThis(),
  });
});

describe('getListItems', () => {
  it('returns list items from SharePoint', async () => {
    mockGet.mockResolvedValueOnce({
      value: [
        { id: '1', fields: { Title: 'Task A', Status: 'Active' } },
        { id: '2', fields: { Title: 'Task B', Status: 'Done' } },
      ],
    });

    const items = await getListItems(fakeClient, 'site-id', 'list-id');
    expect(items).toHaveLength(2);
    expect(items[0]).toEqual({ id: '1', fields: { Title: 'Task A', Status: 'Active' } });
  });

  it('returns AuthError on 403', async () => {
    mockGet.mockRejectedValueOnce({ statusCode: 403 });
    const result = await getListItems(fakeClient, 'site-id', 'list-id');
    expect(result).toEqual([{ error: 'AuthError: insufficient permissions for this SharePoint site' }]);
  });
});

describe('createListItem', () => {
  it('creates a list item and returns the created item', async () => {
    mockPost.mockResolvedValueOnce({ id: '3', fields: { Title: 'New Task' } });

    const item = await createListItem(fakeClient, 'site-id', 'list-id', { Title: 'New Task' });
    expect(item).toEqual({ id: '3', fields: { Title: 'New Task' } });
  });
});

describe('updateListItem', () => {
  it('updates a list item and returns confirmation', async () => {
    mockPatch.mockResolvedValueOnce({ id: '1', fields: { Status: 'Done' } });

    const result = await updateListItem(fakeClient, 'site-id', 'list-id', '1', { Status: 'Done' });
    expect(result).toEqual({ id: '1', fields: { Status: 'Done' } });
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx jest tests/tools/lists.test.ts --no-coverage
```

Expected: FAIL — `Cannot find module '../../src/tools/lists'`

- [ ] **Step 3: Create `agent-service/src/tools/lists.ts`**

```typescript
import { Client } from '@microsoft/microsoft-graph-client';

type ListItem = Record<string, unknown>;
type ListItemResult = ListItem | { error: string };

function classifyGraphError(err: unknown): string {
  const e = err as { statusCode?: number; message?: string };
  if (e.statusCode === 401 || e.statusCode === 403)
    return 'AuthError: insufficient permissions for this SharePoint site';
  if (e.statusCode === 404)
    return 'NotFound: list or site does not exist';
  if (e.statusCode === 429)
    return 'RateLimited: Graph API rate limit hit, retry shortly';
  return `ServiceError: ${e.message ?? 'Graph API unavailable'}`;
}

export async function getListItems(
  client: Client,
  siteId: string,
  listId: string,
  filter?: string
): Promise<ListItemResult[]> {
  try {
    let req = client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .top(100);
    if (filter) req = req.filter(filter);
    const res = await req.get();
    return (res.value ?? []) as ListItem[];
  } catch (err) {
    return [{ error: classifyGraphError(err) }];
  }
}

export async function createListItem(
  client: Client,
  siteId: string,
  listId: string,
  fields: Record<string, unknown>
): Promise<ListItemResult> {
  try {
    return await client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .post({ fields });
  } catch (err) {
    return { error: classifyGraphError(err) };
  }
}

export async function updateListItem(
  client: Client,
  siteId: string,
  listId: string,
  itemId: string,
  fields: Record<string, unknown>
): Promise<ListItemResult> {
  try {
    return await client
      .api(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`)
      .patch(fields);
  } catch (err) {
    return { error: classifyGraphError(err) };
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

```bash
npx jest tests/tools/lists.test.ts --no-coverage
```

Expected: PASS — 5 tests

- [ ] **Step 5: Commit**

```bash
git add agent-service/src/tools/lists.ts agent-service/tests/tools/lists.test.ts
git commit -m "feat: add SharePoint list tools (getListItems, createListItem, updateListItem)"
```

---

## Task 5: Documents, Pages, and Notifications Tools

**Files:**
- Create: `agent-service/src/tools/documents.ts`
- Create: `agent-service/src/tools/pages.ts`
- Create: `agent-service/src/tools/notifications.ts`
- Create: `agent-service/tests/tools/documents.test.ts`
- Create: `agent-service/tests/tools/pages.test.ts`
- Create: `agent-service/tests/tools/notifications.test.ts`

- [ ] **Step 1: Write failing tests for all three tools**

Create `agent-service/tests/tools/documents.test.ts`:

```typescript
import { uploadDocument } from '../../src/tools/documents';
import { Client } from '@microsoft/microsoft-graph-client';

const mockPut = jest.fn();
const mockApi = jest.fn().mockReturnValue({ put: mockPut });
const fakeClient = { api: mockApi } as unknown as Client;

describe('uploadDocument', () => {
  it('uploads a document and returns the created item', async () => {
    mockPut.mockResolvedValueOnce({ id: 'file-1', name: 'report.pdf', webUrl: 'https://sp/report.pdf' });

    const result = await uploadDocument(
      fakeClient,
      'drive-id',
      'report.pdf',
      Buffer.from('pdf content'),
      'application/pdf'
    );

    expect(result).toEqual({ id: 'file-1', name: 'report.pdf', webUrl: 'https://sp/report.pdf' });
    expect(mockApi).toHaveBeenCalledWith('/drives/drive-id/root:/report.pdf:/content');
  });

  it('returns AuthError on 403', async () => {
    mockPut.mockRejectedValueOnce({ statusCode: 403 });
    const result = await uploadDocument(fakeClient, 'drive-id', 'x.pdf', Buffer.from(''), 'application/pdf');
    expect(result).toEqual({ error: 'AuthError: insufficient permissions for this SharePoint site' });
  });
});
```

Create `agent-service/tests/tools/pages.test.ts`:

```typescript
import { createPage } from '../../src/tools/pages';
import { Client } from '@microsoft/microsoft-graph-client';

const mockPost = jest.fn();
const mockApi = jest.fn().mockReturnValue({ post: mockPost });
const fakeClient = { api: mockApi } as unknown as Client;

describe('createPage', () => {
  it('creates a SharePoint page and returns the page URL', async () => {
    mockPost.mockResolvedValueOnce({ id: 'page-1', title: 'Summary', webUrl: 'https://sp/Summary' });

    const result = await createPage(fakeClient, 'site-id', 'Summary', '<p>Hello</p>');

    expect(result).toEqual({ id: 'page-1', title: 'Summary', webUrl: 'https://sp/Summary' });
    expect(mockApi).toHaveBeenCalledWith('/sites/site-id/pages');
  });

  it('returns AuthError on 403', async () => {
    mockPost.mockRejectedValueOnce({ statusCode: 403 });
    const result = await createPage(fakeClient, 'site-id', 'Title', '<p>x</p>');
    expect(result).toEqual({ error: 'AuthError: insufficient permissions for this SharePoint site' });
  });
});
```

Create `agent-service/tests/tools/notifications.test.ts`:

```typescript
import { sendNotification } from '../../src/tools/notifications';
import { Client } from '@microsoft/microsoft-graph-client';

const mockPost = jest.fn();
const mockApi = jest.fn().mockReturnValue({ post: mockPost });
const fakeClient = { api: mockApi } as unknown as Client;

describe('sendNotification', () => {
  it('sends a mail notification and returns success', async () => {
    mockPost.mockResolvedValueOnce(undefined);

    const result = await sendNotification(
      fakeClient,
      'me',
      ['alice@contoso.com'],
      'Action required',
      'Please review the updated list.'
    );

    expect(result).toEqual({ sent: true });
    expect(mockApi).toHaveBeenCalledWith('/users/me/sendMail');
  });

  it('returns AuthError on 403', async () => {
    mockPost.mockRejectedValueOnce({ statusCode: 403 });
    const result = await sendNotification(fakeClient, 'me', ['x@y.com'], 'Sub', 'Body');
    expect(result).toEqual({ error: 'AuthError: insufficient permissions for this SharePoint site' });
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
npx jest tests/tools/documents.test.ts tests/tools/pages.test.ts tests/tools/notifications.test.ts --no-coverage
```

Expected: FAIL — all three modules not found

- [ ] **Step 3: Create `agent-service/src/tools/documents.ts`**

```typescript
import { Client } from '@microsoft/microsoft-graph-client';

type DocumentResult = Record<string, unknown> | { error: string };

function classifyGraphError(err: unknown): string {
  const e = err as { statusCode?: number; message?: string };
  if (e.statusCode === 401 || e.statusCode === 403)
    return 'AuthError: insufficient permissions for this SharePoint site';
  if (e.statusCode === 404) return 'NotFound: drive or folder does not exist';
  if (e.statusCode === 429) return 'RateLimited: Graph API rate limit hit, retry shortly';
  return `ServiceError: ${e.message ?? 'Graph API unavailable'}`;
}

export async function uploadDocument(
  client: Client,
  driveId: string,
  fileName: string,
  content: Buffer,
  contentType: string
): Promise<DocumentResult> {
  try {
    return await client
      .api(`/drives/${driveId}/root:/${fileName}:/content`)
      .header('Content-Type', contentType)
      .put(content);
  } catch (err) {
    return { error: classifyGraphError(err) };
  }
}
```

- [ ] **Step 4: Create `agent-service/src/tools/pages.ts`**

```typescript
import { Client } from '@microsoft/microsoft-graph-client';

type PageResult = Record<string, unknown> | { error: string };

function classifyGraphError(err: unknown): string {
  const e = err as { statusCode?: number; message?: string };
  if (e.statusCode === 401 || e.statusCode === 403)
    return 'AuthError: insufficient permissions for this SharePoint site';
  if (e.statusCode === 404) return 'NotFound: SharePoint site does not exist';
  if (e.statusCode === 429) return 'RateLimited: Graph API rate limit hit, retry shortly';
  return `ServiceError: ${e.message ?? 'Graph API unavailable'}`;
}

export async function createPage(
  client: Client,
  siteId: string,
  title: string,
  htmlContent: string
): Promise<PageResult> {
  try {
    return await client.api(`/sites/${siteId}/pages`).post({
      '@odata.type': '#microsoft.graph.sitePage',
      title,
      pageLayout: 'article',
      canvasLayout: {
        horizontalSections: [
          {
            layout: 'oneColumn',
            columns: [
              {
                webparts: [
                  {
                    '@odata.type': '#microsoft.graph.textWebPart',
                    innerHtml: htmlContent,
                  },
                ],
              },
            ],
          },
        ],
      },
    });
  } catch (err) {
    return { error: classifyGraphError(err) };
  }
}
```

- [ ] **Step 5: Create `agent-service/src/tools/notifications.ts`**

```typescript
import { Client } from '@microsoft/microsoft-graph-client';

type NotificationResult = { sent: boolean } | { error: string };

function classifyGraphError(err: unknown): string {
  const e = err as { statusCode?: number; message?: string };
  if (e.statusCode === 401 || e.statusCode === 403)
    return 'AuthError: insufficient permissions for this SharePoint site';
  if (e.statusCode === 429) return 'RateLimited: Graph API rate limit hit, retry shortly';
  return `ServiceError: ${e.message ?? 'Graph API unavailable'}`;
}

export async function sendNotification(
  client: Client,
  senderUserId: string,
  toAddresses: string[],
  subject: string,
  body: string
): Promise<NotificationResult> {
  try {
    await client.api(`/users/${senderUserId}/sendMail`).post({
      message: {
        subject,
        body: { contentType: 'Text', content: body },
        toRecipients: toAddresses.map((a) => ({
          emailAddress: { address: a },
        })),
      },
    });
    return { sent: true };
  } catch (err) {
    return { error: classifyGraphError(err) };
  }
}
```

- [ ] **Step 6: Run all tool tests**

```bash
npx jest tests/tools/ --no-coverage
```

Expected: PASS — all 10 tool tests across 5 files

- [ ] **Step 7: Commit**

```bash
git add agent-service/src/tools/ agent-service/tests/tools/
git commit -m "feat: add documents, pages, and notifications tools"
```

---

## Task 6: Express Server (Tool Endpoint Wiring)

**Files:**
- Create: `agent-service/src/server.ts`

- [ ] **Step 1: Create `agent-service/src/server.ts`**

```typescript
import express, { Request, Response } from 'express';
import { loadConfig } from './config';
import { createGraphClient } from './auth/graphClient';
import { searchKnowledge } from './tools/knowledge';
import { getListItems, createListItem, updateListItem } from './tools/lists';
import { uploadDocument } from './tools/documents';
import { createPage } from './tools/pages';
import { sendNotification } from './tools/notifications';

const app = express();
app.use(express.json({ limit: '10mb' }));

const config = loadConfig();
const graphClient = createGraphClient({
  tenantId: config.azure.tenantId,
  clientId: config.azure.clientId,
  clientSecret: config.azure.clientSecret,
});

// Health check
app.get('/health', (_req: Request, res: Response) => {
  res.json({ status: 'ok', model: config.reasoningModel });
});

// Tool: searchKnowledge
app.post('/tools/searchKnowledge', async (req: Request, res: Response) => {
  const { query } = req.body as { query: string };
  const results = await searchKnowledge(query, {
    searchEndpoint: config.azure.searchEndpoint,
    indexName: config.azure.searchIndexName,
  });
  res.json({ results });
});

// Tool: getListItems
app.post('/tools/getListItems', async (req: Request, res: Response) => {
  const { siteId, listId, filter } = req.body as {
    siteId: string;
    listId: string;
    filter?: string;
  };
  const items = await getListItems(graphClient, siteId, listId, filter);
  res.json({ items });
});

// Tool: createListItem
app.post('/tools/createListItem', async (req: Request, res: Response) => {
  const { siteId, listId, fields } = req.body as {
    siteId: string;
    listId: string;
    fields: Record<string, unknown>;
  };
  const item = await createListItem(graphClient, siteId, listId, fields);
  res.json({ item });
});

// Tool: updateListItem
app.post('/tools/updateListItem', async (req: Request, res: Response) => {
  const { siteId, listId, itemId, fields } = req.body as {
    siteId: string;
    listId: string;
    itemId: string;
    fields: Record<string, unknown>;
  };
  const item = await updateListItem(graphClient, siteId, listId, itemId, fields);
  res.json({ item });
});

// Tool: uploadDocument
app.post('/tools/uploadDocument', async (req: Request, res: Response) => {
  const { driveId, fileName, contentBase64, contentType } = req.body as {
    driveId: string;
    fileName: string;
    contentBase64: string;
    contentType: string;
  };
  const content = Buffer.from(contentBase64, 'base64');
  const result = await uploadDocument(graphClient, driveId, fileName, content, contentType);
  res.json({ result });
});

// Tool: createPage
app.post('/tools/createPage', async (req: Request, res: Response) => {
  const { siteId, title, htmlContent } = req.body as {
    siteId: string;
    title: string;
    htmlContent: string;
  };
  const result = await createPage(graphClient, siteId, title, htmlContent);
  res.json({ result });
});

// Tool: sendNotification
app.post('/tools/sendNotification', async (req: Request, res: Response) => {
  const { senderUserId, toAddresses, subject, body } = req.body as {
    senderUserId: string;
    toAddresses: string[];
    subject: string;
    body: string;
  };
  const result = await sendNotification(graphClient, senderUserId, toAddresses, subject, body);
  res.json({ result });
});

app.listen(config.port, () => {
  console.log(`Agent service running on port ${config.port}`);
  console.log(`Reasoning model: ${config.reasoningModel}`);
});

export default app;
```

- [ ] **Step 2: Do a build check**

```bash
npm run build
```

Expected: `dist/` built with no TypeScript errors.

- [ ] **Step 3: Commit**

```bash
git add agent-service/src/server.ts
git commit -m "feat: add Express server with all tool endpoints"
```

---

## Task 7: Dockerfile

**Files:**
- Create: `agent-service/Dockerfile`
- Create: `agent-service/.dockerignore`

- [ ] **Step 1: Create `agent-service/Dockerfile`**

```dockerfile
FROM node:20-alpine AS builder
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY tsconfig.json ./
COPY src ./src
RUN npm run build

FROM node:20-alpine AS runtime
WORKDIR /app
ENV NODE_ENV=production
COPY package*.json ./
RUN npm ci --omit=dev
COPY --from=builder /app/dist ./dist
EXPOSE 3000
CMD ["node", "dist/server.js"]
```

- [ ] **Step 2: Create `agent-service/.dockerignore`**

```
node_modules
dist
tests
.env
.env.*
*.test.ts
```

- [ ] **Step 3: Build and verify Docker image locally**

```bash
cd agent-service
docker build -t sharepoint-foundry-agent:local .
```

Expected: Image built successfully, `Successfully tagged sharepoint-foundry-agent:local`

- [ ] **Step 4: Commit**

```bash
git add agent-service/Dockerfile agent-service/.dockerignore
git commit -m "feat: add multi-stage Dockerfile for agent service"
```

---

## Task 8: ACA Infrastructure (Bicep)

**Files:**
- Create: `infra/agent.bicep`

- [ ] **Step 1: Create `infra/agent.bicep`**

```bicep
// ===============================================
// SharePoint Foundry Agent — ACA Infrastructure
// Provisions: Container Apps Environment, ACA App,
//             System-assigned Managed Identity,
//             ACR pull role assignment
// ===============================================

@description('Resource name prefix (match main.bicep)')
param resourcePrefix string = 'iqseries'

@description('Azure region')
param location string = 'eastus2'

@description('Container Registry login server (e.g. myregistry.azurecr.io)')
param acrLoginServer string

@description('Docker image tag to deploy')
param imageTag string = 'latest'

@description('Azure AI endpoint from main.bicep outputs')
param azureAiEndpoint string

@description('Foundry project name from main.bicep')
param foundryProjectName string

@description('AI Search endpoint from main.bicep outputs')
param searchEndpoint string

@description('AI Search index name (knowledge base)')
param searchIndexName string

@description('Reasoning model string (e.g. azure/gpt-4o)')
param reasoningModel string = 'azure/gpt-4o'

var uniqueSuffix = uniqueString(resourceGroup().id)
var containerAppName = '${resourcePrefix}-agent-${uniqueSuffix}'
var acrPullRoleId = '7f951dda-4ed3-4680-a7ca-43fe172d538d'

// -----------------------------------------------
// Container Apps Environment
// -----------------------------------------------

resource acaEnvironment 'Microsoft.App/managedEnvironments@2024-03-01' = {
  name: '${resourcePrefix}-aca-env-${uniqueSuffix}'
  location: location
  properties: {
    zoneRedundant: false
  }
}

// -----------------------------------------------
// Container App
// -----------------------------------------------

resource containerApp 'Microsoft.App/containerApps@2024-03-01' = {
  name: containerAppName
  location: location
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    managedEnvironmentId: acaEnvironment.id
    configuration: {
      ingress: {
        external: true
        targetPort: 3000
        transport: 'auto'
      }
      registries: [
        {
          server: acrLoginServer
          identity: 'system'
        }
      ]
    }
    template: {
      containers: [
        {
          name: 'agent-service'
          image: '${acrLoginServer}/sharepoint-foundry-agent:${imageTag}'
          resources: {
            cpu: json('0.5')
            memory: '1Gi'
          }
          env: [
            { name: 'AZURE_AI_ENDPOINT', value: azureAiEndpoint }
            { name: 'FOUNDRY_PROJECT_NAME', value: foundryProjectName }
            { name: 'AZURE_SEARCH_ENDPOINT', value: searchEndpoint }
            { name: 'AZURE_SEARCH_INDEX_NAME', value: searchIndexName }
            { name: 'REASONING_MODEL', value: reasoningModel }
            { name: 'PORT', value: '3000' }
          ]
        }
      ]
      scale: {
        minReplicas: 1
        maxReplicas: 3
      }
    }
  }
}

// -----------------------------------------------
// Grant ACA identity ACR pull permission
// -----------------------------------------------

resource acrResource 'Microsoft.ContainerRegistry/registries@2023-07-01' existing = {
  name: split(acrLoginServer, '.')[0]
}

resource acrPullRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(resourceGroup().id, containerApp.name, acrPullRoleId)
  scope: acrResource
  properties: {
    principalId: containerApp.identity.principalId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', acrPullRoleId)
  }
}

// -----------------------------------------------
// Outputs
// -----------------------------------------------

output acaEndpoint string = 'https://${containerApp.properties.configuration.ingress.fqdn}'
output acaManagedIdentityPrincipalId string = containerApp.identity.principalId
```

- [ ] **Step 2: Validate Bicep**

```bash
az bicep build --file infra/agent.bicep
```

Expected: `infra/agent.json` generated with no errors.

- [ ] **Step 3: Commit**

```bash
git add infra/agent.bicep infra/agent.json
git commit -m "feat: add ACA infrastructure Bicep for agent service"
```

---

## Task 9: Foundry Agent Registration Script

**Files:**
- Create: `scripts/register-agent.ts`
- Create: `scripts/package.json`
- Create: `scripts/tsconfig.json`

- [ ] **Step 1: Create `scripts/package.json`**

```json
{
  "name": "iq-scripts",
  "private": true,
  "scripts": {
    "register": "ts-node register-agent.ts"
  },
  "dependencies": {
    "@azure/ai-projects": "^1.0.0",
    "@azure/identity": "^4.4.0"
  },
  "devDependencies": {
    "ts-node": "^10.9.2",
    "typescript": "^5.5.0",
    "@types/node": "^20.14.0"
  }
}
```

- [ ] **Step 2: Create `scripts/tsconfig.json`**

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "commonjs",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true
  }
}
```

- [ ] **Step 3: Create `scripts/register-agent.ts`**

```typescript
import { AIProjectClient } from '@azure/ai-projects';
import { DefaultAzureCredential } from '@azure/identity';

const AZURE_AI_ENDPOINT = process.env.AZURE_AI_ENDPOINT!;
const FOUNDRY_PROJECT_NAME = process.env.FOUNDRY_PROJECT_NAME!;
const ACA_ENDPOINT = process.env.ACA_ENDPOINT!;
const REASONING_MODEL = process.env.REASONING_MODEL ?? 'azure/gpt-4o';

const AGENT_NAME = 'sharepoint-iq-agent';
const AGENT_DESCRIPTION = 'SharePoint IQ Agent — reads SharePoint documents, lists, and pages via Foundry IQ, and can create list items, upload documents, create pages, and send notifications.';

const SYSTEM_PROMPT = `You are the SharePoint IQ Agent. You help users find information from SharePoint and take actions on their behalf.

You have access to these tools:
- searchKnowledge: Search the Foundry IQ knowledge base for documents and pages
- getListItems: Read items from a SharePoint list
- createListItem: Create a new item in a SharePoint list
- updateListItem: Update an existing item in a SharePoint list
- uploadDocument: Upload a file to a SharePoint document library
- createPage: Create a new SharePoint page
- sendNotification: Send an email notification via Microsoft Graph

Always confirm actions with the user before writing or updating data.
When you get an error from a tool, explain it clearly to the user and suggest next steps.`;

const TOOL_DEFINITIONS = [
  {
    type: 'function',
    function: {
      name: 'searchKnowledge',
      description: 'Search the Foundry IQ knowledge base for relevant documents, pages, and content.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'The search query' },
        },
        required: ['query'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'getListItems',
      description: 'Get items from a SharePoint list, optionally filtered.',
      parameters: {
        type: 'object',
        properties: {
          siteId: { type: 'string', description: 'SharePoint site ID' },
          listId: { type: 'string', description: 'SharePoint list ID or name' },
          filter: { type: 'string', description: 'OData filter expression (optional)' },
        },
        required: ['siteId', 'listId'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'createListItem',
      description: 'Create a new item in a SharePoint list.',
      parameters: {
        type: 'object',
        properties: {
          siteId: { type: 'string' },
          listId: { type: 'string' },
          fields: { type: 'object', description: 'Key-value pairs for the new list item fields' },
        },
        required: ['siteId', 'listId', 'fields'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'updateListItem',
      description: 'Update fields on an existing SharePoint list item.',
      parameters: {
        type: 'object',
        properties: {
          siteId: { type: 'string' },
          listId: { type: 'string' },
          itemId: { type: 'string' },
          fields: { type: 'object' },
        },
        required: ['siteId', 'listId', 'itemId', 'fields'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'uploadDocument',
      description: 'Upload a file to a SharePoint document library.',
      parameters: {
        type: 'object',
        properties: {
          driveId: { type: 'string', description: 'SharePoint drive ID' },
          fileName: { type: 'string' },
          contentBase64: { type: 'string', description: 'Base64-encoded file content' },
          contentType: { type: 'string', description: 'MIME type (e.g. application/pdf)' },
        },
        required: ['driveId', 'fileName', 'contentBase64', 'contentType'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'createPage',
      description: 'Create a new SharePoint page with HTML content.',
      parameters: {
        type: 'object',
        properties: {
          siteId: { type: 'string' },
          title: { type: 'string' },
          htmlContent: { type: 'string', description: 'HTML body content for the page' },
        },
        required: ['siteId', 'title', 'htmlContent'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'sendNotification',
      description: 'Send an email notification via Microsoft Graph.',
      parameters: {
        type: 'object',
        properties: {
          senderUserId: { type: 'string', description: 'UPN or user ID of the sender' },
          toAddresses: { type: 'array', items: { type: 'string' }, description: 'Recipient email addresses' },
          subject: { type: 'string' },
          body: { type: 'string' },
        },
        required: ['senderUserId', 'toAddresses', 'subject', 'body'],
      },
    },
  },
];

async function main() {
  console.log(`Registering agent "${AGENT_NAME}" in project "${FOUNDRY_PROJECT_NAME}"...`);
  console.log(`Endpoint: ${AZURE_AI_ENDPOINT}`);
  console.log(`ACA tool endpoint: ${ACA_ENDPOINT}`);
  console.log(`Reasoning model: ${REASONING_MODEL}`);

  const client = new AIProjectClient(
    AZURE_AI_ENDPOINT,
    FOUNDRY_PROJECT_NAME,
    new DefaultAzureCredential()
  );

  const agents = client.agents;

  // Check if agent already exists
  const existing = await agents.listAgents();
  const found = existing.data?.find((a: { name: string }) => a.name === AGENT_NAME);

  if (found) {
    console.log(`Agent "${AGENT_NAME}" already exists (id: ${found.id}), updating...`);
    await agents.updateAgent(found.id, {
      model: REASONING_MODEL,
      name: AGENT_NAME,
      description: AGENT_DESCRIPTION,
      instructions: SYSTEM_PROMPT,
      tools: TOOL_DEFINITIONS as never,
      toolResources: {
        azureAiSearch: {
          indexList: [],
        },
      },
      metadata: { acaEndpoint: ACA_ENDPOINT },
    });
    console.log(`Agent updated successfully.`);
  } else {
    const agent = await agents.createAgent({
      model: REASONING_MODEL,
      name: AGENT_NAME,
      description: AGENT_DESCRIPTION,
      instructions: SYSTEM_PROMPT,
      tools: TOOL_DEFINITIONS as never,
      metadata: { acaEndpoint: ACA_ENDPOINT },
    });
    console.log(`Agent created successfully. ID: ${agent.id}`);
  }

  console.log('Done. Open Azure AI Foundry portal → Agents to test in the playground.');
}

main().catch((err) => {
  console.error('Registration failed:', err.message ?? err);
  process.exit(1);
});
```

- [ ] **Step 4: Install scripts dependencies**

```bash
cd scripts
npm install
```

- [ ] **Step 5: Compile check**

```bash
npx tsc --noEmit
```

Expected: No TypeScript errors.

- [ ] **Step 6: Commit**

```bash
git add scripts/
git commit -m "feat: add Foundry Agent registration script"
```

---

## Task 10: GitHub Actions Deployment Pipeline

**Files:**
- Create: `.github/workflows/deploy-agent.yml`

- [ ] **Step 1: Create `.github/workflows/deploy-agent.yml`**

```yaml
name: Deploy SharePoint Foundry Agent

on:
  push:
    branches: [main]
    paths:
      - 'agent-service/**'
      - 'scripts/register-agent.ts'
      - 'infra/agent.bicep'
  workflow_dispatch:
    inputs:
      reasoning_model:
        description: 'Reasoning model (e.g. azure/gpt-4o, google/gemini-2.0-flash)'
        required: false
        default: 'azure/gpt-4o'

env:
  ACR_NAME: ${{ secrets.ACR_NAME }}
  RESOURCE_GROUP: ${{ secrets.RESOURCE_GROUP }}
  AZURE_AI_ENDPOINT: ${{ secrets.AZURE_AI_ENDPOINT }}
  FOUNDRY_PROJECT_NAME: ${{ secrets.FOUNDRY_PROJECT_NAME }}
  AZURE_SEARCH_ENDPOINT: ${{ secrets.AZURE_SEARCH_ENDPOINT }}
  AZURE_SEARCH_INDEX_NAME: ${{ secrets.AZURE_SEARCH_INDEX_NAME }}
  REASONING_MODEL: ${{ github.event.inputs.reasoning_model || 'azure/gpt-4o' }}

jobs:
  build-and-deploy:
    runs-on: ubuntu-latest
    permissions:
      id-token: write
      contents: read

    steps:
      - uses: actions/checkout@v4

      - name: Azure Login (OIDC)
        uses: azure/login@v2
        with:
          client-id: ${{ secrets.AZURE_CLIENT_ID }}
          tenant-id: ${{ secrets.AZURE_TENANT_ID }}
          subscription-id: ${{ secrets.AZURE_SUBSCRIPTION_ID }}

      - name: Run unit tests
        working-directory: agent-service
        run: |
          npm ci
          npm test -- --no-coverage

      - name: Build TypeScript
        working-directory: agent-service
        run: npm run build

      - name: Build & push Docker image
        run: |
          IMAGE_TAG=$(echo $GITHUB_SHA | cut -c1-8)
          az acr build \
            --registry $ACR_NAME \
            --image sharepoint-foundry-agent:$IMAGE_TAG \
            --image sharepoint-foundry-agent:latest \
            agent-service/

      - name: Deploy ACA (Bicep)
        run: |
          IMAGE_TAG=$(echo $GITHUB_SHA | cut -c1-8)
          az deployment group create \
            --resource-group $RESOURCE_GROUP \
            --template-file infra/agent.bicep \
            --parameters \
              acrLoginServer="${ACR_NAME}.azurecr.io" \
              imageTag="$IMAGE_TAG" \
              azureAiEndpoint="$AZURE_AI_ENDPOINT" \
              foundryProjectName="$FOUNDRY_PROJECT_NAME" \
              searchEndpoint="$AZURE_SEARCH_ENDPOINT" \
              searchIndexName="$AZURE_SEARCH_INDEX_NAME" \
              reasoningModel="$REASONING_MODEL"

      - name: Get ACA endpoint
        id: aca
        run: |
          ACA_ENDPOINT=$(az deployment group show \
            --resource-group $RESOURCE_GROUP \
            --name agent \
            --query properties.outputs.acaEndpoint.value -o tsv)
          echo "endpoint=$ACA_ENDPOINT" >> $GITHUB_OUTPUT

      - name: Register Foundry Agent
        working-directory: scripts
        run: |
          npm ci
          ACA_ENDPOINT=${{ steps.aca.outputs.endpoint }} npx ts-node register-agent.ts
```

- [ ] **Step 2: Add required GitHub secrets (manual step)**

In your GitHub repo → Settings → Secrets → Actions, add:

| Secret | Value |
|---|---|
| `ACR_NAME` | Your Azure Container Registry name (without `.azurecr.io`) |
| `RESOURCE_GROUP` | Your resource group (e.g. `iq-series-rg`) |
| `AZURE_AI_ENDPOINT` | From Bicep outputs: `https://<aiservices>.cognitiveservices.azure.com` |
| `FOUNDRY_PROJECT_NAME` | `iqseries-project` (or your prefix + `-project`) |
| `AZURE_SEARCH_ENDPOINT` | From Bicep outputs: `https://<search>.search.windows.net` |
| `AZURE_SEARCH_INDEX_NAME` | Your Foundry IQ knowledge base index name |
| `AZURE_CLIENT_ID` | Service principal / Managed Identity client ID for OIDC |
| `AZURE_TENANT_ID` | Your Azure tenant ID |
| `AZURE_SUBSCRIPTION_ID` | Your Azure subscription ID |

- [ ] **Step 3: Commit**

```bash
git add .github/workflows/deploy-agent.yml
git commit -m "feat: add GitHub Actions deployment pipeline for agent service"
```

---

## Task 11: Grant Microsoft Graph Permissions (Manual Azure Step)

This step cannot be automated — it requires an Azure AD admin.

- [ ] **Step 1: Go to Azure Portal → Azure Active Directory → App registrations → [your ACA Managed Identity]**

- [ ] **Step 2: Under "API permissions", click "Add a permission" → "Microsoft Graph" → "Application permissions"**

Add the following:
  - `Sites.ReadWrite.All` — read/write SharePoint sites, lists, and pages
  - `Files.ReadWrite.All` — upload documents to drives
  - `Mail.Send` — send notification emails

- [ ] **Step 3: Click "Grant admin consent" for your organization**

- [ ] **Step 4: Verify in the ACA container logs that Graph calls succeed**

```bash
az containerapp logs show \
  --name <your-aca-app-name> \
  --resource-group iq-series-rg \
  --follow
```

Expected: No `401`/`403` errors when the agent calls Graph tools.

---

## Task 12: End-to-End Foundry Playground Test

- [ ] **Step 1: Open Azure AI Foundry portal**

Navigate to: `https://ai.azure.com` → your project → **Agents**

- [ ] **Step 2: Find `sharepoint-iq-agent` and open it**

Verify:
- Model shows your `REASONING_MODEL` value
- All 7 tools are listed (searchKnowledge, getListItems, createListItem, updateListItem, uploadDocument, createPage, sendNotification)

- [ ] **Step 3: Test knowledge retrieval in the playground**

Prompt: `What documents do we have about [a topic in your knowledge base]?`

Expected: Agent calls `searchKnowledge`, returns grounded answer with source citations.

- [ ] **Step 4: Test list read**

Prompt: `Show me all items in the Projects list on site [your-site-id]`

Expected: Agent calls `getListItems`, formats items as a readable list.

- [ ] **Step 5: Test list write**

Prompt: `Create a new task called "Review Q1 Report" in the Tasks list on site [your-site-id]`

Expected: Agent confirms intent, calls `createListItem`, confirms success.

- [ ] **Step 6: Test model switching**

Re-deploy with `REASONING_MODEL=google/gemini-2.0-flash` (via `workflow_dispatch` in GitHub Actions), re-run the same prompts.

Expected: Same quality responses from a different model — no code changes needed.

---

## Self-Review Checklist

- **Spec coverage:** Architecture ✓, 7 tools ✓, configurable model ✓, Managed Identity + App Reg auth ✓, error handling ✓, ACA deployment ✓, Foundry registration ✓, GitHub Actions ✓, playground testing ✓
- **No placeholders:** All code blocks complete, no TBDs
- **Type consistency:** `classifyGraphError` used consistently across lists/documents/pages/notifications; `KnowledgeResult` used across knowledge.ts and its test; `loadConfig()` → `Config` used in server.ts
- **Graph permissions manual step:** Documented as Task 11 with exact permissions
