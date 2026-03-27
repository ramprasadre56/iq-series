import express, { Request, Response, NextFunction } from 'express';
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

type AsyncHandler = (req: Request, res: Response, next: NextFunction) => Promise<void>;

function asyncHandler(fn: AsyncHandler) {
  return (req: Request, res: Response, next: NextFunction) => {
    fn(req, res, next).catch(next);
  };
}

// Health check
app.get('/health', (_req: Request, res: Response) => {
  res.json({ status: 'ok', model: config.reasoningModel });
});

// API key authentication (required for all /tools/* endpoints)
if (config.apiKey) {
  app.use('/tools', (req: Request, res: Response, next: NextFunction) => {
    const key = req.headers['x-api-key'];
    if (key !== config.apiKey) {
      return res.status(401).json({ error: 'Unauthorized: invalid or missing API key' });
    }
    next();
  });
}

function requireFields<T extends Record<string, unknown>>(
  body: T,
  fields: (keyof T)[]
): string | null {
  for (const field of fields) {
    if (body[field] === undefined || body[field] === null || body[field] === '') {
      return `Missing required field: ${String(field)}`;
    }
  }
  return null;
}

// Tool: searchKnowledge
app.post('/tools/searchKnowledge', asyncHandler(async (req, res) => {
  const body = req.body as { query: string };
  const err = requireFields(body, ['query']);
  if (err) { res.status(400).json({ error: err }); return; }
  const results = await searchKnowledge(body.query, {
    searchEndpoint: config.azure.searchEndpoint,
    indexName: config.azure.searchIndexName,
  });
  res.json({ results });
}));

// Tool: getListItems
app.post('/tools/getListItems', asyncHandler(async (req, res) => {
  const body = req.body as { siteId: string; listId: string; filter?: string };
  const err = requireFields(body, ['siteId', 'listId']);
  if (err) { res.status(400).json({ error: err }); return; }
  const items = await getListItems(graphClient, body.siteId, body.listId, body.filter);
  res.json({ items });
}));

// Tool: createListItem
app.post('/tools/createListItem', asyncHandler(async (req, res) => {
  const body = req.body as { siteId: string; listId: string; fields: Record<string, unknown> };
  const err = requireFields(body, ['siteId', 'listId', 'fields']);
  if (err) { res.status(400).json({ error: err }); return; }
  const item = await createListItem(graphClient, body.siteId, body.listId, body.fields);
  res.json({ item });
}));

// Tool: updateListItem
app.post('/tools/updateListItem', asyncHandler(async (req, res) => {
  const body = req.body as { siteId: string; listId: string; itemId: string; fields: Record<string, unknown> };
  const err = requireFields(body, ['siteId', 'listId', 'itemId', 'fields']);
  if (err) { res.status(400).json({ error: err }); return; }
  const item = await updateListItem(graphClient, body.siteId, body.listId, body.itemId, body.fields);
  res.json({ item });
}));

// Tool: uploadDocument
app.post('/tools/uploadDocument', asyncHandler(async (req, res) => {
  const body = req.body as { driveId: string; fileName: string; contentBase64: string; contentType: string };
  const err = requireFields(body, ['driveId', 'fileName', 'contentBase64', 'contentType']);
  if (err) { res.status(400).json({ error: err }); return; }
  // Sanitize fileName to prevent path traversal
  const safeFileName = body.fileName.replace(/[/\\]/g, '_').replace(/\.\./g, '_');
  const content = Buffer.from(body.contentBase64, 'base64');
  const result = await uploadDocument(graphClient, body.driveId, safeFileName, content, body.contentType);
  res.json({ result });
}));

// Tool: createPage
app.post('/tools/createPage', asyncHandler(async (req, res) => {
  const body = req.body as { siteId: string; title: string; htmlContent: string };
  const err = requireFields(body, ['siteId', 'title', 'htmlContent']);
  if (err) { res.status(400).json({ error: err }); return; }
  const result = await createPage(graphClient, body.siteId, body.title, body.htmlContent);
  res.json({ result });
}));

// Tool: sendNotification
app.post('/tools/sendNotification', asyncHandler(async (req, res) => {
  const body = req.body as { senderUserId: string; toAddresses: string[]; subject: string; body: string };
  const err = requireFields(body, ['senderUserId', 'toAddresses', 'subject', 'body']);
  if (err) { res.status(400).json({ error: err }); return; }
  const result = await sendNotification(graphClient, body.senderUserId, body.toAddresses, body.subject, body.body);
  res.json({ result });
}));

// Global error handler
// eslint-disable-next-line @typescript-eslint/no-unused-vars
app.use((err: Error, _req: Request, res: Response, _next: NextFunction) => {
  console.error('Unhandled error:', err.message);
  res.status(500).json({ error: 'Internal server error' });
});

app.listen(config.port, () => {
  console.log(`Agent service running on port ${config.port}`);
  console.log(`Reasoning model: ${config.reasoningModel}`);
});

export default app;
