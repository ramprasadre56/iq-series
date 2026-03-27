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
