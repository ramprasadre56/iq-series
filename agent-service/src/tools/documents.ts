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
