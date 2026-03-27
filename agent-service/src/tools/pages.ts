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
