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
