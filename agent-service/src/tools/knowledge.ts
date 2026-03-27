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
