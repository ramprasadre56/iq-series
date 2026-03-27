import { searchKnowledge, KnowledgeResult } from '../../src/tools/knowledge';
import { SearchClient } from '@azure/search-documents';

jest.mock('@azure/search-documents');

const mockSearch = jest.fn();
(SearchClient as jest.MockedClass<typeof SearchClient>).mockImplementation(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  () => ({ search: mockSearch } as any)
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
