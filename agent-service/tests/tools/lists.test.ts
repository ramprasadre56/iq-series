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
