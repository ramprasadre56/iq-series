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
