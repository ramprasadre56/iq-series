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
