import { uploadDocument } from '../../src/tools/documents';
import { Client } from '@microsoft/microsoft-graph-client';

const mockPut = jest.fn();
const mockHeader = jest.fn().mockReturnThis();
const mockApi = jest.fn().mockReturnValue({ put: mockPut, header: mockHeader });
const fakeClient = { api: mockApi } as unknown as Client;

describe('uploadDocument', () => {
  it('uploads a document and returns the created item', async () => {
    mockPut.mockResolvedValueOnce({ id: 'file-1', name: 'report.pdf', webUrl: 'https://sp/report.pdf' });

    const result = await uploadDocument(
      fakeClient,
      'drive-id',
      'report.pdf',
      Buffer.from('pdf content'),
      'application/pdf'
    );

    expect(result).toEqual({ id: 'file-1', name: 'report.pdf', webUrl: 'https://sp/report.pdf' });
    expect(mockApi).toHaveBeenCalledWith('/drives/drive-id/root:/report.pdf:/content');
  });

  it('returns AuthError on 403', async () => {
    mockPut.mockRejectedValueOnce({ statusCode: 403 });
    const result = await uploadDocument(fakeClient, 'drive-id', 'x.pdf', Buffer.from(''), 'application/pdf');
    expect(result).toEqual({ error: 'AuthError: insufficient permissions for this SharePoint site' });
  });
});
