import { Client } from '@microsoft/microsoft-graph-client';

type NotificationResult = { sent: boolean } | { error: string };

function classifyGraphError(err: unknown): string {
  const e = err as { statusCode?: number; message?: string };
  if (e.statusCode === 401 || e.statusCode === 403)
    return 'AuthError: insufficient permissions for this SharePoint site';
  if (e.statusCode === 429) return 'RateLimited: Graph API rate limit hit, retry shortly';
  return `ServiceError: ${e.message ?? 'Graph API unavailable'}`;
}

export async function sendNotification(
  client: Client,
  senderUserId: string,
  toAddresses: string[],
  subject: string,
  body: string
): Promise<NotificationResult> {
  try {
    await client.api(`/users/${senderUserId}/sendMail`).post({
      message: {
        subject,
        body: { contentType: 'Text', content: body },
        toRecipients: toAddresses.map((a) => ({
          emailAddress: { address: a },
        })),
      },
    });
    return { sent: true };
  } catch (err) {
    return { error: classifyGraphError(err) };
  }
}
