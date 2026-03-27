import { AIProjectClient } from '@azure/ai-projects';
import { DefaultAzureCredential } from '@azure/identity';

const AZURE_AI_ENDPOINT = process.env.AZURE_AI_ENDPOINT!;
const FOUNDRY_PROJECT_NAME = process.env.FOUNDRY_PROJECT_NAME!;
const ACA_ENDPOINT = process.env.ACA_ENDPOINT!;
const REASONING_MODEL = process.env.REASONING_MODEL ?? 'azure/gpt-4o';

const AGENT_NAME = 'sharepoint-iq-agent';
const AGENT_DESCRIPTION = 'SharePoint IQ Agent — reads SharePoint documents, lists, and pages via Foundry IQ, and can create list items, upload documents, create pages, and send notifications.';

const SYSTEM_PROMPT = `You are the SharePoint IQ Agent. You help users find information from SharePoint and take actions on their behalf.

You have access to these tools:
- searchKnowledge: Search the Foundry IQ knowledge base for documents and pages
- getListItems: Read items from a SharePoint list
- createListItem: Create a new item in a SharePoint list
- updateListItem: Update an existing item in a SharePoint list
- uploadDocument: Upload a file to a SharePoint document library
- createPage: Create a new SharePoint page
- sendNotification: Send an email notification via Microsoft Graph

Always confirm actions with the user before writing or updating data.
When you get an error from a tool, explain it clearly to the user and suggest next steps.`;

const TOOL_DEFINITIONS = [
  {
    type: 'function',
    function: {
      name: 'searchKnowledge',
      description: 'Search the Foundry IQ knowledge base for relevant documents, pages, and content.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'The search query' },
        },
        required: ['query'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'getListItems',
      description: 'Get items from a SharePoint list, optionally filtered.',
      parameters: {
        type: 'object',
        properties: {
          siteId: { type: 'string', description: 'SharePoint site ID' },
          listId: { type: 'string', description: 'SharePoint list ID or name' },
          filter: { type: 'string', description: 'OData filter expression (optional)' },
        },
        required: ['siteId', 'listId'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'createListItem',
      description: 'Create a new item in a SharePoint list.',
      parameters: {
        type: 'object',
        properties: {
          siteId: { type: 'string' },
          listId: { type: 'string' },
          fields: { type: 'object', description: 'Key-value pairs for the new list item fields' },
        },
        required: ['siteId', 'listId', 'fields'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'updateListItem',
      description: 'Update fields on an existing SharePoint list item.',
      parameters: {
        type: 'object',
        properties: {
          siteId: { type: 'string' },
          listId: { type: 'string' },
          itemId: { type: 'string' },
          fields: { type: 'object' },
        },
        required: ['siteId', 'listId', 'itemId', 'fields'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'uploadDocument',
      description: 'Upload a file to a SharePoint document library.',
      parameters: {
        type: 'object',
        properties: {
          driveId: { type: 'string', description: 'SharePoint drive ID' },
          fileName: { type: 'string' },
          contentBase64: { type: 'string', description: 'Base64-encoded file content' },
          contentType: { type: 'string', description: 'MIME type (e.g. application/pdf)' },
        },
        required: ['driveId', 'fileName', 'contentBase64', 'contentType'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'createPage',
      description: 'Create a new SharePoint page with HTML content.',
      parameters: {
        type: 'object',
        properties: {
          siteId: { type: 'string' },
          title: { type: 'string' },
          htmlContent: { type: 'string', description: 'HTML body content for the page' },
        },
        required: ['siteId', 'title', 'htmlContent'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'sendNotification',
      description: 'Send an email notification via Microsoft Graph.',
      parameters: {
        type: 'object',
        properties: {
          senderUserId: { type: 'string', description: 'UPN or user ID of the sender' },
          toAddresses: { type: 'array', items: { type: 'string' }, description: 'Recipient email addresses' },
          subject: { type: 'string' },
          body: { type: 'string' },
        },
        required: ['senderUserId', 'toAddresses', 'subject', 'body'],
      },
    },
  },
];

async function main() {
  console.log(`Registering agent "${AGENT_NAME}" in project "${FOUNDRY_PROJECT_NAME}"...`);
  console.log(`Endpoint: ${AZURE_AI_ENDPOINT}`);
  console.log(`ACA tool endpoint: ${ACA_ENDPOINT}`);
  console.log(`Reasoning model: ${REASONING_MODEL}`);

  const client = new AIProjectClient(
    AZURE_AI_ENDPOINT,
    new DefaultAzureCredential()
  );

  const agents = client.agents;

  // Check if agent already exists
  let found: { id: string; name: string | null } | undefined;
  for await (const agent of agents.listAgents()) {
    if (agent.name === AGENT_NAME) {
      found = agent;
      break;
    }
  }

  if (found) {
    console.log(`Agent "${AGENT_NAME}" already exists (id: ${found.id}), updating...`);
    await agents.updateAgent(found.id, {
      model: REASONING_MODEL,
      name: AGENT_NAME,
      description: AGENT_DESCRIPTION,
      instructions: SYSTEM_PROMPT,
      tools: TOOL_DEFINITIONS as never,
      metadata: { acaEndpoint: ACA_ENDPOINT },
    });
    console.log(`Agent updated successfully.`);
  } else {
    const agent = await agents.createAgent(REASONING_MODEL, {
      name: AGENT_NAME,
      description: AGENT_DESCRIPTION,
      instructions: SYSTEM_PROMPT,
      tools: TOOL_DEFINITIONS as never,
      metadata: { acaEndpoint: ACA_ENDPOINT },
    });
    console.log(`Agent created successfully. ID: ${agent.id}`);
  }

  console.log('Done. Open Azure AI Foundry portal → Agents to test in the playground.');
}

main().catch((err) => {
  console.error('Registration failed:', err.message ?? err);
  process.exit(1);
});
