# SharePoint Foundry IQ Agent — Design Spec

**Date:** 2026-03-27
**Status:** Approved
**Repo:** iq-series

---

## 1. Overview

Build a TypeScript Azure AI Foundry Agent that integrates with SharePoint as both a **knowledge source** (read) and an **action target** (write). The agent is registered natively in Azure AI Foundry, backed by an Azure Container Apps (ACA) service, and testable via the Foundry portal playground.

---

## 2. Goals

- Answer questions grounded in SharePoint content: document libraries, lists, and pages
- Take actions on SharePoint: create/update list items, upload documents, create pages, send notifications
- Support configurable reasoning models (GPT-4o, Gemini, Claude, etc.) via Azure AI Gateway
- Deploy to Azure with a single pipeline step; testable immediately in Foundry playground
- Use Managed Identity in production, App Registration for local development

---

## 3. Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Azure AI Foundry                         │
│                                                             │
│   ┌─────────────────────────────────────────────────────┐  │
│   │  Foundry Agent (SharePoint IQ Agent)                │  │
│   │  - System prompt + tool definitions                 │  │
│   │  - Reasoning loop (configurable model)             │  │
│   │  - Foundry IQ Knowledge Base (AI Search index)     │  │
│   └──────────────────┬──────────────────────────────────┘  │
│                       │ tool calls                          │
└───────────────────────┼─────────────────────────────────────┘
                        │
              ┌─────────▼──────────┐
              │  ACA: agent-service│  (TypeScript / Node.js)
              │  Express HTTP API  │
              │  Tool handlers:    │
              │  - searchKnowledge │◄── Foundry IQ / AI Search
              │  - getListItems    │◄── Microsoft Graph API
              │  - createListItem  │◄── Microsoft Graph API
              │  - updateListItem  │◄── Microsoft Graph API
              │  - uploadDocument  │◄── Microsoft Graph API
              │  - createPage      │◄── Microsoft Graph API
              │  - sendNotification│◄── Microsoft Graph API
              └────────────────────┘
                        │
              ┌─────────▼──────────┐
              │  Microsoft Graph   │
              │  (SharePoint Sites,│
              │   Lists, Drives,   │
              │   Pages, Mail)     │
              └────────────────────┘
```

---

## 4. Components

| Component | Location | Responsibility |
|---|---|---|
| `agent-service/` | ACA (TypeScript) | Express API, tool handlers, Graph auth |
| `agent-service/tools/knowledge.ts` | same | AI Search / Foundry IQ read path |
| `agent-service/tools/lists.ts` | same | Graph API: get/create/update list items |
| `agent-service/tools/documents.ts` | same | Graph API: upload documents to drives |
| `agent-service/tools/pages.ts` | same | Graph API: create SharePoint pages |
| `agent-service/tools/notifications.ts` | same | Graph API: send mail/Teams notifications |
| `agent-service/auth/graphClient.ts` | same | Auth factory — Managed Identity vs App Reg |
| `agent-service/config.ts` | same | Env var loading incl. `REASONING_MODEL` |
| `infra/agent.bicep` | Azure | ACA app + identity + role assignments |
| `scripts/register-agent.ts` | local/CI | Upsert Foundry Agent via `@azure/ai-projects` |

---

## 5. Configurable Reasoning Model

The agent reads `REASONING_MODEL` from environment variables at startup. Any model available via Azure AI Gateway is valid:

```
REASONING_MODEL=azure/gpt-4o            # default
REASONING_MODEL=google/gemini-2.0-flash
REASONING_MODEL=anthropic/claude-sonnet-4-6
```

Swapping models requires only an env var change and ACA restart — no code change. The Foundry Agent registration script reads the same config and updates the model reference in Foundry accordingly.

---

## 6. Data Flows

### Read — Knowledge Q&A

```
User prompt
  → Foundry Agent reasoning loop
    → tool: searchKnowledge(query)
        → Foundry IQ AI Search index → ranked document/page chunks
    → tool: getListItems(siteId, listId, filter?)
        → Microsoft Graph → list rows as structured JSON
  → agent synthesizes answer → streams response to user
```

### Write — Actions

```
User prompt (e.g. "Create a task for John on the Projects list")
  → Foundry Agent reasoning loop
    → tool: createListItem / updateListItem / uploadDocument / createPage / sendNotification
        → Microsoft Graph API
            ← Managed Identity token (prod) or ClientSecretCredential (dev)
  → agent confirms action result → responds to user
```

---

## 7. Authentication

### Local Development
- Uses `ClientSecretCredential` from `@azure/identity`
- Requires `.env`: `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `AZURE_TENANT_ID`
- App Registration needs Graph API permissions: `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `Mail.Send`

### Production (ACA)
- Uses `DefaultAzureCredential` from `@azure/identity` — automatically picks up Managed Identity
- ACA system-assigned identity granted Graph API permissions via admin consent
- No secrets in environment — zero credential rotation needed

---

## 8. Error Handling

All tool handlers return structured results. Errors are classified and returned as readable messages so the agent can reason about failures:

| Error Code | Message Returned to Agent |
|---|---|
| 401 / 403 | `"AuthError: insufficient permissions for [resource]"` |
| 404 | `"NotFound: [list/document/site] does not exist"` |
| 429 | `"RateLimited: retry after N seconds"` |
| 5xx | `"ServiceError: upstream unavailable, try again shortly"` |

- Graph API calls: 1 automatic retry on 429/503 with exponential backoff
- AI Search: no retry (Foundry IQ handles internally)
- No raw stack traces ever reach the user — all errors surface through the agent's natural language response

---

## 9. Testing Strategy

| Layer | Tooling | Notes |
|---|---|---|
| Tool unit tests | Jest | Mock Graph API responses; test each handler in isolation |
| Auth tests | Jest | Cover both `ClientSecretCredential` and `DefaultAzureCredential` paths |
| Integration tests | Jest + real Azure | Requires `.env.test` with a dedicated SharePoint test site |
| Foundry playground | Manual | End-to-end prompt testing after each deployment |
| Model switching | Manual smoke test | Change `REASONING_MODEL`, restart ACA, run standard prompt set |

---

## 10. Deployment Pipeline

```
git push (main)
  → GitHub Actions
      → tsc build + Jest tests
      → Docker build → push to Azure Container Registry
      → az containerapp update (rolling deploy to ACA)
      → npx ts-node scripts/register-agent.ts (upsert Foundry Agent)
  → Agent live in Foundry playground
```

---

## 11. Infrastructure Changes

Extend existing `infra/` Bicep to add:

- `infra/agent.bicep` — ACA environment, ACA app, system-assigned Managed Identity
- Role assignment: `Graph API permissions` via admin consent (manual step, documented)
- Output: ACA endpoint URL (used by `register-agent.ts`)

Existing resources (AI Search, Azure OpenAI, Foundry project) remain unchanged.

---

## 12. New Directory Structure

```
iq-series/
├── agent-service/              ← NEW: TypeScript ACA service
│   ├── src/
│   │   ├── config.ts
│   │   ├── server.ts
│   │   ├── auth/
│   │   │   └── graphClient.ts
│   │   └── tools/
│   │       ├── knowledge.ts
│   │       ├── lists.ts
│   │       ├── documents.ts
│   │       ├── pages.ts
│   │       └── notifications.ts
│   ├── tests/
│   ├── Dockerfile
│   └── package.json
├── scripts/
│   └── register-agent.ts       ← NEW: Foundry Agent registration
├── infra/
│   ├── agent.bicep              ← NEW: ACA + identity infra
│   └── ... (existing unchanged)
└── docs/superpowers/specs/
    └── 2026-03-27-sharepoint-foundry-agent-design.md
```

---

## 13. Out of Scope

- Frontend chat UI (agent is tested via Foundry playground)
- SharePoint webhook/event subscriptions (push-based triggers)
- Multi-tenant SharePoint support (single tenant only)
- Fabric IQ / Work IQ integration (future episodes)
