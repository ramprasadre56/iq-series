# SharePoint Foundry IQ Agent — Azure Deployment Plan

**Status:** Draft

---

## Azure Context
- **Subscription**: 12203a55-a187-4908-a625-1a7284605135
- **User**: ramprasadre56@gmail.com
- **Resource Group**: rg-ramprasadre56-5875
- **Location**: eastus2
- **Recipe**: Bicep + AZCLI (azd not installed; infra/*.bicep already exists)

---

## Existing Resources
| Resource | Name | Status |
|---|---|---|
| AI Services (Foundry) | ramprasadre56-5875-resource | ✅ Exists |
| Foundry Project | ramprasadre56-5875 | ✅ Exists |
| Resource Group | rg-ramprasadre56-5875 | ✅ Exists |

---

## Resources to Create
| Resource | Name | Bicep File |
|---|---|---|
| Azure AI Search (Standard) | auto-named from main.bicep | infra/main.bicep |
| Azure OpenAI | auto-named from main.bicep | infra/main.bicep |
| Azure Blob Storage | auto-named from main.bicep | infra/main.bicep |
| Azure Container Registry (Basic) | iqseriesacr{suffix} | az CLI (new) |
| Container Apps Environment | infra/agent.bicep | infra/agent.bicep |
| Container App (agent-service) | infra/agent.bicep | infra/agent.bicep |
| Entra App Registration | sharepoint-iq-agent-local | az CLI (for local dev) |

---

## Deployment Steps
- [ ] 1. Deploy infra/main.bicep → creates AI Search, OpenAI, Blob Storage
- [ ] 2. Get main.bicep outputs (search endpoint, OpenAI endpoint)
- [ ] 3. Create Azure Container Registry via az CLI
- [ ] 4. Build Docker image and push to ACR (via az acr build)
- [ ] 5. Deploy infra/agent.bicep → creates ACA Environment + App
- [ ] 6. Grant ACA Managed Identity Graph API permissions (admin consent)
- [ ] 7. Run scripts/register-agent.ts → registers Foundry Agent
- [ ] 8. Create Entra App Registration for local dev Graph auth

---

## Configuration
- **AI Services Endpoint**: https://ramprasadre56-5875-resource.cognitiveservices.azure.com/
- **Foundry Project**: ramprasadre56-5875
- **Reasoning Model**: azure/gpt-4o
- **API Key**: to be generated and stored in ACA secret
