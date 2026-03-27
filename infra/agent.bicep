// ===============================================
// SharePoint Foundry Agent — ACA Infrastructure
// Provisions: Container Apps Environment, ACA App,
//             System-assigned Managed Identity,
//             ACR pull role assignment
// ===============================================

@description('Resource name prefix (match main.bicep)')
param resourcePrefix string = 'iqseries'

@description('Azure region')
param location string = 'eastus2'

@description('Container Registry login server (e.g. myregistry.azurecr.io)')
param acrLoginServer string

@description('Docker image tag to deploy')
param imageTag string = 'latest'

@description('Azure AI endpoint from main.bicep outputs')
param azureAiEndpoint string

@description('Foundry project name from main.bicep')
param foundryProjectName string

@description('AI Search endpoint from main.bicep outputs')
param searchEndpoint string

@description('AI Search index name (knowledge base)')
param searchIndexName string

@description('Reasoning model string (e.g. azure/gpt-4o)')
param reasoningModel string = 'azure/gpt-4o'

var uniqueSuffix = uniqueString(resourceGroup().id)
var containerAppName = '${resourcePrefix}-agent-${uniqueSuffix}'
var acrPullRoleId = '7f951dda-4ed3-4680-a7ca-43fe172d538d'

// -----------------------------------------------
// Container Apps Environment
// -----------------------------------------------

resource acaEnvironment 'Microsoft.App/managedEnvironments@2024-03-01' = {
  name: '${resourcePrefix}-aca-env-${uniqueSuffix}'
  location: location
  properties: {
    zoneRedundant: false
  }
}

// -----------------------------------------------
// Container App
// -----------------------------------------------

resource containerApp 'Microsoft.App/containerApps@2024-03-01' = {
  name: containerAppName
  location: location
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    managedEnvironmentId: acaEnvironment.id
    configuration: {
      ingress: {
        external: true
        targetPort: 3000
        transport: 'auto'
      }
      registries: [
        {
          server: acrLoginServer
          identity: 'system'
        }
      ]
    }
    template: {
      containers: [
        {
          name: 'agent-service'
          image: '${acrLoginServer}/sharepoint-foundry-agent:${imageTag}'
          resources: {
            cpu: json('0.5')
            memory: '1Gi'
          }
          env: [
            { name: 'AZURE_AI_ENDPOINT', value: azureAiEndpoint }
            { name: 'FOUNDRY_PROJECT_NAME', value: foundryProjectName }
            { name: 'AZURE_SEARCH_ENDPOINT', value: searchEndpoint }
            { name: 'AZURE_SEARCH_INDEX_NAME', value: searchIndexName }
            { name: 'REASONING_MODEL', value: reasoningModel }
            { name: 'PORT', value: '3000' }
          ]
        }
      ]
      scale: {
        minReplicas: 1
        maxReplicas: 3
      }
    }
  }
}

// -----------------------------------------------
// Grant ACA identity ACR pull permission
// -----------------------------------------------

resource acrResource 'Microsoft.ContainerRegistry/registries@2023-07-01' existing = {
  name: split(acrLoginServer, '.')[0]
}

resource acrPullRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(resourceGroup().id, containerApp.name, acrPullRoleId)
  scope: acrResource
  properties: {
    principalId: containerApp.identity.principalId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', acrPullRoleId)
  }
}

// -----------------------------------------------
// Outputs
// -----------------------------------------------

output acaEndpoint string = 'https://${containerApp.properties.configuration.ingress.fqdn}'
output acaManagedIdentityPrincipalId string = containerApp.identity.principalId
