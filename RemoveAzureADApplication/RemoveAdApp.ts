import tl = require('azure-pipelines-task-lib/task');
import msRestNodeAuth = require('@azure/ms-rest-nodeauth');
import azureGraph = require('@azure/graph');

async function LoginToAzure(servicePrincipalId:string, servicePrincipalKey:string, tenantId:string) {
    return await msRestNodeAuth.loginWithServicePrincipalSecret(servicePrincipalId, servicePrincipalKey, tenantId );
};

async function FindAzureAdApplication(applicationId:string, graphClient:azureGraph.GraphRbacManagementClient){
    var appFilterValue = "appId eq '" + applicationId + "'";
    var appFilter = {
        filter: appFilterValue 
    };
    var searchResults = await graphClient.applications.list(appFilter);
    if(searchResults.length === 0){
        return null;
    } else {
        return searchResults[0];
    }
}

async function DeleteAzureADApplication(applicationObjectId:string, graphClient:azureGraph.GraphRbacManagementClient){
  return await graphClient.applications.deleteMethod(applicationObjectId);
}

async function run() {
    try {
        
        var azureEndpointSubscription = tl.getInput("azureSubscriptionEndpoint", true) as string;
        var applicationId = tl.getInput("applicationId", true) as string;
        
        var subcriptionId = tl.getEndpointDataParameter(azureEndpointSubscription, "subscriptionId", false) as string;
    
        var servicePrincipalId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalid", false) as string;
        var servicePrincipalKey = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalkey", false) as string;
        var tenantId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription,"tenantid", false) as string;
    
        console.log("SubscriptionId: " + subcriptionId);
        console.log("ServicePrincipalId: " + servicePrincipalId);
        console.log("ServicePrincipalKey: " + servicePrincipalKey);
        console.log("TenantId: " + tenantId);
    
        console.log("Application Id: " + applicationId);

        const azureCredentials = await LoginToAzure(servicePrincipalId, servicePrincipalKey, tenantId);

        var pipeCreds:any = new msRestNodeAuth.ApplicationTokenCredentials(azureCredentials.clientId, tenantId, azureCredentials.secret, 'graph');
        var graphClient = new azureGraph.GraphRbacManagementClient(pipeCreds, tenantId, { baseUri: 'https://graph.windows.net' });

        var applicationInstance = await FindAzureAdApplication(applicationId, graphClient);
        if(applicationInstance == null){
            console.log("Azure AD Application with id '" + applicationId + "' does not exists");
        } else {
          console.log("Azure AD Application with id '" + applicationId + "' is found");
          console.log("   Removing Azure AD Application with id '" + applicationId + "' ...");
          await DeleteAzureADApplication(applicationInstance.objectId, graphClient);
        }
} catch (err) {
        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
    }
}

run();