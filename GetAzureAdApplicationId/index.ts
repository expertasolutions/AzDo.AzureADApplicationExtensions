import tl = require('azure-pipelines-task-lib/task');
import msRestNodeAuth = require('@azure/ms-rest-nodeauth');
import azureGraph = require('@azure/graph');

async function LoginToAzure(servicePrincipalId:string, servicePrincipalKey:string, tenantId:string) {
  return await msRestNodeAuth.loginWithServicePrincipalSecret(servicePrincipalId, servicePrincipalKey, tenantId );
};

async function FindAzureServicePrincipal(applicationName:string, graphClient:azureGraph.GraphRbacManagementClient){
  var appFilterValue = "displayName eq '" + applicationName + "'";
  var appFilter = {
      filter: appFilterValue 
  };
  var searchResults = await graphClient.servicePrincipals.list(appFilter);
  if(searchResults.length === 0){
      return null;
  } else {
      return searchResults[0];
  }
}

async function run() {
  try {

    let azureEndpointSubscription = tl.getInput("azureSubscriptionEndpoint", true) as string;
    let applicationName = tl.getInput("applicationName", true) as string;
    
    let subcriptionId = tl.getEndpointDataParameter(azureEndpointSubscription, "subscriptionId", false) as string;

    let servicePrincipalId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalid", false) as string;
    let servicePrincipalKey = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalkey", false) as string;
    let tenantId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription,"tenantid", false) as string;

    console.log("SubscriptionId: " + subcriptionId);
    console.log("ServicePrincipalId: " + servicePrincipalId);
    console.log("ServicePrincipalKey: " + servicePrincipalKey);
    console.log("TenantId: " + tenantId);
    console.log("Application Name: " + applicationName);

    const azureCredentials = await LoginToAzure(servicePrincipalId, servicePrincipalKey, tenantId);
    var pipeCreds:any = new msRestNodeAuth.ApplicationTokenCredentials(azureCredentials.clientId, tenantId, azureCredentials.secret, 'graph');
    var graphClient = new azureGraph.GraphRbacManagementClient(pipeCreds, tenantId, { baseUri: 'https://graph.windows.net' });

    var servicePrincipalInstance = await FindAzureServicePrincipal(applicationName, graphClient);
    if(servicePrincipalInstance === null){
      console.log("Azure AD Application with name '" + applicationName + "' isn't found");
      tl.setResult(tl.TaskResult.Failed, "Application Name not found in Azure AD directory");
    } else {
      console.log("Set the Azure AD application id ... ");
      tl.setVariable("azureAdApplicationId", servicePrincipalInstance.appId as string);

      console.log("Set the Azure Permission Access ...");
      let permissionString = JSON.stringify(servicePrincipalInstance.oauth2Permissions);
      tl.setVariable("azureAdApplicationResourceAccessJson", permissionString);
    }

  } catch(err) {
    tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
  }
}

run();