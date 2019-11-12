var tl = require('azure-pipelines-task-lib/task');
var msRestNodeAuth = require('@azure/ms-rest-nodeauth');
var azureGraph = require('@azure/graph');
async;
function LoginToAzure(servicePrincipalId, servicePrincipalKey, tenantId) {
    return await;
    msRestNodeAuth.loginWithServicePrincipalSecret(servicePrincipalId, servicePrincipalKey, tenantId);
}
async;
function FindAzureAdApplication(applicationName, graphClient) {
    var appFilterValue = "displayName eq '" + applicationName + "'";
    var appFilter = {
        filter: appFilterValue
    };
    var searchResults = await, graphClient, applications, list = (appFilter);
    if (searchResults.length === 0) {
        return null;
    }
    else {
        return searchResults[0];
    }
}
async;
function FindServicePrincipalByAppId(appId, graphClient) {
    return null;
}
async;
function run() {
    try {
        var azureEndpointSubscription = tl.getInput("azureSubscriptionEndpoint", true);
        var applicationName = tl.getInput("applicationName", true);
        var ownerId = tl.getInput("applicationOwnerId", true);
        var rootDomain = tl.getInput("rootDomain", true);
        var applicationSecret = tl.getInput("applicationSecretPassword", true);
        var requiredResource = tl.getInput("requiredResource", true);
        var homeUrl = tl.getInput("homeUrl", true);
        var taskReplyUrls = tl.getInput("replyUrls", false);
        var subcriptionId = tl.getEndpointDataParameter(azureEndpointSubscription, "subscriptionId", false);
        var servicePrincipalId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalid", false);
        var servicePrincipalKey = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalkey", false);
        var tenantId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "tenantid", false);
        console.log("SubscriptionId: " + subcriptionId);
        console.log("ServicePrincipalId: " + servicePrincipalId);
        console.log("ServicePrincipalKey: " + servicePrincipalKey);
        console.log("TenantId: " + tenantId);
        console.log("Application Name: " + applicationName);
        console.log("Root Domain: " + rootDomain);
        console.log("Home url: " + homeUrl);
        console.log("Reply Urls: " + taskReplyUrls);
        console.log("OwnerId: " + ownerId);
        console.log("");
        var azureCredentials = await, LoginToAzure_1 = (servicePrincipalId, servicePrincipalKey, tenantId);
        console.log("Azure Credentials");
        console.log(azureCredentials);
        var pipeCreds = new msRestNodeAuth.ApplicationTokenCredentials(azureCredentials.clientId, tenantId, azureCredentials.secret, 'graph');
        var graphClient = new azureGraph.GraphRbacManagementClient(pipeCreds, tenantId, { baseUri: 'https://graph.windows.net' });
        var applicationInstance = FindAzureAdApplication(applicationName, graphClient);
        if (applicationInstance == null) {
            console.log("Application not found");
        }
        else {
            console.log("Application found");
            console.log(applicationInstance);
        }
    }
    catch (err) {
        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
    }
}
run();
