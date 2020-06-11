import tl = require('azure-pipelines-task-lib/task');
import msRestNodeAuth = require('@azure/ms-rest-nodeauth');
import azureGraph = require('@azure/graph');

function delay(ms: number) {
    return new Promise( resolve => setTimeout(resolve, ms) );
}

async function LoginToAzure(servicePrincipalId:string, servicePrincipalKey:string, tenantId:string) {
    return await msRestNodeAuth.loginWithServicePrincipalSecret(servicePrincipalId, servicePrincipalKey, tenantId );
};

async function FindAzureAdApplication(applicationName:string, graphClient:azureGraph.GraphRbacManagementClient){
    var appFilterValue = "displayName eq '" + applicationName + "'";
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

async function FindServicePrincipal(
    applicationId:string
  , graphClient:azureGraph.GraphRbacManagementClient
) {
  console.log("List Service Principal ...");
  let result = await graphClient.servicePrincipals.list();
  return result.find(x=> x.appId === applicationId);
}

async function CreateServicePrincipal(
      applicationName:string
    , applicationId:string
    , graphClient:azureGraph.GraphRbacManagementClient
) {
    console.log("Create Service Principal ...");
    var serviceParms = {
        displayName: applicationName,
        appId: applicationId
    };
    let result = await graphClient.servicePrincipals.create(serviceParms);
    
    // Delay for the Azure AD Application and Service Principal...
    await delay(60000);
    return result;
}

async function AddADApplicationOwner(
        applicationObjectId:string
    ,   ownerId:string
    ,   tenantId:string
    ,   graphClient:azureGraph.GraphRbacManagementClient) 
{
    let urlGraph = 'https://graph.windows.net/' + tenantId + '/directoryObjects/' + ownerId;
    var ownerParm = {
        url: urlGraph
    };
    console.log("   Adding owner to Azure ActiveDirectory Application ...");
    return await graphClient.applications.addOwner(applicationObjectId, ownerParm);
}

async function CreateOrUpdateADApplication(
          appObjectId:string | null
        , applicationName:string
        , rootDomain:string
        , applicationSecret:string
        , homeUrl:string
        , taskReplyUrls:string
        , requiredResource:string
        , graphClient:azureGraph.GraphRbacManagementClient
) {
    if(appObjectId === null)
        console.log("Creating new Azure ActiveDirectory AD Application...");
    else
        console.log("Updating Azure ActiveDirectory AD Application...");

    var now = new Date();
    const nextYear = new Date(now.getFullYear()+1, now.getMonth(), now.getDay());
    var newPwdCreds = [{
        endDate: nextYear,
        value: applicationSecret
    }];

    var taskUrlArray:Array<string>;
    if(taskReplyUrls === undefined || taskReplyUrls.length === 0){
        taskUrlArray = [
            'http://' + applicationName + '.' + rootDomain,
            'http://' + applicationName + '.' + rootDomain + '/signin-oidc',
            'http://' + applicationName + '.' + rootDomain + '/signin-aad'
        ];
    } else {
        taskUrlArray = JSON.parse(taskReplyUrls);
    }

    var newAppParms = {
        displayName: applicationName,
        homepage: homeUrl,
        passwordCredentials: newPwdCreds,
        replyUrls: taskUrlArray,
        requiredResourceAccess: JSON.parse(requiredResource)
    };

    if(appObjectId == null){
        await graphClient.applications.create(newAppParms);

        // Delay for the Azure AD Application and Service Principal...
        await delay(10000);
        return await FindAzureAdApplication(applicationName, graphClient);
    }
    else {
        await graphClient.applications.patch(appObjectId, newAppParms);
        // Delay for the Azure AD Application and Service Principal...
        await delay(10000);
        return await FindAzureAdApplication(applicationName, graphClient);
    }
}

async function grantAuth2Permissions (
        rqAccess: any
    ,   servicePrincipalId:string
    ,   graphClient:azureGraph.GraphRbacManagementClient
) {
    console.log("Grant Auth2Permissions ...");
    var resourceAppFilter = {
        filter: "appId eq '" + rqAccess.resourceAppId + "'"
    };
    
    var rs = await graphClient.servicePrincipals.list(resourceAppFilter);
    var srv = rs[0];
    var desiredScope = "";
    for(var i=0;i<rqAccess.resourceAccess.length;i++){
        var rAccess = rqAccess.resourceAccess[i];
        if(srv.oauth2Permissions != null) {
            var p = srv.oauth2Permissions.find(p=> {
                return p.id === rAccess.id;
            }) as azureGraph.GraphRbacManagementModels.OAuth2Permission;
            desiredScope += p.value + " ";
        }
    }

    var now = new Date();
    const nextYear = new Date(now.getFullYear()+1, now.getMonth(), now.getDay());
    
    var permissions = {
        body: {
            clientId: servicePrincipalId,
            consentType: 'AllPrincipals',
            scope: desiredScope,
            resourceId: srv.objectId,
            expiryTime: nextYear.toISOString()
        }
    } as azureGraph.GraphRbacManagementModels.OAuth2PermissionGrantCreateOptionalParams;
    return await graphClient.oAuth2PermissionGrant.create(permissions)
}

async function run() {
    try {
        
        var azureEndpointSubscription = tl.getInput("azureSubscriptionEndpoint", true) as string;
        var applicationName = tl.getInput("applicationName", true) as string;
        var ownerId = tl.getInput("applicationOwnerId", true) as string;
        var rootDomain = tl.getInput("rootDomain", true) as string;
        var applicationSecret = tl.getInput("applicationSecretPassword", true) as string;
        var requiredResource = tl.getInput("requiredResource", true) as string;
        var homeUrl = tl.getInput("homeUrl", true) as string;
        var taskReplyUrls = tl.getInput("replyUrls", false) as string;
        
        var subcriptionId = tl.getEndpointDataParameter(azureEndpointSubscription, "subscriptionId", false) as string;

        var servicePrincipalId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalid", false) as string;
        var servicePrincipalKey = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalkey", false) as string;
        var tenantId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription,"tenantid", false) as string;

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

        const azureCredentials = await LoginToAzure(servicePrincipalId, servicePrincipalKey, tenantId);

        var pipeCreds:any = new msRestNodeAuth.ApplicationTokenCredentials(azureCredentials.clientId, tenantId, azureCredentials.secret, 'graph');
        var graphClient = new azureGraph.GraphRbacManagementClient(pipeCreds, tenantId, { baseUri: 'https://graph.windows.net' });

        let applicationInstance = await FindAzureAdApplication(applicationName, graphClient);
        if(applicationInstance === null) {
            // Create new Azure AD Application
            applicationInstance = await CreateOrUpdateADApplication(null, applicationName, rootDomain, applicationSecret, homeUrl, taskReplyUrls, requiredResource, graphClient);

            // Add Owner to new Azure AD Application
            await AddADApplicationOwner(applicationInstance.objectId as string, ownerId, tenantId, graphClient);

            // Create Service Principal for Azure AD Application
            let newServicePrincipal = await CreateServicePrincipal(applicationName, applicationInstance.appId as string, graphClient);

            // Set Application Permission
            for(var i=0;i<applicationInstance.requiredResourceAccess.length;i++){
                var rqAccess = applicationInstance.requiredResourceAccess[i];
                await grantAuth2Permissions(rqAccess, newServicePrincipal.objectId as string, graphClient);
            }

            // Update Application IdentifierUrisApplicationInstance
            var appUpdateParms = {
                identifierUris: ['https://' + rootDomain + '/' + applicationInstance.appId ]
            };
            await graphClient.applications.patch(applicationInstance.objectId, appUpdateParms);
        } 
        else {
            applicationInstance = await CreateOrUpdateADApplication(applicationInstance.objectId as string, applicationName, rootDomain, applicationSecret, homeUrl, taskReplyUrls, requiredResource, graphClient);
            let service = await FindServicePrincipal(applicationInstance.appId, graphClient);

            // Add Owner to new Azure AD Application
            await AddADApplicationOwner(applicationInstance.objectId as string, ownerId, tenantId, graphClient);

            // Set Application Permission
            for(var i=0;i<applicationInstance.requiredResourceAccess.length;i++){
                var rqAccess = applicationInstance.requiredResourceAccess[i];
                await grantAuth2Permissions(rqAccess, service.objectId as string, graphClient);
            }

            // Update Application IdentifierUrisApplicationInstance
            var appUpdateParms = {
                identifierUris: ['https://' + rootDomain + '/' + applicationInstance.appId ]
            };
            await graphClient.applications.patch(applicationInstance.objectId, appUpdateParms);
        }

        tl.setVariable("azureAdApplicationId", applicationInstance.appId as string);

        // Delay for the Azure AD Application and Service Principal...
        await delay(10000);
    } catch (err) {
        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
    }
}

run();