import tl = require('azure-pipelines-task-lib/task');
import msRestNodeAuth = require('@azure/ms-rest-nodeauth');
import azureGraph = require('@azure/graph');
import { RequiredResourceAccess, ServicePrincipal, OAuth2PermissionGrantListOptionalParams } from '@azure/graph/src/models';
import { ServicePrincipalObjectResult } from '@azure/graph/esm/models/mappers';
import { async, race } from 'q';
import { RequestPrepareOptions } from '@azure/ms-rest-js';

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

async function FindServicePrincipal (
    applicationId:string
  , graphClient:azureGraph.GraphRbacManagementClient
) {
  console.log("List Service Principal ...");
  
  var resourceAppFilter = {
      filter: "appId eq '" + applicationId + "'"
  };

  let searchResults = await graphClient.servicePrincipals.list(resourceAppFilter);
  if(searchResults.length === 0){
    return undefined;
  } else {
    let srvPrincipal = searchResults[0];
    srvPrincipal = await graphClient.servicePrincipals.get(srvPrincipal.objectId);
    return srvPrincipal;
  }
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
    console.log("Add Application Owner ... ");
    let urlGraph = 'https://graph.windows.net/' + tenantId + '/directoryObjects/' + ownerId;
    var ownerParm = {
        url: urlGraph
    };

    let owners = await graphClient.applications.listOwners(applicationObjectId);
    if(owners.find(x=> x.objectId === ownerId)) {
        console.log("   Owner already existing on the Azure ActiveDirectory Application");
        return undefined;
    } else {
        console.log("   Adding owner to Azure ActiveDirectory Application ...");
        return await graphClient.applications.addOwner(applicationObjectId, ownerParm);
    }
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

async function deleteAuth2Permissions (
    objectId: string
,   graphClient:azureGraph.GraphRbacManagementClient
) {
    console.log("Delete Auth2Permissions ... of : " + objectId);
    let options: OAuth2PermissionGrantListOptionalParams = {
        filters: objectId
    };

    let result = graphClient.oAuth2PermissionGrant.list(options);

    console.log("------");
    console.log(JSON.stringify(result));
    console.log("------");
    return await graphClient.oAuth2PermissionGrant.deleteMethod(objectId)
}

async function grantAuth2Permissions (
        rqAccess: any
    ,   servicePrincipalId:string
    ,   graphClient:azureGraph.GraphRbacManagementClient
) {
    console.log("Grant Auth2Permissions '" + rqAccess.resourceAppId + "' ...");
    var resourceAppFilter = {
        filter: "appId eq '" + rqAccess.resourceAppId + "'"
    };
    
    let rs = await graphClient.servicePrincipals.list(resourceAppFilter);
    let srv = rs[0];

    let desiredScope = "";
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

    try {
        await graphClient.oAuth2PermissionGrant.create(permissions);
        console.log("   Permissions granted for '" + rqAccess.resourceAppId + "'");
    } catch {
        console.log("   Permissions already granted for '" + rqAccess.resourceAppId + "'");
    }
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

            // Create Service Principal for Azure AD Application
            let newServicePrincipal = await CreateServicePrincipal(applicationName, applicationInstance.appId as string, graphClient);

            // Set Application Permissions
            for(var i=0;i<applicationInstance.requiredResourceAccess.length;i++){
                var rqAccess = applicationInstance.requiredResourceAccess[i];
                await grantAuth2Permissions(rqAccess, newServicePrincipal.objectId as string, graphClient);
            }
        }
        else {
            applicationInstance = await CreateOrUpdateADApplication(applicationInstance.objectId as string, applicationName, rootDomain, applicationSecret, homeUrl, taskReplyUrls, requiredResource, graphClient);
            let service = await FindServicePrincipal(applicationInstance.appId, graphClient);
            let newPermissions: RequiredResourceAccess[] = JSON.parse(requiredResource);

            
            //https://graph.microsoft.com/v1.0/oauth2PermissionGrants/{id}
            
            var currentGrants = (await graphClient.oAuth2PermissionGrant.list()).filter(x=> x.clientId === service.objectId);
            console.log("-----");
            console.log(JSON.stringify(currentGrants));
            console.log("-----")
            for(let i=0;i<currentGrants.length;i++) {
                let prm = currentGrants[i];

                let requestUrl: RequestPrepareOptions = {
                    url: "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/" + prm.objectId,
                    method: "DELETE"
                };

                console.log("Delete: " + prm.objectId);
                let requestResult = await graphClient.sendRequest(requestUrl);
                console.log(JSON.stringify(requestResult));
                await graphClient.oAuth2PermissionGrant.deleteMethod(prm.objectId);
            }

            // Set Application Permissions
            for(var i=0;i<newPermissions.length;i++){
                var newPerm = newPermissions[i];
                //await grantAuth2Permissions(newPerm, service.objectId as string, graphClient);
            }
        }

        // Add Owner to new Azure AD Application
        await AddADApplicationOwner(applicationInstance.objectId as string, ownerId, tenantId, graphClient);

        // Update Application IdentifierUrisApplicationInstance
        var appUpdateParms = {
            identifierUris: ['https://' + rootDomain + '/' + applicationInstance.appId ]
        };
        await graphClient.applications.patch(applicationInstance.objectId, appUpdateParms);

        tl.setVariable("azureAdApplicationId", applicationInstance.appId as string);

        // Delay for the Azure AD Application and Service Principal...
        await delay(10000);
    } catch (err) {
        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
    }
}

run();