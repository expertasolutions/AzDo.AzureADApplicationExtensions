const tl = require('azure-pipelines-task-lib/task');
const msRestNodeAuth = require('@azure/ms-rest-nodeauth');
const azureGraph = require('@azure/graph');

async function LoginToAzure(servicePrincipalId:string, servicePrincipalKey:string, tenantId:string) {
    return await msRestNodeAuth.loginWithServicePrincipalSecret(servicePrincipalId, servicePrincipalKey, tenantId );
};

async function FindAzureAdApplication(applicationName:string, graphClient:any){
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

async function FindServicePrincipalByAppId(appId, graphClient){
    return null;
}

async function run() {
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
        var tenantId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription,"tenantid", false);

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
        console.log("Azure Credentials");
        console.log(azureCredentials);

        var pipeCreds = new msRestNodeAuth.ApplicationTokenCredentials(azureCredentials.clientId, tenantId, azureCredentials.secret, 'graph');
        var graphClient = new azureGraph.GraphRbacManagementClient(pipeCreds, tenantId, { baseUri: 'https://graph.windows.net' });

        var applicationInstance = FindAzureAdApplication(applicationName, graphClient);

        if(applicationInstance == null){
            console.log("Application not found");
        } else {
            console.log("Application found");
            console.log(applicationInstance);
        }

        /*
        msRestNodeAuth.loginWithServicePrincipalSecret(
            servicePrincipalId, servicePrincipalKey, tenantId
        ).then(creds => {
            var pipeCreds = new msRestNodeAuth.ApplicationTokenCredentials(creds.clientId, tenantId, creds.secret, 'graph');
            var graphClient = new azureGraph.GraphRbacManagementClient(pipeCreds, tenantId, { baseUri: 'https://graph.windows.net' });
            
            var appFilterValue = "displayName eq '" + applicationName + "'"
            var appFilter = {
                filter: appFilterValue 
            };

            graphClient.applications.list(appFilter)
            .then(apps => {
                var appObject = apps[0];

                var now = new Date();
                const nextYear = new Date(now.getFullYear()+1, now.getMonth(), now.getDay());

                // Use UpdatePasswordCredentials
                var newPwdCreds = [{
                    endDate: nextYear,
                    value: applicationSecret,
                }];

                if(apps.length == 0){
                    console.log("Creating new Azure Active Directory application...");
                    var taskUrlArray;
                    if(taskReplyUrls.length === 0){
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
                    graphClient.applications.create(newAppParms)
                    .then(applicationCreateResult => {
                        var serviceParms = {
                            displayName: applicationName,
                            appId: applicationCreateResult.appId,
                        };

                        var ownerParm = {
                            url: 'https://graph.windows.net/' + tenantId + '/directoryObjects/' + ownerId
                        };

                        console.log("Adding owner to Azure Active Directory Application ...");
                        graphClient.applications.addOwner(applicationCreateResult.objectId, ownerParm)
                        .catch(err=> {
                            tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                        });

                        console.log("Creating Application Service Principal ...");
                        graphClient.servicePrincipals.create(serviceParms)
                        .then(serviceCreateResult => {

                            var applicationServicePrincipalObjectId = serviceCreateResult.objectId;
                            for(var i=0;i<applicationCreateResult.requiredResourceAccess.length;i++) {
                                var rqAccess = applicationCreateResult.requiredResourceAccess[i];
                                var resourceAppFilter = {
                                    filter: "appId eq '" + rqAccess.resourceAppId + "'"
                                };
                                graphClient.servicePrincipals.list(resourceAppFilter)
                                .then(rs => {
                                    var srv = rs[0];
                                    var desiredScope = "";
                                    for(var j=0;j<rqAccess.resourceAccess.length;j++){
                                        var rAccess = rqAccess.resourceAccess[j];
                                        var permission = srv.oauth2Permissions.find(p=> {
                                            return p.id === rAccess.id;
                                        });
                                        desiredScope += permission.value + " ";
                                    }
                                    var permission = {
                                        body: {
                                            clientId: applicationServicePrincipalObjectId,
                                            consentType: 'AllPrincipals',
                                            scope: desiredScope,
                                            resourceId: srv.objectId,
                                            expiryTime: nextYear.toISOString()
                                        }
                                    };

                                    graphClient.oAuth2PermissionGrant.create(permission)
                                    .then(p=> {
                                        console.log("Permissions granted");
                                    }).catch(err => {
                                        console.dir(err, {depth: null, colors: true});
                                        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                                    });
                                }).catch(err=> {
                                    tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                                });
                            }

                            var appUpdateParm = {
                                identifierUris: [ 'https://' + rootDomain + '/' + applicationCreateResult.appId ]
                            };
                            graphClient.applications.patch(applicationCreateResult.objectId, appUpdateParm)
                            .then(rs => {
                                tl.setVariable("azureAdApplicationId", applicationCreateResult.appId);
                            })
                            .catch(err => {
                                tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                            });
                        }).catch(err => {
                            tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                        });
                    }).catch(err => {
                        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                    })
                } else {

                    // Add expected owner
                    var ownerParm = {
                        url: 'https://graph.windows.net/' + tenantId + '/directoryObjects/' + ownerId
                    };
                    graphClient.applications.addOwner(appObject.objectId, ownerParm)
                    .catch(err=> {
                        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                    });

                    // Update Azure AD application
                    var updateAppParms = {
                        displayName: applicationName,
                        homepage: homeUrl,
                        passwordCredentials: newPwdCreds,
                        replyUrls: taskReplyUrls,
                        identifierUris: [ 'https://' + rootDomain + '/' + appObject.appId ],
                        requiredResourceAccess: JSON.parse(requiredResource)
                    };

                    console.log("Updating Azure Active Directory application ...");
                    graphClient.applications.patch(appObject.objectId, updateAppParms)
                    .catch(err=> {
                        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                    });
                }
            }).catch(err=> {
                tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
            });
        }).catch(err=> {
            tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
        });
    */
    } catch (err) {
        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
    }
}

run();