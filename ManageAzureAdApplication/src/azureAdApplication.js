"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });

var tl = require('azure-pipelines-task-lib');
var os = require('os');
var path = require('path');
var fs = require('fs');
var uuidV4 = require('uuid/v4');

const msRestNodeAuth = require('@azure/ms-rest-nodeauth');
const azureGraph = require('@azure/graph');

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

    if(taskReplyUrls.length === 0){
        //replyUrls = '[]'
    }
    
    // Create manifest.json file from requiredResource content
    var tempDirectory = tl.getVariable('agent.tempDirectory');
    tl.checkPath(tempDirectory, `${tempDirectory} (agent.tempDirectory)`);
    var filePath = path.join(tempDirectory, uuidV4() + '.json');
    fs.writeFile(filePath, requiredResource, { encoding: 'utf8' });

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
            var azureApplicationId;
            var servicePrincipalId;

            /*
            var taskReplyUrls = [
                'http://' + applicationName + '.' + rootDomain,
                'http://' + applicationName + '.' + rootDomain + '/signin-oidc',
                'http://' + applicationName + '.' + rootDomain + '/signin-aad'
            ];
            */

            var now = new Date();
            const nextYear = new Date(now.getFullYear()+1, now.getMonth(), now.getDay());

            if(apps.length == 0){
                console.log("application not found");
                
                // Use UpdatePasswordCredentials
                var newPwdCreds = [{
                    endDate: nextYear,
                    value: applicationSecret,
                }]

                console.log(taskReplyUrls);
                var rp = JSON.parse(taskReplyUrls);
                console.log(rp);

                var newAppParms = {
                    displayName: applicationName,
                    homepage: homeUrl,
                    passwordCredentials: newPwdCreds,
                    replyUrls = rp
                };

                console.log("---------------------------------------------");
                console.log("");
                console.log("Creating new application name " + applicationName + " ...");
                graphClient.applications.create(newAppParms)
                .then(applicationCreateResult => {
                    console.log("");
                    console.log("---------------------------------------------");
                    console.log(applicationCreateResult);
                    console.log("---------------------------------------------");
                    console.log("");
                    console.log("Create Application Service principal ...");

                    var serviceParms = {
                        displayName: applicationName,
                        appId: applicationCreateResult.appId,
                    };

                    var ownerParm = {
                        url: 'https://graph.windows.net/' + tenantId + '/directoryObjects/' + ownerId
                    };
                    graphClient.applications.addOwner(applicationCreateResult.objectId, ownerParm)
                    .catch(err=> {
                        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                    });

                    graphClient.servicePrincipals.create(serviceParms)
                    .then(serviceCreateResult => {
                        console.log("");
                        console.log("---------------------------------------------");
                        console.log("Service Principal creation result:");
                        console.log(serviceCreateResult);
                        console.log("");
                        console.log("---------------------------------------------");

                        var appUpdateParm = {
                            identifierUris: [ 'https://' + rootDomain + '/' + applicationCreateResult.appId ]
                        };
                        graphClient.applications.patch(applicationCreateResult.objectId, appUpdateParm)
                        .then(appUpdateResult => {
                            console.log("");
                            console.log("AppUpdateResult: ");
                            console.log("");
                            console.log("---------------------------------------------");
                            console.log(appUpdateResult);
                        }).catch(err => {
                            tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                        });
                        

                    }).catch(err => {
                        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                    });
                }).catch(err => {
                    tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
                })
            } else {
                console.log("application found");
                azureApplicationId = appObject.appId;
            }

        }).catch(err=> {
            tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
        });
    }).catch(err=> {
        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
    });
} catch (err) {
    tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
}