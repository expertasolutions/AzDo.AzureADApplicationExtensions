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
    var homeUrl = tl.getInput("homeUrl", false);
    var replyUrls = tl.getInput("replyUrls", false);
    
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
    console.log("Reply Urls: " + replyUrls);
    console.log("OwnerId: " + ownerId);
    console.log("");
    
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

            /*
            var replyUrls = [
                'http://' + applicationName + '.' + rootDomain,
                'http://' + applicationName + '.' + rootDomain + '/signin-oidc',
                'http://' + applicationName + '.' + rootDomain + '/signin-aad'
            ];
            */

            if(apps.length == 0){
                console.log("application not found");
                
                var newAppParms = {
                    displayName: applicationName,
                };

                console.log("Creating new application name " + applicationName + " ...");
                graphClient.applications.create(newAppParms)
                .then(createResult => {
                    console.log("createResult:");
                    console.log(createResult);
                }).catch(err => {
                    console.log(err);
                })
            } else {
                console.log("application found");
                azureApplicationId = appObject.appId;
            }

            //var identifierUrls = 'https://' + rootDomain + '/' + azureApplicationId;

        }).catch(err=> {
            tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
        });
    }).catch(err=> {
        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
    });
} catch (err) {
    tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
}