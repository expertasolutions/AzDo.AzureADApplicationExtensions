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
const msRestNodeAuth = require('@azure/ms-rest-nodeauth');
const azureGraph = require('@azure/graph');

try {
    
    var azureEndpointSubscription = tl.getInput("azureSubscriptionEndpoint", true);
    var applicationName = tl.getInput("applicationName", true);
    
    var subcriptionId = tl.getEndpointDataParameter(azureEndpointSubscription, "subscriptionId", false);

    var servicePrincipalId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalid", false);
    var servicePrincipalKey = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalkey", false);
    var tenantId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription,"tenantid", false);

    console.log("SubscriptionId: " + subcriptionId);
    console.log("ServicePrincipalId: " + servicePrincipalId);
    console.log("ServicePrincipalKey: " + servicePrincipalKey);
    console.log("TenantId: " + tenantId);

    console.log("Application Name: " + applicationName);

    msRestNodeAuth.loginWithServicePrincipalSecret(
        servicePrincipalId, servicePrincipalKey, tenantId
    ).then(creds => {

        var pipeCreds = new msRestNodeAuth.ApplicationTokenCredentials(creds.clientId, tenantId, creds.secret, 'graph');
        var graphClient = new azureGraph.GraphRbacManagementClient(pipeCreds, tenantId, { baseUri: 'https://graph.windows.net' });
        
        var appFilterValue = "displayName eq '" + applicationName + "'"
        var appFilter = {
            filter: appFilterValue 
        };

        graphClient.servicePrincipals.list(appFilter)
        .then(appResults => {
            var appEntity = appResults[0];
            console.log(appEntity);            
            //console.log("Set the Azure AD Application id...");
            //tl.setVariable("azureAdApplicationId", appEntity.appId);
            console.log("Set the Azure Permission access ...");
            var permissionString = JSON.stringify(appEntity.oauth2Permissions);
            console.log(permissionString);
            //tl.setVariable("azureAdApplicationResourceAccessJson", permissionString);
        })
        .catch(err => {
            tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
        });
    }).catch(err => {
        tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
    });
} catch (err) {
    tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
}