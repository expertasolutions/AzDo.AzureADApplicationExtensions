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

var tl = require('vsts-task-lib');
var shell = require('node-powershell');

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
   
    var pwsh = new shell({
        executionPolicy: 'Bypass',
        noProfile: true
    });
    
    pwsh.addCommand(__dirname  + "/createAzureApp.ps1 -subscriptionId '" + subcriptionId + "'"
        + " -servicePrincipalId '" + servicePrincipalId + "' -servicePrincipalKey '" + servicePrincipalKey + "' -tenantId '" + tenantId + "'"
        + " -applicationName '" + applicationName + "'")
        .then(function(){
            return pwsh.invoke();
        })
        .then(function(output){
            console.log(output);
            
            console.log("Getting the Azure ApplicationId value...");
            var regx = "(Azure ApplicationID): ([A-Za-z0-9\-]*)";
            var result = output.match(regx);
            var appId = result[2];
            
            console.log("Setting the AzureAdApplicationId ...");
            tl.setVariable("azureAdApplicationId", appId);
            
            console.log("Getting the Azure Permission access ...");
            var permissionJsonRegx = "(Azure Permission Access Info-json):([\[\w\{\": -.]*\}\])";
            //result = output.match(permissionJsonRegx);
            //var permissionJson = result[2];
            //console.log("Setting azureAdApplicationResourceAccessJson ...");
            //tl.setVariable("azureAdApplicationResourceAccessJson", permissionJson);
            
            pwsh.dispose();
        }).catch(function(err){
            console.log(err);
            tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
            pwsh.dispose();
        });
} catch (err) {
    console.log(err);
    tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
}