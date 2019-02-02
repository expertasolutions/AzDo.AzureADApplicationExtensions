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
var os = require('os');
var path = require('path');
var fs = require('fs');
var uuidV4 = require('uuid/v4');

try {
    
    var azureEndpointSubscription = tl.getInput("azureSubscriptionEndpoint", true);
    var applicationName = tl.getInput("applicationName", true);
    var ownerId = tl.getInput("applicationOwnerId", true);
    var rootDomain = tl.getInput("rootDomain", true);
    var applicationSecret = tl.getInput("applicationSecretPassword", true);
    var adminuser = tl.getInput("azadadminuser", true);
    var adminpwd = tl.getInput("azadadminpwd", true);
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
    console.log("ApplicationSecret: " + applicationSecret);
    console.log("Home url: " + homeUrl);
    console.log("Reply Urls: " + replyUrls);
    console.log("OwnerId: " + ownerId);
    
    // Create manifest.json file from requiredResource content
    var tempDirectory = tl.getVariable('agent.tempDirectory');
    tl.checkPath(tempDirectory, `${tempDirectory} (agent.tempDirectory)`);
    var filePath = path.join(tempDirectory, uuidV4() + '.json');
    fs.writeFile(filePath, requiredResource, { encoding: 'utf8' });
    
    var pwsh = new shell({
        executionPolicy: 'Bypass',
        noProfile: true
    });
    
    pwsh.addCommand(__dirname  + "/createAzureApp.ps1 -subscriptionId '" + subcriptionId + "'"
        + " -servicePrincipalId '" + servicePrincipalId + "' -servicePrincipalKey '" + servicePrincipalKey + "' -tenantId '" + tenantId + "'"
        + " -applicationName '" + applicationName + "'"
        + " -rootDomain '" + rootDomain + "' -applicationSecret '" + applicationSecret + "'"
        + " -manifestFile '" + filePath + "' -homeUrl '" + homeUrl + "' -replyUrls '" + replyUrls + "' -ownerId '" + ownerId + "'")
        .then(function() {
            return pwsh.invoke();
        }).then(function(output){
            console.log(output);
            var regx = "(Azure ApplicationID): ([A-Za-z0-9\\-]*)";
            var result = output.match(regx);
            var appId = result[2];
            tl.setVariable("azureAdApplicationId", appId);
            pwsh.dispose();
        }).catch(function(err){
            console.log(err);
            var regx = "(Azure ApplicationID): ([A-Za-z0-9\\-]*)";
            var result = err.match(regx);
            if(result != null) {
                var appId = result[2];
                tl.setVariable("azureAdApplicationId", appId);
            } else {
                console.log("not application id returned");
                tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
            }
            pwsh.dispose();
        });
} catch (err) {
    tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
}