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
    
    var azureEndpointSubscriptionId = tl.getInput("azureSubscriptionEndpoint", true);
    var applicationName = tl.getInput("applicationName", true);
    var adminuser = tl.getInput("azadadminuser", true);
    var adminpwd = tl.getInput("azadadminpwd", true);
    
    var subcriptionId = tl.getEndpointDataParameter(azureEndpointSubscriptionId, "subscriptionId", false);

    console.log("SubscriptionId: " + subcriptionId);
    console.log("AdminAdUser: " + adminuser);
    console.log("AdminAdPwd: " + adminpwd);
    console.log("Application Name: " + applicationName);
   
    var pwsh = new shell({
        executionPolicy: 'Bypass',
        noProfile: true
    });
    
    pwsh.addCommand(__dirname  + "/createAzureApp.ps1 -subscriptionId '" + subcriptionId + "' -applicationName '" + applicationName 
        + "' -adUser '" + adminuser + "' -adPwd '" + adminpwd + "'");
    
    pwsh.invoke()
        .then(function(output) {
            console.log(output);
            
            var regx = "(Azure ApplicationID): ([A-Za-z0-9\\-]*)";
            var result = output.match(regx);
            var appId = result[2];
            
            tl.setVariable("azureAdApplicationId", appId);
            
            var permissionJsonRegx = "(Azure Permission Access Info-json):([\\[\\w\\{\": -.]*\\}\\])";
            result = output.match(permissionJsonRegx);
            var permissionJson = result[2];
            tl.setVariable("azureAdApplicationResourceAccessJson", permissionJson);
            
            pwsh.dispose();
        }).catch(function(err){
            tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
            pwsh.dispose();
        });
    
} catch (err) {
    tl.setResult(tl.TaskResult.Failed, err.message || 'run() failed');
}