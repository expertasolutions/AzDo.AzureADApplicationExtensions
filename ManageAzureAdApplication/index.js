"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
exports.__esModule = true;
var tl = require("azure-pipelines-task-lib/task");
var msRestNodeAuth = require("@azure/ms-rest-nodeauth");
var azureGraph = require("@azure/graph");
function LoginToAzure(servicePrincipalId, servicePrincipalKey, tenantId) {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, msRestNodeAuth.loginWithServicePrincipalSecret(servicePrincipalId, servicePrincipalKey, tenantId)];
                case 1: return [2 /*return*/, _a.sent()];
            }
        });
    });
}
;
function FindAzureAdApplication(applicationName, graphClient) {
    return __awaiter(this, void 0, void 0, function () {
        var appFilterValue, appFilter, searchResults;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    appFilterValue = "displayName eq '" + applicationName + "'";
                    appFilter = {
                        filter: appFilterValue
                    };
                    return [4 /*yield*/, graphClient.applications.list(appFilter)];
                case 1:
                    searchResults = _a.sent();
                    if (searchResults.length === 0) {
                        return [2 /*return*/, null];
                    }
                    else {
                        return [2 /*return*/, searchResults[0]];
                    }
                    return [2 /*return*/];
            }
        });
    });
}
function CreateServicePrincipal(applicationName, applicationId, graphClient) {
    return __awaiter(this, void 0, void 0, function () {
        var serviceParms;
        return __generator(this, function (_a) {
            serviceParms = {
                displayName: applicationName,
                appId: applicationId
            };
            return [2 /*return*/, graphClient.servicePrincipals.create(serviceParms)];
        });
    });
}
function AddADApplicationOwner(applicationObjectId, ownerId, tenantId, graphClient) {
    return __awaiter(this, void 0, void 0, function () {
        var ownerParm;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    ownerParm = {
                        url: 'https://graph.windows.net/' + tenantId + '/directoryObjects/' + ownerId
                    };
                    console.log("Adding owner to Azure ActiveDirectory Application ...");
                    return [4 /*yield*/, graphClient.applications.addOwner(applicationObjectId, ownerParm)];
                case 1: return [2 /*return*/, _a.sent()];
            }
        });
    });
}
function CreateOrUpdateADApplication(appObjectId, applicationName, rootDomain, applicationSecret, homeUrl, taskReplyUrls, requiredResource, graphClient) {
    return __awaiter(this, void 0, void 0, function () {
        var now, nextYear, newPwdCreds, taskUrlArray, newAppParms;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    console.log("Creating new Azure ActiveDirectory AD Application...");
                    now = new Date();
                    nextYear = new Date(now.getFullYear() + 1, now.getMonth(), now.getDay());
                    newPwdCreds = [{
                            endDate: nextYear,
                            value: applicationSecret
                        }];
                    if (taskReplyUrls.length === 0) {
                        taskUrlArray = [
                            'http://' + applicationName + '.' + rootDomain,
                            'http://' + applicationName + '.' + rootDomain + '/signin-oidc',
                            'http://' + applicationName + '.' + rootDomain + '/signin-aad'
                        ];
                    }
                    else {
                        taskUrlArray = JSON.parse(taskReplyUrls);
                    }
                    newAppParms = {
                        displayName: applicationName,
                        homepage: homeUrl,
                        passwordCredentials: newPwdCreds,
                        replyUrls: taskUrlArray,
                        requiredResourceAccess: JSON.parse(requiredResource)
                    };
                    if (!(appObjectId == null)) return [3 /*break*/, 2];
                    return [4 /*yield*/, graphClient.applications.create(newAppParms)];
                case 1: return [2 /*return*/, _a.sent()];
                case 2: return [4 /*yield*/, graphClient.applications.patch(appObjectId, newAppParms)];
                case 3: return [2 /*return*/, _a.sent()];
            }
        });
    });
}
function grantAuth2Permissions(rqAccess, servicePrincipalId, graphClient) {
    return __awaiter(this, void 0, void 0, function () {
        var resourceAppFilter, rs, srv, desiredScope, i, rAccess, p, now, nextYear, permissions;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    resourceAppFilter = {
                        filter: "appId eq '" + rqAccess.resourceAppId + "'"
                    };
                    return [4 /*yield*/, graphClient.servicePrincipals.list(resourceAppFilter)];
                case 1:
                    rs = _a.sent();
                    srv = rs[0];
                    desiredScope = "";
                    for (i = 0; i < rqAccess.resourceAccess.length; i++) {
                        rAccess = rqAccess.resourceAccess[i];
                        p = srv.oauth2Permissions.find(function (p) {
                            return p.id === rAccess.id;
                        });
                        desiredScope += p.value + " ";
                    }
                    now = new Date();
                    nextYear = new Date(now.getFullYear() + 1, now.getMonth(), now.getDay());
                    permissions = {
                        body: {
                            clientId: servicePrincipalId,
                            consentType: 'AllPrincipals',
                            scope: desiredScope,
                            resourceId: srv.objectId,
                            expiryTime: nextYear.toISOString()
                        }
                    };
                    return [4 /*yield*/, graphClient.oAuth2PermissionGrant.create(permissions)];
                case 2: return [2 /*return*/, _a.sent()];
            }
        });
    });
}
function run() {
    return __awaiter(this, void 0, void 0, function () {
        var azureEndpointSubscription, applicationName, ownerId, rootDomain, applicationSecret, requiredResource, homeUrl, taskReplyUrls, subcriptionId, servicePrincipalId, servicePrincipalKey, tenantId, azureCredentials, pipeCreds, graphClient, applicationInstance, newServicePrincipal, i, rqAccess, appUpdateParms, err_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 14, , 15]);
                    azureEndpointSubscription = tl.getInput("azureSubscriptionEndpoint", true);
                    applicationName = tl.getInput("applicationName", true);
                    ownerId = tl.getInput("applicationOwnerId", true);
                    rootDomain = tl.getInput("rootDomain", true);
                    applicationSecret = tl.getInput("applicationSecretPassword", true);
                    requiredResource = tl.getInput("requiredResource", true);
                    homeUrl = tl.getInput("homeUrl", true);
                    taskReplyUrls = tl.getInput("replyUrls", false);
                    subcriptionId = tl.getEndpointDataParameter(azureEndpointSubscription, "subscriptionId", false);
                    servicePrincipalId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalid", false);
                    servicePrincipalKey = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "serviceprincipalkey", false);
                    tenantId = tl.getEndpointAuthorizationParameter(azureEndpointSubscription, "tenantid", false);
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
                    return [4 /*yield*/, LoginToAzure(servicePrincipalId, servicePrincipalKey, tenantId)];
                case 1:
                    azureCredentials = _a.sent();
                    pipeCreds = new msRestNodeAuth.ApplicationTokenCredentials(azureCredentials.clientId, tenantId, azureCredentials.secret, 'graph');
                    graphClient = new azureGraph.GraphRbacManagementClient(pipeCreds, tenantId, { baseUri: 'https://graph.windows.net' });
                    return [4 /*yield*/, FindAzureAdApplication(applicationName, graphClient)];
                case 2:
                    applicationInstance = _a.sent();
                    if (!(applicationInstance == null)) return [3 /*break*/, 11];
                    return [4 /*yield*/, CreateOrUpdateADApplication(null, applicationName, rootDomain, applicationSecret, homeUrl, taskReplyUrls, requiredResource, graphClient)];
                case 3:
                    // Create new Azure AD Application
                    applicationInstance = _a.sent();
                    // Add Owner to new Azure AD Application
                    return [4 /*yield*/, AddADApplicationOwner(applicationInstance.objectId, ownerId, tenantId, graphClient)];
                case 4:
                    // Add Owner to new Azure AD Application
                    _a.sent();
                    return [4 /*yield*/, CreateServicePrincipal(applicationName, applicationInstance.appId, graphClient)];
                case 5:
                    newServicePrincipal = _a.sent();
                    i = 0;
                    _a.label = 6;
                case 6:
                    if (!(i < applicationInstance.requiredResourceAccess.length)) return [3 /*break*/, 9];
                    rqAccess = applicationInstance.requiredResourceAccess[i];
                    return [4 /*yield*/, grantAuth2Permissions(rqAccess, newServicePrincipal.objectId, graphClient)];
                case 7:
                    _a.sent();
                    _a.label = 8;
                case 8:
                    i++;
                    return [3 /*break*/, 6];
                case 9:
                    appUpdateParms = {
                        identifierUris: ['https://' + rootDomain + '/' + applicationInstance.appId]
                    };
                    return [4 /*yield*/, graphClient.applications.patch(applicationInstance.objectId, appUpdateParms)];
                case 10:
                    _a.sent();
                    return [3 /*break*/, 13];
                case 11: return [4 /*yield*/, CreateOrUpdateADApplication(applicationInstance.objectId, applicationName, rootDomain, applicationSecret, homeUrl, taskReplyUrls, requiredResource, graphClient)];
                case 12:
                    applicationInstance = _a.sent();
                    _a.label = 13;
                case 13:
                    tl.setVariable("azureAdApplicationId", applicationInstance.appId);
                    return [3 /*break*/, 15];
                case 14:
                    err_1 = _a.sent();
                    tl.setResult(tl.TaskResult.Failed, err_1.message || 'run() failed');
                    return [3 /*break*/, 15];
                case 15: return [2 /*return*/];
            }
        });
    });
}
run();
