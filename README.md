Tasks packages to manage Azure AD Application from Azure DevOps release pipeline.

Available tasks:
- GetAzureAdApplicationId
	- Get the Azure Ad Application id from the specified app name
- ManageAzureAdApplication
	- Create/Update Azure AD Application
	- Create/Update Azure AD Service principal for an Azure Ad Application
	- Create/Update Required resources of the application
	- Set the HomeUrl, Reply Urls
	- Set the Azure AD Application owner
	- Set Azure AD Application client secret

This task package is compatible with:
- Hosted macOS build agent
- Hosted vs2017 (soon)
- Hosted ubuntu 1640 (soon)
- Any private build agent with Azure CLI (v2.0.52 required)

<img src="https://dev.azure.com/experta/ExpertaSolutions/_apis/build/status/AzureAdApplicationExtensions-CI?branchName=master">

## GetAzureAdApplicationId (required parameters)
- Azure subscription
- Azure AD Admin user
- Azure AD Admin password

## ManageAzureApplication (required parameters)
- Azure subscription
- Azure AD Admin user
- Azure AD Admin password
- Application name
- Owner (user objectId - GUID)
- Application domain name
- Secret password
- Required resources (json)
	```
	 [{
		"resourceAppId": "00000002-0000-0000-c000-000000000000",
		"resourceAccess": [
           {
             "id": "311a71cc-e848-46a1-bdf8-97ff7156d8e6",
             "type": "Scope"
           },
           {
             "id": "5778995a-e1bf-45b8-affa-663a9f3f4d04",
             "type": "Scope"
           }
       ]
     }]
    ```
               
- Home Url
- Reply Urls (json)

	```["http://myurl.com","https://myurl.com]```

#Task Output variables
Both task returns the output variable ($ReferenceName.azureAdApplicationId)

# Requirements

- Azure CLI (v2.0.52) must be installed on the build agent

# Next features

- Set the owner of the Azure AD Application Service principal
- Set user/group access to Azure AD Application Service principal
- Upload application icon for the Azure AD Application