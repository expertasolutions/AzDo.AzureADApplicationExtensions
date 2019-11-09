# Summary
Tasks packages to manage Azure AD Application from Azure DevOps release pipeline.

## Available tasks:
- GetAzureAdApplicationId
	- Get the Azure Ad Application id from the specified app name
- ManageAzureAdApplication
	- Create/Update Azure AD Application
	- Create/Update Azure AD Service principal for an Azure Ad Application
	- Create/Update Required resources of the application
	- Set the HomeUrl, Reply Urls
	- Set the Azure AD Application owner
	- Set Azure AD Application client secret
- RemoveAzureADApplication
	- Remove an existing Azure AD Application

## GetAzureAdApplicationId (required parameters)
- Azure subscription

## ManageAzureApplication (required parameters)
- Azure subscription
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

## RemoveAzureADApplication (required parameters)
- Azure AD Application ID

#Task Output variables
Both task returns the output variable ($ReferenceName.azureAdApplicationId)