# Summary
Tasks packages to manage Azure AD Application from Azure DevOps release pipeline.

## Available tasks

### GetAzureAdApplicationId (required parameters)
- Azure subscription

### ManageAzureApplication (required parameters)
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

### RemoveAzureADApplication (required parameters)
- Azure AD Application ID