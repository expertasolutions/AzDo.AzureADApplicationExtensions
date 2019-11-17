param(
    [Parameter(Mandatory=$true, Position=1)]
    [string]$subscriptionId
  , [Parameter(Mandatory=$true, Position=2)]
    [string]$servicePrincipalId
  , [Parameter(Mandatory=$true, Position=3)]
    [string]$servicePrincipalKey
  , [Parameter(Mandatory=$true, Position=4)]
    [string]$tenantId
  , [Parameter(Mandatory=$true, Position=5)]
    [string]$applicationId
)

az login --service-principal -u $servicePrincipalId -p $servicePrincipalKey --tenant $tenantId | Out-Null
az account set --subscription $subscriptionId | Out-Null

$applicationInfo = (az ad app show --id $applicationId) | ConvertFrom-Json

if($applicationInfo.Length -eq 0) {
  write-host "Azure AD application with id '$applicationId' does not exist"
}
write-host ""

# Remove the application here
$result = az ad app delete --id $applicationId

az account clear | out-null

write-host "Azure ApplicationID: $($applicationId)"