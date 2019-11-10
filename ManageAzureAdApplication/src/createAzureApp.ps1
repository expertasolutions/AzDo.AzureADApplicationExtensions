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
    [string]$applicationName
  , [Parameter(Mandatory=$true, Position=6)]
    [string]$rootDomain
  , [Parameter(Mandatory=$true, Position=7)]
    [string]$applicationSecret
  , [Parameter(Mandatory=$true, Position=8)]
    [string]$manifestFile
  , [Parameter(Mandatory=$false, Position=9)]
    [string]$homeUrl
  , [Parameter(Mandatory=$false, Position=10)]
    [string]$replyUrls
  , [Parameter(Mandatory=$false, Position=11)]
    [string]$ownerId
)

az login --service-principal -u $servicePrincipalId -p $servicePrincipalKey --tenant $tenantId | Out-Null
az account set --subscription $subscriptionId | Out-Null

if($homeUrl.length -eq 0)
{
  $homeUrl = "http://$applicationName.$rootDomain"
}

if($replyUrls.length -eq 0)
{
  $replyUrls = "['http://$applicationName.$rootDomain', 'http://$applicationName.$rootDomain/signin-oidc','http://$applicationName.$rootDomain/signin-aad']"
}

$applicationInfo = (az ad app list --filter "displayName eq '$applicationName'") | ConvertFrom-Json
$applicationId = ""

if($applicationInfo.Length -eq 0) {
  write-host "*****"
  $servicePrincipalResult = $(az ad sp create-for-rbac --name $applicationName --password $applicationSecret) | ConvertFrom-Json
  write-host "*****"
  $applicationId = $servicePrincipalResult.appId
} else {
  $applicationId = $applicationInfo.appId
}
write-host ""

# Set the IdentifierUris
write-host "Set IdentifierUris... " -NoNewline
$result = az ad app update --id $applicationId --set identifierUris="['https://$rootDomain/$($applicationId)']"
write-host " Done"

# Set the homepage url
write-host "Set homepage url... " -NoNewline
$result = az ad app update --id $applicationId --set homepage="$homeUrl"
write-host " Done"

# Set the reply urls
write-host "Set Reply urls... " -NoNewline
$result = az ad app update --id $applicationId --set replyUrls=$($replyUrls.replace('"',"'"))
write-host " Done"

# Reset the Application Password
write-host "Set application password... " -NoNewline
$result = az ad app update --id $applicationId --password $applicationSecret
write-host " Done"

# Apply the Required Resources
write-host "Set Required resources accesses... " -NoNewline
$result = az ad app update --id $applicationId --required-resource-accesses $manifestFile
write-host " Done"

# Sets the Application Owner

$ownerList = (az ad app owner list --id $applicationId | ConvertFrom-Json) | Where-Object { $_.objectId -eq $ownerId }
if ($ownerList.length -eq 0)
{
#  write-host "Set Application Owner..." -NoNewline
#  az ad app owner add --id $applicationId --owner-object-id $ownerId
#  write-host " Done"
}

#Granting Permission to service principal
$perms = (az ad app permission list --id $applicationId) | ConvertFrom-Json
$perms | ForEach-Object {
  $appId = $_.resourceAppId
  $granted = $_.grantedTime
  write-host "  Api: '$( $appId )' - " -NoNewline
  if ($granted)
  {
    write-host "Already granted ($granted)" -ForegroundColor Yellow
  }
  else
  {
    $grantResult = az ad app permission grant --id $applicationId --api $appId
    write-host "Granted" -ForegroundColor Green
  }
}

az account clear | Out-Null

write-host "Azure ApplicationID: $($applicationId)"