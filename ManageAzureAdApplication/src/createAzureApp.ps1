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

$loginResult = az login --service-principal -u $servicePrincipalId -p $servicePrincipalKey --tenant $tenantId
$setResult = az account set --subscription $subscriptionId

try {
  $test = az --version
} catch {
  write-host "Azure Cli not installed"
  throw;
}

$versionResult = az --version
$result = [regex]::Match($versionResult, "azure-cli \((([0-9]*).([0-9]*).([0-9]*))\)").captures.groups

if($result.length -eq 5)
{
  $major = $result[2].value
  $minor = $result[3].Value
  $build = $result[4].value
}
$goodVersion = $false

write-host "Azure Cli Version '$major.$minor.$build' installed on build agent"

if($major -ge 2 -and $minor -eq 0 -and $build -ge 52){
  $goodVersion = $true
}

if($homeUrl.length -eq 0)
{
  $homeUrl = "http://$applicationName.$rootDomain"
}

if($replyUrls.length -eq 0)
{
  $replyUrls = "['http://$applicationName.$rootDomain', 'http://$applicationName.$rootDomain/signin-oidc','http://$applicationName.$rootDomain/signin-aad']"
}

$applicationInfo = (az ad app list --filter "displayName eq '$applicationName'" --subscription $subscriptionId) | ConvertFrom-Json

$applicationId = ""

if($applicationInfo.Length -eq 0) {
  write-host "*****"
  $servicePrincipalResult = $(az ad sp create-for-rbac --name "http://$applicationName" --password $applicationSecret --debug --verbose --subscription $subscriptionId) | ConvertFrom-Json
  write-host "*****"
  $applicationId = $servicePrincipalResult.appId
} else {
  $applicationId = $applicationInfo.appId
}
write-host ""

# Set the IdentifierUris
write-host "Set IdentifierUris... " -NoNewline
#$result = az ad app update --id $applicationId --set identifierUris="['https://$rootDomain/$($applicationId)']" --subscription $subscriptionId
write-host " Done"

# Set the homepage url
write-host "Set homepage url... " -NoNewline
#$result = az ad app update --id $applicationId --set homepage="$homeUrl" --subscription $subscriptionId
write-host " Done"

# Set the reply urls
write-host "Set Reply urls... " -NoNewline
#$result = az ad app update --id $applicationId --set replyUrls=$($replyUrls.replace('"',"'")) --subscription $subscriptionId
write-host " Done"

# Reset the Application Password
write-host "Set application password... " -NoNewline
#$result = az ad app update --id $applicationId --password $applicationSecret --subscription $subscriptionId
write-host " Done"

# Apply the Required Resources
write-host "Set Required resources accesses... " -NoNewline
#$result = az ad app update --id $applicationId --required-resource-accesses $manifestFile --subscription $subscriptionId
write-host " Done"

# Sets the Application Owner

if($goodVersion -eq $true)
{
  #$ownerList = (az ad app owner list --id $applicationId --subscription $subscriptionId | ConvertFrom-Json) | Where-Object { $_.objectId -eq $ownerId }
  if ($ownerList.length -eq 0)
  {
    write-host "Set Application Owner..." -NoNewline
    #az ad app owner add --id $applicationId --owner-object-id $ownerId --subscription $subscriptionId
    write-host " Done"
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
      #$grantResult = az ad app permission grant --id $applicationId --api $appId --subscription $subscriptionId
      write-host "Granted" -ForegroundColor Green
    }
  }
} else {
  write-host "Azure Cli Version: $major.$minor.$build doesn't provide set function on Application Owner change and Grant Application permissions"
}

$logoutResult = az account clear

write-host "Azure ApplicationID: $($applicationId)"