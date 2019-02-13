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

$applicationInfo = (az ad app show --id $applicationId --subscription $subscriptionId) | ConvertFrom-Json

if($applicationInfo.Length -eq 0) {
  write-host "Azure AD application with id '$applicationId' does not exist"
}
write-host ""

if($goodVersion -eq $true)
{
  # Remove the application here
  $result = az ad app delete --id $applicationId --subscription $subscriptionId

} else {
  write-host "Azure Cli Version: $major.$minor.$build doesn't provide set function on Application Owner change and Grant Application permissions"
}

$logoutResult = az account clear

write-host "Azure ApplicationID: $($applicationId)"