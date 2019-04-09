# This script calls the Power BI REST API to stagger datasets throughout a 24 hour period
# This script requires administrator permissions and the Azure PowerShell cmdlets
# Ensure the file paths for the cmdlet dlls are accurate for your enviroment

# Parameters - fill these in before running the script!
# =====================================================

$Email = ""     #Email used to filter datasets by owner
$GroupId = ""    # the ID of the group (workspace) that hosts the dataset. Use "me" if this is your My Workspace
$AllDatasets = $true                                 # set if you want all datasets in the workspace to be set
# comma delimited list of datasets to stagger the refresh of.
$DataSets = ""   

# AAD Client ID
# To get this, go to the following page and follow the steps to provision an app
# https://dev.powerbi.com/apps
# To get the sample to work, ensure that you have the following fields:
# App Type: Native app
# Redirect URL: urn:ietf:wg:oauth:2.0:oob
# Level of access: all dataset APIs

$clientId = "" 

# End Parameters =======================================

if ($clientId -eq "") {
    throw "You need to include a client id for this script to authenticate"
}
if (($AllDatasets -ne $false -and $AllDatasets -ne $true) -or ($AllDatasets -eq $false -and $DataSets -eq "")) {
    throw "Check the AllDatasets and DataSets Parameters"
}

# Calls the Active Directory Authentication Library (ADAL) to authenticate against AAD
function GetAuthToken
{
    $adal = "${env:ProgramFiles}\WindowsPowerShell\Modules\AzureRM.profile\5.8.2\Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
    
    $adalforms = "${env:ProgramFiles}\WindowsPowerShell\Modules\AzureRM.profile\5.8.2\Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll"
 
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null

    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null

    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"

    $resourceAppIdURI = "https://analysis.windows.net/powerbi/api"

    $authority = "https://login.microsoftonline.com/common/oauth2/authorize";

    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority

    $authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId, $redirectUri, "Auto")

    return $authResult
}

# Get the auth token from AAD
$token = GetAuthToken

# Building Rest API header with authorization token
$authHeader = @{
   'Content-Type'='application/json'
   'Authorization'=$token.CreateAuthorizationHeader()
}

# properly format groups path
$GroupsPath = ""
if ($GroupId -eq "me") {
    $GroupsPath = "myorg"
} else {
    $GroupsPath = "myorg/groups/$GroupId"
}

# Format dataset list

if ($AllDatasets) {
    # Get datasets
    $uri = "https://api.powerbi.com/v1.0/$GroupsPath/datasets/"
    $Datasets = (Invoke-RestMethod -Uri $uri -Headers $authHeader -Method GET -Verbose).value
    $Datasets = $Datasets.where({$_.configuredBy -eq $Email -and $_.isRefreshable -eq $True}).id
} else {
    $DataSets = $DataSets.Split(",")
}


# set the staggered times so that the refreshes are equally spread
$refreshTimes = @("00:00")
$StaggerLength = [math]::Floor(24/$Datasets.Length)
if ($StaggerLength -lt 0.5) {
    $StaggerLength = 0.5
}
$timesCount = ([math]::Floor($Datasets.Length)) - 1
for ($i=1;$i -le $timesCount;$i++) {
    $time = (($i * $StaggerLength)%24).ToString("00.00").Replace(".",":").Replace("50","30")
    $refreshTimes += $time
}

$counter = 0
foreach($i in $DataSets) {
    $patchBody = @{
        "value" = @{
            "times" = @($refreshTimes[$counter])
        }
    }
    $counter++
    $jsonPatchBody = $patchBody | ConvertTo-JSON
    #PATCH refresh time
    $uri = "https://api.powerbi.com/v1.0/$GroupsPath/datasets/$i/refreshSchedule"
    Invoke-RestMethod -Uri $uri -Headers $authHeader -Method PATCH -Body $jsonPatchBody -Verbose
}
