# This script calls the Power BI REST API to set the refresh schedule of a data set
# This script requires administrator permissions and the Azure PowerShell cmdlets
# Ensure the file paths for the cmdlet dlls are accurate for your enviroment

# Parameters - fill these in before running the script!
# =====================================================

$GroupId = ""    # the ID of the group (workspace) that hosts the dataset. Use "me" if this is your My Workspace
$DatasetId = ""  # the ID of the dataset that you will set the refresh schedule of
$FailureNotification = $true                         # set if you want the failure notification emails, set to $true or $false, if it has any other value it will default to TRUE.
$RefreshDays = ""                                    # set the days you want to refresh on, if left blank all days will be used
$RefreshFrequency = 4                                # set the frequency of refreshes in hours, e.g. set as 3 for every 3 hours, 0.5 for every 30 mins. Will be relative to midnight, must be between 12 and 0.5
$RefreshTimes = "07:00,08:00"                        # set the specific times of refresh, only used if RefreshFrequency is 0, must in the format e.g. "05:00, 13:00, 15:00"
$RefreshTimeZone = "UTC"                             # the ID of the timezone you want the refresh times to refer to. 
                                                     # See https://support.microsoft.com/en-gb/help/973627/microsoft-time-zone-index-values for the IDs

# AAD Client ID
# To get this, go to the following page and follow the steps to provision an app
# https://dev.powerbi.com/apps
# To get the sample to work, ensure that you have the following fields:
# App Type: Native app
# Redirect URL: urn:ietf:wg:oauth:2.0:oob
#  Level of access: all dataset APIs

$clientId = "" 

# End Parameters =======================================

if ($clientId -eq "") {
    throw "You need to include a client id for this script to authenticate"
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

#Format the JSON body

if ($FailureNotification) {
    $FailureNotification = "MailOnFailure"
} else {
    $FailureNotification = "NoNotification"
}

if ($RefreshDays -eq "") {
    $RefreshDays = @("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
} else {
    $RefreshDays =  $RefreshDays.split(",")
}

if ($RefreshFrequency -is [String] ) {
    throw "RefreshFrequency must not be a string, if you want to use RefreshTimes set RefreshFrequency to 0"
}

if  ($RefreshFrequency -eq 0) { #test if RefreshFrequency is 0
    #Use refresh times
    $bodyTimes = $RefreshTimes.split(",")
} else { # use refresh frequency
    $timesCount = [math]::Floor(24/$RefreshFrequency) - 1 #Calculate the number of refreshes needed per 24 hours
    $bodyTimes = @("00:00")
    for ($i=1;$i -le $timesCount;$i++) { # Append the 
        $time = ($i * $RefreshFrequency).ToString("00.00").Replace(".",":").Replace("50","30")
        $bodyTimes += $time
    }
}

$patchBody = @{
    "value" = @{
        "NotifyOption" = $FailureNotification
        "days" = $RefreshDays
        "localTimeZoneId" = $RefreshTimeZone
        "times" = $bodyTimes
    }
}

$jsonPatchBody = $patchBody | ConvertTo-JSON
#PATCH refresh times
$uri = "https://api.powerbi.com/v1.0/$GroupsPath/datasets/$DatasetId/refreshSchedule"
Invoke-RestMethod -Uri $uri -Headers $authHeader -Method PATCH -Body $jsonPatchBody -Verbose
