$date = Get-Date -Format "yyyyMMdd-HH.mm.ss"
$path =  $PSScriptRoot
Start-Transcript -Path $path\LogFile_$date.txt

# Fuction for accessing the Token
function get_token()
{    
    # Application (client) ID, tenant ID and secret
    $clientId = "Your Client ID"
    $tenantId = "Your Tenant ID"
    $clientSecret = 'Your client Secret'

    
    # Construct URI
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
 
    # Construct Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }
 
    # Get OAuth 2.0 Token
    $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
 
    # Access Token
    $token = ($tokenRequest.Content | ConvertFrom-Json).access_token
    return $token
}


# Base URL
$base_URL = "https://graph.microsoft.com/beta/reports/"

# Function extract all reports
function extractM365AdminReports(

    [parameter(mandatory)] [string] $api_endpoint,
    [parameter(mandatory)] [string] $report_period
)
{
    $url = $base_URL+$api_endpoint+"(period='"+$report_period+"')"
    $email = Invoke-RestMethod -Method GET -Uri $url -Headers @{Authorization = "Bearer $token"}
    $email = $email.Replace('ï»¿Report Refresh Date','Report Refresh Date')
    $emailactivity = ConvertFrom-Csv -InputObject $email
    $emailactivity | Export-Csv -Path $path\$api_endpoint.csv -NoTypeInformation
}

# Get the Access Token
$token = get_token

# Calling the function to extract the reports, Please select report period as D7/D30/D60/D90

extractM365AdminReports -api_endpoint "getEmailActivityUserDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getSharePointActivityUserDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getSkypeForBusinessActivityUserDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getTeamsUserActivityUserDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getYammerActivityUserDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getOneDriveActivityUserDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getOffice365ActiveUserDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getOneDriveUsageAccountDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getSharePointSiteUsageDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getOffice365GroupsActivityDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getYammerGroupsActivityDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getTeamsDeviceUsageUserDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getSkypeForBusinessDeviceUsageUserDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getYammerDeviceUsageUserDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getEmailAppUsageUserDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getMailboxUsageDetail" -report_period "D30"
extractM365AdminReports -api_endpoint "getOffice365ActivationsUserDetail" -report_period "D30"


Stop-Transcript
