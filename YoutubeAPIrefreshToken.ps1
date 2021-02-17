$clientID = “”
$clientSecret = “”
$redirectUrl = “http://localhost/oauth2callback”
$scope = “https://www.googleapis.com/auth/youtube.force-ssl”

# Replace with your client id
$authUrl = “https://accounts.google.com/o/oauth2/auth?redirect_uri=$redirectUrl&client_id=$clientID&scope=$scope&approval_prompt=force&access_type=offline&response_type=code”

$authUrl | clip
# Copy the generated url into your browser
PAUSE

$responseCode = “” #paste code here before continueing

$requestUri = “https://www.googleapis.com/oauth2/v3/token”

$body = “code=$([System.Web.HttpUtility]::UrlEncode($responseCode))&redirect_uri=$([System.Web.HttpUtility]::UrlEncode($redirectUrl))&client_id=$clientID&client_secret=$clientSecret&scope=$scope&grant_type=authorization_code”

$tokens = Invoke-RestMethod -Uri $requestUri -Method POST -Body $body -ContentType “application/x-www-form-urlencoded”

Set-Content $PSScriptRoot"\refreshToken.txt" $tokens.refresh_token