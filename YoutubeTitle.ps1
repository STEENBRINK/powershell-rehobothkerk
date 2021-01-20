#imports
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
Add-Type -AssemblyName System.Web

#defines
$clientID = “” #get from youtube api dev console
$clientSecret = “” #get from youtube api dev console
$redirectUrl = “http://localhost/oauth2callback”
$scope = “https://www.googleapis.com/auth/youtube.force-ssl”

#haal de refreshtoken uit de txt
$refreshToken = Get-Content $PSScriptRoot"\refreshToken.txt"

#gegevens om de inlogcode van de youtube api op te halen
$requestUri = “https://www.googleapis.com/oauth2/v3/token”
$refreshTokenParams = @{
  client_id=$clientID;
  client_secret=$clientSecret;
  refresh_token=$refreshToken;
  grant_type="refresh_token"; # Fixed value
}

$body = “refresh_token=$([System.Web.HttpUtility]::UrlEncode($refreshToken))&redirect_uri=$([System.Web.HttpUtility]::UrlEncode($redirectUrl))&client_id=$clientID&client_secret=$clientSecret&scope=$scope&grant_type=refresh_token”

#haal de accessToken op
$tokens = Invoke-RestMethod -Uri $requestUri -Method POST -Body $body

$accessToken = $tokens.access_token

#haal de volgende livestream ID op
$streams = (Invoke-RestMethod -URI "https://youtube.googleapis.com/youtube/v3/liveBroadcasts?part=snippet%2CcontentDetails%2Cstatus&broadcastType=all&mine=true" -METHOD GET -Headers @{Authorization = "Bearer $accessToken"} -ContentType 'application/json')

$streamID = $streams.items.id[0]
$lastStreamID = $streams.items.id[1]
$lastStreamTitle = $streams.items.snippet[1].title

#Kijk of de titel van de vorige livestream al is aangepast, zo niet doe dat dan
if($lastStreamTitle.StartsWith("Livestream ")) {
    #gegevens voor het aanpassen van de titel en beschrijving van de vorige video
    $newTitleLastStream = $lastStreamTitle.Replace("Livestream ", "")
    $requestUri = "https://www.googleapis.com/youtube/v3/liveBroadcasts?part=id&part=snippet&scope=$scope"
    $snippet = @{
        id= $lastStreamID;
        snippet= @{
            title = $newTitleLastStream;
        };
    }

    $body = ConvertTo-Json $snippet

    #pas de titel aan
    Invoke-RestMethod -Uri $requestUri -Method PUT -Headers @{Authorization = "Bearer $accessToken"} -ContentType 'application/json' -Body $body
}

#Haal de dienst op van de kerk api
$services = (Invoke-RestMethod -URI "https://rehobothkerkwoerden.nl/api/v1/agenda/services?page=1&page_size=1&reverseTime=false&search=" -METHOD GET -ContentType 'application/json')

$serviceTitle = $services.results[0].title
$serviceMinister = $services.results[0].minister

#set de variabelen voor de  titel en de beschrijving
$date = "" + (Get-Date).Day + "-" + (Get-Date).Month + "-" + (Get-Date).Year
$YouTubeTitle = "Livestream Rehobothkerk Woerden - " + $serviceTitle + " " + $date + " - " + $serviceMinister
$naam = Get-Content $PSScriptRoot"\naam.txt" #naam van de ds moet met een bestandje door het é teken
$YouTubeDescription =  '' + $serviceTitle + ' van ' + $date + ' door ' + $serviceMinister + '.

De GIVT link voor de collecte: https://bit.ly/givtrehoboth

De rehobothkerk gastvrije en warme gemeente geleid en gedreven door het evangelie, waarbinnen alle mensen zich met hun mogelijkheden voor elkaar inzetten. De gemeente staat bekend om de sterke onderlinge band en om de zorg naar binnen en naar buiten.
De kinderen en jongeren maken deel uit van de gemeente van Christus en krijgen de aandacht en ruimte die nodig is tot opbouw van hun geloof.

De liturgie voor de dienst is te vinden op onze website: https://rehobothkerkwoerden.nl/ (Alleen voor leden, i.v.m. privacyrechten)

Dr. ' + $naam + ' Jansen is sinds 2010 predikant van onze gemeente. Hij is geboren in Pretoria, Zuid-Afrika. Sinds 2000 verkondigt hij Gods Woord in Nederland. Zijn visie: "God verweeft Zijn liefde en plan met onze harten, monden en handen." 
Meer informatie over zijn proefschrift en publicaties is te lezen op zijn website: https://www.hearthandsandvoices.com.'

#gegevens voor het aanpassen van de titel en beschrijving
$requestUri = "https://www.googleapis.com/youtube/v3/liveBroadcasts?part=id&part=snippet&scope=$scope"
$snippet = @{
  id=$streamID;
  snippet= @{
    title = $YouTubeTitle;
    description=$YouTubeDescription;
  };
}

$body = ConvertTo-Json $snippet

#pas beschrijving en titel aan
Invoke-RestMethod -Uri $requestUri -Method PUT -Headers @{Authorization = "Bearer $accessToken"} -ContentType 'application/json' -Body $body