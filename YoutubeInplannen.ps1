#imports
Add-Type -AssemblyName System.Web

#defines
$clientID = “”
$clientSecret = “”
$redirectUrl = “http://localhost/oauth2callback”
$refreshToken = Get-Content $PSScriptRoot"\refreshToken.txt"


############################################################
#  Haal de nieuwe accesscode op om de wijzigingen te doen  #
############################################################

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


###############################################
#  Pas de titel van de vorige livestream aan  #
###############################################

$lastStream = (Invoke-RestMethod -URI "https://youtube.googleapis.com/youtube/v3/liveBroadcasts?part=snippet%2CcontentDetails%2Cstatus&broadcastStatus=completed&broadcastType=all" -METHOD GET -Headers @{Authorization = "Bearer $accessToken"} -ContentType 'application/json')

$lastStreamID = $lastStream.items.id[0]
$lastStreamTitle = $lastStream.items.snippet[0].title
$lastStreamStartTime = $lastStream.items.snippet[0].scheduledStartTime

#Kijk of de titel van de vorige livestream al is aangepast, zo niet doe dat dan
if($lastStreamTitle.StartsWith("Livestream ")) {
    #gegevens voor het aanpassen van de titel en beschrijving van de vorige video
    $newTitleLastStream = $lastStreamTitle.Replace("Livestream - ", "")
    $requestUri = "https://www.googleapis.com/youtube/v3/liveBroadcasts?part=id&part=snippet"
    $data = @{
        id= $lastStreamID;
        snippet= @{
            title = $newTitleLastStream;
            scheduledStartTime = $lastStreamStartTime;
        };
    }

    $body = ConvertTo-Json $data

    #pas de titel aan
    Invoke-RestMethod -Uri $requestUri -Method PUT -Headers @{Authorization = "Bearer $accessToken"} -ContentType 'application/json' -Body $body
}


###########################################
#  Haal De gegevens van de kerkdienst op  #
###########################################

#Haal de dienst op van de kerk api
$services = (Invoke-RestMethod -URI "https://rehobothkerkwoerden.nl/api/v1/agenda/services?page=1&page_size=3&reverseTime=false&search=" -METHOD GET -ContentType 'application/json')
$today = Get-Date -Format "yyyy-MM-dd"

$i = 1

if($services.results[1].startdatetime.Contains($today) -and ((Get-Date).hour -gt 12)){
    $i = 2
}

$serviceTitle = $services.results[$i].title
$serviceMinister = $services.results[$i].minister
$streamStartTime = ($services.results[$i].startdatetime).Replace(":30", ":15") + "+1:00"


##############################################
#  Prepareer de gegevens voor de livestream  #
##############################################

#set de variabelen voor de  titel en de beschrijving
$date = "" + $streamStartTime.Substring(8, 2) + $streamStartTime.Substring(4, 4) + $streamStartTime.Substring(0, 4)
$YouTubeTitle = "Livestream - " + $serviceTitle + " " + $date + " - " + $serviceMinister + " - Rehobothkerk Woerden"
$naamAndre = Get-Content $PSScriptRoot"\naam.txt" #naam van de ds moet met een bestandje door het é teken
$YouTubeDescription =  '' + $serviceTitle + ' van ' + $date + ' door ' + $serviceMinister + '.

De GIVT link voor de collecte: https://bit.ly/givtrehoboth

De rehobothkerk gastvrije en warme gemeente geleid en gedreven door het evangelie, waarbinnen alle mensen zich met hun mogelijkheden voor elkaar inzetten. De gemeente staat bekend om de sterke onderlinge band en om de zorg naar binnen en naar buiten.
De kinderen en jongeren maken deel uit van de gemeente van Christus en krijgen de aandacht en ruimte die nodig is tot opbouw van hun geloof.

De liturgie voor de dienst is te vinden op onze website: https://rehobothkerkwoerden.nl/ (Alleen voor leden, i.v.m. privacyrechten)

Dr. ' + $naamAndre + ' Jansen is sinds 2010 predikant van onze gemeente. Hij is geboren in Pretoria, Zuid-Afrika. Sinds 2000 verkondigt hij Gods Woord in Nederland. Zijn visie: "God verweeft Zijn liefde en plan met onze harten, monden en handen." 
Meer informatie over zijn proefschrift en publicaties is te lezen op zijn website: https://www.hearthandsandvoices.com.'

#gegevens voor het aanpassen van de titel en beschrijving
$requestUri = "https://youtube.googleapis.com/youtube/v3/liveBroadcasts?part=snippet&part=status&part=contentDetails"
$streamData = @{
  snippet= @{
    title = $YouTubeTitle;
    description=$YouTubeDescription;
    scheduledStartTime= $streamStartTime;
  };
  status= @{
    privacyStatus="unlisted";
    selfDeclaredMadeForKids="false";
  };
  contentDetails = @{
    enableAutoStart="true";
    enableAutoStop="true";
  };
}

$body = ConvertTo-Json $streamData

#pas beschrijving en titel aan
Invoke-RestMethod -Uri $requestUri -Method POST -Headers @{Authorization = "Bearer $accessToken"; Accept = "application/json";} -ContentType 'application/json' -Body $body