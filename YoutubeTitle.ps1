#imports
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#defines
$clientID = “”
$clientSecret = “”
$redirectUrl = “http://localhost/oauth2callback”
$refreshToken = Get-Content $PSScriptRoot"\refreshToken.txt"

################################################################
#  Vraag of het om een speciale dienst gaat (rouw/trouw/etc)   #
################################################################

$standaardDienstInput =  [System.Windows.MessageBox]::Show(
'Gaat het om een standaard dienst? 
Kies "No" bij rouwdiensten, trouwdiensten, etc.','Standaard Dienst'
    ,'YesNoCancel')

  switch  ($standaardDienstInput) {

  'Yes' {

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
    #$streamStartTime = ($services.results[$i].startdatetime).Replace(":30", ":15") + "+1:00"

  }

  'No' {
  
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Youtube Titel'
    $form.Size = New-Object System.Drawing.Size(395,175)
    $form.StartPosition = 'CenterScreen'

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(10,100)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'Volgende'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(95,100)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Annuleren'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    $tekstbovenaan = New-Object System.Windows.Forms.Label
    $tekstbovenaan.Location = New-Object System.Drawing.Point(10,10)
    $tekstbovenaan.Size = New-Object System.Drawing.Size(400,25)
    $tekstbovenaan.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $tekstbovenaan.Text = "Stel de titel in (let op spelling)" #verandertekstbovenaan
    $form.Controls.Add($tekstbovenaan)

    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,40)
    $label1.Size = New-Object System.Drawing.Size(160,20)
    $label1.Text = 'Titel (bijv. Trouwdienst X - Y):'
    $form.Controls.Add($label1)

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,65)
    $label2.Size = New-Object System.Drawing.Size(160,20)
    $label2.Text = 'Naam voorganger:               ds.'
    $form.Controls.Add($label2)

    $titelInput = New-Object System.Windows.Forms.TextBox
    $titelInput.Location = New-Object System.Drawing.Point(170,37)
    $titelInput.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($titelInput)

    $voorgangerInput = New-Object System.Windows.Forms.TextBox
    $voorgangerInput.Location = New-Object System.Drawing.Point(170,62)
    $voorgangerInput.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($voorgangerInput)

    $form.Topmost = $true

    $form.Add_Shown({$titelInput.Select()})
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $serviceTitle = $titelInput.Text
        $serviceMinister = $voorgangerInput.Text

        $confirmationForm = New-Object System.Windows.Forms.Form
        $confirmationForm.Text = 'Klopt de Youtube Titel'
        $confirmationForm.Size = New-Object System.Drawing.Size(620,125)
        $confirmationForm.StartPosition = 'CenterScreen'

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(10,50)
        $okButton.Size = New-Object System.Drawing.Size(75,23)
        $okButton.Text = 'Ja'
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $confirmationForm.AcceptButton = $okButton
        $confirmationForm.Controls.Add($okButton)

        $noButton = New-Object System.Windows.Forms.Button
        $noButton.Location = New-Object System.Drawing.Point(95,50)
        $noButton.Size = New-Object System.Drawing.Size(75,23)
        $noButton.Text = 'Nee'
        $noButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $confirmationForm.CancelButton = $noButton
        $confirmationForm.Controls.Add($noButton)

        $tekstbovenaan = New-Object System.Windows.Forms.Label
        $tekstbovenaan.Location = New-Object System.Drawing.Point(10,5)
        $tekstbovenaan.Size = New-Object System.Drawing.Size(400,20)
        $tekstbovenaan.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
        $tekstbovenaan.Text = 'Klopt deze titel?'
        $confirmationForm.Controls.Add($tekstbovenaan)

        $label1 = New-Object System.Windows.Forms.Label
        $label1.Location = New-Object System.Drawing.Point(10,25)
        $label1.Size = New-Object System.Drawing.Size(600,20)
        $label1.Text = ('Livestream ' + $serviceTitle + ' - ' + (Get-Date -Format "dd-MM-yyyy") + ' - ds. ' + $serviceMinister + ' - Rehobothkerk Woerden')
        $confirmationForm.Controls.Add($label1)

        $confirmationForm.Topmost = $true

        $confirmed = $confirmationForm.ShowDialog()

        if ($confirmed -eq [System.Windows.Forms.DialogResult]::Cancel)
        {
            Exit
        }
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel){
        Exit
    }
  }

  'Cancel' {
    Exit
  }
}

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


##############################################
#  Prepareer de gegevens voor de livestream  #
##############################################

#set de variabelen voor de  titel en de beschrijving
#$date = "" + $streamStartTime.Substring(8, 2) + $streamStartTime.Substring(4, 4) + $streamStartTime.Substring(0, 4)
$date = Get-Date -Format "dd-MM-yyyy"
$YouTubeTitle = "Livestream - " + $serviceTitle + " " + $date + " - " + $serviceMinister + " - Rehobothkerk Woerden"
$naamAndre = Get-Content $PSScriptRoot"\naam.txt" #naam van de ds moet met een bestandje door het é teken
$YouTubeDescription =  '' + $serviceTitle + ' van ' + $date + ' door ' + $serviceMinister + '.

De GIVT link voor de collecte: https://bit.ly/givtrehoboth

De rehobothkerk gastvrije en warme gemeente geleid en gedreven door het evangelie, waarbinnen alle mensen zich met hun mogelijkheden voor elkaar inzetten. De gemeente staat bekend om de sterke onderlinge band en om de zorg naar binnen en naar buiten.
De kinderen en jongeren maken deel uit van de gemeente van Christus en krijgen de aandacht en ruimte die nodig is tot opbouw van hun geloof.

De liturgie voor de dienst is te vinden op onze website: https://rehobothkerkwoerden.nl/ (Alleen voor leden, i.v.m. privacyrechten)

Dr. ' + $naamAndre + ' Jansen is sinds 2010 predikant van onze gemeente. Hij is geboren in Pretoria, Zuid-Afrika. Sinds 2000 verkondigt hij Gods Woord in Nederland. Zijn visie: "God verweeft Zijn liefde en plan met onze harten, monden en handen." 
Meer informatie over zijn proefschrift en publicaties is te lezen op zijn website: https://www.hearthandsandvoices.com.'

$streamStartTime = "1970-01-01T00:00:00Z"

$streamID = (Invoke-RestMethod -URI "https://youtube.googleapis.com/youtube/v3/liveBroadcasts?part=snippet%2CcontentDetails%2Cstatus&broadcastStatus=upcoming&broadcastType=all" -METHOD GET -Headers @{Authorization = "Bearer $accessToken"} -ContentType 'application/json').items[0].id


#############################################
#  Pas de gegevens voor de livestream  aan  #
#############################################

#gegevens voor het aanpassen van de titel en beschrijving
$requestUri = "https://youtube.googleapis.com/youtube/v3/liveBroadcasts?part=snippet&part=status"
$streamData = @{
  id = $streamID;
  snippet= @{
    title = $YouTubeTitle;
    description=$YouTubeDescription;
    scheduledStartTime= $streamStartTime;
  };
  status= @{
    privacyStatus="public";
    selfDeclaredMadeForKids="false";
  };
}

$body = ConvertTo-Json $streamData

#pas beschrijving en titel aan
Invoke-RestMethod -Uri $requestUri -Method PUT -Headers @{Authorization = "Bearer $accessToken"; Accept = "application/json";} -ContentType 'application/json' -Body $body