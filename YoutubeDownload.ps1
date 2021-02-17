#imports
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -path $PSScriptRoot'\SharpOSC.dll'

#globals
$companion_url = "" #companion ip
$companion_OSC_port = 12321
$companion_videos_page = 9
$spotify_color = [PSCustomObject]@{
    r = 100
    g = 25
    b = 15
}
$normal_color = [PSCustomObject]@{
    r = 0
    g = 0
    b = 0
}
$companion_button_numbers = [PSCustomObject]@{
    kindermoment = 0
    kinderlied = 0
    lied1 = 3
    lied2 = 0
    lied3 = 0
    lied4 = 0
    lied5 = 0
    lied6 = 0
    lied7 = 0
    extra1 = 0
    extra2 = 0
    extra3 = 0
}

#set button title for companion based on the video name (e.g. lied1)
function Set-Companion-Button-Name {
    Param([String]$video_name, [String]$title, [bool]$checked)

    $button_number = -1;

    $companion_button_numbers.PSObject.Properties | ForEach-Object {
        if($_.Name -eq $video_name){
            $button_number = $_.Value
        }
    }

    if($button_number -ne -1) {
        $sender = new-object SharpOSC.UDPSender $companion_url, $companion_OSC_port
        $message = new-object SharpOSC.OscMessage ("/style/text/" + $companion_videos_page + "/" + $button_number), $title;
        $sender.Send($message);

        if($checked){
            $color_message = new-object SharpOSC.OscMessage ("/style/bgcolor/" + $companion_videos_page + "/" + $button_number), ($spotify_color.r, $spotify_color.g, $spotify_color.b);
        }else {
            $color_message = new-object SharpOSC.OscMessage ("/style/bgcolor/" + $companion_videos_page + "/" + $button_number), ($normal_color.r, $normal_color.g, $normal_color.b)
        }
        $sender.Send($color_message);
    }
    else {
        [System.Windows.MessageBox]::Show('Er ging iets fout, bekijk de error voor detail','ERROR','OK','Error')
        Set-Content ([Environment]::GetFolderPath("MyVideos") + "\autogendienst\error.txt") "Er is geen companion knop gevonden voor de videonaam"
    }
}

#check if input is valid ur, then download the video
function Download-Video {
    Param([string]$url,[string]$video_name)

    if((@($url) | .\CheckURL.ps1).IsValid){
        
        $path = ([Environment]::GetFolderPath("MyVideos") + "\autogendienst\" + $video_name + ".mp4")

        try {
            ./youtube-dl -f "bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best" -o $path $url --write-info-json

            $title = (Get-Content ([Environment]::GetFolderPath("MyVideos") + "\autogendienst\" + $video_name + ".info.json") | Out-String | ConvertFrom-Json).title

            $title = $title.Replace("Nederland Zingt: ", "")
            
            Set-Companion-Button-Name -video_name $video_name -title $title -checked $false
        }
        catch
        {
            [System.Windows.MessageBox]::Show('Er ging iets fout, bekijk de error voor detail','ERROR','OK','Error')
            Set-Content ([Environment]::GetFolderPath("MyVideos") + "\autogendienst\error.txt") $_
        }
    }else{
        [System.Windows.MessageBox]::Show($url + ' voor ' + $video_name + ' is geen geldige URL','ERROR','OK','Error')
    }
}

function Show-Form {
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Youtube Downloader'
    $form.Size = New-Object System.Drawing.Size(340,430)
    $form.StartPosition = 'CenterScreen'

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75,355)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(170,355)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    $tekstbovenaan = New-Object System.Windows.Forms.Label
    $tekstbovenaan.Location = New-Object System.Drawing.Point(10,20)
    $tekstbovenaan.Size = New-Object System.Drawing.Size(300,15)
    $tekstbovenaan.Text = "Voer de URL's in    (Vink voor spotify, voor dan naam in)" #verandertekstbovenaan
    $form.Controls.Add($tekstbovenaan)

    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,50)
    $label1.Size = New-Object System.Drawing.Size(80,20)
    $label1.Text = 'Kindermoment'
    $form.Controls.Add($label1)

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,75)
    $label2.Size = New-Object System.Drawing.Size(80,20)
    $label2.Text = 'Kinderlied'
    $form.Controls.Add($label2)

    $label3 = New-Object System.Windows.Forms.Label
    $label3.Location = New-Object System.Drawing.Point(10,100)
    $label3.Size = New-Object System.Drawing.Size(50,20)
    $label3.Text = 'Lied 1'
    $form.Controls.Add($label3)

    $label4 = New-Object System.Windows.Forms.Label
    $label4.Location = New-Object System.Drawing.Point(10,125)
    $label4.Size = New-Object System.Drawing.Size(50,20)
    $label4.Text = 'Lied 2'
    $form.Controls.Add($label4)

    $label5 = New-Object System.Windows.Forms.Label
    $label5.Location = New-Object System.Drawing.Point(10,150)
    $label5.Size = New-Object System.Drawing.Size(50,20)
    $label5.Text = 'Lied 3'
    $form.Controls.Add($label5)

    $label6 = New-Object System.Windows.Forms.Label
    $label6.Location = New-Object System.Drawing.Point(10,175)
    $label6.Size = New-Object System.Drawing.Size(50,20)
    $label6.Text = 'Lied 4'
    $form.Controls.Add($label6)

    $label7 = New-Object System.Windows.Forms.Label
    $label7.Location = New-Object System.Drawing.Point(10,200)
    $label7.Size = New-Object System.Drawing.Size(50,20)
    $label7.Text = 'Lied 5'
    $form.Controls.Add($label7)

    $label8 = New-Object System.Windows.Forms.Label
    $label8.Location = New-Object System.Drawing.Point(10,225)
    $label8.Size = New-Object System.Drawing.Size(50,20)
    $label8.Text = 'Lied 6'
    $form.Controls.Add($label8)

    $label9 = New-Object System.Windows.Forms.Label
    $label9.Location = New-Object System.Drawing.Point(10,250)
    $label9.Size = New-Object System.Drawing.Size(50,20)
    $label9.Text = 'Lied 7'
    $form.Controls.Add($label9)

    $label10 = New-Object System.Windows.Forms.Label
    $label10.Location = New-Object System.Drawing.Point(10,275)
    $label10.Size = New-Object System.Drawing.Size(50,20)
    $label10.Text = 'Extra 1'
    $form.Controls.Add($label10)

    $label11 = New-Object System.Windows.Forms.Label
    $label11.Location = New-Object System.Drawing.Point(10,300)
    $label11.Size = New-Object System.Drawing.Size(50,20)
    $label11.Text = 'Extra 2'
    $form.Controls.Add($label11)

    $label12 = New-Object System.Windows.Forms.Label
    $label12.Location = New-Object System.Drawing.Point(10,325)
    $label12.Size = New-Object System.Drawing.Size(50,20)
    $label12.Text = 'Extra 3'
    $form.Controls.Add($label12)

    $urlKindermoment = New-Object System.Windows.Forms.TextBox
    $urlKindermoment.Location = New-Object System.Drawing.Point(90,47)
    $urlKindermoment.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlKindermoment)

    $urlKinderlied = New-Object System.Windows.Forms.TextBox
    $urlKinderlied.Location = New-Object System.Drawing.Point(90,72)
    $urlKinderlied.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlKinderlied)

    $checkBoxKinderlied = New-Object System.Windows.Forms.CheckBox
    $checkBoxKinderlied.Location = New-Object System.Drawing.Point(300, 72)
    $checkBoxKinderlied.Size = New-Object System.Drawing.Size(20,20)
    $form.Controls.Add($checkBoxKinderlied)

    $urlLied1 = New-Object System.Windows.Forms.TextBox
    $urlLied1.Location = New-Object System.Drawing.Point(90,97)
    $urlLied1.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlLied1)

    $checkBoxlied1 = New-Object System.Windows.Forms.CheckBox
    $checkBoxlied1.Location = New-Object System.Drawing.Point(300, 97)
    $checkBoxlied1.Size = New-Object System.Drawing.Size(20,20)
    $form.Controls.Add($checkBoxlied1)

    $urlLied2 = New-Object System.Windows.Forms.TextBox
    $urlLied2.Location = New-Object System.Drawing.Point(90,122)
    $urlLied2.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlLied2)

    $checkBoxlied2 = New-Object System.Windows.Forms.CheckBox
    $checkBoxlied2.Location = New-Object System.Drawing.Point(300, 122)
    $checkBoxlied2.Size = New-Object System.Drawing.Size(20,20)
    $form.Controls.Add($checkBoxlied2)

    $urlLied3 = New-Object System.Windows.Forms.TextBox
    $urlLied3.Location = New-Object System.Drawing.Point(90,147)
    $urlLied3.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlLied3)

    $checkBoxlied3 = New-Object System.Windows.Forms.CheckBox
    $checkBoxlied3.Location = New-Object System.Drawing.Point(300, 147)
    $checkBoxlied3.Size = New-Object System.Drawing.Size(20,20)
    $form.Controls.Add($checkBoxlied3)

    $urlLied4 = New-Object System.Windows.Forms.TextBox
    $urlLied4.Location = New-Object System.Drawing.Point(90,172)
    $urlLied4.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlLied4)

    $checkBoxlied4 = New-Object System.Windows.Forms.CheckBox
    $checkBoxlied4.Location = New-Object System.Drawing.Point(300, 172)
    $checkBoxlied4.Size = New-Object System.Drawing.Size(20,20)
    $form.Controls.Add($checkBoxlied4)

    $urlLied5 = New-Object System.Windows.Forms.TextBox
    $urlLied5.Location = New-Object System.Drawing.Point(90,197)
    $urlLied5.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlLied5)

    $checkBoxlied5 = New-Object System.Windows.Forms.CheckBox
    $checkBoxlied5.Location = New-Object System.Drawing.Point(300, 197)
    $checkBoxlied5.Size = New-Object System.Drawing.Size(20,20)
    $form.Controls.Add($checkBoxlied5)

    $urlLied6 = New-Object System.Windows.Forms.TextBox
    $urlLied6.Location = New-Object System.Drawing.Point(90,222)
    $urlLied6.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlLied6)

    $checkBoxlied6 = New-Object System.Windows.Forms.CheckBox
    $checkBoxlied6.Location = New-Object System.Drawing.Point(300, 222)
    $checkBoxlied6.Size = New-Object System.Drawing.Size(20,20)
    $form.Controls.Add($checkBoxlied6)

    $urlLied7 = New-Object System.Windows.Forms.TextBox
    $urlLied7.Location = New-Object System.Drawing.Point(90,247)
    $urlLied7.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlLied7)

    $checkBoxlied7 = New-Object System.Windows.Forms.CheckBox
    $checkBoxlied7.Location = New-Object System.Drawing.Point(300, 247)
    $checkBoxlied7.Size = New-Object System.Drawing.Size(20,20)
    $form.Controls.Add($checkBoxlied7)

    $urlExtra1 = New-Object System.Windows.Forms.TextBox
    $urlExtra1.Location = New-Object System.Drawing.Point(90,272)
    $urlExtra1.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlExtra1)

    $urlExtra2 = New-Object System.Windows.Forms.TextBox
    $urlExtra2.Location = New-Object System.Drawing.Point(90,297)
    $urlExtra2.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlExtra2)

    $urlExtra3 = New-Object System.Windows.Forms.TextBox
    $urlExtra3.Location = New-Object System.Drawing.Point(90,322)
    $urlExtra3.Size = New-Object System.Drawing.Size(200,20)
    $form.Controls.Add($urlExtra3)

    $form.Topmost = $true

    $form.Add_Shown({$urlKindermoment.Select()})

    $result = $form.ShowDialog()

    $urlList = [PSCustomObject]@{
        form = $result
        kindermoment = [PSCustomObject]@{
            url = $urlKindermoment.Text
            checked = $false
            }
        kinderlied = [PSCustomObject]@{
            url = $urlKinderlied.Text
            checked = $checkBoxKinderlied.Checked
            }
        lied1 = [PSCustomObject]@{
            url = $urlLied1.Text
            checked = $checkBoxlied1.Checked
            }
        lied2 = [PSCustomObject]@{
            url = $urlLied2.Text
            checked = $checkBoxlied2.Checked
            }
        lied3 = [PSCustomObject]@{
            url = $urlLied3.Text
            checked = $checkBoxlied3.Checked
            }
        lied4 = [PSCustomObject]@{
            url = $urlLied4.Text
            checked = $checkBoxlied4.Checked
            }
        lied5 = [PSCustomObject]@{
            url = $urlLied5.Text
            checked = $checkBoxlied5.Checked
            }
        lied6 = [PSCustomObject]@{
            url = $urlLied6.Text
            checked = $checkBoxlied6.Checked
            }
        lied7 = [PSCustomObject]@{
            url = $urlLied7.Text
            checked = $checkBoxlied7.Checked
            }
        extra1 = [PSCustomObject]@{
            url = $urlExtra1.Text
            checked = $false
            }
        extra2 = [PSCustomObject]@{
            url = $urlExtra2.Text
            checked = $false
            }
        extra3 = [PSCustomObject]@{
            url = $urlExtra3.Text
            checked = $false
            }
    }

    Write-Output $urlList
}

#empty folder 
Get-ChildItem -Path ([Environment]::GetFolderPath("MyVideos") + "\autogendienst") | foreach { $_.Delete()}

#show form and get results
$results = Show-Form

#if OK button pressed, for each string that is not empty run download function
if ($results.form -eq [System.Windows.Forms.DialogResult]::OK)
{
    $results.PSObject.Properties | ForEach-Object {
    if(-not [string]::IsNullOrEmpty($_.Value.url))
        {
            if($_.Value.checked){
                Set-Companion-Button-Name -video_name $_.Name -title $_.Value.url -checked $true
            }
            else{
                Download-Video -url $_.Value.url -video_name $_.Name
            }
        }
    }
}