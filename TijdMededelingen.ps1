$WshShell = New-Object -ComObject wscript.shell
$Time = (Get-Date).hour

if ($Time -gt 12)
        {
        Set-Content -Path 'C:\Users\rehob\Dropbox\Easy Worship\Andere bestanden\DienstTijden\aanvang.txt' -Value '16:30u'
        Set-Content -Path 'C:\Users\rehob\Dropbox\Easy Worship\Andere bestanden\DienstTijden\dagdeel.txt' -Value 'volgende week'
        Set-Content -Path 'C:\Users\rehob\Dropbox\Easy Worship\Andere bestanden\DienstTijden\volgendedienst.txt' -Value '9:30u'
        }
else 
        {
        Set-Content -Path 'C:\Users\rehob\Dropbox\Easy Worship\Andere bestanden\DienstTijden\aanvang.txt' -Value '9:30u'
        Set-Content -Path 'C:\Users\rehob\Dropbox\Easy Worship\Andere bestanden\DienstTijden\dagdeel.txt' -Value 'vanmiddag'
        Set-Content -Path 'C:\Users\rehob\Dropbox\Easy Worship\Andere bestanden\DienstTijden\volgendedienst.txt' -Value '16:30u'
}

Set-Content -Path 'C:\Users\rehob\Desktop\Extra Mededeling.txt' -Value ''
            