$code = @"
    [DllImport("user32.dll")]
    public static extern bool BlockInput(bool fBlockIt);
"@

$userInput = Add-Type -MemberDefinition $code -Name UserInput -Namespace UserInput -PassThru

function Disable-UserInput($seconds) {
    $userInput::BlockInput($true)
    Start-Sleep $seconds
    $userInput::BlockInput($false)
}


$wshell = New-Object -ComObject wscript.shell;
#param([string] $proc="obs64", [string]$adm)
$proc = "obs64"
cls

Add-Type @"
  using System;
  using System.Runtime.InteropServices;
  public class WinAp {
     [DllImport("user32.dll")]
     [return: MarshalAs(UnmanagedType.Bool)]
     public static extern bool SetForegroundWindow(IntPtr hWnd);

     [DllImport("user32.dll")]
     [return: MarshalAs(UnmanagedType.Bool)]
     public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
  }

"@
$p = Get-Process |where {$_.mainWindowTItle }|where {$_.Name -like "$proc"}

if (($p -eq $null) -and ($adm -ne ""))
{
    Start-Process "$proc" -Verb runAs
}
elseif (($p -eq $null) -and ($adm -eq ""))
{
    Start-Process "$proc" #-Verb runAs
}
else
{
    $h = $p.MainWindowHandle

    [void] [WinAp]::SetForegroundWindow($h)
    [void] [WinAp]::ShowWindow($h,3);
}

#|format-table id,name,mainwindowtitle â€“AutoSize
# static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
# powershell.exe -windowstyle hidden -file *.ps1 -adm "a"

Disable-UserInput -second 1 | Out-Null 
$wshell.SendKeys('%f')
$wshell.SendKeys('s')
Disable-UserInput -second 1 | Out-Null 
$wshell.SendKeys('s')
Disable-UserInput -second 1 | Out-Null 
$wshell.SendKeys('{TAB}')
Disable-UserInput -second 1 | Out-Null 
$wshell.SendKeys('{TAB}')
Disable-UserInput -second 1 | Out-Null 
$wshell.SendKeys('{TAB}')
Disable-UserInput -second 1 | Out-Null 
$wshell.SendKeys('') #verander hier de YouTube key wanneer nodig.
Disable-UserInput -second 1 | Out-Null 
$wshell.SendKeys('{TAB}')
Disable-UserInput -second 1 | Out-Null 
$wshell.SendKeys('{ENTER}') 
Disable-UserInput -second 1 | Out-Null 
$wshell.SendKeys('{TAB}')
Disable-UserInput -second 1 | Out-Null 
$wshell.SendKeys('{TAB}')
Disable-UserInput -second 1 | Out-Null 
$wshell.SendKeys('{ENTER}') 
start ‘https://studio.youtube.com/channel/UCXVY17JezGcpyD1xHzLmaVg/livestreaming/dashboard’
Disable-UserInput -second 5 | Out-Null 