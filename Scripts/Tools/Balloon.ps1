Add-Type -AssemblyName System.Windows.Forms
$global:balloon = New-Object System.Windows.Forms.NotifyIcon
$path = (Get-Process -id $pid).Path
$balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\explorer.exe")
$balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
$balloon.BalloonTipText = "Clic to open"
$balloon.BalloonTipTitle = "Exported File"
$balloon.Visible = $true
Register-ObjectEvent -InputObject $balloon -EventName BalloonTipClicked -Action{Start-Process explorer.exe} | Out-Null
$balloon.ShowBalloonTip(1000)