# Optional: start Land Bot automatically when you log in to Windows (hidden window).
#   powershell -ExecutionPolicy Bypass -File .\scripts\autostart-on.ps1
# To turn it off again:
#   Unregister-ScheduledTask -TaskName "Probridge Land Bot" -Confirm:$false

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$root\start.ps1`"" `
    -WorkingDirectory $root
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0
Register-ScheduledTask -TaskName "Probridge Land Bot" -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
Write-Host "Land Bot will now start automatically when you log in." -ForegroundColor Green
