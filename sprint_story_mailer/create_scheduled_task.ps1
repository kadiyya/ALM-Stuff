# Windows Gorev Zamanlayici'ya her gun 16:00'da calisan gorevi ekler.
# Script sadece sprint bitis gununde (veya kacirildiysa sonraki gunlerde) mail atar.
# PC 16:00'da kapaliysa, acildiginda calisir (StartWhenAvailable).
$py = (Get-Command pythonw -ErrorAction Stop).Source
$script = Join-Path $PSScriptRoot "sprint_story_mailer.py"
$action = New-ScheduledTaskAction -Execute $py -Argument "`"$script`"" -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Daily -At 16:00
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 4) `
            -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName "RTC Sprint Story Mailer" -Action $action -Trigger $trigger `
                       -Settings $settings -Force | Out-Null
Write-Host "Gorev olusturuldu: 'RTC Sprint Story Mailer' (her gun 16:00)."
Write-Host "Elle test: Start-ScheduledTask -TaskName 'RTC Sprint Story Mailer'"
