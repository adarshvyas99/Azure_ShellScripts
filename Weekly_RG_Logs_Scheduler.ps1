$username = "$env:USERNAME"
$password = "Iamback"
$Action = New-ScheduledTaskAction -Execute 'C:\Program Files\PowerShell\7\pwsh.exe' -Argument '-NonInteractive -NoLogo -NoProfile -File "C:\Users\Adarsh\rg_log_schedule.ps1"'
$Trigger = New-ScheduledTaskTrigger -Once -At "08:57pm" <# -Weekly -WeeksInterval -DaysOfWeek Monday -At 8am #>
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable
$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings
Register-ScheduledTask -TaskName 'Automation script for RG logs' -InputObject $Task -User $username -Password $password