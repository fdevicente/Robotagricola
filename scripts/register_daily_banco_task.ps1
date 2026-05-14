# Registra Task Scheduler: corre daily_banco_18h.py todos los dias a las 18:00.
# Uso: ejecutar como Administrador.

$TaskName = "AgricolaSantaElisa-DailyBanco-18h"
$Python = "py.exe"
$Args = "-3.11 $PSScriptRoot\..\daily_banco_18h.py"
$WorkingDir = "$PSScriptRoot\.."

$Action = New-ScheduledTaskAction -Execute $Python -Argument $Args -WorkingDirectory $WorkingDir
$Trigger = New-ScheduledTaskTrigger -Daily -At 6:00pm
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 60)

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger `
    -Principal $Principal -Settings $Settings -Force

Write-Host "Tarea registrada: $TaskName (corre cada dia a las 18:00)"
