@echo off
chcp 65001 >nul
REM 평일(월~금) 아침에 teams_alarm.py 를 자동 실행하도록 Windows 작업 스케줄러에 등록합니다.
REM 사용법: setup_schedule.bat            -> 평일 09:00
REM         setup_schedule.bat 08:30      -> 평일 08:30
REM 해제:   schtasks /delete /tn TeamsAlarmGAGC /f

set RUN_AT=%1
if "%RUN_AT%"=="" set RUN_AT=09:00

powershell -NoProfile -Command ^
  "$a = New-ScheduledTaskAction -Execute 'py' -Argument 'teams_alarm.py' -WorkingDirectory '%~dp0';" ^
  "$t = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At '%RUN_AT%';" ^
  "$s = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries;" ^
  "Register-ScheduledTask -TaskName 'TeamsAlarmGAGC' -Action $a -Trigger $t -Settings $s -Force | Out-Null;" ^
  "Write-Host '등록 완료: 평일 %RUN_AT% 에 실행됩니다.'"

pause
