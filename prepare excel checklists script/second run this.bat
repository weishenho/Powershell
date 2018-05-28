@ECHO OFF
powershell.exe -NoProfile -ExecutionPolicy Bypass -File edit_checklist.ps1
timeout 1 > nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File edit_2_checklist.ps1
PAUSE