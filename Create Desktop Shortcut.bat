@echo off
REM One-time helper: put a "Jasco Order Picker" shortcut on the Windows Desktop.
REM
REM Double-click this once after setup. It creates a Desktop shortcut that points
REM back to "Order Picker.bat" in this repo folder (with the repo as its working
REM directory), so the owner can launch the picker straight from the Desktop.
REM Safe to run more than once -- it just overwrites the existing shortcut.

cd /d "%~dp0"

REM GetFolderPath('Desktop') respects a OneDrive-redirected Desktop, unlike a hard
REM-coded %USERPROFILE%\Desktop. The shortcut targets the .bat and runs it from here.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$d=[Environment]::GetFolderPath('Desktop'); $lnk=Join-Path $d 'Jasco Order Picker.lnk'; $w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut($lnk); $s.TargetPath='%~dp0Order Picker.bat'; $s.WorkingDirectory='%CD%'; $s.Description='Launch the Jasco order picker'; $s.Save(); Write-Host ('Created: ' + $lnk)"

if errorlevel 1 (
    echo.
    echo Could not create the shortcut automatically. You can make one by hand:
    echo   right-click "Order Picker.bat" -^> Send to -^> Desktop ^(create shortcut^).
    pause
) else (
    echo.
    echo Done -- look for "Jasco Order Picker" on your Desktop.
    pause
)
