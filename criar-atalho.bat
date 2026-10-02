@echo off
REM Recria o atalho do Pooling SURF na area de trabalho (abre como janela de app no Chrome).
setlocal
set APP=%~dp0index.html
set ICO=%~dp0icon.ico
set CHROME=
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe
if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe
if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" set CHROME=%LocalAppData%\Google\Chrome\Application\chrome.exe

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "$d=[Environment]::GetFolderPath('Desktop');" ^
 "$s=(New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $d 'Pooling SURF.lnk'));" ^
 "if('%CHROME%' -ne ''){ $s.TargetPath='%CHROME%'; $s.Arguments='--app=file:///' + ('%APP%' -replace '\\','/') } else { $s.TargetPath='%APP%' };" ^
 "$s.IconLocation='%ICO%,0'; $s.WorkingDirectory='%~dp0'; $s.Description='Pooling SURF - ISPX'; $s.Save();" ^
 "Write-Host 'Atalho criado em' $d"

echo.
pause
