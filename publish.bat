@echo off
setlocal

set "ROOT_DIR=%~dp0"
set "PROJECT_FILE=%ROOT_DIR%BillMatch.Wpf\BillMatch.Wpf.csproj"
set "PUBLISH_EXE=%ROOT_DIR%BillMatch.Wpf\bin\Release\net8.0-windows\win-x64\publish\BillMatch.Wpf.exe"
set "TARGET_EXE=%ROOT_DIR%BillMatch.exe"

echo Publishing BillMatch.Wpf...
dotnet publish "%PROJECT_FILE%" -c Release -r win-x64 --no-self-contained -p:PublishSingleFile=true
if errorlevel 1 goto :publish_failed

if not exist "%PUBLISH_EXE%" goto :missing_output

copy /Y "%PUBLISH_EXE%" "%TARGET_EXE%" >nul
if errorlevel 1 goto :copy_failed

echo Done. Copied to %TARGET_EXE%
pause
exit /b 0

:publish_failed
echo Publish failed.
pause
exit /b 1

:missing_output
echo Publish output not found: %PUBLISH_EXE%
pause
exit /b 1

:copy_failed
echo Copy failed.
pause
exit /b 1
