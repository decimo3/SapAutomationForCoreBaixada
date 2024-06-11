@echo off
git diff --exit-code comunicado.txt > nul
if %ERRORLEVEL% equ 0 (
  echo Updating the 'comunicado.txt' file is necessary to publish
  exit /B
)
set "UPDATE_FOLDER=\\192.168.10.213\chatbot\"
if not exist %UPDATE_FOLDER% (
  echo The shared folder path '%UPDATE_FOLDER%' was not found!
  exit /B
)
call build.bat
cd ..
for /f "tokens=2 delims==" %%a in ('wmic os get localdatetime /value') do set datetime=%%a
set "datestamp=%datetime:~0,4%%datetime:~4,2%%datetime:~6,2%"
7z u -x!database.db -x!.env -x!sap.conf -x!ofs.conf .\releases\%datestamp%.zip .\tmp\*
copy .\releases\%datestamp%.zip %UPDATE_FOLDER%\%datestamp%.zip
cd .\SapAutomationForCoreBaixada
