@echo off
git diff --exit-code comunicado.txt > nul
if %ERRORLEVEL% equ 0 (
  echo Updating the 'comunicado.txt' file is necessary to publish
  exit
)
set "UPDATE_FOLDER=\\192.168.10.213\chatbot\"
if not exist %UPDATE_FOLDER% (
  echo The shared folder path '%UPDATE_FOLDER%' was not found!
  exit
)
call build.bat
cd ..
for /f "tokens=2 delims==" %%a in ('wmic os get localdatetime /value') do set datetime=%%a
set "datestamp=%datetime:~0,4%%datetime:~4,2%%datetime:~6,2%"
echo %datestamp% > ".\tmp\version"
7z u -x!database.db -x!sap.conf -x!.env %UPDATE_FOLDER%%datestamp%.zip .\tmp\*
cd .\SapAutomationForCoreBaixada
git add comunicado.txt
git commit -m "Update 'comunicado.txt' to version %datestamp%"