@echo OFF
if defined VIRTUAL_ENV (
git diff --exit-code comunicado.txt > nul
if %ERRORLEVEL% equ 1 (
set UPDATE_FOLDER="\\srv-isjpa\chatbot\"
if exist "%UPDATE_FOLDER%" (
for /f "tokens=2 delims==" %%a in ('wmic os get localdatetime /value') do set datetime=%%a
set "datestamp=%datetime:~0,4%%datetime:~4,2%%datetime:~6,2%"
pyinstaller --onefile .\\src\\sap.py
pyinstaller --onefile .\\src\\img.py
del .\\src\\sap.db
sqlite3 .\\src\\sap.db < .\\src\\sap.sql
cd ..
mkdir tmp
copy .\SapAutomationForCoreBaixada\src\sap.db tmp
copy .\SapAutomationForCoreBaixada\src\sap.conf tmp
copy .\SapAutomationForCoreBaixada\dist\sap.exe tmp
copy .\SapAutomationForCoreBaixada\dist\img.exe tmp
copy .\SapAutomationForCoreBaixada\comunicado.txt tmp
copy .\SapAutomationForCoreBaixada\src\erroDialog.vbs tmp
copy .\SapAutomationForCoreBaixada\src\fileDialog.vbs tmp
cd .\TelegramBotForFieldTeamHelper
echo %datestamp% > ".\version"
dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained true --output ..\tmp\
cd ..
7z u -x!database.db -x!sap.conf -x!.env %UPDATE_FOLDER%%datestamp%.zip .\tmp\*
REM https://superuser.com/questions/1654994/how-to-copy-folder-structure-but-exclude-certain-files-in-windows
rem robocopy tmp %USERPROFILE%\MestreRuan\ /XF database.db /XF .env /XF sap.conf
cd .\SapAutomationForCoreBaixada
git add comunicado.txt
git commit -m "Update 'comunicado.txt' to version %datestamp%"
) else (
  echo The shared folder path '%UPDATE_FOLDER%' was not found!
)
) else (
  echo Updating the 'comunicado.txt' file is necessary to build
)
) else (
echo Environment variable 'VIRTUAL_ENV' was not defined
)
