@echo OFF
if defined VIRTUAL_ENV (
git diff --exit-code comunicado.txt > nul
if %errorlevel% equ 1 (
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
copy .\SapAutomationForCoreBaixada\src\erroDialog.vbs tmp
copy .\SapAutomationForCoreBaixada\src\fileDialog.vbs tmp
cd .\TelegramBotForFieldTeamHelper
dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained true --output ..\tmp\
cd ..
7z u -x!database.db -x!sap.conf -x!.env mestreruan.zip .\tmp\*
REM https://superuser.com/questions/1654994/how-to-copy-folder-structure-but-exclude-certain-files-in-windows
robocopy tmp %USERPROFILE%\MestreRuan\ /XF database.db /XF .env /XF sap.conf
cd .\SapAutomationForCoreBaixada
) else (
  echo Updating the 'comunicado.txt' file is necessary to build
)
) else (
echo Environment variable VIRTUAL_ENV was not defined
)
