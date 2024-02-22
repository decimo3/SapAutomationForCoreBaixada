@echo OFF
if defined VIRTUAL_ENV (
pyinstaller --onefile .\\src\\sap.py
pyinstaller --onefile .\\src\\img.py
pyinstaller --onefile .\\src\\etc.py
rm .\\src\\sap.db
sqlite3 .\\src\\sap.db < .\\src\\sap.sql
cd ..
mkdir tmp
copy .\SapAutomationForCoreBaixada\src\sap.db tmp
copy .\SapAutomationForCoreBaixada\src\sap.conf tmp
copy .\SapAutomationForCoreBaixada\src\erroDialog.vbs tmp
copy .\SapAutomationForCoreBaixada\src\fileDialog.vbs tmp
copy .\SapAutomationForCoreBaixada\dist\* tmp
cd .\TelegramBotForFieldTeamHelper
dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained true --output ..\tmp\
cd ..
robocopy tmp %USERPROFILE%\MestreRuan\ /XF database.db /XF .env
7z u -x!database.db -x!sap.conf -x!.env mestreruan.zip .\tmp\*
REM https://superuser.com/questions/1654994/how-to-copy-folder-structure-but-exclude-certain-files-in-windows
cd .\SapAutomationForCoreBaixada
) else (
echo "Environment variable VIRTUAL_ENV was not defined"
)
