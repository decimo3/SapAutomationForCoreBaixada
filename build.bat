@echo OFF
if not defined VIRTUAL_ENV (
  echo Environment variable 'VIRTUAL_ENV' was not defined
  exit
)
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
dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained true --output ..\tmp\
cd ..
cd .\SapAutomationForCoreBaixada
