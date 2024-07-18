@echo OFF
if not defined VIRTUAL_ENV (
  call .venv\Scripts\activate
)
pyinstaller --onefile --icon appicon.ico .\\src\\sap.py
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
for /f "tokens=2 delims==" %%a in ('wmic os get localdatetime /value') do set datetime=%%a
set "datestamp=%datetime:~0,4%%datetime:~4,2%%datetime:~6,2%"
echo %datestamp% > version
dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained true --output ..\tmp\
cd ..
cd .\monitoring-fieldteam\src
dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained true --output ..\..\tmp\
cd ..
cd ..
cd .\SapAutomationForWeb
dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained true --output ..\tmp\
cd ..
cd .\SapAutomationForCoreBaixada
