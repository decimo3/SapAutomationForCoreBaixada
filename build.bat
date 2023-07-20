REM @echo OFF
REM cls
venv\Scripts\Activate
pyinstaller --onefile .\\src\\sap.py
pyinstaller --onefile .\\src\\img.py
cd ..
mkdir tmp
copy .\SapAutomationForCoreBaixada\src\fileDialog.vbs tmp
copy .\SapAutomationForCoreBaixada\dist\* tmp
cd .\TelegramBotForFieldTeamHelper
dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained true --output ..\tmp\
cd ..
copy .\tmp\* %USERPROFILE%\MestreRuan\
cd .\SapAutomationForCoreBaixada
venv\Scripts\deactivate.bat
pause
