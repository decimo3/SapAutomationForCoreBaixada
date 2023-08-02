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
robocopy tmp %USERPROFILE%\MestreRuan\ /XF database.db
REM https://superuser.com/questions/1654994/how-to-copy-folder-structure-but-exclude-certain-files-in-windows
cd .\SapAutomationForCoreBaixada
venv\Scripts\deactivate.bat
pause
