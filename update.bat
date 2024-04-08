@echo off
call build.bat
cd ..
robocopy tmp %USERPROFILE%\MestreRuan\ /XF database.db /XF .env /XF sap.conf
cd .\SapAutomationForCoreBaixada
