@echo off
call build.bat
cd ..
robocopy tmp %USERPROFILE%\MestreRuan\ /XF database.db /XF .env /XF sap.conf /XF ofs.conf
cd .\SapAutomationForCoreBaixada
