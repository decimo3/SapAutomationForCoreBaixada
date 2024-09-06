@echo off
call build.bat
cd ..
robocopy tmp\net7.0\win-x64\publish tmp\ /S
robocopy tmp %USERPROFILE%\MestreRuan\ /S /XD net7.0 /XF database.db /XF .env /XF sap.conf /XF ofs.conf
cd .\SapAutomationForCoreBaixada
