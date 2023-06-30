REM @echo OFF
REM cls
venv\Scripts\Activate
pyinstaller --onefile .\\src\\sap.py
pyinstaller --onefile .\\src\\img.py
cd ..
copy .\Automacao\src\fileDialog.vbs .\telbot
copy .\Automacao\dist\* .\telbot
cd .\telbot
dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained true
cd ..
copy .\telbot\bin\Debug\net6.0\win-x64\publish\* %USERPROFILE%\MestreRuan\
cd .\Automacao
venv\Scripts\deactivate.bat
pause
