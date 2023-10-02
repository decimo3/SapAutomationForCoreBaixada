@echo off
setlocal enabledelayedexpansion
rem build a new executable
pyinstaller --onefile .\\src\sap.py
rem clear the screen before next prompt
cls
rem get the current directory where this batch script is located
set "current_directory=%~dp0"
rem activate the virtual environment
rem call "%current_directory%\venv\Scripts\activate"
rem define an array of integers
set "argumentos=1376992838 1376992952 1376992860 1376992570 1376993162 1376992628 1376992801 1376993215 1376993219 1376992424 1376992835 1376992676 1376993242 1376992695 1376991969 1376992955 1376992542 1376992927 1376991914 1376992857 1376992959 1376993104 1376992240 1376991916 1376992254 1376992328 1376992059 1376993241 1376992886 1376992888 1376992765 1376993169 1376992110 1376992038 1376992129"
rem define an array os strings
set "aplicacoes=pendente testando testando2"
rem path to your python script
set "python_script=.\\src\\sap.py"
rem define the log file path
set "log_file=execution_log.log"
rem loop through the array of integers
for %%i in (%argumentos%) do (
  for %%j in (%aplicacoes%) do (
    echo aplicacao %%j argumento %%i
    python "%python_script%" %%j %%i
  )
  pause
)
rem deactivate the virtual environment
rem deactivate
endlocal
