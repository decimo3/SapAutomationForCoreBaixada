#!/usr/bin/python
# coding: utf8
import os
import time
import dotenv
import subprocess
import win32com.client
if __name__ == "__main__":
  dotenv_path = os.path.dirname(__file__)
  dotenv_file = os.path.join(dotenv_path, 'sap.conf')
  dotenv.load_dotenv(dotenv_file)
  # Get scripting
  try:
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
  except:
    saplogon = "C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"
    subprocess.Popen(saplogon)
    time.sleep(3)
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
  if not type(SapGuiAuto) == win32com.client.CDispatch:
      raise Exception("ERRO: SAP GUI Scripting API is not available.")
  application = SapGuiAuto.GetScriptingEngine
  # Get connection
  if not (len(application.connections) > 0):
    try:
      connection = application.OpenConnection("#PCL", True)
      print(connection.connectionString)
    except:
      raise Exception("ERRO: SAP FrontEnd connection is not available.")
  else:
    connection = application.connections[0]
  # Get session
  session = connection.Children(0)
  session.findById("wnd[0]/usr/txtRSYST-BNAME").text = os.environ.get("USUARIO")
  session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = os.environ.get("PALAVRA")
  session.findById("wnd[0]/tbar[0]/btn[0]").Press()
  # /app/con[0]/ses[0]/wnd[0]/usr/txtRSYST-BNAME