import win32com.client
sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
sap = sap.FindById("ses[0]")
