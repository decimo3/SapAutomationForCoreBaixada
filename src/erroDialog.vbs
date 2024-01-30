Set Wshell = CreateObject("WScript.Shell")
Dim title, window
title = "SAP GUI for Windows 730"
Do 
  window = Wshell.AppActivate(title)
  wscript.sleep 1000
  if (window) Then
    Wshell.appActivate title
    Wshell.SendKeys "{ESC}"
    WScript.Sleep 100
  end if
Loop

