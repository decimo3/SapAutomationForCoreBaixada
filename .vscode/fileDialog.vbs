
FileName =  "testando"&".pdf"

Set Wshell = CreateObject("WScript.Shell")
Do 
  bWindowFound = Wshell.AppActivate("Save As") 
  wscript.sleep 1000
  cWindowFound = Wshell.AppActivate("Import file") 
  wscript.sleep 1000
Loop Until bWindowFound or cWindowFound

' and probably the least elegant solution around - using tab sendkeys to access the necessary input fields. 
' the number of tabs depends on what you want to access - might be different for you. Trial and error are recommended ;)

if (bWindowFound) Then

Wshell.appActivate "Save As"
Wshell.SendKeys "{TAB}"
WScript.Sleep 100
Wshell.SendKeys "{TAB}"
WScript.Sleep 100
Wshell.SendKeys "{TAB}"
WScript.Sleep 100
Wshell.SendKeys "{TAB}"
WScript.Sleep 100
Wshell.SendKeys "{TAB}"
WScript.Sleep 100
Wshell.SendKeys FileName
WScript.Sleep 100
Wshell.SendKeys "{ENTER}"
end if

if (cWindowFound) Then

Wshell.appActivate "Import file"
Wshell.SendKeys "{TAB}"
WScript.Sleep 100
Wshell.SendKeys "{TAB}"
WScript.Sleep 100
Wshell.SendKeys "{TAB}"
WScript.Sleep 100
Wshell.SendKeys FileName
WScript.Sleep 1000
Wshell.SendKeys "{ENTER}"
end if