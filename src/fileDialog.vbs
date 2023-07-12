Set Wshell = CreateObject("WScript.Shell")

Do 
  bWindowFound = Wshell.AppActivate("Salvar como") 
  wscript.sleep 10
  cWindowFound = Wshell.AppActivate("Importar arquivo") 
  wscript.sleep 10
Loop Until bWindowFound or cWindowFound

' and probably the least elegant solution around - using tab sendkeys to access the necessary input fields. 
' the number of tabs depends on what you want to access - might be different for you. Trial and error are recommended ;)

if (bWindowFound) Then
  Wshell.appActivate "Salvar como"
  Wshell.SendKeys "{TAB}"
  WScript.Sleep 10
  Wshell.SendKeys "{TAB}"
  WScript.Sleep 10
  Wshell.SendKeys "{TAB}"
  WScript.Sleep 10
  Wshell.SendKeys "{TAB}"
  WScript.Sleep 10
  Wshell.SendKeys "{TAB}"
  WScript.Sleep 10
  Wshell.SendKeys "{ENTER}"
  WScript.Sleep 10
  bWindowFound = Wshell.AppActivate("Salvar como")
  if (bWindowFound) Then
    WScript.Sleep 10
    Wshell.SendKeys "+{TAB}"
    WScript.Sleep 10
    Wshell.SendKeys " "
  end if
end if

if (cWindowFound) Then
  Wshell.appActivate "Importar arquivo"
  Wshell.SendKeys "{TAB}"
  WScript.Sleep 10
  Wshell.SendKeys "{TAB}"
  WScript.Sleep 10
  Wshell.SendKeys "{TAB}"
  WScript.Sleep 10
  Wshell.SendKeys "{ENTER}"
end if