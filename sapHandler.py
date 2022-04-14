# coding: utf8

import datetime
import ctypes

import win32com.client


class sap:
  def relatorio(dia=7):
    # TODO: adicionar a verificação do tipo do argumento
    try:
      SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
      session = SapGui.FindById("ses[0]")
      session.StartTransaction(Transaction="ZSVC20")
      session.FindById("wnd[0]/usr/btn%_SO_QMART_%_APP_%-VALU_PUSH").Press()
      session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "B1"
      session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "BL"
      session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "BR"
      session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
      hoje = datetime.date.today()
      semana = hoje - datetime.timedelta(days=dia)
      session.FindById("wnd[0]/usr/ctxtSO_QMDAT-LOW").text = semana.strftime("%d.%m.%Y")
      session.FindById("wnd[0]/usr/ctxtSO_QMDAT-HIGH").text = hoje.strftime("%d.%m.%Y")
      session.FindById("wnd[0]/usr/btn%_SO_USUAR_%_APP_%-VALU_PUSH").Press()
      session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "ENVI"
      session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "LIBE"
      session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
      session.FindById("wnd[0]/usr/ctxtSO_BEBER-LOW").text = "RB"
      session.FindById("wnd[0]/usr/ctxtP_LAYOUT").text = "/MANSERVRELC"
      print("Aguarde relatório sendo processado...")
      session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      print("Relatório processado com sucesso!")
      ctypes.windll.user32.FlashWindow(ctypes.windll.kernel32.GetConsoleWindow(), True )
    except:
      print("Verifique se o SAP está aberto ou se há uma seção ativa!")
    def leiturista(nota):
      pass
    def debito(nota):
      pass