# coding: utf8

import datetime
import ctypes

import win32com.client


class sap:
  def __init__(self):
    print("Inicializando o módulo de automação do SAP Frontend...")
    try:
      self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine  
      self.session = self.SapGui.FindById("ses[0]")
    except:
      raise Exception("O módulo do SAP Frontend não pode ser iniciado!\nVerifique se o SAP está aberto ou se há uma seção ativa!")
  def relatorio(self, dia=7):
    try:
      self.session.StartTransaction(Transaction="ZSVC20")
      self.session.FindById("wnd[0]/usr/btn%_SO_QMART_%_APP_%-VALU_PUSH").Press()
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "B1"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "BL"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "BR"
      self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
      hoje = datetime.date.today()
      semana = hoje - datetime.timedelta(days=dia)
      self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-LOW").text = semana.strftime("%d.%m.%Y")
      self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-HIGH").text = hoje.strftime("%d.%m.%Y")
      self.session.FindById("wnd[0]/usr/btn%_SO_USUAR_%_APP_%-VALU_PUSH").Press()
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "ENVI"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "LIBE"
      self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
      self.session.FindById("wnd[0]/usr/ctxtSO_BEBER-LOW").text = "RB"
      self.session.FindById("wnd[0]/usr/ctxtP_LAYOUT").text = "/MANSERVRELC"
      print("Aguarde relatório sendo processado...")
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      print("Relatório processado com sucesso!")
      ctypes.windll.user32.FlashWindow(ctypes.windll.kernel32.GetConsoleWindow(), True )
    except:
      print("Verifique se o SAP está aberto ou se há uma seção ativa!")
  def leiturista(self, nota):
    try:
      self.session.StartTransaction(Transaction="IW53")
      self.session.FindById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = nota
      self.session.FindById("wnd[0]/tbar[1]/btn[5]").Press()
      self.session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09").Press()
      instalacao = self.session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09/ssubSUB_GROUP_10:SAPLIQS0:7217/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/ctxtVIQMEL-ZZINSTLN").text
      print(instalacao)
      self.session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09/ssubSUB_GROUP_10:SAPLIQS0:7217/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/ctxtVIQMEL-ZZINSTLN").Press()
      unidade = self.session.FindById("wnd[0]/usr/tblSAPLES30TC_TIMESL/ctxtEANLD-ABLEINH[9,0]").text
      print(unidade)
      self.session.StartTransaction(Transaction="ZMED89")
      self.session.FindById("wnd[0]/usr/txtP_ABL_Z-LOW").text = "015"
      self.session.FindById("wnd[0]/usr/ctxtP_LOTE-LOW").text = "03"
      self.session.FindById("wnd[0]/usr/ctxtP_MESANO").text = "04/2022"
      self.session.FindById("wnd[0]/usr/ctxtP_UNID-LOW").text = "03L61414"
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").press
    except:
      print("Verifique se o SAP está aberto ou se há uma seção ativa!")
  def debito(self, nota):
    pass