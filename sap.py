# coding: utf8

import datetime
import ctypes

import win32com.client


class sap:
  def __init__(self):
      self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
      # print(dir(self.SapGui)) ['AddHistoryEntry', 'CreateGuiCollection', 'DropHistory', 'FindById', 'GetScriptingEngine', 'Ignore', 'OpenConnection', 'OpenConnectionByConnectionString', 'OpenWDConnection', 'Quit', 'RegisterROT', 'RevokeROT']
      self.session = self.SapGui.FindById("ses[0]")
      # print(dir(self.session)) ['AsStdNumberFormat', 'ClearErrorList', 'CreateSession', 'EnableJawsEvents', 'EndTransaction', 'FindById', 'FindByPosition', 'GetIconResourceName', 'GetVKeyDescription', 'LockSessionUI', 'RunScriptControl', 'SendCommand', 'SendCommandAsync', 'SendMenu', 'StartTransaction', 'UnlockSessionUI']
  def relatorio(self, dia=7):
    try:
      self.session.StartTransaction(Transaction="ZSVC20")
      # print(dir(self.session.FindById("wnd[0]/usr/btn%_SO_QMART_%_APP_%-VALU_PUSH"))) ['DumpState', 'Press', 'SetFocus', 'ShowContextMenu', 'Visualize']
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
      start_time = datetime.datetime.now()
      print("Aguarde relatório sendo processado...")
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      print("Relatório processado com sucesso!")
      end_time = datetime.datetime.now()
      print(f"Relatório gerado em {end_time - start_time}")
      ctypes.windll.user32.FlashWindow(ctypes.windll.kernel32.GetConsoleWindow(), True )
    except:
      print("Verifique se o SAP está aberto ou se há uma seção ativa!")
  def leiturista(self, nota):
    try:
      instalacao = self.instalacao(nota)
      self.session.StartTransaction(Transaction="ES32")
      self.session.FindById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = instalacao
      self.session.findById("wnd[0]/tbar[0]/btn[0]").Press()
      unidade = self.session.FindById("wnd[0]/usr/tblSAPLES30TC_TIMESL/ctxtEANLD-ABLEINH[9,0]").text
      print(unidade)
      self.session.StartTransaction(Transaction="ZMED89")
      livro = f"{unidade[0]}{unidade[1]}"
      local = f"{unidade[2]}{unidade[3]}{unidade[4]}{unidade[5]}"
      if (local == "L645"): centro = "017"
      elif (local == "L644"): centro = "017"
      elif (local == "L643"): centro = "017"
      elif (local == "L624"): centro = "017"
      elif (local == "L622"): centro = "017"
      elif (local == "L613"): centro = "017"
      elif (local == "L616"): centro = "015"
      elif (local == "L615"): centro = "015"
      elif (local == "L614"): centro = "015"
      elif (local == "L612"): centro = "015"
      elif (local == "L610"): centro = "014"
      elif (local == "L623"): centro = "014"
      elif (local == "L611"): centro = "014"
      elif (local == "L620"): centro = "015"
      elif (local == "L617"): centro = "015"
      elif (local == "L625"): centro = "013"
      elif (local == "L635"): centro = "013"
      elif (local == "L636"): centro = "013"
      elif (local == "L637"): centro = "013"
      elif (local == "L630"): centro = "012"
      elif (local == "L632"): centro = "012"
      else: raise ValueError("A localidade pesquisada é desconhecida")
      mes = datetime.date.today()
      mes = mes.replace(day=1)
      mes = mes - datetime.timedelta(days=1)
      print(mes.strftime("%m/%Y"))
      self.session.FindById("wnd[0]/usr/txtP_ABL_Z-LOW").text = centro
      self.session.FindById("wnd[0]/usr/ctxtP_LOTE-LOW").text = livro
      self.session.FindById("wnd[0]/usr/ctxtP_MESANO").text = mes.strftime("%m/%Y")
      self.session.FindById("wnd[0]/usr/ctxtP_UNID-LOW").text = unidade
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      self.session.FindById("wnd[0]/tbar[1]/btn[33]").Press()
      self.session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(201,"DEFAULT")
      self.session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
      self.session.FindById("wnd[0]/tbar[0]/btn[71]").Press()
      self.session.FindById("wnd[1]/usr/txtGS_SEARCH-VALUE").text = instalacao
      self.session.FindById("wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
      self.session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
      self.session.FindById("wnd[1]").Close()
      celula = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow
      if (celula > 28):
        self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = celula -14
      self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").setColumnWidth("ZZ_NUMSEQ",5)
      self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").setColumnWidth("ZHORALEIT",7)
      self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").setColumnWidth("GERAET",8)
      self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").setColumnWidth("ZENDERECO",65)
      self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = celula
    except:
      print("Verifique se o SAP está aberto ou se há uma seção ativa!")
  def debito(self, nota):
      instalacao = self.instalacao(nota)
      self.session.StartTransaction(Transaction="ZARC140")
      self.session.FindById("wnd[0]/usr/ctxtP_ANLAGE").text = instalacao
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
  def instalacao(self, nota):
      self.session.StartTransaction(Transaction="IW53")
      self.session.FindById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = nota
      self.session.FindById("wnd[0]/tbar[1]/btn[5]").Press()
      self.session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09").Select()
      return self.session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09/ssubSUB_GROUP_10:SAPLIQS0:7217/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/ctxtVIQMEL-ZZINSTLN").text
  def historico(self, nota):
    instalacao = self.instalacao(nota)
    self.session.StartTransaction(Transaction="ZSVC20")
    self.session.FindById("wnd[0]/usr/ctxtSO_ANLAG-LOW").text = instalacao
    self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-LOW").text = ""
    self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-HIGH").text = ""
    self.session.FindById("wnd[0]/usr/ctxtP_LAYOUT").text = "/WILLIAM"
    start_time = datetime.datetime.now()
    print("Aguarde relatório sendo processado...")
    self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    print("Relatório processado com sucesso!")
    end_time = datetime.datetime.now()
    print(f"Relatório gerado em {end_time - start_time}")

