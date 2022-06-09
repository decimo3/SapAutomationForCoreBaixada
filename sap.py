#!/usr/bin/python
# coding: utf8

import datetime
import re
import win32com.client
from win10toast import ToastNotifier

from log import log

class sap:
  def __init__(self):
      self.log = log()
      # print(dir(self.SapGui)) ['AddHistoryEntry', 'CreateGuiCollection', 'DropHistory', 'FindById', 'GetScriptingEngine', 'Ignore', 'OpenConnection', 'OpenConnectionByConnectionString', 'OpenWDConnection', 'Quit', 'RegisterROT', 'RevokeROT']
      self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
      # print(dir(self.session)) ['AsStdNumberFormat', 'ClearErrorList', 'CreateSession', 'EnableJawsEvents', 'EndTransaction', 'FindById', 'FindByPosition', 'GetIconResourceName', 'GetVKeyDescription', 'LockSessionUI', 'RunScriptControl', 'SendCommand', 'SendCommandAsync', 'SendMenu', 'StartTransaction', 'UnlockSessionUI']
      self.session = self.SapGui.FindById("ses[0]")
      self.toaster = ToastNotifier()
  def relatorio(self, dia=7):
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
      self.toaster.show_toast("Relatório está pronto!")
  def manobra(self, dia=7):
      self.session.StartTransaction(Transaction="ZSVC20")
      self.session.FindById("wnd[0]/usr/btn%_SO_QMART_%_APP_%-VALU_PUSH").Press()
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "BP"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "BB"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "BD"
      self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
      hoje = datetime.date.today()
      semana = hoje - datetime.timedelta(days=dia)
      self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-LOW").text = semana.strftime("%d.%m.%Y")
      self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-HIGH").text = hoje.strftime("%d.%m.%Y")
      self.session.FindById("wnd[0]/usr/btn%_SO_FECOD_%_APP_%-VALU_PUSH").Press()
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "AP04"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "AP99"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "AP11"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "AP25"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "AP79"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "APRA"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "APRT"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "APTC"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 1
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "APCI"
      self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
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
      self.toaster.show_toast("Relatório está pronto!")
  def leiturista(self, nota):
      instalacao = self.instalacao(nota)
      self.session.StartTransaction(Transaction="ES32")
      self.session.FindById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = instalacao
      self.session.findById("wnd[0]/tbar[0]/btn[0]").Press()
      unidade = self.session.FindById("wnd[0]/usr/tblSAPLES30TC_TIMESL/ctxtEANLD-ABLEINH[9,0]").text
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
      self.session.FindById("wnd[0]/usr/txtP_ABL_Z-LOW").text = centro
      self.session.FindById("wnd[0]/usr/ctxtP_LOTE-LOW").text = livro
      self.session.FindById("wnd[0]/usr/ctxtP_MESANO").text = mes.strftime("%m/%Y")
      self.session.FindById("wnd[0]/usr/ctxtP_UNID-LOW").text = unidade
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      self.session.FindById("wnd[0]/tbar[1]/btn[33]").Press()
      self.session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(27,"DEFAULT")
      self.session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
      self.session.FindById("wnd[0]/tbar[0]/btn[71]").Press()
      self.session.FindById("wnd[1]/usr/txtGS_SEARCH-VALUE").text = instalacao
      self.session.FindById("wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
      self.session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
      self.session.FindById("wnd[1]").Close()
      celula = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow
      linhas = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
      if (celula >= 14):
        self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = celula - 14
        apontador = celula - 14
        limite = celula + 14
      else:
        self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 0
        apontador = 0
        limite = 28
      self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = celula
      leitString = "Seq\tEndereço\tBairro\tMedidor\tHora\tCod\n"
      print(f"Celula: {celula} Apontador: {apontador} Limite: {limite}")
      print(leitString)
      while (apontador < limite and apontador < linhas):
        sequencia = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ZZ_NUMSEQ")
        endereco = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ZENDERECO")
        subbairro = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"BAIRRO")
        medidor = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"GERAET")
        horaleit = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ZHORALEIT")
        codleit = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ABLHINW")
        leitString = f"{leitString}{sequencia}\t{endereco}\t{subbairro}\t{medidor}\t{horaleit}\t{codleit}\n"
        print(f"{sequencia}\t{endereco}\t{subbairro}\t{medidor}\t{horaleit}\t{codleit}")
        apontador = apontador + 1
      return leitString
  def debito(self, nota):
      instalacao = self.instalacao(nota)
      self.session.StartTransaction(Transaction="ZARC140")
      self.session.FindById("wnd[0]/usr/ctxtP_ANLAGE").text = instalacao
      start_time = datetime.datetime.now()
      print("Aguarde relatório sendo processado...")
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      linhas = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").RowCount
      print("Relatório processado com sucesso!")
      end_time = datetime.datetime.now()
      print(f"Relatório gerado em {end_time - start_time}")
      debString = 'Referência\tVencimento\tValor\tTipo\n'
      apontador = 1
      while (apontador < linhas):
        referencia = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"BILLING_PERIOD")
        vencimento = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"FAEDN")
        pendente = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"TOTAL_AMNT")
        faturamento = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"TIP_FATURA")
        debString = f"{debString}{referencia}\t{vencimento}\tR$:{pendente}\t{faturamento}\n"
        apontador = apontador + 1
        print(f"{referencia}\t{vencimento}\tR$:{pendente}\t{faturamento}")
      return debString
  def instalacao(self, arg):
    arg = str(arg)
    if re.search("[0-9]{10}", arg):
      self.session.StartTransaction(Transaction="IW53")
      self.session.FindById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = arg
      self.session.FindById("wnd[0]/tbar[1]/btn[5]").Press()
      self.session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09").Select()
      return self.session.FindById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09/ssubSUB_GROUP_10:SAPLIQS0:7217/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/ctxtVIQMEL-ZZINSTLN").text
    elif re.search("[0-9]{9}", arg):
      return arg
    else:
      return 0
  def historico(self, nota):
    instalacao = self.instalacao(nota)
    self.session.StartTransaction(Transaction="ZSVC20")
    self.session.FindById("wnd[0]/usr/ctxtSO_ANLAG-LOW").text = instalacao
    self.session.FindById("wnd[0]/usr/ctxtSO_QMART-LOW").text = ""
    self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-LOW").text = ""
    self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-HIGH").text = ""
    self.session.FindById("wnd[0]/usr/ctxtP_LAYOUT").text = "/WILLIAM"
    start_time = datetime.datetime.now()
    print("Aguarde relatório sendo processado...")
    self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    print("Relatório processado com sucesso!")
    end_time = datetime.datetime.now()
    print(f"Relatório gerado em {end_time - start_time}")
  def agrupamento(self, nota):
    instalacao = self.instalacao(nota)
    self.session.StartTransaction(Transaction="ES32")
    self.session.FindById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = instalacao
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    consumo = self.session.FindById("wnd[0]/usr/ctxtEANLD-VSTELLE").text
    self.session.StartTransaction(Transaction="ES61")
    self.session.findById("wnd[0]/usr/ctxtEVBSD-VSTELLE").text = consumo
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    ligacao = self.session.FindById("wnd[0]/usr/ctxtEVBSD-HAUS").text
    self.session.StartTransaction(Transaction="ES57")
    self.session.FindById("wnd[0]/usr/ctxtEHAUD-HAUS").text = ligacao
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    logradouro = self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME_CO").text
    numero = self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-HOUSE_NUM1").text
    self.session.StartTransaction(Transaction="ZMED95")
    self.session.FindById("wnd[0]/usr/ctxtADRSTREET-STRT_CODE").text = logradouro
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    self.session.FindById("wnd[0]/tbar[1]/btn[9]").Press()
    if (numero == "1SN" or numero == "SN"):
      return "O agrupamento não pode ser analizado automaticamente"
    linhas = self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").RowCount
    apontador = 1
    while (apontador < linhas):
      # num = self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").getCellValue(apontador,"TI_NUMSX-NUMERO")
      num = self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX/txtTI_NUMSX-NUMERO[0,{apontador}]").text
      if num == numero:
        break
      apontador = apontador + 1
    print(apontador)
    self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").getAbsoluteRow(apontador).selected = True
