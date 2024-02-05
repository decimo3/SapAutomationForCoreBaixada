#!/usr/bin/python
# coding: utf8
import os
import sys
import time
import datetime
import re
import shutil
import subprocess
from os import makedirs
from os import listdir
import win32com.client
import pandas
import dotenv

class sap:
  def __init__(self, instancia) -> None:
      self.CURRENT_FOLDER = os.getcwd() + "\\tmp\\"
      if (not(os.path.exists(self.CURRENT_FOLDER))):
        makedirs(self.CURRENT_FOLDER)
      self.DESTAQUE_AMARELO = 3
      self.DESTAQUE_VERMELHO = 2
      self.DESTAQUE_VERDEJANTE = 4
      self.DESTAQUE_AUSENTE = 0
      self.instancia = instancia
      self.inicializar()
  def inicializar(self) -> bool:
    subprocess.Popen("cscript erroDialog.vbs", stdin=subprocess.PIPE, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
    dotenv.load_dotenv('sap.conf')
    # Get scripting
    try:
      self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    except:
      saplogon = "C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"
      subprocess.Popen(saplogon, start_new_session=True)
      time.sleep(3)
      self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    if not type(self.SapGui) == win32com.client.CDispatch:
        raise Exception("ERRO: SAP GUI Scripting API is not available.")
    # Get connection
    if not (len(self.SapGui.connections) > 0):
      try:
        self.connection = self.SapGui.OpenConnection("#PCL", True)
      except:
        raise Exception("ERRO: SAP FrontEnd connection is not available.")
    else:
      self.connection = self.SapGui.connections[0]
    # Get session
    self.session = self.connection.Children(self.instancia)
    if (self.session.info.user == ''):
      self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = os.environ.get("USUARIO")
      self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = os.environ.get("PALAVRA")
      self.session.findById("wnd[0]/tbar[0]/btn[0]").Press()
      if (self.session.findById("wnd[1]", False) != None):
        self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").Press()
    return (self.session.info.user != '')
  def relatorio(self, dia=7) -> str:
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
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "TABL"
      self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
      self.session.FindById("wnd[0]/usr/ctxtSO_BEBER-LOW").text = os.environ.get("REGIAO")
      self.session.FindById("wnd[0]/usr/ctxtP_LAYOUT").text = os.environ.get("ZSVC20")
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      try:
        exist = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell", False)
        if(exist):
          subprocess.Popen(f"cscript fileDialog.vbs")
          self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
          self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").selectContextMenuItem("&XXL")
          return "FEITO: relatorio salvo no local padrao!"
        else:
          raise Exception("O relatorio de notas em aberto esto vazio!")
      except:
        raise Exception("O relatorio de notas em aberto esto vazio!")
  def manobra(self, dia=0) -> str:
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
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      try:
        exist = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell", False)
        if(exist):
          subprocess.Popen(f"cscript fileDialog.vbs")
          self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
          self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").selectContextMenuItem("&XXL")
          return "FEITO: relatorio salvo no local padrao!"
        else:
          raise Exception("O relatorio de notas em aberto esto vazio!")
      except:
        raise Exception("O relatorio de notas em aberto esto vazio!")
  def leiturista(self, nota, retry=False, order_by_sequence: bool=False) -> str:
      instalacao = self.instalacao(nota)
      self.session.StartTransaction(Transaction="ES32")
      self.session.FindById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = instalacao
      self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
      unidade = self.session.FindById("wnd[0]/usr/tblSAPLES30TC_TIMESL/ctxtEANLD-ABLEINH[9,0]").text
      self.session.FindById("wnd[0]/usr/tblSAPLES30TC_TIMESL/ctxtEANLD-ABLEINH[9,0]").setFocus()
      self.session.FindById("wnd[0]").SendVKey(2)
      centro = self.session.findById("wnd[0]/usr/ctxtTE422-ABL_Z").text
      self.session.StartTransaction(Transaction="ZMED89")
      livro = f"{unidade[0]}{unidade[1]}"
      if (retry): centro = "001"
      mes = datetime.date.today()
      mes = mes.replace(day=1)
      mes = mes - datetime.timedelta(days=1)
      self.session.FindById("wnd[0]/usr/txtP_ABL_Z-LOW").text = centro
      if(retry): self.session.FindById("wnd[0]/usr/txtP_ABL_Z-HIGH").text = "100"
      self.session.FindById("wnd[0]/usr/ctxtP_LOTE-LOW").text = livro
      self.session.FindById("wnd[0]/usr/ctxtP_MESANO").text = mes.strftime("%m/%Y")
      self.session.FindById("wnd[0]/usr/ctxtP_UNID-LOW").text = unidade
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      try:
        self.session.FindById("wnd[0]/tbar[1]/btn[33]").Press()
        self.session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell(0,"DEFAULT")
        self.session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
      except:
        raise Exception("Nao ho relatorio de leitura para o periodo especificado!")
      if(retry or order_by_sequence):
        self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn("ZZ_NUMSEQ")
        self.session.FindById("wnd[0]/tbar[1]/btn[28]").Press()
      self.session.FindById("wnd[0]/tbar[0]/btn[71]").Press()
      self.session.FindById("wnd[1]/usr/txtGS_SEARCH-VALUE").text = instalacao
      self.session.FindById("wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
      self.session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
      # statusBar = self.session.FindById("/app/con[0]/ses[{self.instancia}]/wnd[0]/sbar").text
      # if(statusBar == "Nenhuma ocorrência encontrada"):
        # raise Exception("A instalacao nao foi encontrada no relatorio!")
      self.session.FindById("wnd[1]").Close()
      celula = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow
      if(celula == 0 and instalacao != int(self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(0,"ANLAGE"))):
        raise Exception("A instalacao nao foi encontrada no relatorio!")
      linhas = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
      # se a linhaAtual e menor que 14, a primeiraVisivel e 0 e offset e igual a linha atual
      # se a linhaAtual e maior que linhasTotais - 14, entao primeiraVisivel e linhasTotais - 28 e offset e igual a
      apontador = 0
      limite = 0
      offset = 0
      # Se a instalacao foi encontrada no início do relatorio
      if (celula <= 14):
        self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 0
        apontador = 0
        limite = 28
        offset = celula + 1
      # Se a instalacao foi encontrada no meio do relatorio
      if (celula > 14 and celula < (linhas - 14)):
        self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = celula - 14
        apontador = celula - 14
        limite = celula + 14
        offset = 14 + 1
      # Se a instalacao foi encontrada no final do relatorio
      if(celula > 14 and celula >= (linhas - 14)):
        self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = linhas - 28
        apontador = celula - 28
        limite = linhas
        offset = (28 - (linhas - celula)) + 1
      self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = celula
      leitString = "Cor,Seq,Instalacao,Endereco,Bairro,Medidor,Hora,Cod\n"
      tamanhos = [0,4,10,0,0,8,8,4]
      while (apontador < limite and apontador < linhas):
        destaque = self.DESTAQUE_AMARELO if(apontador == celula) else self.DESTAQUE_AUSENTE
        sequencia = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ZZ_NUMSEQ")
        instalRoteiro = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ANLAGE")
        endereco = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ZENDERECO")
        tamanhos[3] = len(endereco) if (len(endereco) > tamanhos[3]) else tamanhos[3]
        subbairro = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"BAIRRO")
        tamanhos[4] = len(subbairro) if (len(subbairro) > tamanhos[4]) else tamanhos[4]
        medidor = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"GERAET")
        horaleit = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ZHORALEIT")
        codleit = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ABLHINW")
        leitString = f"{leitString}{destaque},{sequencia},{instalRoteiro},{endereco},{subbairro},{medidor},{horaleit},{codleit}\n"
        apontador = apontador + 1
      tamanhosString = f"{tamanhos[0]},{tamanhos[1]},{tamanhos[2]},{tamanhos[3]},{tamanhos[4]},{tamanhos[5]},{tamanhos[6]},{tamanhos[7]}\n"
      leitString = tamanhosString + leitString
      return leitString
  def debito(self, nota, reavisos: bool=False) -> None:
    instalacao = self.instalacao(nota)
    contrato = self.session.FindById("wnd[0]/usr/txtEANLD-VERTRAG").text
    self.session.StartTransaction(Transaction="ZARC140")
    self.session.FindById("wnd[0]/usr/ctxtP_PARTNR").text = ""
    self.session.findById("wnd[0]/usr/ctxtP_VERTRG").text = contrato
    self.session.FindById("wnd[0]/usr/ctxtP_ANLAGE").text = instalacao
    if(reavisos):
      self.session.FindById("wnd[0]/usr/chkC_REAV").Selected = True
    self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
  def escrever(self, nota) -> str:
    self.debito(nota)
    self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110").Select()
    linhas = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").RowCount
    if(linhas < 1): raise Exception("Cliente nao possui faturas pendentes!")
    debString = 'Cor,Mes ref.,Vencimento,Valor,Tipo,Status\n'
    apontador = 1
    tamanhos = [0,7,10,12,0,0]
    while (apontador < linhas):
      referencia = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"BILLING_PERIOD")
      vencimento = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"FAEDN")
      valorPendente = self.sanitizar(self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"TOTAL_AMNT"))
      tipoDebito = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"TIP_FATURA")
      tamanhos[4] = len(tipoDebito) if (len(tipoDebito) > tamanhos[4]) else tamanhos[4]
      statusFat = self.session.findById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador, "STATUS")
      if(statusFat == "@5B@"): # Status no prazo de vencimento da fatura
        destaque = self.DESTAQUE_VERDEJANTE
        textStatus = "Fat. no prazo"
      elif(statusFat == "@5C@"): # Status prazo de pagamento vencido
        destaque = self.DESTAQUE_VERMELHO
        textStatus = "Fat. vencida"
      elif(statusFat == "@06@"): # Status prazo de pagamento vencido
        destaque = self.DESTAQUE_AMARELO
        textStatus = "Fat. Retida"
      else:
        destaque = self.DESTAQUE_AUSENTE
        textStatus = "Consultar"
      tamanhos[5] = len(textStatus) if (len(textStatus) > tamanhos[5]) else tamanhos[5]
      debString = f"{debString}{destaque},{referencia},{vencimento},R$:{valorPendente},{tipoDebito},{textStatus}\n"
      apontador = apontador + 1
    tamanhosString = f"{tamanhos[0]},{tamanhos[1]},{tamanhos[2]},{tamanhos[3]},{tamanhos[4]},{tamanhos[5]}\n"
    debString = tamanhosString + debString
    return debString
  def imprimir(self, documento):
    self.session.StartTransaction(Transaction="ZATC73")
    shutil.rmtree(self.CURRENT_FOLDER)
    makedirs(self.CURRENT_FOLDER)
    self.session.FindById("wnd[0]/usr/chkP_LOCL").selected = True
    self.session.FindById("wnd[0]/usr/chkP_IMPLOC").selected = True
    apontador = 0
    while(apontador < len(documento)):
      self.session.FindById("wnd[0]/usr/ctxtP_OPBEL").text = documento[apontador]
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      win = self.session.findById("wnd[1]/usr/btnSPOP-OPTION1", False)
      if(win != None): win.Press()
      apontador = apontador + 1
  def fatura(self, nota) -> str:
    debitos = []
    apontador = 0
    self.debito(nota)
    linhas = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").RowCount
    while(apontador < linhas):
      if not (self.analisar(apontador)):
        apontador = apontador + 1
        continue
      debitos.append(self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"ZIMPRES"))
      apontador = apontador + 1
    if(len(debitos) > 6 and self.instancia == 0):
      raise Exception(f"Cliente possui muitas faturas ({len(debitos)}) pendentes")
    if(len(debitos) == 0):
      raise Exception("Cliente nao possui faturas vencidas!")
    self.imprimir(debitos)
    return self.monitorar(len(debitos))
  def instalacao(self, arg) -> int:
    try:
      arg = int(arg)
    except:
      raise Exception("Informacao nao e um numero valido!")
    if (arg > 999999999):
      self.session.StartTransaction(Transaction="IW53")
      self.session.FindById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = arg
      self.session.FindById("wnd[0]/tbar[1]/btn[5]").Press()
      try:
        self.session.FindById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09").Select()
      except:
        raise Exception("A nota informada e invalida!")
      instalacao = self.session.FindById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09/ssubSUB_GROUP_10:SAPLIQS0:7217/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/ctxtVIQMEL-ZZINSTLN").text
      self.instalacao(instalacao)
      return instalacao
    if (arg < 999999999 and arg > 99999999):
      self.session.StartTransaction(Transaction="ES32")
      self.session.FindById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = arg
      self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
      return arg
    if(arg < 99999999):
      self.session.StartTransaction(Transaction="IQ03")
      self.session.FindById("wnd[0]/usr/ctxtRISA0-SERNR").text = arg
      try:
        self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
        self.session.FindById(r'wnd[0]/usr/tabsTABSTRIP/tabpT\03/ssubSUB_DATA:SAPMIEQ0:0500/subISUSUB:SAPLE10R:1000/btnBUTTON_ISABL').Press()
        instalacao = self.session.findById("wnd[0]/usr/txtIEANL-ANLAGE").text
        self.instalacao(instalacao)
        return instalacao
      except:
        raise Exception("O numero informado nao eh nota, instalacao ou medidor")
      pass
    raise Exception("O numero informado nao eh nota, instalacao ou medidor")
  def historico(self, nota) -> str:
    instalacao = self.instalacao(nota)
    self.session.StartTransaction(Transaction="ZSVC20")
    self.session.FindById("wnd[0]/usr/ctxtSO_ANLAG-LOW").text = instalacao
    self.session.FindById("wnd[0]/usr/ctxtSO_QMART-LOW").text = ""
    self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-LOW").text = ""
    self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-HIGH").text = ""
    self.session.FindById("wnd[0]/usr/ctxtP_LAYOUT").text = "/WILLIAM"
    self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    linhas = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").RowCount
    apontador = 0
    tamanhos = [0,10,4,0,0,10]
    historico = "Cor,Nota,Tipo,Texto breve para dano,Texto breve para code,Data\n"
    while(apontador < linhas and apontador < 10):
      destaque = 0
      notaServico = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador,"QMNUM")
      dano = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador, "QMART")
      textoDano = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador, "KURZTEXT")
      tamanhos[3] = len(textoDano) if (len(textoDano) > tamanhos[3]) else tamanhos[3]
      textoCode = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador,"MATXT")
      tamanhos[4] = len(textoCode) if (len(textoCode) > tamanhos[4]) else tamanhos[4]
      FimAvaria = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador,"AUSBS")
      historico = f"{historico}{destaque},{notaServico},{dano},{textoDano},{textoCode},{FimAvaria}\n"
      apontador = apontador + 1
    tamanho = f"{tamanhos[0]},{tamanhos[1]},{tamanhos[2]},{tamanhos[3]},{tamanhos[4]},{tamanhos[5]}\n"
    return tamanho + historico
  def agrupamento(self, nota, have_authorization: bool):
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
    if (numero == "1SN" or numero == "SN"):
      raise Exception("O agrupamento nao pode ser analisado automaticamente")
    numero_sem_letra = re.search("[0-9]{1,5}", numero)
    if(numero_sem_letra == None):
      raise Exception("O agrupamento nao pode ser analisado automaticamente")
    self.session.StartTransaction(Transaction="ZMED95")
    self.session.FindById("wnd[0]/usr/ctxtADRSTREET-STRT_CODE").text = logradouro
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    self.session.FindById("wnd[0]/tbar[1]/btn[9]").Press()
    linhas = self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").RowCount
    # Determinar tamanho máximo do grid
    apontador = 0
    tamanho_maximo = 0
    while(True):
      try:
        self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX/txtTI_NUMSX-NUMERO[0,{apontador}]")
        apontador = apontador + 1
        continue
      except:
        tamanho_maximo = apontador - 1
        break
    apontador = 0
    while (apontador < linhas):
      num10 = self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX/txtTI_NUMSX-NUMERO[0,{tamanho_maximo}]").text
      num10_sem_letra = re.search("[0-9]{1,5}", num10)
      if ((num10_sem_letra != None) and (int(num10_sem_letra.group()) < int(numero_sem_letra.group()))):
        apontador = apontador + tamanho_maximo
        self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").verticalScrollbar.position = apontador
        continue
      num = self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX/txtTI_NUMSX-NUMERO[0,0]").text
      if num == numero:
        quantidade = int(self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX/txtTI_NUMSX-QTD_INSTAL[1,0]").text)
        if(quantidade > 12):
          raise Exception(f"Agrupamento possui instalacoes demais ({quantidade})")
        self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").verticalScrollbar.position = apontador
        self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").GetAbsoluteRow(apontador).selected = True
        self.session.FindById("wnd[0]/usr/btn%#AUTOTEXT005").Press()
        break
      apontador = apontador + 1
      self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").verticalScrollbar.position = apontador
    enderecos = []
    instalacoes = []
    nomeCliente = []
    tipoinstal = []
    statusInstalacao = []
    textoDescricao = []
    destaques = []
    tamanhos = [0,0,10,0,10,20]
    agrupamentoString = "Cor,End.,Instalacao,Nome cliente,Tipo cliente,Observacao\n"
    # Coleta das informacões do agrupamento
    apontador = 0
    ultima_instalacao = 0
    while (True):
      if(self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX/txtTI_INSTALX-ANLAGE[1,0]").text == ultima_instalacao):
        break
      enderecos.append(self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX/txtTI_INSTALX-COMPLS[0,0]").text)
      tamanhos[1] = len(enderecos[apontador]) if (len(enderecos[apontador]) > tamanhos[1]) else tamanhos[1]
      instalacoes.append(self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX/txtTI_INSTALX-ANLAGE[1,0]").text)
      nomeCliente.append(self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX/txtTI_INSTALX-NOME[2,0]").text)
      tamanhos[3] = len(nomeCliente[apontador]) if (len(nomeCliente[apontador]) > tamanhos[3]) else tamanhos[3]
      tipoinstal.append(self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX/txtTI_INSTALX-CLASSE[3,0]").text)
      tamanhos[4] = len(tipoinstal[apontador]) if (len(tipoinstal[apontador]) > tamanhos[4]) else tamanhos[4]
      ultima_instalacao = self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX/txtTI_INSTALX-ANLAGE[1,0]").text
      apontador = apontador + 1
      self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX").verticalScrollbar.position = apontador
    if(apontador == 1):
      raise Exception("Instalacao unica para o numero no sistema")
    if(apontador > 12 and self.instancia == 0):
      raise Exception(f"Agrupamento possui instalacoes demais ({apontador})")
    apontador = 0
    # Coleta da situacao das instalacões
    while (apontador < len(instalacoes)):
      self.instalacao(instalacoes[apontador])
      statusInstalacao.append(self.session.findById("wnd[0]/usr/txtEANLD-DISCSTAT").text)
      if(int(instalacoes[apontador]) == instalacao):
        textoDescricao.append("Instalacao da nota")
        destaques.append(self.DESTAQUE_VERMELHO)
        apontador = apontador + 1
        continue
      if(self.session.findById("wnd[0]/usr/txtEANLD-VERTRAG").text == ""):
        textoDescricao.append("Sem contrato ativo")
        destaques.append(self.DESTAQUE_VERMELHO)
        apontador = apontador + 1
        continue
      if(statusInstalacao[apontador] == " Instalação complet.suspensa"):
        textoDescricao.append("Suspensa no sistema")
        destaques.append(self.DESTAQUE_VERMELHO)
        apontador = apontador + 1
        continue
      if(statusInstalacao[apontador] == "Supensao iniciada"):
        textoDescricao.append("Tem ordem de corte")
        destaques.append(self.DESTAQUE_VERMELHO)
        apontador = apontador + 1
        continue
      if(have_authorization): temp = self.novo_analisar(instalacoes[apontador])
      else: temp = self.passivas_novo(instalacoes[apontador])
      if(temp):
        textoDescricao.append("Tem contas passivas")
        destaques.append(self.DESTAQUE_VERMELHO)
        apontador = apontador + 1
        continue
      # caso nao encontre nenhum impedimento
      textoDescricao.append("Cliente nao passivel")
      destaques.append(self.DESTAQUE_VERDEJANTE)
      apontador = apontador + 1
    apontador = 0
    # Preparacao da string final
    tamanhosString = f"{tamanhos[0]},{tamanhos[1]},{tamanhos[2]},{tamanhos[3]},{tamanhos[4]},{tamanhos[5]}\n"
    while (apontador < len(instalacoes)):
      agrupamentoString = f"{agrupamentoString}{destaques[apontador]},{enderecos[apontador]},{instalacoes[apontador]},{nomeCliente[apontador]},{tipoinstal[apontador]},{textoDescricao[apontador]}\n"
      apontador = apontador + 1
    return f"{tamanhosString}{agrupamentoString}\n"
  def coordenadas(self, nota) -> str:
    instalacao = self.instalacao(nota)
    self.session.StartTransaction(Transaction="ES32")
    self.session.FindById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = instalacao
    self.session.findById("wnd[0]/tbar[0]/btn[0]").Press()
    consumo = self.session.FindById("wnd[0]/usr/ctxtEANLD-VSTELLE").text
    self.session.StartTransaction(Transaction="ES61")
    self.session.findById("wnd[0]/usr/ctxtEVBSD-VSTELLE").text = consumo
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    self.session.FindById("wnd[0]/usr/ssubSUB:SAPLXES60:0100/tabsTS0100/tabpTAB1/ssubSUB1:SAPLXES60:0101/ctxtEVBSD-ZZ_TPCAIXA").text = ""
    self.session.FindById("wnd[0]/usr/ssubSUB:SAPLXES60:0100/tabsTS0100/tabpTAB1/ssubSUB1:SAPLXES60:0101/ctxtEVBSD-ZZ_MDCAIXA").text = ""
    self.session.FindById("wnd[0]/usr/ssubSUB:SAPLXES60:0100/tabsTS0100/tabpTAB1/ssubSUB1:SAPLXES60:0101/txtEVBSD-ZZ_NUMCAIXA").text = ""
    self.session.FindById("wnd[0]/usr/ssubSUB:SAPLXES60:0100/tabsTS0100/tabpTAB2").Select()
    coordenada = self.session.FindById("wnd[0]/usr/ssubSUB:SAPLXES60:0100/tabsTS0100/tabpTAB2/ssubSUB1:SAPLXES60:0201/txtEVBSD-ZZ_COORDENADAS").text
    if (len(coordenada) > 0):
      coordenada = re.sub(',', '.', coordenada)
      coordenada = re.findall("-[0-9]{2}.[0-9]*", coordenada)
      return f"https://www.google.com/maps?z=12&t=m&q=loc:{coordenada[0]}+{coordenada[1]}"
    else:
      raise Exception("A instalacao nao possui coordenada cadastrada!")
  def telefone(self, arg) -> str:
    instalacao = self.instalacao(arg)
    parceiro = self.session.findById("wnd[0]/usr/txtEANLD-PARTNER").text
    phone_field_partial_string = self.parceiro(parceiro)
    telefone = []
    nome_cliente = self.session.FindById(phone_field_partial_string + "subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1510/txtBUS_JOEL_MAIN-CHANGE_DESCRIPTION").text
    nome_cliente = str.split(nome_cliente, "/")[0]
    self.session.FindById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01").Select()
    def coletor(self):
      try:
        coleta = []
        coleta.append(self.session.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL2/txtADTEL-TEL_NUMBER[2,0]").text)
        coleta.append(self.session.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL2/txtADTEL-TEL_NUMBER[2,1]").text)
        coleta.append(self.session.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL2/txtADTEL-TEL_NUMBER[2,2]").text)
        coleta.append(self.session.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL2/txtADTEL-TEL_NUMBER[2,3]").text)
        self.session.FindById("wnd[1]").Close()
        return coleta
      except:
        return ""
    if(have_authorization):
      telefone.extend(coletor(self))
      self.session.FindById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/btnG_ICON_MOB").Press()
      telefone.extend(coletor(self))
      self.session.FindById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/subA06P01:SAPLBUA0:0700/subADDR_ICOMM:SAPLSZA11:0100/btnG_ICON_TEL").Press()
      telefone.extend(coletor(self))
      self.session.FindById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/subA06P01:SAPLBUA0:0700/subADDR_ICOMM:SAPLSZA11:0100/btnG_ICON_MOB").Press()
      telefone.extend(coletor(self))
    else:
      telefone.append(self.session.FindById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/txtSZA7_D0400-TEL_NUMBER").text)
      telefone.append(self.session.FindById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/subA05P01:SAPLBUA0:0400/subADDRESS:SAPLSZA7:0600/subCOUNTRY_SCREEN:SAPLSZA7:0601/txtSZA7_D0400-MOB_NUMBER").text)
      telefone.append(self.session.FindById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/subA06P01:SAPLBUA0:0700/subADDR_ICOMM:SAPLSZA11:0100/txtSZA11_0100-TEL_NUMBER").text)
      telefone.append(self.session.FindById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7016/subA06P01:SAPLBUA0:0700/subADDR_ICOMM:SAPLSZA11:0100/txtSZA11_0100-MOB_NUMBER").text)
    # Remove duplicadas do array
    telefone = list(dict.fromkeys(telefone))
    try:
      telefone.remove("______________________________")
    except:
      pass
    texto = nome_cliente + " "
    for tel in telefone:
      texto += tel + " " if (len(tel) > 0) else ""
    return texto
  def medidor(self, nota) -> str:
    instalacao = self.instalacao(nota)
    self.session.StartTransaction(Transaction="ZATC66")
    self.session.FindById("wnd[0]/usr/ctxtP_ANLAGE").text = instalacao
    self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    try:
      self.session.FindById("wnd[0]/usr/subSUB1:SAPLZATC_INFO_CRM:0900/radXSCREEN-HEADER-RB_LEIT").Select()
    except:
      raise Exception(f"A instalacao {instalacao} nao possui historico de consumo para o contrato atual.")
    linhas = self.session.FindById("wnd[0]/usr/cntlCONTROL/shellcont/shell").RowCount
    apontador = 0
    while(apontador < linhas):
      codigo = self.session.FindById("wnd[0]/usr/cntlCONTROL/shellcont/shell").getCellValue(apontador,"OCORRENCIA")
      if ((codigo == "3201") or (codigo == "3202") or (codigo == "3203") or (codigo == "3251")):
        medidor = int(self.session.FindById("wnd[0]/usr/cntlCONTROL/shellcont/shell").getCellValue(apontador,"GERNR"))
        leitura = self.session.FindById("wnd[0]/usr/cntlCONTROL/shellcont/shell").getCellValue(apontador,"ADATSOLL")
        return f"Medidor {medidor} com codigo de retirado pelo leiturista desde {leitura}"
      apontador = apontador + 1
    return "Medidor *nao* consta como retirado"
  def analisar(self, apontador=0, verificar_15_dias=False) -> bool:
    if(apontador == 0): return False
    if (self.session.findById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador, "STATUS") != "@5C@"): return False
    if(verificar_15_dias):
      vencimento = self.session.findById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"FAEDN")
      vencimento = datetime.datetime.strptime(vencimento, f"%d.%m.%Y").date()
      prazo_mais_15_dias = vencimento + datetime.timedelta(days=15)
      if (datetime.date.today() > prazo_mais_15_dias): return False
    return True
  def monitorar(self, qnt) -> str:
    while(len(listdir(self.CURRENT_FOLDER)) < qnt):
      time.sleep(3)
    return "\n".join(listdir(self.CURRENT_FOLDER))
  def novo_medidor(self, arg) -> str:
    instalacao = self.instalacao(arg)
    statusInstalacao = self.session.findById('wnd[0]/usr/txtEANLD-DISCSTAT').text
    endereco = self.session.FindById("wnd[0]/usr/txtEANLD-LINE1").text
    endereco = str.split(endereco, ",")[1]
    cliente = self.session.FindById("wnd[0]/usr/txtEANLD-PARTTEXT").text
    cliente = str.split(cliente, "/")[0]
    dataRetirado = None
    textoStatus = None
    codMedidor = None
    txtCodMedidor = None
    try:
      self.session.FindById("wnd[0]/usr/btnEANLD-DEVSBUT").Press()
    except:
      raise Exception("Instalacao nao tem medidor")
    try:
      dataRetirado = self.session.FindById("wnd[1]/usr/tblSAPLET03UTS_TC/txtPERIODS-BIS[1,0]").text
      self.session.FindById("wnd[1]").SendVKey(2)
    except:
      pass
    medidor = self.session.FindById("wnd[0]/usr/tblSAPLEG70TC_DEVRATE_C/ctxtREG70_D-GERAET[0,0]").text
    self.session.StartTransaction(Transaction="IQ03")
    self.session.FindById("wnd[0]/usr/ctxtRISA0-SERNR").text = medidor
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    codMedidor = self.session.findById("wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/subSUB_0152A:SAPLITO0:1521/ctxtITOB-MATNR").text
    if (codMedidor == "392030"): txtCodMedidor = "CONVENCIONAL MONO"
    elif (codMedidor == "392016"): txtCodMedidor = "CONVENCIONAL MONO"
    elif (codMedidor == "392107"): txtCodMedidor = "CONVENCIONAL BIFASICO"
    elif (codMedidor == "392031"): txtCodMedidor = "CONVENCIAL TRIFASICO"
    elif (codMedidor == "392106"): txtCodMedidor = "CONVENCIAL RURAL"
    elif (codMedidor == "392158"): txtCodMedidor = "CONVENCIONAL MONO"
    elif (codMedidor == "392200"): txtCodMedidor = "TARIFA BRANCA MONO"
    elif (codMedidor == "392201"): txtCodMedidor = "TARIFA BRANCA BIFASICO"
    elif (codMedidor == "392202"): txtCodMedidor = "TARIFA BRANCA TRIFASICA"
    elif (codMedidor == "392143"): txtCodMedidor = "MICRO GERACAO MONO"
    elif (codMedidor == "392144"): txtCodMedidor = "MICRO GERACAO RURAL"
    elif (codMedidor == "392145"): txtCodMedidor = "MICRO GERACAO TRI"
    elif (codMedidor == "392146"): txtCodMedidor = "MICRO GERACAO 30A"
    elif (codMedidor == "392147"): txtCodMedidor = "MICRO GERACAO INDIRETO"
    elif (codMedidor == "391105"): txtCodMedidor = "MODULO MONO"
    elif (codMedidor == "392150"): txtCodMedidor = "MODULO BI"
    elif (codMedidor == "392151"): txtCodMedidor = "MODULO TRI"
    elif (codMedidor == "391108"): txtCodMedidor = "200 AMP"
    elif (codMedidor == "392032"): txtCodMedidor = "BT INDIRETO"
    elif (codMedidor == "392181"): txtCodMedidor = "E-450"
    elif (codMedidor == "601178"): txtCodMedidor = "TELEMEDIDO-SHUNT"
    elif (codMedidor == "313043"): txtCodMedidor = "TRANSF 400-5A"
    elif (codMedidor == "313044"): txtCodMedidor = "TRANSF 1000-5A"
    else: txtCodMedidor = "tipo medidor desconhecido"
    if not(dataRetirado == None):
      textoStatus = f"retirado no sistema desde {dataRetirado}"
      return f"*Medidor:* {medidor}\n*Tipo:* {txtCodMedidor}\n*Status do medidor:* {textoStatus}\n*Instalacao:* {instalacao}\n*Status Instalacao:* {statusInstalacao}\n*Endereco:* {endereco}\n*Cliente:* {cliente}"
    self.session.FindById(r'wnd[0]/usr/tabsTABSTRIP/tabpT\03/ssubSUB_DATA:SAPMIEQ0:0500/subISUSUB:SAPLE10R:1000/btnBUTTON_ISABL').Press()
    apontador = 0
    linhas = self.session.FindById("wnd[0]/usr/cntlBCALVC_EVENT2_D100_C1/shellcont/shell").RowCount
    while(apontador < linhas and apontador < 12):
      status = self.session.findById("wnd[0]/usr/cntlBCALVC_EVENT2_D100_C1/shellcont/shell").getCellValue(apontador,"ABLHINW")
      if(status == "3201"):
        textoStatus = "3201 - medidor retirado"
        break
      if(status == "3202"):
        textoStatus = "3202 - medidor retirado"
        break
      if(status == "3203"):
        textoStatus = "3203 - retirado telemedido"
        break
      if(status == "5800"):
        textoStatus = "5800 - incendiado/demolido"
        break
      apontador = apontador + 1
    if(textoStatus == None):
      textoStatus = "nao esta retirado"
    return f"*Instalacao:* {instalacao}\n*Status Instalacao:* {statusInstalacao}\n*Medidor:* {medidor}\n*Status medidor:* {textoStatus}\n*Tipo:* {txtCodMedidor}"
  def novo_analisar(self, arg) -> bool:
    self.debito(arg, True)
    apontador = 0
    try:
      self.session.findById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF190").Select()
    except:
      return False
    linhas = self.session.FindById(r"wnd[0]/usr/tabsTAB_STRIP_100/tabpF190/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0190/cntlCONTAINER_190/shellcont/shell").RowCount
    if(linhas == 0): return False
    while(apontador < linhas):
      status = self.session.findById(r"wnd[0]/usr/tabsTAB_STRIP_100/tabpF190/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0190/cntlCONTAINER_190/shellcont/shell").getCellValue(apontador, "STATUS")
      dtMax = self.session.findById(r"wnd[0]/usr/tabsTAB_STRIP_100/tabpF190/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0190/cntlCONTAINER_190/shellcont/shell").getCellValue(apontador, "DT_MAX_CRT")
      dtMin = self.session.findById(r"wnd[0]/usr/tabsTAB_STRIP_100/tabpF190/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0190/cntlCONTAINER_190/shellcont/shell").getCellValue(apontador, "DT_MIN_CRT")
      if(status == "@45@"): return True
      if(dtMin == "" or dtMax == ""):
        apontador = apontador + 1
        continue
      dtMax = datetime.datetime.strptime(dtMax, f"%d.%m.%Y").date()
      dtMin = datetime.datetime.strptime(dtMin, f"%d.%m.%Y").date()
      if(datetime.date.today() > dtMin and datetime.date.today() < dtMax): return True
      apontador = apontador + 1
    return False
  def passivas(self, arg) -> str:
    self.debito(arg, True)
    apontador = 0
    passiveis = []
    try:
      self.session.findById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF190").Select()
    except:
      raise Exception("Nao ha faturas passiveis!")
    linhas = self.session.FindById(r"wnd[0]/usr/tabsTAB_STRIP_100/tabpF190/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0190/cntlCONTAINER_190/shellcont/shell").RowCount
    if(linhas == 0): raise Exception("Nao ha faturas passiveis!")
    while(apontador < linhas):
      fatura = self.session.findById(r"wnd[0]/usr/tabsTAB_STRIP_100/tabpF190/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0190/cntlCONTAINER_190/shellcont/shell").getCellValue(apontador, "ZIMPRES")
      status = self.session.findById(r"wnd[0]/usr/tabsTAB_STRIP_100/tabpF190/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0190/cntlCONTAINER_190/shellcont/shell").getCellValue(apontador, "STATUS")
      dtMax = self.session.findById(r"wnd[0]/usr/tabsTAB_STRIP_100/tabpF190/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0190/cntlCONTAINER_190/shellcont/shell").getCellValue(apontador, "DT_MAX_CRT")
      dtMin = self.session.findById(r"wnd[0]/usr/tabsTAB_STRIP_100/tabpF190/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0190/cntlCONTAINER_190/shellcont/shell").getCellValue(apontador, "DT_MIN_CRT")
      if(status == "@45@"):
        passiveis.append(fatura)
        apontador = apontador + 1
        continue
      if(dtMin == "" or dtMax == ""):
        apontador = apontador + 1
        continue
      dtMax = datetime.datetime.strptime(dtMax, f"%d.%m.%Y").date()
      dtMin = datetime.datetime.strptime(dtMin, f"%d.%m.%Y").date()
      if(datetime.date.today() > dtMin and datetime.date.today() < dtMax):
        passiveis.append(fatura)
      apontador = apontador + 1
    if(len(passiveis) > 5 and self.instancia == 0):
      raise Exception(f"Cliente possui muitas faturas ({len(passiveis)}) passivas")
    self.imprimir(passiveis)
    return self.monitorar(len(passiveis))
  def sanitizar(self, arg) -> str:
    arg = str.replace(arg, ' ', '')
    arg = str.replace(arg, '.', '')
    arg = str.replace(arg,',','.')
    if '-' in arg:
      arg = '-' + arg.replace('-', '')
    return arg
  def escrever_novo(self, arg, doc_impressao: bool=False, so_passivas: bool=False) -> str | list[str]:
    self.instalacao(arg)
    contrato = self.session.FindById("wnd[0]/usr/txtEANLD-VERTRAG").text
    self.session.StartTransaction(Transaction="FPL9")
    self.session.findById("wnd[0]/usr/ctxtFKKL1-GPART").text = ""
    self.session.findById("wnd[0]/usr/ctxtFKKL1-VTREF").text = contrato
    self.session.findById("wnd[0]/tbar[0]/btn[0]").Press()
    self.session.findById("wnd[0]/tbar[1]/btn[39]").Press()
    # Check if you are still on the FPL9 transaction form screen
    if(self.session.findById("wnd[0]/usr/ctxtFKKL1-VTREF", False) != None):
      if(so_passivas == True): return []
      raise Exception("Cliente nao possui faturas pendentes!")
    #[char, line]
    col = 1
    row = 0
    tamanhos = [0,13,10,12,10,9]
    colunas = [0,0,0,0,0,0]
    datasete = {
      "destaques": [],
      "status": [],
      "referencia": [],
      "impressao": [],
      "vencimento": [],
      "valores": []
    }
    hasLines = True
    MAX_LINES = 32
    firstLine = False
    while(hasLines):
      # Aponta para a linha atual ou para a última caso o índice da linha tenha ultrapassado o máximo
      linha = MAX_LINES if(row >= MAX_LINES) else row
      # Verifica se a coluna VENCIMENTO não está vazia, se estiver, encerra a leitura dos labels
      if(colunas[2] > 0 and self.session.FindById(f"wnd[0]/usr/lbl[1,{linha}]", False) == None):
        if(firstLine == False):
          firstLine = True
          row = row + 1
          continue
        else:
          hasLines = False
          continue
      # Verifica se tem label na linha na coluna 1, caso não tenha a linha é vazia, então pula ela
      if(self.session.FindById(f"wnd[0]/usr/lbl[1,{linha}]", False) == None):
        row = row + 1
        if(row >= MAX_LINES):
          self.session.FindById("wnd[0]/usr").verticalScrollbar.position = row - MAX_LINES
        continue
      while(col < 99 and colunas[5] == 0):
        label = self.session.FindById(f"wnd[0]/usr/lbl[{col},{linha}]", False)
        # Verifica se a coordenada do objeto retorna um objeto, se não pula para a próxima coluna
        if(label == None):
          col = col + 1
          continue
        if(label.text == "Sts"): colunas[1] = col
        if(label.text == "Mês Refer"): colunas[2] = col
        if(label.text == "Doc. Faturam"): colunas[3] = col
        if(label.text == "Vencimento"): colunas[4] = col
        if(label.text == "Valor"): colunas[5] = col
        col = col + 1
      if(colunas[2] > 0 and firstLine == True):
        status = self.session.FindById(f"wnd[0]/usr/lbl[{colunas[1]},{linha}]", False).iconName
        if(status == "S_TL_R"):
          datasete["destaques"].append(self.DESTAQUE_VERMELHO)
          datasete["status"].append("Fat. vencida")
        if(status == "S_TL_Y"):
          datasete["destaques"].append(self.DESTAQUE_VERDEJANTE)
          datasete["status"].append("Fat. no prazo")
        if(status == "S_TL_G"):
          datasete["destaques"].append(self.DESTAQUE_VERDEJANTE)
          datasete["status"].append("Fat. no prazo")
        if(status != "S_TL_R" and status != "S_TL_Y" and status != "S_TL_G"):
          datasete["destaques"].append(self.DESTAQUE_AUSENTE)
          datasete["status"].append("")
        datasete["referencia"].append(self.session.FindById(f"wnd[0]/usr/lbl[{colunas[2]},{linha}]", False).text)
        datasete["impressao"].append(self.session.FindById(f"wnd[0]/usr/lbl[{colunas[3]},{linha}]", False).text)
        datasete["vencimento"].append(datetime.datetime.strptime(self.session.FindById(f"wnd[0]/usr/lbl[{colunas[4]},{linha}]", False).text ,"%d.%m.%Y"))
        datasete["valores"].append(float(self.sanitizar(self.session.FindById(f"wnd[0]/usr/lbl[{colunas[5]},{linha}]", False).text)))
      row = row + 1
      if(row >= MAX_LINES):
        self.session.FindById("wnd[0]/usr").verticalScrollbar.position = row - MAX_LINES
      col = 1
    dt1 = pandas.DataFrame(datasete)
    dt2 = dt1.groupby('impressao')['valores'].sum().reset_index()
    dt1.drop_duplicates(subset="impressao",inplace=True)
    dt3 = dt1.merge(dt2, on="impressao")
    del dt3['valores_x']
    dt3 = dt3.rename(columns={'valores_y': 'valores'})
    dt3['impressao'].replace('', pandas.NA, inplace=True)
    dt3 = dt3.dropna(subset=['impressao'])
    tamanhoString = f"{tamanhos[0]},{tamanhos[1]},{tamanhos[2]},{tamanhos[3]},{tamanhos[4]},{tamanhos[5]}\n"
    if(so_passivas):
      dt3["vencimento"] = pandas.to_datetime(dt3['vencimento'])
      prazo_minimo = datetime.date.today() - datetime.timedelta(days=15)
      prazo_maximo = datetime.date.today() - datetime.timedelta(days=90)
      dt3 = dt3[dt3['vencimento'] < pandas.to_datetime(prazo_minimo)]
      dt3 = dt3[dt3['vencimento'] > pandas.to_datetime(prazo_maximo)]
      return dt3['impressao'].to_list()
    if(doc_impressao):
      dt3 = dt3[dt3['status'] != "Fat. no prazo"]
      return dt3['impressao'].to_list()
    return tamanhoString + dt3.to_csv(index = False)
  def fatura_novo(self, arg) -> str:
    debitos = self.escrever_novo(arg, True)
    if(len(debitos) > 6): raise Exception(f"Cliente possui muitas faturas ({len(debitos)}) pendentes")
    self.imprimir(debitos)
    return self.monitorar(len(debitos))
  def passivas_novo(self, arg) -> bool:
    debitos = self.escrever_novo(arg, False, True)
    return (len(debitos) > 0)
  def informacao(self, arg) -> str:
    result = self.novo_medidor(arg)
    instalacao = self.instalacao(arg)
    parceiro = self.session.findById("wnd[0]/usr/txtEANLD-PARTNER").text
    phone_field_partial_string = self.parceiro(parceiro)
    nome_cliente = self.session.FindById(phone_field_partial_string + "subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1510/txtBUS_JOEL_MAIN-CHANGE_DESCRIPTION").text
    nome_cliente = str.split(nome_cliente, "/")[0]
    self.session.findById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_04").Select()
    pessoa_fisica = self.session.findById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_04/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7006/subA04P01:SAPLBUPA_BUTX_DIALOG:0100/tblSAPLBUPA_BUTX_DIALOGTCTRL_BPTAX/txtDFKKBPTAXNUM-TAXNUM[2,0]").text
    return result + f"\n*Cod. do cliente:* {parceiro}\n*Cadastro Pessoa Fisica (CPF):* {pessoa_fisica}\n*Nome do cliente:* {nome_cliente}"
  def parceiro(self, parceiro, have_authorization: bool=True) -> str:
    SAPLBUS_LOCATOR = "2000" if(have_authorization) else "2036"
    phone_field_partial_string = f"wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:{SAPLBUS_LOCATOR}/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
    self.session.StartTransaction(Transaction="BP")
    # close search side panel
    self.session.findById("wnd[0]/tbar[1]/btn[9]").Press()
    # Click 'Open PN' button
    self.session.findById("wnd[0]/tbar[1]/btn[17]").Press()
    self.session.findById("wnd[1]/usr/ctxtBUS_JOEL_MAIN-OPEN_NUMBER").text = parceiro
    self.session.findById("wnd[1]/tbar[0]/btn[0]").Press()
    # Click 'dados gerais' button
    self.session.findById("wnd[0]/tbar[1]/btn[25]").Press()
    self.session.findById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/subSCREEN_1100_ROLE_AND_TIME_AREA:SAPLBUPA_DIALOG_JOEL:1110/cmbBUS_JOEL_MAIN-PARTNER_ROLE").key = "MKK"
    if(self.session.findById("wnd[1]", False) != None): self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").Press()
    return phone_field_partial_string
  def cruzamento(self, arg) -> str:
    instalacao = self.instalacao(arg)
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
    self.session.StartTransaction(Transaction="ZMED95")
    self.session.FindById("wnd[0]/usr/ctxtADRSTREET-STRT_CODE").text = logradouro
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    self.session.FindById("wnd[0]/tbar[1]/btn[7]").Press()
    container = self.session.FindById("wnd[0]/usr/cntlCONTAINER_200/shellcont/shell")
    linhas = container.RowCount
    apontador = 0
    dataframe = {
      "#": [],
      "Num. 1": [],
      "Cod. Cruza": [],
      "Cod. rua": [],
      "Logradouro": [],
      "Num. 2": [],
      "Latitude": [],
      "Longitude": [],
      "Cidade": [],
    }
    tamanhos = [0,6,10,9,0,6,16,16,0]
    while(apontador < linhas):
      # 0. "COLORACAO"
      dataframe['#'].append("0")
      # 1. "NUMERO"
      dataframe['Num. 1'].append(container.getCellValue(apontador, "NUMERO"))
      # 2. "COD_CRUZAMENTO"
      dataframe['Cod. Cruza'].append(container.getCellValue(apontador, "COD_CRUZAMENTO"))
      # 3. "STRT_CODE"
      dataframe['Cod. rua'].append(container.getCellValue(apontador, "STRT_CODE"))
      # 4. "STREET"
      dataframe['Logradouro'].append(container.getCellValue(apontador, "STREET"))
      tamanho_campo = len(container.getCellValue(apontador, "STREET"))
      tamanhos[4] = tamanhos[4] if(tamanhos[4] > tamanho_campo) else tamanho_campo
      # 5. "NUMERO_OUTR"
      dataframe['Num. 2'].append(container.getCellValue(apontador, "NUMERO_OUTR"))
      # 6. "LATITUDE"
      dataframe['Latitude'].append(str.replace(container.getCellValue(apontador, "LATITUDE"), ',', '.'))
      # 7. "LONGITUDE"
      dataframe['Longitude'].append(str.replace(container.getCellValue(apontador, "LONGITUDE"), ',', '.'))
      # 8. "CITY_NAME"
      dataframe['Cidade'].append(container.getCellValue(apontador, "CITY_NAME"))
      tamanho_campo = len(container.getCellValue(apontador, "CITY_NAME"))
      tamanhos[8] = tamanhos[8] if(tamanhos[8] > tamanho_campo) else tamanho_campo
      apontador = apontador + 1
    tamanhos_string = ",".join([str(x) for x in tamanhos])
    resultados_csv = pandas.DataFrame(dataframe).to_csv(index=False, sep=',')
    return tamanhos_string + '\n' + resultados_csv

if __name__ == "__main__":
  # Validação dos argumentos da linha de comando:
  # 1. Será possível realizar a tarefa se no mínimo
  #    três argumentos forem passados para a aplicação
  #    (script path mais dois do usuário). Caso contrário
  #    não será possível inferir a posição dos argumentos;
  # 2. Os argumentos devem obedecer a ordem:
  #    0. Caminho do script e nome;
  #    1. Nome da aplicação desejada (somente letras);
  #    2. Número associado ao serviço;
  #    3. Número da instância utilizada;
  #    4. Outros argumentos opcionais;
  # 3. Se houver somente 3 argumentos, então é uma consulta simples
  #    e o script é configurado automaticamente para usar a 0.
  if(len(sys.argv) < 3): raise Exception("Falta argumentos necessarios!")
  if(not str.isalpha(sys.argv[1])): raise Exception("O primeiro argumento é inválido!")
  aplicacao = sys.argv[1]
  argumento = int(sys.argv[2])
  if(len(sys.argv) == 3): instancia = 0
  else: instancia = int(sys.argv[3])
  # Attempts to connect to SAP FrontEnd on the specified instance
  try: robo = sap(instancia)
  except: raise Exception("ERRO: Nao pode se conectar ao sistema SAP!")
  have_authorization = True
  # If the number of arguments is greater than the minimum (4),
  # then it checks the other arguments (now, only one optional argument is accepted).
  if(len(sys.argv) > 4):
    apontador = 4
    while(apontador < len(sys.argv)):
      if ('--sap-restrito' == sys.argv[apontador]): have_authorization = False
      else: raise Exception("O argumento fornecido nao eh valido!")
      apontador = apontador + 1
  # Attempts to execute the method requested in the first argument
  try:
    if (aplicacao == "coordenada"):
      print(robo.coordenadas(argumento))
    elif ((aplicacao == "telefone") or (aplicacao == "contato")):
      if(not have_authorization): raise Exception("Nao eh possivel consultar essas informacoes no modo restrito")
      print(robo.telefone(argumento))
    elif (aplicacao == "medidor"):
      print(robo.novo_medidor(argumento))
    elif ((aplicacao == "leiturista") or (aplicacao == "roteiro")):
      try:
        if(aplicacao == "roteiro"):
          print(robo.leiturista(argumento, False, True))
        else:
          print(robo.leiturista(argumento, False, False))
      except:
        if(aplicacao == "roteiro"):
          print(robo.leiturista(argumento, True, True))
        else:
          print(robo.leiturista(argumento, True, False))
    elif ((aplicacao == "debito") or (aplicacao == "fatura")):
      if(have_authorization):
        print(robo.fatura(argumento))
      else:
        print(robo.fatura_novo(argumento))
    elif (aplicacao == "relatorio"):
      robo.relatorio(argumento)
    elif ((aplicacao == "historico") or (aplicacao == "historico")):
      print(robo.historico(argumento))
    elif (aplicacao == "agrupamento"):
        print(robo.agrupamento(argumento, have_authorization))
    elif (aplicacao == "pendente"):
      if(have_authorization):
        print(robo.escrever(argumento))
      else:
        print(robo.escrever_novo(argumento))
    elif (aplicacao == "manobra"):
      print(robo.manobra(argumento))
    elif(aplicacao == "informacao"):
      if(not have_authorization): raise Exception("Nao eh possivel consultar essas informacoes no modo restrito")
      else: print(robo.informacao(argumento))
    elif(aplicacao == "desperta"):
      print(robo.instalacao(argumento))
    elif(aplicacao == "cruzamento"):
      print(robo.cruzamento(argumento))
    else:
      raise Exception("Nao entendi o comando, verifique se esto correto!")
  # Returns the error with an 'ERROR:' prefix on method failure
  except Exception as erro:
    print(f"ERRO: {erro.args[0]}")