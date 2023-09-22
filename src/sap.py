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

class sap:
  def __init__(self, instancia=0) -> None:
      self.CURRENT_FOLDER = os.getcwd() + "\\tmp\\"
      if (not(os.path.exists(self.CURRENT_FOLDER))):
        makedirs(self.CURRENT_FOLDER)
      self.DESTAQUE_AMARELO = 3
      self.DESTAQUE_VERMELHO = 2
      self.DESTAQUE_VERDEJANTE = 4
      self.DESTAQUE_AUSENTE = 0
      self.instancia = instancia
      self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
      self.session = self.SapGui.FindById(f"ses[{self.instancia}]")
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
  def leiturista(self, nota, retry=False) -> str:
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
      if(retry):
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
    linhas = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").RowCount
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
      # if not (self.session.FindById("wnd[1]/usr/btnSPOP-OPTION1") == None):
      #   self.session.FindById("wnd[1]/usr/btnSPOP-OPTION1").Press()
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
    if(len(debitos) > 5 and self.instancia == 0):
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
    tamanhos = [0,10,0,0,10]
    historico = "Cor,Nota,Texto breve para dano,Texto breve para code,Data\n"
    while(apontador < linhas and apontador < 10):
      destaque = 0
      notaServico = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador,"QMNUM")
      textoDano = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador, "KURZTEXT")
      tamanhos[2] = len(textoDano) if (len(textoDano) > tamanhos[2]) else tamanhos[2]
      textoCode = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador,"MATXT")
      tamanhos[3] = len(textoCode) if (len(textoCode) > tamanhos[3]) else tamanhos[3]
      FimAvaria = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador,"AUSBS")
      historico = f"{historico}{destaque},{notaServico},{textoDano},{textoCode},{FimAvaria}\n"
      apontador = apontador + 1
    tamanho = f"{tamanhos[0]},{tamanhos[1]},{tamanhos[2]},{tamanhos[3]},{tamanhos[4]}\n"
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
  def telefone(self, info, have_authorization: bool) -> str:
    SAPLBUS_LOCATOR = "2000" if(have_authorization) else "2036"
    phone_field_partial_string = f"wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:{SAPLBUS_LOCATOR}/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/"
    telefone = []
    nome_solicitante = ""
    try:
      info = int(info)
    except:
      raise Exception("Informacao nao e um numero valido!")
    if (info > 999999999):
      self.session.StartTransaction(Transaction="IW53")
      self.session.FindById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = info
      try:
        self.session.FindById("wnd[0]/tbar[1]/btn[5]").Press()
      except:
        raise Exception("Numero da nota e invalido!")
      self.session.FindById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09").Select()
      nome_solicitante = self.session.FindById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09/ssubSUB_GROUP_10:SAPLIQS0:7217/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/txtVIQMEL-ZZ_NOME_SOLICIT").text
      telefone.append(self.session.FindById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09/ssubSUB_GROUP_10:SAPLIQS0:7217/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/txtVIQMEL-ZZ_TEL_SOLICIT").text)
      telefone.append(self.session.FindById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09/ssubSUB_GROUP_10:SAPLIQS0:7217/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/txtVIQMEL-ZZ_CEL_SOLICIT").text)
      info = self.session.FindById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09/ssubSUB_GROUP_10:SAPLIQS0:7217/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/ctxtVIQMEL-ZZINSTLN").text
    self.session.StartTransaction(Transaction="ES32")
    self.session.FindById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = info
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    cliente = self.session.FindById("wnd[0]/usr/txtEANLD-PARTNER").text
    cliente = str.split(cliente, "/")[0]
    self.session.StartTransaction(Transaction="BP")
    try:
      self.session.FindById(phone_field_partial_string + "subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1510/ctxtBUS_JOEL_MAIN-CHANGE_NUMBER").text = cliente
    except:
      self.session.FindById(f"wnd[0]/tbar[1]/btn[9]").Press()
      self.session.FindById(phone_field_partial_string + "subSCREEN_1000_HEADER_AREA:SAPLBUPA_DIALOG_JOEL:1510/ctxtBUS_JOEL_MAIN-CHANGE_NUMBER").text = cliente
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
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
    texto = nome_solicitante + " " if (len(nome_solicitante) > 0) else nome_cliente + " "
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
    self.session.StartTransaction(Transaction="ES32")
    self.session.FindById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = instalacao
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
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
    return f"*Medidor:* {medidor}\n*Tipo:* {txtCodMedidor}\n*Status medidor:* {textoStatus}\n*Instalacao:* {instalacao}\n*Status Instalacao:* {statusInstalacao}\n*Endereco:* {endereco}\n*Cliente:* {cliente}"
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
    arg = str.replace(arg,',','.')
    arg = arg.strip()
    return arg
  def escrever_novo(self, arg, doc_impressao: bool=False) -> str | list[str]:
    self.instalacao(arg)
    contrato = self.session.FindById("wnd[0]/usr/txtEANLD-VERTRAG").text
    self.session.StartTransaction(Transaction="FPL9")
    self.session.findById("wnd[0]/usr/ctxtFKKL1-VTREF").text = contrato
    self.session.findById("wnd[0]/tbar[0]/btn[0]").Press()
    self.session.findById("wnd[0]/tbar[1]/btn[39]").Press()
    #[collum, line]
    col = 1
    row = 0
    tamanhos = [0,7,12,10,9]
    colunas = [0,0,0,0,0]
    response = ""
    debitos = []
    while(row < 99):
      while(col < 99):
        label = self.session.FindById(f"wnd[0]/usr/lbl[{col},{row}]", False)
        if(label == None):
          col = col + 1
          continue
        # Verifica se a coluna DOC_IMPRESSAO não está preenchida com VOID
        # if(colunas[2] > 0):
        #   temp = self.session.FindById(f"wnd[0]/usr/lbl[{colunas[2]},{row}]", False)
        #   if(temp == None): break
        if(col == colunas[0]): response = response + f"0," # Preenchimento da COR_DESTAQUE
        if(col == colunas[1]): response = response + f"{label.text}," # Preenchimento do REFERENCIA
        # Preenchimento do DOC_IMPRESSAO
        if(col == colunas[2]):
          response = response + f"{label.text},"
          debitos.append(label.text)
        if(col == colunas[3]): response = response + f"{label.text}," # Preenchimento da VENCIMENTO
        if(col == colunas[4]): response = response + f"R$: {self.sanitizar(label.text)}\n" # Preenchimento do VALOR
        if(label.text == "Sts"): colunas[0] = col
        if(label.text == "Mês Refer"): colunas[1] = col
        if(label.text == "Doc. Faturam"): colunas[2] = col
        if(label.text == "Vencimento"): colunas[3] = col
        if(label.text == "Valor"): colunas[4] = col
        col = col + 1
      row = row + 1
      col = 1
    tamanhoString = f"{tamanhos[0]},{tamanhos[1]},{tamanhos[2]},{tamanhos[3]},{tamanhos[4]}\n"
    respostaString = f"{tamanhoString}Cor,Mes ref.,Doc. Faturam,Vencimento,Valor\n{response}"
    if(doc_impressao):
      return debitos
    else:
      return respostaString
  def fatura_novo(self, arg) -> str:
    debitos = self.escrever_novo(arg, True)
    if(len(debitos) > 6): raise Exception(f"Cliente possui muitas faturas ({len(debitos)}) pendentes")
    self.imprimir(debitos)
    return self.monitorar(len(debitos))
  def passivas_novo(self, arg) -> str:
    
    return ""

if __name__ == "__main__":
  if (len(sys.argv) < 3):
    raise Exception("Falta argumentos para relizar alguma acao!")
  if (len(sys.argv) > 4):
    raise Exception("Script nao foi programado para essa quantidade de argumentos!")
  try:
    robo = sap() if (len(sys.argv) == 3) else sap(int(sys.argv[3]))
  except:
    raise Exception("ERRO: Nao pode se conectar ao sistema SAP!")
  if(os.getenv("SAP_PERMISSIONS") == None):
    have_authorization = True
  else:
    have_authorization = bool(int(os.getenv("SAP_PERMISSIONS", "0")))
  try:
    if ((sys.argv[1] == "coordenada") or (sys.argv[1] == "localizacao")):
      print(robo.coordenadas(int(sys.argv[2])))
    elif ((sys.argv[1] == "telefone") or (sys.argv[1] == "contato")):
      print(robo.telefone(int(sys.argv[2]), have_authorization))
    elif (sys.argv[1] == "medidor"):
      print(robo.novo_medidor(int(sys.argv[2])))
    elif ((sys.argv[1] == "leiturista") or (sys.argv[1] == "roteiro")):
      try:
        print(robo.leiturista(int(sys.argv[2])))
      except:
        print(robo.leiturista(int(sys.argv[2]), True))
    elif ((sys.argv[1] == "debito") or (sys.argv[1] == "fatura") or (sys.argv[1] == "debito")):
      if(have_authorization):
        print(robo.fatura(int(sys.argv[2])))
      else:
        print(robo.fatura_novo(int(sys.argv[2])))
    elif (sys.argv[1] == "relatorio"):
      robo.relatorio(int(sys.argv[2]))
    elif ((sys.argv[1] == "historico") or (sys.argv[1] == "historico")):
      print(robo.historico(sys.argv[2]))
    elif (sys.argv[1] == "agrupamento"):
        print(robo.agrupamento(sys.argv[2], have_authorization))
    elif (sys.argv[1] == "pendente"):
      if(have_authorization):
        print(robo.escrever(int(sys.argv[2])))
      else:
        print(robo.escrever_novo(int(sys.argv[2])))
    elif (sys.argv[1] == "manobra"):
      print(robo.manobra(int(sys.argv[2])))
    else:
      raise Exception("Nao entendi o comando, verifique se esto correto!")
  except Exception as erro:
    print(f"ERRO: {erro.args[0]}")