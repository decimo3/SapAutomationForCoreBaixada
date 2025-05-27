# pylint: skip-file
#!/usr/bin/python
# coding: utf8
#region imports
import os
import io
import sys
import time
import datetime
import re
import subprocess
import win32com.client
import pandas
import dotenv
import sqlite3
import logging
from logging.handlers import RotatingFileHandler
#endregion

class sap:
  def __init__(self, instancia) -> None:
    environment = "--em-desenvolvimento" in sys.argv
    if environment:
      self.BASE_FOLDER = os.path.expandvars("%USERPROFILE%\\MestreRuan\\")
    else:
      if getattr(sys, 'frozen', False):
        # The application is frozen (PyInstaller bundled)
        self.BASE_FOLDER = os.path.dirname(sys.executable)
      else:
        # Running in normal Python mode
        self.BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))
    logfilename = os.path.join(self.BASE_FOLDER, f"logfile_{instancia}.log")
    logging.basicConfig(
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        level=logging.DEBUG,
        handlers=[
          RotatingFileHandler(logfilename, maxBytes=10000000, backupCount=5),
          logging.StreamHandler()
        ] 
    )
    self.logger = logging.getLogger(__name__)
    self.LOCKFILE = os.path.join(self.BASE_FOLDER, 'sap.lock')
    CONF_FILE = os.path.join(self.BASE_FOLDER, 'sap.conf')
    dotenv.load_dotenv(CONF_FILE)
    self.NOTUSE = str(os.environ.get("NOTUSE")).split(',')
    self.SETOR = os.environ.get("SETOR")
    if(self.SETOR == None): raise Exception("500: A variavel SETOR no arquivo `sap.config` nao esta definida!")
    self.REGIAO = os.environ.get("REGIAO")
    if(self.REGIAO == None): raise Exception("500: A variavel REGIAO no arquivo `sap.config` nao esta definida!")
    self.LAYOUT = os.environ.get("LAYOUT")
    if(self.LAYOUT == None): raise Exception("500: A variavel LAYOUT no arquivo `sap.config` nao esta definida!")
    self.ATIVIDADES = self.depara('setor_atividades', self.SETOR).split(',')
    self.DESTAQUE_AMARELO = 3
    self.DESTAQUE_VERMELHO = 2
    self.DESTAQUE_VERDEJANTE = 4
    self.DESTAQUE_AUSENTE = 0
    self.instancia = instancia - 1
  def create_lock(self) -> None:
    if not os.path.exists(self.LOCKFILE):
      with open(self.LOCKFILE, 'w') as file:
        pass
  def delete_lock(self) -> None:
    if(os.path.exists(self.LOCKFILE)):
      os.remove(self.LOCKFILE)
  def check_lock(self) -> bool:
    return os.path.exists(self.LOCKFILE)
  def health_check(self, desire_instances: int = 5) -> None:
    """ Check, create and recreate SAP instances """
    TIME_WAIT = 5
    INTERVAL_CHECK = 10
    self.create_lock()
    self.logger.info("Starting instance checker...")
    while(True):
      if not self.check_lock():
        time.sleep(INTERVAL_CHECK)
        continue
      try:
        # Get scripting
        try:
          self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
        except:
          self.create_lock()
          self.logger.warning("Cannot able to attach 'ScriptingEngine'. Starting SAP...")
          saplogon = "C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"
          subprocess.Popen(saplogon, start_new_session=True)
          time.sleep(TIME_WAIT)
          self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
        if not type(self.SapGui) == win32com.client.CDispatch:
          self.create_lock()
          self.logger.error("Cannot able to attach 'ScriptingEngine', even after start SAP Frontend.")
          raise Exception("500: SAP GUI Scripting API is not available.")
        # Get connection
        if not (len(self.SapGui.connections) > 0):
          self.create_lock()
          self.logger.warning("Connection is not open. trying open...")
          try:
            self.connection = self.SapGui.OpenConnection("#PCL", True)
          except:
            self.create_lock()
            self.logger.error("Cannot able to open connection with SAP server.")
            raise Exception("500: SAP FrontEnd connection is not available.")
        else:
          self.connection = self.SapGui.connections[0]
        self.session = self.connection.Children(0)
        # Check and get authentication
        if (self.session.info.user == ''):
          self.create_lock()
          self.logger.warning("User is not authenticated yet. Authenticating...")
          self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = os.environ.get("USUARIO")
          self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = os.environ.get("PALAVRA")
          self.session.findById("wnd[0]/tbar[0]/btn[0]").Press()
          if (self.session.findById("wnd[1]", False) != None):
            if (self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1", False) != None):
              self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").Press()
        # re-check authentication
        if(self.session.info.user == ''):
          self.create_lock()
          self.logger.error("User cannot be authenticated.")
          raise Exception("500: User cannot be authenticated!")
        # Create sessions
        number_of_sessions = len(self.connection.sessions)
        if(number_of_sessions < desire_instances):
          self.create_lock()
          self.logger.warning("Less instances that desire, creating new ones...")
          for i in range(desire_instances - number_of_sessions):
            self.connection.Children(0).createSession()
            time.sleep(TIME_WAIT)
        # Re-check number of sessions
        number_of_sessions = len(self.connection.sessions)
        if(number_of_sessions > desire_instances):
          self.create_lock()
          self.logger.warning("More instances that desire, closing excess...")
          for i in range(number_of_sessions, desire_instances, -1):
            self.connection.closeSession(self.connection.sessions[i - 1].Id)
        # Unlock instances
        self.delete_lock()
        self.logger.info("SAP Frontend is ready to receive requests.")
      except:
        pass
  def inicializar(self) -> None:
    try:
      while(self.check_lock()): time.sleep(1)
      self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
      self.connection = self.SapGui.connections[0]
      self.session = self.connection.Children(self.instancia)
    except:
      self.create_lock()
      self.inicializar()
  def relatorio(self, dia=7, filtrar_dias=False) -> str:
      tipos_de_nota = []
      danos_filtrar = []
      hoje = datetime.date.today()
      semana = hoje - datetime.timedelta(days=dia)
      janela = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
      self.session.StartTransaction(Transaction="ZSVC20")
      self.session.FindById("wnd[0]/usr/btn%_SO_QMART_%_APP_%-VALU_PUSH").Press()
      for atividade in self.ATIVIDADES:
        tipos_de_nota.extend(self.depara('relatorio_tipo', atividade).split(','))
        danos_filtrar.extend(self.depara('relatorio_filtro', atividade).split(','))
      for i in range(len(tipos_de_nota)):
        self.session.FindById(janela + f"ctxtRSCSEL_255-SLOW_I[1,{i}]").text = tipos_de_nota[i]
        self.session.FindById(janela).verticalScrollbar.position = i
      self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
      self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-LOW").text = semana.strftime("%d.%m.%Y")
      self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-HIGH").text = hoje.strftime("%d.%m.%Y")
      self.session.FindById("wnd[0]/usr/btn%_SO_USUAR_%_APP_%-VALU_PUSH").Press()
      self.session.FindById(janela + "ctxtRSCSEL_255-SLOW_I[1,0]").text = "ENVI"
      self.session.FindById(janela + "ctxtRSCSEL_255-SLOW_I[1,1]").text = "LIBE"
      self.session.FindById(janela + "ctxtRSCSEL_255-SLOW_I[1,2]").text = "TABL"
      self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
      self.session.FindById("wnd[0]/usr/ctxtSO_BEBER-LOW").text = self.REGIAO
      self.session.FindById("wnd[0]/usr/ctxtP_LAYOUT").text = self.LAYOUT
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      tabela = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell", False)
      if(tabela == None): raise Exception("404: O relatorio de notas em aberto esto vazio!")
      if(tabela.RowCount == 0): raise Exception("404: O relatorio de notas em aberto esto vazio!")
      dataframe = {
        "Cor": [],
        "Nota": [],
        "Instalacao": [],
        "Tipo": [],
        "Dano": [],
        "Data": [],
        "Hora": [],
        "Status": [],
      }
      for i in range(tabela.RowCount):
        dataframe["Cor"].append(str(self.DESTAQUE_AUSENTE))
        dataframe["Nota"].append(tabela.getCellValue(i, "QMNUM"))
        dataframe["Instalacao"].append(tabela.getCellValue(i, "ZZINSTLN"))
        dataframe["Tipo"].append(tabela.getCellValue(i, "QMART"))
        dataframe["Dano"].append(tabela.getCellValue(i, "FECOD"))
        dataframe["Data"].append(tabela.getCellValue(i, "LTRMN"))
        dataframe["Hora"].append(tabela.getCellValue(i, "LTRUR"))
        dataframe["Status"].append(tabela.getCellValue(i, "ZZ_ST_USUARIO"))
        tabela.firstVisibleRow = i
      dataframe = pandas.DataFrame(dataframe)
      dataframe["Data"] = pandas.to_datetime(dataframe["Data"], format="%d/%m/%Y", errors='coerce')
      dataframe["Hora"] = pandas.to_datetime(dataframe["Hora"], format="%H:%M:%S")
      for dano in danos_filtrar:
        dataframe = dataframe[dataframe["Dano"] != dano]
      if(filtrar_dias):
        hoje = pandas.to_datetime(datetime.date.today())
        dataframe = dataframe[dataframe["Data"] <= hoje]
      quantidade_total = len(dataframe)
      if(quantidade_total == 0): raise Exception("404: O relatorio de notas em aberto esto vazio!")
      csv = dataframe.to_csv(index=False, sep=';')
      if not (isinstance(csv, str)):
        raise Exception("500: O relatorio nao pode ser convertido em CSV!")
      return csv
  def bandeirada(self, dia=45, filtrar_dias=True) -> str:
      hoje = datetime.date.today()
      agora = datetime.datetime.now()
      semana = hoje - datetime.timedelta(days=dia)
      self.session.StartTransaction(Transaction="ZSVC20")
      self.session.FindById("wnd[0]/usr/btn%_SO_QMART_%_APP_%-VALU_PUSH").Press()
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "BA"
      self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
      hoje = datetime.date.today()
      semana = hoje - datetime.timedelta(days=dia)
      self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-LOW").text = semana.strftime("%d.%m.%Y")
      self.session.FindById("wnd[0]/usr/ctxtSO_QMDAT-HIGH").text = hoje.strftime("%d.%m.%Y")
      self.session.FindById("wnd[0]/usr/btn%_SO_FECOD_%_APP_%-VALU_PUSH").Press()
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "OSTA"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "OSJD"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "OSFT"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "OSAT"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "OSAR"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "OATI"
      self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
      self.session.FindById("wnd[0]/usr/btn%_SO_USUAR_%_APP_%-VALU_PUSH").Press()
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "ANAL"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "POSB"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "PCOM"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "NEXE"
      self.session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "EXEC"
      self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
      self.session.FindById("wnd[0]/usr/ctxtSO_BEBER-LOW").text = self.REGIAO
      self.session.FindById("wnd[0]/usr/ctxtP_LAYOUT").text = self.LAYOUT
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      tabela = self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell", False)
      if(tabela == None): raise Exception("404: O relatorio de notas em aberto esto vazio!")
      if(tabela.RowCount == 0): raise Exception("404: O relatorio de notas em aberto esto vazio!")
      dataframe = {
        "Cor": [],
        "Nota": [],
        "Instalacao": [],
        "Tipo": [],
        "Dano": [],
        "Data": [],
        "Hora": [],
        "Status": [],
        "Encerramento": [],
      }
      for i in range(tabela.RowCount):
        dataframe["Cor"].append(str(self.DESTAQUE_AUSENTE))
        dataframe["Nota"].append(tabela.getCellValue(i, "QMNUM"))
        dataframe["Instalacao"].append(tabela.getCellValue(i, "ZZINSTLN"))
        dataframe["Tipo"].append(tabela.getCellValue(i, "QMART"))
        dataframe["Dano"].append(tabela.getCellValue(i, "FECOD"))
        dataframe["Data"].append(tabela.getCellValue(i, "LTRMN"))
        dataframe["Hora"].append(tabela.getCellValue(i, "LTRUR"))
        dataframe["Status"].append(tabela.getCellValue(i, "ZZ_ST_USUARIO"))
        dataframe["Encerramento"].append(tabela.getCellValue(i, "QMDAB"))
        tabela.firstVisibleRow = i
      dataframe = pandas.DataFrame(dataframe)
      dataframe["Data"] = pandas.to_datetime(dataframe["Data"], format="%d/%m/%Y", errors='coerce')
      dataframe["Hora"] = pandas.to_datetime(dataframe["Hora"], format="%H:%M:%S")
      dataframe["Encerramento"] = pandas.to_datetime(dataframe["Encerramento"], format="%d/%m/%Y", errors='coerce')
      if(filtrar_dias):
        dataframe = dataframe[dataframe["Encerramento"].isnull()]
      quantidade_total = len(dataframe)
      if(quantidade_total == 0): raise Exception("404: O relatorio de notas em aberto esto vazio!")
      csv = dataframe.to_csv(index=False, sep=';')
      if not (isinstance(csv, str)):
        raise Exception("500: O relatorio nao pode ser convertido em CSV!")
      return csv
  def leiturista(self, nota, retry:bool=False, order_by_sequence:bool=False, interval:int=30) -> str:
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
        raise Exception("404: Nao ho relatorio de leitura para o periodo especificado!")
      if(retry or order_by_sequence):
        self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn("ZZ_NUMSEQ")
        self.session.FindById("wnd[0]/tbar[1]/btn[28]").Press()
      self.session.FindById("wnd[0]/tbar[0]/btn[71]").Press()
      self.session.FindById("wnd[1]/usr/txtGS_SEARCH-VALUE").text = instalacao
      self.session.FindById("wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
      self.session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
      # statusBar = self.session.FindById("/app/con[0]/ses[{self.instancia}]/wnd[0]/sbar").text
      # if(statusBar == "Nenhuma ocorrência encontrada"):
        # raise Exception("404: A instalacao nao foi encontrada no relatorio!")
      self.session.FindById("wnd[1]").Close()
      celula = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow
      if(celula == 0 and instalacao != int(self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(0,"ANLAGE"))):
        raise Exception("404: A instalacao nao foi encontrada no relatorio!")
      linhas = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
      # se a linhaAtual e menor que 14, a primeiraVisivel e 0 e offset e igual a linha atual
      # se a linhaAtual e maior que linhasTotais - 14, entao primeiraVisivel e linhasTotais - 28 e offset e igual a
      apontador = 0
      limite = 0
      # Se a quantidade de linhas for menor que o padrão
      if (linhas <= interval * 2):
        apontador = 0
        limite = linhas
      # Se a instalacao foi encontrada no início do relatorio
      elif (celula <= interval):
        apontador = 0
        limite = interval * 2
      # Se a instalacao foi encontrada no final do relatorio
      elif(celula > (linhas - interval)):
        apontador = linhas - (interval * 2)
        limite = linhas
      # Se a instalacao foi encontrada no meio do relatorio
      else:
        apontador = celula - interval
        limite = celula + interval
      self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = celula
      leitString = {
        "Cor": [],
        "Seq": [],
        "Instalacao": [],
        "Endereco": [],
        "Bairro": [],
        "Medidor": [],
        "Hora": [],
        "Cod": [],
      }
      texto_codigo_leiturista = ""
      while (apontador < limite):
        leitString["Cor"].append(str(self.DESTAQUE_AMARELO if(apontador == celula) else self.DESTAQUE_AUSENTE))
        leitString["Seq"].append(self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ZZ_NUMSEQ"))
        leitString["Instalacao"].append(self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ANLAGE"))
        leitString["Endereco"].append(self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ZENDERECO"))
        leitString["Bairro"].append(self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"BAIRRO"))
        leitString["Medidor"].append(self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"GERAET"))
        leitString["Hora"].append(self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ZHORALEIT"))
        cod = self.session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador,"ABLHINW")
        leitString["Cod"].append(cod)
        texto_codigo_leiturista = self.depara("leitura_codigo", cod) if(apontador == celula and cod != "") else texto_codigo_leiturista
        apontador = apontador + 1
      dataframe = pandas.DataFrame(leitString)
      if(texto_codigo_leiturista != ""):
        apontador = len(dataframe)
        dataframe.loc[apontador, "Cor"] = str(self.DESTAQUE_VERMELHO)
        dataframe.loc[apontador, "Endereco"] = texto_codigo_leiturista
      return dataframe.to_csv(index=False)
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
    if(self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110", False) == None):
      raise Exception("404: Instalação consultada não tem registro de faturas!")
    self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110").Select()
    linhas = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").RowCount
    if(linhas < 1): raise Exception("404: Cliente nao possui faturas pendentes!")
    dataframe = {
      "Cor": [],
      "Mes ref": [],
      "Vencimento": [],
      "Valor": [],
      "Tipo": [],
      "Status": [],
    }
    apontador = 1
    while (apontador < linhas):
      dataframe["Mes ref"].append(self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"BILLING_PERIOD"))
      dataframe["Vencimento"].append(self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"FAEDN"))
      dataframe["Valor"].append(self.sanitizar(self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"TOTAL_AMNT")))
      dataframe["Tipo"].append(self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"TIP_FATURA"))
      statusFat = self.session.findById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador, "STATUS")
      if(statusFat == "@5B@"): # Status no prazo de vencimento da fatura
        dataframe["Cor"].append(str(self.DESTAQUE_VERDEJANTE))
        dataframe["Status"].append("Fat. no prazo")
      elif(statusFat == "@5C@"): # Status prazo de pagamento vencido
        dataframe["Cor"].append(str(self.DESTAQUE_VERMELHO))
        dataframe["Status"].append("Fat. vencida")
      elif(statusFat == "@06@"): # Status prazo de pagamento vencido
        dataframe["Cor"].append(str(self.DESTAQUE_AMARELO))
        dataframe["Status"].append("Fat. Retida")
      else:
        dataframe["Cor"].append(str(self.DESTAQUE_AUSENTE))
        dataframe["Status"].append("Consultar")
      apontador = apontador + 1
    return pandas.DataFrame(dataframe).to_csv(index=False)
  def imprimir(self, documento) -> None:
    # command = f"taskkill /F /FI \"IMAGENAME eq SAPLPD.exe\""
    # subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    self.session.StartTransaction(Transaction="ZATC73")
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
      raise Exception(f"429: Cliente possui muitas faturas ({len(debitos)}) pendentes")
    if(len(debitos) == 0):
      raise Exception("404: Cliente nao possui faturas vencidas!")
    self.imprimir(debitos)
    return str(len(debitos))
  def instalacao(self, arg) -> int:
    try:
      arg = int(arg)
    except:
      raise Exception("400: Informacao nao e um numero valido!")
    if (arg > 999999999):
      self.session.StartTransaction(Transaction="IW53")
      self.session.FindById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = arg
      self.session.FindById("wnd[0]/tbar[1]/btn[5]").Press()
      try:
        self.session.FindById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09").Select()
      except:
        raise Exception("400: A nota informada e invalida!")
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
      self.session.FindById("wnd[0]/usr/ctxtRISA0-MATNR").text = ""
      self.session.FindById("wnd[0]/usr/ctxtRISA0-SERNR").text = arg
      self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
      if(self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell", False) != None):
        linhas = self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
        equipamentos = []
        for linha in range(linhas):
          equipamentos.append(self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(linha, "MATNR"))
        lista_informacoes = "####################\n"
        for equipamento in equipamentos:
          lista_informacoes = lista_informacoes + self.info_medidor(str(arg), str(equipamento)) + "\n####################\n"
        raise Exception(f"*409: Ha mais de um equipamento com esse mesmo numero de serie! Verifique qual lhe atende e solicite pela instalacao!*\n\n{lista_informacoes}")
      try:
        self.session.FindById(r'wnd[0]/usr/tabsTABSTRIP/tabpT\03/ssubSUB_DATA:SAPMIEQ0:0500/subISUSUB:SAPLE10R:1000/btnBUTTON_ISABL').Press()
        instalacao = self.session.findById("wnd[0]/usr/txtIEANL-ANLAGE").text
        self.instalacao(instalacao)
        return instalacao
      except:
        raise Exception("400: O numero informado nao eh nota, instalacao ou medidor")
    raise Exception("400: O numero informado nao eh nota, instalacao ou medidor")
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
    historico = {
      "Cor": [],
      "Nota": [],
      "Tipo": [],
      "Texto breve para dano": [],
      "Texto breve para code": [],
      "Status": [],
      "Data": [],
    }
    while(apontador < linhas and apontador < 10):
      historico["Cor"].append(0)
      historico["Nota"].append(self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador,"QMNUM"))
      historico["Tipo"].append(self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador, "QMART"))
      historico["Texto breve para dano"].append(self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador, "KURZTEXT"))
      historico["Texto breve para code"].append(self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador,"MATXT"))
      historico["Data"].append(self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador,"AUSBS"))
      historico["Status"].append(self.session.FindById("wnd[0]/usr/cntlCONTAINER_100/shellcont/shell").getCellValue(apontador,"ZZ_ST_USUARIO"))
      apontador = apontador + 1
    return pandas.DataFrame(historico).to_csv(index=False)
  def agrupamento(self, nota, have_authorization: bool = True, debitos: bool = False) -> str:
    instalacao = self.instalacao(nota)
    self.session.StartTransaction(Transaction="ES32")
    self.session.FindById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = instalacao
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    if("ES61" in self.NOTUSE):
      self.session.FindById("wnd[0]/usr/ctxtEANLD-VSTELLE").setFocus()
      self.session.FindById("wnd[0]").SendVKey(2)
    else:
      consumo = self.session.FindById("wnd[0]/usr/ctxtEANLD-VSTELLE").text
      self.session.StartTransaction(Transaction="ES61")
      self.session.findById("wnd[0]/usr/ctxtEVBSD-VSTELLE").text = consumo
      self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    if("ES57" in self.NOTUSE):
      self.session.FindById("wnd[0]/usr/ctxtEVBSD-HAUS").setFocus()
      self.session.FindById("wnd[0]").SendVKey(2)
    else:
      ligacao = self.session.FindById("wnd[0]/usr/ctxtEVBSD-HAUS").text
      self.session.StartTransaction(Transaction="ES57")
      self.session.FindById("wnd[0]/usr/ctxtEHAUD-HAUS").text = ligacao
      self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    # Gets street number and verify if is a valid number
    numero = self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-HOUSE_NUM1").text
    match_number = re.search("[0-9]+", numero)
    if not (match_number):
      raise Exception("422: Instalacao sem numero de rua! O agrupamento nao pode ser analisado automaticamente.")
    numero_sem_letra = int(match_number.group())
    if (numero == "1SN" or numero == "SN"):
      raise Exception("422: Instalacao sem numero de rua! O agrupamento nao pode ser analisado automaticamente.")
    if(numero_sem_letra == None):
      raise Exception("422: Instalacao sem numero de rua! O agrupamento nao pode ser analisado automaticamente.")
    # Gets 'logradouro' and go to 'logradouro' aplication details
    logradouro = self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME_CO").text
    if("ZMED95" in self.NOTUSE):
      self.session.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME_CO").setFocus()
      self.session.FindById("wnd[0]").SendVKey(2)
    else:
      self.session.StartTransaction(Transaction="ZMED95")
      self.session.FindById("wnd[0]/usr/ctxtADRSTREET-STRT_CODE").text = logradouro
      self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    # Go to list of street numbers aplication 
    self.session.FindById("wnd[0]/tbar[1]/btn[9]").Press()
    linhas = self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").RowCount
    tamanho_maximo = self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").rows.length
    apontador = 0
    dataframe = {
      "Cor": [],
      "Endereco": [],
      "Instalacao": [],
      "Medidor": [],
      "Cliente": [],
      "Status": [],
      "Tipo": [],
      "Montante": [],
      "Observacao": [],
    }
    # Colecting information about clients in the same street number that client street number
    while (apontador < linhas):
      num10_com_letra = self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX/txtTI_NUMSX-NUMERO[0,{tamanho_maximo - 1}]").text
      match = re.search("[0-9]+", num10_com_letra)
      num10_sem_letra = int(match.group()) if (match != None) else 999999
      if (num10_sem_letra < numero_sem_letra):
        apontador = apontador + tamanho_maximo
        self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").verticalScrollbar.position = apontador
        continue
      num_atual_texto = self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX/txtTI_NUMSX-NUMERO[0,0]").text
      match = re.search("[0-9]+", num_atual_texto)
      if(match == None):
        apontador = apontador + 1
        self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").verticalScrollbar.position = apontador
        continue
      num_atual = int(match.group())
      if(num_atual > numero_sem_letra): break
      if(num_atual == numero_sem_letra):
        quantidade = int(self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX/txtTI_NUMSX-QTD_INSTAL[1,0]").text)
        self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").verticalScrollbar.position = apontador
        self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").GetAbsoluteRow(apontador).selected = True
        self.session.FindById("wnd[0]/usr/btn%#AUTOTEXT005").Press()
        for i in range(1, quantidade + 1):
          complemento = self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX/txtTI_INSTALX-COMPLS[0,0]").text
          dataframe["Endereco"].append(f"{num_atual_texto} {complemento}")
          dataframe["Instalacao"].append(self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX/txtTI_INSTALX-ANLAGE[1,0]").text)
          dataframe["Cliente"].append(self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX/txtTI_INSTALX-NOME[2,0]").text)
          dataframe["Tipo"].append(self.session.FindById(f"wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX/txtTI_INSTALX-CLASSE[3,0]").text)
          self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_INSTALX").verticalScrollbar.position = i
      apontador = apontador + 1
      self.session.FindById("wnd[0]/usr/tblSAPLZMED_ENDERECOSTC_NUMSX").verticalScrollbar.position = apontador
    # Checking if number of clients is excessive
    quantidade = len(dataframe['Instalacao'])
    tolerancia = 12 if debitos == False else 24
    if(quantidade > tolerancia):
      raise Exception(f"429: Agrupamento possui instalacoes demais ({quantidade})")
    apontador = 0
    # Coleta da situacao das instalacões
    for apontador in range(quantidade):
      instalacao_corrente = dataframe["Instalacao"][apontador]
      self.instalacao(instalacao_corrente)
      dataframe["Status"].append(self.session.findById("wnd[0]/usr/txtEANLD-DISCSTAT").text)
      if(debitos):
        try:
          csv_data = self.escrever(instalacao_corrente)
          debitos_tabela = pandas.read_csv(io.StringIO(csv_data))
          debitos_tabela = debitos_tabela[debitos_tabela['Cor'] == self.DESTAQUE_VERMELHO]
          debitos_tabela['Valor'] = pandas.to_numeric(debitos_tabela['Valor'], 'coerce').astype(float)
          dataframe["Montante"].append(debitos_tabela['Valor'].sum())
        except:
          dataframe["Montante"].append(0)
        dataframe["Observacao"].append("")
        dataframe["Cor"].append(self.DESTAQUE_AUSENTE)
        apontador = apontador + 1
        continue
      if(instalacao_corrente == instalacao):
        dataframe["Observacao"].append("Instalacao da nota")
        dataframe["Cor"].append(self.DESTAQUE_VERMELHO)
        apontador = apontador + 1
        continue
      if(self.session.findById("wnd[0]/usr/txtEANLD-VERTRAG").text == ""):
        dataframe["Observacao"].append("Sem contrato ativo")
        dataframe["Cor"].append(self.DESTAQUE_VERMELHO)
        apontador = apontador + 1
        continue
      if(dataframe["Status"][apontador] == " Instalação complet.suspensa"):
        dataframe["Observacao"].append("Suspensa no sistema")
        dataframe["Cor"].append(self.DESTAQUE_VERMELHO)
        apontador = apontador + 1
        continue
      if(dataframe["Status"][apontador] == "Supensao iniciada"):
        dataframe["Observacao"].append("Tem ordem de corte")
        dataframe["Cor"].append(self.DESTAQUE_VERMELHO)
        apontador = apontador + 1
        continue
      if(have_authorization): temp = self.novo_analisar(instalacao_corrente)
      else: temp = self.passivas_novo(instalacao_corrente)
      if(temp):
        dataframe["Observacao"].append("Tem contas passivas")
        dataframe["Cor"].append(self.DESTAQUE_VERMELHO)
        apontador = apontador + 1
        continue
      # caso nao encontre nenhum impedimento
      dataframe["Observacao"].append("Cliente nao passivel")
      dataframe["Cor"].append(self.DESTAQUE_VERDEJANTE)
      apontador = apontador + 1
    apontador = 0
    # Preparacao da string final
    del dataframe["Medidor"]
    if not (debitos):
      del dataframe["Montante"]
    else:
      del dataframe["Observacao"]
    dataframe = pandas.DataFrame(dataframe)
    return dataframe.to_csv(index=False)
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
      return re.sub(',', '.', coordenada)
    else:
      raise Exception("404: A instalacao nao possui coordenada cadastrada!")
  def telefone(self, arg) -> str:
    instalacao = self.instalacao(arg)
    parceiro = self.session.findById("wnd[0]/usr/txtEANLD-PARTNER").text
    cliente_nome = self.session.findById("wnd[0]/usr/txtEANLD-PARTTEXT").text
    if(str(cliente_nome) == ""): raise Exception("404: Instalacao sem cliente! Sem telefone!")
    if(str(cliente_nome).startswith("UNIDADE C/ CONSUMO")): raise Exception("404: Cliente ficticio! Sem telefone!")
    if(str(cliente_nome).startswith("PARCEIRO DE NEGOCIO")): raise Exception("404: Cliente ficticio! Sem telefone!")
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
    espaco_vazio = "______________________________"
    try:
      telefone.remove(espaco_vazio)
    except:
      pass
    if(len(telefone) == 0):
      return "404: " + nome_cliente + " NAO TEM NUMERO DE TELEFONE CADASTRADO!"
    else:
      return nome_cliente + ' '.join(telefone)
  def consumo(self, nota) -> str:
    instalacao = self.instalacao(nota)
    self.session.StartTransaction(Transaction="ZATC66")
    self.session.FindById("wnd[0]/usr/ctxtP_ANLAGE").text = instalacao
    self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    try:
      self.session.FindById("wnd[0]/usr/subSUB1:SAPLZATC_INFO_CRM:0900/radXSCREEN-HEADER-RB_LEIT").Select()
    except:
      raise Exception(f"404: A instalacao {instalacao} nao possui historico de consumo para o contrato atual.")
    tabela = "wnd[0]/usr/cntlCONTROL/shellcont/shell"
    linhas = self.session.FindById(tabela).RowCount
    dataframe = {
      "Cor": [],
      "Mes ref.": [],
      "Data leit.": [],
      "Medidor": [],
      "Leitura": [],
      "Consumo": [],
      "Registrador": [],
      "Tipo de leitura": [],
      "Motivo da leitura": [],
      "Nota do leiturista": [],
    }
    apontador = 0
    while(apontador < linhas):
      dataframe["Cor"].append(self.DESTAQUE_AUSENTE)
      dataframe["Mes ref."].append(self.session.FindById(tabela).getCellValue(apontador, "MES_ANO"))
      dataframe["Data leit."].append(self.session.FindById(tabela).getCellValue(apontador, "ADATSOLL"))
      dataframe["Medidor"].append(int(self.session.FindById(tabela).getCellValue(apontador, "GERNR")))
      dataframe["Leitura"].append(int(self.sanitizar(self.session.FindById(tabela).getCellValue(apontador, "LEIT_FATURADA"))))
      dataframe["Consumo"].append(0)
      # Código do registrador e texto breve descritivo
      registrador = self.session.FindById(tabela).getCellValue(apontador, "ZWNUMMER")
      if (registrador != ""):
        registrador = "0" + str(registrador) if len(registrador) == 1 else registrador
        texto_registrador = self.depara("medidor_registrador", registrador)
        dataframe["Registrador"].append(f"{registrador} - {texto_registrador}")
      else:
        dataframe["Registrador"].append("00 - Sem codigo do registrador")
      # Código do leiturista e texto breve descritivo
      codigo_leiturista = self.session.FindById(tabela).getCellValue(apontador, "OCORRENCIA")
      texto_codigo_leiturista = codigo_leiturista + " - " + self.depara("leitura_codigo", codigo_leiturista) if codigo_leiturista else ""
      dataframe["Nota do leiturista"].append(texto_codigo_leiturista)
      # Código do tipo de leitura e texto breve descritivo
      tipo_leitura = self.session.FindById(tabela).getCellValue(apontador, "TIPO_LEITURA")
      texto_tipo_leitura = self.depara("leitura_tipo", tipo_leitura)
      dataframe["Tipo de leitura"].append(f"{tipo_leitura} - {texto_tipo_leitura}")
      # Código do motivo da leitura e texto breve descritivo
      motivo_leitura = self.session.FindById(tabela).getCellValue(apontador, "MOTIVO_LEITURA")
      texto_motivo_leitura = self.depara("leitura_motivo", motivo_leitura)
      dataframe["Motivo da leitura"].append(f"{motivo_leitura} - {texto_motivo_leitura}")
      apontador = apontador + 1
    dataframe = pandas.DataFrame(dataframe)
    leitura_anterior = dataframe["Leitura"].shift(-1)
    consumo = dataframe["Leitura"] - leitura_anterior
    dataframe["Consumo"] = consumo
    return dataframe.to_csv(index=False, sep=',')
  def analisar(self, apontador=0, verificar_15_dias=False) -> bool:
    if(apontador == 0): return False
    if (self.session.findById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador, "STATUS") != "@5C@"): return False
    if(verificar_15_dias):
      vencimento = self.session.findById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(apontador,"FAEDN")
      vencimento = datetime.datetime.strptime(vencimento, f"%d.%m.%Y").date()
      prazo_mais_15_dias = vencimento + datetime.timedelta(days=15)
      if (datetime.date.today() > prazo_mais_15_dias): return False
    return True
  def info_medidor(self, medidor_serial = "", medidor_codigo = "") -> str:
    texto_retorno = []
    informar_instalacao = medidor_serial != ""
    if(medidor_serial == ""):
      txtCodMedidor = None
      self.session.FindById("wnd[0]/usr/btnEANLD-DEVSBUT").Press()
      if(self.session.FindById("wnd[1]", False) != None):
        return "404: *INSTALACAO SEM MEDIDOR VINCULADO!*"
        dataRetirado = self.session.FindById("wnd[1]/usr/tblSAPLET03UTS_TC/txtPERIODS-BIS[1,0]").text
        self.session.FindById("wnd[1]").SendVKey(2)
      if(self.session.FindById("wnd[0]/usr/tblSAPLEG70TC_DEVRATE_C", False) == None):
        return "404: *INSTALACAO SEM MEDIDOR VINCULADO!*"
      medidor_codigo = self.session.FindById("wnd[0]/usr/tblSAPLEG70TC_DEVRATE_C/ctxtREG70_D-MATNR[8,0]").text
      medidor_serial = self.session.FindById("wnd[0]/usr/tblSAPLEG70TC_DEVRATE_C/ctxtREG70_D-GERAET[0,0]").text
    txtCodMedidor = self.depara("material_codigo", medidor_codigo)
    self.session.StartTransaction(Transaction="IQ03")
    self.session.FindById("wnd[0]/usr/ctxtRISA0-MATNR").text = medidor_codigo
    self.session.FindById("wnd[0]/usr/ctxtRISA0-SERNR").text = medidor_serial
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    code_montagem_medidor = self.session.FindById("wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/subSUB_0152C:SAPLITO0:1526/txtITOBATTR-STTXT").text
    code_status_medidor = self.session.FindById("wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/subSUB_0152C:SAPLITO0:1526/txtITOBATTR-STTXU").text
    texto_montagem_medidor = code_montagem_medidor + ' - ' + self.depara('medidor_montagem', code_montagem_medidor)
    texto_status_medidor = code_status_medidor + ' - ' + self.depara('medidor_status', code_status_medidor)
    texto_retorno.append(f"*Equipamento:* {medidor_serial}")
    texto_retorno.append(f"*Tipo:* {txtCodMedidor}")
    texto_retorno.append(f"*Montagem do equipamento:* {texto_montagem_medidor}")
    texto_retorno.append(f"*Status do equipamento:* {texto_status_medidor}")
    self.session.FindById(r'wnd[0]/usr/tabsTABSTRIP/tabpT\03/ssubSUB_DATA:SAPMIEQ0:0500/subISUSUB:SAPLE10R:1000/btnBUTTON_ISABL').Press()
    if(self.session.FindById("wnd[0]/usr/cntlBCALVC_EVENT2_D100_C1/shellcont/shell", False) == None):
      if(informar_instalacao): texto_retorno.append(f"*Instalacao:* `SEM ACESSO A INSTALACAO!`")
      return '\n'.join(texto_retorno)
    instalacao = self.session.FindById("wnd[0]/usr/txtIEANL-ANLAGE").text
    if(informar_instalacao):
      texto_retorno.append(f"*Instalacao:* `{instalacao}`")
      self.instalacao(instalacao)
      return '\n'.join(texto_retorno) + "\n" + self.info_instalacao()
    apontador = 0
    linhas = self.session.FindById("wnd[0]/usr/cntlBCALVC_EVENT2_D100_C1/shellcont/shell").RowCount
    limite = 12 if linhas > 12 else linhas
    dataframe = {
      "Data": [],
      "Codigo": [],
      "Descricao": [],
    }
    while(apontador < limite):
      data = self.session.FindById("wnd[0]/usr/cntlBCALVC_EVENT2_D100_C1/shellcont/shell").getCellValue(apontador,"ADATSOLL")
      status = self.session.FindById("wnd[0]/usr/cntlBCALVC_EVENT2_D100_C1/shellcont/shell").getCellValue(apontador,"ABLHINW")
      if(status != ""):
        dataframe['Data'].append(data)
        dataframe['Codigo'].append(status)
        texto_status = self.depara('leitura_codigo', status)
        dataframe['Descricao'].append(texto_status)
      apontador = apontador + 1
    dataframe = pandas.DataFrame(dataframe)
    if(len(dataframe) > 0):
      texto_retorno.append(f"*Codigos de leitura nas ultimas {limite} leituras:*")
      texto_retorno.append(dataframe.to_string(index=False))
    else:
      texto_retorno.append(f"Sem codigos de leitura nas ultimas {limite} leituras!")
    return '\n'.join(texto_retorno)
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
      raise Exception("404: Nao ha faturas passiveis!")
    linhas = self.session.FindById(r"wnd[0]/usr/tabsTAB_STRIP_100/tabpF190/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0190/cntlCONTAINER_190/shellcont/shell").RowCount
    if(linhas == 0): raise Exception("404: Nao ha faturas passiveis!")
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
      raise Exception(f"429: Cliente possui muitas faturas ({len(passiveis)}) passivas")
    self.imprimir(passiveis)
    return str(len(passiveis))
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
      raise Exception("404: Cliente nao possui faturas pendentes!")
    #[char, line]
    col = 1
    row = 0
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
    return dt3.to_csv(index = False)
  def fatura_novo(self, arg) -> str:
    debitos = self.escrever_novo(arg, True)
    if(len(debitos) > 6): raise Exception(f"429: Cliente possui muitas faturas ({len(debitos)}) pendentes")
    self.imprimir(debitos)
    return str(len(debitos))
  def passivas_novo(self, arg) -> bool:
    debitos = self.escrever_novo(arg, False, True)
    return (len(debitos) > 0)
  def info_parceiro(self) -> str:
    parceiro = self.session.findById("wnd[0]/usr/txtEANLD-PARTNER").text
    nome_cliente = self.session.findById("wnd[0]/usr/txtEANLD-PARTTEXT").text
    if(len(parceiro) == 0): return "404: *INSTALACAO SEM CLIENTE VINCULADO!*"
    if(str(nome_cliente).startswith("UNIDADE C/ CONSUMO")): return "404: *INSTALACAO SEM CLIENTE VINCULADO!*"
    if(str(nome_cliente).startswith("PARCEIRO DE NEGOCIO")): return "404: *INSTALACAO SEM CLIENTE VINCULADO!*"
    nome_cliente = str(nome_cliente).split("/")[0]
    phone_field_partial_string = self.parceiro(parceiro)
    self.session.findById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_04").Select()
    pessoa_fisica = self.session.findById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_04/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7006/subA04P01:SAPLBUPA_BUTX_DIALOG:0100/tblSAPLBUPA_BUTX_DIALOGTCTRL_BPTAX/txtDFKKBPTAXNUM-TAXNUM[2,0]").text
    return f"*Cod. do cliente:* {parceiro}\n*Cadastro Pessoa Fisica (CPF):* {pessoa_fisica}\n*Nome do cliente:* {nome_cliente}"
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
      # 5. "NUMERO_OUTR"
      dataframe['Num. 2'].append(container.getCellValue(apontador, "NUMERO_OUTR"))
      # 6. "LATITUDE"
      dataframe['Latitude'].append(str.replace(container.getCellValue(apontador, "LATITUDE"), ',', '.'))
      # 7. "LONGITUDE"
      dataframe['Longitude'].append(str.replace(container.getCellValue(apontador, "LONGITUDE"), ',', '.'))
      # 8. "CITY_NAME"
      dataframe['Cidade'].append(container.getCellValue(apontador, "CITY_NAME"))
      apontador = apontador + 1
    return pandas.DataFrame(dataframe).to_csv(index=False, sep=',')
  def depara(self, tipo: str, de: str) -> str:
    try:
      filename = os.path.join(self.BASE_FOLDER, 'sap.db')
      connection = sqlite3.connect(filename)
      cursor = connection.execute(f"SELECT para FROM depara WHERE tipo = '{tipo}' AND de = '{de}'")
      result = cursor.fetchone()
      return result[0] if result else "Codigo desconhecido!"
    except Exception as e:
      return f"An error occurred: {e}"
  def retorno(self) -> None:
    self.session.sendCommand('/n')
  def inspecao(self, arg) -> str:
    instalacao = self.instalacao(arg)
    retorno = f"401: A instalacao {instalacao} nao esta apta para abertura de nota de recuperacao devido "
    # Collecting installation information
    statusInstalacao = self.session.findById('wnd[0]/usr/txtEANLD-DISCSTAT').text
    if(statusInstalacao != ' Instalação não suspensa'): return retorno + "nao estar ativa!"
    parceiro = self.session.findById("wnd[0]/usr/txtEANLD-PARTNER").text
    cliente = str(self.session.FindById("wnd[0]/usr/txtEANLD-PARTTEXT").text)
    if(cliente == ""): return retorno + "nao ter cliente vinculado"
    if(str(cliente).startswith("UNIDADE C/ CONSUMO")): return retorno + "nao ter cliente vinculado"
    if(str(cliente).startswith("PARCEIRO DE NEGOCIO")): return retorno + "nao ter cliente vinculado"
    consumo = self.session.FindById("wnd[0]/usr/ctxtEANLD-VSTELLE").text
    classe = self.session.FindById("wnd[0]/usr/tblSAPLES30TC_TIMESL/ctxtEANLD-ISTYPE[5,0]").text
    is_residencial = int(classe) > 1000 and int(classe) < 2000
    subclasse = self.depara('classe_subclasse', classe)
    is_baixa_renda = is_residencial and subclasse.find('Baixa Renda') >= 0
    if(is_baixa_renda): return retorno + "devido instalacao ser baixa renda"
    unidade = self.session.FindById("wnd[0]/usr/tblSAPLES30TC_TIMESL/ctxtEANLD-ABLEINH[9,0]").text
    localidade = unidade[2:6:1]
    # Collecting measurement information
    try:
      self.session.FindById("wnd[0]/usr/btnEANLD-DEVSBUT").Press()
    except:
      return retorno + " nao tem medidor vinculado"
    medidor = self.session.FindById("wnd[0]/usr/tblSAPLEG70TC_DEVRATE_C/ctxtREG70_D-GERAET[0,0]").text
    self.session.StartTransaction(Transaction="IQ03")
    self.session.FindById("wnd[0]/usr/ctxtRISA0-MATNR").text = ""
    self.session.FindById("wnd[0]/usr/ctxtRISA0-SERNR").text = medidor
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    statusMedidor = self.session.findById("wnd[0]/usr/subSUB_EQKO:SAPLITO0:0152/subSUB_0152C:SAPLITO0:1526/txtITOBATTR-STTXU").text
    if(statusMedidor != "INST"): return retorno + f"ao medidor {medidor} estar com status {statusMedidor}!"
    # Collecting installation type information
    self.session.StartTransaction(Transaction="ES61")
    self.session.FindById("wnd[0]/usr/ctxtEVBSD-VSTELLE").text = consumo
    self.session.FindById("wnd[0]/tbar[0]/btn[0]").Press()
    tipoInstalacao = int(self.session.FindById("wnd[0]/usr/ssubSUB:SAPLXES60:0100/tabsTS0100/tabpTAB1/ssubSUB1:SAPLXES60:0101/ctxtEVBSD-ZZ_TP_LIGACAO").text)
    # if(not self.is_passivel_ren(tipoInstalacao, is_residencial, localidade)): return retorno + "ser residencial em area restrita de inspecao pelo tipo de instalacao"
    # Collecting customer registration information
    phone_field_partial_string = self.parceiro(parceiro)
    self.session.findById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_04").Select()
    tipo_documento = self.session.findById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_04/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7006/subA04P01:SAPLBUPA_BUTX_DIALOG:0100/tblSAPLBUPA_BUTX_DIALOGTCTRL_BPTAX/txtTFKTAXNUMTYPE_T-TEXT[1,0]").text
    if(tipo_documento != "Brasil: nº CPF" and tipo_documento != "Brasil: nº CNPJ"): return retorno + "ao cliente nao tem CPF ou CNPJ no cadastro"
    pessoa_fisica = self.session.findById(phone_field_partial_string + "ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_04/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7006/subA04P01:SAPLBUPA_BUTX_DIALOG:0100/tblSAPLBUPA_BUTX_DIALOGTCTRL_BPTAX/txtDFKKBPTAXNUM-TAXNUM[2,0]").text
    if(pessoa_fisica == ""): return retorno + "ao cliente nao tem CPF ou CNPJ no cadastro"
    # Collecting information on outstanding debts
    # debitos = pandas.read_csv(io.StringIO(self.escrever(instalacao)))
    # debitos = debitos[debitos["Cor"] != str(self.DESTAQUE_VERMELHO)]
    # if(len(debitos) > 0): return retorno + "o cliente possuir debito(s) pendente(s)!"
    # Collecting service history information
    meses_verificacao_inspecoes =  6  # if not is_residencial else 12
    historico = pandas.read_csv(io.StringIO(self.historico(instalacao)))
    historico["Data"] = pandas.to_datetime(historico['Data'], format="%d.%m.%Y")
    prazo_maximo = datetime.date.today() - datetime.timedelta(days=meses_verificacao_inspecoes * 30)
    historico = historico[historico["Data"] >= pandas.to_datetime(prazo_maximo)]
    historico = historico[(historico["Tipo"] == "BI") | (historico["Tipo"] == "BU")]
    historico = historico[historico["Status"] == "EXEC"]
    if(len(historico) > 0): return retorno + f"ja possuir nota {historico['Nota'].to_string(index=False)} de recuperacao executada!"
    return f"A instalacao {instalacao} esta apta sim para abertura de nota de recuperacao!"
  def is_passivel_ren(self, fases_instalacao:int, residencial:bool, localidade:str) -> bool:
    if(fases_instalacao == 3): return True
    if(residencial == False): return True
    if(localidade == 'L539'): return True
    if(localidade == 'L595'): return True
    return False
  def procurar(self, arg) -> str:
    dataframe = {
      'Cor': [],
      'Instalacao': [],
      'Endereco': [],
      'Medidor': [],
      ' 01 ': [],
      ' 02 ': [],
      ' 03 ': [],
      ' 04 ': [],
      ' 05 ': [],
      ' 06 ': [],
      ' 07 ': [],
      ' 08 ': [],
      ' 09 ': [],
      ' 10 ': [],
      ' 11 ': [],
      ' 12 ': [],
      # 'Passivel': [],
    }
    instalacao = self.instalacao(arg)
    leiturista = self.leiturista(instalacao, False, False, 10)
    leiturista = pandas.read_csv(io.StringIO(leiturista))
    leiturista['Instalacao'] = pandas.to_numeric(leiturista['Instalacao'], 'coerce').astype('Int64')
    leiturista['Medidor'] = pandas.to_numeric(leiturista['Medidor'], 'coerce').astype('Int64')
    leiturista = leiturista[leiturista['Instalacao'].notna()]
    dataframe['Endereco'].extend(leiturista['Endereco'].to_list())
    dataframe['Instalacao'].extend(leiturista['Instalacao'].to_list())
    dataframe['Medidor'].extend(leiturista['Medidor'].to_list())
    meses_de_referencia = []
    hoje = datetime.datetime.today()
    ano = hoje.year
    mes = hoje.month
    for i in range(1, 13):
      mes_atual = mes - i if ((mes - i) > 0) else  (mes - i + 12)
      ano_atual = ano if ((mes - i) > 0) else ano - 1
      # Calcula o primeiro dia do mês atual menos i meses
      data = datetime.datetime(year=ano_atual, month= mes_atual, day=1)
      # Adiciona a data à lista
      meses_de_referencia.append(data)
    for instalacao_atual in dataframe['Instalacao']:
      # Retirada análise de passividade devido ao nocaute de solicitações
      # passivel = self.inspecao(instalacao_atual)
      # if(passivel.startswith(f'A instalacao {instalacao_atual} nao')):
      #   dataframe['Cor'].append(str(self.DESTAQUE_VERMELHO))
      #   dataframe['Passivel'].append('Nao')
      # else:
      #   dataframe['Cor'].append(str(self.DESTAQUE_VERDEJANTE))
      #   dataframe['Passivel'].append('Sim')
      cor_destaque = self.DESTAQUE_AMARELO if(instalacao_atual == instalacao) else self.DESTAQUE_AUSENTE
      dataframe['Cor'].append(cor_destaque)
      try:
        consumos = self.consumo(instalacao_atual)
        consumos = pandas.read_csv(io.StringIO(consumos))
        consumos['Mes ref.'] = pandas.to_datetime(consumos['Mes ref.'], format='%m/%Y')
        consumos = consumos[consumos['Mes ref.'].notna()]
        consumos = consumos[consumos['Motivo da leitura'] == '01 - Leitura periódica']
        for i in range(1, 13):
          mes_indice = f" 0{i} " if i < 10 else f" {i} "
          try:
            consumo = consumos[consumos['Mes ref.'] == meses_de_referencia[i-1]]['Consumo'].values[0]
            dataframe[mes_indice].append(consumo)
          except:
            dataframe[mes_indice].append(pandas.NA)
      except:
        for i in range(1, 13):
          mes_indice = f" 0{i} " if i < 10 else f" {i} "
          dataframe[mes_indice].append(pandas.NA)
    dataframe = pandas.DataFrame(dataframe)
    dataframe['Cor'] = pandas.to_numeric(dataframe['Cor'], 'coerce').astype('Int64')
    dataframe['Instalacao'] = pandas.to_numeric(dataframe['Instalacao'], 'coerce').astype('Int64')
    dataframe['Medidor'] = pandas.to_numeric(dataframe['Medidor'], 'coerce').astype('Int64')
    for i in range(1, 13):
      mes_indice = f" 0{i} " if i < 10 else f" {i} "
      dataframe[mes_indice] = pandas.to_numeric(dataframe[mes_indice], 'coerce').astype('Int64')
    return dataframe.to_csv(index=False)
  def info_instalacao(self) -> str:
    instalacao_status = self.session.findById('wnd[0]/usr/txtEANLD-DISCSTAT').text
    endereco = str.split(self.session.FindById("wnd[0]/usr/txtEANLD-LINE1").text, ",")[1]
    classe = self.session.FindById("wnd[0]/usr/tblSAPLES30TC_TIMESL/ctxtEANLD-ISTYPE[5,0]").text
    texto_classe = self.depara('classe_subclasse', classe)
    return f"*Status Instalacao:* {instalacao_status}\n*Classe da instalacao:* {texto_classe}\n*Endereco:* {endereco}"
  def novo_informacao(self, arg) -> str:
    inst = self.instalacao(arg)
    info_instalacao = self.info_instalacao()
    self.instalacao(inst)
    info_cliente = self.info_parceiro()
    self.instalacao(inst)
    info_medicao = self.info_medidor()
    return f"*Instalacao:* {inst}\n{info_instalacao}\n{info_cliente}\n{info_medicao}"
  def codbarra(self, arg, tel) -> str:
    instalacao = self.instalacao(arg)
    parceiro = self.session.findById("wnd[0]/usr/txtEANLD-PARTNER").text
    self.session.StartTransaction(Transaction="ZATC45")
    self.session.FindById("wnd[0]/usr/radP2VIA").Select()
    self.session.findById("wnd[0]/usr/ctxtPPARTNER").text = parceiro
    self.session.findById("wnd[0]/usr/ctxtPANLAGE").text = instalacao
    self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    statusBar = self.session.FindById("wnd[0]/sbar").text
    if(statusBar != ''): raise Exception(f"400: {statusBar}")
    self.session.FindById("wnd[0]/usr/radCOD_BARRAS").Select()
    linhas = int(self.session.FindById("wnd[0]/usr/txtZATCE_MENGE_BETRW-MENGE").text)
    if(linhas > 10): raise Exception(f"429: Cliente possui muitas faturas ({linhas})")
    for apontador in range(linhas):
      situacao = self.session.FindById(f"wnd[0]/usr/tblSAPLZCRM_METODOSTC_FATURAS/txtIT_SAIDA-STATUS[7,{apontador}]").text
      if(situacao != "A Vencer" and situacao != "Retida"):
        self.session.FindById(f"wnd[0]/usr/tblSAPLZCRM_METODOSTC_FATURAS/chkIT_SAIDA-SELFAT[3,{apontador}]").selected = True
    self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    statusBar = self.session.FindById("wnd[0]/sbar").text
    if(statusBar != ''): raise Exception(f"400: {statusBar}")
    self.session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select()
    self.session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
    self.session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
    linhas = self.session.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").RowCount
    for apontador in range(linhas):
      telefone = self.session.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").getCellValue(apontador, "CELULAR")
      if(telefone == "NENHUM DOS ANTERIORES"):
        self.session.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").setCurrentCell(apontador, "CELULAR")
        self.session.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").clickCurrentCell()
    self.session.FindById("wnd[1]/usr/txtSPOP-VARVALUE1").text = tel
    self.session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
    self.session.FindById("wnd[1]/usr/btnBUTTON_1").Press()
    self.session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
    return "Solicitado envio do codigo de barras!"
  def ZATC45(self, instalacao:int, parceiro:int, documentos:list=[]) -> None:
    self.session.StartTransaction(Transaction="ZATC45")
    self.session.FindById("wnd[0]/usr/radP2VIA").Select()
    self.session.findById("wnd[0]/usr/ctxtPPARTNER").text = parceiro
    self.session.findById("wnd[0]/usr/ctxtPANLAGE").text = instalacao
    self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    statusBar = self.session.FindById("wnd[0]/sbar").text
    if(statusBar != '' and statusBar != 'Nenhum débito foi encontrado!'): raise Exception(f"400: {statusBar}")
    linhas = float(str(self.session.FindById("wnd[0]/usr/txtZATCE_MENGE_BETRW-MENGE").text).replace(',','.'))
    if(linhas != len(documentos)):
      raise Exception("401: A quantidade de faturas nao bate com o ZARC140!")
    indices = []
    for apontador in range(int(linhas)):
      documento = self.session.FindById(f"wnd[0]/usr/tblSAPLZCRM_METODOSTC_FATURAS/txtIT_SAIDA-ZIMPRES[8,{apontador}]").text
      if(documento in documentos): indices.append(apontador)
    if(len(indices) != len(documentos)):
      raise Exception("401: A quantidade de faturas nao bate com o ZARC140!")
    self.session.FindById("wnd[0]/usr/rad2VIA").Select()
    for i in range(len(indices)):
      if(i > 0):
        self.session.FindById(f"wnd[0]/usr/tblSAPLZCRM_METODOSTC_FATURAS/chkIT_SAIDA-SELFAT[3,{indices[i - 1]}]").selected = False
      self.session.FindById(f"wnd[0]/usr/tblSAPLZCRM_METODOSTC_FATURAS/chkIT_SAIDA-SELFAT[3,{indices[i]}]").selected = True
      self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
      justificativa = self.session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]", False)
      if(justificativa != None):
        justificativa.Select()
        self.session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
      if(self.session.FindById("wnd[1]", False) != None):
        self.session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
    statusBar = self.session.FindById("wnd[0]/sbar").text
    if(statusBar != ''): raise Exception(f"400: {statusBar}")
  def fatura_ZATC45(self, arg) -> str:
    debitos = []
    instalacao = self.instalacao(arg)
    parceiro = int(self.session.findById("wnd[0]/usr/txtEANLD-PARTNER").text)
    self.debito(instalacao)
    linhas = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").RowCount
    for linha in range(linhas):
      documento = self.session.FindById("wnd[0]/usr/tabsTAB_STRIP_100/tabpF110/ssubSUB_100:SAPLZARC_DEBITOS_CCS_V2:0110/cntlCONTAINER_110/shellcont/shell").getCellValue(linha,"ZIMPRES")
      if(str(documento).isdigit()):
        debitos.append(documento)
    if(len(debitos) == 0): raise Exception("404: Cliente nao possui faturas vencidas!")
    if(len(debitos) > 6): raise Exception(f"429: Cliente possui muitas faturas ({len(debitos)}) pendentes")
    self.ZATC45(instalacao, parceiro, debitos)
    return str(len(debitos))
  def ZSVC168(self, instalacoes) -> str:
    dataframe = {
      "Nota": [],
      "Inst": [],
      "Data": [],
      "Hora": [],
      "Stay": [],
      "Poste": [],
      "Trafo": [],
      "Serie": [],
      "Ramal": [],
      "Fases": [],
      "Lance": [],
      "Local": [],
      "Vaos": [],
      "Lado": []
    }
    self.session.StartTransaction(Transaction="ZSVC168")
    self.session.FindById("wnd[0]/usr/btn%_S_INSTAL_%_APP_%-VALU_PUSH").Press()
    tabela = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
    for i in range(len(instalacoes)):
      self.session.FindById(f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = instalacoes[i]
      self.session.FindById(f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = i + 1
    self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
    self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    tabela = self.session.FindById("wnd[0]/usr/cntlALV/shellcont/shell")
    rows = tabela.rowCount
    if(rows == 0): raise Exception("404: Nao foram encontrados dados do Croqui Digital")
    for i in range(rows):
      dataframe["Nota"].append(tabela.getCellValue(i, "QMNUM"))
      dataframe["Inst"].append(tabela.getCellValue(i, "INSTALACAO"))
      dataframe["Local"].append(tabela.getCellValue(i, "LOCAL_CONSUMO"))
      dataframe["Data"].append(tabela.getCellValue(i, "DT_REGISTRO"))
      dataframe["Hora"].append(tabela.getCellValue(i, "HORA_REGISTRO"))
      dataframe["Stay"].append(tabela.getCellValue(i, "PING_COOR"))
      dataframe["Poste"].append(tabela.getCellValue(i, "RAMAL_COOR"))
      dataframe["Trafo"].append(tabela.getCellValue(i, "TRAFO_COOR"))
      dataframe["Serie"].append(tabela.getCellValue(i, "NUM_TRAFO"))
      dataframe["Ramal"].append(tabela.getCellValue(i, "ESPECIFICACAO_RAMAL"))
      dataframe["Fases"].append(tabela.getCellValue(i, "TIPO_FASE"))
      dataframe["Lance"].append(tabela.getCellValue(i, "TAM_RAMAL"))
      dataframe["Vaos"].append(tabela.getCellValue(i, "NUM_VAO"))
      dataframe["Lado"].append(tabela.getCellValue(i, "LADO_REDE"))
    return pandas.DataFrame(dataframe).to_csv(index=False)
  def zona(self, arg) -> str:
    instalacao = self.instalacao(arg)
    leiturista = self.leiturista(instalacao)
    dataframe = pandas.read_csv(io.StringIO(leiturista))
    instalacoes = dataframe["Instalacao"].values
    return self.ZSVC168(instalacoes)
  def get_dataframe_from_table(self, xpath) -> str:
    data = {}
    cols = 0
    rows = 0
    tabela = self.session.FindById(xpath)
    if(tabela.type == 'GuiShell'):
      if(tabela.subType == 'GridView'):
        cols = tabela.columnCount
        rows = tabela.rowCount
        pass
    return ""
  def try_catch(self, prefixo) -> None:
    for i in range(0, 50):
      transacao = prefixo + str(i)
      self.session.StartTransaction(Transaction=transacao)
      barra = self.session.FindById("wnd[0]/sbar").text
      print(f"{transacao} : {barra}")
      time.sleep(1)
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
  try:
    if(len(sys.argv) < 3): raise Exception("500: Falta argumentos necessarios!")
    aplicacao = sys.argv[1]
    argumento = int(sys.argv[2])
    if(len(sys.argv) == 3): instancia = 0
    else: instancia = int(sys.argv[3])
    # Attempts to connect to SAP FrontEnd on the specified instance
    try: robo = sap(instancia)
    except: raise Exception("500: Nao pode se conectar ao sistema SAP!")
    if(argumento == 'instancia'):
      robo.create_lock()
    have_authorization = True
    telefone = None
    # If the number of arguments is greater than the minimum (4),
    # then it checks the other arguments (now, only one optional argument is accepted).
    if(len(sys.argv) > 4):
      apontador = 4
      while(apontador < len(sys.argv)):
        if ('--sap-restrito' == sys.argv[apontador]): have_authorization = False
        elif ('--baixada' == sys.argv[apontador]): robo.REGIAO = 'RB'
        elif ('--oeste' == sys.argv[apontador]): robo.REGIAO = 'RO'
        elif ('--leste' == sys.argv[apontador]): robo.REGIAO = 'RL'
        elif (str(sys.argv[apontador]).startswith("--telefone")):
          telefone = str(sys.argv[apontador]).split('=')[1]
        elif ('--em-desenvolvimento'): pass
        else: raise Exception("500: O argumento fornecido nao eh valido!")
        apontador = apontador + 1
    if(telefone != None):
      telefone = re.search("[0-9]{11}$", telefone)
      if(telefone != None):
        telefone = telefone.group()
    # Attempts to execute the method requested in the first argument
    try:
      if(aplicacao == "instancia"):
        robo.health_check(argumento)
        sys.exit()
      else:
        robo.inicializar()
      if (aplicacao == "coordenada"):
        print(robo.coordenadas(argumento))
      elif ((aplicacao == "telefone") or (aplicacao == "contato")):
        if(not have_authorization): raise Exception("401: Nao eh possivel consultar essas informacoes no modo restrito")
        print(robo.telefone(argumento))
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
        # Recreate lockfile to force to check number of instances
        # If instance numbers is greater that 5, it will be fail
        robo.create_lock()
        robo.inicializar()
        if("ZATC73" in robo.NOTUSE):
          print(robo.fatura_ZATC45(argumento))
        elif(have_authorization):
          print(robo.fatura(argumento))
        else:
          print(robo.fatura_novo(argumento))
      elif (aplicacao == "relatorio"):
        print(robo.relatorio(argumento))
      elif ((aplicacao == "historico") or (aplicacao == "historico")):
        print(robo.historico(argumento))
      elif (aplicacao == "agrupamento"):
          print(robo.agrupamento(argumento, have_authorization))
      elif (aplicacao == "pendente"):
        if(have_authorization):
          print(robo.escrever(argumento))
        else:
          print(robo.escrever_novo(argumento))
      elif (aplicacao == "bandeirada"):
        print(robo.bandeirada(argumento))
      elif((aplicacao == "informacao") or (aplicacao == "medidor")):
        if(not have_authorization): raise Exception("401: Nao eh possivel consultar essas informacoes no modo restrito")
        else: print(robo.novo_informacao(argumento))
      elif(aplicacao == "instalacao"):
        print(robo.instalacao(argumento))
      elif(aplicacao == "cruzamento"):
        print(robo.cruzamento(argumento))
      elif((aplicacao == "consumo") or (aplicacao == "leitura")):
        print(robo.consumo(argumento))
      elif(aplicacao == "abertura"):
        print(robo.inspecao(argumento))
      elif(aplicacao == "vencimento"):
        print(robo.relatorio(argumento, True))
      elif(aplicacao == "ren360"):
        print(robo.procurar(argumento))
      elif(aplicacao == "codbarra"):
        if(telefone == None): raise Exception("500: Nao foi informado telefone")
        print(robo.codbarra(argumento, telefone))
      elif(aplicacao == "zona"):
        print(robo.zona(argumento))
      elif(aplicacao == "fuga"):
        print(robo.agrupamento(nota=argumento, have_authorization=True, debitos=True))
      else:
        raise Exception("400: Nao entendi o comando, verifique se esto correto!")
      robo.retorno()
    except Exception as erro:
      if not isinstance(erro.args[0], str):
        print("500: Um erro inesperado aconteceu.")
      elif(re.match("^[0-9]{3}", erro.args[0]) == None):
        print("500: " + str(erro.args[0]))
      else:
        print(erro.args[0])
  except Exception as erro:
    print(erro)