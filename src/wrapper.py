#!/usr/bin/python
''' Module to wraper SAPGUI scripting engine automation '''
# coding: utf8
#region imports
import re
import os
import sys
import time
import datetime
import subprocess
import logging
from logging.handlers import RotatingFileHandler
import win32com.client
import pandas
from constants import (
  SHORT_TIME_WAIT,
  LONG_TIME_WAIT,
  LOCKFILE,
  BASE_FOLDER,
)
from helpers import depara, STRINGPATH
from exceptions import (
  WrapperBaseException,
  TooMannyRequests,
  SomethingGoesWrong,
  UnavailableSap,
  ArgumentException,
  InformationNotFound,
)
from models import (
  InstalacaoInfo,
  LigacaoInfo,
  LogradouroInfo,
  ServicoInfo,
  ParceiroInfo,
  MedidorInfo,
)
from enumerators import (
  DESTAQUES,
  IW53_FLAGS,
  ES32_FLAGS,
  ZMED89_FLAGS,
  ZARC140_FLAGS,
  ES61_FLAGS,
  ES57_FLAGS,
  ZMED95_FLAGS,
  BP_FLAGS,
  IQ03_FLAGS,
)
__ALL__ = (
  'check_lock',
  'create_lock',
  'delete_lock',
  'attach_session',
  'HOME_PAGE',
  'SEND_ENTER',
  'CHECK_STATUS',
  'GETBY_XY', # get element by col and row numbers
  'ZATC73', # print by invoice document number
  'ZSVC20', # get services report
  'IW53', # get information about service
  'ES32', # get information about instalation
  'ZATC45', # print invoice when ZATC73 unavaliable
  'ZMED89', # get reading report
  'ZARC140', # get invoice report
  'ES61', # get information about ligacao
  'ES57', # get information about street
  'ZMED95', # get information about logradouro
  'FPL9', # get invoice report then ZARC140 unavaliable
  'BP', # get information about costumer
  'ZATC66', # get consume report
  'IQ03', # get information about meter
)
#endregion

class SapBot:
  ''' Wrapper class to SAP FrontEnd automation '''
  def check_lock(self) -> bool:
    ''' Check if LOCKFILE exists in base folder '''
    return os.path.exists(LOCKFILE)
  def create_lock(self) -> None:
    ''' Create LOCKFILE if it is not exist '''
    if not self.check_lock():
      with open(LOCKFILE, 'w', encoding='UTF8'):
        pass
  def delete_lock(self) -> None:
    ''' Remove LOCKFILE if it is not exist '''
    if self.check_lock():
      os.remove(LOCKFILE)
  def attach_session(self, instancia: int) -> None:
    ''' Function attach session in SAPGUI scripting engine '''
    try:
      while self.check_lock():
        time.sleep(1)
      self.sap_gui = win32com.client.GetObject('SAPGUI').GetScriptingEngine
      self.connection = self.sap_gui.connections[0]
      self.session = self.connection.Children(instancia)
    except:
      self.create_lock()
      self.attach_session(instancia)
  def create_session(self, instancia: int) -> None:
    ''' Function create session in SAPGUI scripting engine '''
    self.logger.info('Starting instance checker...')
    while True:
      if not self.check_lock():
        time.sleep(LONG_TIME_WAIT)
        continue
      try:
        # Get scripting
        try:
          self.sap_gui = win32com.client.GetObject('SAPGUI').GetScriptingEngine
        except:
          self.create_lock()
          self.logger.warning('Cannot able to attach ScriptingEngine, starting SAP...')
          saplogon = STRINGPATH['SAP_EXECUTABLE_FILEPATH']
          # pylint: disable-next=consider-using-with
          subprocess.Popen(saplogon, start_new_session=True)
          time.sleep(SHORT_TIME_WAIT)
          self.sap_gui = win32com.client.GetObject('SAPGUI').GetScriptingEngine
        if not isinstance(self.sap_gui, win32com.client.CDispatch):
          self.create_lock()
          self.logger.error('Cannot able to attach ScriptingEngine, even after start SAP Frontend.')
          raise UnavailableSap('SAP GUI Scripting API is not available.')
        # Get connection
        if not len(self.sap_gui.connections) > 0:
          self.create_lock()
          self.logger.warning('Connection is not open. trying open...')
          try:
            self.connection = self.sap_gui.OpenConnection('#PCL', True)
          except:
            self.create_lock()
            self.logger.error('Cannot able to open connection with SAP server.')
            raise UnavailableSap('SAP FrontEnd connection is not available.')
        else:
          self.connection = self.sap_gui.connections[0]
        self.session = self.connection.Children(0)
        # Check and get authentication
        if self.session.info.user == '':
          self.create_lock()
          self.logger.warning('User is not authenticated yet. Authenticating...')
          self.session.FindById(STRINGPATH['LOGIN_AUTH_USERNAME']).text = os.environ.get('USUARIO')
          self.session.FindById(STRINGPATH['LOGIN_AUTH_PASSWORD']).text = os.environ.get('PALAVRA')
          self.session.FindById(STRINGPATH['GLOBAL_ENTER_BUTTON']).Press()
          if self.session.findById('wnd[1]', False) is not None:
            if self.session.findById(STRINGPATH['LOGIN_POPUP_OPTION'], False) is not None:
              self.session.findById(STRINGPATH['LOGIN_POPUP_OPTION']).Select()
            self.session.findById(STRINGPATH['POPUP_ENTER_BUTTON']).Press()
        # re-check authentication
        if self.session.info.user == '':
          self.create_lock()
          self.logger.error('User cannot be authenticated.')
          raise UnavailableSap('User cannot be authenticated!')
        # Create sessions
        number_of_sessions = len(self.connection.sessions)
        if number_of_sessions < instancia:
          self.create_lock()
          self.logger.warning('Less instances that desire, creating new ones...')
          for i in range(instancia - number_of_sessions):
            self.connection.Children(0).createSession()
            time.sleep(SHORT_TIME_WAIT)
        # Re-check number of sessions
        number_of_sessions = len(self.connection.sessions)
        if number_of_sessions > instancia:
          self.create_lock()
          self.logger.warning('More instances that desire, closing excess...')
          for i in range(number_of_sessions, instancia, -1):
            self.connection.closeSession(self.connection.sessions[i - 1].Id)
        # Unlock instances
        self.delete_lock()
        self.logger.info('SAP Frontend is ready to receive requests.')
      except WrapperBaseException as erro:
        self.logger.error(erro.message)
      except Exception as erro:
        self.logger.error(erro.args[0])
  def __init__(self, instancia: int) -> None:
    ''' Define instance number and config logger '''
    caminho_logs = os.path.join(BASE_FOLDER, 'log')
    if not os.path.exists(caminho_logs):
      os.mkdir(caminho_logs)
    logfilename = os.path.join(caminho_logs, f'logfile_{instancia}.log')
    logging.basicConfig(
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        level=logging.DEBUG,
        handlers=[
          RotatingFileHandler(logfilename, maxBytes=10000000, backupCount=5),
          logging.StreamHandler(sys.stdout) #! Remove to do not display log messages
        ]
    )
    self.logger = logging.getLogger(__name__)
  def HOME_PAGE(self) -> None:
    ''' Function to return to home page '''
    self.session.sendCommand('/n')
  def SEND_ENTER(self) -> None:
    ''' Function to send 'Enter' key '''
    self.session.FindById("wnd[0]").SendVKey(2)
  def CHECK_STATUSBAR(self, check_false_sucess_text: str = '') -> str: # TODO - pass a clear check to this function
    ''' Function to check errors message on status bar '''
    statusbar = self.session.findById(STRINGPATH['STATUS_BAR_MESSAGE'])
    '''
    Property 'messageType':
    
    S - Sucess
    W - Warn
    E - Error
    A - Abort
    I - Info
    '''
    if statusbar.messageType == 'E':
      raise InformationNotFound(statusbar.text)
    if check_false_sucess_text:
      if check_false_sucess_text in statusbar.text:
        raise InformationNotFound(statusbar.text)
    return statusbar.text
  def GETBY_XY(self, id_template: str, col: int, row: int):
    ''' function to get element replace col and row from array id '''
    id_string = id_template.replace('¿', str(col)).replace('?', str(row))
    return self.session.FindById(id_string, False)
  def ZATC73(
      self,
      documentos: list[int]
      ) -> None:
    ''' Function that send list of invoice document number to print '''
    self.session.StartTransaction(Transaction="ZATC73")
    self.CHECK_STATUSBAR()
    self.session.FindById(STRINGPATH['ZATC73_PRINT_DEFAULT']).selected = True
    self.session.FindById(STRINGPATH['ZATC73_PRINT_DEFAULT2']).selected = True
    for documento in enumerate(documentos):
      self.session.FindById(STRINGPATH['ZATC73_PRINT_ASDDASD']).text = documento
      self.session.FindById(STRINGPATH['GLOBAL_ACCEPT_BUTTON']).Press()
      if self.session.FindById(STRINGPATH['ZATC73_POPUP_OPTION'], False) is not None:
        self.session.FindById(STRINGPATH['ZATC73_POPUP_OPTION']).Press()
  def ZSVC20(
      self,
      tipos_notas: list[str],
      min_data: datetime.date,
      max_data: datetime.date,
      danos_filtro: list[str],
      statuses: list[str],
      regional: str,
      layout: str,
      colluns: list[str],
      colluns_names: list[str],
      instalacao: InstalacaoInfo | None = None
    ) -> pandas.DataFrame:
    ''' Function to run ZSVC20 transaction and return table data '''
    self.session.StartTransaction(Transaction='ZSVC20')
    self.CHECK_STATUSBAR()
    #
    if instalacao is not None:
      self.session.FindById(STRINGPATH['ZSVC20_INSTALLATION_INPUT']).text = instalacao.instalacao
    # Insere a lista de tipo de nota no formulário
    self.session.FindById(STRINGPATH['ZSVC20_MULTIPLE_TYPES']).Press()
    for i, tipo in enumerate(tipos_notas):
      self.session.FindById(STRINGPATH['ZSVC20_POPUP_OPTION'].replace('?',str(i))).text = tipo
      self.session.FindById(STRINGPATH['ZSVC20_POPUP_WINDOW']).verticalScrollbar.position = i
    self.session.FindById(STRINGPATH['POPUP_ACCEPT_BUTTON']).Press()
    # Insere as datas de janela de datas do relatório
    self.session.FindById(STRINGPATH['ZSVC20_PERIODO_INICIO']).text = min_data.strftime('%d.%m.%Y')
    self.session.FindById(STRINGPATH['ZSVC20_PERIODO_FINAL']).text = max_data.strftime('%d.%m.%Y')
    # Insere a lista de danos permitidos no formulário
    self.session.FindById(STRINGPATH['ZSVC20_MULTIPLE_DANO']).Press()
    for i, dano in enumerate(danos_filtro):
      self.session.FindById(STRINGPATH['ZSVC20_POPUP_OPTION'].replace('?',str(i))).text = dano
      self.session.FindById(STRINGPATH['ZSVC20_POPUP_WINDOW']).verticalScrollbar.position = i
    self.session.FindById(STRINGPATH['POPUP_ACCEPT_BUTTON']).Press()
    # Insere a lista de status permitidos no formulário
    self.session.FindById(STRINGPATH['ZSVC20_MULTIPLE_STATUS']).Press()
    for i, status in enumerate(statuses):
      self.session.FindById(STRINGPATH['ZSVC20_POPUP_OPTION'].replace('?',str(i))).text = status
      self.session.FindById(STRINGPATH['ZSVC20_POPUP_WINDOW']).verticalScrollbar.position = i
    self.session.FindById(STRINGPATH['POPUP_ACCEPT_BUTTON']).Press()
    # Insere a regional, layout e roda o formulário
    self.session.FindById(STRINGPATH['ZSVC20_REGIONAL_TEXT']).text = regional
    self.session.FindById(STRINGPATH['ZSVC20_LAYOUT_TEXT']).text = layout
    self.session.FindById(STRINGPATH['GLOBAL_ACCEPT_BUTTON']).Press()
    self.CHECK_STATUSBAR()
    # Verifica se há resultados no relatório
    tabela = self.session.FindById(STRINGPATH['ZSVC20_RESULT_TABLE'], False)
    if tabela is None:
      raise InformationNotFound('O relatorio de notas esta vazio!')
    if tabela.RowCount == 0:
      raise InformationNotFound('O relatorio de notas esta vazio!')
    # Coleta as informações da tabela de acordo com as colunas esperadas
    dataframe = {key: [] for key in colluns_names}
    dataframe['#'] = []
    for i in range(tabela.RowCount):
      for j, column in enumerate(colluns):
        dataframe[colluns_names[j]].append(tabela.getCellValue(i, column))
      tabela.firstVisibleRow = i
      dataframe['#'].append(DESTAQUES.AUSENTE)
    dataframe = pandas.DataFrame(dataframe)
    reordered_columns = ['#'] + [col for col in dataframe.columns if col != '#']
    dataframe = dataframe[reordered_columns]
    return dataframe
  def IW53(
    self,
    nota: int,
    flags: list[IW53_FLAGS] = [IW53_FLAGS.GET_INST]
    ) -> ServicoInfo:
    ''' Function to get information about service '''
    servico = ServicoInfo()
    servico.nota = nota
    self.session.StartTransaction(Transaction="IW53")
    self.CHECK_STATUSBAR()
    self.session.FindById(STRINGPATH['IW53_SERVICE_INPUT']).text = nota
    self.session.FindById(STRINGPATH['IW53_ENTER_BUTTON']).Press() # TODO - Verificar se não pode ser trocado por GLOBAL_ENTER_BUTTON
    self.CHECK_STATUSBAR()
    if IW53_FLAGS.GET_INST in flags:
      self.session.FindById(STRINGPATH['IW53_INSTALLATION_TAB']).Select()
      servico.instalacao = self.session.FindById(STRINGPATH['IW53_INSTALLATION_TEXT']).text
    if IW53_FLAGS.GET_INFO in flags:
      # TODO - COLETAR O RESTO DAS INFORMAÇÔES DO SERVICO
      pass
    return servico
  def ES32(
    self,
    instalacao: int,
    flags: list[ES32_FLAGS] = [ES32_FLAGS.ONLY_INST]
    ) -> InstalacaoInfo:
    ''' Function to get information about installation '''
    self.session.StartTransaction(Transaction="ES32")
    self.session.FindById(STRINGPATH['ES32_INSTALLATION_INPUT']).text = instalacao
    self.session.FindById(STRINGPATH['GLOBAL_ENTER_BUTTON']).Press()
    self.CHECK_STATUSBAR(check_false_sucess_text='não existe')
    data = InstalacaoInfo()
    data.instalacao = instalacao
    data.status = self.session.findById(STRINGPATH['ES32_STATUS_TEXT']).text
    data.classe = int(self.session.FindById(STRINGPATH['ES32_CLASSE_TEXT']).text)
    data.texto_classe = str(data.classe) + ' - ' + depara('classe_subclasse', str(data.classe))
    data.consumo = int(self.session.FindById(STRINGPATH['ES32_CONSUMO_TEXT']).text)
    data.contrato = int(self.session.FindById(STRINGPATH['ES32_CONTRATO_TEXT']).text)
    data.parceiro = int(self.session.findById(STRINGPATH['ES32_PARCEIRO_TEXT']).text)
    data.unidade = self.session.FindById(STRINGPATH['ES32_UNIDADE_TEXT']).text
    data.endereco = self.session.FindById(STRINGPATH['ES32_NOME_ENDERECO_TEXT']).text
    data.nome_cliente = self.session.findById(STRINGPATH['ES32_NOMECLIENTE_TEXT']).text
    if ES32_FLAGS.ENTER_CONSUMO in flags:
      self.session.FindById(STRINGPATH['ES32_CONSUMO_TEXT']).setFocus()
      self.SEND_ENTER()
      return data
    if ES32_FLAGS.GET_CENTER in flags:
      self.session.FindById(STRINGPATH['ES32_UNIDADE_TEXT']).setFocus()
      self.SEND_ENTER()
      data.centro = int(self.session.findById(STRINGPATH['ES32_CENTRO_TEXT']).text)
      return data
    if ES32_FLAGS.GET_METER in flags:
      self.session.FindById(STRINGPATH['ES32_MEDIDOR_BUTTON']).Press()
      if self.session.FindById(STRINGPATH['POPUP'], False) is not None:
        self.session.FindById(STRINGPATH['POPUP_ENTER_BUTTON']).Press()
        return data
      ESPACO_VAZIO = '__________________'
      for i in range(self.session.findById(STRINGPATH['ES32_EQUIPAMENTO_TABLE']).RowCount):
        material = self.session.findById(STRINGPATH['ES32_EQUIPAMENTO_CODIGO'].replace('?',str(i))).text
        serial = self.session.findById(STRINGPATH['ES32_EQUIPAMENTO_SERIAL'].replace('?',str(i))).text
        if material == ESPACO_VAZIO:
          break
        medidor = MedidorInfo()
        medidor.serial = int(serial)
        medidor.material = int(material)
        medidor.texto_material = material + ' - ' + depara('material_codigo', material)
        data.equipamento.append(medidor)
      data.equipamento = [eq for eq in data.equipamento if eq.serial != ESPACO_VAZIO]
    return data
  def ZATC45(
      self,
      instalacao: InstalacaoInfo,
      documentos:list[int]
      ) -> None:
    ''' Function that request print invoice trought `ZATC45` transaction when `ZATC73` is unavaliable '''
    self.session.StartTransaction(Transaction="ZATC45")
    self.CHECK_STATUSBAR()
    self.session.FindById(STRINGPATH['ZATC45_2THVIA_RADIO']).Select()
    self.session.findById(STRINGPATH['ZATC45_PARCEIRO_INPUT']).text = instalacao.parceiro
    self.session.findById(STRINGPATH['ZATC45_INSTALLATION_INPUT']).text = instalacao.instalacao
    self.session.FindById(STRINGPATH['GLOBAL_ACCEPT_BUTTON']).Press()
    self.CHECK_STATUSBAR()
    tabela = self.session.FindById(STRINGPATH['ZATC45_RESULT_TABLE'])
    # Verificando se as faturas solicitadas estão na tabela
    indices = []
    for i in range(tabela.RowCount):
      documento = self.session.FindById(STRINGPATH['ZATC45_RESULT_TABLE']).text
      if documento in documentos:
        indices.append(i)
    if len(indices) != len(documentos):
      raise ArgumentException('A quantidade de faturas nao bate com o esperado!')
    self.session.FindById(STRINGPATH['ZATC45_2THVIA_RADIO2']).Select()
    for i in indices:
      self.session.FindById(STRINGPATH['ZATC45_PRINT_CHECK'].replace('?', i)).selected = True
      self.session.FindById(STRINGPATH['GLOBAL_ACCEPT_BUTTON']).Press()
      if self.session.FindById(STRINGPATH['ZATC45_POPUP_JUSTIFICATIVA'], False) is not None:
        self.session.FindById(STRINGPATH['ZATC45_POPUP_JUSTIFICATIVA']).Select()
        self.session.FindById(STRINGPATH['POPUP_ENTER_BUTTON']).Press()
      if self.session.FindById(STRINGPATH['POPUP'], False) is not None:
        self.session.FindById(STRINGPATH['POPUP_ENTER_BUTTON']).Press()
      self.session.FindById(STRINGPATH['ZATC45_PRINT_CHECK'].replace('?', i)).selected = False
      self.CHECK_STATUSBAR()
  def ZMED89(
      self,
      instalacao: InstalacaoInfo,
      quantidade: int,
      collumns: list[str],
      collumns_names: list[str],
      flag: ZMED89_FLAGS
      ) -> pandas.DataFrame:
    ''' Function that get information about reading report '''
    if instalacao.centro is None:
      raise SomethingGoesWrong('A propriedade `centro` não foi definida!')
    self.session.StartTransaction(Transaction="ZMED89")
    mes = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    lote = instalacao.unidade[:2]
    # Checks if query all meter centers or only one
    if flag == ZMED89_FLAGS.TELEMEDIDO:
      self.session.FindById(STRINGPATH['ZMED89_CENTRO_MIN']).text = '001'
      self.session.FindById(STRINGPATH['ZMED89_CENTRO_MAX']).text = '100'
    else:
      self.session.FindById(STRINGPATH['ZMED89_CENTRO_MIN']).text = instalacao.centro
      self.session.FindById(STRINGPATH['ZMED89_CENTRO_MAX']).text = ''
    # Fill the rest of form
    self.session.FindById(STRINGPATH['ZMED89_LOTE_INPUT']).text = lote
    self.session.FindById(STRINGPATH['ZMED89_MES_REFERENCIA']).text = mes.strftime("%m/%Y")
    self.session.FindById(STRINGPATH['ZMED89_UNIDADE_INPUT']).text = instalacao.unidade
    self.session.FindById(STRINGPATH['GLOBAL_ACCEPT_BUTTON']).Press()
    # Check if has error
    self.CHECK_STATUSBAR()
    # Select first layout
    self.session.FindById(STRINGPATH['ZMED89_LAYOUT_BUTTON']).Press()
    self.session.FindById(STRINGPATH['ZMED89_LAYOUT_TABLE']).setCurrentCell(0,'DEFAULT')
    self.session.FindById(STRINGPATH['ZMED89_LAYOUT_TABLE']).clickCurrentCell()
    # Order by sequence number
    if flag in { ZMED89_FLAGS.SEQ_ORDER, ZMED89_FLAGS.TELEMEDIDO }:
      self.session.FindById(STRINGPATH['ZMED89_RESULT_TABLE']).selectColumn('ZZ_NUMSEQ')
      self.session.FindById(STRINGPATH['ZMED89_ORDER_ASC_BUTTON']).Press()
    # Search for instalation
    self.session.FindById(STRINGPATH['ZMED89_LOCALIZAR_BUTTON']).Press()
    self.session.FindById(STRINGPATH['ZMED89_LOCALIZAR_INPUT']).text = instalacao.instalacao
    self.session.FindById(STRINGPATH['ZMED89_LOCALIZAR_ORDER']).key = '0'
    self.session.FindById(STRINGPATH['POPUP_ENTER_BUTTON']).Press()
    self.session.FindById(STRINGPATH['POPUP']).Close()
    celula = self.session.FindById(STRINGPATH['ZMED89_RESULT_TABLE']).currentCellRow
    # Checking if searched instalation is found
    if int(self.session.FindById(STRINGPATH['ZMED89_RESULT_TABLE']).getCellValue(celula,"ANLAGE")) != instalacao.instalacao:
      raise InformationNotFound('A instalacao nao foi encontrada no relatorio!')
    # Get informations to collect report
    linhas = self.session.FindById(STRINGPATH['ZMED89_RESULT_TABLE']).RowCount
    dataframe = {key: [] for key in collumns_names}
    dataframe['#'] = []
    quantidade = int(quantidade / 2)
    min_row = 0 if celula <= quantidade else celula - quantidade
    max_row = linhas if linhas <= celula + quantidade else celula + quantidade
    # Collect report information
    for i in range(min_row, max_row + 1):
      for j, collumn in enumerate(collumns):
        value = self.session.FindById(STRINGPATH['ZMED89_RESULT_TABLE']).getCellValue(i, collumn)
        dataframe[collumns_names[j]].append(value)
      destaque = DESTAQUES.AMARELO if i == celula else DESTAQUES.AUSENTE
      dataframe['#'].append(destaque)
    dataframe = pandas.DataFrame(dataframe)
    reordered_columns = ['#'] + [col for col in dataframe.columns if col != '#']
    dataframe = dataframe[reordered_columns]
    return dataframe
  def ZARC140(
    self,
    instalacao: InstalacaoInfo,
    flag: ZARC140_FLAGS = ZARC140_FLAGS.GET_PENDING
    ) -> pandas.DataFrame:
    ''' Function that get information about pending invoice report '''
    self.session.StartTransaction(Transaction="ZARC140")
    self.CHECK_STATUSBAR()
    self.session.FindById(STRINGPATH['ZARC140_PARCEIRO_INPUT']).text = instalacao.parceiro
    self.session.FindById(STRINGPATH['ZARC140_CONTRATO_INPUT']).text = instalacao.contrato
    self.session.FindById(STRINGPATH['ZARC140_INSTALACAO_INPUT']).text = instalacao.instalacao
    self.session.FindById(STRINGPATH['ZARC140_REAVISOS_CHECK']).Selected = True
    self.session.FindById(STRINGPATH['GLOBAL_ACCEPT_BUTTON']).Press()
    self.CHECK_STATUSBAR()
    if flag == ZARC140_FLAGS.GET_PENDING:
      # Check if has pending invoices
      if self.session.FindById(STRINGPATH['ZARC140_PENDENTES_TAB'], False) is None:
        raise InformationNotFound('Instalacao consultada nao tem registro de faturas!')
      self.session.FindById(STRINGPATH['ZARC140_PENDENTES_TAB']).Select()
      tabela = self.session.FindById(STRINGPATH['ZARC140_PENDENTES_TABLE'])
      linhas = tabela.RowCount
      if linhas == 0:
        raise InformationNotFound('Cliente nao possui faturas pendentes!')
      # Collect information on pending invoice report
      collumns = ['BILLING_PERIOD', 'FAEDN', 'ZIMPRES', 'TOTAL_AMNT', 'TIP_FATURA', 'STATUS']
      collumns_names = ['Mes ref', 'Vencimento', 'Documento', 'Valor', 'Tipo', 'Status']
      dataframe = {key: [] for key in collumns_names}
      dataframe['#'] = []
      dataframe['Observacao'] = []
      for i in range(1, linhas):
        for j, collumn in enumerate(collumns):
          dataframe[collumns_names[j]].append(tabela.getCellValue(i, collumn))
        if dataframe['Status'][i-1] == '@5B@':
          dataframe['#'].append(DESTAQUES.VERDE)
          dataframe['Observacao'].append('Fat. no prazo')
          continue
        if dataframe['Status'][i-1] == '@5C@':
          dataframe['#'].append(DESTAQUES.VERMELHO)
          dataframe['Observacao'].append('Fat. vencida')
          continue
        if dataframe['Status'][i-1] == '@06@':
          dataframe['#'].append(DESTAQUES.AMARELO)
          dataframe['Observacao'].append('Fat. Retida')
          continue
        dataframe["#"].append(DESTAQUES.AUSENTE)
        dataframe['Observacao'].append('Consultar')
      dataframe = pandas.DataFrame(dataframe)
      reordered_columns = ['#'] + [col for col in dataframe.columns if col != '#']
      dataframe = dataframe[reordered_columns]
      return dataframe
    if flag == ZARC140_FLAGS.GET_RENOTICE:
      if self.session.FindById(STRINGPATH['ZARC140_RENOTICE_TAB'], False) is None:
        raise InformationNotFound('Instalacao consultada nao tem registro de reavisos!')
      self.session.FindById(STRINGPATH['ZARC140_RENOTICE_TAB']).Select()
      tabela = self.session.FindById(STRINGPATH['ZARC140_RENOTICE_TABLE'])
      linhas = tabela.RowCount
      if linhas == 0:
        raise InformationNotFound('Cliente nao possui reaviso de faturas!')
      collumns = ['STATUS', 'DT_MAX_CRT', 'DT_MIN_CRT']
      collumns_names = ['Status', 'Data min', 'Data max']
      dataframe = {key: [] for key in collumns_names}
      dataframe['#'] = []
      dataframe['Observacao'] = []
      for i in range(1, linhas + 1):
        for j, collumn in enumerate(collumns):
          dataframe[collumns_names[j]].append(tabela.getCellValue(i, collumn))
        if dataframe['Status'][i] == '@45@':
          dataframe['#'].append(DESTAQUES.VERMELHO)
          dataframe['Observacao'].append('Com reaviso')
          continue
        if dataframe['Data min'][i] == '' or dataframe['Data max'][i] == '':
          dataframe['#'].append(DESTAQUES.VERDE)
          dataframe['Observacao'].append('Sem reaviso')
          continue
        dtMax = datetime.datetime.strptime(dataframe['Data min'][i], '%d.%m.%Y').date()
        dtMin = datetime.datetime.strptime(dataframe['Data max'][i], '%d.%m.%Y').date()
        if datetime.date.today() > dtMin and datetime.date.today() < dtMax:
          dataframe['#'].append(DESTAQUES.VERMELHO)
          dataframe['Observacao'].append('Com reaviso')
          continue
        dataframe['#'].append(DESTAQUES.VERDE)
        dataframe['Observacao'].append('Sem reaviso')
      dataframe = pandas.DataFrame(dataframe)
      reordered_columns = ['#'] + [col for col in dataframe.columns if col != '#']
      dataframe = dataframe[reordered_columns]
      return dataframe
    raise SomethingGoesWrong('Flag argument value is unknow!')
  def ES61(
    self,
    instalacao: InstalacaoInfo,
    flags: list[ES61_FLAGS] = [ES61_FLAGS.ENTER_ENTER]
    ) -> LigacaoInfo:
    ''' Function to get information about objeto de ligacao '''
    ligacao = LigacaoInfo()
    if not ES61_FLAGS.SKIPT_ENTER in flags:
      self.session.StartTransaction(Transaction="ES61")
      self.CHECK_STATUSBAR()
      self.session.findById(STRINGPATH['ES61_CONSUMO_INPUT']).text = instalacao.consumo
      self.session.FindById(STRINGPATH['GLOBAL_ENTER_BUTTON']).Press()
    ligacao.ligacao = int(self.session.FindById(STRINGPATH['ES61_LIGACAO_TEXT']).text)
    if ES61_FLAGS.GET_COORD in flags:
      self.session.FindById(STRINGPATH['ES61_COORDENADAS_TAB']).Select()
      coordenadas = str(self.session.FindById(STRINGPATH['ES61_COORDENADAS_TEXT']).text)
      if not coordenadas:
        raise InformationNotFound('A instalacao nao possui coordenada cadastrada!')
      ligacao.coordenadas = re.sub(',', '.', coordenadas)
    if ES61_FLAGS.ENTER_LIGACAO in flags:
      self.session.FindById(STRINGPATH['ES61_LIGACAO_TEXT']).setFocus()
      self.SEND_ENTER()
    return ligacao
  def ES57(
    self,
    ligacao: LigacaoInfo,
    flag: ES57_FLAGS = ES57_FLAGS.ENTER_ENTER
    ) -> LogradouroInfo:
    ''' Function to get information about about street '''
    if not flag in {ES57_FLAGS.SKIPT_ENTER, ES57_FLAGS.SKIPT_ENTER_LOGRADOURO_ENTER}:
      self.session.StartTransaction(Transaction="ES57")
      self.CHECK_STATUSBAR()
      self.session.FindById(STRINGPATH['ES57_LIGACAO_INPUT']).text = ligacao.ligacao
      self.session.FindById(STRINGPATH['GLOBAL_ENTER_BUTTON']).Press()
    logradouro = LogradouroInfo(
      logradouro = self.session.FindById(STRINGPATH['ES57_LOGRADOURO_TEXT']).text,
      numero = self.session.FindById(STRINGPATH['ES57_NUMERO_TEXT']).text
    )
    # TODO - Colect the rest of information
    if flag in {ES57_FLAGS.ENTER_LOGRADOURO, ES57_FLAGS.SKIPT_ENTER_LOGRADOURO_ENTER}:
      self.session.FindById(STRINGPATH['ES57_LOGRADOURO_TEXT']).setFocus()
      self.SEND_ENTER()
    return logradouro
  def ZMED95(
    self,
    logradouro: LogradouroInfo,
    flag: ZMED95_FLAGS = ZMED95_FLAGS.ENTER_ENTER
    ) -> pandas.DataFrame:
    ''' Function to get information about group of instalations '''
    if flag is not ZMED95_FLAGS.SKIPT_ENTER:
      self.session.StartTransaction(Transaction="ZMED95")
      self.CHECK_STATUSBAR()
      self.session.FindById(STRINGPATH['ZMED95_LOGRADOURO_INPUT']).text = logradouro.logradouro
      self.session.FindById(STRINGPATH['GLOBAL_ENTER_BUTTON']).Press()
    self.session.FindById(STRINGPATH['ZMED95_ENDERECOS_BUTTON']).Press()
    tamanho = self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TABLE']).rows.length
    apontador = 0
    colunas = ['Endereco', 'Instalacao', 'Cliente', 'Tipo']
    dataframe = {key: [] for key in colunas}
    while apontador < self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TABLE']).RowCount:
      self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TABLE']).verticalScrollbar.position = apontador
      # Check if greater number on screen is less that expected number
      greater_str = self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TEXT'].replace('?',str(tamanho))).text
      match = re.search("[0-9]+", greater_str)
      greater_int = int(match.group()) if match is not None else 99999
      if greater_int < logradouro.numero_int:
        apontador += tamanho
        continue
      current_str = self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TEXT'].replace('?',str(0))).text
      match = re.search("[0-9]+", current_str)
      if match is None:
        apontador += 1
        continue
      current_int = int(match.group())
      if current_int > logradouro.numero_int:
        break
      if current_int == logradouro.numero_int:
        quantidade = int(self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TEXT']).text)
        self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TEXT']).GetAbsoluteRow(apontador).selected = True
        self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_BUTTON']).Press()
        for i in range(1, quantidade + 1):
          complemento = self.session.FindById(STRINGPATH['ZMED95_INSTALACOES_COMPLEMENTO']).text
          dataframe['Endereco'].append(logradouro.numero_str + complemento)
          dataframe["Instalacao"].append(self.session.FindById(STRINGPATH['ZMED95_INSTALACOES_INSTALACAO']).text)
          dataframe["Cliente"].append(self.session.FindById(STRINGPATH['ZMED95_INSTALACOES_NOMECLIENTE']).text)
          dataframe["Tipo"].append(self.session.FindById(STRINGPATH['ZMED95_INSTALACOES_CLASSE_INSTALACAO']).text)
          self.session.FindById(STRINGPATH['ZMED95_INSTALACOES_TABLE']).verticalScrollbar.position = i
      apontador += 1
    return pandas.DataFrame(dataframe)
  def FPL9(
    self,
    instalacao: InstalacaoInfo
    ) -> pandas.DataFrame:
    ''' Function that get information about pending invoice report
        throught `FPL9` transaction when `ZARC140` is unavaliable '''
    self.session.StartTransaction(Transaction="FPL9")
    self.CHECK_STATUSBAR()
    self.session.findById(STRINGPATH['FPL9_PARCEIRO_INPUT']).text = instalacao.parceiro
    self.session.findById(STRINGPATH['FPL9_CONTRATO_INPUT']).text = instalacao.contrato
    self.session.findById(STRINGPATH['GLOBAL_ENTER_BUTTON']).Press()
    self.session.findById(STRINGPATH['FPL9_UNKNOW_BUTTON']).Press() # TODO - Search for button need
    self.CHECK_STATUSBAR()
    # Prepare variables to
    #wnd[0]/usr/lbl[char, line]
    col = 1
    row = 0
    MAX_ROW = 32
    MAX_COL = 99
    hasLines = True
    firstLine = False
    indices = [0,0,0,0,0,0]
    colunas = ['#', 'Referencia', 'Impressao', 'Vencimento', 'Valores', 'Observacao']
    dataframe = {key: [] for key in colunas}
    while hasLines:
      # Aponta para a linha atual ou para a última caso o índice da linha ultrapasse o máximo
      linha = MAX_ROW if row >= MAX_ROW else row
      # Rola a barra vertical caso o índice da linha tenha ultrapassado o máximo
      if row >= MAX_ROW:
        self.session.FindById(STRINGPATH['GLOBAL_USER_AREA']).verticalScrollbar.position = row - MAX_ROW
      # Obtém o valor do label na coluna 1 da linha atual
      primeiro_caractere = self.GETBY_XY(STRINGPATH['FPL9_LABEL_CANVAS'], 1, linha)
      # Verifica se a coleta já foi iniciada e se o label está vazio, se sim, encerra a coleta
      if indices[2] > 0 and primeiro_caractere is None:
        if firstLine is False:
          firstLine = True
          row += 1
          continue
        hasLines = False
        continue
      # Verifica se tem label na linha na coluna 1, caso não tenha a linha é vazia, então pula ela
      if primeiro_caractere is None:
        row += 1
        continue
      # iteração sobre todos os caracteres da linha para coletar os índices dos labels
      while (col < MAX_COL and indices[5] == 0):
        # Obtém o valor do label na coluna atual da linha atual
        label = self.GETBY_XY(STRINGPATH['FPL9_LABEL_CANVAS'], col, linha)
        # Verifica se a coordenada do objeto retorna um objeto, se não pula para a próxima coluna
        if label is None:
          col += 1
          continue
        if label.text == "Sts":
          indices[1] = col
        if label.text == "Mês Refer":
          indices[2] = col
        if label.text == "Doc. Faturam":
          indices[3] = col
        if label.text == "Vencimento":
          indices[4] = col
        if label.text == "Valor":
          indices[5] = col
        col += 1
      if indices[2] > 0 and firstLine == True:
        
        status_icon_name = self.GETBY_XY(STRINGPATH['FPL9_LABEL_CANVAS'], indices[1], linha).iconName
        if status_icon_name == "S_TL_R":
          dataframe["#"].append(DESTAQUES.VERMELHO)
          dataframe["Observacao"].append("Fat. vencida")
        if status_icon_name == "S_TL_Y":
          dataframe["#"].append(DESTAQUES.VERDE)
          dataframe["Observacao"].append("Fat. no prazo")
        if status_icon_name == "S_TL_G":
          dataframe["#"].append(DESTAQUES.VERDE)
          dataframe["Observacao"].append("Fat. no prazo")
        if not status_icon_name in {"S_TL_R", "S_TL_Y", "S_TL_G"}:
          dataframe["#"].append(DESTAQUES.AUSENTE)
          dataframe["Observacao"].append("")
        dataframe["Referencia"].append(self.GETBY_XY(STRINGPATH['FPL9_LABEL_CANVAS'], indices[2], linha).text)
        dataframe["Impressao"].append(self.GETBY_XY(STRINGPATH['FPL9_LABEL_CANVAS'], indices[3], linha).text)
        vence = datetime.datetime.strptime(self.GETBY_XY(STRINGPATH['FPL9_LABEL_CANVAS'], indices[4], linha).text ,"%d.%m.%Y")
        dataframe["Vencimento"].append(vence)
        valor = float(str.replace(self.GETBY_XY(STRINGPATH['FPL9_LABEL_CANVAS'], indices[5], linha).text, ',', '.'))
        dataframe["Valores"].append(valor)
      row += 1
      col = 1
    dataframe1 = pandas.DataFrame(dataframe)
    # agrupa os valores por documento de impressão
    dataframe2 = dataframe1.groupby('Impressao')['Valores'].sum().reset_index()
    # remove as duplicatas para ter somente um documento de impressão por linha
    dataframe1.drop_duplicates(subset="Impressao", inplace=True)
    # mescla o dataframe com a soma dos valores com o dataframe com as informações
    dataframe3 = dataframe1.merge(dataframe2, on="Impressao")
    # remove a coluna antiga com o valor errado
    del dataframe3['Valores_x']
    dataframe3 = dataframe3.rename(columns={'Valores_y': 'valores'})
    dataframe3['Impressao'] = dataframe3['Impressao'].replace('', pandas.NA)
    dataframe3 = dataframe3.dropna(subset=['Impressao'])
    return dataframe3
  def BP(
    self,
    instalacao: InstalacaoInfo,
    flag: BP_FLAGS = BP_FLAGS.GET_PHONES
  ) -> ParceiroInfo:
    ''' Function to get information about costumer'''
    parceiro = ParceiroInfo()
    parceiro.parceiro = instalacao.parceiro
    parceiro.nome_cliente = instalacao.nome_cliente
    self.session.StartTransaction(Transaction="BP")
    self.session.FindById(STRINGPATH['BP_CLOSE_SIDE_PANEL']).Press()
    self.session.FindById(STRINGPATH['BP_PN_OPEN_POPUP']).Press()
    self.session.findById(STRINGPATH['BP_PN_POPUP_INPUT']).text = parceiro.parceiro
    self.session.findById(STRINGPATH['POPUP_ENTER_BUTTON']).Press()
    self.session.FindById(STRINGPATH['BP_DADOS_GERAIS_BUTTON']).Press()
    self.session.findById(STRINGPATH['BP_TIPO_PN_SELECT']).key = "MKK"
    if not self.session.findById(STRINGPATH['POPUP'], False) is None:
      self.session.findById(STRINGPATH['BP_DENY_WRITE_PN_BUTTON']).Press()
    if flag is BP_FLAGS.GET_DOCS:
      self.session.findById(STRINGPATH['BP_DOCS_CPF_TAB']).Select()
      parceiro.documento_tipo = self.session.findById(STRINGPATH['BP_DOCS_TIPO_TEXT']).text
      parceiro.documento_numero = self.session.findById(STRINGPATH['BP_DOCS_CPF_TEXT']).text
    if flag is BP_FLAGS.GET_PHONES:
      self.session.findById(STRINGPATH['BP_PHONE_TAB']).Select()
      for lista in {'BP_PHONE_LIST1', 'BP_PHONE_LIST2', 'BP_PHONE_LIST3'}:
        self.session.FindById(STRINGPATH[lista]).Press()
        for i in range(4):
          parceiro.telefones.append(self.GETBY_XY('BP_PHONE_TEXT', 2, i))
      parceiro.telefones = list(dict.fromkeys(parceiro.telefones))
      ESPACO_VAZIO = "______________________________"
      if ESPACO_VAZIO in parceiro.telefones:
        parceiro.telefones.remove(ESPACO_VAZIO)
      if not parceiro.telefones:
        raise InformationNotFound(f'Cliente {parceiro.nome_cliente} não tem telefone cadastrado!')
    return parceiro
  def ZATC66(
    self,
    instalacao: InstalacaoInfo
    ) -> pandas.DataFrame:
    ''' Function to get information about consumption '''
    self.session.StartTransaction(Transaction="ZATC66")
    self.session.FindById(STRINGPATH['ZATC66_INSTALACAO_INPUT']).text = instalacao.instalacao
    self.session.FindById(STRINGPATH['GLOBAL_ACCEPT_BUTTON']).Press()
    self.CHECK_STATUSBAR()
    self.session.FindById(STRINGPATH['ZATC66_LEITURA_RADIO']).Select()
    tabela = self.session.FindById(STRINGPATH['ZATC66_TABELA_RESULT'])
    nome_colunas = [
      '#',
      'Mes ref.',
      'Data leit.',
      'Medidor',
      'Leitura',
      'Consumo',
      'Registrador',
      'Tipo de leitura',
      'Motivo da leitura',
      'Nota do leiturista'
    ]
    dataframe = {key: [] for key in nome_colunas}
    temporario_code = None
    temporario_texto = None
    for i in range(tabela.RowCount):
      dataframe['#'].append(DESTAQUES.AUSENTE)
      dataframe["Mes ref."].append(tabela.getCellValue(i, "MES_ANO"))
      dataframe["Data leit."].append(tabela.getCellValue(i, "ADATSOLL"))
      dataframe["Medidor"].append(int(tabela.getCellValue(i, "GERNR")))
      temporario_code = int(str(tabela.getCellValue(i, "LEIT_FATURADA")).replace(',', '.'))
      dataframe["Leitura"].append(temporario_code)
      dataframe["Consumo"].append(0)
      # Código do registrador e texto breve descritivo
      temporario_code = tabela.getCellValue(i, "ZWNUMMER")
      if temporario_code != "":
        temporario_code = "0" + str(temporario_code) if len(temporario_code) == 1 else temporario_code
        temporario_texto = depara("medidor_registrador", temporario_code)
        dataframe["Registrador"].append(f"{temporario_code} - {temporario_texto}")
      else:
        dataframe["Registrador"].append("00 - Sem codigo do registrador")
      temporario_code = temporario_texto = None #! clear values of variables
      # Código do leiturista e texto breve descritivo
      temporario_code = tabela.getCellValue(i, "OCORRENCIA")
      if temporario_code != '':
        temporario_texto = depara("leitura_codigo", temporario_code)
        dataframe["Nota do leiturista"].append(f"{temporario_code} - {temporario_texto}")
      else:
        dataframe["Nota do leiturista"].append("")
      temporario_code = temporario_texto = None #! clear values of variables
      # Código do tipo de leitura e texto breve descritivo
      temporario_code = tabela.getCellValue(i, "TIPO_LEITURA")
      if temporario_code != '':
        temporario_texto = depara("leitura_tipo", temporario_code)
        dataframe["Tipo de leitura"].append(f"{temporario_code} - {temporario_texto}")
      else:
        dataframe["Tipo de leitura"].append("")
      temporario_code = temporario_texto = None #! clear values of variables
      # Código do motivo da leitura e texto breve descritivo
      temporario_code = tabela.getCellValue(i, "MOTIVO_LEITURA")
      if temporario_code != '':
        temporario_texto = depara("leitura_motivo", temporario_code)
        dataframe["Motivo da leitura"].append(f"{temporario_code} - {temporario_texto}")
      else:
        dataframe["Motivo da leitura"].append("00 - Sem motivo para leitura")
      temporario_code = temporario_texto = None #! clear values of variables
    dataframe = pandas.DataFrame(dataframe)
    leitura_anterior = dataframe["Leitura"].shift(-1)
    dataframe["Consumo"] = dataframe["Leitura"] - leitura_anterior
    return dataframe
  def IQ03(
    self,
    serial: int,
    material: int = 0,
    flags: list[IQ03_FLAGS] = [IQ03_FLAGS.ONLY_INST]
    ) -> list[MedidorInfo]:
    self.session.StartTransaction(Transaction="IQ03")
    self.session.FindById(STRINGPATH['IQ03_MATERIAL_INPUT']).text = material if material > 0 else ""
    self.session.FindById(STRINGPATH['IQ03_SERIAL_INPUT']).text = serial
    self.session.FindById(STRINGPATH['GLOBAL_ENTER_BUTTON']).Press()
    self.CHECK_STATUSBAR()
    varios_table = self.session.FindById(STRINGPATH['IQ03_VARIOUS_TABLE'], False)
    if not varios_table is None:
      equipamentos = []
      for i in range(varios_table.RowCount):
        equipamentos.append(int(varios_table.getCellValue(i, "MATNR")))
      medidores_info = []
      for i, equipamento in enumerate(equipamentos):
        medidores_info.extend(self.IQ03(equipamento, serial))
      return medidores_info
    medidor = MedidorInfo()
    medidor.serial = serial
    medidor.material = material
    medidor.texto_material = depara("material_codigo", str(medidor.material)) or ""
    medidor.code_montagem = self.session.FindById(STRINGPATH['IQ03_MONTAGEM_CODE']).text
    medidor.code_status = self.session.FindById(STRINGPATH['IQ03_STATUS_CODE']).text
    medidor.texto_montagem = f"{medidor.code_montagem}  -  {depara('medidor_montagem', medidor.code_montagem)}"
    medidor.texto_status = f"{medidor.code_status}  -  {depara('medidor_status', medidor.code_status)}"
    # Get the instalation attached to meter
    self.session.FindById(STRINGPATH['IQ03_INSTALATION_BUTTON']).Press()
    status = self.CHECK_STATUSBAR()
    if status: # TODO - Verificar sem lançar exceção
      medidor.observacao = status
      return [medidor]
    medidor.instalacao = int(self.session.findById(STRINGPATH['IQ03_INSTALATION_VALUE']).text)
    if IQ03_FLAGS.ONLY_INST in flags:
      return [medidor]
    leituras = self.session.FindById(STRINGPATH['IQ03_LEITURAS_TABLE'], False)
    if leituras is None:
      medidor.observacao = "Equipamento sem historico de leituras!"
      return [medidor]
    dataframe = {key: [] for key in ['Data', 'Codigo', 'Descricao']}
    limite = 12 if leituras.RowCount > 12 else leituras.RowCount
    for i in range(limite):
      data = leituras.getCellValue(i, "ADATSOLL")
      status = leituras.getCellValue(i, "ABLHINW")
      if status == '':
        continue
      dataframe['Data'].append(data)
      dataframe['Codigo'].append(status)
      texto = depara('leitura_codigo', status)
      dataframe['Descricao'].append(texto)
    medidor.leituras = pandas.DataFrame(dataframe)
    if len(medidor.leituras) > 0:
      medidor.observacao = f"Codigos de leitura nas ultimas {limite} leituras:"
    else:
      medidor.observacao = f"Sem codigos de leitura nas ultimas {limite} leituras!"
    return [medidor]
