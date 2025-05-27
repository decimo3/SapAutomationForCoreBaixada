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
import numpy
import pandas
from constants import (
  SHORT_TIME_WAIT,
  LONG_TIME_WAIT,
  LOCKFILE,
  BASE_FOLDER,
)
from helpers import depara, STRINGPATH
from conversor import conversor
from exceptions import (
  ElementNotFound,
  UnavailableTransaction,
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
  FPL9_FLAGS,
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
  'create_session',
  'attach_session',
  'HOME_PAGE',
  'SEND_ENTER',
  'CHECK_STATUS',
  'GETBY_XY', # get element by col and row numbers
  'GET_ROWS', # get data from table element
  'ZATC73', # print by invoice document number
  'ZSVC20', # get services report
  'IW53', # get information about service
  'ES32', # get information about instalation
  'ZATC45', # print invoice when ZATC73 unavaliable
  'ZMED89', # get reading report
  'ZARC140', # get invoice report
  'ES61', # get information about ligacao
  'ES62', # get information about ligacao
  'ES57', # get information about street
  'ZMED95', # get information about logradouro
  'FPL9', # get invoice report then ZARC140 unavaliable
  'BP', # get information about costumer
  'ZATC66', # get consume report
  'IQ03', # get information about meter
  'ZHISCON', # get information with CPF/CNPJ
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
          try:
            self.CHECK_STATUSBAR()
          except InformationNotFound as erro:
            self.logger.error(erro.message)
            sys.exit(1)
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
    _handlers: list[logging.Handler] = [
      RotatingFileHandler(logfilename, maxBytes=10000000, backupCount=5)
    ]
    if instancia < 0:
      _handlers.append(logging.StreamHandler(sys.stdout))
    logging.basicConfig(
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        level=logging.DEBUG,
        handlers=_handlers
    )
    self.logger = logging.getLogger(__name__)
  def HOME_PAGE(self) -> None:
    ''' Function to return to home page '''
    self.session.sendCommand('/n')
  def SEND_ENTER(self) -> None:
    ''' Function to send 'Enter' key '''
    self.session.FindById("wnd[0]").SendVKey(2)
  def CHECK_STATUSBAR(self, check_false_sucess_text: str = '', throw_error: bool = True) -> str:
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
    if not throw_error:
      return statusbar.text
    if str(statusbar.text).startswith('Sem autorização'):
      raise UnavailableTransaction(statusbar.text)
    if statusbar.messageType == 'E':
      raise InformationNotFound(statusbar.text)
    if check_false_sucess_text:
      if check_false_sucess_text in statusbar.text:
        raise InformationNotFound(statusbar.text)
    return statusbar.text
  def GETBY_XY(self, id_template: str, col: int, row: int, throw: bool = True):
    ''' function to get element replace col and row from array id '''
    id_string = STRINGPATH[id_template].replace('¿', str(col)).replace('?', str(row))
    elemento = self.session.FindById(id_string, False)
    if not elemento and throw:
      raise ElementNotFound(f'O elemento {id_template} nao foi encontrado')
    return elemento
  def GET_ROWS(self, table_id: str, columns_ids: str, columns_names: str, data_types: str, offset: int = 1, limit: int = 0) -> pandas.DataFrame:
    ''' function to get values from shell table '''
    data_types_list = STRINGPATH[data_types].split('/')
    columns_ids_list = STRINGPATH[columns_ids].split('/')
    columns_names_list = STRINGPATH[columns_names].split('/')
    dataframe = {key: [] for key in columns_names_list}
    tabela = self.session.FindById(STRINGPATH[table_id], False)
    if not tabela:
      return pandas.DataFrame(dataframe)
    if tabela.type == 'GuiShell':
      shell_or_table = True
    elif tabela.type == 'GuiTableControl':
      shell_or_table = False
    else:
      raise SomethingGoesWrong(f'The element {table_id} is not table or shell!')
    if tabela.RowCount == 0:
      return pandas.DataFrame(dataframe)
    limit = tabela.RowCount if limit == 0 or tabela.RowCount < limit else limit
    for i in range(offset, limit):
      for j, column in enumerate(columns_ids_list):
        try:
          if shell_or_table:
            valor = tabela.getCellValue(i, column)
          else:
            identificador = f'{STRINGPATH[table_id]}/{column}[{j},0]'
            valor = self.session.FindById(identificador).text
          valor = conversor[data_types_list[j]](valor)
          dataframe[columns_names_list[j]].append(valor)
        except:
          dataframe[columns_names_list[j]].append(None)
      if shell_or_table:
        try:
          tabela.firstVisibleRow = i
        except:
          pass
      else:
        identificador = f'{STRINGPATH[table_id]}/{columns_ids_list[0]}[0,1]'
        if not str(self.session.FindById(identificador).text).strip('_'):
          break
        else:
          # NOTE - Usando 'FindById' para 'table_id' novamente pois
          # a referência para variavel 'table' muda após o scroll
          self.session.FindById(STRINGPATH[table_id]).verticalScrollbar.position = i
    dataframe = pandas.DataFrame(dataframe)
    dataframe = dataframe.dropna(axis=0, how='all')
    dataframe = dataframe.dropna(axis=1, how='all')
    return pandas.DataFrame(dataframe)
  def ZATC73(
      self,
      documentos: list[int]
      ) -> None:
    ''' Function that send list of invoice document number to print '''
    self.session.StartTransaction(Transaction="ZATC73")
    self.CHECK_STATUSBAR()
    self.session.FindById(STRINGPATH['ZATC73_PRINT_DEFAULT1']).selected = True
    self.session.FindById(STRINGPATH['ZATC73_PRINT_DEFAULT2']).selected = True
    for documento in documentos:
      self.session.FindById(STRINGPATH['ZATC73_PRINT_DOCUMENT']).text = documento
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
    if self.session.FindById(STRINGPATH['ZSVC20_ERROR_POPUP_TEXT'], False):
      raise InformationNotFound(self.session.FindById(STRINGPATH['ZSVC20_ERROR_POPUP_TEXT']).text)
    return self.GET_ROWS(
      'ZSVC20_RESULT_TABLE',
      'ZSVC20_COLUMNS_IDS',
      'ZSVC20_COLUMNS_NAMES',
      'ZSVC20_COLUMNS_TYPES'
    )
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
      return servico
    if IW53_FLAGS.GET_INFO in flags:
      # Coletar informações sobre o serviço
      self.session.FindById(STRINGPATH['IW53_SOLICITACAO_TAB']).Select()
      servico.tipo = self.session.FindById(STRINGPATH['IW53_TIPO_SERVICO']).text
      servico.status = self.session.FindById(STRINGPATH['IW53_CODIGO_STATUS']).text
      servico.dano = self.session.FindById(STRINGPATH['IW53_CODIGO_DANO']).text
      servico.texto_dano = self.session.FindById(STRINGPATH['IW53_DANO_DESCRICAO']).text
      servico.descricao = self.session.FindById(STRINGPATH['IW53_DESCRICAO_NOTA']).text
      servico.observacao = self.session.FindById(STRINGPATH['IW53_OBSERVACAO']).text
      # Coletar informações sobre o cliente e endereço
      self.session.FindById(STRINGPATH['IW53_SOLICITANTE_TAB']).Select()
      self.session.FindById(STRINGPATH['IW53_CLIENTE_SUBTAB']).Select()
      servico.parceiro = self.session.FindById(STRINGPATH['IW53_CLIENTE_NUM']).text
      servico.nome_cliente = self.session.FindById(STRINGPATH['IW53_CLIENTE_NOME']).text
      servico.telefones.append(self.session.FindById(STRINGPATH['IW53_CLIENTE_TEL']).text)
      self.session.FindById(STRINGPATH['IW53_OBJETO_SUBTAB']).Select()
      servico.endereco = self.session.FindById(STRINGPATH['IW53_OBJETO_ENDERECO']).text
      servico.cod_postal = self.session.FindById(STRINGPATH['IW53_OBJETO_CEP']).text
      self.session.FindById(STRINGPATH['IW53_LOCALIZACAO_TAB']).Select()
      servico.local = self.session.FindById(STRINGPATH['IW53_LOCALIZACAO_LOCAL']).text
      servico.texto_local = self.session.FindById(STRINGPATH['IW53_LOCALIZACAO_TEXTO']).text
      # Coletar informações da instalação
      self.session.FindById(STRINGPATH['IW53_INSTALLATION_TAB']).Select()
      servico.instalacao = self.session.FindById(STRINGPATH['IW53_INSTALLATION_TEXT']).text
      servico.atendimento_obs = self.session.FindById(STRINGPATH['IW53_ATENDIMENTO_OBS']).text
      if self.session.FindById(STRINGPATH['IW53_ATENDIMENTO_TEL']).text:
        servico.telefones.append(self.session.FindById(STRINGPATH['IW53_ATENDIMENTO_TEL']).text)
      if self.session.FindById(STRINGPATH['IW53_ATENDIMENTO_CEL']).text:
        servico.telefones.append(self.session.FindById(STRINGPATH['IW53_ATENDIMENTO_CEL']).text)
      if servico.status in {'ENVI', 'LIBE', 'PEND', 'TABL'}:
        return servico
      # Coletar informações sobre os tempos
      self.session.FindById(STRINGPATH['IW53_DATAHORA_TAB']).Select()
      servico.data_nota = conversor['data'](self.session.FindById(STRINGPATH['IW53_DATA_NOTA']).text)
      servico.hora_nota = conversor['hora'](self.session.FindById(STRINGPATH['IW53_HORA_NOTA']).text)
      servico.avaria_inicio_data = conversor['data'](self.session.FindById(STRINGPATH['IW53_INICIO_AVARIA_DATA']).text)
      servico.avaria_inicio_hora = conversor['hora'](self.session.FindById(STRINGPATH['IW53_INICIO_AVARIA_HORA']).text)
      servico.desejado_inicio_data = conversor['data'](self.session.FindById(STRINGPATH['IW53_INICIO_DESEJADO_DATA']).text)
      servico.desejado_inicio_hora = conversor['hora'](self.session.FindById(STRINGPATH['IW53_INICIO_DESEJADO_HORA']).text)
      servico.avaria_final_data = conversor['data'](self.session.FindById(STRINGPATH['IW53_FINAL_AVARIA_DATA']).text)
      servico.avaria_final_hora = conversor['hora'](self.session.FindById(STRINGPATH['IW53_FINAL_AVARIA_HORA']).text)
      servico.desejado_final_data = conversor['data'](self.session.FindById(STRINGPATH['IW53_FINAL_DESEJADO_DATA']).text)
      servico.desejado_final_hora = conversor['hora'](self.session.FindById(STRINGPATH['IW53_FINAL_DESEJADO_HORA']).text)
      servico.encerramento_data = conversor['data'](self.session.FindById(STRINGPATH['IW53_ENCERRAMENTO_DATA']).text)
      servico.encerramento_hora = conversor['hora'](self.session.FindById(STRINGPATH['IW53_ENCERRAMENTO_HORA']).text)
      # Coletar informações da finalização
      self.session.FindById(STRINGPATH['IW53_FINALIZACAO_TAB']).Select()
      servico.finalizacao = self.GET_ROWS(
        'IW53_FINALIZACAO_TABLE',
        'IW53_FINALIZACAO_IDS',
        'IW53_FINALIZACAO_NAMES',
        'IW53_FINALIZACAO_TYPES'
      )
      self.session.FindById(STRINGPATH['IW53_EQUIPAMENTO_TAB']).Select()
      servico.equipamentos_inst = self.GET_ROWS(
        'IW53_EQUIPAMENTO_INST_TABLE',
        'IW53_EQUIPAMENTO_INST_IDS',
        'IW53_EQUIPAMENTO_INST_NAMES',
        'IW53_EQUIPAMENTO_INST_TYPES'
      )
      servico.equipamentos_inst = servico.equipamentos_inst.loc[servico.equipamentos_inst['Equipamento'] != 0]
      servico.equipamentos_inst['Texto breve para o registrador'] = servico.equipamentos_inst['Reg'].apply(lambda x:
          depara('medidor_registrador', str(x).zfill(2)) or 'Sem codigo do registrador')
      servico.equipamentos_remo = self.GET_ROWS(
        'IW53_EQUIPAMENTO_REMO_TABLE',
        'IW53_EQUIPAMENTO_REMO_IDS',
        'IW53_EQUIPAMENTO_REMO_NAMES',
        'IW53_EQUIPAMENTO_REMO_TYPES'
      )
      servico.equipamentos_remo = servico.equipamentos_remo.loc[servico.equipamentos_remo['Equipamento'] != 0]
      servico.equipamentos_remo['Texto breve para o registrador'] = servico.equipamentos_remo['Reg'].apply(lambda x:
          depara('medidor_registrador', str(x).zfill(2)) or 'Sem codigo do registrador')
      return servico
    raise SomethingGoesWrong('Flag argument value is unknow!')
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
    data.instalacao = int(instalacao)
    data.status = self.session.findById(STRINGPATH['ES32_STATUS_TEXT']).text
    data.classe = int(self.session.FindById(STRINGPATH['ES32_CLASSE_TEXT']).text)
    data.texto_classe = str(data.classe) + ' - ' + depara('classe_subclasse', str(data.classe))
    data.consumo = int(self.session.FindById(STRINGPATH['ES32_CONSUMO_TEXT']).text)
    temporario = self.session.FindById(STRINGPATH['ES32_CONTRATO_TEXT']).text
    data.contrato = int(temporario) if temporario else 0
    temporario = self.session.findById(STRINGPATH['ES32_PARCEIRO_TEXT']).text
    data.parceiro = int(temporario) if temporario else 0
    data.unidade = self.session.FindById(STRINGPATH['ES32_UNIDADE_TEXT']).text
    data.endereco = self.session.FindById(STRINGPATH['ES32_NOME_ENDERECO_TEXT']).text
    data.nome_cliente = str(self.session.findById(STRINGPATH['ES32_NOMECLIENTE_TEXT']).text).split('/')[0]
    data.tipo_instalacao = self.session.findById(STRINGPATH['ES32_TIPO_INSTALACAO']).text
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
        if ES32_FLAGS.DONOT_THROW in flags:
          return data
        raise InformationNotFound('Instalacao nao possui medidor!')
      if self.session.FindById(STRINGPATH['ES32_EQUIPAMENTO_TABLE'], False) is None:
        if ES32_FLAGS.DONOT_THROW in flags:
          return data
        raise InformationNotFound('Instalacao nao possui medidor!')
      ESPACO_VAZIO = '__________________'
      for i in range(self.session.FindById(STRINGPATH['ES32_EQUIPAMENTO_TABLE']).RowCount - 4):
        material = self.session.findById(STRINGPATH['ES32_EQUIPAMENTO_CODIGO'].replace('?','0')).text
        serial = self.session.findById(STRINGPATH['ES32_EQUIPAMENTO_SERIAL'].replace('?','0')).text
        if material == ESPACO_VAZIO:
          break
        medidor = MedidorInfo()
        medidor.serial = int(serial)
        medidor.material = int(material)
        medidor.texto_material = material + ' - ' + depara('material_codigo', material) or ''
        data.equipamento.append(medidor)
        self.session.FindById(STRINGPATH['ES32_EQUIPAMENTO_TABLE']).verticalScrollbar.position = i + 1
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
    status_text = self.CHECK_STATUSBAR(throw_error=False)
    if status_text == 'Nenhum débito foi encontrado!':
      raise InformationNotFound('A quantidade de faturas nao bate com o esperado!')
    if status_text:
      raise InformationNotFound(status_text)
    # Verificando se as faturas solicitadas estão na tabela
    indices = []
    quantidade = conversor['numero'](self.session.FindById(STRINGPATH['ZATC45_QUANTIDADE_TEXT']).text)
    if quantidade == 0:
      raise InformationNotFound('A quantidade de faturas nao bate com o esperado!')
    for i in range(quantidade):
      documento = conversor['numero'](self.GETBY_XY('ZATC45_DOCUMENT_NUMBER', 8, i).text)
      if documento in documentos:
        indices.append(i)
    if len(indices) != len(documentos):
      raise ArgumentException('A quantidade de faturas nao bate com o esperado!')
    self.session.FindById(STRINGPATH['ZATC45_2THVIA_RADIO2']).Select()
    for i in indices:
      self.GETBY_XY('ZATC45_PRINT_CHECK', 3, i).selected = True
      self.session.FindById(STRINGPATH['GLOBAL_ACCEPT_BUTTON']).Press()
      if self.session.FindById(STRINGPATH['ZATC45_POPUP_JUSTIFICATIVA'], False) is not None:
        self.session.FindById(STRINGPATH['ZATC45_POPUP_JUSTIFICATIVA']).Select()
        self.session.FindById(STRINGPATH['POPUP_ENTER_BUTTON']).Press()
      if self.session.FindById(STRINGPATH['POPUP'], False) is not None:
        self.session.FindById(STRINGPATH['POPUP_ENTER_BUTTON']).Press()
      self.GETBY_XY('ZATC45_PRINT_CHECK', 3, i).selected = False
      self.CHECK_STATUSBAR()
  def ZMED89(
      self,
      instalacao: InstalacaoInfo,
      flags: list[ZMED89_FLAGS],
      quantidade: int = 30
      ) -> pandas.DataFrame:
    ''' Function that get information about reading report '''
    if instalacao.centro is None:
      raise SomethingGoesWrong('A propriedade `centro` não foi definida!')
    self.session.StartTransaction(Transaction="ZMED89")
    mes = datetime.date.today() - datetime.timedelta(days=45)
    lote = instalacao.unidade[:2]
    # Checks if query all meter centers or only one
    if ZMED89_FLAGS.TELEMEDIDO in flags:
      self.session.FindById(STRINGPATH['ZMED89_CENTRO_MIN']).text = '001'
      self.session.FindById(STRINGPATH['ZMED89_CENTRO_MAX']).text = '100'
    else:
      self.session.FindById(STRINGPATH['ZMED89_CENTRO_MIN']).text = str(instalacao.centro).zfill(3)
      self.session.FindById(STRINGPATH['ZMED89_CENTRO_MAX']).text = ''
    # Fill the rest of form
    self.session.FindById(STRINGPATH['ZMED89_LOTE_INPUT']).text = lote
    self.session.FindById(STRINGPATH['ZMED89_MES_REFERENCIA']).text = mes.strftime("%m/%Y")
    self.session.FindById(STRINGPATH['ZMED89_UNIDADE_INPUT']).text = instalacao.unidade
    self.session.FindById(STRINGPATH['GLOBAL_ACCEPT_BUTTON']).Press()
    # Check if has error
    self.CHECK_STATUSBAR()
    # check if has error window
    if self.session.FindById(STRINGPATH['POPUP'], False) is not None:
      texto = self.session.FindById(STRINGPATH['ZMED89_POPUP_ERROR']).text
      self.session.FindById(STRINGPATH['POPUP']).Close()
      raise InformationNotFound(texto)
    # Select first layout
    self.session.FindById(STRINGPATH['ZMED89_LAYOUT_BUTTON']).Press()
    self.session.FindById(STRINGPATH['ZMED89_LAYOUT_TABLE']).setCurrentCell(0,'DEFAULT')
    self.session.FindById(STRINGPATH['ZMED89_LAYOUT_TABLE']).clickCurrentCell()
    # Order by sequence number
    if ZMED89_FLAGS.SEQ_ORDER in flags or ZMED89_FLAGS.TELEMEDIDO in flags:
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
    quantidade = int(quantidade / 2)
    min_row = 0 if celula <= quantidade else celula - quantidade
    max_row = linhas if linhas <= celula + quantidade else celula + quantidade
    # Collect report information
    relatorio = self.GET_ROWS(
      'ZMED89_RESULT_TABLE',
      'ZMED89_COLUMNS_IDS',
      'ZMED89_COLUMNS_NAMES',
      'ZMED89_COLUMNS_TYPES',
      min_row,
      max_row + 1
    )
    relatorio['#'] = [DESTAQUES.AMARELO if x == instalacao.instalacao
                      else DESTAQUES.AUSENTE for x in relatorio['Instalacao']]
    reordered_columns = ['#'] + [col for col in relatorio.columns if col != '#']
    relatorio = relatorio[reordered_columns]
    return relatorio
  def ZARC140(
    self,
    instalacao: InstalacaoInfo,
    flags: list[ZARC140_FLAGS] = [ZARC140_FLAGS.GET_PENDING]
    ) -> pandas.DataFrame:
    ''' Function that get information about pending invoice report '''
    self.session.StartTransaction(Transaction="ZARC140")
    self.CHECK_STATUSBAR()
    self.session.FindById(STRINGPATH['ZARC140_PARCEIRO_INPUT']).text = instalacao.parceiro
    self.session.FindById(STRINGPATH['ZARC140_CONTRATO_INPUT']).text = instalacao.contrato
    self.session.FindById(STRINGPATH['ZARC140_INSTALACAO_INPUT']).text = instalacao.instalacao
    self.session.FindById(STRINGPATH['ZARC140_REAVISOS_CHECK']).Selected = (ZARC140_FLAGS.GET_RENOTICE in flags)
    self.session.FindById(STRINGPATH['GLOBAL_ACCEPT_BUTTON']).Press()
    self.CHECK_STATUSBAR()
    if ZARC140_FLAGS.GET_PENDING in flags:
      # Check if has pending invoices
      if self.session.FindById(STRINGPATH['ZARC140_PENDENTES_TAB'], False) is None:
        if ZARC140_FLAGS.DONOT_THROW in flags:
          return pandas.DataFrame()
        raise InformationNotFound('Instalacao consultada nao tem registro de faturas!')
      self.session.FindById(STRINGPATH['ZARC140_PENDENTES_TAB']).Select()
      dataframe = self.GET_ROWS(
        'ZARC140_PENDENTES_TABLE',
        'ZARC140_PENDENTES_IDS',
        'ZARC140_PENDENTES_NAMES',
        'ZARC140_PENDENTES_TYPES'
      )
      if dataframe.shape[0] == 0:
        if ZARC140_FLAGS.DONOT_THROW in flags:
          return dataframe
        raise InformationNotFound('Instalacao consultada nao tem registro de faturas!')
      __status = {
        '@5B@': DESTAQUES.VERDE,
        '@5C@': DESTAQUES.VERMELHO,
        '@06@': DESTAQUES.AMARELO
      }
      dataframe['#'] = dataframe['Status'].apply(lambda x: __status.get(x, DESTAQUES.AUSENTE))
      reordered_columns = ['#'] + [col for col in dataframe.columns if col != '#']
      dataframe = dataframe[reordered_columns]
      return dataframe.drop('Status', axis=1)
    if ZARC140_FLAGS.GET_RENOTICE in flags:
      if self.session.FindById(STRINGPATH['ZARC140_RENOTICE_TAB'], False) is None:
        raise InformationNotFound('Instalacao consultada nao tem registro de reavisos!')
      self.session.FindById(STRINGPATH['ZARC140_RENOTICE_TAB']).Select()
      dataframe = self.GET_ROWS(
        'ZARC140_RENOTICE_TABLE',
        'ZARC140_RENOTICE_IDS',
        'ZARC140_RENOTICE_NAMES',
        'ZARC140_RENOTICE_TYPES'
      )
      if dataframe.shape[0] == 0:
        if ZARC140_FLAGS.DONOT_THROW in flags:
          return pandas.DataFrame()
        raise InformationNotFound('Instalacao consultada nao tem registro de reavisos!')
      __conditions = [
        dataframe['Status'] == '@45@',
        (dataframe['Data min'].isna() | dataframe['Data max'].isna()),
        (datetime.date.today() > dataframe['Data min']) & (datetime.date.today() < dataframe['Data max'])
      ]
      __choices = [DESTAQUES.VERMELHO, DESTAQUES.VERDE, DESTAQUES.VERMELHO]
      dataframe['#'] = numpy.select(__conditions, __choices, default=DESTAQUES.VERDE)
      dataframe['Observacao'] = dataframe['#'].apply(lambda x: 
          'Com reaviso' if x == DESTAQUES.VERMELHO else DESTAQUES.VERDE)
      reordered_columns = ['#'] + [col for col in dataframe.columns if col != '#']
      dataframe = dataframe[reordered_columns]
      return dataframe
    raise SomethingGoesWrong('Flag argument value is unknow!')
  def ES6X(
    self,
    instalacao: InstalacaoInfo,
    flags: list[ES61_FLAGS] = [ES61_FLAGS.ENTER_ENTER],
    transaction: str = 'ES61'
    ) -> LigacaoInfo:
    ''' Function to get information about local de consumo '''
    ligacao = LigacaoInfo()
    if not ES61_FLAGS.SKIPT_ENTER in flags:
      self.session.StartTransaction(Transaction=transaction)
      self.CHECK_STATUSBAR()
      self.session.findById(STRINGPATH['ES61_CONSUMO_INPUT']).text = instalacao.consumo
      self.session.FindById(STRINGPATH['GLOBAL_ENTER_BUTTON']).Press()
    ligacao.ligacao = int(self.session.FindById(STRINGPATH['ES61_LIGACAO_TEXT']).text)
    self.session.FindById(STRINGPATH['ES61_DADOS_TECNICOS_TAB']).Select()
    ligacao.tipo_instalacao = self.session.FindById(STRINGPATH['ES61_TIPO_INSTALACAO']).text
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
  def ES61(
    self,
    instalacao: InstalacaoInfo,
    flags: list[ES61_FLAGS] = [ES61_FLAGS.ENTER_ENTER]
  ) -> LigacaoInfo:
    return self.ES6X(instalacao, flags, 'ES61')
  def ES62(
    self,
    instalacao: InstalacaoInfo,
    flags: list[ES61_FLAGS] = [ES61_FLAGS.ENTER_ENTER]
  ) -> LigacaoInfo:
    return self.ES6X(instalacao, flags, 'ES62')
  def ES57(
    self,
    ligacao: LigacaoInfo,
    flag: list[ES57_FLAGS] = [ES57_FLAGS.ENTER_ENTER]
    ) -> LogradouroInfo:
    ''' Function to get information about objeto de ligacao '''
    if ES57_FLAGS.ENTER_ENTER in flag:
      self.session.StartTransaction(Transaction="ES57")
      self.CHECK_STATUSBAR()
      self.session.FindById(STRINGPATH['ES57_LIGACAO_INPUT']).text = ligacao.ligacao
      self.session.FindById(STRINGPATH['GLOBAL_ENTER_BUTTON']).Press()
    logradouro = LogradouroInfo(
      logradouro = self.session.FindById(STRINGPATH['ES57_LOGRADOURO_TEXT']).text,
      numero = self.session.FindById(STRINGPATH['ES57_NUMERO_TEXT']).text
    )
    # TODO - Colect the rest of information
    if ES57_FLAGS.ENTER_LOGRADOURO in flag:
      self.session.FindById(STRINGPATH['ES57_LOGRADOURO_TEXT']).setFocus()
      self.SEND_ENTER()
    return logradouro
  def ZMED95(
    self,
    logradouro: LogradouroInfo,
    flags: list[ZMED95_FLAGS] = [ZMED95_FLAGS.ENTER_ENTER, ZMED95_FLAGS.GET_GROUPING]
    ) -> pandas.DataFrame:
    ''' Function to get information about group of instalations '''
    if not logradouro.numero_int:
      raise InformationNotFound('Instalacao sem numero de rua, nao agrupado!')
    if logradouro.numero_str.find('SN') >= 0:
      raise InformationNotFound('Instalacao sem numero de rua, nao agrupado!')
    if not ZMED95_FLAGS.SKIPT_ENTER in flags:
      self.session.StartTransaction(Transaction="ZMED95")
      self.CHECK_STATUSBAR()
      self.session.FindById(STRINGPATH['ZMED95_LOGRADOURO_INPUT']).text = logradouro.logradouro
      self.session.FindById(STRINGPATH['GLOBAL_ENTER_BUTTON']).Press()
    if ZMED95_FLAGS.GET_GROUPING in flags:
      self.session.FindById(STRINGPATH['ZMED95_ENDERECOS_BUTTON']).Press()
      tamanho = self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TABLE']).rows.length
      apontador = 0
      colunas = ['Endereco', 'Instalacao', 'Cliente', 'Tipo']
      dataframe = {key: [] for key in colunas}
      while apontador < self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TABLE']).RowCount:
        self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TABLE']).verticalScrollbar.position = apontador
        # Check if greater number on screen is less that expected number
        greater_str = self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TEXT'].replace('?',str(tamanho - 1))).text
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
          quantidade = int(self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_QNTD']).text)
          if quantidade > 12:
            raise TooMannyRequests('Agrupamento possui instalacoes demais!')
          self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_TABLE']).GetAbsoluteRow(apontador).selected = True
          self.session.FindById(STRINGPATH['ZMED95_NUMBERS_LIST_BUTTON']).Press()
          for i in range(1, quantidade + 1):
            complemento = self.session.FindById(STRINGPATH['ZMED95_INSTALACOES_COMPLEMENTO']).text
            dataframe['Endereco'].append(current_str + ' ' + complemento)
            dataframe["Instalacao"].append(self.session.FindById(STRINGPATH['ZMED95_INSTALACOES_INSTALACAO']).text)
            dataframe["Cliente"].append(self.session.FindById(STRINGPATH['ZMED95_INSTALACOES_NOMECLIENTE']).text)
            dataframe["Tipo"].append(self.session.FindById(STRINGPATH['ZMED95_INSTALACOES_CLASSE_INSTALACAO']).text)
            self.session.FindById(STRINGPATH['ZMED95_INSTALACOES_TABLE']).verticalScrollbar.position = i
        apontador += 1
      dataframe = pandas.DataFrame(dataframe)
      if dataframe.shape[0] == 1:
        raise InformationNotFound('Instalacao unica no sistema!')
      return dataframe
    if ZMED95_FLAGS.GET_CROSSING in flags:
      self.session.FindById(STRINGPATH['ZMED95_CRUZAMENTOS_BUTTON']).Press()
      container = self.session.FindById(STRINGPATH['ZMED95_CRUZAMENTOS_TABLE'])
      if container.RowCount == 0:
        raise InformationNotFound('Nao ha informacao de cruzamentos')
      return self.GET_ROWS(
        'ZMED95_CRUZAMENTOS_TABLE',
        'ZMED95_COLUMNS_IDS',
        'ZMED95_COLUMNS_NAMES',
        'ZMED95_COLUMNS_TYPES'
      )
    raise SomethingGoesWrong('Probably the lack of a flag made you fall here')
  def FPL9(
    self,
    instalacao: InstalacaoInfo,
    flags: list[FPL9_FLAGS] = [FPL9_FLAGS.GET_PENDING]
    ) -> pandas.DataFrame:
    ''' Function that get information about pending invoice report
        throught `FPL9` transaction when `ZARC140` is unavaliable '''
    self.session.StartTransaction(Transaction="FPL9")
    self.CHECK_STATUSBAR()
    self.session.findById(STRINGPATH['FPL9_PARCEIRO_INPUT']).text = instalacao.parceiro
    self.session.findById(STRINGPATH['FPL9_CONTRATO_INPUT']).text = instalacao.contrato
    self.session.findById(STRINGPATH['GLOBAL_ENTER_BUTTON']).Press()
    self.session.findById(STRINGPATH['FPL9_EXPAND_BUTTON']).Press()
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
    colunas = ['#', 'Referencia', 'Documento', 'Vencimento', 'Valor', 'Observacao']
    dataframe = {key: [] for key in colunas}
    while hasLines:
      # Aponta para a linha atual ou para a última caso o índice da linha ultrapasse o máximo
      linha = MAX_ROW if row >= MAX_ROW else row
      # Rola a barra vertical caso o índice da linha tenha ultrapassado o máximo
      if row >= MAX_ROW:
        self.session.FindById(STRINGPATH['GLOBAL_USER_AREA']).verticalScrollbar.position = row - MAX_ROW
      # Obtém o valor do label na coluna 1 da linha atual
      primeiro_caractere = self.GETBY_XY('FPL9_LABEL_CANVAS', 1, linha, False)
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
        label = self.GETBY_XY('FPL9_LABEL_CANVAS', col, linha, False)
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
        
        status_icon_name = self.GETBY_XY('FPL9_LABEL_CANVAS', indices[1], linha).iconName
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
        dataframe["Referencia"].append(self.GETBY_XY('FPL9_LABEL_CANVAS', indices[2], linha).text)
        dataframe["Documento"].append(conversor['numero'](self.GETBY_XY('FPL9_LABEL_CANVAS', indices[3], linha).text))
        dataframe["Vencimento"].append(conversor['data'](self.GETBY_XY('FPL9_LABEL_CANVAS', indices[4], linha).text))
        dataframe["Valor"].append(conversor['decimal'](self.GETBY_XY('FPL9_LABEL_CANVAS', indices[5], linha).text))
      row += 1
      col = 1
    dataframe1 = pandas.DataFrame(dataframe)
    # agrupa os valores por documento de impressão
    dataframe2 = dataframe1.groupby('Documento')['Valor'].sum().reset_index()
    # remove as duplicatas para ter somente um documento de impressão por linha
    dataframe1.drop_duplicates(subset="Documento", inplace=True)
    # mescla o dataframe com a soma dos valores com o dataframe com as informações
    dataframe3 = dataframe1.merge(dataframe2, on="Documento")
    # remove a coluna antiga com o valor errado
    del dataframe3['Valor_x']
    dataframe3 = dataframe3.rename(columns={'Valor_y': 'Valor'})
    dataframe3['Documento'] = dataframe3['Documento'].replace('', pandas.NA)
    dataframe3 = dataframe3.dropna(subset=['Documento'])
    return dataframe3
  def BP(
    self,
    instalacao: InstalacaoInfo,
    flags: list[BP_FLAGS]
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
    if BP_FLAGS.GET_DOCS in flags:
      self.session.findById(STRINGPATH['BP_DOCS_CPF_TAB']).Select()
      parceiro.documento_tipo = self.session.findById(STRINGPATH['BP_DOCS_TIPO_TEXT']).text
      parceiro.documento_numero = self.session.findById(STRINGPATH['BP_DOCS_CPF_TEXT']).text
    if BP_FLAGS.GET_PHONES in flags:
      ESPACO_VAZIO = "______________________________"
      self.session.findById(STRINGPATH['BP_PHONE_TAB']).Select()
      for lista in {'BP_PHONE_LIST1', 'BP_PHONE_LIST2', 'BP_PHONE_LIST3', 'BP_PHONE_LIST4', 'BP_PHONE_LIST5'}:
        self.session.FindById(STRINGPATH[lista]).Press()
        if self.session.FindById(STRINGPATH['POPUP'], False) is None:
          continue
        for i in range(4):
          telefone = self.GETBY_XY('BP_PHONE_TEXT', 2, i).text
          if telefone == '' or telefone == ESPACO_VAZIO:
            continue
          parceiro.telefones.append(telefone)
        self.session.FindById(STRINGPATH['POPUP']).Close()
      parceiro.telefones = list(dict.fromkeys(parceiro.telefones))
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
    relatorio = self.GET_ROWS(
      'ZATC66_TABELA_RESULT',
      'ZATC66_COLUMNS_IDS',
      'ZATC66_COLUMNS_NAMES',
      'ZATC66_COLUMNS_TYPES'
    )
    relatorio['Texto registrador'] = relatorio['Reg'].apply(lambda x: 
        depara('medidor_registrador', str(x).zfill(2)) or 'Sem codigo do registrador')
    relatorio['Texto do leiturista'] = relatorio['Cod leit'].apply(lambda x: depara('leitura_codigo', x))
    relatorio['Texto tipo leitura'] = relatorio['Cod tipo'].apply(lambda x: depara('leitura_tipo', str(x).zfill(2)))
    relatorio['Texto motivo leitura'] = relatorio['Cod motivo'].apply(lambda x: depara('leitura_tipo', str(x).zfill(2)))
    leitura_anterior = relatorio["Leitura"].shift(-1)
    relatorio["Consumo"] = relatorio["Leitura"] - leitura_anterior
    return relatorio[['Mes ref','Data leit','Medidor','Leitura','Consumo','Reg','Texto registrador','Cod tipo','Texto tipo leitura','Cod motivo','Texto motivo leitura','Cod leit','Texto do leiturista']]
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
    dataframe = self.GET_ROWS(
      'IQ03_VARIOUS_TABLE',
      'IQ03_VARIOUS_COLUMNS_IDS',
      'IQ03_VARIOUS_COLUMNS_NAMES',
      'IQ03_VARIOUS_COLUMNS_TYPES'
    )
    if dataframe.shape[0] != 0:
      medidores_info = []
      for index, row in dataframe.iterrows():
        medidores_info.extend(self.IQ03(row['Material'], serial))
      return medidores_info
    medidor = MedidorInfo()
    medidor.serial = serial
    medidor.material = material
    medidor.texto_material = str(material) + ' - ' + depara("material_codigo", str(medidor.material)) or ''
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
    dataframe = self.GET_ROWS(
      'IQ03_LEITURAS_TABLE',
      'IQ03_LEITURAS_COLUMNS_IDS',
      'IQ03_LEITURAS_COLUMNS_NAMES',
      'IQ03_LEITURAS_COLUMNS_TYPES',
      0,
      12
    )
    if dataframe.shape[0] == 0:
      medidor.observacao = "Equipamento sem historico de leituras!"
      return [medidor]
    dataframe['Descricao'] = dataframe['Codigo'].apply(lambda x: depara('leitura_codigo', x))
    if medidor.leituras.shape[0] > 0:
      medidor.observacao = f"Codigos de leitura nas ultimas {medidor.leituras.shape[0]} leituras:"
    else:
      medidor.observacao = f"Sem codigos de leitura nas ultimas {medidor.leituras.shape[0]} leituras!"
    dataframe = dataframe[dataframe['Codigo'] != 0]
    medidor.leituras = dataframe
    return [medidor]
  def ZHISCON(self, cpf_cnpj: int) -> object:
    ''' Exibe informações do cliente pelo CPF/CNPJ '''
    numero = str(cpf_cnpj).zfill(11) if len(str(cpf_cnpj)) <= 11 else str(cpf_cnpj).zfill(14)
    self.session.StartTransaction(Transaction="ZHISCON")
    self.session.FindById(STRINGPATH['ZHISCON_CPFCNPJ_INPUT']).text = numero
    self.session.FindById(STRINGPATH['ZHISCON_PERIODO_INICIO']).text = ''
    self.session.FindById(STRINGPATH['ZHISCON_PERIODO_FINAL']).text = ''
    self.session.FindById(STRINGPATH['GLOBAL_ACCEPT_BUTTON']).Press()
    parceiros_lista = self.GET_ROWS(
      'ZHISCON_TABLE_RESULT',
      'ZHISCON_COLUMNS_IDS',
      'ZHISCON_COLUMNS_NAMES',
      'ZHISCON_COLUMNS_TYPES',
      0,
      0
    )
    if parceiros_lista.shape[0] == 0:
      raise InformationNotFound('Nao foram encontrados clientes!')
    return '\n\n'.join([row.__str__() for _, row in parceiros_lista.iterrows()])
