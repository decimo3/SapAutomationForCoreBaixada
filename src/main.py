#!/usr/bin/python
''' Module to wraper SAPGUI scripting engine automation '''
# coding: utf8
#region imports
import sys
import datetime
import pandas
from dateutil.relativedelta import relativedelta
from wrapper import SapBot
from helpers import depara
from constants import (
  SEPARADOR,
  NOTUSE,
)
from exceptions import (
  SomethingGoesWrong,
  UnavailableSap,
  ArgumentException,
  InformationNotFound,
  TooMannyRequests,
  UnavailableTransaction,
  WrapperBaseException
)
from enumerators import (
  BP_FLAGS,
  DESTAQUES,
  ES32_FLAGS,
  ES57_FLAGS,
  ES61_FLAGS,
  FPL9_FLAGS,
  IQ03_FLAGS,
  ZMED89_FLAGS,
  ZARC140_FLAGS,
  IW53_FLAGS,
  ZMED95_FLAGS,
)
from models import (
  InstalacaoInfo,
  LigacaoInfo,
  MedidorInfo,
  ParceiroInfo,
  ServicoInfo,
)
#endregion

def obter_servico_por_servico(robo: SapBot, numero_servico: int, flags: list[IW53_FLAGS]) -> ServicoInfo:
  ''' Obtém informações do serviço a partir do número do servico. '''
  servico = robo.IW53(numero_servico, flags)
  return servico

def obter_medidor_por_medidor(robo: SapBot, numero_serial: int, numero_material: int, flags: list[IQ03_FLAGS]) -> MedidorInfo:
  ''' Obtém informações do medidor a partir do número do medidor. '''
  medidor = robo.IQ03(numero_serial, numero_material, flags)
  return medidor[0]

def obter_instalacao_por_instalacao(robo: SapBot, numero_instalacao: int, flags: list[ES32_FLAGS]) -> InstalacaoInfo:
  ''' Obtem informações da instalação a partir do número da instalacao. '''
  return robo.ES32(numero_instalacao, flags)

def obter_medidor_por_servico(robo: SapBot, numero_servico: int, flags: list[IQ03_FLAGS]) -> MedidorInfo:
  ''' Obtém informações do medidor a partir do número de serviço. '''
  servico = obter_servico_por_servico(robo, numero_servico, [IW53_FLAGS.GET_INST])
  instalacao = obter_instalacao_por_instalacao(robo, servico.instalacao, [ES32_FLAGS.GET_METER])
  equipamento = instalacao.get_medidor()
  return obter_medidor_por_medidor(robo, equipamento.serial, equipamento.material, flags)

def obter_medidor_por_instalacao(robo: SapBot, numero_instalacao: int, flags: list[IQ03_FLAGS]) -> MedidorInfo:
  ''' Obtém informações do medidor a partir do número de instalação. '''
  instalacao = obter_instalacao_por_instalacao(robo, numero_instalacao, [ES32_FLAGS.GET_METER])
  equipamento = instalacao.get_medidor()
  return obter_medidor_por_medidor(robo, equipamento.serial, equipamento.material, flags)

def obter_historico(robo: SapBot, argumento: int, _layout: str = '/WILLIAM') -> pandas.DataFrame:
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.ONLY_INST])
  return robo.ZSVC20(
    instalacao = instalacao_info,
    tipos_notas = [''],
    min_data = datetime.date.today() - datetime.timedelta(days = 90),
    max_data = datetime.date.today(),
    danos_filtro = [''],
    statuses = [''],
    regional = '',
    layout = _layout
  )

def obter_servico_por_instalacao(robo: SapBot, numero_instalacao: int, flags: list[IW53_FLAGS]) -> ServicoInfo:
  ''' Obtém informações do servico a partir do número da instalacao '''
  historico = obter_historico(robo, numero_instalacao)
  return obter_servico_por_servico(robo, historico.at[0, 'Nota'], flags)

def obter_servico_por_medidor(robo: SapBot, numero_medidor: int, flags: list[IW53_FLAGS]) -> ServicoInfo:
  ''' Obtém informações do servico a partir do número de medidor. '''
  medidor = obter_medidor_por_medidor(robo, numero_medidor, 0, [IQ03_FLAGS.ONLY_INST])
  return obter_servico_por_instalacao(robo, medidor.instalacao, flags)

def obter_instalacao_por_servico(robo: SapBot, numero_servico: int, flags: list[ES32_FLAGS]) -> InstalacaoInfo:
  ''' Obtém informações do servico a partir do número de serviço. '''
  servico = obter_servico_por_servico(robo, numero_servico, [IW53_FLAGS.GET_INST])
  return obter_instalacao_por_instalacao(robo, servico.instalacao, flags)

def obter_instalacao_por_medidor(robo: SapBot, numero_medidor: int, flags: list[ES32_FLAGS]) -> InstalacaoInfo:
  ''' Obtém informações do servico a partir do número do medidor. '''
  medidor = obter_medidor_por_medidor(robo, numero_medidor, 0, [IQ03_FLAGS.ONLY_INST])
  return obter_instalacao_por_instalacao(robo, medidor.instalacao, flags)

def obter_instalacao(robo: SapBot, argumento: int, flags: list[ES32_FLAGS] | None = None) -> InstalacaoInfo:
  if flags is None:
    flags = [ES32_FLAGS.GET_METER]
  if argumento > 999999999:
    return obter_instalacao_por_servico(robo, argumento, flags)
  if argumento < 99999999:
    return obter_instalacao_por_medidor(robo, argumento, flags)
  return obter_instalacao_por_instalacao(robo, argumento, flags)

def obter_medidor(robo: SapBot, argumento: int, flags: list[IQ03_FLAGS] | None = None) -> MedidorInfo:
  if flags is None:
    flags = [IQ03_FLAGS.READ_REPORT]
  if argumento > 999999999:
    return obter_medidor_por_servico(robo, argumento, flags)
  if argumento < 99999999:
    return obter_medidor_por_medidor(robo, argumento, 0, flags)
  return obter_medidor_por_instalacao(robo, argumento, flags)

def obter_servico(robo: SapBot, argumento: int, flags: list[IW53_FLAGS] | None = None) -> ServicoInfo:
  if flags is None:
    flags = [IW53_FLAGS.GET_INFO]
  if argumento > 999999999:
    return obter_servico_por_servico(robo, argumento, flags)
  if argumento < 99999999:
    return obter_servico_por_medidor(robo, argumento, flags)
  return obter_servico_por_instalacao(robo, argumento, flags)

def obter_religacao(robo: SapBot, argumento: int) -> pandas.DataFrame:
  if argumento > 90:
    raise ArgumentException('O numero de dias eh superior ao permitido!')
  data_inicio = datetime.date.today() - datetime.timedelta(days=argumento)
  return robo.ZSVC20(
    tipos_notas = ['B1', 'BL', 'BR'],
    min_data = data_inicio,
    max_data = datetime.date.today(),
    statuses = ['ENVI', 'LIBE', 'TABL'],
    danos_filtro = [],
    regional = 'RO',
    layout = '//TT'
  )

def obter_bandeirada(robo: SapBot, argumento: int) -> pandas.DataFrame:
  if argumento > 90:
    raise ArgumentException('O numero de dias eh superior ao permitido!')
  data_inicio = datetime.date.today() - datetime.timedelta(days=argumento)
  return robo.ZSVC20(
    tipos_notas = ['BA'],
    min_data = data_inicio,
    max_data = datetime.date.today(),
    danos_filtro = ['OSTA', 'OSJD', 'OSFT', 'OSAT', 'OSAR', 'OATI'],
    statuses = ['ANAL', 'POSB', 'PCOM', 'NEXE', 'EXEC'],
    regional = 'RO',
    layout = '/WILLIAM'
  )

def obter_lideanexo(robo: SapBot, argumento: int) -> pandas.DataFrame:
  if argumento > 90:
    raise ArgumentException('O numero de dias eh superior ao permitido!')
  data_inicio = datetime.date.today() - datetime.timedelta(days=argumento)
  relatorio = robo.ZSVC20(
    tipos_notas = ['B5', 'B8', 'BA', 'BC', 'BN', 'BS', 'BV'],
    min_data = data_inicio,
    max_data = datetime.date.today(),
    danos_filtro = [],
    statuses = ['ENVI', 'LIBE', 'TABL'],
    regional = 'RO',
    layout = '//TT'
  )
  filtrar_danos = depara('relatorio_filtro', 'LIDE').split(',')
  filtrar_danos.extend(depara('relatorio_filtro', 'ANEXO').split(','))
  relatorio = relatorio[~relatorio['Dano'].isin(filtrar_danos)]
  return relatorio

def obter_pendente(robo: SapBot, argumento: int, raise_error: bool = True) -> pandas.DataFrame:
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.ONLY_INST])
  # Getting the pending invoice report
  if 'ZARC140' not in NOTUSE:
    flags = [ZARC140_FLAGS.GET_PENDING]
    if not raise_error:
      flags.extend([ZARC140_FLAGS.DONOT_THROW])
    return robo.ZARC140(
      instalacao = instalacao_info,
      flags = flags
    )
  if 'FPL9' not in NOTUSE:
    flags = [FPL9_FLAGS.GET_PENDING]
    if not raise_error:
      flags.extend([FPL9_FLAGS.DONOT_THROW])
    return robo.FPL9(
      instalacao = instalacao_info,
      flags = flags
    )
  raise UnavailableTransaction('Sem acesso a transacao no sistema SAP!')

def print_pendentes(robo: SapBot, documentos: list[int], instalacao: InstalacaoInfo) -> int:
  if 'ZATC73' not in NOTUSE:
    robo.ZATC73(
      documentos = documentos
    )
  elif 'ZATC45' not in NOTUSE:
    robo.ZATC45(
      instalacao = instalacao,
      documentos = documentos
    )
  else:
    raise UnavailableTransaction('Sem acesso a transacao no sistema SAP!')
  return len(documentos)

def obter_ligacao(robo: SapBot, instalacao: InstalacaoInfo, _flags: list[ES61_FLAGS] | None = None) -> LigacaoInfo:
  if _flags == None:
    _flags = [ES61_FLAGS.ENTER_ENTER]
  if ES61_FLAGS.SKIPT_ENTER in _flags:
    return robo.ES6X(
      instalacao = instalacao,
      flags = _flags,
      transaction = ''
    )
  elif 'ES61' not in NOTUSE:
    return robo.ES61(
      instalacao = instalacao,
      flags = _flags
    )
  elif 'ES62' not in NOTUSE:
    return robo.ES62(
      instalacao = instalacao,
      flags = _flags
    )
  else:
    raise UnavailableTransaction('Sem acesso a transacao no sistema SAP!')

def obter_faturas(robo: SapBot, argumento: int) -> int:
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.ONLY_INST])
  # Getting the pending invoice report
  relatorio = obter_pendente(robo, argumento, False)
  if relatorio.shape[0] == 0:
    raise InformationNotFound('Cliente nao possui faturas pendentes!')
  if relatorio.shape[0] > 6:
    raise TooMannyRequests(f'O cliente possui faturas demais ({relatorio.shape[0]})')
  # Printing pending invoices
  relatorio = relatorio[relatorio['#'] != DESTAQUES.AMARELO]
  return print_pendentes(robo, relatorio['Documento'].to_list(), instalacao_info)

def obter_parceiro(robo: SapBot, argumento: int, flags: list[BP_FLAGS]) -> ParceiroInfo:
  if 'BP' in NOTUSE:
    raise UnavailableTransaction('Sem acesso a transacao no sistema SAP!')
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.ONLY_INST])
  return robo.BP(instalacao_info, flags)

def obter_documento(robo: SapBot, argumento: int) -> ParceiroInfo:
  return obter_parceiro(robo, argumento, [BP_FLAGS.GET_DOCS])

def obter_telefone(robo: SapBot, argumento: int) -> ParceiroInfo:
  return obter_parceiro(robo, argumento, [BP_FLAGS.GET_PHONES])

def obter_coordenadas(robo: SapBot, argumento: int) -> str:
  flag = [ES32_FLAGS.ENTER_CONSUMO] if 'ES61' in NOTUSE else [ES32_FLAGS.ONLY_INST]
  instalacao_info = obter_instalacao(robo, argumento, flag)
  flag = [ES61_FLAGS.GET_COORD]
  flag.extend([ES61_FLAGS.SKIPT_ENTER]) if 'ES61' in NOTUSE else flag.extend([ES61_FLAGS.ENTER_ENTER])
  consumo_info = obter_ligacao(robo, instalacao_info, flag)
  if not consumo_info.coordenadas:
    raise InformationNotFound('A instalacao nao possui coordenada cadastrada!')
  return consumo_info.coordenadas

def obter_leiturista(robo: SapBot, argumento: int, _flags: list[ZMED89_FLAGS]) -> pandas.DataFrame:
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.GET_CENTER])
  if not instalacao_info.tipo_instalacao in {'', 'B'}:
    _flags.extend([ZMED89_FLAGS.TELEMEDIDO])
  return robo.ZMED89(
    instalacao = instalacao_info,
    flags = _flags,
    quantidade = 30,
  )

def obter_sequencial(robo: SapBot, argumento: int) -> pandas.DataFrame:
  return obter_leiturista(robo, argumento, [ZMED89_FLAGS.SEQ_ORDER])

def obter_horariado(robo: SapBot, argumento: int) -> pandas.DataFrame:
  return obter_leiturista(robo, argumento, [ZMED89_FLAGS.TIME_ORDER])

def obter_agrupamento(robo: SapBot, argumento: int) -> pandas.DataFrame:
  flag = [ES32_FLAGS.ENTER_CONSUMO] if 'ES61' in NOTUSE else [ES32_FLAGS.ONLY_INST]
  instalacao_info = obter_instalacao(robo, argumento, flag)
  flag = [ES61_FLAGS.SKIPT_ENTER] if 'ES61' in NOTUSE else [ES61_FLAGS.ENTER_ENTER]
  if 'ES57' in NOTUSE:
    flag.extend([ES61_FLAGS.ENTER_LIGACAO])
  ligacao_info = obter_ligacao(robo, instalacao_info, flag)
  flag = [ES57_FLAGS.SKIPT_ENTER] if 'ES57' in NOTUSE else [ES57_FLAGS.ENTER_ENTER]
  logradouro_info = robo.ES57(ligacao_info, flag)
  flag = [ZMED95_FLAGS.SKIPT_ENTER] if 'ZMED95' in NOTUSE else [ZMED95_FLAGS.ENTER_ENTER]
  flag.extend([ZMED95_FLAGS.GET_GROUPING])
  return robo.ZMED95(logradouro_info, flag)

def obter_consumo(robo: SapBot, argumento: int) -> pandas.DataFrame:
  if 'ZATC66' in NOTUSE:
    raise UnavailableTransaction('Sem acesso a transacao no sistema SAP!')
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.ONLY_INST])
  return robo.ZATC66(instalacao_info)

def obter_consumos(robo: SapBot, argumento: int) -> pandas.DataFrame:
  if 'ZATC66' in NOTUSE:
    raise UnavailableTransaction('Sem acesso a transacao no sistema SAP!')
  consumos = pandas.DataFrame()
  # Get reader data
  leiturista = obter_leiturista(robo, argumento, [ZMED89_FLAGS.TIME_ORDER])
  # Trim excessive data if needed
  if leiturista.shape[0] >= 30:
    offset = leiturista.shape[0] // 3
    leiturista = leiturista.iloc[offset:-offset]
  # Convert columns to numeric types
  leiturista['#'] = [0 for i in range(leiturista.shape[0])]
  leiturista['Medidor'] = pandas.to_numeric(leiturista['Medidor'], 'coerce').astype('Int64')
  leiturista['Instalacao'] = pandas.to_numeric(leiturista['Instalacao'], 'coerce').astype('Int64')
  # Generate the last 12 months formatted as "MMM.YY" (e.g., "fev.25")
  meses_de_referencia = [
      (datetime.datetime.today().replace(day=1) - relativedelta(months=i)).strftime('%b.%y')
      for i in range(12)
  ]
  # Get all consumption data
  for _, leitura in leiturista.iterrows():
    try:
      consumo = obter_consumo(robo, leitura['Instalacao'])
      consumos = pandas.concat([consumos, consumo], ignore_index=True)
    except:
      ...
  if consumos.shape[0] == 0:
    raise InformationNotFound('Nao foi possivel obter consumos!')
  # Pivot the consumption data
  consumos = consumos.drop_duplicates(subset=['Medidor', 'Mes ref.'])
  df_pivot = consumos.pivot(index='Medidor', columns='Mes ref.', values='Consumo')
  # Rename columns to desired format (e.g., "fev.25" for "2025-02"), handling NaN values safely
  df_pivot.columns = pandas.to_datetime(df_pivot.columns, format='%m/%Y', errors='coerce').strftime('%b.%y')
  df_pivot.columns = [col if pandas.notna(col) else "Unknown" for col in df_pivot.columns]
  # Filter columns to match the last 12 months (ensure all months are present in order)
  df_pivot = df_pivot.reindex(columns=meses_de_referencia, fill_value=0)
  # Reset index to merge
  df_pivot.reset_index(inplace=True)
  # Merge with consumer information
  df_final = pandas.merge(leiturista, df_pivot, on='Medidor', how='left')
  # Sort columns in reverse order (latest month first)
  df_final = df_final[['#', 'Endereco', 'Instalacao', 'Medidor'] + meses_de_referencia]
  return df_final

def obter_debitos_passiveis(robo: SapBot, argumento: int, filtrar_passivas: bool = True) -> int:
  debito = obter_pendente(robo, argumento, False)
  if debito.empty:
    return 0
  if filtrar_passivas:
    debito = debito[debito['#'] == DESTAQUES.VERMELHO]
    prazo_minimo = datetime.date.today() - datetime.timedelta(days=15)
    prazo_maximo = datetime.date.today() - datetime.timedelta(days=90)
    debito = debito[debito['Vencimento'] < pandas.to_datetime(prazo_minimo)]
    debito = debito[debito['Vencimento'] > pandas.to_datetime(prazo_maximo)]
  return debito['Valor'].sum()

def obter_passivo_corte(robo: SapBot, argumento: int) -> str:
  try:
    instalacao = obter_instalacao(robo, argumento, [ES32_FLAGS.GET_METER])
  except InformationNotFound as erro:
    if erro.message == 'Instalacao nao possui medidor!':
      return erro.message
  except Exception as erro:
    raise erro
  if instalacao.status == ' Instalação complet.suspensa':
    return 'Ja suspenso no sistema!'
  if instalacao.contrato == 0:
    return 'Instalacao sem contrato!'
  if not instalacao.nome_cliente:
    return 'Instalacao sem parceiro!'
  if instalacao.nome_cliente.startswith("UNIDADE C/ CONSUMO"):
    return 'Instalacao sem parceiro!'
  if instalacao.nome_cliente.startswith("PARCEIRO DE NEGOCIO"):
    return 'Instalacao sem parceiro!'
  try:
    instalacao.get_medidor()
  except InformationNotFound as erro:
    return 'Instalacao sem medidor!'
  except:
    ...
  # TODO - Verificar se baixa renda é critério para passividade
  # if 'Baixa Renda' in instalacao.texto_classe:
  #   return 'Instalacao nao passivel'
  try:
    if obter_debitos_passiveis(robo, instalacao.instalacao) > 0:
      return 'Debitos passiveis de corte!'
  except:
    ...
  return 'Instalacao nao passivel'

def obter_devedores(robo: SapBot, argumento: int) -> pandas.DataFrame:
  passividade = pandas.DataFrame({ 'Instalacao': [], 'Observacao': [] })
  agrupamentos = obter_agrupamento(robo, argumento)
  agrupamentos['Instalacao'] = pandas.to_numeric(agrupamentos['Instalacao'], 'coerce').astype('Int64')
  for _, agrupamento in agrupamentos.iterrows():
    instalacao = agrupamento['Instalacao']
    passivo = obter_passivo_corte(robo, instalacao)
    status = DESTAQUES.VERDE if passivo == 'Instalacao nao passivel' else DESTAQUES.VERMELHO
    linha = [{
      '#': status,
      'Instalacao': instalacao,
      'Observacao': passivo
    }]
    passividade = pandas.concat([passividade, pandas.DataFrame(linha)], ignore_index=True)
  agrupamentos = agrupamentos.merge(passividade, how='left', on='Instalacao')
  agrupamentos = agrupamentos[['#'] + [col for col in agrupamentos.columns if col != '#']]
  return agrupamentos

def obter_fugitivos(robo: SapBot, argumento: int) -> pandas.DataFrame:
  passividade = pandas.DataFrame({ 'Instalacao': [], 'Pendente': [] })
  agrupamentos = obter_agrupamento(robo, argumento)
  agrupamentos['Instalacao'] = pandas.to_numeric(agrupamentos['Instalacao'], 'coerce').astype('Int64')
  for _, agrupamento in agrupamentos.iterrows():
    instalacao = agrupamento['Instalacao']
    try:
      passivo = obter_debitos_passiveis(robo, instalacao, False)
    except:
      passivo = 0
    linha = [{
      '#': DESTAQUES.AUSENTE,
      'Instalacao': instalacao,
      'Valor': passivo
    }]
    passividade = pandas.concat([passividade, pandas.DataFrame(linha)], ignore_index=True)
  agrupamentos = agrupamentos.merge(passividade, how='left', on='Instalacao')
  agrupamentos = agrupamentos[['#'] + [col for col in agrupamentos.columns if col != '#']]
  return agrupamentos

def obter_cruzamento(robo: SapBot, argumento: int) -> pandas.DataFrame:
  flag = [ES32_FLAGS.ENTER_CONSUMO] if 'ES61' in NOTUSE else [ES32_FLAGS.ONLY_INST]
  instalacao_info = obter_instalacao(robo, argumento, flag)
  flag = [ES61_FLAGS.SKIPT_ENTER] if 'ES61' in NOTUSE else [ES61_FLAGS.ENTER_ENTER]
  if 'ES57' in NOTUSE:
    flag.extend([ES61_FLAGS.ENTER_LIGACAO])
  ligacao_info = robo.ES61(instalacao_info, flag)
  flag = [ES57_FLAGS.SKIPT_ENTER] if 'ES57' in NOTUSE else [ES57_FLAGS.ENTER_ENTER]
  logradouro_info = robo.ES57(ligacao_info, flag) 
  flag = [ZMED95_FLAGS.SKIPT_ENTER] if 'ZMED95' in NOTUSE else [ZMED95_FLAGS.ENTER_ENTER]
  flag.extend([ZMED95_FLAGS.GET_CROSSING])
  return robo.ZMED95(logradouro_info, flag)

def obter_informacao(robo: SapBot, argumento: int) -> str:
  if 'BP' in NOTUSE:
    raise UnavailableTransaction('Sem acesso a transacao no sistema SAP!')
  texto = ''
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.GET_METER, ES32_FLAGS.DONOT_THROW])
  if instalacao_info.parceiro:
    try:
      parceiro_info = obter_parceiro(robo, instalacao_info.instalacao, [BP_FLAGS.GET_DOCS])
      texto += parceiro_info.__str__() + '\n'
    except WrapperBaseException as erro:
      texto += erro.message
  texto += instalacao_info.__str__() + '\n'
  if instalacao_info.equipamento:
    try:
      medidor_info = instalacao_info.get_medidor()
      medidor_info = robo.IQ03(medidor_info.serial, medidor_info.material, [IQ03_FLAGS.READ_REPORT])[0]
      texto += medidor_info.__str__()
    except WrapperBaseException as erro:
      texto += erro.message
  output = []
  for linha in texto.split('\n'):
    if not linha.strip():
      continue
    if not linha in output:
      output.append(linha)
  return '\n'.join(output)

def checar_inspecao(robo: SapBot, argumento: int) -> str:
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.GET_METER, ES32_FLAGS.DONOT_THROW])
  texto = f'Instalacao {instalacao_info.instalacao} nao esta apta para abertura de nota de recuperacao devido '
  # Checking installation information
  if instalacao_info.status != ' Instalação não suspensa':
    raise InformationNotFound(texto + 'nao estar ativa!')
  if not instalacao_info.parceiro:
    raise InformationNotFound(texto + 'nao ter cliente vinculado!')
  if instalacao_info.nome_cliente.startswith('UNIDADE C/ CONSUMO'):
    raise InformationNotFound(texto + 'nao ter cliente vinculado!')
  if instalacao_info.nome_cliente.startswith('PARCEIRO DE NEGOCIO'):
    raise InformationNotFound(texto + 'nao ter cliente vinculado!')
  if instalacao_info.texto_classe.find('Baixa Renda') >= 0:
    raise InformationNotFound(texto + 'instalacao ser baixa renda!')
  # REMOVED - Checking restriction information
  # ligacao_info = obter_ligacao(robo, instalacao_info)
  # localidade = instalacao_info.unidade[2:6:1]
  # is_residencial = instalacao_info.classe > 1000 and instalacao_info.classe < 2000
  # if ligacao_info.tipo_instalacao == 1 and is_residencial and (localidade != 'L539' and localidade != 'L595'):
  #   raise InformationNotFound(texto + 'ser residencial em area restrita de inspecao pelo tipo de instalacao')
  # if ligacao_info.tipo_instalacao == 2 and is_residencial:
  #   raise InformationNotFound(texto + 'ser residencial em area restrita de inspecao pelo tipo de instalacao')
  # Checking measurement information
  if not instalacao_info.equipamento:
    raise InformationNotFound(texto + 'nao tem medidor vinculado!')
  medidor = instalacao_info.get_medidor() if len(instalacao_info.equipamento) > 1 else instalacao_info.equipamento[0]
  medidor_info = robo.IQ03(medidor.serial, medidor.material)[0]
  if medidor_info.code_status != 'INST':
    raise InformationNotFound(texto + f'medidor `{medidor_info.serial}` com status `{medidor_info.code_status}`!')
  # Checking customer information
  parceiro_info = obter_parceiro(robo, instalacao_info.instalacao, [BP_FLAGS.GET_DOCS])
  if not parceiro_info.documento_numero:
    raise InformationNotFound(texto + 'ao cliente nao ter CPF ou CNPJ no cadastro!')
  if parceiro_info.documento_tipo != 'Brasil: nº CPF' and parceiro_info.documento_tipo != 'Brasil: nº CNPJ':
    raise InformationNotFound(texto + 'ao cliente nao ter CPF ou CNPJ no cadastro!')
  # REMOVED - Checking outstanding debts information
  # if obter_debitos_passiveis(robo, instalacao_info.instalacao) > 0:
  #   raise InformationNotFound(texto + 'o cliente possuir debito(s) pendente(s)')
  # Checking if installation already has order
  meses_verificacao_inspecoes =  6  # if not is_residencial else 12
  prazo_maximo_verificacao = datetime.date.today() - datetime.timedelta(days=meses_verificacao_inspecoes * 30)
  historico_info = obter_historico(robo, instalacao_info.instalacao, '/VENCIMENTOS')
  historico_info = historico_info[pandas.to_datetime(historico_info["Data"]) >= pandas.to_datetime(prazo_maximo_verificacao)]
  historico_info = historico_info[(historico_info["Tipo"] == "BI") | (historico_info["Tipo"] == "BU")]
  historico_info = historico_info[historico_info["Status"] == "EXEC"]
  if historico_info.shape[0] > 0:
    raise InformationNotFound(texto + f'ja tem nota {historico_info['Nota'].to_string(index=False)} executada!')
  return f"A instalacao {instalacao_info.instalacao} esta apta sim para abertura de nota de recuperacao!"

def obter_cliente(robo: SapBot, argumento: int) -> object:
  if 'ZHISCON' in NOTUSE:
    raise UnavailableSap('Sem acesso a transacao no sistema SAP!')
  return robo.ZHISCON(argumento)

aplicacoes = {
  'cliente': obter_cliente,
  'instalacao': obter_instalacao,
  'servico': obter_servico,
  'medidor': obter_medidor,
  'telefone': obter_telefone,
  'contato': obter_telefone,
  'fatura': obter_faturas,
  'debito': obter_faturas,
  'coordenada': obter_coordenadas,
  'localizacao': obter_coordenadas,
  'religacao': obter_religacao,
  'bandeirada': obter_bandeirada,
  'historico': obter_historico,
  'leiturista': obter_horariado,
  'roteiro': obter_sequencial,
  'consumo': obter_consumo,
  'ren360': obter_consumos,
  'agrupamento': obter_devedores,
  'fuga': obter_fugitivos,
  'pendente': obter_pendente,
  'cruzamento': obter_cruzamento,
  'lideanexo': obter_lideanexo,
  'abertura': checar_inspecao,
  'informacao': obter_informacao,
  'leitura': obter_consumo,
}

if __name__ == '__main__':
  try:
    if len(sys.argv) < 4:
      raise ArgumentException('Falta argumentos necessarios!')
    aplicacao = sys.argv[1]
    argumento = int(sys.argv[2])
    instancia_argumento = int(sys.argv[3])
    # Attempts to connect to SAP FrontEnd on the specified instance
    robo = SapBot(instancia_argumento)
    if aplicacao == 'instancia':
      robo.create_session(argumento)
    else:
      robo.attach_session(instancia_argumento)
    # Attempts to execute the method requested in the first argument
    retorno = aplicacoes[aplicacao](robo, argumento)
    if isinstance(retorno, pandas.DataFrame):
      if '#' in retorno.columns:
        retorno['#'] = retorno['#'].astype(int)
      print(retorno.to_csv(index=False,sep=SEPARADOR))
    else:
      print(retorno)
    robo.HOME_PAGE()
  except KeyError as erro:
    print(f'400: aplicacao \'{erro.args[0]}\' desconhecida!')
  except ArgumentException as erro:
    print(f'400: {erro.message}')
  except InformationNotFound as erro:
    print(f'404: {erro.message}')
  except TooMannyRequests as erro:
    print(f'409: {erro.message}')
  except UnavailableSap as erro:
    print(f'408: {erro.message}')
  except UnavailableTransaction as erro:
    print(f'408: {erro.message}')
  except SomethingGoesWrong as erro:
    print(f'500: {erro.message}')
  except WrapperBaseException as erro:
    print(f'500: {erro.message}')
  except Exception as erro:
    print(f'500: Unhandled exception! {erro.args[0]}')
