#!/usr/bin/python
''' Module to wraper SAPGUI scripting engine automation '''
# coding: utf8
#region imports
import sys
import datetime
import pandas
from wrapper import SapBot
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
  WrapperBaseException
)
from enumerators import (
  BP_FLAGS,
  ES32_FLAGS,
  ES57_FLAGS,
  ES61_FLAGS,
  IQ03_FLAGS,
  ZMED89_FLAGS,
  ZARC140_FLAGS,
  IW53_FLAGS,
  ZMED95_FLAGS,
)
from models import (
  InstalacaoInfo,
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
  if len(instalacao.equipamento) > 1:
    raise TooMannyRequests("Instalação possui mais de um equipamento!")
  return obter_medidor_por_medidor(
    robo,
    instalacao.equipamento[0].serial,
    instalacao.equipamento[0].material,
    [IQ03_FLAGS.READ_REPORT])

def obter_medidor_por_instalacao(robo: SapBot, numero_instalacao: int, flags: list[IQ03_FLAGS]) -> MedidorInfo:
  ''' Obtém informações do medidor a partir do número de instalação. '''
  instalacao = obter_instalacao_por_instalacao(robo, numero_instalacao, [ES32_FLAGS.GET_METER])
  equipamento = instalacao.get_medidor()
  return obter_medidor_por_medidor(robo, equipamento.serial, equipamento.material, flags)

def obter_historico(robo: SapBot, argumento: int) -> pandas.DataFrame:
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.ONLY_INST])
  return robo.ZSVC20(
    instalacao = instalacao_info,
    tipos_notas = [''],
    min_data = datetime.date.today() - datetime.timedelta(days = 90),
    max_data = datetime.date.today(),
    danos_filtro = [''],
    statuses = [''],
    regional = '',
    layout = '/WILLIAN',
    colluns = ['QMNUM', 'QMART', 'KURZTEXT', 'MATXT', 'AUSBS', 'ZZ_ST_USUARIO'],
    colluns_names = ['Nota', 'Tipo', 'Texto do dano', 'Texto do code', 'Data', 'Status']
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

def obter_instalacao(robo: SapBot, argumento: int, flags: list[ES32_FLAGS] = [ES32_FLAGS.GET_METER]) -> InstalacaoInfo:
  if argumento > 999999999:
    return obter_instalacao_por_servico(robo, argumento, flags)
  elif argumento < 99999999:
    return obter_instalacao_por_medidor(robo, argumento, flags)
  else:
    return obter_instalacao_por_instalacao(robo, argumento, flags)

def obter_medidor(robo: SapBot, argumento: int, flags: list[IQ03_FLAGS] = [IQ03_FLAGS.READ_REPORT]) -> MedidorInfo:
  if argumento > 999999999:
    return obter_medidor_por_servico(robo, argumento, flags)
  elif argumento < 99999999:
    return obter_medidor_por_medidor(robo, argumento, 0, flags)
  else:
    return obter_medidor_por_instalacao(robo, argumento, flags)

def obter_servico(robo: SapBot, argumento: int, flags: list[IW53_FLAGS] = [IW53_FLAGS.GET_INFO]) -> ServicoInfo:
  if argumento > 999999999:
    return obter_servico_por_servico(robo, argumento, flags)
  elif argumento < 99999999:
    return obter_servico_por_medidor(robo, argumento, flags)
  else:
    return obter_servico_por_instalacao(robo, argumento, flags)

def obter_religacao(robo: SapBot, argumento: int) -> pandas.DataFrame:
  if argumento > 90:
    raise ArgumentException('O numero de dias eh superior ao permitido!')
  if 'ZSVC20' in NOTUSE:
    raise UnavailableSap('A aplicacao necessaria esta indisponivel!')
  data_inicio = datetime.date.today() - datetime.timedelta(days=argumento)
  return robo.ZSVC20(
    tipos_notas = ['B1', 'BL', 'BR'],
    min_data = data_inicio,
    max_data = datetime.date.today(),
    statuses = ['ENVI', 'LIBE', 'TABL'],
    danos_filtro = [],
    regional = 'RO',
    layout = '/VENCIMENTOS',
    colluns = ['QMNUM', 'ZZINSTLN', 'QMART', 'FECOD', 'LTRMN', 'LTRUR', 'ZZ_ST_USUARIO', 'QMDAB'],
    colluns_names = ['Nota', 'Instalacao', 'Tipo', 'Dano', 'Data', 'Hora', 'Status', 'Fim avaria']
  )

def obter_bandeirada(robo: SapBot, argumento: int) -> pandas.DataFrame:
  if argumento > 90:
    raise ArgumentException('O numero de dias eh superior ao permitido!')
  if 'ZSVC20' in NOTUSE:
    raise UnavailableSap('A aplicacao necessaria esta indisponivel!')
  data_inicio = datetime.date.today() - datetime.timedelta(days=argumento)
  return robo.ZSVC20(
    tipos_notas = ['BA'],
    min_data = data_inicio,
    max_data = datetime.date.today(),
    danos_filtro = ['OSTA', 'OSJD', 'OSFT', 'OSAT', 'OSAR', 'OATI'],
    statuses = ['ANAL', 'POSB', 'PCOM', 'NEXE', 'EXEC'],
    regional = 'RO',
    layout = '/VENCIMENTOS',
    colluns = ['QMNUM', 'ZZINSTLN', 'QMART', 'FECOD', 'LTRMN', 'LTRUR', 'ZZ_ST_USUARIO', 'QMDAB'],
    colluns_names = ['Nota', 'Instalacao', 'Tipo', 'Dano', 'Data', 'Hora', 'Status', 'Fim avaria']
  )

def obter_pendente(robo: SapBot, argumento: int) -> pandas.DataFrame:
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.ONLY_INST])
  # Getting the pending invoice report
  if 'ZARC140' not in NOTUSE:
    return robo.ZARC140(
      instalacao = instalacao_info,
      flag = ZARC140_FLAGS.GET_PENDING
    )
  elif 'FPL9' not in NOTUSE:
    return robo.FPL9(
      instalacao = instalacao_info
    )
  else:
    raise UnavailableSap('Sem acesso a transacao no sistema SAP!')

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
    raise UnavailableSap('Sem acesso a transacao no sistema SAP!')
  return len(documentos)

def obter_faturas(robo: SapBot, argumento: int) -> int:
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.ONLY_INST])
  # Getting the pending invoice report
  relatorio = obter_pendente(robo, argumento)
  # Printing pending invoices
  return print_pendentes(robo, relatorio['Documento'].to_list(), instalacao_info)

def obter_parceiro(robo: SapBot, argumento: int, flags: list[BP_FLAGS]) -> ParceiroInfo:
  if 'BP' in NOTUSE:
    raise UnavailableSap('Sem acesso a transacao no sistema SAP!')
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
  consumo_info = robo.ES61(instalacao_info, flag)
  if not consumo_info.coordenadas:
    raise InformationNotFound('A instalacao nao possui coordenada cadastrada!')
  return consumo_info.coordenadas

def obter_leiturista(robo: SapBot, argumento: int, _flag: ZMED89_FLAGS) -> pandas.DataFrame:
  instalacao_info = obter_instalacao(robo, argumento, [ES32_FLAGS.GET_CENTER])
  return robo.ZMED89(
    instalacao = instalacao_info,
    quantidade = 30,
    collumns = ['ZZ_NUMSEQ', 'ANLAGE', 'ZENDERECO', 'BAIRRO', 'GERAET', 'ZHORALEIT', 'ABLHINW'],
    collumns_names = ['Seq', 'Instalacao', 'Endereco', 'Bairro', 'Medidor', 'Horario', 'Codigo'],
    flag = _flag
  )

def obter_sequencial(robo: SapBot, argumento: int) -> pandas.DataFrame:
  return obter_leiturista(robo, argumento, ZMED89_FLAGS.SEQ_ORDER)

def obter_horariado(robo: SapBot, argumento: int) -> pandas.DataFrame:
  return obter_leiturista(robo, argumento, ZMED89_FLAGS.TIME_ORDER)

def obter_agrupamento(robo: SapBot, argumento: int) -> pandas.DataFrame:
  flag = [ES32_FLAGS.ENTER_CONSUMO] if 'ES61' in NOTUSE else [ES32_FLAGS.ONLY_INST]
  instalacao_info = obter_instalacao(robo, argumento, flag)
  flag = [ES61_FLAGS.SKIPT_ENTER] if 'ES61' in NOTUSE else [ES61_FLAGS.ENTER_ENTER]
  if 'ES57' in NOTUSE:
    flag.extend([ES61_FLAGS.ENTER_LIGACAO])
  ligacao_info = robo.ES61(instalacao_info, flag)
  flag = [ES57_FLAGS.SKIPT_ENTER] if 'ES57' in NOTUSE else [ES57_FLAGS.ENTER_ENTER]
  logradouro_info = robo.ES57(ligacao_info, flag)
  flag = [ZMED95_FLAGS.SKIPT_ENTER] if 'ZMED95' in NOTUSE else [ZMED95_FLAGS.ENTER_ENTER]
  return robo.ZMED95(logradouro_info, flag)

aplicacoes = {
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
  'informacao': obter_documento,
  'leiturista': obter_horariado,
  'roteiro': obter_sequencial,
  # 'consumo': obter_consumo,
  # 'ren360': obter_consumos,
  # 'agrupamento': obter_devedores,
  # 'fuga': obter_devedores,
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
      print(retorno.to_csv(index=False,sep=SEPARADOR))
    else:
      print(retorno)
    robo.HOME_PAGE()
  except ArgumentException as erro:
    print(f'400: {erro.message}')
  except InformationNotFound as erro:
    print(f'404: {erro.message}')
  except TooMannyRequests as erro:
    print(f'409: {erro.message}')
  except UnavailableSap as erro:
    print(f'408: {erro.message}')
  except SomethingGoesWrong as erro:
    print(f'500: {erro.message}')
  except WrapperBaseException as erro:
    print(f'500: {erro.message}')
  except Exception as erro:
    print(f'500: Something goes wrong! {erro.args[0]}')
