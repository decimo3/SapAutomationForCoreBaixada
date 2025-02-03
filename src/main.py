#!/usr/bin/python
''' Module to wraper SAPGUI scripting engine automation '''
# coding: utf8
#region imports
import sys
import datetime
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
  ES32_FLAGS,
  ES61_FLAGS,
  ZMED89_FLAGS,
  ZARC140_FLAGS,
)
#endregion

if __name__ == '__main__':
  try:
    if len(sys.argv) < 4:
      raise ArgumentException('Falta argumentos necessarios!')
    aplicacao = sys.argv[1]
    argumento = int(sys.argv[2])
    instancia_argumento = int(sys.argv[3])
    # Attempts to connect to SAP FrontEnd on the specified instance
    robo = SapBot(instancia_argumento)
    # Attempts to execute the method requested in the first argument
    if aplicacao == 'instancia':
      if instancia_argumento != 0:
        raise ArgumentException('A instancia para essa aplicacao dever ser zero!')
      robo.create_session()
      sys.exit()
    robo.attach_session(instancia_argumento)
    if aplicacao == 'vencimento':
      if argumento > 90:
        raise ArgumentException('O numero de dias eh superior ao permitido!')
      if 'ZSVC20' in NOTUSE:
        raise UnavailableSap('A aplicacao necessaria esta indisponivel!')
      data_inicio = datetime.date.today() - datetime.timedelta(days=argumento)
      relatorio = robo.ZSVC20(
        tipos_notas = ['B1', 'BL', 'BR'],
        min_data = data_inicio,
        max_data = datetime.date.today(),
        statuses = ['ENVI', 'LIBE', 'TABL'],
        danos_filtro = [],
        regional = 'RB',
        layout = '/VENCIMENTOS',
        colluns = ['QMNUM', 'ZZINSTLN', 'QMART', 'FECOD', 'LTRMN', 'LTRUR', 'ZZ_ST_USUARIO', 'QMDAB'],
        colluns_names = ['Nota', 'Instalacao', 'Tipo', 'Dano', 'Data', 'Hora', 'Status', 'Fim avaria']
      )
      print(relatorio.to_csv(index=False,sep=SEPARADOR))
    elif aplicacao == 'bandeirada':
      if argumento > 90:
        raise ArgumentException('O numero de dias eh superior ao permitido!')
      if 'ZSVC20' in NOTUSE:
        raise UnavailableSap('A aplicacao necessaria esta indisponivel!')
      data_inicio = datetime.date.today() - datetime.timedelta(days=argumento)
      relatorio = robo.ZSVC20(
        tipos_notas = ['BA'],
        min_data = data_inicio,
        max_data = datetime.date.today(),
        danos_filtro = ['OSTA', 'OSJD', 'OSFT', 'OSAT', 'OSAR', 'OATI'],
        statuses = ['ANAL', 'POSB', 'PCOM', 'NEXE', 'EXEC'],
        regional = 'RB',
        layout = '/VENCIMENTOS',
        colluns = ['QMNUM', 'ZZINSTLN', 'QMART', 'FECOD', 'LTRMN', 'LTRUR', 'ZZ_ST_USUARIO', 'QMDAB'],
        colluns_names = ['Nota', 'Instalacao', 'Tipo', 'Dano', 'Data', 'Hora', 'Status', 'Fim avaria']
      )
      print(relatorio.to_csv(index=False,sep=SEPARADOR))
    elif aplicacao == 'historico':
      instalacao_info = robo.ES32(argumento, ES32_FLAGS.ONLY_INST)
      relatorio = robo.ZSVC20(
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
      print(relatorio.to_csv(index=False,sep=SEPARADOR))
    elif aplicacao == 'leiturista':
      instalacao_info = robo.ES32(argumento, ES32_FLAGS.GET_CENTER)
      relatorio = robo.ZMED89(
        instalacao = instalacao_info,
        quantidade = 30,
        collumns = ['ZZ_NUMSEQ', 'ANLAGE', 'ZENDERECO', 'BAIRRO', 'GERAET', 'ZHORALEIT', 'ABLHINW'],
        collumns_names = ['Seq', 'Instalacao', 'Endereco', 'Bairro', 'Medidor', 'Horario', 'Codigo'],
        flag = ZMED89_FLAGS.TIME_ORDER
      )
      print(relatorio.to_csv(index=False,sep=SEPARADOR))
    elif aplicacao == 'roteiro':
      instalacao_info = robo.ES32(argumento, ES32_FLAGS.GET_CENTER)
      relatorio = robo.ZMED89(
        instalacao = instalacao_info,
        quantidade = 30,
        collumns = ['ZZ_NUMSEQ', 'ANLAGE', 'ZENDERECO', 'BAIRRO', 'GERAET', 'ZHORALEIT', 'ABLHINW'],
        collumns_names = ['Seq', 'Instalacao', 'Endereco', 'Bairro', 'Medidor', 'Horario', 'Codigo'],
        flag = ZMED89_FLAGS.SEQ_ORDER
      )
      print(relatorio.to_csv(index=False,sep=SEPARADOR))
    elif aplicacao == 'pendente':
      instalacao_info = robo.ES32(argumento, ES32_FLAGS.ONLY_INST)
      # Getting the pending invoice report
      if 'ZARC140' not in NOTUSE:
        relatorio = robo.ZARC140(
          instalacao = instalacao_info,
          flag = ZARC140_FLAGS.GET_PENDING
        )
      elif 'FPL9' not in NOTUSE:
        relatorio = robo.FPL9(
          instalacao = instalacao_info
        )
      else:
        raise UnavailableSap('Sem acesso a transacao no sistema SAP!')
      print(relatorio.to_csv(index=False,sep=SEPARADOR))
    elif aplicacao in {'fatura', 'debito'}:
      instalacao_info = robo.ES32(argumento, ES32_FLAGS.ONLY_INST)
      # Getting the pending invoice report
      if 'ZARC140' not in NOTUSE:
        relatorio = robo.ZARC140(
          instalacao = instalacao_info,
          flag = ZARC140_FLAGS.GET_PENDING
        )
      elif 'FPL9' not in NOTUSE:
        relatorio = robo.FPL9(
          instalacao = instalacao_info
        )
      else:
        raise UnavailableSap('Sem acesso a transacao no sistema SAP!')
      # Printing pending invoices
      if 'ZATC73' not in NOTUSE:
        robo.ZATC73(
          documentos = relatorio['documentos'].to_list()
        )
      elif 'ZATC45' not in NOTUSE:
        robo.ZATC45(
          instalacao = instalacao_info,
          documentos = relatorio['documentos'].to_list()
        )
      else:
        raise UnavailableSap('Sem acesso a transacao no sistema SAP!')
      print(relatorio.shape[0])
    elif aplicacao == 'coordenada':
      if 'ES61' in NOTUSE:
        instalacao_info = robo.ES32(argumento, ES32_FLAGS.ENTER_CONSUMO)
      else:
        instalacao_info = robo.ES32(argumento, ES32_FLAGS.ONLY_INST)
      if 'ES61' in NOTUSE:
        print(robo.ES61(instalacao_info, [ES61_FLAGS.SKIPT_ENTER, ES61_FLAGS.GET_COORD]).coordenadas)
      else:
        print(robo.ES61(instalacao_info, [ES61_FLAGS.ENTER_ENTER, ES61_FLAGS.GET_COORD]).coordenadas)
    else:
      raise ArgumentException('Nao entendi o comando, verifique se esto correto!')
    robo.HOME_PAGE()
  except ArgumentException as erro:
    print(f'400: {erro.message}')
  except InformationNotFound as erro:
    print(f'404: {erro.message}')
  except TooMannyRequests as erro:
    print(f'409: {erro.message}')
  except UnavailableSap as erro:
    print(f'500: {erro.message}')
  except SomethingGoesWrong as erro:
    print(f'500: {erro.message}')
  except WrapperBaseException as erro:
    print(f'500: {erro.message}')
  except Exception as erro:
    print(f'500: Something goes wrong! {erro.args[0]}')
