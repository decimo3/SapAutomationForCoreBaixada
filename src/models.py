''' Module to hold classes used on SapWrapper module '''
import re
import datetime

class InstalacaoInfo():
  ''' Class to hold information about instalation '''
  instalacao: int
  status: str
  parceiro: int
  nome_cliente: str
  contrato: int
  endereco: str
  consumo: int
  classe: int
  unidade: str
  vigencia_inicio: datetime.datetime
  vigencia_final: datetime.datetime
  equipamento: list[dict]
  centro: int

class ServicoInfo():
  ''' Class to hold information about service '''
  nota: int
  status: str
  descricao: str
  observacao: str
  data_criacao: datetime.datetime
  fim_avaria: datetime.datetime
  instalacao: int
  finalizacao: list[str]


class LigacaoInfo():
  ''' Class to hold information about linker object '''
  ligacao: int

class LogradouroInfo():
  ''' Class to hold information about street '''
  logradouro: int
  numero_str: str
  numero_int: int
  def __init__(
    self,
    logradouro: str,
    numero: str
    ) -> None:
    self.logradouro = int(logradouro)
    self.numero_str = numero
    match = re.search("[0-9]+", numero)
    if match is None:
      self.numero_int = 0
    else:
      self.numero_int = int(match.group())

class ParceiroInfo():
  ''' class to hold information about client '''
  parceiro: int
  nome_cliente: str
  documento_tipo: str
  documento_numero: str
  telefones: list[str]