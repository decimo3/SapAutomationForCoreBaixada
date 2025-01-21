''' Module to hold classes used on SapWrapper module '''
import re
import datetime
import pandas

class MedidorInfo():
  ''' class to hold information about meter '''
  instalacao: int
  serial: int
  material: int
  texto_material: str
  code_montagem: str
  code_status: str
  texto_montagem: str
  texto_status: str
  observacao: str
  leituras: pandas.DataFrame
  def __str__(self) -> str:
    texto = f'*Instalacao:* {self.instalacao}'
    texto += f'*Serial:* {self.serial}\n'
    texto += f'*Material:* {self.texto_material}\n'
    texto += f'*Status:* {self.texto_status}\n'
    texto += f'*Montagem:* {self.texto_montagem}\n'
    texto += f'*Observacao:* {self.observacao}' if self.observacao else ''
    return texto

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
  equipamento: list[MedidorInfo]
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
  coordenadas: str

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
