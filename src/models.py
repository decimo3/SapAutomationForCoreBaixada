''' Module to hold classes used on SapWrapper module '''
from operator import eq
import re
import datetime
import pandas
from dataclasses import dataclass, field

from exceptions import InformationNotFound

@dataclass
class MedidorInfo():
  ''' class to hold information about meter '''
  instalacao: int = 0
  serial: int = 0
  material: int = 0
  texto_material: str = ''
  code_montagem: str = ''
  code_status: str = ''
  texto_montagem: str = ''
  texto_status: str = ''
  observacao: str = ''
  leituras: pandas.DataFrame = field(default_factory=pandas.DataFrame)
  def __str__(self) -> str:
    texto = f'*Instalacao:* {self.instalacao}\n' if self.instalacao else ''
    texto += f'*Serial:* {self.serial}\n' if self.serial else ''
    texto += f'*Material:* {self.texto_material}\n' if self.texto_material else ''
    texto += f'*Status:* {self.texto_status}\n' if self.texto_status else ''
    texto += f'*Montagem:* {self.texto_montagem}\n' if self.texto_montagem else ''
    texto += f'*Observacao:* {self.observacao}\n' if self.observacao else ''
    texto += self.leituras.to_string(index=False) if not self.leituras.empty else ''
    return texto

@dataclass
class InstalacaoInfo():
  ''' Class to hold information about instalation '''
  instalacao: int = 0
  status: str = ''
  parceiro: int = 0
  nome_cliente: str = ''
  contrato: int = 0
  endereco: str = ''
  consumo: int = 0
  classe: int = 0
  texto_classe: str = ''
  unidade: str = ''
  vigencia_inicio: datetime.datetime = field(default_factory=lambda: datetime.datetime.min)
  vigencia_final: datetime.datetime = field(default_factory=lambda: datetime.datetime.min)
  equipamento: list[MedidorInfo] = field(default_factory=list)
  centro: int = 0
  def __str__(self) -> str:
    texto = f'*Instalacao:* {self.instalacao}\n' if self.instalacao else ''
    texto += f'*Status:* {self.status}\n' if self.status else ''
    texto += f'*Contrato:* {self.contrato}\n' if self.contrato else ''
    texto += f'*Parceiro:* {self.parceiro}\n' if self.parceiro else ''
    texto += f'*Nome:* {self.nome_cliente}\n' if self.nome_cliente else ''
    texto += f'*Endereco:* {self.endereco}\n' if self.endereco else ''
    texto += f'*Classe:* {self.texto_classe}\n' if self.texto_classe else ''
    if self.equipamento:
      for i, medidor in enumerate(self.equipamento):
        texto += f'*Serial:* {medidor.serial}\n'
        texto += f'*Material:* {medidor.texto_material}\n'
    else:
      texto += f'Instalacao nao possui equipamentos vinculados!'
    return texto
  def get_medidor(self) -> MedidorInfo:
    counts = {}
    for equipamento in self.equipamento:
      counts[equipamento.material] = counts.get(equipamento.material, 0) + 1  # Increment count
    unique_material = [item for item, count in counts.items() if count == 1]
    if(len(unique_material) > 1):
      raise InformationNotFound('Nao foi possivel encontrar o medidor!')
    unique_material = unique_material[0]
    return [eq for eq in self.equipamento if eq.material == unique_material][0]


@dataclass
class ServicoInfo():
  ''' Class to hold information about service '''
  nota: int = 0
  tipo: str = ''
  dano: str = ''
  texto_dano: str = ''
  parceiro: int = 0
  nome_cliente: str = ''
  telefones: list[str] = field(default_factory=list)
  instalacao: int = 0
  endereco: str = ''
  local: str = ''
  cod_postal: str = ''
  texto_local: str = ''
  status: str = ''
  descricao: str = ''
  atendimento_obs: str = ''
  data_nota: datetime.date = field(default_factory=lambda: datetime.date.min)
  hora_nota: datetime.time = field(default_factory=lambda: datetime.time.min)
  avaria_inicio_data: datetime.date = field(default_factory=lambda: datetime.date.min)
  avaria_inicio_hora: datetime.time = field(default_factory=lambda: datetime.time.min)
  desejado_inicio_data: datetime.date = field(default_factory=lambda: datetime.date.min)
  desejado_inicio_hora: datetime.time = field(default_factory=lambda: datetime.time.min)
  avaria_final_data: datetime.date = field(default_factory=lambda: datetime.date.min)
  avaria_final_hora: datetime.time = field(default_factory=lambda: datetime.time.min)
  desejado_final_data: datetime.date = field(default_factory=lambda: datetime.date.min)
  desejado_final_hora: datetime.time = field(default_factory=lambda: datetime.time.min)
  encerramento_data: datetime.date = field(default_factory=lambda: datetime.date.min)
  encerramento_hora: datetime.time = field(default_factory=lambda: datetime.time.min)
  finalizacao: pandas.DataFrame = field(default_factory=pandas.DataFrame)
  equipamentos_inst: pandas.DataFrame = field(default_factory=pandas.DataFrame)
  equipamentos_remo: pandas.DataFrame = field(default_factory=pandas.DataFrame)
  observacao: str = ''
  def __str__(self) -> str:
    texto = f'*Nota:* {self.nota}\n' if self.nota else ''
    texto += f'*Tipo:* {self.tipo} ' if self.tipo else ''
    texto += f'*Dano:* {self.dano}\n' if self.dano else '\n'
    texto += f'*Texto dano:* {self.texto_dano}\n' if self.texto_dano else ''
    texto += f'*Status:* {self.status}\n' if self.status else ''
    texto += f'*Descricao:* {self.descricao}\n' if self.descricao else ''
    texto += f'*Parceiro:* {self.parceiro}\n' if self.parceiro else ''
    texto += f'*Nome:* {self.nome_cliente}\n' if self.nome_cliente else ''
    texto += f'*Telefones:* {' '.join(self.telefones)}\n' if len(self.telefones) > 0 else '\n'
    texto += f'*Instalacao:* {self.instalacao}\n' if self.instalacao else ''
    texto += f'*Endereco:* {self.endereco}\n' if self.endereco else ''
    texto += f'*Local:* {self.local} ' if self.local else ''
    texto += f'*CEP:* {self.cod_postal}\n' if self.cod_postal else '\n'
    texto += f'*Localidade:* {self.texto_local}\n' if self.texto_local else ''
    texto += f'*Observacao do atendente:* {self.atendimento_obs}\n' if self.atendimento_obs else ''
    texto += f'*Criação:* {self.data_nota} ' if self.data_nota > datetime.date.min else ''
    texto += f'{self.hora_nota}\n' if self.hora_nota > datetime.time.min else '\n'
    texto += f'*Inicio avaria:* {self.avaria_inicio_data} ' if self.avaria_inicio_data > datetime.date.min else ''
    texto += f'{self.avaria_inicio_hora}\n' if self.avaria_inicio_hora > datetime.time.min else '\n'
    texto += f'*Inicio desejado:* {self.desejado_inicio_data} ' if self.desejado_inicio_data > datetime.date.min else ''
    texto += f'{self.desejado_inicio_hora}\n' if self.desejado_inicio_hora > datetime.time.min else '\n'
    texto += f'*Final avaria:* {self.avaria_final_data} ' if self.avaria_final_data > datetime.date.min else ''
    texto += f'{self.avaria_final_hora}\n' if self.avaria_final_hora > datetime.time.min else '\n'
    texto += f'*Final desejado:* {self.desejado_final_data} ' if self.desejado_final_data > datetime.date.min else ''
    texto += f'{self.desejado_final_hora}\n' if self.desejado_final_hora > datetime.time.min else '\n'
    texto += f'*Encerramento:* {self.encerramento_data} ' if self.encerramento_data > datetime.date.min else ''
    texto += f'{self.encerramento_hora}\n' if self.encerramento_hora > datetime.time.min else '\n'
    texto += f'*Finalizacao:*\n{self.finalizacao.to_string(index=False)}\n' if self.finalizacao.shape[0] > 0 else ''
    texto += f'*Equipamentos instalados:*\n{self.equipamentos_inst.to_string(index=False)}\n' if self.equipamentos_inst.shape[0] > 0 else ''
    texto += f'*Equipamentos removidos:*\n{self.equipamentos_remo.to_string(index=False)}\n' if self.equipamentos_remo.shape[0] > 0 else ''
    texto += f'*Observacao:* {self.observacao}\n' if self.observacao else ''
    return texto

@dataclass
class ParceiroInfo():
  ''' class to hold information about client '''
  parceiro: int = 0
  nome_cliente: str = ''
  documento_tipo: str = ''
  documento_numero: str = ''
  telefones: list[str] = field(default_factory=list)
  def __str__(self) -> str:
    texto = f'*Parceiro:* {self.parceiro}\n' if self.parceiro else ''
    texto += f'*Nome:* {self.nome_cliente}\n' if self.nome_cliente else ''
    texto += f'*Tipo do documento:* {self.documento_tipo}\n' if self.documento_tipo else ''
    texto += f'*Numero do documento:* {self.documento_numero}\n' if self.documento_numero else ''
    texto += f'*Telefones:* {' '.join(self.telefones)}' if len(self.telefones) > 0 else ''
    return  texto

class LigacaoInfo():
  ''' Class to hold information about linker object '''
  ligacao: int
  tipo_instalacao: str
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

