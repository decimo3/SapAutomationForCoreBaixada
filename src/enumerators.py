''' Hodule that holds enumerates '''
from enum import IntEnum

class DESTAQUES(IntEnum):
  ''' Class to hold enumerate with colors to csv highlights '''
  AUSENTE = 0
  AMARELO = 1
  VERMELHO = 2
  VERDE = 3

class ES32_FLAGS(IntEnum):
  ONLY_INST = 0
  GET_CENTER = 1
  GET_METER = 2
  ENTER_CONSUMO = 3

class IW53_FLAGS(IntEnum):
  GET_INST = 0
  GET_INFO = 1

class ZMED89_FLAGS(IntEnum):
  TIME_ORDER = 0
  SEQ_ORDER = 1
  TELEMEDIDO = 2

class ZARC140_FLAGS(IntEnum):
  GET_PENDING = 0
  GET_RENOTICE = 1

class ES61_FLAGS(IntEnum):
  ENTER_ENTER = 0
  SKIPT_ENTER = 1
  ENTER_LIGACAO = 2
  GET_COORD = 4

class ES57_FLAGS(IntEnum):
  ENTER_ENTER = 0
  SKIPT_ENTER = 1
  ENTER_LOGRADOURO = 2
  SKIPT_ENTER_LOGRADOURO_ENTER = 3

class ZMED95_FLAGS(IntEnum):
  ENTER_ENTER = 0
  SKIPT_ENTER = 1
  CHECK_PASIVE = 2
  CHECK_VALUES = 3

class BP_FLAGS(IntEnum):
  GET_PHONES = 0
  GET_DOCS = 1

class IQ03_FLAGS(IntEnum):
  ONLY_INST = 0
  READ_REPORT = 1