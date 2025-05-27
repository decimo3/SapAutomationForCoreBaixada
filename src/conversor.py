''' Module to convert objects and strings to desire type '''
import datetime

__accents = {
  'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a', 'Á': 'A',
  'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e', 'É': 'E', 'Ê': 'E',
  'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i', 'Í': 'I',
  'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o', 'Õ': 'O', 'Ó': 'O',
  'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u', 'Ú': 'U',
  'ç': 'c', 'Ç': 'C'
}

def __texto(arg) -> str:
  ''' Function to get string from argument '''
  arg = str(arg).strip()
  return ''.join(__accents.get(char, char) for char in arg)
def __numero(arg) -> int:
  ''' Function to get decimal from argument '''
  arg = str(arg).strip()
  if not arg:
    return 0
  arg = arg.replace('.', '')
  arg = arg.replace(',', '.')
  return int(arg)
def __decimal(arg) -> float:
  ''' Function to get decimal from argument '''
  arg = str(arg).strip()
  if not arg:
    return 0
  arg = arg.replace('.', '')
  arg = arg.replace(',', '.')
  return float(arg)
def __data(arg) -> datetime.date:
  ''' Function to get dateonly from argument '''
  arg = str(arg).strip()
  if not arg:
    return datetime.date.min
  arg = arg.replace('.', '/')
  try:
    return datetime.datetime.strptime(arg, '%d/%m/%Y').date()
  except:
    return datetime.date.min
def __hora(arg) -> datetime.time:
  ''' Function to get timeonly from argument '''
  arg = str(arg).strip()
  if not arg:
    return datetime.time.min
  return datetime.datetime.strptime(arg, '%H:%M:%S').time()
def __datahora(arg) -> datetime.datetime:
  ''' Function to get datetime from argument '''
  arg = str(arg).strip()
  if not arg:
    return datetime.datetime.min
  return datetime.datetime.strptime(arg, '%d/%m/%Y %H:%M:%S')

conversor = {
  'texto': __texto,
  'numero': __numero,
  'decimal': __decimal,
  'data': __data,
  'hora': __hora,
  'datahora': __datahora
}