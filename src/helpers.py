#!/usr/bin/env python
''' Module to hold helper functions '''
# coding: utf8
#region imports
import os
import sqlite3
from constants import BASE_FOLDER
#endregion
def arquivo_configuracao(filepath: str, separador: str = '=') -> dict[str, str]:
  ''' Function to retrieve configurations from file '''
  if not os.path.exists(filepath):
    raise FileNotFoundError('O arquivo {filepath} nao foi encontrado')
  dicionario = {}
  with open(filepath, 'r', encoding='utf-8') as file:
    for line in file:
      argumentos = line.split(separador)
      if len(argumentos) != 2:
        continue
      dicionario[argumentos[0]] = argumentos[1]
  return dicionario
def depara(tipo: str, de: str) -> str | None:
  ''' Function wraper to get value from database '''
  try:
    filename = os.path.join(BASE_FOLDER, 'sap.db')
    connection = sqlite3.connect(filename)
    sql_instruction = f'SELECT para FROM depara WHERE tipo = \'{tipo}\' AND de = \'{de}\''
    cursor = connection.execute(sql_instruction)
    result = cursor.fetchone()
    return result[0] if result else 'Codigo desconhecido!'
  except ValueError:
    return None
