#!/usr/bin/python
# coding: utf8

import datetime
from sqlalchemy import create_engine

class Log:
  _instance = None
  connection = create_engine('sqlite:///database.db').connect()
  def __init__(self) -> None:
    self.some_attribute = None
  def getInstance(self):
    if self._instance is None:
      self._instance = self()
    return self._instance
  def get(self) -> str:
    query = self.connection.execute("select * from user")
    result = [dict(zip(tuple(query.keys()), i)) for i in query.cursor]
    print(result)
  def set(self, registro):
    self.connection.execute(f"insert into user values(null, '{registro.timestamp}', '{registro.telefone}','{registro.modulo}','{registro.telefone}')")

class Registro:
  telefone = ""
  modulo = ""
  funcao = ""
  timestamp = ""
  def __init__(self, telefone, modulo, funcao) -> None:
    self.telefone = telefone
    self.modulo = modulo
    self.funcao = funcao
    self.timestamp = datetime.datetime.now().__str__




