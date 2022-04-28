#!/usr/bin/python
# coding: utf8
import re

from wpp import wpp
from sap import sap
from xls import xls

class index:
  def __init__(self):
    self.whats = wpp()
    self.sape = sap()
    self.xlsx = xls()
    self.ultimo_texto = ''
    while True:
      self.resposta = ''
      self.texto = self.whats.escuta()
      if self.texto != self.ultimo_texto:
        self.ultimo_texto = self.texto
        argumentos = self.texto.split(' ')
        # if re.search("débito", self.texto):
        # if re.search("^\:*", argumentos[0]):
        # print(argumentos[0])
        if re.search("débito", self.texto):
          self.resposta = self.sape.debito(argumentos[1])
          self.resposta = self.xlsx.escrever(self.resposta)
          self.whats.responde(self.resposta)
        if argumentos[0] == ":r":
          self.resposta = self.sape.relatorio(argumentos[1])
          self.whats.responde(self.resposta)
        if argumentos[0] == ":l":
          self.resposta = self.sape.leiturista(argumentos[1])
          self.whats.responde(self.resposta)
        if argumentos[0] == ":h":
          self.resposta = self.sape.historico(argumentos[1])
          self.whats.responde(self.resposta)
robo = index()