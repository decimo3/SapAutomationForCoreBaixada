#!/usr/bin/python
# coding: utf8

from sap import sap
from wpp import wpp

class main:
  def __init__(self):
    print("Programa de automação de rotinas do Mestre Ruan")
    print("===============================================")
    print("")
    self.sape = sap()
    self.wapp = wpp()
    print("Digite o comando abaixo que deseja executar ou AJUDA para mostrar as opções disponíveis.")
    print("")
    while True:
      resposta = input("> ")
      resposta = resposta.lower()
      argumentos = resposta.split(' ')
      if (argumentos[0] == "sair"):
        break
      elif (argumentos[0] == "ajuda"):
        print("")
        print("RELATORIO dias")
        print("\tAutomatiza o relatório de religa retroativo a quantidade de dias informado.")
        print("LEITURISTA nota")
        print("\tAutomatiza o relatório de leitura do mês anterior para a nota informada.")
        print("DEBITO nota")
        print("\tAutomatiza o relatório de débitos da instalação referente a nota informada.")
        print("HISTORICO nota")
        print("\tAutomatiza o relatório com o histórico de notas referentes a nota informada.")
      elif (argumentos[0] == "relatorio"):
        try:
          self.sape.relatorio(int(argumentos[1]))
        except:
          self.sape.relatorio()
      elif (argumentos[0] == "leiturista"):
        try:
          self.sape.leiturista(int(argumentos[1]))
        except:
          print("É necessário fornecer um número de nota válido")
      elif (argumentos[0] == "debito"):
        try:
          self.sape.debito(int(argumentos[1]))
        except:
          print("É necessário fornecer um número de nota válido")
      elif (argumentos[0] == "historico"):
        try:
          self.sape.historico(int(argumentos[1]))
        except:
          print("É necessário fornecer um número de nota válido")
      else: print("Selecione uma opção válida. Digite AJUDA para saber as consultas suportadas ou SAIR para terminar o programa!")
robo = main()