#!/usr/bin/python
# coding: utf8

from sap import sap

class main:
  def __init__(self):
    print("Programa de automação de rotinas do Mestre Ruan")
    print("===============================================")
    print("")
    self.sape = sap()
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
        self.sape.relatorio(int(argumentos[1]))
      elif (argumentos[0] == "leiturista"):
        self.sape.leiturista(int(argumentos[1]))
      elif (argumentos[0] == "debito"):
        self.sape.debito(int(argumentos[1]))
      elif (argumentos[0] == "historico"):
        self.sape.historico(int(argumentos[1]))
      else: print("Selecione uma opção válida. Digite AJUDA para saber as consultas suportadas ou SAIR para terminar o programa!")
robo = main()