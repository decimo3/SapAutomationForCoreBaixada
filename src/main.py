#!/usr/bin/python
# coding: utf8

from sap import sap
import sys
import os
import datetime

class main:
  def __init__(self):
#   if (datetime.datetime.now() > datetime.datetime.strptime("13.01.2023", "%d.%m.%Y")):
#     raise Exception("O prazo de avaliação acabou!")
    #sys.argv
    if 'sys.argv[1]' in locals():
      self.instancia = int(sys.argv[1])
    else:
      self.instancia = 0
    print("Programa de automação de rotinas do Mestre Ruan")
    print("===============================================")
    print("")
    self.sape = sap(self.instancia)
    self.ultima_nota = ""
    print("Digite o comando abaixo que deseja executar ou AJUDA para mostrar as opções disponíveis.")
    print("")
    while True:
      resposta = input("> ")
      resposta = resposta.lower()
      argumentos = resposta.split(' ')
      if (len(argumentos) < 2):
        argumentos.append(self.ultima_nota)
      if (argumentos[0] == "sair"):
        break
      elif (argumentos[0] == ""):
        continue
      elif (argumentos[0] == "ajuda"):
        print("")
        print("RELATORIO dias\n\tAutomatiza o relatório de religa retroativo a quantidade de dias informado.")
        print("LEITURISTA nota\n\tAutomatiza o relatório de leitura do mês anterior para a nota informada.")
        print("DEBITO nota\n\tAutomatiza o relatório de débitos da instalação referente a nota informada.")
        print("HISTORICO nota\n\tAutomatiza o relatório com o histórico de notas referentes a nota informada.")
        print("AGRUPAMENTO nota\n\tAutomatiza o relatório de consulta de débitos para o agrupamento da nota informada.")
        print("MANOBRA dias\n\tAutomatiza o relatório de CORE MT retroativo a quantidade de dias informado.")
        print("COORDENADA nota\n\tAutomatiza a coleta das coordenadas da instalação e monta o link do Google Maps")
        print("TELEFONE nota\n\tAutomatiza a coleta de telefones do cliente, tanto do atendimento, quanto do cadastro")
        print("MEDIDOR nota\n\tAutomatiza a verificação de código de retirada de medidor na nota")
        print("FATURA nota\n\tAutomatiza a impressão de 2a via de fatura através da nota informada")
      elif (argumentos[0] == "relatorio"):
        self.sape.relatorio(int(argumentos[1]))
      elif (argumentos[0] == "leiturista"):
        self.sape.leiturista(int(argumentos[1]))
      elif (argumentos[0] == "debito"):
        self.sape.debito(int(argumentos[1]))
      elif (argumentos[0] == "historico"):
        self.sape.historico(int(argumentos[1]))
      elif (argumentos[0] == "agrupamento"):
        self.sape.agrupamento(int(argumentos[1]))
      elif (argumentos[0] == "manobra"):
        self.sape.manobra(int(argumentos[1]))
      elif (argumentos[0] == "coordenada"):
        self.sape.coordenadas(int(argumentos[1]))
      elif (argumentos[0] == "telefone"):
        self.sape.telefone(int(argumentos[1]))
      elif (argumentos[0] == "medidor"):
        self.sape.medidor(int(argumentos[1]))
      elif (argumentos[0] == "fatura"):
        self.sape.fatura_novo(int(argumentos[1]))
      elif (argumentos[0] == "analise"):
        self.sape.analisar(int(argumentos[1]))
      else: print("Selecione uma opção válida. Digite AJUDA para saber as consultas suportadas ou SAIR para terminar o programa!")
      self.ultima_nota = argumentos[1]
robo = main()