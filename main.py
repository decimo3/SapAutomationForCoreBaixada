# coding: utf8

import sapHandler

class main:
  print("Programa de automação de rotinas do Mestre Ruan")
  print("===============================================")
  print("")
  sape = sapHandler.sap()
  print("Digite o comando abaixo que deseja executar ou AJUDA para mostrar as opções disponíveis.")
  print("")
  while True:
    resposta = input("> ")
    resposta = resposta.lower()
    argumentos = resposta.split(' ')
    if argumentos[0] == "sair":
      break
    if argumentos[0] == "ajuda":
      print("")
      print("RELATORIO dias")
      print("\tAutomatiza o relatório de religa retroativo a quantidade de dias informado")
      print("LEITURISTA nota")
      print("\tAutomatiza o relatório de leitura do mês anterior para a nota informada")
    if argumentos[0] == "relatorio":
      try:
        sape.relatorio(int(argumentos[1]))
      except:
        sape.relatorio()
    if argumentos[0] == "leiturista":
      try:
        sape.leiturista(int(argumentos[1]))
      except:
        print("É necessário fornecer um número de nota válido")