#!/usr/bin/python
# coding: utf8

import os 
import time

from sap import sap
from selenium import webdriver

class wpp:
  def __init__(self):
      self.sape = sap()
    # try:
      self.directory = os.getenv('USERPROFILE')
      self.temporary = os.getcwd() + '\\.whatsapp'
      self.chrome = self.directory + '\\AppData\\Local\\SeleniumBasic\\chromedriver.exe'
      self.options = webdriver.ChromeOptions()
      self.options.add_argument(f"user-data-dir={self.temporary}")
      self.driver = webdriver.Chrome(self.chrome, chrome_options=self.options)
      self.inicia('Teste robo')
      self.ultimo_texto = ''
      while True:
        self.texto = self.escuta()
        if self.texto != self.ultimo_texto:
          self.ultimo_texto = self.texto
          argumentos = self.texto.split(' ')
          print(argumentos[0])
          if argumentos[0] == ":i":
            if (argumentos[1] == "sair"):
              break
            elif (argumentos[1] == "ajuda"):
              print("")
              print("RELATORIO dias")
              print("\tAutomatiza o relatório de religa retroativo a quantidade de dias informado.")
              print("LEITURISTA nota")
              print("\tAutomatiza o relatório de leitura do mês anterior para a nota informada.")
              print("DEBITO nota")
              print("\tAutomatiza o relatório de débitos da instalação referente a nota informada.")
              print("HISTORICO nota")
              print("\tAutomatiza o relatório com o histórico de notas referentes a nota informada.")
            elif (argumentos[1] == "relatorio"):
              try:
                self.sape.relatorio(int(argumentos[2]))
              except:
                self.sape.relatorio()
            elif (argumentos[1] == "leiturista"):
              try:
                self.sape.leiturista(int(argumentos[2]))
              except:
                print("É necessário fornecer um número de nota válido")
            elif (argumentos[1] == "debito"):
              try:
                self.sape.debito(int(argumentos[2]))
              except:
                print("É necessário fornecer um número de nota válido")
            elif (argumentos[1] == "historico"):
              try:
                self.sape.historico(int(argumentos[2]))
              except:
                print("É necessário fornecer um número de nota válido")
            else: print("Selecione uma opção válida. Digite AJUDA para saber as consultas suportadas ou SAIR para terminar o programa!")
          print(self.texto)
          self.responde(f"{self.texto}")
    # except:
      print("O módulo do WhatsApp não pode ser iniciado!")
      raise Exception("O módulo do WhatsApp não pode ser iniciado!\nVerifique se o chromedriver está atualizado.")
  def inicia(self, nome_contato):
    self.driver.get(r'https://web.whatsapp.com/')
    self.driver.implicitly_wait(15)
    self.caixa_de_pesquisa = self.driver.find_element_by_class_name('_13NKt')
    self.caixa_de_pesquisa.send_keys(nome_contato)
    time.sleep(2)
    self.contato = self.driver.find_element_by_xpath('//span[@title = "{}"]'.format(nome_contato))
    self.contato.click()
    time.sleep(2)
    return True
  def escuta(self):
    post = self.driver.find_elements_by_class_name('i0jNr')
    ultimo = len(post) - 1
    post = post[ultimo].find_elements_by_tag_name("span")
    texto = post[0].text
    return texto
  def responde(self, texto):
    response = texto
    self.caixa_de_mensagem = self.driver.find_element_by_xpath('//div[@title = "{}"]'.format("Mensagem"))
    self.caixa_de_mensagem.send_keys(response)
    time.sleep(1)
    self.botao_enviar = self.driver.find_element_by_class_name('_4sWnG')
    self.botao_enviar.click()