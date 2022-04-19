#!/usr/bin/python
# coding: utf8

import os 
import time

from selenium import webdriver

class wpp:
  def __init__(self):
      self.directory = os.getenv('USERPROFILE')
      self.temporary = os.getcwd() + '\\.whatsapp'
      self.chrome = self.directory + '\\AppData\\Local\\SeleniumBasic\\chromedriver.exe'
      self.options = webdriver.ChromeOptions()
      self.options.add_argument(f"user-data-dir={self.temporary}")
      self.driver = webdriver.Chrome(self.chrome, chrome_options=self.options)
      self.inicia('Teste robo')
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