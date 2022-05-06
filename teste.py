import re
import win32com.client

class testes:
  def buscaRegex(self):
    argumento = 'Boa tarde 1204293380 roteiro prfv'
    print(argumento)
    print(re.search("[0-9]{10}", argumento))
    resposta = re.search("[0-9]{10}", argumento)
    print(resposta.match)
  def criarSessão(self):
    self.SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    self.connection = self.SapGui.OpenConnection("teste")
init = testes()
init.criarSessão()
