#!/usr/bin/python
# coding: utf8

import win32com.client

class xls:
  def __init__(self):
    self.excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    self.excel.Visible = True
    self.wb = self.excel.Workbooks.Add()
    self.ws1 = self.wb.Worksheets('Planilha1')
    self.ws1.Name = 'DEBITOS'
    cabecalho = list(['Referência', 'Vencimento', 'Valor', 'Tipo'])
    self.ws1.Range(self.ws1.Cells(1,1),self.ws1.Cells(1,4)).Value = cabecalho
    self.wb.Worksheets.Add()
    self.ws2 = self.wb.Worksheets('Planilha2')
    self.ws2.Name = 'LEITURISTA'
    cabecalho = list(["Seq", "Instalação", "Endereço", "Bairro", "Medidor", "Hora", "Data", "Valor", "Leiturista", "Cod"])
    self.ws2.Range(self.ws2.Cells(1,1),self.ws2.Cells(1,10)).Value = cabecalho
    # print(dir(self.ws1))
  def debito(self, arg):
    self.ws1.Select()
    self.ws1.Range(self.ws1.Cells(2,1),self.ws1.Cells(self.ws1.Cells(2,1).End(-4121).Row,4)).Delete()
    apontador = 0
    contador = 0
    linhas = arg.split('\n')
    while (apontador < len(linhas)):
      colunas = linhas[apontador].split('\t')
      apontador = apontador + 1
      contador = 0
      while (contador < len(colunas)):
        contador = contador + 1
        self.ws1.Cells(apontador + 1,contador).Value = colunas[contador - 1]
    self.ws1.Columns.AutoFit()
    self.ws1.Range(self.ws1.Cells(1,1),self.ws1.Cells(len(linhas),4)).Copy()
  def leitura(self, arg):
    self.ws2.Select()
    conteudo = arg.split('|')
    a = 2
    for c in conteudo:
      if c == '#':
        a =+ 1
      self.ws2.Range(self.ws2.Cells(a,c),self.ws2.Cells(a,c)).Copy()
# a = xls()
# a.inicia()
# a.debito("2022/04\t09.05.2022\tR$:348,65\tFat. Normal\n2022/03\t08.04.2022\tR$:438,54\tFat. Normal\n2022/02\t08.03.2022\tR$:569,27\tFat. Normal\n2022/01\t08.02.2022\tR$:750,78\tFat. Normal\n2021/12\t10.01.2022\tR$:1.178,74\tFat. Normal\n")
# a.debito("2022/04\t25.04.2022\tR$:143,11\tFat. Normal\n2022/03\t25.03.2022\tR$:144,69\tFat. Normal\n")