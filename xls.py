#!/usr/bin/python
# coding: utf8

import win32com.client

class xls:
  def __init__(self):
    self.excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    self.excel.Visible = True
    self.book = self.excel.Workbooks.Add()
    self.sheet = self.book.Worksheets('Planilha1')
    self.sheet.Select()
    # self.sheet.Name = 'DEBITOS'
    # cabecalho = list(['Referência', 'Vencimento', 'Valor', 'Tipo'])
    # self.sheet.Range(self.sheet.Cells(1,1),self.sheet.Cells(1,4)).Value = cabecalho
    # self.book.Worksheets.Add()
    # self.sheet = self.book.Worksheets('Planilha2')
    # self.sheet.Name = 'LEITURISTA'
    # cabecalho = list(["Seq", "Instalação", "Endereço", "Bairro", "Medidor", "Hora", "Data", "Valor", "Leiturista", "Cod"])
    # self.sheet.Range(self.sheet.Cells(1,1),self.sheet.Cells(1,10)).Value = cabecalho
    # # print(dir(self.sheet))
  def escrever(self, arg):
    # TODO: algoritmo para apagar o registro anterior (apesar de ser irrelevante)
    linhas = self.sheet.Rows.Count
    self.sheet.Rows(linhas).Delete()
    apontador = 0
    linhas = arg.split('\n')
    while (apontador < (len(linhas) - 1)):
      colunas = linhas[apontador].split('\t')
      contador = 1
      while (contador <= len(colunas)):
        self.sheet.Cells(apontador + 1, contador).Value = colunas[contador - 1]
        contador = contador + 1
      apontador = apontador + 1
    self.sheet.Columns.AutoFit()
    self.sheet.Range(self.sheet.Cells(1,1),self.sheet.Cells(apontador,len(colunas))).Copy()
    print(f"Linhas: {apontador}, Colunas: {len(colunas)}")
# a = xls()
# a.escrever("Referência\tVencimento\tValor\tTipo\n2022/04\t09.05.2022\tR$:348,65\tFat. Normal\n2022/03\t08.04.2022\tR$:438,54\tFat. Normal\n2022/02\t08.03.2022\tR$:569,27\tFat. Normal\n2022/01\t08.02.2022\tR$:750,78\tFat. Normal\n2021/12\t10.01.2022\tR$:1.178,74\tFat. Normal\n")
# a.escrever("Referência\tVencimento\tValor\tTipo\n2022/04\t25.04.2022\tR$:143,11\tFat. Normal\n2022/03\t25.03.2022\tR$:144,69\tFat. Normal\n")