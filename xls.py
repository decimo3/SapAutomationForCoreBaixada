#!/usr/bin/python
# coding: utf8

import sys
import win32com.client

class xls:
  def __init__(self):
    self.excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    self.excel.Visible = True
    self.book = self.excel.Workbooks.Add()
    self.sheet = self.book.Worksheets('Planilha1')
    self.sheet.Select()
  def escrever(self, arg: str):
    # TODO: algoritmo para apagar o registro anterior (apesar de ser irrelevante)
    linhas = self.sheet.Rows.Count
    self.sheet.Rows(linhas).Delete()
    apontador = 0
    linhas = arg.split('\n')
    while (apontador <= (len(linhas) - 1)):
      colunas = linhas[apontador].split('\t')
      contador = 1
      while (contador <= len(colunas)):
        self.sheet.Cells(apontador + 1, contador).Value = colunas[contador - 1]
        contador = contador + 1
      apontador = apontador + 1
    self.sheet.Columns.AutoFit()
    self.sheet.Range(self.sheet.Cells(1,1),self.sheet.Cells(apontador,len(colunas))).Copy()
    print(f"Linhas: {apontador}, Colunas: {len(colunas)}")
    # self.book.Close(SaveChanges=False)


if __name__ == "__main__":
  excel = xls()
  excel.escrever(sys.argv[1])