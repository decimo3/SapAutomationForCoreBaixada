#!/usr/bin/python
# coding: utf8

from os import path
from sys import argv
from wand.image import Image 
from wand.drawing import Drawing 
from wand.color import Color

SEPARADOR_ENTRE_COLUNAS = "|"
TAMANHO_COLUNAS_RELATORIO = [(5 * 13),0,0,(10 * 13),(10 * 13),0]

MARGEM_ESQUERDA = 13
nRow = 0 # contador de linha atual
nCol = 0 # contador de coluna atual
nEnd = 0 # tamanho da string do endereço
nBairro = 0 # tamanho da string do bairro
cursor = MARGEM_ESQUERDA # distância a esqueda da escrita do texto
linhas = argv[1].split("\n")

LARGURA_TOTAL = 0
while(nRow < len(linhas)):
  colunas = linhas[nRow].split(SEPARADOR_ENTRE_COLUNAS)
  # Verifica o maior tamanho do endereço
  if(len(colunas[1]) > nEnd):
    nEnd = len(colunas[1])
  # Verifica o maior tamanho do sub-bairro
  if(len(colunas[2]) > nBairro):
    nBairro = len(colunas[2])
  # Verifica o tamanho total da linha
  if(len(linhas[nRow]) > LARGURA_TOTAL):
    LARGURA_TOTAL = len(linhas[nRow])
  nRow = nRow + 1

TAMANHO_COLUNAS_RELATORIO[1] = nEnd * 13
TAMANHO_COLUNAS_RELATORIO[2] = nBairro * 13

nRow = 0
nCol = 0

print(f"A maior linha tem {LARGURA_TOTAL} caracteres!")
print(f"O maior endereço tem {nEnd} caracteres!")


with Drawing() as draw:
  with Image(width = (LARGURA_TOTAL * 13), height = (len(linhas) * 20 + 5), background = Color('white')) as img:
    draw.font_family = 'Arial'
    draw.font = 'monospace'
    draw.font_size = 20 # 15x15 cada letra
    while(nRow < len(linhas)):
      colunas = linhas[nRow].split(SEPARADOR_ENTRE_COLUNAS)
      while(nCol < len(colunas)):
        col = "0" if (colunas[nCol] == None or colunas[nCol] == "") else colunas[nCol]
        draw.text(x = cursor, y = (nRow + 1) * 20, body = col)
        cursor = cursor + TAMANHO_COLUNAS_RELATORIO[nCol]
        nCol = nCol + 1
      nCol = 0
      cursor = MARGEM_ESQUERDA
      nRow = nRow + 1
    draw(img) 
    img.save(filename = "temporary.png")
