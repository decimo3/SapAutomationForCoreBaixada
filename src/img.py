#!/usr/bin/python
# coding: utf8

from os import path
from sys import argv
from wand.image import Image 
from wand.drawing import Drawing 
from wand.color import Color

SEPARADOR_ENTRE_COLUNAS = "|"
SEPARADOR_ENTRE_LINHAS = "\n"
SEPARADOR_METADADOS = ":"
MARGEM_ESQUERDA = 13
LARGURA_CARACTERE = 14
ALTURA_CARACTERE = 20
nRow = 0 # contador de linha atual
nCol = 0 # contador de coluna atual
cursor = MARGEM_ESQUERDA # distância a esqueda da escrita do texto
CORES = ['rgb(255,255,255)', 'rgb(0,0,0)', 'rgb(255,0,0)', 'rgb(255,255,0)', 'rgb(0,255,0)']
# OFFSET_LINHAS_RELATORIO = []

linhas = argv[1].split(SEPARADOR_ENTRE_LINHAS)
metadados = []
if(linhas[0] == "0"):
  metadados.append(0)
else:
  metadados.append(int(linhas[0]))
  #metadados = linhas[0].split(SEPARADOR_METADADOS)
  #metadados = [int(x) for x in metadados]
# for metadado in metadados:
#   OFFSET_LINHAS_RELATORIO[0] = metadado.split(SEPARADOR_METADADOS)[0]
#   CORES_LINHAS_RELATORIO[0] = metadado.split(SEPARADOR_METADADOS)[1]
TAMANHO_COLUNAS_RELATORIO = linhas[1].split(SEPARADOR_ENTRE_COLUNAS)
TAMANHO_COLUNAS_RELATORIO = [int(x) for x in TAMANHO_COLUNAS_RELATORIO]
CARACTERES_TOTAL = sum(TAMANHO_COLUNAS_RELATORIO)

LARGURA_TOTAL_IMAGEM = CARACTERES_TOTAL * LARGURA_CARACTERE
ALTURA_TOTAL_IMAGEM = len(linhas) * ALTURA_CARACTERE

nRow = 2

with Drawing() as draw:
  with Image(width = LARGURA_TOTAL_IMAGEM, height = ALTURA_TOTAL_IMAGEM, background = Color(CORES[0])) as img:
    draw.font_family = 'Consolas'
    draw.font = 'Consolas'
    draw.font_size = ALTURA_CARACTERE # 15x15 cada letra
    while(nRow < len(linhas)):
      if(metadados[0] != 0 and (nRow == metadados[0] + 1)):
        draw.fill_color = Color(CORES[3])
        draw.rectangle(left = 0, top = (nRow * ALTURA_CARACTERE) + 1, right = LARGURA_TOTAL_IMAGEM, bottom = (nRow * ALTURA_CARACTERE) + ALTURA_CARACTERE + 1)
        draw.fill_color = Color(CORES[1])
      if(linhas[nRow] == ""):
        nRow = nRow + 1
        continue
      colunas = linhas[nRow].split(SEPARADOR_ENTRE_COLUNAS)
      while(nCol < len(colunas)):
        col = " " if (colunas[nCol] == None or colunas[nCol] == "") else colunas[nCol]
        draw.text(x = cursor, y = (nRow * ALTURA_CARACTERE), body = col)
        cursor = cursor + (TAMANHO_COLUNAS_RELATORIO[nCol] * LARGURA_CARACTERE)
        nCol = nCol + 1
      nCol = 0
      cursor = MARGEM_ESQUERDA
      nRow = nRow + 1
    draw(img)
    img.save(filename = "C:\\Users\\ruan.camello\\Documents\\Temporario\\temporario.png")
