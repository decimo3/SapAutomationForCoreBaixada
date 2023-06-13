#!/usr/bin/python
# coding: utf8
import os
import sys
from os import path
from wand.image import Image 
from wand.drawing import Drawing 
from wand.color import Color

SEPARADOR_ENTRE_COLUNAS = "|"
SEPARADOR_ENTRE_LINHAS = "\n"
MARGEM_ESQUERDA = 13
LARGURA_CARACTERE = 14
ALTURA_CARACTERE = 20
nRow = 0 # contador de linha atual
nCol = 0 # contador de coluna atual
cursor = MARGEM_ESQUERDA # distância a esqueda da escrita do texto
# CORES: branco, preto, vermelho, amarelo, verde
CORES = ['rgb(255,255,255)', 'rgb(0,0,0)', 'rgb(255,128,128)', 'rgb(255,255,128)', 'rgb(128,255,128)']
# OFFSET_LINHAS_RELATORIO = []

valores = sys.stdin.read() if (len(sys.argv) < 2) else sys.argv[1]
linhas = valores.split(SEPARADOR_ENTRE_LINHAS)
TAMANHO_COLUNAS_RELATORIO = linhas[0].split(SEPARADOR_ENTRE_COLUNAS)
TAMANHO_COLUNAS_RELATORIO = [int(x) for x in TAMANHO_COLUNAS_RELATORIO]
CARACTERES_TOTAL = sum(TAMANHO_COLUNAS_RELATORIO)

LARGURA_TOTAL_IMAGEM = CARACTERES_TOTAL * LARGURA_CARACTERE
ALTURA_TOTAL_IMAGEM = len(linhas) * ALTURA_CARACTERE

nRow = 1

with Drawing() as draw:
  with Image(width = LARGURA_TOTAL_IMAGEM, height = ALTURA_TOTAL_IMAGEM, background = Color(CORES[0])) as img:
    draw.font_family = 'Consolas'
    draw.font = 'Consolas'
    draw.font_size = ALTURA_CARACTERE # 15x15 cada letra
    while(nRow < len(linhas)):
      if(linhas[nRow] == ""):
        nRow = nRow + 1
        continue
      colunas = linhas[nRow].split(SEPARADOR_ENTRE_COLUNAS)
      while(nCol < len(colunas)):
        if(nCol == 0):
          try:
            cor = int(colunas[nCol])
            if(cor > 0):
              draw.fill_color = Color(CORES[int(colunas[nCol])])
              draw.rectangle(left = 0, top = ((nRow - 1) * ALTURA_CARACTERE) + 1, right = LARGURA_TOTAL_IMAGEM, bottom = ((nRow - 1) * ALTURA_CARACTERE) + ALTURA_CARACTERE + 1)
              draw.fill_color = Color(CORES[1])
          except:
            pass
          finally:
            nCol = nCol + 1
            continue
        col = " " if (colunas[nCol] == None or colunas[nCol] == "") else colunas[nCol]
        draw.text(x = cursor, y = (nRow * ALTURA_CARACTERE), body = col)
        cursor = cursor + (TAMANHO_COLUNAS_RELATORIO[nCol] * LARGURA_CARACTERE)
        nCol = nCol + 1
      nCol = 0
      cursor = MARGEM_ESQUERDA
      nRow = nRow + 1
    draw(img)
    savefilename = os.getcwd() + "\\tmp\\temporario.png"
    img.save(filename = savefilename)
