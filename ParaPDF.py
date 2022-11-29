#!/usr/bin/python
# coding: utf8

import win32print

# Lista até 5 impressoras instaladas
impressoras = win32print.EnumPrinters(2)
print(impressoras)
# Pega a impressora padrão do sistema
padrao = win32print.GetDefaultPrinter()
print(padrao)
# Abre um canal de comunicação com a impressora
printer_handle = win32print.OpenPrinter(padrao)
# Lista até 10 tarefas pendentes da impresora
print(win32print.EnumJobs(printer_handle, 0, 10))