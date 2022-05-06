#!/usr/bin/python
# coding: utf8

import datetime

class log:
  def escrever(self, app, code, texto):
    file = open('log.log', 'a')
    agora = datetime.datetime.now()
    file.write(f"{agora.strftime('%Y.%m.%d %H:%M:%S')} - {app}:{code} - {texto}\n")
    file.close()
    print(texto)
