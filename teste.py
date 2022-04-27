import re
argumento = 'Boa tarde 1204293380 roteiro prfv'
print(argumento)
print(re.search("[0-9]{10}", argumento))
resposta = re.search("[0-9]{10}", argumento)
print(resposta.match)