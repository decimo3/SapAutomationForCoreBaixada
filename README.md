# runner de escritório do Mestre Ruan

1. Crie um ambiente virtual:
```bash
python3 -m venv venv
```
2. Ative o ambiente virtual:
```bash
venv\Scripts\Activate
```
3. Instale as dependências:
```bash
pip install -r requirements.txt
```
4. Configure `sap.conf` com sua autenticação
```txt
USUARIO=sua_matricula_aqui
PALAVRA=sua_senha_aqui
```
6. Para atualizar o executável, execute:
```bash
pyinstaller --onefile sap.py
```
5. Para finalizar, execute:
```bash
Deactivate
```

Obs.0: Esse script contém dependências específicas para o sistema operacional `Windowns`.

Obs.1: É necessário ter o `SAP FrontEnd` instalado e estar autenticado para usar o _runner_.

Obs.2: Esse _runner_ foi construído para o uso exclusivo do `CORE BAIXADA` e não se aplica para outros usos.

