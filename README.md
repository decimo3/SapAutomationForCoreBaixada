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
NOTUSE=transacoes_sem_acesso
```
5. Para atualizar o banco de dados, execute:
```bash
rm src\\sap.db
sqlite3 src\\sap.db < src\\sap.sql
```
6. Para atualizar o executável, execute:
```bash
pyinstaller --name sap --icon appicon.ico --onefile src\\main.py
```
7. Copie todos os arquivos para a pasta `dist`:
```bash
md tmp
copy dist\\sap.exe tmp
copy src\\sap.conf tmp
copy src\\sap.path tmp
copy src\\sap.db tmp
copy src\\erroDialog.vbs tmp
copy src\\fileDialog.vbs tmp
``` 
8. Monte o pacote com o comando abaixo:
```bash
7z u sap_bot.zip .\\tmp\\*
```
9. Para finalizar, execute:
```bash
Deactivate
```

flask --app webserver run --debug


Obs.0: Esse script contém dependências específicas para o sistema operacional `Windowns`.

Obs.1: É necessário ter o `SAP FrontEnd` instalado e estar autenticado para usar o _runner_.

Obs.2: Esse _runner_ foi construído para o uso exclusivo do `CORE BAIXADA` e não se aplica para outros usos.

