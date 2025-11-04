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

5. Para atualizar o executável, execute:

    ```bash
    pyinstaller -n sap --icon appicon.ico --onefile src/main.py
    sqlite3 dist/sap.db < src/sap.sql
    cp 
    ```

6. Para finalizar, execute:

    ```bash
    Deactivate
    ```

Obs.1: Esse script contém dependências específicas para o sistema operacional `Windowns`.

Obs.2: É necessário ter o `SAP FrontEnd` instalado para usar o _runner_.
