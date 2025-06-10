''' Web server to handle GET requests with application name and information ID.'''
from flask import Flask, request
from main import aplicacoes

app = Flask(__name__)

@app.route("/", methods=["GET"])
def index_get():
  ''' Endpoint to handle GET requests with application name and information ID. '''
  if not request.args.get('aplicacao') or not request.args.get('informacao'):
    return "Parâmetros 'aplicacao' e 'informacao' são obrigatórios.", 400
  # Attempts to connect to SAP FrontEnd on the specified instance
  aplicacao = request.args.get('aplicacao')
  argumento = int(request.args.get('informacao'))
  instancia = load_balancer()
  robo = SapBot(instancia)
  if aplicacao == 'instancia':
    robo.create_session(argumento)
  else:
    robo.attach_session(instancia)
  try:
    retorno = aplicacoes[aplicacao](robo, argumento)
    if isinstance(retorno, pandas.DataFrame):
      if '#' in retorno.columns:
        retorno['#'] = retorno['#'].astype(int)
        return retorno.to_csv(index=False,sep=SEPARADOR), 200
      else:
        return str(retorno), 200
  except:
    return f"Erro ao executar a aplicação '{aplicacao}' com o argumento '{argumento}'.", 500

if __name__ == "__main__":
  ''' Run the Flask web server. '''
  app.run(host="0.0.0.0", port=8080)
