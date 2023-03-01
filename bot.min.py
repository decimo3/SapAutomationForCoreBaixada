from os import environ
from dotenv import load_dotenv
from sap import sap
from telegram import Update, ForceReply, InlineKeyboardMarkup, InlineKeyboardButton, ParseMode
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, CallbackContext, CallbackQueryHandler, ContextTypes

class telbot:
  def __init__(self) -> None:
    # Load enviorement variables from '.env' file
    load_dotenv()
    # create a new instance of SAP runner
    self.sape = sap(1)
    self.updater = Updater(environ.get("TOKEN"))
    # Get the dispatcher to register handlers
    self.dispatcher = self.updater.dispatcher
    # Register commands each handler and the conditions the update must meet to trigger it
    self.dispatcher.add_handler(CommandHandler("start", self.start))
    self.dispatcher.add_handler(CommandHandler("ajuda", self.ajuda))
    # executa a funcao 'write' para toda mensagem que não é um comando
    self.dispatcher.add_handler(MessageHandler(~Filters.command, self.write))
    self.dispatcher.add_error_handler(self.error)
    # Start the Bot
    self.updater.start_polling()
    # Run the bot until you press Ctrl-C
    self.updater.idle()
  def write(self, update: Update, context: CallbackContext) -> None:
    # This function would be added to the dispatcher as a handler for messages coming from the Bot API
    print(f'{update.message.from_user.first_name} wrote {update.message.text}')
    if(not(update.message.text)):
      return
    resposta = update.message.text.lower()
    argumentos = resposta.split(' ')
    if (len(argumentos) < 2):
      self.ajuda(update, context)
      return
    # Métodos homologados
    if (argumentos[0] == "telefone"):
      resposta = self.sape.telefone(int(argumentos[1]))
      context.bot.send_message(update.message.chat_id, resposta)
    elif (argumentos[0] == "localização"):
      context.bot.send_message(update.message.chat_id, self.sape.coordenadas(int(argumentos[1])))
    elif (argumentos[0] == "coordenada"):
      context.bot.send_message(update.message.chat_id, self.sape.coordenadas(int(argumentos[1])))
    # métodos não homologados
    # elif (argumentos[0] == "leiturista"):
    #   self.sap.leiturista(int(argumentos[1]))
    # elif (argumentos[0] == "debito"):
    #   self.sap.debito(int(argumentos[1]))
    # elif (argumentos[0] == "historico"):
    #   self.sap.historico(int(argumentos[1]))
    # elif (argumentos[0] == "agrupamento"):
    #   self.sap.agrupamento(int(argumentos[1]))
    # elif (argumentos[0] == "manobra"):
    #   self.sap.manobra(int(argumentos[1]))
    # elif (argumentos[0] == "consulta"):
    #   self.sap.consulta(argumentos[1])
    # elif (argumentos[0] == "medidor"):
    #   self.sap.medidor(int(argumentos[1]))
    # elif (argumentos[0] == "fatura"):
    #   self.sap.fatura_novo(int(argumentos[1]))
    else:
      context.bot.send_message(update.message.chat_id, "Selecione uma opção válida.")
      context.bot.send_message(update.message.chat_id, "Digite /AJUDA para saber as consultas suportadas")
      self.ajuda(update, context)
  #
  def error(self, update: Update, context: CallbackContext) -> None:
    context.bot.send_message(update.message.chat_id, "Ocorreu um erro ao tentar processar a informação solicitada!")
    context.bot.send_message(update.message.chat_id, "Solicito que verifique com o monitor(a) responsável")
    return
  def start(self, update: Update, context: CallbackContext) -> None:
    context.bot.send_message(update.message.chat_id, "Bem vindo ao bot do MestreRuan!")
    self.ajuda(update, context)
    return
  def ajuda(self, update: Update, context: CallbackContext) -> None:
    context.bot.send_message(update.message.chat_id, "Digite o tipo de informação que deseja e depois o número da nota ou instalação. Por exemplo:\n")
    context.bot.send_message(update.message.chat_id, "leiturista 1012456598")
    context.bot.send_message(update.message.chat_id, "No momento temos as informações: TELEFONE, LOCALIZAÇÃO")
    context.bot.send_message(update.message.chat_id, "Estou trabalhando para trazer mais funções em breve")
    return

if __name__ == '__main__':
  bot = telbot()