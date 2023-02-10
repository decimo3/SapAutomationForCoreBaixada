import datetime

class Debito:
  def __init__(self, status, referencia, vencimento, valor, documento) -> None:
    self.status = True if (status == "@5C@") else False
    self.referencia = referencia
    self.vencimento = datetime.datetime.strptime(vencimento, "%d.%m.%Y").date()
    self.valor = valor
    self.documento = documento
    self.passivel = self.setPassivel()
  def setPassivel(self) -> bool:
    # Se o estado não estiver em "cobrando"
    if (self.status != "@5C@"): return False
    # Verifica se tem menos de 15 dias de vencido
    prazo_15_dias = self.vencimento - datetime.timedelta(days=15)
    if (self.vencimento < prazo_15_dias): return False
    return True

