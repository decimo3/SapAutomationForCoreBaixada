class Solicitacao():
    aplicacao: str = ''
    informacao: int = 0
    instancia: int = 0
    def __init__(self, aplicacao: str, informacao: int) -> None:
        self.aplicacao = aplicacao
        self.informacao = informacao
