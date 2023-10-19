from openpyxl import Workbook

class GerenciadorDePlanilhas:
    def __init__(self):
        self.workbook = Workbook()
        self.planilha_ativa = None

    def criaPlanilha(self, titulo: str = ''):
        nova_planilha = self.workbook.create_sheet(titulo)
        self.workbook.active = nova_planilha
        self.planilha_ativa = nova_planilha

        return nova_planilha