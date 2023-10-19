class LeitorDeAcoes:
    def __init__ (self, caminho_do_arquivo: str = ''):
        self.caminho_do_arquivo = caminho_do_arquivo
        self.dados = []
    
    def processaArquivo(self, acao: str):
        with open(f'{self.caminho_do_arquivo}{acao}.txt', 'r') as arquivo_cotacao:
        # Pega todas as linhas do arquivo
            linhas = arquivo_cotacao.readlines()
            
            # formata linha a linha da forma que queremos
            self.dados = [linha.replace('\n', '').split(';') for linha in linhas]