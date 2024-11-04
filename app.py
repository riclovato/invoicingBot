import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename


class InvoicingBot:
    def __init__(self):
        pass

    def select_file(self):
        Tk().withdraw() # esconde janela principal do Tkinter
        file_path = askopenfilename(title="Selecione o arquivo de fatura", filetypes=[("Arquivos Excel", "*.xlsx *.xls")])

        if not file_path:
            print("Nenhum arquivo selecionado.")
            return None
        return file_path

    def read_file(self, file_path):
        try:
            df = pd.read_excel(file_path)

            if 'date' not in df.columns.str.lower() and 'data'not in df.columns.str.lower():
                print("O arquivo não possui a coluna 'Data'")
                return None
            if 'account' not in df.columns.str.lower() and 'conta' not in df.columns.str.lower():
                print("O arquivo não possui a coluna 'Conta'")
                return None
            if 'contact' not in df.columns.str.lower() and 'contato' not in df.columns.str.lower():
                print("O arquivo não possui a coluna 'Contato'")
                return None
            if 'amount' not in df.columns.str.lower() and 'valor' not in df.columns.str.lower():
                print("O arquivo não possui a coluna 'Valor'")
                return None
            if 'Status' not in df.columns.str.lower():
                print("O arquivo não possui a coluna 'Status'")
                return None

            print(df)
        except Exception as e:
            print(f"Erro ao ler o arquivo: {e}")
            return None

bot = InvoicingBot()
file_path = bot.select_file()
if file_path:
    bot.read_file(file_path)