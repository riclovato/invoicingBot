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
            print(df)
        except Exception as e:
            print(f"Erro ao ler o arquivo: {e}")
            return None

bot = InvoicingBot()
file_path = bot.select_file()
if file_path:
    bot.read_file(file_path)