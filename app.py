import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename


class InvoicingBot:
    def __init__(self):
        pass

    def select_file(self):
       # Tk().withdraw() # esconde janela principal do Tkinter
        file_path = askopenfilename(title="Selecione o arquivo de fatura", filetypes=[("Arquivos Excel", "*.xlsx *.xls")])

        if not file_path:
            print("Nenhum arquivo selecionado.")
            return None
        return file_path