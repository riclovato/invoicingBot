import time

import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from pywinauto.application import Application, ProcessNotFoundError


class InvoicingBot:

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
            if 'status' not in df.columns.str.lower():
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
try:
    app = Application(backend="uia").connect(path=r"C:\Program Files (x86)\Contoso, Inc\Contoso Invoicing\LegacyInvoicingApp.exe")
except(ProcessNotFoundError, TimeoutError):
    app = Application(backend="uia").start(r"C:\Program Files (x86)\Contoso, Inc\Contoso Invoicing\LegacyInvoicingApp.exe")
main_window = app.window(title="Contoso Invoicing")
time.sleep(10)
invoices_element = main_window.child_window(class_name="TextBlock", title="Invoices")
if invoices_element.exists():
    invoices_element.click_input()
    print("Clicando no elemento 'Invoices'")
else:
    print("Não foi possível encontrar o elemento 'Invoices'")

date_field = main_window.child_window(auto_id="txtDate103", control_type="Edit")

# Interagir com o campo de texto "Date"
if date_field.exists():
    date_field.set_focus()
    date_field.type_keys("^a")
    date_field.type_keys("08/11/2024")
    print("Digitando a data no campo 'Date'")
else:
    print("Não foi possível encontrar o campo 'Date'")

account_field = main_window.child_window(auto_id="txtAccount103", control_type="Edit")
if account_field.exists():
    account_field.set_focus()
    account_field.type_keys("^a")
    account_field.type_keys("NomeCont")
    print("Preenchendo a conta.")
else:
    print("Não foi possível encontrar o campo 'account'")

contact_field = main_window.child_window(auto_id="txtContactEmail103", control_type="Edit")
if contact_field.exists():
    contact_field.set_focus()
    contact_field.type_keys("^a")
    contact_field.type_keys("contato@email.com")
    print("Preenchendo o contato.")
else:
    print("Não foi possível encontrar o campo 'contato'")

amount_field = main_window.child_window(auto_id="txtAmount103", control_type="Edit")
if amount_field.exists():
    amount_field.set_focus()
    amount_field.type_keys("^a")
    amount_field.type_keys("1324562")
    print("Preenchendo o valor.")
else:
    print("Não foi possível encontrar o campo 'amount'")

status_field = main_window.child_window(auto_id="cmbStatusChooser103", control_type="ComboBox")
if status_field.exists():
    status_field.select('Paid')
    print("Atualizando o status.")
else:
    print("Não foi possível encontrar o campo 'status'")

save_button = main_window.child_window(auto_id="btnSave", control_type="Button")
save_button.click()