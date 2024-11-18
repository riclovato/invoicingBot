import time

import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from pywinauto.application import Application, ProcessNotFoundError


class InvoicingBot:
    def open_program(self):
        try:
            app = Application(backend="uia").connect(
                path=r"C:\Program Files (x86)\Contoso, Inc\Contoso Invoicing\LegacyInvoicingApp.exe")
        except(ProcessNotFoundError, TimeoutError):
            app = Application(backend="uia").start(
                r"C:\Program Files (x86)\Contoso, Inc\Contoso Invoicing\LegacyInvoicingApp.exe")
        main_window = app.window(title="Contoso Invoicing")
        time.sleep(10)
        invoices_element = main_window.child_window(class_name="TextBlock", title="Invoices")
        if invoices_element.exists():
            invoices_element.click_input()
            print("Clicando no elemento 'Invoices'")
        else:
            print("Não foi possível encontrar o elemento 'Invoices'")
        return main_window
    def select_file(self):
        Tk().withdraw() # esconde janela principal do Tkinter
        file_path = askopenfilename(title="Selecione o arquivo de fatura", filetypes=[("Arquivos Excel", "*.xlsx *.xls")])

        if not file_path:
            print("Nenhum arquivo selecionado.")
            return None
        return file_path

    def read_file(self, file_path):
        main_window = self.open_program()

        try:

            df = pd.read_excel(file_path)


            print(df.columns)
            required_columns = ['Date', 'Account', 'Contact', 'Amount', 'Status']
            for col in required_columns:
                if col.lower() not in [c.lower() for c in df.columns]:
                    print(f"O arquivo não possui a coluna '{col}'")
                    return None
            #itera sobre cada linha do DataFrame
            for index, row in df.iterrows():
                try:
                    date_field = main_window.child_window(auto_id="txtDate103", control_type="Edit")
                    if date_field.exists():
                        date_field.set_focus()
                        date_field.type_keys("^a")
                        date_value = pd.to_datetime(row['Date'])
                        date_field.type_keys(date_value.strftime('%d/%m/%Y'))

                    account_field = main_window.child_window(auto_id="txtAccount103", control_type="Edit")
                    if account_field.exists():
                        account_field.set_focus()
                        account_field.type_keys("^a")
                        account_field.type_keys(str(row['Account']))

                    contact_field = main_window.child_window(auto_id="txtContactEmail103", control_type="Edit")
                    if contact_field.exists():
                        contact_field.set_focus()
                        contact_field.type_keys("^a")
                        contact_field.type_keys(str(row['Contact']))

                    amount_field = main_window.child_window(auto_id="txtAmount103", control_type="Edit")
                    if amount_field.exists():
                        amount_field.set_focus()
                        amount_field.type_keys("^a")
                        amount_field.type_keys(str(row['Amount']))

                    status_field = main_window.child_window(auto_id="cmbStatusChooser103", control_type="ComboBox")
                    if status_field.exists():
                        status_field.select(str(row['Status']))

                    save_button = main_window.child_window(auto_id="btnSave", control_type="Button")
                    save_button.click()

                    new_button = main_window.child_window(auto_id="btnNew", control_type="Button")
                    new_button.click()
                except Exception as row_error:
                    print(f"Erro ao processar registro {index + 1}: {row_error}")
            print("Todos os registros processados")
            return df



        except Exception as e:
            print(f"Erro ao ler o arquivo: {e}")
            return None

bot = InvoicingBot()
file_path = bot.select_file()
if file_path:
    bot.read_file(file_path)


