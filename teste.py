import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# Função para selecionar o arquivo PDF
def selecionar_pdf():
    global caminho_pdf
    caminho_pdf = filedialog.askopenfilename(
        filetypes=[("Arquivos PDF", "*.pdf")])
    if caminho_pdf:
        messagebox.showinfo("Arquivo Selecionado", f"PDF selecionado: {caminho_pdf}")

# Função principal para enviar mensagens e PDFs
def enviar_mensagens():
    if not caminho_pdf:
        messagebox.showerror("Erro", "Por favor, selecione um arquivo PDF.")
        return

    webbrowser.open('https://web.whatsapp.com/')
    sleep(5)

    workbook = openpyxl.load_workbook('clientes.xlsx')
    pagina_clientes = workbook['DM25']

    for linha in pagina_clientes.iter_rows(min_row=2):
        nome = linha[0].value
        telefone = linha[1].value

        mensagem = f'Opa, {nome}! Tudo bem? Estou enviando um arquivo importante junto com essa mensagem.'
        try:
            link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link_mensagem_whatsapp)
            sleep(10)

            # Enviar mensagem
            pyautogui.click(x=4150, y=1025)
            sleep(2)

            # Anexar e enviar o PDF
            pyautogui.click(x=4500, y=1200)  # Coordenada para botão de anexar (ajuste necessário)
            sleep(1)
            pyautogui.write(caminho_pdf)
            pyautogui.press('enter')
            sleep(2)

        except:
            print(f'Não foi possível enviar mensagem para {nome}')
            with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                arquivo.write(f'{nome},{telefone}{os.linesep}')

# Criar interface gráfica
app = tk.Tk()
app.title("Envio de Mensagens e PDFs")
app.geometry("400x200")

tk.Label(app, text="Selecione o arquivo PDF para enviar:").pack(pady=10)
tk.Button(app, text="Selecionar PDF", command=selecionar_pdf).pack(pady=5)
tk.Button(app, text="Iniciar Envio", command=enviar_mensagens).pack(pady=20)

app.mainloop()
