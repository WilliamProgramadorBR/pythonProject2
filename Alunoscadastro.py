import tkinter as tk
import sqlite3

import pandas
import pandas as pd
from tkinter import messagebox as M
import win32com.client as win32




conexao = sqlite3.connect('banco_clientes.db')

c = conexao.cursor()

# c.execute(''' CREATE TABLE clientes (
#     nome text,
#     sobrenome text,
#     email text,
#     telefone text,
#     Id
#     )
# ''')
# conexao.commit()
#
# conexao.close()


def _cadastrar_alunos():
    conexao = sqlite3.connect('banco_clientes.db')
    nome=entry_nome.get()
    sobrenome=entry_Sobrenome.get()
    email=entry_email.get()
    telefone=entry_Telefone.get()
    Id=entry_Id.get()
    if nome == "" or sobrenome== "" or email=="" or telefone == "" and Id:
        M.showerror("Erro", "Preenche os campos")
        print(nome,email,sobrenome,telefone,Id)
    else:
        c = conexao.cursor()


        c.execute(f"SELECT * FROM clientes WHERE nome='{entry_Id.get()}' or email='{entry_email.get()}' ")
        retorno = c.fetchall()
        if retorno !=[]:
            M.showerror("ERRO", "Usuário/email já cadastrado")
        else:
            c.execute(" INSERT INTO clientes VALUES (:nome, :sobrenome, :email, :telefone, :Id)",
                      {
                          'nome': entry_nome.get(),
                          'sobrenome': entry_Sobrenome.get(),
                          'email': entry_email.get(),
                          'telefone': entry_Telefone.get(),
                          'Id': entry_Id.get()
                      }
                      )

        conexao.commit()

        conexao.close()

        entry_nome.delete(0, "end")
        entry_email.delete(0, "end")
        entry_Telefone.delete(0, "end")
        entry_Sobrenome.delete(0, "end")
        entry_Id.delete(0, "end")


def _exporta_alunos():
     conexao = sqlite3.connect('banco_clientes.db')

     c = conexao.cursor()

     c.execute("SELECT * FROM clientes")
     clientes_cadastrados = c.fetchall()
     clientes_cadastrados = pd.DataFrame(clientes_cadastrados, columns=['nome','sobrenome','email','telefone','Id'])
     print(clientes_cadastrados)
     clientes_cadastrados.to_excel('Alunos.xlsx')
     conexao.commit()

     conexao.close()


def _enviar_email(anexo=None):
    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    # configurar as informações do seu e-mail
    email.To = "william100william@hotmail.com; willian-100boladao@hotmail.com"
    email.Subject = "E-mail automático do Python"
    email.HTMLBody = f"""
    <p>Olá Lira, aqui é o código Python</p>


    <p>Olá eu sou um programa Python, por favor não responda a esse e-mail</p>

    <p>Abs,</p>
    <p>Código Python</p>
    """
    anexo = (r"C:\Users\willi\PycharmProjects\pythonProject2\Alunos.xlsx")
    email.Attachments.Add(anexo)

    email.Send()
    print("Email Enviado")








janela = tk.Tk()
janela.geometry("500x350")
janela.resizable(False,False)
janela.title('ferramenta para mapear alunos com problemas')

# Labels:

label_nome = tk.Label(janela, text='Nome', width=30)
label_nome.grid(row=0, column=0, padx=10, pady=10)

label_Sobrenome = tk.Label(janela, text='Sobrenome')
label_Sobrenome.grid(row=1, column=0, padx=10, pady=10)

label_email = tk.Label(janela, text='E-mail')
label_email.grid(row=2, column=0, padx=10, pady=10)

label_Telefone = tk.Label(janela, text= 'Telefone')
label_Telefone.grid(row=3, column=0, padx=10, pady=10)

label_Id = tk.Label(janela, text= 'Id')
label_Id.grid(row=4, column=0, padx=10, pady=10)

# Entrys:


entry_nome = tk.Entry(janela, width_=30)
entry_nome.grid(row=0, column=1, padx=10, pady=10)

entry_Sobrenome = tk.Entry(janela, width_=30)
entry_Sobrenome.grid(row=1, column=1, padx=10, pady=10)

entry_email = tk.Entry(janela, width_=30)
entry_email.grid(row=2, column=1, padx=10, pady=10)

entry_Telefone = tk.Entry(janela, width_=30)
entry_Telefone.grid(row=3, column=1, padx=10, pady=10)

entry_Id = tk.Entry(janela, width_=30)
entry_Id.grid(row=4, column=1, padx=10, pady=10)

# botões:

botao_cadastrar = tk.Button(janela, text='Cadastrar Alunos', command_=_cadastrar_alunos)
botao_cadastrar.grid(row=5, column=0, padx=10, pady=10, columnspan=2, ipadx=80)

botao_exportar = tk.Button(janela, text='Exportar Base de alunos', command_=_exporta_alunos)
botao_exportar.grid(row=6, column=0, padx=10, pady=10, columnspan=2, ipadx=80)

botao_Enviar = tk.Button(janela, text='Enviar para email', command_=_enviar_email)
botao_Enviar.grid(row=7, column=0, padx=10, pady=10, columnspan=2, ipadx=80)




janela.mainloop()

