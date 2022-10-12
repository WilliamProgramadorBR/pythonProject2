import tkinter as tk
import sqlite3
import pandas as pd

# conexao = sqlite3.connect('banco_clientes.db')
#
# c = conexao.cursor()
#
# c.execute(''' CREATE TABLE clientes (
#     nome text,
#     sobrenome text,
#     email text,
#     telefone text
#     )
# ''')
# conexao.commit()
#
# conexao.close()

janela = tk.Tk()
janela.title('ferramenta para mapear alunos com problemas')

# Labels:

label_nome = tk.Label(janela, text='Nome')
label_nome.grid(row=0, column=0, padx=10, pady=10)

label_Sobrenome = tk.Label(janela, text='Sobrenome')
label_Sobrenome.grid(row=1, column=0, padx=10, pady=10)

label_email = tk.Label(janela, text='E-mail')
label_email.grid(row=2, column=0, padx=10, pady=10)

label_Telefone = tk.Label(janela, text= 'Telefone')
label_Telefone.grid(row=3, column=0, padx=10, pady=10)




janela.mainloop()