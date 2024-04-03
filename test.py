import sqlite3

banco = sqlite3.connect(r'.\dados\banco_cadastro.db')
cursor = banco.cursor()
cursor.execute("select * from funcoes")
dados = cursor.fetchall()
tables = len(dados)



print(dados)
