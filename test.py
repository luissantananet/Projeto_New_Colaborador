import sqlite3
nomecompleto = "afdsagha"
banco = sqlite3.connect(r'.\dados\banco_cadastro.db')
cursor = banco.cursor()
cursor.execute(f"SELECT * FROM cadastro_colaborador WHERE nome_completo = '{nomecompleto}';")
dados = cursor.fetchall()




print(dados)
