import os
import openpyxl
from openpyxl import Workbook
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QMessageBox
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import pandas as pd

import sqlite3
# Criando o Bando de Dados
# Caminho para a pasta onde o banco de dados será armazenado
pasta_dados = r'.\dados'
# Verifica se a pasta "dados" existe, se não, cria a pasta
if not os.path.exists(pasta_dados):
    os.makedirs(pasta_dados)
# Caminho completo para o arquivo do banco de dados
database = os.path.join(pasta_dados, 'banco_cadastro.db')
# Verifica se o arquivo do banco de dados já existe
if not os.path.exists(database):
    # Conecta ao banco de dados (isso criará o arquivo se ele não existir)
    banco = sqlite3.connect(database)
    # Fechando a conexão
    banco.close()
# criando tabelas se ele nao exixtir
banco = sqlite3.connect(database)
cursor = banco.cursor()
# Cria a tabela 'tabela' se ela não existir
cursor.execute("""CREATE TABLE IF NOT EXISTS tabela ( 
id INTEGER PRIMARY KEY AUTOINCREMENT, 
diaria decimal(5,2) NOT NULL, 
hextra decimal(5,2) NOT NULL, 
vtransp decimal(5,2) NOT NULL,
vref decimal(5,2) NOT NULL);""")
# Cria a tabela 'funcoes' se ela não existir
cursor.execute("""CREATE TABLE IF NOT EXISTS funcoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    descricao TEXT NOT NULL);""")
# Salva as alterações no banco de dados
banco.commit()
# Fecha a conexão com o banco de dados
banco.close()

# Caminho do arquivo
arquivo_xlsx = 'RegistrosColaboradores.xlsx'
def gerarPlanilha():
    # Verifica se o arquivo já existe
    if not os.path.isfile(arquivo_xlsx):
        # Criando um novo livro
        wb = Workbook()
        # Renomeando a aba padrão para 'Registro'
        wsRegistro =wb.active
        wsRegistro.title = "Registros"
        # Criando e renomeando as outras abas
        wsColaboradores = wb.create_sheet("Colaboradores")
        wsFuncoes = wb.create_sheet("Funções")
        wsTaxas = wb.create_sheet("Taxas")
        # Salvar
        wb.save(r'.\RegistrosColaboradores.xlsx')
    else:
        QMessageBox.information(frm_principal, "Aviso", f"O arquivo {arquivo_xlsx} já existe.")

def salvarRegistro():
    datainicial = frm_principal.datainicial.text()
    datafinal = frm_principal.datafinal.text()
    nome = frm_principal.edt_nome.text()
    advale = frm_principal.edt_advale.text()
    dias = frm_principal.edt_dias.text()
    he = frm_principal.edt_he.text()
    sobtotal = frm_principal.edt_subtotal.text()
    total = frm_principal.edt_total.text()
    vale = frm_principal.edt_vale.text()
    vr = frm_principal.edt_vr.text()
    vt = frm_principal.edt_vt.text()
    try:
       
        if ids == "":
            cursor.execute("INSERT INTO registro VALUES(NULL,'"+datainicial+"','"+datafinal+"','"+nome+"','"+dias+"','"+he+"','"+vr+"','"+vt+"','"+advale+"','"+vale+"','"+sobtotal+"','"+total+"');")
            banco.commit()
            banco.close()
            frm_principal.edt_nome.setText('')
            frm_principal.edt_advale.setText('')
            frm_principal.edt_dias.setText('')
            frm_principal.edt_he.setText('')
            frm_principal.edt_subtotal.setText('')
            frm_principal.edt_total.setText('')
            frm_principal.edt_vale.setText('')
            frm_principal.edt_vr.setText('')
            frm_principal.edt_vt.setText('')
            QMessageBox.information(frm_principal, "Aviso", "Registro cadastrado com sucesso")
        else:
            cursor.execute("UPDATE registro SET data_inicial = '"+datainicial+"', data_final = '"+datafinal+"',nome_completo = '"+nome+"',dias_tr = '"+dias+"', he = '"+he+"', vr = '"+vr+"', vt = '"+vt+"',ad_vale = '"+advale+"', vale = '"+vale+"', subtotal = '"+sobtotal+"', total = '"+total+"'")
            banco.commit()
            banco.close()
            frm_principal.edt_nome.setText('')
            frm_principal.edt_advale.setText('')
            frm_principal.edt_dias.setText('')
            frm_principal.edt_he.setText('')
            frm_principal.edt_subtotal.setText('')
            frm_principal.edt_total.setText('')
            frm_principal.edt_vale.setText('')
            frm_principal.edt_vr.setText('')
            frm_principal.edt_vt.setText('')
            QMessageBox.information(frm_principal, "Aviso", "Registro atualizado com sucesso")
    except sqlite3.Error as erro:
        print("Erro ao cadastrar registro: ",erro)

def cadastroColaborador():
    id = frm_principal.edt_id.text()
    nomecompleto = frm_principal.edt_nome.text()
    funcao = frm_principal.comboBox_funcao.currentText()
    cpf = frm_principal.edt_cpf.text()
    rg = frm_principal.edt_rg.text()
    cnh = frm_principal.edt_cnh.text()
    endereco = frm_principal.edt_endereco.text()
    numero = frm_principal.edt_numeroEnd.text()
    bairro = frm_principal.edt_bairro.text()
    cidade = frm_principal.comboBox_cidade.currentText()
    uf = frm_principal.comboBox_uf.currentText()
    try:
        banco = sqlite3.connect(r'.\dados\banco_cadastro.db') 
        cursor = banco.cursor()
        # cria o bando se ele nao exixtir 
        cursor.execute("""CREATE TABLE IF NOT EXISTS cadastro_colaborador ( 
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        nome_completa varchar(100)NOT NULL, 
        funcao varchar(100)NOT NULL, 
        cpf varchar(100)NOT NULL,
        rg varchar(100), 
        cnh varchar(100), 
        endereco varchar(100),
        numero varchar(10),
        bairro varchar(100), 
        cidade varchar(100), 
        uf varchar(2));""")
        # verifica se o colaborador já existe 
        if id == "":
            # inserir dados na tabela
            cursor.execute("INSERT INTO cadastro_colaborador VALUES(NULL,'"+nomecompleto+"','"+funcao+"','"+cpf+"','"+rg+"','"+cnh+"','"+endereco+"','"+numero+"','"+bairro+"','"+cidade+"','"+uf+"')")
            banco.commit()
            banco.close()
            frm_principal.edt_nome.setText('')
            frm_principal.edt_cpf.setText('')
            frm_principal.edt_rg.setText('')
            frm_principal.edt_cnh.setText('')
            frm_principal.edt_endereco.setText('')           
            QMessageBox.information(frm_principal, "Aviso", "Colaborador cadastrado com sucesso")
        else:
            cursor.execute("UPDATE cadastro_colaborador SET nome_completa = '"+nomecompleto+"', funcao = '"+funcao+"',cpf = '"+cpf+"', rg = '"+rg+"', cnh = '"+cnh+"', endereco = '"+endereco+"', numero = '"+numero+"', bairro = '"+endereco+"', cidade = '"+cidade+"', uf = '"+uf+"' WHERE id = '"+id+"'")
            banco.commit()
            banco.close()
            frm_principal.edt_nome.setText('')
            frm_principal.edt_cpf.setText('')
            frm_principal.edt_rg.setText('')
            frm_principal.edt_cnh.setText('')
            frm_principal.edt_endereco.setText('')
            frm_principal.show()
            QMessageBox.information(frm_principal, "Aviso", "Colaborador atualizado com sucesso")
    except sqlite3.Error as erro:
        print("Erro ao inserir os dados: ",erro)
        QMessageBox.about(frm_principal, "ERRO","Erro ao inserir os dados")
        banco.close()   

# Função unificada para capturar os dados do formulário e adicionar na aba 'Taxas'
def cadastro_e_adicionar_taxas():
    # Captura os dados do formulário
    diaria = str(frm_principal.edt_diariaCad.text().replace(',','.')) # 85,00
    hextra = str(frm_principal.edt_hextraCad.text().replace(',','.')) # 16,00
    vtrans = str(frm_principal.edt_vtranspCad.text().replace(',','.')) # 12,30
    vresf = str(frm_principal.edt_vrefCad.text().replace(',','.')) # 17,00
    try:
        banco = sqlite3.connect(r'.\dados\banco_cadastro.db') 
        cursor = banco.cursor()
        cursor.execute("select * from tabela")
        dados = cursor.fetchall()
        tables = len(dados)
        if tables == 0:
            # inserir dados na tabela
            cursor.execute("INSERT INTO tabela VALUES(NULL,'"+diaria+"','"+hextra+"','"+vtrans+"','"+vresf+"')")
            banco.commit()
            banco.close()
            QMessageBox.information(frm_principal, "Aviso", "Tabela cadastrado com sucesso")
        else:
            cursor.execute("UPDATE tabela SET diaria = '"+diaria+"', hextra = '"+hextra+"',vtransp = '"+vtrans+"', vref = '"+vresf+"'")
            banco.commit()
            banco.close()
            QMessageBox.information(frm_principal, "Aviso", "Tabela atualizado com sucesso")
    except sqlite3.Error as erro:
        print("Erro ao inserir os dados: ",erro)
        QMessageBox.about(frm_principal, "ERRO","Erro ao inserir os dados")
    banco.close()

def preencherComboBoxFuncao():
    # Caminho para o arquivo do banco de dados
    database = r'.\dados\banco_cadastro.db'
    # Conecta ao banco de dados
    banco = sqlite3.connect(database)
    # Cria um cursor para executar operações no banco de dados
    cursor = banco.cursor()
    # Seleciona todas as descrições da tabela 'funcoes'
    cursor.execute("SELECT descricao FROM funcoes")
    # Recupera todos os resultados
    funcoes = cursor.fetchall()
    # Lista para armazenar as descrições
    lista_funcoes = [funcao[0] for funcao in funcoes]
    # Fecha a conexão com o banco de dados
    banco.close()
    return lista_funcoes
def cadastroFuncao():
    funcao = frm_principal.comboBox_funcaoCad.currentText()
    
    banco = sqlite3.connect(r'.\dados\banco_cadastro.db') 
    cursor = banco.cursor()
    # Verifica se a descrição já existe na tabela 'funcoes'
    cursor.execute("SELECT id FROM funcoes WHERE descricao = ?", (funcao,))
    dados = cursor.fetchone()
    
    try:
        if dados is None:
            # Insere a nova função na tabela 'funcoes' se ela não existir
            cursor.execute("INSERT INTO funcoes (descricao) VALUES (?)", (funcao,))
            banco.commit()
            QMessageBox.information(frm_principal, "Aviso", "Função cadastrada com sucesso.")
    except sqlite3.Error as erro:
        print("Erro ao inserir os dados: ",erro)
        QMessageBox.about(frm_principal, "ERRO","Erro ao inserir os dados")
    banco.close()
    frm_principal.show()
def excluirFuncao():
    # Obtém o texto do item selecionado no comboBox
    funcao = frm_principal.comboBox_funcaoCad.currentText()
    # Conecta ao banco de dados
    banco = sqlite3.connect(r'.\dados\banco_cadastro.db')
    cursor = banco.cursor()
    try:
        # Exclui a função da tabela 'funcoes' com base na descrição
        cursor.execute("DELETE FROM funcoes WHERE descricao = ?", (funcao,))
        banco.commit()
        # Verifica se a função foi realmente excluída
        if cursor.rowcount > 0:
            QMessageBox.information(frm_principal, "Aviso", "Função excluída com sucesso.")
        else:
            QMessageBox.warning(frm_principal, "Aviso", "Função não encontrada.")
    except sqlite3.Error as erro:
        print("Erro ao excluir a função: ", erro)
        QMessageBox.critical(frm_principal, "ERRO", "Erro ao excluir a função.")
    finally:
        # Fecha a conexão com o banco de dados
        banco.close()

def calcularRegistro():
    # +
    dias = str(frm_principal.edt_dias.text()).replace(',','.')
    he = str(frm_principal.edt_he.text()).replace(',','.')
    vr = str(frm_principal.edt_vr.text()).replace(',','.')
    vt = str(frm_principal.edt_vt.text()).replace(',','.')
    # -
    advale = str(frm_principal.edt_advale.text()).replace(',','.')
    vale = str(frm_principal.edt_vale.text()).replace(',','.')
    # tabela
    diaria = str(frm_principal.edt_diaria.text()).replace(',','.')
    hextra = str(frm_principal.edt_hextra.text()).replace(',','.')
    vtransp = str(frm_principal.edt_vtransp.text()).replace(',','.')
    vref = str(frm_principal.edt_vref.text()).replace(',','.')
    # calcular
    dias1 = float(dias) if dias else 0.0
    he1 = float(he) if he else 0.0
    vr1 = float(vr) if vr else 0.0
    vt1 = float(vt) if vt else 0.0
    advale1 = float(advale) if advale else 0.0
    vale1 = float(vale) if vale else 0.0
    diaria1 = float(diaria) if diaria else 0.0
    hextra1 = float(hextra) if hextra else 0.0
    vref1 = float(vref) if vref else 0.0
    vtransp1 = float(vtransp) if vtransp else 0.0
    tdias = (dias1 * diaria1)
    the = (he1 * hextra1)
    tvr = (vr1 * vref1)
    tvt = (vt1 * vtransp1)
    sobtotal = (tdias + the + tvr + tvt)
    subt = (advale1 + vale1)
    total = (sobtotal - subt)

    frm_principal.edt_subtotal.setText("{:.2f}".format(sobtotal))
    frm_principal.edt_total.setText("{:.2f}".format(total))

def gerarWord ():
    # pegar dados da planilha
    for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
        nome_curso = linha[0].value
        nome_participante = linha[1].value
        tipo_participacao = linha[2].value
        data_inicio = linha[3].value
        data_final = linha[4].value
        carga_horaria = linha[5].value
        data_emissao = linha[6].value
            
        # Definindo a fonte a ser usada
        fonte_nome = ImageFont.truetype('./tahomabd.ttf',90)
        fonte_geral = ImageFont.truetype('./tahoma.ttf',80)
        fonte_data = ImageFont.truetype('./tahoma.ttf',55)
        
        # definir certificado
        image = Image.open('./certificado_padrao.jpg')
        desenhar = ImageDraw.Draw(image)
        
        # Transferir os dados da planilha para a imagem do certificado
        desenhar.text((1020,827), nome_participante,fill='black',font=fonte_nome)
        desenhar.text((1060,950),nome_curso, fill='black',font=fonte_geral)
        desenhar.text((1435,1065),tipo_participacao, fill='black',font=fonte_geral)
        desenhar.text((1480, 1182),str(carga_horaria),fill='black',font=fonte_geral)
        
        desenhar.text((750, 1770),data_inicio,fill='blue',font=fonte_data)
        desenhar.text((750, 1930),data_final,fill='blue',font=fonte_data)
        
        desenhar.text((2220, 1930),data_emissao,fill='blue',font=fonte_data)
        image.save(f'./teste/{indice} {nome_participante} certificado.png')
def atualizarInterface():
    # Atualiza o comboBox com funções
    frm_principal.comboBox_funcaoCad.clear()
    frm_principal.comboBox_funcaoCad.addItems(preencherComboBoxFuncao())
    
    # Atualiza os campos de texto com as taxas mais recentes
    try:
        banco = sqlite3.connect(r'.\dados\banco_cadastro.db') 
        cursor = banco.cursor()
        cursor.execute("SELECT * FROM tabela")
        dados_lidos = cursor.fetchall()
        frm_principal.edt_diariaCad.setText(str('%.2f'%dados_lidos[0][1]).replace('.',','))
        frm_principal.edt_hextraCad.setText(str('%.2f'%dados_lidos[0][2]).replace('.',','))
        frm_principal.edt_vtranspCad.setText(str('%.2f'%dados_lidos[0][3]).replace('.',','))
        frm_principal.edt_vrefCad.setText(str('%.2f'%dados_lidos[0][4]).replace('.',','))
        frm_principal.edt_diaria.setText(str('%.2f'%dados_lidos[0][1]).replace('.',','))
        frm_principal.edt_hextra.setText(str('%.2f'%dados_lidos[0][2]).replace('.',','))
        frm_principal.edt_vtransp.setText(str('%.2f'%dados_lidos[0][3]).replace('.',','))
        frm_principal.edt_vref.setText(str('%.2f'%dados_lidos[0][4]).replace('.',','))
    except sqlite3.Error as erro:
        print("Erro ao atualizar os dados: ", erro)



if __name__ == '__main__':
    App = QtWidgets.QApplication([])
    frm_principal = uic.loadUi(r'.\frms\frm_principal.ui')
    # Conecta os eventos aos métodos correspondentes
    frm_principal.btn_GerarTabela.clicked.connect(gerarPlanilha)
    frm_principal.btn_salvarTabela.clicked.connect(lambda: [cadastro_e_adicionar_taxas(), atualizarInterface()])
    frm_principal.btn_salvarFucao.clicked.connect(lambda: [cadastroFuncao(), atualizarInterface()])
    frm_principal.btn_excluirFucao.clicked.connect(lambda: [excluirFuncao(), atualizarInterface()])
    
    # Atualiza a interface ao iniciar o aplicativo
    atualizarInterface()
    frm_principal.show()
    App.exec()