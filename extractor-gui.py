#!/bin/venv python
# -*- coding: utf-8 -*-
# Autor: fant9sy <fant9sy@deadspace.wtf>
# Descrição: Extrai dados específicos de uma planilha do Excel

###################
#      TODOs      #
###################
# Adicionar as duas linhas acima das colunas na planilha com o título
    # BUSCA POR PROCEDIMENTO
    # O título segue o padrão: {Cidade} - {Procedimento} {Data iniíco - final}

# Não precisa pedir o arquivo de novo na hora de escolher outro QRadioButton (cliente_ocorrencia, atendente, ranking), manter o mesmo

# Tentar fazer um filtro por data, usando as datas que estarão na planilha clientes, quando baixada do Graces
# Aparecer todas as ocorrências quando pesquisar por n vezes (e.g atendente -> Ariane, procedimento -> Metodo Recover, n visitas -> 2)
    # Printar as duas linhas que aparecem a cliente

# Adicionar grade contornada de preto na planilha em todas as células preenchidas

###################
#     Exemplos    #
###################
####### SUB: cliente_ocorrencia
# Retorna os clientes que visitaram apenas uma vez a clínica
    # python3 extractor.py cliente_ocorrencia --arquivo planilha-teste.xlsx0

# Retorna os clientes que visitam a clínica 5 vezes
    # python3 extractor.py cliente_ocorrencia --arquivo planilha-teste.xlsx --n 5

######## SUB: atendente
# Retorna os clientes atendidos por Karol
    # python3 extractor.py atendente --arquivo planilha-teste.xlsx --nome_atendente karol

# Retorna os clientes atendidos por Karol 5 vezes
    # python3 extractor.py atendente --arquivo planilha-teste.xlsx --nome_atendente karol --n 5

import pandas as pd
import argparse
from colorama import Fore, Style
import sys
from PyQt5 import QtGui
from PyQt5.QtGui import QIntValidator
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QRadioButton, QFormLayout, QLineEdit, QLabel, QPushButton, QFileDialog, QHBoxLayout, QTextEdit, QMessageBox
from PyQt5.QtCore import QTextStream
from io import StringIO
import re

# Ignorando avisos do openpyxl
import warnings
warnings.simplefilter("ignore")

# Configurações
caminho_output = "/home/flame/self/mae/outputs/"
colunas_output = ['Cliente', 'Data comanda', 'Serviço/Produto', 'Profissional', 'Celular', 'Observação']
tamanho_colunas = {'Cliente': 22, 'Data comanda': 10.5, 'Serviço/Produto': 21, 'Profissional': 8, 'Celular': 18, 'Observação': 37}
coluna_cliente = "Cliente"
coluna_atendente = "Profissional"
coluna_procedimento = "Serviço/Produto"

# Funções de exibição
def info(message):
    print(f'[{Fore.BLUE}*{Style.RESET_ALL}] {message}')
def success(message):
    print(f'[{Fore.GREEN}+{Style.RESET_ALL}] {message}')
def error(message):
    print(f'[{Fore.RED}!{Style.RESET_ALL}] {message}')

def gen_output(index, output, custom_header=None):
    info(f"Enviando os resultados para planilha: {Fore.RED}{output}.xlsx{Style.RESET_ALL}")
    writer = pd.ExcelWriter(f'{caminho_output + output}.xlsx', engine='xlsxwriter')
    index.reset_index().to_excel(writer, index=False, columns=colunas_output) 
    workbook = writer.book  

    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]   
        for col, tamanho in tamanho_colunas.items():
            col_index = colunas_output.index(col)
            worksheet.set_column(col_index, col_index, tamanho) 
            if col in ['Data comanda', 'Celular']:
                worksheet.set_column(col_index, col_index, tamanho, cell_format=workbook.add_format({'align': 'center'}))   

    writer.save()

def extrair_clientes_por_ocorrencia(arquivo, nome_cliente, n, output):
    df = pd.read_excel(arquivo)

    # buscar pelo nome do cliente e retornar as linhas que o cliente aparece
    if nome_cliente:
        # TODO: quero conseguir filtrar por cliente usando apenas o primeiro nome, ou nomes incompletos, ao inves do nome exato
        index_final = df[df[coluna_cliente].str.contains(nome_cliente, case=False)]
    else:
        count_names = df[f'{coluna_cliente}'].value_counts()
        unique_names = count_names[count_names == n].index
        index_final = df[df[coluna_cliente].isin(unique_names)].copy()
    
    index_final['Observação'] = ''

    print(index_final.to_string(index=False, columns=colunas_output))

    if output:
        gen_output(index_final, output)

    success(f"Encontrados {Fore.RED}{len(index_final)}{Style.RESET_ALL} resultados")

def extrair_por_atendente(arquivo, nome_atendente, procedimento, n, output):
    df = pd.read_excel(arquivo)
    nome_atendente = nome_atendente.strip()
    linhas_atendente = df[df[f'{coluna_atendente}'].str.strip() == nome_atendente]

    if n:
        contagem_clientes = linhas_atendente[coluna_cliente].value_counts() # Verificar a contagem de clientes específicos para a atendente desejada
        clientes_n_vezes = contagem_clientes[contagem_clientes == n].index # Filtrar clientes que aparecem exatamente n vezes
        linhas_atendente = linhas_atendente[linhas_atendente[coluna_cliente].isin(clientes_n_vezes)] # Filtrar linhas para os clientes específicos que aparecem n vezes

        if linhas_atendente.empty:
            error(f"Nenhum resultado para atendente {Fore.RED}{nome_atendente}{Style.RESET_ALL} com clientes que apareceram {Fore.RED}{n}{Style.RESET_ALL} vezes.")
            return

        success(f"Resultados para atendente {Fore.RED}{nome_atendente}{Style.RESET_ALL} com clientes que apareceram {Fore.RED}{n}{Style.RESET_ALL} vezes:\n")

    if procedimento:
        # Verificar se o procedimento contém a substring desejada
        linhas_atendente = linhas_atendente[linhas_atendente[f'{coluna_procedimento}'].str.contains(procedimento, case=False)]
        
        if linhas_atendente.empty:
            error(f"Nenhum resultado para atendente {Fore.RED}{nome_atendente}{Style.RESET_ALL} com procedimento {Fore.RED}{procedimento}{Style.RESET_ALL}.")
            return

        success(f"Resultados para atendente {Fore.RED}{nome_atendente}{Style.RESET_ALL} com procedimento {Fore.RED}{procedimento}{Style.RESET_ALL}:\n")

    # Restante do seu código...
    linhas_atendente['Observação'] = ''
    print(linhas_atendente.to_string(index=False, columns=colunas_output))

    if output:
        gen_output(linhas_atendente, output)

    success(f"Encontrados {Fore.RED}{len(linhas_atendente)}{Style.RESET_ALL} resultados")

def gerar_ranking(arquivo, tipo, procedimento, nome_atendente, output):
    df = pd.read_excel(arquivo)

    if tipo == 'clientes': # Gera o ranking de clientes com mais visitas/procedimentos
        if procedimento: # Ranking por procedimento
            df_procedimento = df[df[f'{coluna_procedimento}'] == procedimento]
            count_names = df_procedimento[f'{coluna_cliente}'].value_counts()
            count_names = count_names.sort_values(ascending=False)

            if count_names.empty:
                error(f"Nenhum resultado para procedimento {Fore.RED}{procedimento}{Style.RESET_ALL}.")
                return
            
            success(f"Ranking de clientes que mais fizeram o procedimento {Fore.RED}{procedimento}{Style.RESET_ALL}:\n")
        else: # Ranking geral (clientes com mais visitas)
            count_names = df[f'{coluna_cliente}'].value_counts()
            count_names = count_names.sort_values(ascending=False)
            success("Ranking de clientes mais assíduos:\n")
            
        print(f"{count_names.reset_index().to_string(index=False, header=['CLIENTE', 'VISITAS'])}\n")

        if output:
            info(f"Enviando os resultados para planilha: {Fore.RED}{output}.xlsx{Style.RESET_ALL}")
            count_names.reset_index().to_excel(f'{caminho_output + output}.xlsx', index=False, header=['CLIENTE', 'VISITAS'])

    elif tipo == 'procedimentos': 
        if nome_atendente:
            nome_atendente = nome_atendente.strip()
            procedimentos_count = df[df[coluna_atendente].str.strip() == nome_atendente][f'{coluna_procedimento}'].value_counts()
        else:
            procedimentos_count = df[f'{coluna_procedimento}'].value_counts()

        procedimentos_count = procedimentos_count.sort_values(ascending=False)

        if procedimentos_count.empty:
            error("Nenhum resultado para ranking de procedimentos.")
            return

        success("Ranking de procedimentos mais realizados:\n")
        print(procedimentos_count.reset_index().to_string(index=False, header=['PROCEDIMENTO', 'QUANTIDADE']))

        if output:
            info(f"Enviando os resultados para planilha: {Fore.RED}{output}.xlsx{Style.RESET_ALL}")
            procedimentos_count.reset_index().to_excel(f'{caminho_output + output}.xlsx', index=False, header=['PROCEDIMENTO', 'QUANTIDADE'])


    elif tipo == 'atendentes': # Gera o ranking de atendentes
        atendentes_count = df[f'{coluna_atendente}'].value_counts()
        atendentes_count = atendentes_count.sort_values(ascending=False)

        if atendentes_count.empty:
            error("Nenhum resultado para ranking de atendentes.")
            return

        success("Ranking de profissionais com mais atendimentos:\n")
        print(atendentes_count.reset_index().to_string(index=False, header=['ATENDENTE', 'ATENDIMENTOS']))

        if output:
            info(f"Enviando os resultados para planilha: {Fore.RED}{output}.xlsx{Style.RESET_ALL}")
            atendentes_count.reset_index().to_excel(f'{caminho_output + output}.xlsx', index=False, header=['ATENDENTE', 'ATENDIMENTOS'])

class MyGUI(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Extractor de Dados')
        self.setGeometry(600, 400, 600, 400)

        # Layout for subcommand selection
        subcommand_layout = QVBoxLayout()

        self.radio_buttons = []
        subcommands = ['cliente_ocorrencia', 'atendente', 'ranking']

        for subcommand in subcommands:
            radio_button = QRadioButton(subcommand)
            radio_button.clicked.connect(self.onRadioButtonClicked)
            subcommand_layout.addWidget(radio_button)
            self.radio_buttons.append(radio_button)

        # Layout for argument fields
        self.argument_layout = QFormLayout()

        # Create and set default argument fields
        self.createArgumentFields('cliente_ocorrencia')

        # OK button to execute the selected subcommand
        self.ok_button = QPushButton('OK')
        self.ok_button.clicked.connect(self.executeSubcommand)
        
        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)  # Make it read-only
        self.output_text_edit.setMinimumHeight(100)

        # Main layout
        main_layout = QVBoxLayout()
        main_layout.addLayout(subcommand_layout)
        main_layout.addLayout(self.argument_layout)
        main_layout.addWidget(self.ok_button)
        main_layout.addWidget(self.output_text_edit)

        self.setLayout(main_layout)

    def onRadioButtonClicked(self):
        selected_subcommand = [button.text() for button in self.radio_buttons if button.isChecked()][0]
        self.createArgumentFields(selected_subcommand)

    def createArgumentFields(self, subcommand):
        # Clear existing argument fields, unless self.arquivo_edit and self.output_edit are the same as in the previous subcommand
        for i in reversed(range(self.argument_layout.count())):
            item = self.argument_layout.takeAt(i)
            if item.widget():
                item.widget().deleteLater()

        self.arquivo_edit = self.addFileInputField('Caminho para planilha do Excel')
        # Create argument fields based on the selected subcommand
        if subcommand == 'cliente_ocorrencia':
            self.nome_cliente_edit = self.addArgumentField('Nome do cliente')
            self.n_edit = self.addArgumentField('Quantidade de visitas', is_numeric=True)
            # self.output_edit = self.addArgumentField('Planilha de destino')

        elif subcommand == 'atendente':
            # self.arquivo_edit = self.addFileInputField('Caminho para planilha do Excel')
            self.nome_atendente_edit = self.addArgumentField('Nome da atendente')
            self.procedimento_edit = self.addArgumentField('Especificar procedimento')
            self.n_edit = self.addArgumentField('Quantidade de visitas', is_numeric=True)
            # self.output_edit = self.addArgumentField('Planilha de destino')

        elif subcommand == 'ranking':
            # self.arquivo_edit = self.addFileInputField('Caminho para planilha do Excel')
            self.tipo_edit = self.addArgumentField('Especificar tipo de ranking')
            self.procedimento_edit = self.addArgumentField('Especificar procedimento')
            self.nome_atendente_edit = self.addArgumentField('Nome da atendente')

        self.output_edit = self.addArgumentField('Planilha de destino')

    def addFileInputField(self, label_text):
        label = QLabel(label_text)
        edit = QLineEdit()
        browse_button = QPushButton('Browse')
        browse_button.clicked.connect(lambda: self.browseFile(edit))
        
        layout = QHBoxLayout()
        layout.addWidget(edit)
        layout.addWidget(browse_button)

        self.argument_layout.addRow(label, layout)
        return edit


    def addArgumentField(self, label_text, is_numeric=False):
        label = QLabel(label_text)
        edit = QLineEdit()
        if is_numeric:
            edit.setValidator(QIntValidator())
        self.argument_layout.addRow(label, edit)
        return edit

    def browseFile(self, edit):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_path, _ = QFileDialog.getOpenFileName(self, "Escolha a planilha do Excel", "", "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if file_path:
            edit.setText(file_path)

    def executeSubcommand(self):
        # Get the selected subcommand
        selected_subcommand = [button.text() for button in self.radio_buttons if button.isChecked()][0]

        # Get values from argument fields
        arquivo = self.arquivo_edit.text()
        output = self.output_edit.text()

        output_text = StringIO()
        sys.stdout = output_text

        try:
            if selected_subcommand == 'cliente_ocorrencia':
                nome_cliente = self.nome_cliente_edit.text()
                n = int(self.n_edit.text()) if self.n_edit.text() else 1
                # Call your function with the selected subcommand and arguments
                extrair_clientes_por_ocorrencia(arquivo, nome_cliente, n, output)

            elif selected_subcommand == 'atendente':
                nome_atendente = self.nome_atendente_edit.text()
                procedimento = self.procedimento_edit.text()
                n = int(self.n_edit.text()) if self.n_edit.text() else None
                # Call your function with the selected subcommand and arguments
                extrair_por_atendente(arquivo, nome_atendente, procedimento, n, output)

            elif selected_subcommand == 'ranking':
                tipo = self.tipo_edit.text()
                procedimento = self.procedimento_edit.text()
                nome_atendente = self.nome_atendente_edit.text()
                # Call your function with the selected subcommand and arguments
                gerar_ranking(arquivo, tipo, procedimento, nome_atendente, output)
        finally:
            sys.stdout = sys.__stdout__

            # Get the captured output, strip ANSI escape codes, and update the QTextEdit widget
            output_text.seek(0)
            text = self.strip_ansi_escape_codes(output_text.read())
            self.output_text_edit.setPlainText(text)

            # Show an alert message box only if a valid output file is provided
            if output:
                alert_message = f"Dados enviados para planilha {output}"
                self.showMessageBox("Informação", alert_message)

    def strip_ansi_escape_codes(self, text):
        # Use a regular expression to remove ANSI escape codes
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
        return ansi_escape.sub('', text)

    def showMessageBox(self, title, message):
        msg_box = QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.exec_()


# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     ex = MyGUI()
#     ex.show()
#     sys.exit(app.exec_())
        
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Retorna dados customizados de uma planilha do Excel.')

    subparsers = parser.add_subparsers(title='subcommands', dest='subcommand')
    subparsers.required = True

    # Subparser para subcomando "cliente_ocorrencia"
    unique_parser = subparsers.add_parser('cliente_ocorrencia', help='Retorna clientes baseado no número de visitas')
    unique_parser.add_argument('--arquivo', dest='arquivo', help='Caminho para planilha do Excel', required=True)
    unique_parser.add_argument('--n', type=int, help='Quantidade de visitas', nargs='?', default=1)
    unique_parser.add_argument('--output', dest='output', help='Planilha de destino', required=False)
    unique_parser.add_argument('--nome_cliente', dest='nome_cliente', help='Nome do cliente', required=False)

    # Subparser para subcomando "atendente"
    custom_times_parser = subparsers.add_parser('atendente', help='Filtro de clientes atendidos por x atendente')
    custom_times_parser.add_argument('--arquivo', dest='arquivo', help='Caminho para planilha do Excel', required=True)
    custom_times_parser.add_argument('--nome_atendente', dest='nome_atendente', help='Nome da atendente', required=True)
    custom_times_parser.add_argument('--procedimento', dest='procedimento', help='Especificar procedimento', required=False)
    custom_times_parser.add_argument('--n', type=int, help='Quantidade de visitas')
    custom_times_parser.add_argument('--output', dest='output', help='Planilha de destino', required=False)
    
    # Subparser para subcomando "ranking"
    gerar_ranking_parser = subparsers.add_parser('ranking', help='Ranking de clientes')
    gerar_ranking_parser.add_argument('--arquivo', dest='arquivo', help='Caminho para planilha do Excel', required=True)
    gerar_ranking_parser.add_argument('--tipo', dest='tipo', help='Especificar tipo de ranking', required=True)
    gerar_ranking_parser.add_argument('--procedimento', dest='procedimento', help='Especificar procedimento', required=False)
    gerar_ranking_parser.add_argument('--nome_atendente', dest='nome_atendente', help='Nome da atendente', required=False)
    gerar_ranking_parser.add_argument('--output', dest='output', help='Planilha de destino', required=False)

    args = parser.parse_args()

    if args.subcommand == 'cliente_ocorrencia':
        extrair_clientes_por_ocorrencia(args.arquivo, args.nome_cliente, args.n, args.output)
    elif args.subcommand == 'atendente':
        extrair_por_atendente(args.arquivo, args.nome_atendente, args.procedimento, args.n, args.output)
    elif args.subcommand == 'ranking':
        gerar_ranking(args.arquivo, args.tipo, args.procedimento, args.nome_atendente, args.output)
