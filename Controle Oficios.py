#Importações
import tkinter as tk
from tkinter import filedialog
import sys
import cv2
from PIL import Image
import numpy as np
import pytesseract
import re
import pandas as pd
import os

#Caminho Tesseract
# Função para obter o caminho do Tesseract
def get_tesseract_path():
    config_file = 'config.txt'
    
    if os.path.exists(config_file):
        # Se o arquivo de configuração existir, leia o caminho do Tesseract a partir dele
        with open(config_file, 'r') as f:
            tesseract_path = f.readline().strip()
    else:
        # Se o arquivo de configuração não existir, exiba a janela de seleção de arquivo
        janela = tk.Tk()
        janela.withdraw()  # Oculta a janela principal
        caminho_do_arquivo = filedialog.askopenfilename(title="Selecione o arquivo Tesseract OCR")
        janela.destroy()  # Fecha a janela após a seleção
        
        if not caminho_do_arquivo:
            # Caso o usuário cancele a seleção, você pode tratar isso aqui
            raise Exception("A seleção do caminho do Tesseract foi cancelada pelo usuário.")
        
        # Salve o caminho em um arquivo de configuração
        with open(config_file, 'w') as f:
            f.write(caminho_do_arquivo)
        
        tesseract_path = caminho_do_arquivo

    return tesseract_path

# Obtém o caminho do Tesseract
caminho_tesseract = get_tesseract_path()
print(f'Caminho do Tesseract: {caminho_tesseract}')

#Escolher Arquivo da Imagem
# Função para lidar com o arquivo selecionado
def selecionar_arquivo():
    global caminho_do_arquivo
    caminho_do_arquivo = filedialog.askopenfilename()
    if caminho_do_arquivo:
        label.config(text=f'Caminho do arquivo selecionado: {caminho_do_arquivo}')
        janela.destroy()
# Configuração da janela
janela = tk.Tk()
janela.title("Seleção de Arquivo")
# Botão para selecionar arquivo
btn_selecionar = tk.Button(janela, text="Selecionar Imagem do Oficio", command=selecionar_arquivo)
btn_selecionar.pack(pady=20)
# Rótulo para exibir o caminho do arquivo selecionado
label = tk.Label(janela, text="")
label.pack()
# Variável para armazenar o caminho do arquivo
caminho_do_arquivo = ""
# Iniciar a GUI
janela.mainloop()
# Agora você pode usar a variável "caminho_do_arquivo" em outras partes do seu programa
print(f'Caminho do arquivo selecionado: {caminho_do_arquivo}')

#Le Imagem
imagem = cv2.imread(fr'{caminho_do_arquivo}')
print (imagem)

#Transforma Imagem em texto
pytesseract.pytesseract.tesseract_cmd = caminho_tesseract
texto = pytesseract.image_to_string(imagem, lang='por')
print(texto)

#Extrair informações necessárias
padroes_rua = [r"Ruas ([^,]+)",r"rua ([^,]+)",r"Rua ([^,]+)"]
padroes_nome_vereador = [r"Mo:(.+)"]
padroes_n_oficio =[r'Oficio Nº (.+)' ]
padroes_servico =[r'limpeza e capinação',r'capinação e limpeza',r'aterrar',r'container',r'retirada de resíduos',r'placa poribitiva']
padroes_bairro = [r'Genipabu',r'Centro',r'Araturi',r'Cumbuco',]
# Encontra correspondências usando regex para cada padrão
ruas = []
for padrao_rua in padroes_rua:
    correspondencia1 = re.search(padrao_rua, texto)
    if correspondencia1:
        rua = correspondencia1.group(1).strip()
        ruas.append(rua)
nomes_vereador = []
for padrao_nome_vereador in padroes_nome_vereador:
    correspondencia2 = re.search(padrao_nome_vereador, texto)
    if correspondencia2:
        nome_vereador = correspondencia2.group(1).strip()
        nomes_vereador.append(nome_vereador)
ns_oficio = []
for padrao_n_oficio in padroes_n_oficio:
    correspondencia3 = re.search(padrao_n_oficio, texto)
    if correspondencia3:
        n_oficio = correspondencia3.group(1).strip()
        ns_oficio.append(n_oficio)
servicos = []
for padrao_servico in padroes_servico:
    correspondencia4 = re.search(padrao_servico, texto)
    if correspondencia4:
        servico = correspondencia4.group().strip()
        servicos.append(servico)
bairros = []
for padrao_bairro in padroes_bairro:
    correspondencia5 = re.search(padrao_bairro, texto)
    if correspondencia5:
        bairro = correspondencia5.group().strip()
        bairros.append(bairro)
# Extrai as informações se encontradas
if ruas:
    for i, rua in enumerate(ruas):
        print(f"Rua:{rua}")
else:
    print("Nenhuma rua encontrada.")

if nomes_vereador:
    for i, nome_vereador in enumerate(nomes_vereador):
        print(f"Nome do Vereador: {nome_vereador}")
else:
    print("Nenhum vereador encontrado.")
if ns_oficio:
    for i, n_oficio in enumerate(ns_oficio):
        print(f"Número Oficio: {n_oficio}")
else:
    print("Nenhum oficio encontrado.")
if servicos:
    for i, servico in enumerate(servicos):
        print(f"Tipo de Serviço: {servico}")
else:
    print("Nenhum oficio encontrado.")

#Transforma informações importantes em Dataframe
informacoes = {
    "Endereço": ruas,
    "Vereador": nomes_vereador,
    "N.Oficio": ns_oficio,
    "Tipo de Serviço:": servicos
}
dataframe = pd.DataFrame(informacoes)
print(dataframe)

#A partir daqui crendeciamento do google e importação para a planilha
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SPREADSHEET_ID = '1kAtn-w62aKq7lS1r_mmW3aZTcOlnrN1vFec5HXq25rU'


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        #Procura a ultima linha com algo escrito
        sheet_name = 'BASE'
        # Recupere os valores atuais na guia
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=f'{sheet_name}!A:A').execute()
        values = result.get('values', [])

        # Determine a última linha preenchida na coluna A
        last_row = len(values) if values else 0
        
        #Determina o range em que as informações serão adicionadas
        range_to_update = f'{sheet_name}!A{last_row + 1}'
        #Define quais valores devem ser adicionados
        valores_adicionar =[
            [n_oficio,servico,'Rua ' + rua,bairro,"",nome_vereador]
        
        ]
        #Adiciona os valores na planilha
        result = sheet.values().update(spreadsheetId = SPREADSHEET_ID,
                                     range = range_to_update,
                                     valueInputOption = "RAW",
                                     body ={"values": valores_adicionar}).execute()
        
    except HttpError as err:
        print(err)

if __name__ == '__main__':
    main()