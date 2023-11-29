import requests
import xml.etree.ElementTree as ET
from tabulate import tabulate
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from tkinter import Tk, filedialog

now = datetime.now()
formatted_date = now.strftime("%Y-%m-%d-%H-%M")

def consultar_awb(site_id, password, awb, historico=False):
    url = 'https://xmlpi-ea.dhl.com/XMLShippingServlet'
    headers = {'Content-Type': 'application/xml'}

    data = f'''<?xml version="1.0" encoding="UTF-8"?>
        <req:KnownTrackingRequest xmlns:req="http://www.dhl.com" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.dhl.com TrackingRequestKnown.xsd" schemaVersion="1.0">
            <Request>
                <ServiceHeader>
                    <SiteID>{site_id}</SiteID>
                    <Password>{password}</Password>
                </ServiceHeader>
            </Request>
            <LanguageCode>BR</LanguageCode>
            <AWBNumber>{awb}</AWBNumber>
            <LevelOfDetails>ALL_CHECK_POINTS</LevelOfDetails>
        </req:KnownTrackingRequest>'''

    # Faz a requisição
    response = requests.post(url, headers=headers, data=data)

    # Verifica se a requisição foi bem-sucedida
    if response.status_code == 200:
        # Analisa a resposta XML
        root = ET.fromstring(response.text)

        # Extrai informações relevantes
        events = []
        for event in root.findall('.//ShipmentEvent'):
            date = event.find('Date').text
            time = event.find('Time').text
            event_code = event.find('ServiceEvent/EventCode').text
            description = event.find('ServiceEvent/Description').text
            events.append([date, time, event_code, description])

        # Ordena os eventos por data e hora
        events_sorted = sorted(events, key=lambda x: (x[0], x[1]))

        # Retorna todo o histórico ou apenas o último evento, conforme a opção
        return events_sorted if historico else events_sorted[-1:]

    else:
        print(f'A requisição para a AWB {awb} falhou com o código de status {response.status_code}')
        print(f'Resposta XML:\n{response.text}')  # Adiciona esta linha para imprimir a resposta XML em caso de erro
        return None

def main(input_file, output_file, historico=False):
    site_id = input("Digite o Site id do cliente: ")
    password = input("Digite a senha: ")

    # Carrega o arquivo de entrada
    df = pd.read_excel(input_file)

    # Cria um novo arquivo Excel para armazenar os resultados
    wb = Workbook()
    ws = wb.active

    # Adiciona cabeçalhos às colunas no novo arquivo Excel
    headers = ['AWB', 'Date', 'Time', 'Event Code', 'Description']
    for col_num, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_num, value=header)

    # Itera sobre as linhas do DataFrame
    for index, row in df.iterrows():
        awb = str(row['AWB'])  # Converte para string, pois pode haver AWBs que são números inteiros
        events_sorted = consultar_awb(site_id, password, awb, historico)

        # Adiciona resultados ao novo arquivo Excel
        if events_sorted:
            if historico:
                # Salva todo o histórico
                for event in events_sorted:
                    ws.append([awb] + event)
            else:
                # Salva apenas o último evento
                ws.append([awb] + events_sorted[0] if events_sorted else [awb] + ['N/A', 'N/A', 'N/A', 'N/A'])

    # Salva o novo arquivo Excel
    wb.save(output_file)

if __name__ == "__main__":
    # Cria a janela de diálogo para seleção de arquivo
    root = Tk()
    root.withdraw()  # Oculta a janela principal

    print("Selecione o arquivo de entrada")

    # Obtém o caminho do arquivo de entrada
    input_file = filedialog.askopenfilename(title="Selecione o arquivo de entrada", filetypes=[("Arquivos Excel", "*.xlsx")])

    # Define se deve salvar o histórico completo ou apenas o último evento
    historico = input("Deseja obter o histórico completo? (S/N): ").upper() == 'S'

    # Define o caminho do arquivo de saída
    output_file = f'{formatted_date}_results{"_historico" if historico else ""}.xlsx'

    # Executa a função principal
    main(input_file, output_file, historico)