import os
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from tkinter import Tk, filedialog
from dotenv import load_dotenv

class DHLRequestBuilder:
    def __init__(self, site_id, password, awb):
        self.site_id = site_id
        self.password = password
        self.awb = awb
        
    def build_xml(self):
        xml_template = """<?xml version="1.0" encoding="UTF-8"?>
            <req:KnownTrackingRequest xmlns:req="http://www.dhl.com" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.dhl.com TrackingRequestKnown.xsd" schemaVersion="1.0">
                <Request>
                    <ServiceHeader>
                        <SiteID>{}</SiteID>
                        <Password>{}</Password>
                    </ServiceHeader>
                </Request>
                <LanguageCode>BR</LanguageCode>
                <AWBNumber>{}</AWBNumber>
                <LevelOfDetails>ALL_CHECK_POINTS</LevelOfDetails>
            </req:KnownTrackingRequest>"""
        return xml_template.format(self.site_id, self.password, self.awb)

class DHLTracker:
    def __init__(self, site_id, password):
        self.site_id = site_id
        self.password = password

    def get_awb(self, awb, exists_history=False):
        try:
            load_dotenv()
            url = os.getenv("URL")
            headers_str = os.getenv("HEADERS")

            # Convertendo a string JSON em um dicionário
            headers = eval(headers_str)

            request_builder = DHLRequestBuilder(self.site_id, self.password, awb)

            data = request_builder.build_xml()

            response = requests.post(url, headers=headers, data=data)
            response.raise_for_status()  # Adiciona esta linha para verificar se houve um erro HTTP

            if response.status_code == 200:
                root = ET.fromstring(response.text)
                events = self.extract_events(root)
                events_sorted = sorted(events, key=lambda x: (x[0], x[1]))
                return events_sorted if exists_history else events_sorted[-1:]
            else:
                print(f'A requisição para a AWB {awb} falhou com o código de status {response.status_code}')
                print(f'Resposta XML:\n{response.text}')
                return None
        except requests.exceptions.RequestException as e:
            print(f'Erro na requisição para a AWB {awb}: {e}')
            return None

    def extract_events(self, root):
        events = []
        for event in root.findall('.//ShipmentEvent'):
            date = event.find('Date').text
            time = event.find('Time').text
            event_code = event.find('ServiceEvent/EventCode').text
            description = event.find('ServiceEvent/Description').text
            events.append([date, time, event_code, description])
        return events        

class Tracking:
    def __init__(self):
        self.now = datetime.now()
        self.formatted_date = self.now.strftime("%Y-%m-%d-%H-%M")
        
    def get_tracking_and_generate_report(self, input_file, output_file, exists_history=False):
        try:
            site_id = input("Digite o Site id do cliente: ")
            password = input("Digite a senha: ")

            tracker = DHLTracker(site_id, password)

            dataframe = pd.read_excel(input_file)
            workbook = Workbook()
            worksheet = workbook.active

            headers_file = ['AWB', 'Date', 'Time', 'Event Code', 'Description']
            for column_number, headers_file in enumerate(headers_file, 1):
                worksheet.cell(row=1, column=column_number, value=headers_file)

            for index, row in dataframe.iterrows():
                awb = str(row["AWB"])
                events_sorted = tracker.get_awb(awb, exists_history)

                # Correção aqui: Alterando a condição para verificar se há eventos
                if events_sorted:
                    if exists_history:
                        for event in events_sorted:
                            worksheet.append([awb] + event)
                    else:
                        worksheet.append([awb] + events_sorted[0] if events_sorted else [awb] + ['N/A', 'N/A', 'N/A', 'N/A'])
            workbook.save(output_file)
        except Exception as e:
            print(f'Erro durante a execução: {e}')

if __name__ == "__main__":
    try:
        root = Tk()
        root.withdraw()

        print("Selecione o arquivo de entrada")

        input_file = filedialog.askopenfilename(title="Selecione o arquivo de entrada", filetypes=[("Arquivos Excel", "*.xlsx")])

        history = input("Deseja obter o histórico completo? (S/N): ").upper() == 'S'

        app = Tracking()

        output_file = f'{app.formatted_date}_results{"_historico" if history else ""}.xlsx'

        app.get_tracking_and_generate_report(input_file, output_file, history)
    except Exception as e:
        print(f'Erro durante a execução do programa: {e}')
