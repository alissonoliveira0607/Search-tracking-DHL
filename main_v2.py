import os
import requests
import json
import xml.etree.ElementTree as ET
import pandas as pd
from openpyxl import Workbook
from datetime import datetime
from tkinter import Tk, filedialog
from dotenv import load_dotenv
import logging

# Configurando o sistema de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DHLRequestTracking:
    def get_tracking_request(self, site_id, password, awb):
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
        return xml_template.format(site_id, password, awb)

class DHLTracker:
    def get_awb(self, site_id, password, awb, exists_history=False):
        try:
            load_dotenv()
            url = os.getenv("URL")
            headers_env = os.getenv("HEADERS")

            # Convertendo a string JSON em um dicionário
            headers = json.loads(headers_env)

            request_builder = DHLRequestTracking()

            data = request_builder.get_tracking_request(site_id, password, awb)

            response = requests.post(url, headers=headers, data=data)
            response.raise_for_status()

            if response.status_code == 200:
                root = ET.fromstring(response.text)
                events = self.extract_events(root)
                events_sorted = sorted(events, key=lambda event: (event['date'], event['time']))
                return events_sorted if exists_history else events_sorted[-1:]
            else:
                logger.error(f'A requisição para a AWB {awb} falhou com o código de status {response.status_code}')
                logger.error(f'Resposta XML:\n{response.text}')
        except requests.exceptions.RequestException as e:
            logger.error(f'Erro na requisição para a AWB {awb}: {e}')

    def extract_events(self, root):
        events = []
        for event in root.findall('.//ShipmentEvent'):
            date = event.find('Date').text
            time = event.find('Time').text
            event_code = event.find('ServiceEvent/EventCode').text
            description = event.find('ServiceEvent/Description').text
            events.append({'date': date, 'time': time, 'event_code': event_code, 'description': description})
        return events

class Tracking:
    def __init__(self):
        self.now = datetime.now()
        self.formatted_date = self.now.strftime("%Y-%m-%d-%H-%M")

    def generate_report(self, input_file, output_file, exists_history=False):
        try:
            site_id = input("Digite o Site id do cliente: ")
            password = input("Digite a senha: ")

            tracker = DHLTracker()

            dataframe = pd.read_excel(input_file)
            workbook = Workbook()
            worksheet = workbook.active

            headers_file = ['AWB', 'Date', 'Time', 'Event Code', 'Description']
            for column_number, header in enumerate(headers_file, 1):
                worksheet.cell(row=1, column=column_number, value=header)

            for index, row in dataframe.iterrows():
                awb = str(row["AWB"])
                events_sorted = tracker.get_awb(site_id, password, awb, exists_history)

                if events_sorted:
                    if exists_history:
                        for event in events_sorted:
                            worksheet.append([awb, event['date'], event['time'], event['event_code'], event['description']])
                    else:
                        worksheet.append([awb, events_sorted[0]['date'], events_sorted[0]['time'], events_sorted[0]['event_code'], events_sorted[0]['description']] if events_sorted else [awb, 'N/A', 'N/A', 'N/A', 'N/A'])
            workbook.save(output_file)
        except Exception as e:
            logger.error(f'Erro durante a execução: {e}')

if __name__ == "__main__":
    try:
        root = Tk()
        root.withdraw()

        print("Selecione o arquivo de entrada")

        input_file = filedialog.askopenfilename(title="Selecione o arquivo de entrada", filetypes=[("Arquivos Excel", "*.xlsx")])

        history = input("Deseja obter o histórico completo? (S/N): ").upper() == 'S'

        app = Tracking()

        output_file = f'{app.formatted_date}_results{"_historico" if history else ""}.xlsx'

        app.generate_report(input_file, output_file, history)
    except Exception as e:
        logger.error(f'Erro durante a execução do programa: {e}')
