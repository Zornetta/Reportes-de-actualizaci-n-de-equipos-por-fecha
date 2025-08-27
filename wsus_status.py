import subprocess
import csv
from datetime import datetime, timedelta
import os
import json

class WSUSStatusChecker:
    def __init__(self, ps_script_name='export-wsus-data.ps1', config_file='config.json'):
        self.current_directory = os.path.dirname(os.path.abspath(__file__))
        self.ps_script_path = os.path.join(self.current_directory, ps_script_name)
        self.csv_path = os.path.join(self.current_directory, 'WSUSData.csv')
        self.config_file = os.path.join(self.current_directory, config_file)
        self.tolerance_days = self.load_config()

    def load_config(self):
        with open(self.config_file, 'r') as file:
            config = json.load(file)
            return config.get('wsus_tolerance_days', 7)  # Default to 7 days if not specified

    def run_ps_script(self):
        subprocess.run(["powershell.exe", "-File", self.ps_script_path], check=True)

    def read_csv(self):
        wsus_data = []
        with open(self.csv_path, mode='r') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                wsus_data.append(row)
        return wsus_data

    def generate_status_table(self, wsus_data):
        status_table = []
        for entry in wsus_data:
            computer_name = entry['Equipo']
            last_report_date = datetime.strptime(entry['FechaUltimoReporte'], '%Y-%m-%d %H:%M:%S')
            current_date = datetime.now()
            if current_date - last_report_date > timedelta(days=self.tolerance_days):
                status = "Mal"
            else:
                status = "OK"
            status_table.append({'Nombre': computer_name, 'Estado': status})
        return status_table

    def get_status_table(self):
        self.run_ps_script()
        wsus_data = self.read_csv()
        status_table = self.generate_status_table(wsus_data)
        os.remove(self.csv_path)  # Eliminar el archivo CSV después de procesarlo
        return status_table

# Example usage:
if __name__ == "__main__":
    checker = WSUSStatusChecker()
    status_table = checker.get_status_table()
    for item in status_table:
        print(item)