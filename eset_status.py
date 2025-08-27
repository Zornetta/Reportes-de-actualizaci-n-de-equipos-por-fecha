import csv
import os
from datetime import datetime, timedelta
from tkinter import Tk
from tkinter.filedialog import askopenfilenames
import json

class ESETStatusChecker:
    def __init__(self, config_file='config.json'):
        self.datos = []
        self.resultado = []
        self.current_directory = os.path.dirname(os.path.abspath(__file__))
        self.config_file = os.path.join(self.current_directory, config_file)
        self.tolerance_days = self.load_config()

    def load_config(self):
        with open(self.config_file, 'r') as file:
            config = json.load(file)
            return config.get('eset_tolerance_days', 5)  # Valor por defecto: 5 días si no está especificado

    def leer_csv(self, file_path):
        with open(file_path, mode='r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                self.datos.append({
                    'Nombre': row['Nombre'],
                    'Última conexión': row['Última conexión']
                })

    def procesar_datos(self):
        hoy = datetime.now()
        for item in self.datos:
            ultima_conexion = datetime.strptime(item['Última conexión'], '%d/%m/%Y %H:%M:%S')
            if (hoy - ultima_conexion).days <= self.tolerance_days:
                estado = 'OK'
            else:
                estado = 'Mal'
            self.resultado.append({
                'Nombre': item['Nombre'],
                'Estado': estado
            })

    def seleccionar_archivos(self):
        Tk().withdraw()  # Ocultar la ventana principal
        archivos = askopenfilenames(title="Selecciona las planillas CSV", filetypes=[("CSV files", "*.csv")])
        return archivos

    def get_status_table(self):
        return self.resultado

    def run(self):
        # Seleccionar archivos de entrada
        archivos = self.seleccionar_archivos()

        # Leer datos de los archivos CSV seleccionados
        for archivo in archivos:
            self.leer_csv(archivo)

        # Procesar datos
        self.procesar_datos()

        # Obtener la tabla de estado
        return self.get_status_table()

# Ejemplo de uso:
if __name__ == "__main__":
    processor = ESETStatusChecker()
    status_table = processor.run()
    for item in status_table:
        print(item)