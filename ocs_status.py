import csv
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

class OCSStatusChecker:
    def __init__(self):
        self.datos = []

    def leer_csv(self, file_path):
        with open(file_path, mode='r', encoding='utf-8') as file:
            reader = csv.DictReader(file, delimiter=';')  # Usar punto y coma como delimitador
            # Verificar las columnas disponibles en el archivo CSV
            columns = reader.fieldnames
            if 'Equipo' not in columns:
                raise KeyError(f"La columna 'Equipo' no se encuentra en el archivo CSV. Columnas disponibles: {columns}")
            for row in reader:
                self.datos.append({
                    'Equipo': row['Equipo']
                })

    def seleccionar_archivo(self):
        Tk().withdraw()  # Ocultar la ventana principal
        archivo = askopenfilename(title="Selecciona la planilla CSV de OCS", filetypes=[("CSV files", "*.csv")])
        return archivo

    def get_status_table(self):
        return self.datos

    def run(self):
        # Seleccionar archivo de entrada
        archivo = self.seleccionar_archivo()

        # Leer datos del archivo CSV seleccionado
        if archivo:
            self.leer_csv(archivo)

        # Obtener la tabla de estado
        return self.get_status_table()

# Example usage:
if __name__ == "__main__":
    checker = OCSStatusChecker()
    status_table = checker.run()
    for item in status_table:
        print(item)