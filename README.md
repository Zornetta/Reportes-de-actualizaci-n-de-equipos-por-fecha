
# Reportes de actualización de equipos por fecha

**Autor:** Anzor

Este proyecto permite generar un reporte consolidado del estado de actualización de equipos en una red, combinando información de tres fuentes principales: ESET (antivirus), WSUS (actualizaciones de Windows) y OCS Inventory (inventario de equipos).

## Objetivo

El objetivo es facilitar la supervisión del estado de los equipos, mostrando en un solo reporte si cada máquina está actualizada según los criterios de ESET, WSUS y si está presente en el inventario de OCS. El reporte final se genera en formato Excel, con colores que facilitan la visualización del estado de cada equipo.

## ¿Cómo funciona?

1. **Obtención de datos**:
   - **ESET**: Selecciona uno o varios archivos CSV exportados desde la consola de ESET. El script verifica la última conexión de cada equipo y determina si está dentro del rango de tolerancia configurado en `config.json`.
   - **WSUS**: El script ejecuta automáticamente `export-wsus-data.ps1` para exportar los datos de WSUS a un archivo CSV. Compara la fecha del último reporte de cada equipo con la tolerancia configurada en `config.json`.
   - **OCS**: Selecciona el archivo CSV exportado desde OCS Inventory, que contiene la lista de equipos detectados.

2. **Procesamiento y consolidación**:
   - Los datos de los tres sistemas se normalizan y se combinan en una sola tabla.
   - Se ignoran los equipos que aparecen en una lista negra (blacklist).
   - Para cada equipo, se muestra el estado en ESET, WSUS y si está presente en OCS.

3. **Generación de reporte**:
   - Se crea un archivo Excel (`status_report.xlsx`) con el estado de cada equipo.
   - Los estados se colorean: verde para "OK", rojo para "Mal" y azul para "N/A".

## Dependencias

- Python 3.x
- Paquetes:
  - `openpyxl`
  - `tkinter` (incluido en la mayoría de instalaciones de Python)
- Acceso a los archivos CSV exportados desde ESET, WSUS y OCS Inventory.
- PowerShell (para ejecutar el script de WSUS)

## Archivos principales

- `index.py`: Script principal que coordina la obtención, procesamiento y reporte de los datos.
- `eset_status.py`: Clase para procesar los datos exportados de ESET.
- `wsus_status.py`: Clase para procesar los datos exportados de WSUS.
- `ocs_status.py`: Clase para procesar los datos exportados de OCS Inventory.
- `config.json`: Configuración de tolerancia de días para ESET y WSUS.
- `export-wsus-data.ps1`: Script de PowerShell para exportar datos de WSUS.
- `status_report.xlsx`: Archivo generado con el reporte consolidado.

## Uso detallado

1. **Ejecuta el script principal**  
   Abre una terminal en la carpeta del proyecto y ejecuta:
   ```
   python index.py
   ```

2. **Selecciona los archivos de entrada**  
   El programa te pedirá seleccionar los archivos CSV necesarios:
   - Primero, selecciona uno o varios archivos CSV exportados desde la consola de ESET.  
   - Luego, selecciona el archivo CSV exportado desde OCS Inventory.

3. **Exportación de datos de WSUS**  
   El script ejecutará automáticamente el archivo `export-wsus-data.ps1` para generar el archivo `WSUSData.csv` con la información de WSUS. No necesitas hacerlo manualmente.

4. **Procesamiento de datos**  
   El programa:
   - Lee y procesa los datos de ESET, WSUS y OCS.
   - Normaliza los nombres de los equipos para evitar duplicados.
   - Compara las fechas de última conexión/reporte con la tolerancia configurada en `config.json`.
   - Ignora los equipos que están en la lista negra (blacklist).

5. **Generación del reporte**  
   Se crea el archivo `status_report.xlsx` con el estado de cada equipo:
   - Verde ("OK") si el equipo está actualizado.
   - Rojo ("Mal") si el equipo no está actualizado.
   - Azul ("N/A") si no hay información disponible.

6. **Resultado final**  
   Al finalizar, verás un mensaje indicando la ubicación del archivo generado.  
   Abre `status_report.xlsx` para consultar el reporte consolidado.


## Notas
Si necesitas agregar nuevas fuentes de datos o modificar la lógica de tolerancia, edita el archivo `config.json` y/o los scripts correspondientes.
