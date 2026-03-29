import pandas as pd
import os
from tkinter import Tk, filedialog, messagebox

# Ocultar ventana principal
Tk().withdraw()

# Seleccionar carpeta
input_folder = filedialog.askdirectory(title="Selecciona la carpeta con los reportes")

if not input_folder:
    print("❌ No seleccionaste ninguna carpeta")
    exit()

# Crear carpeta /fix
output_folder = os.path.join(input_folder, "fix")
os.makedirs(output_folder, exist_ok=True)

columnas_estandar = [
    "Estado", "Entrega de Carga ", "Entregado a", "Modelo",
    "Descripción", "Numero de Parte", "Serial", "Piezas", "Fecha de Salida"
]

procesados = 0
errores = 0

for archivo in os.listdir(input_folder):
    if archivo.endswith(".xls") or archivo.endswith(".xlsx"):
        ruta = os.path.join(input_folder, archivo)

        try:
            df = pd.read_excel(ruta, header=None)
            df = df.iloc[1:].reset_index(drop=True)
            df.columns = columnas_estandar[:len(df.columns)]

            # 1. Convertir a datetime (esto es necesario para que Excel lo reconozca como fecha)
            df["Fecha de Salida"] = pd.to_datetime(df["Fecha de Salida"], errors="coerce")

            for col in columnas_estandar:
                if col not in df.columns:
                    df[col] = None
            df = df[columnas_estandar]

            # Definir ruta de salida antes de escribir
            nombre_salida = archivo.replace(".xls", ".xlsx") if archivo.endswith(".xls") else archivo
            salida = os.path.join(output_folder, nombre_salida)

            # 2. Guardar forzando el formato de celda de Excel
            with pd.ExcelWriter(salida, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Reporte')
                
                # Acceder a la hoja para dar formato específico a la columna
                workbook = writer.book
                worksheet = writer.sheets['Reporte']
                
                # El formato 'dd/mm/yyyy' obliga a Excel a mostrarlo como Fecha Corta
                # La columna "Fecha de Salida" es la 9na (índice I en Excel)
                for cell in worksheet['I']: 
                    cell.number_format = 'MM/DD/YYYY'

            print(f"✔ Procesado: {archivo}")
            procesados += 1

        except Exception as e:
            print(f"❌ Error en {archivo}: {e}")
            errores += 1

print("\n✅ Proceso completado")

# Notificación final
root = Tk()
root.withdraw()
messagebox.showinfo("Proceso finalizado", f"Procesados: {procesados}\nErrores: {errores}")
root.destroy()