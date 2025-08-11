import win32com.client
import datetime
import pandas as pd
import re
import os

# === CONFIGURACIÓN ===
excel_path = r"Z:\EOLICA\MAT REF\BASE DE DATOS ESTACIONES\CorreosDatosEstaciones\control_mails.xlsx"
#fecha_objetivo = datetime.date.today()  # Fecha de verificación (hoy)
fecha_objetivo = datetime.date(2025, 8, 5)

# === LEER EXCEL ===
df = pd.read_excel(excel_path)
sistemas = df.iloc[:, 0].astype(str).tolist()  # Primera columna: sistemas
print(sistemas)

# === CONECTAR A OUTLOOK ===
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
bandeja_entrada = outlook.GetDefaultFolder(6)  # 6 = Bandeja de entrada

mensajes = bandeja_entrada.Items
mensajes.Sort("[ReceivedTime]", True)  # Ordenar por fecha descendente

# === RANGO DE FECHA ===
inicio_dia = datetime.datetime.combine(fecha_objetivo, datetime.time.min)
fin_dia = datetime.datetime.combine(fecha_objetivo, datetime.time.max)

# === BUSCAR CORREOS ===
correos_recibidos = set()

for mensaje in mensajes:
    try:
        fecha_msg = mensaje.ReceivedTime.replace(tzinfo=None)
        if inicio_dia <= fecha_msg <= fin_dia:
            match = re.match(r"LIDAR_(\w+)_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}", mensaje.Subject)
            if match:
                sistema_id = match.group(1)
                correos_recibidos.add(sistema_id)
    except AttributeError:
        # Evita errores si no es un correo (puede haber elementos de calendario, tareas, etc.)
        continue

# === NOMBRE DE COLUMNA ===
columna_fecha = fecha_objetivo.strftime("%Y-%m-%d")

# === AÑADIR O ACTUALIZAR COLUMNA ===
df[columna_fecha] = [1 if s in correos_recibidos else 0 for s in sistemas]

# === GUARDAR SOBRE EL MISMO EXCEL ===
backup_path = excel_path.replace(".xlsx", "_backup.xlsx")
os.replace(excel_path, backup_path)  # Guardar copia de seguridad ; Hace falta que el archivo no esté abierto
df.to_excel(excel_path, index=False)

print(f"Proceso completado. Columna '{columna_fecha}' actualizada en {excel_path}")
print(f"Se creó un backup en: {backup_path}")

