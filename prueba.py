import win32com.client
import datetime
import pandas as pd
import re
import os

# === CONFIGURACIÓN ===
excel_path = r"Z:\EOLICA\MAT REF\BASE DE DATOS ESTACIONES\CorreosDatosEstaciones\control_mails.xlsx"
fecha_objetivo = datetime.date.today()  # Fecha de verificación = hoy
nombre_cuenta = "energias.renovables.es@dekra.com"  # Nombre exacto del buzón

# === LEER EXCEL ===
df = pd.read_excel(excel_path)
sistemas = df.iloc[:, 0].astype(str).tolist()
print("📄 Sistemas en Excel:", sistemas[:10], "...")

# === CONECTAR A OUTLOOK ===
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Mostrar todas las cuentas disponibles
print("\n📬 Cuentas de Outlook detectadas:")
for store in outlook.Stores:
    print(" -", store.DisplayName)

# Seleccionar la cuenta correcta
try:
    store = outlook.Stores[nombre_cuenta]
except Exception:
    raise ValueError(f"No se encontró la cuenta '{nombre_cuenta}'. Verifica el nombre en la lista anterior.")

bandeja_entrada = store.GetDefaultFolder(6)  # 6 = Inbox
mensajes = bandeja_entrada.Items
mensajes.Sort("[ReceivedTime]", True)  # Más recientes primero

# === PATRÓN DEL ASUNTO ===
patron_asunto = re.compile(r"LIDAR\s(.+?)_(\d{4}-\d{2}-\d{2})_\d{2}-\d{2}-\d{2}")

correos_recibidos = set()
contador = 0

print(f"\n🔍 Buscando correos con fecha en asunto = {fecha_objetivo}\n")

for mensaje in mensajes:
    try:
        asunto = mensaje.Subject
        fecha_recepcion = mensaje.ReceivedTime
        match = patron_asunto.match(asunto)

        if match:
            sistema_id = match.group(1).strip()
            fecha_asunto = datetime.datetime.strptime(match.group(2), "%Y-%m-%d").date()

            if fecha_asunto == fecha_objetivo:
                print(f"✅ {fecha_recepcion} | {asunto} → Coincide (Sistema: {sistema_id})")
                correos_recibidos.add(sistema_id)
            else:
                print(f"❌ {fecha_recepcion} | {asunto} → Fecha asunto {fecha_asunto} no coincide")
        else:
            print(f"⚠️ {fecha_recepcion} | {asunto} → No coincide con patrón LIDAR")

        contador += 1
        if contador >= 500:  # Limita la búsqueda
            break

    except AttributeError:
        continue

print(f"\n📧 Sistemas con correo en {fecha_objetivo}: {correos_recibidos}")

# === AÑADIR COLUMNA ===
columna_fecha = fecha_objetivo.strftime("%Y-%m-%d")
df[columna_fecha] = [1 if s in correos_recibidos else 0 for s in sistemas]

# === GUARDAR CON BACKUP ===
#backup_path = excel_path.replace(".xlsx", "_backup.xlsx")
#if os.path.exists(excel_path):
#    os.replace(excel_path, backup_path)
#df.to_excel(excel_path, index=False)

print(f"\n✅ Columna '{columna_fecha}' actualizada en {excel_path}")
#print(f"💾 Backup creado en: {backup_path}")


