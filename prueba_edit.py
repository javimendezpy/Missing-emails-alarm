''' 
Contexto: Cada día se envian correos a energias.renovables.es@dekra.com en los cuales se adjuntan los 
datos medidos por cada estación (LiDAR, SoDAR o TM). Cuando alguno de los sistemas falla no se envía el 
correo correspondiente. No hay forma de ver de forma directa si faltan correos y por tanto alguna de los 
sistemas está fallando. Se quiere automatizar un proceso por el cual se identifique si faltan correos o 
no y guardar en una hoja de cálculo un 1/0 por sistema y día además de un identificador del remitente.

El objetivo es tener un histórico de si se han recibido correos o no para cada sistema y día (filas=sistemas,
columnas=días)

Ejemplo, recibimos un correo de windcubeinsights@vaisala.info para la estación de Potrillo con fecha 2025-08-11
debemos guardar en la hoja de cálculo "1(windcube)" en la fila de Potrillo y en la columna del día 2025-08-11.
'''


import pandas as pd 
import re 
''' Regular expression operations (regex) -> Reconocer patrones en strings
- re.match(patron, texto) : Comprueba si el texto empieza con el patrón
- re.search(patron, texto) : Busca el patrón en cualquier parte del texto
- re.findall(patron, texto) : Devuelve todas las ocurrencias del patrón en el texto y las pone en una lista
- re.sub(patron, reemplazo, texto) : Reemplaza las ocurrencias del patrón en el texto
- re.compile(patron) : Compila el patrón para reutilizarlo varias veces de forma más eficiente
'''
import win32com.client

df = pd.read_excel('sistemas_id_asuntos.xlsx')
sistemas = df.iloc[:, 0].tolist()
#print("Sistemas: ", sistemas)
#print(df)

''' 
Asuntos de ejemplo y formato de fechas para cada remitente
estacionesmeteo -> LIDAR Punago-9_2025-08-12_00-10-00 ; YYYY-MM-DD_HH-MM-SS
windcube -> WindCube Insights Fleet: New STA File from WLS71497  2025/07/31  00:10:00 ; YYYY/MM/DD  HH-MM-SS
emailrelay -> LIDAR Villalube-6A_2025-08-11_00-10-00 ; YYYY-MM-DD_HH-MM-SS
zx -> Daily Data: Wind10_1148@Y2025_M08_D02.CSV (Averaged data) ; YYYYY_MMMM_DDD
molas -> Data of Molas B300-2150——2025/05/28 ; YYYY/MM/DD
'''

# Conectar a Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#print("\n📬 Cuentas de Outlook detectadas:")
#for store in outlook.Stores:
    #print(" -", store.DisplayName)


nombre_cuenta = "energias.renovables.es@dekra.com"
store = outlook.Stores[nombre_cuenta]
bandeja_entrada = store.GetDefaultFolder(6)
mensajes = bandeja_entrada.Items
mensajes.Sort("[ReceivedTime]", True)

fecha_actual = pd.Timestamp.now().normalize() # Formato 
fecha_actual_ini = fecha_actual.date()  # Fecha actual a las 00:00:00
fecha_actual_fin = (fecha_actual + pd.Timedelta(days=1)).date()  # Un día después

print("Fecha inicial: ", fecha_actual_ini)
print("Fecha final: ", fecha_actual_fin)

for sistema in sistemas:

    print(sistema)

    id = df[df.iloc[:, 0] == sistema].iloc[0, 1]
    remitente = df[df.iloc[:, 0] == sistema].iloc[0, 2]

    if remitente == 'estaciones.meteo@dekra-industrial.es':
        patron_asunto = re.compile(rf"^LIDAR {id}_(\d{{4}}-\d{{2}}-\d{{2}})_(\d{{2}}-\d{{2}}-\d{{2}})$")
        rem = 'estacionesmeteo'
    elif remitente == 'windcubeinsights@vaisala.info':
        patron_asunto = re.compile(rf"`WindCube Insights Fleet: New STA File from {id}  (\d{{4}}/\d{{2}}/\d{{2}})  (\d{{2}}-\d{{2}}-\d{{2}})$")
        rem = 'windcube'
    elif remitente == 'emailrelay@konectgds.com':
        patron_asunto = re.compile(rf"^LIDAR {id}_(\d{{4}}-\d{{2}}-\d{{2}})_(\d{{2}}-\d{{2}}-\d{{2}})$")
        rem = 'emailrelay'
    elif remitente == 'status@support.zxlidars.com':
        patron_asunto = re.compile(rf"^Daily Data: Wind10_{id}@Y(\d{{4}})_M(\d{{2}})_D(\d{{2}})\.CSV \(Averaged data\)$")
        rem = 'zx'
    elif remitente == 'molas-b300@wind.molascloud.com':
        patron_asunto = re.compile(rf"^Data of {id} ——(\d{{4}}/\d{{2}}/\d{{2}})$")
        rem = 'molas'

    print(patron_asunto)

    # Identificar si se ha recibido el correo
    #print(mensajes_filtrados)

    def contiene_elementos(asunto: str) -> str:
        if patron_asunto.match(asunto):  # .match = desde el inicio de la cadena
            return "sí"
        else:
            return "no"
        
    for msg in mensajes:
        print(msg.Subject)
        print(pd.Timestamp(msg.ReceivedTime).date())

        if fecha_actual_ini <= pd.Timestamp(msg.ReceivedTime).date() <= fecha_actual_fin:
            if msg.SenderEmailAddress == remitente:
                if contiene_elementos(msg.Subject) == "sí":
                    out = f"1 ({rem})"
                else:
                    out = f"0 ({rem})"

    print(out)
