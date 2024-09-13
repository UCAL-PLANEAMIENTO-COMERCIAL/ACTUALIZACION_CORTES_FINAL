from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import shutil
import time
from datetime import datetime,timedelta
import pytz
import pandas as pd
import numpy as np
import openpyxl
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive



# Set the timezone to Lima
lima_timezone = pytz.timezone('America/Lima')
lima_time = datetime.now(lima_timezone)
current_hour= lima_time.hour
if current_hour < 11:
    # If it's before 12 AM, set the date to yesterday
    today_date_lima = lima_time.date() - timedelta(days=1)
else:
    # If it's 12 AM or later, set the date to today
    today_date_lima = lima_time.date()        
    
today_string = today_date_lima.strftime('%d-%m-%Y')
    
def borrar_carpeta(folder):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

### Usuario
usuario = "MGARRIDOUCALPRE"
contrasena = "Admin@24"

print("Deleting old reports...")
borrar_carpeta('reportes_descarga\\')

######## Opciones para descarga de base
chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")  # Linux only

# Set a larger window size (replace 1920x1080 with your preferred size)
chrome_options.add_argument("--window-size=1920,1080")
current_dir = os.getcwd() + "\\reportes_descarga\\"
prefs = {"download.default_directory": current_dir}
chrome_options.add_experimental_option("prefs", prefs)
if not os.path.exists(current_dir):
    os.makedirs(current_dir)

# Path to the ChromeDriver executable
chrome_driver_path = 'C:\\Users\\TIP\\Desktop\\chromedriver-win64\\chromedriver.exe'
service = Service(chrome_driver_path)

### Abrir Prometei en chrome
print("Launching browser...")
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.get("https://prometeo.ieduca.pe/Login.aspx")

#### Inicia Sesion
print("Logging in...")
# Log in
usuario_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "txtUsuario"))
)
usuario_input.send_keys(usuario)
contrasena_input = driver.find_element(By.ID, "txtClave")
contrasena_input.send_keys(contrasena)
elemento_boton = driver.find_element(By.ID, "btnIngresar")
elemento_boton.click()

# Navigate to the report page
print("Navigating to report page...")
driver.execute_script("window.open('');")
driver.switch_to.window(driver.window_handles[1])
driver.get("https://prometeo.ieduca.pe/Reportes/ReporteSeguimiento.aspx")
time.sleep(2)  # Give the page some time to load

# Select options for the report
print("Selecting report options...")
tlmk = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_GvTipoAtencion_ctl22_ChkTipoAtencionG"))
)
tlmk.click()
driver.switch_to.default_content()
time.sleep(1)

wsp = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_GvTipoAtencion_ctl29_ChkTipoAtencionG"))
)
wsp.click()
driver.switch_to.default_content()
time.sleep(1)

traslado = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_GvTipoAtencion_ctl23_ChkTipoAtencionG"))
)
traslado.click()
driver.switch_to.default_content()
time.sleep(1)

inicio = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtFechaInicio"))
)
inicio.send_keys(today_string)
time.sleep(1)

fin = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtFechaFin"))
)
fin.send_keys(today_string)
time.sleep(1)

print("Generating report...")
generar = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnExportar"))
)

# Ejecutar el clic en el botón
driver.execute_script("arguments[0].click();", generar)

# Verificar si se hizo clic correctamente
print("Clicked on 'Generar reporte' button.")

# Esperar a que el archivo se descargue completamente
print("Waiting for file to download...")
time.sleep(10)  # Aumenta este tiempo si es necesario para archivos más grandes

# Obtener la lista de archivos descargados
files = os.listdir(current_dir)
if len(files) > 0:
    print(f"Downloaded file: {files[0]}")
else:
    print("No files downloaded.")

# Close the browser
print("Closing browser...")
driver.quit()
print("Automation completed successfully!")
 

download_folder = r"C:\Users\TIP\Documents\POYECTO_REPORT\reportes_descarga"

# Buscar el archivo Excel en la carpeta de descarga
def find_excel_file(folder):
    for filename in os.listdir(folder):
        if filename.endswith(".xls") or filename.endswith(".xlsx"):
            return os.path.join(folder, filename)
    return None

# Obtener la ruta del archivo Excel
path = find_excel_file(download_folder)

if path:
    # Cargar el archivo Excel encontrado
    data = pd.read_excel(path)
    dataCole = pd.read_excel('dataCole.xlsx')
    print("Archivo Excel cargado exitosamente!")
else:
    print("No se encontró ningún archivo Excel en la carpeta de descarga.")


datos = data[data['vendedor'] != 'TI Integrador']

# Separar campañas
#campania_24 = datos[(datos['campania'] != 'Admision 25-1') & (datos['vendedor'] != 'Stefano Napuri') & (datos['vendedor'] != 'Cinthia Orosco')]
campania_25 = datos[datos['campania'] == 'Admision 25-1']

data_cole = campania_25[campania_25['id_cliente'].isin(dataCole['sc_prospecto'])]
data_no_cole = campania_25[~campania_25['id_cliente'].isin(dataCole['sc_prospecto'])]

vendedores_unicos2 = ['Cinthia Orosco','Fiorella Lanegra','Angelica Iparraguirre','Andrea Araujo Antara','Maria Luque']

#vendedores_unicos = campania_24['vendedor'].unique()
#vendedores_unicos=[vendedor.title() for vendedor in vendedores_unicos]
"""
meta_values_campania24_2 = {
    'meta_gestiones': 144,
    'meta_contactos_unicos': 30,
    'meta_contactos_efectivos': 24,
    'meta_valoraciones_positivas': 8,
    'meta_promesas_de_pago': 1,
    'meta_ventas': 1
}
"""
meta_values_campania25_1 = {
    'meta_gestiones': 100,
    'meta_contactos_unicos': 30,
    'meta_contactos_efectivos': 24,
    'meta_valoraciones_positivas': 4,
    'meta_promesas_de_pago': 1,
    'meta_ventas': 1
}

# Función para procesar los datos y generar el reporte
def process_campaign_data2(df, meta_values):
    reporte2 = pd.DataFrame(columns=[
        'Asesor', 'COLE/NO COLE', 'Visita', 'llamada', 'Sin contacto', 'Perdidos',
        'Real Gestionados', 'Meta Gestionados', 
        'Real Contactados', 'Meta Contactados', 
        'Real Contactos Efectivos', 'Meta Contactos Efectivos', 
        'Real Valoraciones Positivas', 'Meta Valoraciones Positivas', 
        'Real Promesas de Pago', 'Meta Promesas de Pago', 
        'Real Ventas', 'Meta Ventas', 'COLE', 'Pool'
    ])
    
    for vendedor in vendedores_unicos2:
        for grupo, nombre_grupo in [(data_cole, 'COLE'), (data_no_cole, 'NO COLE')]:
            df_vendedor = grupo[grupo['vendedor'] == vendedor]
            
            llamada = df_vendedor.shape[0]
            sin_contacto = df_vendedor[df_vendedor['respuesta'] == 'Sin contacto']['id_cliente'].nunique()
            perdido = df_vendedor[df_vendedor['respuesta'] == 'Perdido'].shape[0]
            gestiones_por_vendedor = df_vendedor['id_cliente'].nunique()
            contactos_unicos = df_vendedor[df_vendedor['respuesta'] != 'Sin contacto']['id_cliente'].nunique()
            contactos_efectivos = df_vendedor[
                df_vendedor['respuesta_2_nivel'].isin([
                    'Programa visita a campus', 'Desea visita con el coordinador',
                    'Con fecha', 'Costo muy alto - hasta 600', 'Costo muy alto - hasta 800',
                    'Decide por otra institución - Otros', 'Decide por otra institución - Toulouse',
                    'Decide por otra institución - UPC', 'Motivos personales/laborales',
                    'Por distancia', 'Próxima campaña', 'Carrera de interés',
                    'Pendiente decisión de padres', 'Revisión de convalidacion',
                    'Revisión de escala de pago', 'Volver a llamar','Está ocupado - fecha'
                ])
            ]['id_cliente'].nunique()
            valoracion_positivas = df_vendedor[
                df_vendedor['respuesta_2_nivel'].isin([
                    'Programa visita a campus', 'Desea visita con el coordinador',
                    'Carrera de interés', 'Pendiente decisión de padres',
                    'Revisión de convalidacion', 'Revisión de escala de pago',
                    'Volver a llamar'
                ])
            ]['id_cliente'].nunique()
            promesa_pago = df_vendedor[df_vendedor['respuesta_2_nivel'] == 'Con fecha']['id_cliente'].nunique()
            ventas = df_vendedor[df_vendedor['respuesta_2_nivel'] == 'Venta']['id_cliente'].nunique()
            
            reporte2 = reporte2.append({
                'Asesor': vendedor,
                'COLE/NO COLE': nombre_grupo,
                'Visita': 0,
                'llamada': llamada,
                'Sin contacto': sin_contacto,
                'Perdidos': perdido,
                'Real Gestionados': gestiones_por_vendedor,
                'Meta Gestionados': meta_values['meta_gestiones'],
                'Real Contactados': contactos_unicos,
                'Meta Contactados': meta_values['meta_contactos_unicos'],
                'Real Contactos Efectivos': contactos_efectivos,
                'Meta Contactos Efectivos': meta_values['meta_contactos_efectivos'],
                'Real Valoraciones Positivas': valoracion_positivas,
                'Meta Valoraciones Positivas': meta_values['meta_valoraciones_positivas'],
                'Real Promesas de Pago': promesa_pago,
                'Meta Promesas de Pago': meta_values['meta_promesas_de_pago'],
                'Real Ventas': ventas,
                'Meta Ventas': meta_values['meta_ventas'],
                'COLE': nombre_grupo,
                'Pool': 'NO COLE' if nombre_grupo == 'COLE' else 'COLE'
            }, ignore_index=True)
    reporte2.loc[reporte2['Asesor'] == 'Maria Luque', 'Asesor'] = 'Maria Paz'
    reporte2.set_index(['Asesor', 'COLE/NO COLE'], inplace=True)
    reporte2.sort_values(by='Real Gestionados', ascending=False)
   

    return reporte2
reporte2 = process_campaign_data2(campania_25, meta_values_campania25_1)
#reporte2.reset_index(drop=True, inplace=True)
"""
def process_campaign_data(df, meta_values):
    llamada = df.groupby('vendedor').size().reindex(df['vendedor'].unique(), fill_value=0)
    Sin_Contacto = df[df['respuesta'] == 'Sin contacto'].groupby('vendedor')['id_cliente'].nunique().reindex(df['vendedor'].unique(), fill_value=0)
    perdido = df[df['respuesta'] == 'Perdido'].groupby('vendedor').size().reindex(df['vendedor'].unique(), fill_value=0)
    gestiones_por_vendedor = df.groupby('vendedor')['id_cliente'].nunique().reindex(df['vendedor'].unique(), fill_value=0)
    contactos_unicos = df[df['respuesta'] != 'Sin contacto'].groupby('vendedor')['id_cliente'].nunique().reindex(df['vendedor'].unique(), fill_value=0)
    contactos_efectivos = datos[
        datos['respuesta_2_nivel'].isin([
            'Programa visita a campus', 'Desea visita con el coordinador',
            'Con fecha', 'Costo muy alto - hasta 600', 'Costo muy alto - hasta 800',
            'Decide por otra institución - Otros', 'Decide por otra institución - Toulouse',
            'Decide por otra institución - UPC', 'Motivos personales/laborales',
            'Por distancia', 'Próxima campaña', 'Carrera de interés',
            'Pendiente decisión de padres', 'Revisión de convalidacion',
            'Revisión de escala de pago', 'Volver a llamar'
        ])
    ].groupby('vendedor')['id_cliente'].nunique().reindex(df['vendedor'].unique(), fill_value=0)
    valoracion_positivas = datos[
        datos['respuesta_2_nivel'].isin([
            'Programa visita a campus', 'Desea visita con el coordinador',
            'Carrera de interés', 'Pendiente decisión de padres',
            'Revisión de convalidacion', 'Revisión de escala de pago',
            'Volver a llamar'
        ])
    ].groupby('vendedor')['id_cliente'].nunique().reindex(df['vendedor'].unique(), fill_value=0)
    promesa_pago = datos[datos['respuesta_2_nivel'] == 'Con fecha'].groupby('vendedor')['id_cliente'].nunique().reindex(df['vendedor'].unique(), fill_value=0)
    ventas = datos[datos['vendedor'] == 'dsffs'].groupby('vendedor')['id_cliente'].nunique().reindex(df['vendedor'].unique(), fill_value=0)

    reporte = pd.DataFrame({
        'Asesor': vendedores_unicos,
        'Visita': [0] * len(vendedores_unicos),
        'llamada': llamada.values,
        'Sin contacto': Sin_Contacto.values,
        'Perdidos': perdido.values,
        'Real Gestionados': gestiones_por_vendedor.values,
        'Meta Gestionados': [meta_values['meta_gestiones']] * len(vendedores_unicos),
        'Real Contactados': contactos_unicos.values,
        'Meta Contactados': [meta_values['meta_contactos_unicos']] * len(vendedores_unicos),
        'Real Contactos Efectivos': contactos_efectivos.values,
        'Meta Contactos Efectivos': [meta_values['meta_contactos_efectivos']] * len(vendedores_unicos),
        'Real Valoraciones Positivas': valoracion_positivas.values,
        'Meta Valoraciones Positivas': [meta_values['meta_valoraciones_positivas']] * len(vendedores_unicos),
        'Real Promesas de Pago': promesa_pago.values,
        'Meta Promesas de Pago': [meta_values['meta_promesas_de_pago']] * len(vendedores_unicos),
        'Real Ventas': ventas.values,
        'Meta Ventas': [meta_values['meta_ventas']] * len(vendedores_unicos),
    }).sort_values(by='Real Gestionados', ascending=False)
    
    reporte.loc[reporte['Asesor'] == 'Ingrid Guillermo Rivera', 'Meta Gestionados'] = 100
    
    
    reporte = pd.concat([reporte], ignore_index=True)
    
    return reporte

reporte = process_campaign_data(campania_24, meta_values_campania24_2)
""" 

print("Generando el archivo Excel...")
# Define the path for the final Excel report
output_file = "Reporte_Final.xlsx"
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    #reporte.to_excel(writer, sheet_name="Campania_24.2", index=False)
    reporte2.to_excel(writer, sheet_name="Campania_25.1")

print("¡Reporte final generado exitosamente!")


"""
folder_id="1sUf9tDLftR3UaJaZsb2TtCpvH9nHZ0zf"

def delete_existing_file(drive, file_name):
    # List all files in the drive with the given name
    file_list = drive.ListFile({'q': f"title='{file_name}' and trashed=false"}).GetList()
    for file in file_list:
        if file['title'] == file_name:
            print(f"Deleting existing file: {file['title']} (ID: {file['id']})")
            file.Delete()
            print("File deleted successfully.")
            return
    print("No existing file found with the given name.")
    
    
    
def upload_to_drive(file_path, folder_id):
    # Autenticación y creación del cliente de Google Drive
    gauth = GoogleAuth()
    gauth.LoadCredentialsFile("mycreds.txt")

    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved creds
        gauth.Authorize()

    # Guardar las credenciales para la próxima vez
    gauth.SaveCredentialsFile("mycreds.txt")
    
    drive = GoogleDrive(gauth)
    
    delete_existing_file(drive, os.path.basename(file_path))

    # Crear y cargar el archivo en la carpeta especificada
    file_name = os.path.basename(file_path)
    file = drive.CreateFile({'title': file_name, 'parents': [{'id': folder_id}]})
    file.SetContentFile(file_path)
    file.Upload()
    


# Upload to Google Drive
upload_to_drive(output_file,folder_id)
"""







