import os
import shutil
import glob
import pyautogui
import time
import logging
import subprocess
import flet as ft
from dotenv import load_dotenv

# --- Cargar variables de entorno desde .env ---
load_dotenv()

user_profile = os.path.expanduser("~")

ONEDRIVE_PATH = os.path.join(user_profile, os.getenv("ONEDRIVE_PATH", "OneDrive"))
ONEDRIVE_BACKUP_FOLDER = os.getenv("ONEDRIVE_BACKUP_FOLDER", "backup_correos")
MEGA_PATH = os.path.join(user_profile, os.getenv("MEGA_PATH", "MEGA"))
SCREENSHOTS_PREFIX = os.getenv("SCREENSHOTS_PREFIX", "")
USER_FOLDERS = [c.strip() for c in os.getenv("USER_FOLDERS", "Descargas,Documentos,Imágenes,Música,Videos").split(",")]
BACKUPS_FOLDER = os.path.join(os.getcwd(), os.getenv("BACKUPS_FOLDER", "backups"))
try:
    SCREENSHOT_WAIT = int(os.getenv("SCREENSHOT_WAIT", "2"))
except:
    SCREENSHOT_WAIT = 2

logging.basicConfig(
    filename='soporte_tool.log',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s:%(message)s'
)
if not os.path.exists(BACKUPS_FOLDER):
    os.makedirs(BACKUPS_FOLDER)

def get_onedrive_corporate_folders():
    """
    Retorna una lista de rutas de OneDrive corporativo configuradas para el usuario.
    Ignora la carpeta OneDrive personal.
    """
    onedrives = []
    for item in os.listdir(user_profile):
        if item.startswith("OneDrive -"):
            odrive_path = os.path.join(user_profile, item)
            if os.path.isdir(odrive_path):
                onedrives.append(odrive_path)
    return onedrives

def backup_outlook(page, log_text):
    """
    Realiza una copia de seguridad de los archivos PST/OST de Outlook,
    guardándolos en la cuenta principal de OneDrive definida en .env.

    Si necesitas cambiar la ruta de backup, ajusta ONEDRIVE_PATH y ONEDRIVE_BACKUP_FOLDER en .env.
    """
    try:
        log_text.value += "Iniciando backup de Outlook...\n"
        page.update()
        rutas_busqueda = [
            os.path.join(user_profile, 'Documents', 'Outlook Files'),
            os.path.join(user_profile, 'AppData', 'Local', 'Microsoft', 'Outlook'),
        ]
        archivos_encontrados = []
        for ruta in rutas_busqueda:
            if os.path.exists(ruta):
                archivos_encontrados += glob.glob(os.path.join(ruta, '*.pst'))
                archivos_encontrados += glob.glob(os.path.join(ruta, '*.ost'))
        onedrive_folder = os.path.join(ONEDRIVE_PATH, ONEDRIVE_BACKUP_FOLDER)
        if not os.path.exists(onedrive_folder):
            os.makedirs(onedrive_folder)
        for archivo in archivos_encontrados:
            shutil.copy2(archivo, onedrive_folder)
            logging.info(f'Archivo respaldado: {archivo} → {onedrive_folder}')
            log_text.value += f"Archivo respaldado: {os.path.basename(archivo)}\n"
            page.update()
        if archivos_encontrados:
            log_text.value += "Backup completado en OneDrive.\n"
        else:
            log_text.value += "No se encontraron archivos de Outlook.\n"
        page.update()
    except Exception as e:
        logging.error(f'Error en backup Outlook: {str(e)}')
        log_text.value += f"Error en backup Outlook: {e}\n"
        page.update()

def screenshot_folder(ruta, nombre, log_text, page):
    """
    Toma un screenshot de la carpeta especificada y lo guarda en BACKUPS_FOLDER.
    Usa el prefijo definido en SCREENSHOTS_PREFIX, si existe.
    """
    try:
        os.startfile(ruta)
        time.sleep(SCREENSHOT_WAIT)
        screenshot = pyautogui.screenshot()
        prefix = SCREENSHOTS_PREFIX
        filename = f"{prefix}{nombre}.png" if prefix else f"{nombre}.png"
        path_screenshot = os.path.join(BACKUPS_FOLDER, filename)
        screenshot.save(path_screenshot)
        logging.info(f'Pantallazo guardado: {path_screenshot}')
        log_text.value += f"Pantallazo de {nombre} guardado.\n"
        pyautogui.hotkey('alt', 'f4')
        time.sleep(1)
        page.update()
    except Exception as e:
        logging.error(f'Error screenshot carpeta {nombre}: {str(e)}')
        log_text.value += f"Error screenshot {nombre}: {e}\n"
        page.update()

def screenshot_folders(page, log_text):
    """
    Toma screenshots de:
    - Las carpetas de usuario configuradas en USER_FOLDERS.
    - Todas las cuentas de OneDrive corporativo detectadas (ignora OneDrive personal).
    - La carpeta de MEGA (si existe).

    Los archivos se guardan en BACKUPS_FOLDER, nombrados según la carpeta.
    """
    try:
        log_text.value += "Iniciando pantallazos de carpetas...\n"
        page.update()
        # Carpetas de usuario
        for carpeta in USER_FOLDERS:
            ruta = os.path.join(user_profile, carpeta)
            if os.path.exists(ruta):
                screenshot_folder(ruta, carpeta, log_text, page)
        # OneDrive corporativas
        for odrive_path in get_onedrive_corporate_folders():
            nombre_empresa = odrive_path.split("OneDrive - ", 1)[-1].replace("\\", "_").replace("/", "_")
            screenshot_folder(odrive_path, f"OneDrive_{nombre_empresa}", log_text, page)
        # Mega
        if os.path.exists(MEGA_PATH):
            screenshot_folder(MEGA_PATH, 'Mega', log_text, page)
        log_text.value += f"Pantallazos completados. Guardados en {BACKUPS_FOLDER}.\n"
        page.update()
    except Exception as e:
        logging.error(f'Error en screenshots carpetas: {str(e)}')
        log_text.value += f"Error en screenshots: {e}\n"
        page.update()

def update_windows(page, log_text):
    """
    Verifica si hay actualizaciones de Windows disponibles mediante PowerShell.
    Solo informa, no instala.
    """
    try:
        log_text.value += "Verificando actualizaciones de Windows...\n"
        page.update()
        comando = (
            'powershell "Get-WindowsUpdate | Out-String"'
        )
        proceso = subprocess.run(
            comando,
            shell=True,
            capture_output=True,
            text=True
        )
        salida = proceso.stdout
        if 'No updates are available' in salida or 'No hay actualizaciones disponibles' in salida:
            log_text.value += "El sistema está actualizado.\n"
        else:
            log_text.value += f"Resultado: {salida[:500]}\n"
        logging.info(f'Verificación actualizaciones: {salida}')
        page.update()
    except Exception as e:
        logging.error(f'Error al verificar actualizaciones: {str(e)}')
        log_text.value += f"Error verificando actualizaciones: {e}\n"
        page.update()

def open_folder(path):
    """
    Abre una carpeta en el explorador de archivos de Windows.
    """
    if os.path.exists(path):
        os.startfile(path)

def main(page: ft.Page):
    """
    Función principal que arma la interfaz gráfica con Flet y conecta
    los métodos de backup, pantallazo y actualización de Windows.
    """
    log_text = ft.Text(value="", selectable=True, size=13, color=ft.Colors.BLUE_GREY_900)
    chk_backup = ft.Checkbox(label="Backup de correos Outlook a OneDrive", value=False)
    chk_screens = ft.Checkbox(label="Pantallazos de carpetas del sistema y nubes", value=False)
    chk_update = ft.Checkbox(label="Verificar actualizaciones de Windows", value=False)

    def ejecutar_tareas(e):
        log_text.value = ""
        page.update()
        if chk_backup.value:
            backup_outlook(page, log_text)
        if chk_screens.value:
            screenshot_folders(page, log_text)
        if chk_update.value:
            update_windows(page, log_text)
        log_text.value += "\nTareas ejecutadas. Revisa soporte_tool.log para más detalle."
        page.update()

    def abrir_backups(e):
        open_folder(BACKUPS_FOLDER)

    def abrir_logs(e):
        open_folder(os.path.join(os.getcwd(), "soporte_tool.log"))

    page.title = "Soporte Automatizado"
    page.window_width = 520
    page.window_height = 500
    page.window_resizable = False
    page.scroll = ft.ScrollMode.ADAPTIVE

    page.add(
        ft.Row([
            ft.Text("Herramienta de Soporte Automatizado", size=19, weight="bold"),
        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN, vertical_alignment=ft.CrossAxisAlignment.CENTER, expand=True,
            height=45),
        ft.Divider(),
        ft.Column([
            ft.Text("Selecciona las tareas a ejecutar:", size=16, weight="bold"),
            chk_backup,
            chk_screens,
            chk_update,
            ft.Row([
                ft.ElevatedButton("Ejecutar", on_click=ejecutar_tareas, bgcolor=ft.Colors.BLUE_500,
                                  color=ft.Colors.WHITE),
                ft.ElevatedButton("Abrir Backups", on_click=abrir_backups, bgcolor=ft.Colors.BLUE_GREY_300),
                ft.ElevatedButton("Abrir Logs", on_click=abrir_logs, bgcolor=ft.Colors.BLUE_GREY_300)
            ], alignment=ft.MainAxisAlignment.START, spacing=12),
            ft.Text("Log:", size=14, weight="bold", color=ft.Colors.BLUE_GREY_700),
            ft.Container(
                log_text,
                height=180,
                bgcolor=ft.Colors.BLUE_GREY_100,
                padding=8,
                border_radius=10,
                alignment=ft.alignment.top_left,
            ),
            ft.Row(
                [
                    ft.Text(
                        "desarrollado por: Ing. Alejandro Ariaz",
                        size=12,
                        italic=True,
                        color=ft.Colors.BLUE_GREY_500,
                    )
                ],
                alignment=ft.MainAxisAlignment.END,
            )
        ], tight=True, expand=True, spacing=10, alignment=ft.MainAxisAlignment.START)
    )

ft.app(target=main)
