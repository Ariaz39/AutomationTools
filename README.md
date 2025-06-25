# Herramienta de Soporte Automatizado (Python + Flet)

> Automatiza tareas de respaldo y soporte en equipos Windows:  
> copia de seguridad de correos Outlook, pantallazos de carpetas clave (incluyendo varias cuentas de OneDrive corporativas), verificación de actualizaciones, logs y mucho más.

---

## 🧑‍💻 Para el Usuario Final

### ¿Qué hace esta herramienta?

- **Respalda automáticamente** tus correos de Outlook (archivos .pst/.ost) a la carpeta principal de tu OneDrive corporativo.
- **Toma pantallazos** de las carpetas Descargas, Documentos, Imágenes, Música, Videos y todas las cuentas de OneDrive corporativas configuradas, así como MEGA (si existe).
- **Verifica si hay actualizaciones de Windows pendientes.**
- **Registra todo lo que hace** en un archivo de logs.

---

### ¿Cómo instalar y usar?

1. **Instala Python**  
   Descárgalo desde [python.org/downloads/windows](https://www.python.org/downloads/windows/)  
   Durante la instalación, **marca la opción "Add Python to PATH"**.

2. **Descarga la carpeta del proyecto** (pregunta a tu área de soporte si no la tienes).
   - Incluye al menos: `main.py`, `.env`, y `requirements.txt`.

3. **Abre la carpeta** del proyecto en el Explorador de Archivos de Windows.

4. **Abre una terminal dentro de esa carpeta** (Shift + click derecho > “Abrir ventana de PowerShell aquí”).

5. **Crea y activa el entorno virtual**:
   ```
   python -m venv venv
   venv\Scripts\activate
   ```

6. **Instala las dependencias**:
   ```
   pip install -r requirements.txt
   ```

7. **Lanza la aplicación**:
   ```
   python main.py
   ```

8. **En la ventana que se abre:**
   - Marca las tareas que desees ejecutar.
   - Haz clic en “Ejecutar”.
   - Puedes abrir la carpeta de backups o los logs con los botones correspondientes.

---

### Notas para el usuario

- Los **pantallazos** se guardan en la carpeta `backups` del mismo directorio donde está la app.
- El **respaldo de Outlook** se guarda en tu cuenta de OneDrive corporativa, en la subcarpeta definida en `.env` (por defecto: `backup_correos`).
- El **log de actividad** queda en `soporte_tool.log`.
- **No necesitas modificar el archivo `.env`** a menos que te lo indique soporte.
- Si tu equipo es muy lento o rápido, puedes ajustar el tiempo de espera (`SCREENSHOT_WAIT`) en `.env`.

---

## ⚙️ Para Desarrolladores

### Estructura del Proyecto

```
/HerramientaSoporte/
│   main.py
│   .env
│   requirements.txt
│   soporte_tool.log
│   /backups/
```

---

### Instalación de dependencias

- Crea y activa un entorno virtual:
  ```
  python -m venv venv
  venv\Scripts\activate
  ```
- Instala todas las librerías con:
  ```
  pip install -r requirements.txt
  ```

---

### requirements.txt de ejemplo

```
flet
pyautogui
pillow
python-dotenv
```

---

### Variables clave y personalización

- **main.py**  
  Código principal. Puedes modificar, extender o crear nuevas funciones siguiendo los ejemplos y la documentación en los docstrings de cada función.

- **.env**  
  Aquí se configuran nombres de carpetas, subcarpetas, prefijos y otros parámetros.  
  Ejemplo:
  ```
  ONEDRIVE_PATH=OneDrive
  ONEDRIVE_BACKUP_FOLDER=backup_correos
  MEGA_PATH=MEGA
  SCREENSHOTS_PREFIX=
  USER_FOLDERS=Descargas,Documentos,Imágenes,Música,Videos
  BACKUPS_FOLDER=backups
  SCREENSHOT_WAIT=2
  ```
  > El script arma las rutas dinámicamente usando el usuario logueado.

---

### Cómo funciona la detección de OneDrive

- **Backup de Outlook:**  
  Solo se realiza a la cuenta principal indicada en `.env`.

- **Pantallazos de OneDrive:**  
  Se detectan todas las carpetas en `C:\Users\<usuario>\OneDrive - NombreEmpresa`  
  (solo corporativas, ignora la personal `OneDrive`).
  El pantallazo se nombra como `OneDrive_NombreEmpresa.png`.

---

### Extensión y mantenimiento

- **Agregar nuevas carpetas a capturar:**  
  Modifica la variable `USER_FOLDERS` en `.env`.

- **Cambiar el destino del backup:**  
  Ajusta `ONEDRIVE_PATH` y `ONEDRIVE_BACKUP_FOLDER`.

- **Personalizar los nombres de los screenshots:**  
  Cambia `SCREENSHOTS_PREFIX`.

- **Agregar una nueva funcionalidad:**  
  - Crea una nueva función siguiendo el patrón de `backup_outlook` o `screenshot_folder`.
  - Agrega la casilla y lógica en la función `main()`.

- **Manejo de errores/logs:**  
  Todos los errores y operaciones importantes quedan en `soporte_tool.log`.

---

## 🛠️ Soporte

- Si tienes problemas o necesitas ayuda, revisa primero el archivo `soporte_tool.log` y luego contacta a soporte técnico.
- Para desarrollo, revisa los comentarios y docstrings en el código fuente (`main.py`).

---

*Desarrollado en Python 3.10+ y Flet. Puedes usar, modificar y redistribuir este código bajo los términos de tu organización.*
