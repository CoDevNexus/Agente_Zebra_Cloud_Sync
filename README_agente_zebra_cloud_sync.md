# Zebra Cloud Sync

Aplicación de escritorio para Windows que trabaja en segundo plano, captura códigos de barras desde un escáner (Zebra, SAT, Honeywell u otras marcas), consulta información en MySQL, guarda los escaneos localmente en SQLite y sincroniza los registros a Google Sheets.

Compatible con escáneres en modo **HID Teclado** (sin puerto COM) y en modo **Serial CDC** (puerto COM).

---

## Historial de cambios

### Versión actual

| Cambio | Descripción |
|---|---|
| Soporte HID Teclado | El escáner puede operar como teclado USB sin necesitar puerto COM. Requiere `pynput`. |
| Soporte Serial CDC | Los escáneres Zebra con driver CDC (puerto COM) se conectan con parámetros correctos: `xonxoff=False`, `rtscts=False`, `dsrdtr=False`. |
| Modo de lectura serial universal | Se usa `read(4096)` con `inter_byte_timeout=0.1s` en lugar de `readline()`. Funciona con cualquier terminador: CR, LF, CR+LF o sin terminador. |
| Compatibilidad multi-marca | Funciona con Zebra DS2278, SAT, Honeywell y cualquier escáner que emule teclado HID o use puerto COM. |
| Umbral HID configurable | Campo `hid_inter_char_ms` (default 150 ms) para distinguir el escáner de la escritura humana. Aumentar si el escáner pierde dígitos. |
| Selector de modo escáner en UI | La pestaña Configuración permite elegir entre **Serial (CDC)** y **HID Teclado**. El cambio requiere reiniciar la app. |
| Selector de baudrate en UI | `ComboBox` con 300 / 1200 / 2400 / 4800 / 9600 / 19200 / 38400 / 115200. |
| Rutas del exe corregidas | `BASE_DIR` usa `sys.executable` cuando el programa está compilado con PyInstaller. Los archivos `config_agente.json`, `.db` y `.log` se crean junto al `.exe`, no en una carpeta temporal. |
| Autoinicio con Registro de Windows | `StartupManager` usa `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` en lugar de un archivo `.bat` en Startup. Más confiable, no requiere permisos especiales y no es afectado por OneDrive. |
| MySQL compatible con Python 3.13+ | `use_pure=True` fuerza el backend Python puro del conector y evita el error `RuntimeError: Failed raising error` de la extensión C en Python 3.13 y 3.14. |
| Eliminado `is_connected()` | Reemplazado por `SELECT 1` directo. `is_connected()` tiene bugs en `mysql-connector 9.x`. |
| Traceback completo en Test MySQL | El botón **Test MySQL** captura la causa raíz del error y la escribe en el log con traceback completo. |
| UI compactada para 768px | La pestaña Configuración entra completa en pantallas de 768px de alto. |
| Error de autoinicio visible en UI | Si `StartupManager` falla al activar el autoinicio, se muestra un `messagebox` con el error en lugar de fallar silenciosamente. |
| Descripción de puertos en ComboBox | El selector de puerto COM muestra la descripción del dispositivo (ej. `COM4 — Zebra CDC...`) no solo el nombre. |
| `pynput` añadido a dependencias | `requirements.txt` incluye `pynput>=1.7.6`. Es opcional en ejecución: si no está instalado, se desactiva el modo HID con mensaje de error claro. |
| `mysql-connector` fijado a `9.6.0` | Se fija la versión en `requirements.txt` para evitar problemas con versiones superiores. |
| `.spec` actualizado | Incluye `collect_data_files("customtkinter")`, `hiddenimports` completos para `pynput`, `mysql.connector.plugins.caching_sha2_password`, `gspread`, `openpyxl`, `upx=False`. |

---

## Qué hace

* Lee códigos de barras desde un escáner en modo **HID Teclado** (sin puerto COM) o **Serial CDC** (puerto COM).
* Compatible con escáneres Zebra DS2278/CR2278, SAT, Honeywell y cualquier escáner que emule teclado USB.
* Consulta descripción y stock en MySQL.
* Guarda cada lectura en una base local SQLite para no perder información aunque no haya conexión.
* Sincroniza los registros a Google Sheets.
* Soporta anulación automática por repeticiones dentro de una ventana de tiempo configurable.
* Puede minimizarse a la bandeja del sistema.
* Puede iniciar automáticamente con Windows (vía Registro de Windows).
* Permite exportar historial a Excel por rango de fechas.
* Protege la pestaña de configuración con contraseña.

---

## Requisitos

### Sistema operativo

* Windows 10 o Windows 11

### Python

* Python 3.11 o superior recomendado
* Probado en Python 3.12.3 (despliegue) y Python 3.14.3 (desarrollo)
* En Python 3.13+ el conector MySQL usa `use_pure=True` automáticamente

### Dependencias de Python

```bash
pip install customtkinter pyserial mysql-connector-python==9.6.0 gspread oauth2client pystray pillow openpyxl pynput pyinstaller
```

O usando el archivo de dependencias:

```bash
pip install -r requirements.txt
```

### `requirements.txt`

```
customtkinter>=5.2.2
pyserial>=3.5
mysql-connector-python==9.6.0
gspread>=5.11.0
oauth2client>=4.1.3
pystray>=0.19.5
Pillow>=10.0.0
openpyxl>=3.1.0
pynput>=1.7.6
```

> **Nota:** `pynput` es necesario solo si usas el modo **HID Teclado**. Si no está instalado, la app detecta la ausencia y avisa en la barra de estado.

---

## Estructura recomendada del proyecto

```text
capturador-codigoZebra/
│
├─ agente_zebra_cloud_sync.py
├─ config_agente.json
├─ credentials.json
├─ agente_buffer.db
├─ agente_zebra.log
├─ README_Zebra_Cloud_Sync.md
└─ dist/
```

### Archivos importantes

* `agente_zebra_cloud_sync.py`: programa principal.
* `config_agente.json`: guarda la configuración para no volver a pedirla al abrir.
* `credentials.json`: credenciales de Google para acceder a Sheets.
* `agente_buffer.db`: base local SQLite.
* `agente_zebra.log`: archivo de log para diagnóstico.
* `dist/`: carpeta donde PyInstaller genera el ejecutable.

---

## Cómo ejecutar en modo desarrollo

Ubícate en la carpeta del proyecto y ejecuta:

```bash
python agente_zebra_cloud_sync.py
```

Si quieres que arranque oculto:

```bash
python agente_zebra_cloud_sync.py --hidden
```

---

## Configuración inicial

La primera vez debes completar estos datos:

### 1. Escáner

Primero selecciona el **modo de escáner**:

#### Modo HID Teclado (recomendado para Zebra DS2278 y escáneres modernos)

* El escáner aparece en el Administrador de dispositivos como **Teclado HID**, no como puerto COM.
* No necesitas seleccionar puerto ni baudrate.
* Instala `pynput` si no está instalado.
* Ajusta el campo **Umbral inter-caracteres HID (ms)** si el escáner pierde dígitos (sube el valor) o registra pulsaciones de teclado humanas (baja el valor). El valor por defecto es 150 ms.

#### Modo Serial / CDC

* El escáner aparece en el Administrador de dispositivos como un puerto COM (con driver Zebra CDC u otro driver USB-serial).
* Selecciona el puerto COM correcto (el combo muestra el nombre del dispositivo).
* Selecciona la velocidad (baudrate). Para Zebra DS2278: `9600`.
* El escáner debe estar conectado, encendido y reconocido por Windows.

> **Zebra DS2278 con base CR2278:** instala el driver **Zebra CDC** para que aparezca como puerto COM, o usa directamente el modo HID Teclado sin instalar ningún driver adicional.

### 2. MySQL

* Host (ej. `127.0.0.1` o dominio)
* Usuario
* Password
* Base de datos

### 3. Google Sheets

* ID del Google Sheet
* Ruta al archivo `credentials.json`

### 4. SQLite

* Ruta del archivo `.db` (por defecto junto al ejecutable)

### 5. Operación

* Intervalo de sincronización (segundos, mínimo 3)
* Ventana de corrección en segundos
* Cantidad de repeticiones para anular
* Cantidad de registros por página en el historial

### 6. Autoinicio con Windows

Al activar **Iniciar con Windows en segundo plano**, el programa registra un valor en:

```
HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
```

Para verificar que quedó registrado:

```powershell
Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
```

Deberías ver una entrada llamada `Agente Zebra Cloud Sync`.

### 7. Contraseña de configuración

La primera vez que se entre a la pestaña de configuración, el sistema pedirá crear una contraseña. Después de eso, la pedirá cada vez que se quiera volver a entrar.

---

## Lógica principal del sistema

### Modos de captura del escáner

#### Modo HID Teclado (`scanner_mode: "hid"`)

1. `HIDScannerWorker` registra un listener global de teclado con `pynput`.
2. Los caracteres que llegan con un intervalo menor a `hid_inter_char_ms` se consideran del escáner; los demás se descartan como escritura humana.
3. Al recibir Enter / CR / LF se procesa el código acumulado.
4. El umbral `hid_inter_char_ms` se puede ajustar sin reiniciar la app desde la configuración.

#### Modo Serial CDC (`scanner_mode: "serial"`)

1. `ScannerWorker` abre el puerto COM con parámetros CDC: `xonxoff=False`, `rtscts=False`, `dsrdtr=False`.
2. Lee hasta 4096 bytes con `inter_byte_timeout=0.1s`. No depende de ningún terminador en particular (funciona con CR, LF, CR+LF o sin terminador).
3. Si el puerto no existe o está ocupado, el worker reintenta cada 2 segundos.

### Escaneo (ambos modos)

1. El escáner envía el código.
2. El sistema limpia el valor leído (strip de espacios y saltos de línea).
3. Busca el producto en MySQL.
4. Guarda el registro en SQLite.

### Si MySQL falla

El código igual se guarda localmente con descripción `Error DB`, para no detener la operación.

### Cola local

Todos los escaneos se guardan primero en SQLite.

### Ventana de corrección

Cada producto entra primero en una ventana de espera.

Si el mismo código se repite la cantidad configurada dentro de ese tiempo:

* se anula el grupo de lecturas de ese código dentro de esa ventana,
* el historial lo muestra como anulado,
* y ese registro también puede quedar reflejado en Google Sheets.

### Sincronización

Después de que vence el tiempo de espera:

* si el registro sigue válido, se sincroniza a Google Sheets,
* si fue anulado, se envía con estado de anulado,
* si no hay internet, queda pendiente para el siguiente ciclo.

---

## Qué muestra la aplicación

## Pestaña Historial

Muestra:

* ID
* Código
* Descripción
* Stock
* Fecha
* Tiempo restante
* Estado

Estados posibles:

* En ventana
* Pendiente de envío
* Sincronizado
* Anulado
* Anulado y enviado

También permite:

* refrescar el historial,
* exportar a Excel por rango,
* paginar resultados.

## Pestaña Configuración

Permite editar y probar:

* Puerto COM
* MySQL
* Google Sheets
* SQLite
* auto inicio con Windows
* cambio de contraseña

---

## Botones de prueba

El sistema incluye botones de prueba para validar antes de guardar:

* `Test COM`
* `Test MySQL`
* `Test Google`
* `Test SQLite`

Esto ayuda a confirmar que cada servicio está bien configurado.

---

## Exportación a Excel

La ventana principal permite exportar el historial a Excel por rango de fechas.

### Formato esperado

* Desde: `YYYY-MM-DD`
* Hasta: `YYYY-MM-DD`

### Qué incluye el Excel

* encabezados automáticos,
* registros del rango solicitado,
* tiempo restante,
* estado,
* detalle del error o motivo de anulación.

---

## Inicio automático con Windows

Si activas la opción **Iniciar con Windows en segundo plano**, el programa escribe una entrada en el Registro de Windows:

```
HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
  Agente Zebra Cloud Sync = "C:\ruta\ZebraCloudSync.exe" --hidden
```

Con eso, cuando la computadora encienda y el usuario inicie sesión:

* el agente arrancará automáticamente,
* se ejecutará oculto,
* seguirá trabajando en la bandeja del sistema.

### Por qué Registro y no carpeta Startup

El método anterior usaba un archivo `.bat` en la carpeta Startup del usuario. Con OneDrive activo, esa carpeta puede estar sincronizada o bloqueada, causando que el archivo no se cree. El Registro de Windows no tiene ese problema y es el método estándar recomendado por Microsoft.

### Verificar que está activo

```powershell
Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
```

### Eliminar manualmente si es necesario

```powershell
Remove-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "Agente Zebra Cloud Sync"
```

---

# Configuración de Google Sheets

## 1. Crear el archivo JSON

Debes usar una **cuenta de servicio** en Google Cloud.

Resumen del proceso:

1. Crear proyecto en Google Cloud.
2. Habilitar Google Sheets API y Google Drive API.
3. Crear cuenta de servicio.
4. Crear una clave tipo JSON.
5. Descargar el archivo `credentials.json`.

## 2. Compartir la hoja

Debes abrir el Google Sheet y compartirlo con el correo que aparece dentro del JSON en el campo `client_email`.

## 3. Obtener el Sheet ID

Se toma de la URL del Google Sheet, entre `/d/` y `/edit`.

Ejemplo:

```text
https://docs.google.com/spreadsheets/d/1ABCxyz1234567890/edit#gid=0
```

El ID sería:

```text
1ABCxyz1234567890
```

---

# Cómo crear el ejecutable (.exe)

## Opción recomendada

Para una aplicación de este tipo, lo más recomendable es generar un ejecutable sin consola visible.

## 1. Instalar PyInstaller

```bash
pip install pyinstaller
```

## 2. Ubicarte en la carpeta del proyecto

```bash
cd ruta\de\tu\proyecto
```

## 3. Generar el ejecutable básico

```bash
pyinstaller --noconsole --onefile --name ZebraCloudSync agente_zebra_cloud_sync.py
```

Esto genera:

* carpeta `build/`
* carpeta `dist/`
* archivo `ZebraCloudSync.spec`
* ejecutable dentro de `dist/`

---

## Ejecutable recomendado para este proyecto

### Opción A: onefile

Genera un solo `.exe`.

```bash
pyinstaller --noconsole --onefile --name ZebraCloudSync agente_zebra_cloud_sync.py
```

### Opción B: onedir

Genera una carpeta con el `.exe` y dependencias.
Suele ser más estable para programas con interfaz, bandeja del sistema y varias librerías.

```bash
pyinstaller --noconsole --onedir --name ZebraCloudSync agente_zebra_cloud_sync.py
```

Si buscas facilidad para distribuir, usa `onefile`.
Si buscas más estabilidad y depuración más simple, usa `onedir`.

---

## Si quieres agregar icono al ejecutable

```bash
pyinstaller --noconsole --onefile --name ZebraCloudSync --icon icono.ico agente_zebra_cloud_sync.py
```

---

## Si necesitas incluir archivos adicionales

En general, para este proyecto es mejor dejar estos archivos fuera del `.exe` y colocarlos junto al ejecutable:

* `credentials.json`
* `config_agente.json`
* `agente_buffer.db`

Esto facilita:

* cambiar configuraciones,
* reemplazar credenciales,
* conservar historial,
* evitar recompilar.

Pero si necesitas agregar archivos de datos al empaquetado, puedes usar `--add-data`.

Ejemplo en Windows:

```bash
pyinstaller --noconsole --onefile --name ZebraCloudSync --add-data "archivo_origen;destino" agente_zebra_cloud_sync.py
```

---

## Uso con archivo .spec

Después de la primera compilación, PyInstaller genera un archivo `.spec`.

Ejemplo:

```text
ZebraCloudSync.spec
```

Luego puedes recompilar usando:

```bash
pyinstaller ZebraCloudSync.spec
```

Esto es útil cuando quieres fijar opciones y no escribir el comando completo cada vez.

---

## Comando práctico recomendado

Para este proyecto, una base útil es esta:

```bash
pyinstaller --noconsole --onedir --name ZebraCloudSync agente_zebra_cloud_sync.py
```

Si luego quieres pasar a archivo único:

```bash
pyinstaller --noconsole --onefile --name ZebraCloudSync agente_zebra_cloud_sync.py
```

---

## Dónde queda el ejecutable

Después de compilar:

### Si usas `--onefile`

El ejecutable queda normalmente en:

```text
dist\ZebraCloudSync.exe
```

### Si usas `--onedir`

Queda una carpeta como:

```text
dist\ZebraCloudSync\
```

Y dentro estará el `.exe`.

---

## Cómo distribuirlo

### Si usas onefile

Entrega:

* `ZebraCloudSync.exe`
* `credentials.json` si no lo vas a generar en cliente

### Si usas onedir

Entrega toda la carpeta:

* `dist\ZebraCloudSync\`

Y junto a ella o dentro:

* `credentials.json`
* `config_agente.json` si ya viene preconfigurado

---

## Recomendación importante

Para operación real en cliente:

* primero prueba en modo `.py`,
* luego genera `onedir`,
* valida escáner, MySQL, Google Sheets y auto inicio,
* y solo después genera la versión `onefile` si la necesitas.

---

## Ejecutable con PyInstaller

### Compilar usando el `.spec` incluido (recomendado)

El proyecto incluye `ZebraCloudSync.spec` preconfigurado con:

* `collect_data_files("customtkinter")` — incluye temas e íconos de CustomTkinter.
* `hiddenimports` completos: `pynput`, `mysql.connector.plugins.caching_sha2_password`, `gspread`, `openpyxl`, `pystray._win32`, `serial.tools.list_ports`.
* `upx=False` — evita corrupción del ejecutable en Windows 11.
* `console=False` — sin ventana de consola.
* `onefile=True` — un solo `.exe`.

```bash
pyinstaller ZebraCloudSync.spec
```

### Limpiar antes de compilar

```powershell
Remove-Item -Recurse -Force build, dist
pyinstaller ZebraCloudSync.spec
```

> **Importante:** si el proyecto está en una carpeta de OneDrive, pausa la sincronización antes de compilar para evitar `PermissionError`.

### Generar por primera vez sin `.spec`

```bash
pyinstaller --noconsole --onefile --name ZebraCloudSync agente_zebra_cloud_sync.py
```

---

## Problemas comunes al crear el ejecutable

### 1. El ejecutable abre y se cierra

Prueba primero sin `--noconsole` para ver el error:

```bash
pyinstaller --onefile --name ZebraCloudSync agente_zebra_cloud_sync.py
```

### 2. Archivos de config se crean en carpeta temporal

Esto ocurre si `BASE_DIR` apunta a `sys._MEIPASS` (carpeta temporal de extracción).

La versión actual usa:

```python
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).resolve().parent
```

Así `config_agente.json`, `.db` y `.log` quedan junto al `.exe`.

### 3. El auto inicio no funciona

Verifica con PowerShell:

```powershell
Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
```

Si no aparece la entrada, revisa `agente_zebra.log` para ver el error al registrar el autoinicio.

### 4. Google no conecta

Verifica:

* que el `credentials.json` sea válido,
* que la hoja esté compartida con el correo `client_email` del JSON,
* que el Sheet ID sea correcto.

### 5. Error MySQL: `RuntimeError: Failed raising error`

Esto ocurre con la extensión C de `mysql-connector` en Python 3.13 o 3.14.

La versión actual ya incluye `use_pure=True` en la conexión, lo que fuerza el backend Python puro y evita el error.

Si compilas manualmente el `.exe`, asegúrate de que `mysql.connector.plugins.caching_sha2_password` esté en `hiddenimports` del `.spec`.

### 6. El escáner no responde (modo Serial)

Verifica:

* puerto COM correcto,
* escáner conectado y encendido,
* escáner reconocido por Windows con driver CDC instalado,
* que otro programa no esté usando el mismo COM.

### 7. El escáner pierde dígitos (modo HID)

Aumenta el valor de **Umbral inter-caracteres HID (ms)** en la configuración. Empieza con 200 ms. Si el problema persiste, prueba 300 ms.

### 8. La aplicación registra pulsaciones de teclado humanas como escaneos (modo HID)

Reduce el valor de **Umbral inter-caracteres HID (ms)**. Prueba con 80–100 ms. Un humano raramente escribe más de 10 caracteres por segundo (100 ms entre teclas), mientras que un escáner los envía todos en menos de 50 ms.

### 9. `PermissionError` al compilar con OneDrive activo

Pausa la sincronización de OneDrive antes de ejecutar PyInstaller. Si el error persiste, mueve el proyecto temporalmente fuera de la carpeta de OneDrive para compilar.

---

## Flujo recomendado de entrega

1. Instalar Python y dependencias.
2. Probar el script en desarrollo.
3. Validar COM, MySQL, Google y SQLite con los botones de prueba.
4. Guardar configuración.
5. Probar escaneo real.
6. Probar sincronización.
7. Probar anulación por repetición.
8. Probar exportación a Excel.
9. Activar inicio con Windows si aplica.
10. Generar ejecutable con PyInstaller.
11. Validar el ejecutable en una máquina limpia.

---

## Seguridad y buenas prácticas

* No publiques el `credentials.json` en repositorios públicos.
* Protege el acceso a Configuración con contraseña.
* Haz respaldo de `agente_buffer.db` si el historial es importante.
* Conserva `agente_zebra.log` para soporte técnico.

---

## Resumen final

Este proyecto está pensado para operar de forma silenciosa, continua y tolerante a fallos. Incluso si se cae internet o MySQL, los datos quedan almacenados localmente hasta poder sincronizarse.

La compilación a `.exe` se hace con PyInstaller, recomendando primero probar `onedir` y luego, si todo está estable, generar `onefile` para distribución más simple.


Sí. Desde la carpeta del proyecto, usa este comando para generar un .exe único y sin consola en Windows:

pyinstaller --noconsole --onefile --name ZebraCloudSync agente_zebra_cloud_sync.py

El ejecutable quedará en:

dist\ZebraCloudSync.exe

PyInstaller soporta --onefile para un solo ejecutable y --noconsole/--windowed para ocultar la consola.

Si quieres una versión más estable para pruebas, en carpeta en vez de archivo único, usa:

pyinstaller --noconsole --onedir --name ZebraCloudSync agente_zebra_cloud_sync.py




************************
OneDrive está bloqueando la carpeta build. Elimínala manualmente primero y luego corre sin --clean:

Remove-Item -Recurse -Force "build\ZebraCloudSync"
pyinstaller ZebraCloudSync.spec




