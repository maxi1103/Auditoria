# Outlook Sync Monitor

Monitorea eventos de sincronizacion (Enviar/Recibir) de Microsoft Outlook, captura contexto visual (estilo PSR) y genera informes PDF.

## Requisitos

- Windows 7 o superior
- Microsoft Outlook 2010 o superior
- No requiere Python instalado (usar el `.exe`)

## Instalacion

### Ejecutable (recomendado)

Descarga de `dist/`:

| Archivo | Descripcion |
|---------|-------------|
| `outlook_sync_monitor.exe` | Con ventana de consola (ver estado y Ctrl+C) |
| `outlook_sync_monitor_noconsole.exe` | Sin ventana (segundo plano, matar con Task Manager) |

### Script Python

```powershell
pip install pywin32 pywinauto pillow fpdf2
python outlook_sync_monitor.py
```

## Uso

1. Asegurate de que Outlook este abierto.
2. Ejecuta el `.exe` deseado.
3. Presiona `Enviar/Recibir` en Outlook o espera una sincronizacion automatica.
4. Para detener:
   - **Version console**: Presiona `Ctrl+C` en la ventana.
   - **Version noconsole**: Usa el Administrador de Tareas (Task Manager).

## Archivos generados

| Archivo/Carpeta | Descripcion |
|-----------------|-------------|
| `outlook_log.txt` | Log de texto plano (eventos y estado) |
| `eventos.jsonl` | Log estructurado en JSON Lines (una linea por evento) |
| `captures/` | Capturas de pantalla `.png` del area de Outlook |
| `informe_sincronizacion.pdf` | Informe PDF con eventos y capturas |
| `pdf_state.json` | Estado para generar PDFs incrementales |

## Captura de contexto (estilo PSR)

Al detectar cada evento de sincronizacion, el script captura automaticamente:

- **Ventana activa**: titulo exacto de la ventana de Outlook
- **Elemento enfocado**: nombre y tipo del control UIA (ej. `Enviar y recibir todas las carpetas (Botón)`)
- **Captura de pantalla**: imagen PNG del area de la ventana de Outlook

## Guardado incremental

Cada evento se guarda inmediatamente en `eventos.jsonl` (formato append):

```json
{"timestamp":"2026-05-29T15:00:00","accion":"Inicio de Sincronizacion: Grupo 1","ventana":"Bandeja de entrada - Outlook","elemento_foco":"Enviar y recibir todas las carpetas (Botón)","ruta_captura":"..."}
```

## Generacion de PDF

- Al **iniciar** el programa, se genera un PDF incremental con los eventos no incluidos aun.
- Al presionar `Ctrl+C` (version console), tambien se genera el PDF.
- Si ya existe un PDF, solo se anaden los eventos nuevos (sin duplicados).

## Auto-reconexion

Si cierras y vuelves a abrir Outlook, el monitor se reconecta automaticamente en 5 segundos:

```
[2026-05-29 15:00:00] - Conexion perdida (Outlook cerrado). Reintentando en 5s...
[2026-05-29 15:00:05] - Monitor conectado a Outlook.
```

## Compilar desde codigo fuente

```powershell
pip install pyinstaller pywin32 pywinauto pillow fpdf2

:: Con consola
pyinstaller --onefile --console ^
  --hidden-import=win32com --hidden-import=pythoncom ^
  --hidden-import=win32com.client --hidden-import=pywinauto ^
  --hidden-import=PIL --hidden-import=PIL.ImageGrab ^
  --hidden-import=fpdf --collect-all pywinauto ^
  --collect-all PIL --collect-all fpdf ^
  outlook_sync_monitor.py

:: Sin consola
pyinstaller --onefile --noconsole ^
  --hidden-import=win32com --hidden-import=pythoncom ^
  --hidden-import=win32com.client --hidden-import=pywinauto ^
  --hidden-import=PIL --hidden-import=PIL.ImageGrab ^
  --hidden-import=fpdf --collect-all pywinauto ^
  --collect-all PIL --collect-all fpdf ^
  outlook_sync_monitor.py
```

## Estructura del proyecto

```
outlook_sync_monitor.py              # Script principal
outlook_log.txt                       # Log de texto plano
eventos.jsonl                         # Log estructurado JSON Lines
pdf_state.json                         # Estado del PDF incremental
captures/                             # Capturas de pantalla PNG
  YYYYMMDD_HHMMSS_ffffff_outlook.png
dist/
  outlook_sync_monitor.exe            # Version con consola
  outlook_sync_monitor_noconsole.exe  # Version sin consola
README.md                             # Este archivo
```
