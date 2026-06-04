# Outlook Sync Monitor

Monitorea las sincronizaciones de Outlook 2010 (eventos `SyncStart`/`SyncEnd`) y captura contexto visual tipo PSR (ventana activa, elemento enfocado, screenshot). Genera informes PDF en una carpeta de red compartida.

## Archivos

| Archivo | Descripción |
|---|---|
| `outlook_sync_monitor.py` | Script principal: monitoreo + captura + PDF |
| `close_monitor.py` | Detiene el monitor y genera PDF con los datos pendientes |
| `eventos.jsonl` | Log estructurado de eventos (JSON Lines, local por usuario) |
| `captures/` | Capturas de pantalla del outlook (PNG, local por usuario) |
| `outlook_log.txt` | Log de texto plano para debug (local por usuario) |

## Componentes

### outlook_sync_monitor.py
- Se conecta a Outlook vía COM (pywin32).
- Escucha `SyncStart`/`SyncEnd` en todos los `SyncObjects`.
- En cada evento captura: ventana activa, elemento enfocado (win32gui), screenshot del Outlook.
- Guarda en `eventos.jsonl` (append, crash-safe).
- **Al iniciar**: si `eventos.jsonl` tiene datos (ej. post-reboot), genera PDF y lo limpia.
- **Ctrl+C**: genera PDF con los datos restantes y limpia `eventos.jsonl`.

### close_monitor.py
- Mata el proceso `outlook_sync_monitor.exe` o `outlook_sync_monitor_noconsole.exe`.
- Genera PDF con todos los eventos pendientes y limpia `eventos.jsonl`.

## Destino de los PDFs

Los informes se guardan en una **carpeta de red compartida** (`\\192.168.11.197\test`).

Cada archivo incluye el nombre del usuario de Windows para evitar colisiones:
```
juan_informe_20260601_104300.pdf
maria_informe_20260601_112300.pdf
```

### Configurar credenciales de red

Editar las variables al inicio de `outlook_sync_monitor.py` y `close_monitor.py`:

```python
RUTA_RED = r"\\192.168.11.197\test"   # ruta UNC del servidor
NET_USER = "dominio\\usuario"           # dejar "" si no requiere autenticación
NET_PASS = "contraseña"                # dejar "" si no requiere autenticación
```

- Si `NET_USER` está vacío, no intenta autenticarse (asume permisos de usuario actual o acceso anónimo).
- Si la red no está disponible, el PDF **no se genera** y los datos se conservan en `eventos.jsonl` para reintentar después con `close_monitor.exe`.

## Uso

```powershell
# Directo (requiere dependencias)
python outlook_sync_monitor.py
python close_monitor.py

# Como ejecutable (sin Python)
dist\outlook_sync_monitor.exe
dist\outlook_sync_monitor_noconsole.exe
dist\close_monitor.exe
```

## Compilación

```powershell
# outlook_sync_monitor (consola, Ctrl+C para detener)
pyinstaller --onefile --console --collect-all pywinauto --collect-all PIL --collect-all fpdf --hidden-import=win32com --hidden-import=pythoncom --hidden-import=pywinauto --hidden-import=PIL --hidden-import=PIL.ImageGrab --hidden-import=fpdf --workpath "%TEMP%\opencode\pyi_build" --specpath "%TEMP%\opencode" --distpath dist --name outlook_sync_monitor outlook_sync_monitor.py

# outlook_sync_monitor (sin consola, matar desde Task Manager)
pyinstaller --onefile --noconsole --collect-all pywinauto --collect-all PIL --collect-all fpdf --hidden-import=win32com --hidden-import=pythoncom --hidden-import=pywinauto --hidden-import=PIL --hidden-import=PIL.ImageGrab --hidden-import=fpdf --workpath "%TEMP%\opencode\pyi_build_nc" --specpath "%TEMP%\opencode" --distpath dist --name outlook_sync_monitor_noconsole outlook_sync_monitor.py

# close_monitor
pyinstaller --onefile --console --collect-all fpdf --hidden-import=fpdf --workpath "%TEMP%\opencode\pyi_build_close" --specpath "%TEMP%\opencode" --distpath dist --name close_monitor close_monitor.py
```

## Dependencias

- pywin32
- pywinauto
- Pillow
- fpdf2
- PyInstaller (solo para compilar)

## Flujo multi-usuario

1. Cada usuario Windows ejecuta `outlook_sync_monitor.exe`.
2. Los datos locales (`eventos.jsonl`, `captures/`, `outlook_log.txt`) se guardan donde esté el `.exe`.
3. Al cerrar con Ctrl+C o ejecutar `close_monitor.exe`, se genera el PDF en la carpeta de red con el nombre del usuario.
4. Si la PC se reinicia o matan el proceso, los datos sobreviven localmente. Al reiniciar el monitor, se genera el PDF automáticamente.
5. Cada usuario tiene sus propios archivos locales, los PDFs se centralizan en la carpeta de red compartida.
