# Outlook Sync Monitor

Monitorea las sincronizaciones de Outlook 2010 (eventos `SyncStart`/`SyncEnd`) y captura contexto visual tipo PSR (ventana activa, elemento enfocado, screenshot). Genera informes PDF con todo el historial capturado.

## Archivos

| Archivo | Descripción |
|---|---|
| `outlook_sync_monitor.py` | Script principal: monitoreo + captura + PDF |
| `close_monitor.py` | Detiene el monitor y genera PDF con los datos pendientes |
| `eventos.jsonl` | Log estructurado de eventos (JSON Lines) |
| `captures/` | Capturas de pantalla del outlook (PNG) |
| `reportes/` | Informes PDF generados (`informe_YYYYMMDD_HHMMSS.pdf`) |
| `outlook_log.txt` | Log de texto plano para debug |

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

## Flujo

1. Ejecutar `outlook_sync_monitor.exe` (o la versión noconsole).
2. El monitor captura cada sincronización de Outlook con contexto visual.
3. Al cerrar con Ctrl+C o ejecutar `close_monitor.exe`, se genera un PDF completo en `reportes/` y se limpia `eventos.jsonl`.
4. Si la PC se reinicia o matan el proceso, los datos sobreviven en `eventos.jsonl`. Al reiniciar el monitor, se genera el PDF automáticamente.
