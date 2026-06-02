import pythoncom
import win32com.client
import datetime
import time
import sys
import os
import json
import traceback
from pathlib import Path

# --- Optional dependencies ---
try:
    import pywinauto
    HAS_PYWAUTO = True
except ImportError:
    HAS_PYWAUTO = False

try:
    from PIL import ImageGrab
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    from fpdf import FPDF
    HAS_FPDF = True
except ImportError:
    HAS_FPDF = False

# --- Paths ---
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path.cwd()

CAPTURES_DIR = BASE_DIR / "captures"
LOG_JSONL = BASE_DIR / "eventos.jsonl"
LOG_TXT = BASE_DIR / "outlook_log.txt"
REPORTS_DIR = BASE_DIR / "reportes"

CAPTURES_DIR.mkdir(exist_ok=True)
REPORTS_DIR.mkdir(exist_ok=True)


# --- Logging ---
def log_txt(message):
    try:
        with open(LOG_TXT, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] - {message}\n")
    except Exception:
        pass


def log_jsonl(entry):
    entry["timestamp"] = datetime.datetime.now().isoformat()
    try:
        with open(LOG_JSONL, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass


# --- Context capture (PSR-style) ---
def capture_context():
    ctx = {
        "ventana": "N/A",
        "elemento_foco": "N/A",
        "ruta_captura": ""
    }

    if not HAS_PYWAUTO or not HAS_PIL:
        return ctx

    try:
        app = pywinauto.Application(backend="uia").connect(
            title_re=".*Outlook.*", timeout=3
        )
        dlg = app.top_window()
        ctx["ventana"] = dlg.window_text()
    except Exception as e:
        log_txt(f"capture_context ventana: {e}")
        print(f"capture_context ventana: {e}")
        return ctx

    try:
        import win32gui
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            text = win32gui.GetWindowText(hwnd)
            cls = win32gui.GetClassName(hwnd)
            ctx["elemento_foco"] = f"{text} ({cls})" if text else f"({cls})"
    except Exception as e:
        log_txt(f"capture_context foco: {e}")
        print(f"capture_context foco: {e}")

    try:
        rect = dlg.rectangle()
        if rect.width() > 0 and rect.height() > 0:
            img = ImageGrab.grab(bbox=(
                rect.left, rect.top, rect.right, rect.bottom
            ))
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            filename = f"{ts}_outlook.png"
            filepath = CAPTURES_DIR / filename
            img.save(filepath)
            ctx["ruta_captura"] = str(filepath)
    except Exception as e:
        log_txt(f"capture_context captura: {e}")
        print(f"capture_context captura: {e}")

    return ctx


# --- PDF generation ---
def generar_pdf():
    if not HAS_FPDF:
        return

    entradas = []
    if LOG_JSONL.exists():
        with open(LOG_JSONL, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    entradas.append(json.loads(line))
                except Exception:
                    continue

    if not entradas:
        log_txt("Sin eventos en el log, no se genera PDF.")
        print("Sin eventos en el log, no se genera PDF.")
        return

    now = datetime.datetime.now()
    filename = f"informe_{now:%Y%m%d_%H%M%S}.pdf"
    output_path = REPORTS_DIR / filename

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "Informe de Sincronizacion - Outlook", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(5)

    for entry in entradas:
        ts = entry.get("timestamp", "N/A")
        accion = entry.get("accion", "N/A")

        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 7, f"{ts}  -  {accion}", new_x="LMARGIN", new_y="NEXT")

        pdf.set_font("Helvetica", "", 9)
        ventana = entry.get("ventana", "N/A")
        elemento = entry.get("elemento_foco", "N/A")
        pdf.cell(0, 5, f"Ventana: {ventana}", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 5, f"Elemento: {elemento}", new_x="LMARGIN", new_y="NEXT")

        ruta = entry.get("ruta_captura", "")
        if ruta and Path(ruta).exists():
            try:
                pdf.image(ruta, x=10, w=180)
                pdf.ln(2)
            except Exception:
                pdf.cell(0, 5, "[captura no disponible]", new_x="LMARGIN", new_y="NEXT")

        pdf.ln(5)

    try:
        pdf.output(str(output_path))
    except Exception as e:
        log_txt(f"Error al guardar PDF: {e}")
        return

    try:
        open(LOG_JSONL, "w", encoding="utf-8").close()
    except Exception:
        pass

    log_txt(f"PDF generado: {filename} ({len(entradas)} eventos), jsonl limpiado.")
    print(f"PDF generado: {filename} ({len(entradas)} eventos), jsonl limpiado.")


# --- Event handler factory ---
def build_handler(sync_obj):
    class SyncHandler:
        _sync_obj = sync_obj

        def OnSyncStart(self, *args):
            try:
                name = self._sync_obj.Name
            except Exception:
                name = "Desconocido"
            ctx = capture_context()
            log_jsonl({
                "accion": f"Inicio de Sincronizacion: {name}",
                "ventana": ctx["ventana"],
                "elemento_foco": ctx["elemento_foco"],
                "ruta_captura": ctx["ruta_captura"]
            })
            log_txt(f"Inicio de Sincronizacion: {name}")
            print(f"Inicio de Sincronizacion: {name}")

        def OnSyncEnd(self, *args):
            try:
                name = self._sync_obj.Name
            except Exception:
                name = "Desconocido"
            ctx = capture_context()
            log_jsonl({
                "accion": f"Fin de Sincronizacion: {name}",
                "ventana": ctx["ventana"],
                "elemento_foco": ctx["elemento_foco"],
                "ruta_captura": ctx["ruta_captura"]
            })
            log_txt(f"Fin de Sincronizacion: {name}")
            print(f"Fin de Sincronizacion: {name}")

    return SyncHandler


# --- Outlook connection helpers ---
def is_outlook_running():
    try:
        pythoncom.GetActiveObject("Outlook.Application")
        return True
    except pythoncom.com_error:
        return False


def subscribe_events(outlook):
    namespace = outlook.GetNamespace("MAPI")
    sync_objects = namespace.SyncObjects
    handlers = []
    for i in range(1, sync_objects.Count + 1):
        sync_obj = sync_objects.Item(i)
        try:
            handler = win32com.client.WithEvents(sync_obj, build_handler(sync_obj))
            handlers.append(handler)
        except Exception as e:
            log_txt(f"AVISO: No se pudo suscribir a '{sync_obj.Name}': {e}")
            print(f"AVISO: No se pudo suscribir a '{sync_obj.Name}': {e}")
    return handlers


def connect():
    if not is_outlook_running():
        return None
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        _ = outlook.GetNamespace("MAPI").SyncObjects.Count
        return outlook
    except Exception as e:
        log_txt(f"ERROR al conectar: {e}")
        print(f"ERROR al conectar: {e}")
        return None


# --- Main ---
def main():
    pythoncom.CoInitialize()

    generar_pdf()

    outlook = None
    handlers = []
    running = True
    ping_interval = 0

    log_txt("Monitor de sincronizacion de Outlook iniciado.")
    print("Monitor de sincronizacion de Outlook iniciado.")
    while running:
        try:
            if outlook is None:
                outlook = connect()
                if outlook is None:
                    time.sleep(5)
                    continue
                handlers = subscribe_events(outlook)
                if not handlers:
                    log_txt("ERROR: No se pudo suscribir a ningun SyncObject.")
                    print("ERROR: No se pudo suscribir a ningun SyncObject.")
                    del outlook
                    outlook = None
                    time.sleep(5)
                    continue
                log_txt("Monitor conectado a Outlook.")
                print("Monitor conectado a Outlook.")

            pythoncom.PumpWaitingMessages()
            time.sleep(0.5)

            ping_interval += 1
            if ping_interval >= 30:
                ping_interval = 0
                if not is_outlook_running():
                    log_txt("Conexion perdida (Outlook cerrado). Reintentando en 5s...")
                    print("Conexion perdida (Outlook cerrado). Reintentando en 5s...")
                    handlers.clear()
                    del outlook
                    outlook = None
                    time.sleep(5)

        except (pythoncom.com_error, AttributeError):
            log_txt("Conexion perdida (error COM). Reintentando en 5s...")
            print("Conexion perdida (error COM). Reintentando en 5s...")
            handlers.clear()
            if outlook is not None:
                del outlook
                outlook = None
            time.sleep(5)

        except KeyboardInterrupt:
            log_txt("Monitor detenido por el usuario.")
            print("Monitor detenido por el usuario.")
            generar_pdf()
            running = False

        except Exception:
            log_txt(f"ERROR inesperado: {traceback.format_exc()}")
            print(f"ERROR inesperado: {traceback.format_exc()}")
            time.sleep(5)

    pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
