import pythoncom
import win32com.client
import win32api
import win32con
import datetime
import time
import os
import sys
import json
import traceback
from pathlib import Path

# --- Optional dependencies ---
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


def load_dotenv():
    base = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    dotenv_path = base / ".env"
    if not dotenv_path.exists():
        return
    try:
        with dotenv_path.open("r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"\'')
                os.environ.setdefault(key, value)
    except Exception:
        pass


load_dotenv()

# --- Paths ---
DATA_DIR = Path(os.environ.get('LOCALAPPDATA', Path.cwd())) / "OutlookSyncMonitor"

CAPTURES_DIR = DATA_DIR / "captures"
LOG_JSONL = DATA_DIR / "eventos.jsonl"
LOG_TXT = DATA_DIR / "outlook_log.txt"
REPORTS_DIR = DATA_DIR / "reportes"

DATA_DIR.mkdir(exist_ok=True)
CAPTURES_DIR.mkdir(exist_ok=True)
REPORTS_DIR.mkdir(exist_ok=True)

# --- Network share for PDF reports ---
RUTA_RED = os.environ.get("RUTA_RED", os.environ.get("ruta_red", ""))
HOSTNAME = os.environ.get("COMPUTERNAME", "unknown")
USERNAME = os.environ.get("USERNAME", "unknown")
NET_USER = os.environ.get("NET_USER", "")
NET_PASS = os.environ.get("NET_PASS", "")


def autenticar_red():
    if not NET_USER:
        return True
    try:
        import subprocess
        cmd = ["net", "use", RUTA_RED]
        if NET_USER:
            cmd.extend(["/USER:" + NET_USER, NET_PASS])
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
        if result.returncode != 0:
            log_txt(f"Error autenticando red: {result.stderr.strip()}")
            return False
        return True
    except Exception as e:
        log_txt(f"Error autenticando red: {e}")
        return False


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

    if not HAS_PIL:
        return ctx

    try:
        import win32gui
    except ImportError:
        return ctx

    hwnd = None
    def _enum_cb(w, _):
        nonlocal hwnd
        if win32gui.IsWindowVisible(w) and "Outlook" in win32gui.GetWindowText(w):
            hwnd = w
            return False
        return True
    try:
        win32gui.EnumWindows(_enum_cb, None)
    except Exception as e:
        log_txt(f"capture enum: {e}")

    if hwnd is None:
        return ctx

    try:
        ctx["ventana"] = win32gui.GetWindowText(hwnd)
    except Exception as e:
        log_txt(f"capture ventana: {e}")

    try:
        fg = win32gui.GetForegroundWindow()
        if fg:
            text = win32gui.GetWindowText(fg)
            cls = win32gui.GetClassName(fg)
            ctx["elemento_foco"] = f"{text} ({cls})" if text else f"({cls})"
    except Exception as e:
        log_txt(f"capture foco: {e}")

    try:
        if not win32gui.IsIconic(hwnd):
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            if right > left and bottom > top:
                img = ImageGrab.grab(bbox=(left, top, right, bottom))
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                filename = f"{ts}_outlook.png"
                filepath = CAPTURES_DIR / filename
                img.save(filepath)
                ctx["ruta_captura"] = str(filepath)
    except Exception as e:
        log_txt(f"capture screenshot: {e}")

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
    filename = f"{HOSTNAME}_{USERNAME}_informe_{now:%Y%m%d_%H%M%S}.pdf"

    autenticar_red()

    ts_inicio = entradas[0].get("timestamp", "N/A")[:19]
    ts_fin = entradas[-1].get("timestamp", "N/A")[:19]

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "Informe de Sincronizacion - Outlook", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 7, f"PC: {HOSTNAME}      Usuario: {USERNAME}      Inicio: {ts_inicio}      Fin: {ts_fin}", new_x="LMARGIN", new_y="NEXT", align="C")
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

    pdf_bytes = pdf.output()

    red_ok = False
    local_ok = False

    try:
        red_dir = Path(RUTA_RED) / HOSTNAME
        red_dir.mkdir(parents=True, exist_ok=True)
        (red_dir / filename).write_bytes(pdf_bytes)
        red_ok = True
        log_txt(f"PDF en red: {red_dir}\\{filename}")
    except Exception as e:
        log_txt(f"Error al guardar PDF en red {RUTA_RED}\\{HOSTNAME}: {e}")

    try:
        (REPORTS_DIR / filename).write_bytes(pdf_bytes)
        local_ok = True
        log_txt(f"PDF local: {REPORTS_DIR}\\{filename}")
    except Exception as e:
        log_txt(f"Error al guardar PDF local: {e}")

    if not red_ok and not local_ok:
        print("Error: no se pudo guardar el PDF en ningun destino. Los datos se conservan.")
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


# --- Console control handler (logoff / shutdown) ---
_ctrl_pdf_generated = False

def ctrl_handler(ctrl_type):
    global _ctrl_pdf_generated
    if ctrl_type in (win32con.CTRL_LOGOFF_EVENT, win32con.CTRL_SHUTDOWN_EVENT):
        if not _ctrl_pdf_generated:
            _ctrl_pdf_generated = True
            generar_pdf()
        return True
    return False


# --- Main ---
def main():
    pythoncom.CoInitialize()

    win32api.SetConsoleCtrlHandler(ctrl_handler, True)

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
