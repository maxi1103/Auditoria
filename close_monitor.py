import subprocess
import sys
import datetime
import json
import os
from pathlib import Path

# --- Paths (same as outlook_sync_monitor.py) ---
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path.cwd()

LOG_JSONL = BASE_DIR / "eventos.jsonl"
LOG_TXT = BASE_DIR / "outlook_log.txt"

RUTA_RED = r"\\192.168.11.197\test"
USERNAME = os.environ.get("USERNAME", "unknown")
NET_USER = ""
NET_PASS = ""


def autenticar_red():
    if not NET_USER:
        return True
    try:
        cmd = ["net", "use", RUTA_RED]
        if NET_USER:
            cmd.extend(["/USER:" + NET_USER, NET_PASS])
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
        if result.returncode != 0:
            log_txt(f"close_monitor: Error autenticando red: {result.stderr.strip()}")
            return False
        return True
    except Exception as e:
        log_txt(f"close_monitor: Error autenticando red: {e}")
        return False


# --- Logging ---
def log_txt(message):
    try:
        with open(LOG_TXT, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] - {message}\n")
    except Exception:
        pass


# --- PDF generation (standalone copy from outlook_sync_monitor.py) ---
def generar_pdf():
    try:
        from fpdf import FPDF
    except ImportError:
        log_txt("close_monitor: fpdf2 no disponible.")
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
        log_txt("close_monitor: Sin eventos en el log, no se genera PDF.")
        return

    now = datetime.datetime.now()
    filename = f"{USERNAME}_informe_{now:%Y%m%d_%H%M%S}.pdf"

    autenticar_red()
    output_path = Path(RUTA_RED) / filename

    ts_inicio = entradas[0].get("timestamp", "N/A")[:19]
    ts_fin = entradas[-1].get("timestamp", "N/A")[:19]

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "Informe de Sincronizacion - Outlook", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 7, f"Usuario: {USERNAME}      Inicio: {ts_inicio}      Fin: {ts_fin}", new_x="LMARGIN", new_y="NEXT", align="C")
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
        Path(RUTA_RED).mkdir(parents=True, exist_ok=True)
        pdf.output(str(output_path))
    except Exception as e:
        log_txt(f"close_monitor: Error al guardar PDF en red ({RUTA_RED}): {e}")
        return

    try:
        open(LOG_JSONL, "w", encoding="utf-8").close()
    except Exception:
        pass

    log_txt(f"close_monitor: PDF generado: {filename} ({len(entradas)} eventos), jsonl limpiado.")


# --- Kill monitor process ---
def kill_monitor():
    names = ["outlook_sync_monitor.exe", "outlook_sync_monitor_noconsole.exe"]
    killed = False
    for name in names:
        try:
            result = subprocess.run(
                ["taskkill", "/IM", name, "/F"],
                capture_output=True, text=True, timeout=10
            )
            if result.returncode == 0:
                log_txt(f"close_monitor: Proceso '{name}' terminado.")
                killed = True
            elif "not found" not in result.stderr.lower():
                log_txt(f"close_monitor: Al terminar '{name}': {result.stderr.strip()}")
        except Exception as e:
            log_txt(f"close_monitor: Error al terminar '{name}': {e}")
    return killed


# --- Main ---
def main():
    log_txt("close_monitor: Iniciando cierre del monitor.")

    killed = kill_monitor()
    if killed:
        import time
        time.sleep(1)

    generar_pdf()
    log_txt("close_monitor: Cierre completado.")


if __name__ == "__main__":
    main()
