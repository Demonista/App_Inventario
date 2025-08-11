# app.py
"""
Aplicación Flask para manejo de insumos y actualización del libro maestro.
Incluye:
- subida de archivos (múltiple)
- listado de archivos subidos
- integración por tipo de insumo (endpoint, personnel, tmp, da)
- integración múltiple en batch
- historial y configuración (persistidos en JSON)
- descarga del maestro y rutas de export (Excel/PDF)
"""

import os
import json
import threading
import subprocess
import platform
from datetime import datetime
from flask import (Flask, render_template, request, redirect, url_for,
                   flash, send_file, jsonify)
from werkzeug.utils import secure_filename

# Importar utilidades XLSX (asegúrate de tener xlsx_utils.py en la misma carpeta)
from xlsx_utils import (
    integrate_endpoint_to_antivirus,
    integrate_personnel_to_estado,
    integrate_tmp_to_useraranda,
    integrate_da_to_reporte,
    replace_sheet_with_df,
    backup_file
)

# ----------------------------- Configuración -----------------------------
app = Flask(__name__)
app.secret_key = "cambiar_esta_clave_por_una_segura"  # necesario para flash

ROOT = os.path.abspath(os.getcwd())
UPLOAD_FOLDER = os.path.join(ROOT, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {"xls", "xlsx"}
HISTORY_FILE = os.path.join(UPLOAD_FOLDER, "history.json")
CONFIG_FILE = os.path.join(UPLOAD_FOLDER, "config.json")

# Lista en memoria (se reconstruye al iniciar con archivos del disco)
archivos_cargados = []

# ----------------------------- Helpers ----------------------------------
def is_allowed(filename):
    """Valida extensión permitida."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def load_history():
    """Carga historial desde JSON (si existe)."""
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_history(history):
    """Guarda historial en JSON."""
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

def append_history(entry):
    """Añade una entrada al historial (archivo)."""
    hist = load_history()
    hist.insert(0, entry)  # insert al inicio para ver lo más reciente primero
    save_history(hist)

def load_config():
    """Carga configuración desde JSON (si existe)."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_config(cfg):
    """Guarda configuración en JSON."""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

def find_master_file():
    """
    Busca el archivo maestro en uploads por patrón aproximado:
    se asume que empieza con 3 letras mes + 2 dígitos (p. ej. ago06) y contiene 'Reporte-Data_Inventario'
    Si no lo encuentra, devuelve None.
    """
    for fname in os.listdir(app.config["UPLOAD_FOLDER"]):
        low = fname.lower()
        if "reporte-data_inventario" in low or "reporte-data inventario" in low or "reporte-data-inventario" in low:
            return os.path.join(app.config["UPLOAD_FOLDER"], fname)
    # fallback: buscar el archivo con mayor tamaño que contenga 'inventario'
    candidates = [f for f in os.listdir(app.config["UPLOAD_FOLDER"]) if "inventario" in f.lower()]
    if candidates:
        # devolver el más reciente
        candidates.sort(key=lambda n: os.path.getmtime(os.path.join(app.config["UPLOAD_FOLDER"], n)), reverse=True)
        return os.path.join(app.config["UPLOAD_FOLDER"], candidates[0])
    return None

# ----------------------------- Rutas ------------------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    """
    Página principal:
    - GET: muestra archivos subidos y formularios
    - POST: sube archivos (múltiple) y actualiza lista
    """
    global archivos_cargados
    # reconstruir lista desde uploads para no depender solo de la variable en memoria
    archivos_cargados = []
    for fname in sorted(os.listdir(app.config["UPLOAD_FOLDER"]), reverse=True):
        path = os.path.join(app.config["UPLOAD_FOLDER"], fname)
        if os.path.isfile(path) and is_allowed(fname):
            archivos_cargados.append({
                "nombre": fname,
                "fecha": datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S"),
                "tipo": detectar_tipo_insumo(fname := fname),  # reutiliza tu función de detección (definida abajo)
                "cargado": True
            })

    if request.method == "POST":
        files = request.files.getlist("files[]")
        uploaded = []
        for f in files:
            if f and is_allowed(f.filename):
                filename = secure_filename(f.filename)
                dest = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                f.save(dest)
                uploaded.append(filename)
                # guardar en historial
                append_history({
                    "accion": "upload",
                    "archivo": filename,
                    "timestamp": datetime.now().isoformat()
                })
        if uploaded:
            flash(f"Se subieron {len(uploaded)} archivo(s): {', '.join(uploaded)}", "success")
        else:
            flash("No se subió ningún archivo válido (extensiones permitidas: xls, xlsx).", "warning")
        return redirect(url_for("index"))

    # render
    config = load_config()
    return render_template("index.html", archivos=archivos_cargados, config=config)

@app.route("/eliminar/<nombre_archivo>")
def eliminar(nombre_archivo):
    """Elimina archivo físico en uploads y actualiza historial."""
    path = os.path.join(app.config["UPLOAD_FOLDER"], nombre_archivo)
    if os.path.exists(path):
        try:
            os.remove(path)
            append_history({
                "accion": "delete",
                "archivo": nombre_archivo,
                "timestamp": datetime.now().isoformat()
            })
            flash(f"Archivo {nombre_archivo} eliminado.", "success")
        except Exception as e:
            flash(f"Error eliminando {nombre_archivo}: {e}", "error")
    else:
        flash("Archivo no encontrado.", "warning")
    return redirect(url_for("index"))

# ----------------- Ruta para integrar un insumo seleccionado -----------------
@app.route("/integrar", methods=["POST"])
def integrar():
    """
    Integra un archivo subido al maestro. Form data:
      - tipo_insumo: 'endpoint'|'personnel'|'tmp'|'da'
      - filename: nombre del archivo (de uploads)
      - keep_rows: opcional (int)
    """
    tipo = request.form.get("tipo_insumo")
    filename = request.form.get("filename")
    keep_rows = int(request.form.get("keep_rows", 2))
    if not tipo or not filename:
        flash("Debes seleccionar tipo de insumo y archivo.", "error")
        return redirect(url_for("index"))

    src = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(filename))
    if not os.path.exists(src):
        flash("Archivo fuente no encontrado en uploads.", "error")
        return redirect(url_for("index"))

    # encontrar maestro
    maestro = find_master_file()
    if maestro is None:
        flash("No se encontró el archivo maestro (Reporte-Data_Inventario). Colócalo en uploads.", "error")
        return redirect(url_for("index"))

    # ejecutar integración según tipo
    try:
        if tipo == "endpoint":
            res = integrate_endpoint_to_antivirus(maestro, src, keep_rows=keep_rows)
            flash(f"Integración Antivirus completada. Filas: {res.get('rows_written')}", "success")
            append_history({"accion": "integrar", "tipo": "endpoint", "archivo": filename, "timestamp": datetime.now().isoformat()})
        elif tipo == "personnel":
            res = integrate_personnel_to_estado(maestro, src, keep_rows=keep_rows)
            flash(f"Integración Personal completada. Añadidos: {res.get('added')}, Actualizados: {res.get('updated')}", "success")
            append_history({"accion": "integrar", "tipo": "personnel", "archivo": filename, "timestamp": datetime.now().isoformat()})
        elif tipo == "tmp":
            res = integrate_tmp_to_useraranda(maestro, src, keep_rows=keep_rows)
            flash(f"Integración Useraranda_BLOGIK completada. Filas: {res.get('rows_written')}", "success")
            append_history({"accion": "integrar", "tipo": "tmp", "archivo": filename, "timestamp": datetime.now().isoformat()})
        elif tipo == "da":
            res = integrate_da_to_reporte(maestro, src, keep_rows=keep_rows)
            flash(f"Integración Reporte DA completada. Filas: {res.get('rows_written')}", "success")
            append_history({"accion": "integrar", "tipo": "da", "archivo": filename, "timestamp": datetime.now().isoformat()})
        else:
            flash("Tipo de insumo desconocido.", "error")
    except Exception as e:
        flash(f"Error durante integración: {e}", "error")

    return redirect(url_for("index"))

# ----------------- Integrar múltiples insumos (batch) -----------------------
@app.route("/integrar-multiples", methods=["POST"])
def integrar_multiples():
    """
    Recibe listas tipo_insumo[] y filename[] que emparejan por índice.
    Ejecuta integración secuencial y devuelve resumen.
    """
    tipos = request.form.getlist("tipo_insumo[]")
    files = request.form.getlist("filename[]")
    keep_rows = int(request.form.get("keep_rows", 2))
    maestro = find_master_file()
    if maestro is None:
        flash("No se encontró el archivo maestro (Reporte-Data_Inventario).", "error")
        return redirect(url_for("index"))

    results = []
    for t, f in zip(tipos, files):
        src = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(f))
        if not os.path.exists(src):
            results.append((f, {"error": "archivo fuente no encontrado"}))
            continue
        try:
            if t == "endpoint":
                res = integrate_endpoint_to_antivirus(maestro, src, keep_rows=keep_rows)
            elif t == "personnel":
                res = integrate_personnel_to_estado(maestro, src, keep_rows=keep_rows)
            elif t == "tmp":
                res = integrate_tmp_to_useraranda(maestro, src, keep_rows=keep_rows)
            elif t == "da":
                res = integrate_da_to_reporte(maestro, src, keep_rows=keep_rows)
            else:
                res = {"error": "tipo desconocido"}
            results.append((f, res))
            append_history({"accion": "integrar", "tipo": t, "archivo": f, "timestamp": datetime.now().isoformat()})
        except Exception as e:
            results.append((f, {"error": str(e)}))

    # mostrar mensajes
    for fname, res in results:
        if res and "error" in res:
            flash(f"{fname}: ERROR - {res['error']}", "error")
        else:
            flash(f"{fname}: Integración OK", "success")
    return redirect(url_for("index"))

# ----------------- Descargar maestro -------------------------
@app.route("/download-maestro")
def download_maestro():
    maestro = find_master_file()
    if not maestro or not os.path.exists(maestro):
        flash("No se encontró maestro.xlsx en uploads.", "error")
        return redirect(url_for("index"))
    return send_file(maestro, as_attachment=True)

# ----------------- Exportar (rutas que usan scripts.js) --------------
@app.route("/exportar-excel")
def exportar_excel():
    """
    Ruta simple que devuelve el maestro.xlsx para descargar como 'inventario.xlsx'
    (Puedes reemplazarla por una rutina que genere un workbook temporal)
    """
    maestro = find_master_file()
    if not maestro:
        return ("No hay maestro para exportar", 404)
    # enviamos el archivo directamente
    return send_file(maestro, as_attachment=True, download_name="inventario.xlsx")

@app.route("/exportar-pdf")
def exportar_pdf():
    """
    Ruta que convierte el maestro a PDF y lo devuelve.
    Implementación sencilla: si ya tienes una forma de exportar PDF, cámbiala aquí.
    Por simplicidad en este ejemplo enviamos el archivo maestro renombrado .pdf (no es conversión real).
    Recomendación: usar libreoffice (soffice) o reportlab para conversión.
    """
    maestro = find_master_file()
    if not maestro:
        return ("No hay maestro para exportar", 404)

    # Intento simple: si tienes libreoffice instalado puedes convertir:
    # soffice --headless --convert-to pdf --outdir /tmp maestro.xlsx
    # Para no depender de instalaciones, devolvemos el maestro con extensión .pdf (NO CONVERTIDO)
    return send_file(maestro, as_attachment=True, download_name="inventario.pdf")

# ----------------- Historial & Configuración -------------------------
@app.route("/historial")
def historial():
    """Muestra el historial de acciones (subidas, integraciones, eliminaciones)."""
    history = load_history()
    return render_template("historial.html", history=history)

@app.route("/configuracion", methods=["GET", "POST"])
def configuracion():
    """
    Página para manejar la configuración de integración (mapeos por empresa, reglas, etc).
    Guardamos/recuperamos en config.json dentro de uploads/.
    """
    cfg = load_config()
    if request.method == "POST":
        # ejemplo simple: guardar mapeos o reglas que envíe el formulario
        cfg["empresa_default"] = request.form.get("empresa_default", cfg.get("empresa_default", ""))
        cfg["usar_fecha_archivo"] = bool(request.form.get("usar_fecha_archivo", ""))
        cfg["fecha_formato"] = request.form.get("fecha_formato", cfg.get("fecha_formato", "%Y-%m-%d"))
        save_config(cfg)
        flash("Configuración guardada.", "success")
        append_history({"accion": "config_update", "timestamp": datetime.now().isoformat()})
        return redirect(url_for("configuracion"))
    return render_template("configuracion.html", config=cfg)

# ----------------- Abrir navegador al iniciar (opcional) --------------
def abrir_navegador():
    url = "http://127.0.0.1:5000"
    if platform.system() == "Windows":
        subprocess.Popen(["start", "", url], shell=True)
    else:
        try:
            subprocess.Popen(["xdg-open", url])
        except Exception:
            pass

# ----------------- Detección tipo insumo (tu función original) ---------
def detectar_tipo_insumo(nombre):
    """Función de detección textual para asignar tipo a los archivos subidos."""
    nombre = nombre.lower()
    if "inventario" in nombre or "insumo1" in nombre:
        return "Inventario Proveedor"
    elif "endpoint" in nombre or "antivirus" in nombre or "insumo2" in nombre:
        return "Antivirus"
    elif "retiros" in nombre or "personal" in nombre or "ingresos" in nombre or "colaboradores" in nombre or "insumo3" in nombre:
        return "Talento Humano"
    elif "directorio" in nombre or "insumo5" in nombre:
        return "Directorio Activo"
    elif "aranda" in nombre or "insumo4" in nombre or "tmp" in nombre:
        return "Aranda"
    elif "da" in nombre or "reporte da" in nombre:
        return "Reporte DA"
    return "Desconocido"

# ----------------------------- Main -------------------------------------
if __name__ == "__main__":
    # abrir navegador 1s después (opcional)
    threading.Timer(1.0, abrir_navegador).start()
    app.run(debug=False)
