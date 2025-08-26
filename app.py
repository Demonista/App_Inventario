# app.py
"""
Aplicación Flask para manejo de insumos y actualización del libro Maestro.

Incluye:
- subida de archivos (múltiple)
- listado de archivos subidos
- integración de insumos por tipo (1 a 5)
- historial de operaciones (paginado)
- configuración general editable
- exportación del Maestro
- eliminación de archivos subidos
"""

import os
import json
import pandas as pd
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file, flash

# -------------------
# Configuración inicial
# -------------------
app = Flask(__name__)
app.secret_key = "supersecret"  # Necesario para mensajes flash

UPLOAD_FOLDER = "uploads"
CONFIG_FILE = "config.json"
HISTORIAL_FILE = "historial.json"
MAESTRO_FILE = "static/Maestro.xlsx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs("static", exist_ok=True)

# -------------------
# Funciones auxiliares
# -------------------
def cargar_config():
    """Carga configuración desde archivo JSON."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}
    return {}

def guardar_config(config):
    """Guarda configuración en archivo JSON."""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

def cargar_historial():
    """Carga historial de archivos subidos."""
    if os.path.exists(HISTORIAL_FILE):
        try:
            with open(HISTORIAL_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return []
    return []

def guardar_historial(historial):
    """Guarda historial en JSON."""
    with open(HISTORIAL_FILE, "w", encoding="utf-8") as f:
        json.dump(historial, f, indent=4, ensure_ascii=False)

def agregar_a_historial(nombre, tipo="desconocido"):
    """Agrega un archivo al historial."""
    historial = cargar_historial()
    historial.append({
        "nombre": nombre,
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "tipo": tipo,
        "cargado": True
    })
    guardar_historial(historial)

# -------------------
# Procesadores de insumos
# -------------------
def procesar_insumo_1(path_excel, maestro):
    """
    Insumo 1 (Inventario Maestro).
    - Es la base del libro maestro.
    - No se reemplaza, se mantiene como hoja original.
    """
    df_dict = pd.read_excel(path_excel, sheet_name=None)
    maestro.update(df_dict)  # Cargamos todas las hojas
    return maestro

def procesar_insumo_2(path_excel, maestro):
    """Insumo 2 (Antivirus)."""
    df = pd.read_excel(path_excel)
    maestro["Antivirus"] = df
    return maestro

def procesar_insumo_3(path_excel, maestro):
    """Insumo 3 (Personal)."""
    df = pd.read_excel(path_excel)
    maestro["ESTADO_GEN_USUARIO"] = df
    return maestro

def procesar_insumo_4(path_excel, maestro):
    """Insumo 4 (TMP)."""
    df = pd.read_excel(path_excel)
    maestro["Useraranda_BLOGIK"] = df
    return maestro

def procesar_insumo_5(path_excel, maestro):
    """Insumo 5 (Directorio Activo)."""
    df = pd.read_excel(path_excel)
    maestro["Reporte DA"] = df
    return maestro

# Mapeo de insumos con procesadores
PROCESADORES = {
    "inventario": procesar_insumo_1,
    "endpoint": procesar_insumo_2,
    "personal": procesar_insumo_3,
    "tmp": procesar_insumo_4,
    "da": procesar_insumo_5,
}

# -------------------
# Rutas principales
# -------------------
@app.route("/", methods=["GET", "POST"])
def index():
    """Pantalla principal: subida y listado de archivos."""
    if request.method == "POST":
        if "files[]" not in request.files:
            flash("No seleccionaste archivos")
            return redirect(url_for("index"))

        files = request.files.getlist("files[]")
        for file in files:
            if file and file.filename:
                save_path = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(save_path)
                agregar_a_historial(file.filename)

        flash("Archivos subidos correctamente")
        return redirect(url_for("index"))

    historial = cargar_historial()
    return render_template("index.html", archivos=historial)

@app.route("/integrar", methods=["POST"])
def integrar():
    """Integración de un insumo al Maestro."""
    filename = request.form.get("filename")
    tipo_insumo = request.form.get("tipo_insumo")

    if not filename or not tipo_insumo:
        flash("Faltan datos para la integración")
        return redirect(url_for("index"))

    path_excel = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(path_excel):
        flash("El archivo no existe")
        return redirect(url_for("index"))

    # Si no existe el maestro, crear uno nuevo vacío
    if os.path.exists(MAESTRO_FILE):
        maestro = pd.read_excel(MAESTRO_FILE, sheet_name=None)
    else:
        maestro = {}

    # Procesar insumo según tipo
    if tipo_insumo in PROCESADORES:
        maestro = PROCESADORES[tipo_insumo](path_excel, maestro)
    else:
        flash("Tipo de insumo desconocido")
        return redirect(url_for("index"))

    # Guardar el Maestro actualizado
    with pd.ExcelWriter(MAESTRO_FILE, engine="openpyxl", mode="w") as writer:
        for hoja, df in maestro.items():
            if not df.empty:
                df.to_excel(writer, sheet_name=hoja, index=False)

    flash(f"Archivo {filename} integrado al Maestro como insumo {tipo_insumo}")
    return redirect(url_for("index"))

@app.route("/download_maestro")
def download_maestro():
    """Descargar el archivo Maestro."""
    if os.path.exists(MAESTRO_FILE):
        return send_file(MAESTRO_FILE, as_attachment=True)
    flash("No hay Maestro generado todavía.")
    return redirect(url_for("index"))

@app.route("/eliminar/<nombre_archivo>")
def eliminar(nombre_archivo):
    """Eliminar archivo subido y actualizar historial."""
    path = os.path.join(UPLOAD_FOLDER, nombre_archivo)
    if os.path.exists(path):
        os.remove(path)

    historial = cargar_historial()
    historial = [h for h in historial if h["nombre"] != nombre_archivo]
    guardar_historial(historial)

    flash(f"Archivo {nombre_archivo} eliminado.")
    return redirect(url_for("index"))

@app.route("/historial")
def historial():
    """Vista paginada de historial de archivos."""
    page = int(request.args.get("page", 1))
    per_page = 5
    historial = cargar_historial()
    total_pages = max(1, (len(historial) + per_page - 1) // per_page)

    start = (page - 1) * per_page
    end = start + per_page
    archivos_paginados = historial[start:end]

    return render_template("historial.html",
                           archivos=archivos_paginados,
                           page=page,
                           total_pages=total_pages)

@app.route("/configuracion", methods=["GET", "POST"])
def configuracion():
    """Vista de configuración general."""
    config = cargar_config()

    if request.method == "POST":
        config["empresa_default"] = request.form.get("empresa_default", "")
        config["usar_fecha_archivo"] = "usar_fecha_archivo" in request.form
        config["fecha_formato"] = request.form.get("fecha_formato", "%Y-%m-%d")
        guardar_config(config)
        flash("Configuración guardada correctamente.")
        return redirect(url_for("configuracion"))

    return render_template("configuracion_general.html", config=config)

# -------------------
# Run
# -------------------
if __name__ == "__main__":
    app.run(debug=True)
