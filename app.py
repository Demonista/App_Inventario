# app.py
"""
Aplicación Flask para manejo de insumos y actualización del libro maestro.
Incluye:
- Subida de archivos (con selección de tipo de insumo)
- Listado de archivos subidos
- Integración al inventario maestro según tipo de insumo
- Historial y configuración persistidos en JSON
- Exportación de inventario a Excel y PDF
"""

import os
import json
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# --- Configuración inicial ---
app = Flask(__name__)
app.secret_key = "clave_secreta"  # Necesario para flash messages

UPLOAD_FOLDER = "uploads"
DATA_FOLDER = "data"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

ARCHIVOS_JSON = os.path.join(DATA_FOLDER, "archivos.json")
CONFIG_JSON = os.path.join(DATA_FOLDER, "config.json")

# Archivo maestro
MAESTRO_FILE = os.path.join(DATA_FOLDER, "inventario_maestro.xlsx")


# --- Utilidades JSON ---
def cargar_json(ruta, default):
    """Carga un JSON o retorna un valor por defecto si no existe/corrupción"""
    if os.path.exists(ruta):
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def guardar_json(ruta, data):
    """Guarda un JSON de forma segura"""
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# Inicialización
archivos = cargar_json(ARCHIVOS_JSON, [])
config = cargar_json(CONFIG_JSON, {"autor": "Sistema", "version": "1.0"})


# --- Rutas principales ---
@app.route("/")
def index():
    """Página principal con listado de archivos subidos"""
    return render_template("index.html", archivos=archivos)


@app.route("/upload", methods=["POST"])
def upload():
    """
    Subir archivo con tipo de insumo.
    Se guarda en /uploads y se registra en archivos.json
    """
    if "archivo" not in request.files:
        flash("No se envió archivo", "error")
        return redirect(url_for("index"))

    file = request.files["archivo"]
    tipo = request.form.get("tipo")

    if not file or file.filename == "":
        flash("Archivo no válido", "error")
        return redirect(url_for("index"))

    # Guardar archivo en uploads
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    # Registro
    nuevo_archivo = {
        "nombre": file.filename,
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "tipo": tipo,
        "cargado": True
    }
    archivos.append(nuevo_archivo)
    guardar_json(ARCHIVOS_JSON, archivos)

    flash(f"Archivo {file.filename} subido como {tipo}", "success")
    return redirect(url_for("index"))


@app.route("/integrar", methods=["POST"])
def integrar():
    """
    Integra un archivo subido en el maestro según tipo de insumo.
    - Si es 'maestro', reemplaza el maestro actual.
    - Si es otro insumo, se hace merge con el maestro.
    """
    filename = request.form.get("filename")
    tipo_insumo = request.form.get("tipo_insumo")

    if not filename:
        flash("Debe seleccionar un archivo para integrar", "error")
        return redirect(url_for("index"))

    filepath = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(filepath):
        flash("El archivo no existe en el servidor", "error")
        return redirect(url_for("index"))

    try:
        # Leer archivo subido
        df_insumo = pd.read_excel(filepath)

        if tipo_insumo == "maestro":
            # Guardar como maestro directamente
            df_insumo.to_excel(MAESTRO_FILE, index=False)
            flash(f"El archivo {filename} se estableció como Maestro", "success")

        else:
            if not os.path.exists(MAESTRO_FILE):
                flash("No existe maestro para integrar", "error")
                return redirect(url_for("index"))

            # Cargar maestro existente
            df_maestro = pd.read_excel(MAESTRO_FILE)

            # TODO: aquí se pueden programar reglas de negocio específicas
            # Ejemplo básico → concatenar
            df_resultado = pd.concat([df_maestro, df_insumo], ignore_index=True)

            # Guardar de nuevo el maestro
            df_resultado.to_excel(MAESTRO_FILE, index=False)
            flash(f"Archivo {filename} integrado como {tipo_insumo}", "success")

    except Exception as e:
        flash(f"Error al integrar: {str(e)}", "error")

    return redirect(url_for("index"))


@app.route("/exportar-excel")
def exportar_excel():
    """Exporta el maestro actual a Excel"""
    if not os.path.exists(MAESTRO_FILE):
        flash("No hay maestro disponible para exportar", "error")
        return redirect(url_for("index"))
    return send_file(MAESTRO_FILE, as_attachment=True, download_name="inventario.xlsx")


@app.route("/exportar-pdf")
def exportar_pdf():
    """Exporta el maestro a PDF con ReportLab"""
    if not os.path.exists(MAESTRO_FILE):
        flash("No hay maestro disponible para exportar", "error")
        return redirect(url_for("index"))

    df = pd.read_excel(MAESTRO_FILE)

    pdf_file = os.path.join(DATA_FOLDER, "inventario.pdf")
    c = canvas.Canvas(pdf_file, pagesize=letter)
    width, height = letter

    # Título
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 50, "Inventario Maestro")

    # Imprimir primeras filas
    c.setFont("Helvetica", 10)
    y = height - 80
    for i, row in df.head(30).iterrows():
        line = " | ".join([str(v) for v in row.values])
        c.drawString(50, y, line[:120])  # recorte por ancho de página
        y -= 15
        if y < 50:
            c.showPage()
            y = height - 50

    c.save()

    return send_file(pdf_file, as_attachment=True, download_name="inventario.pdf")


@app.route("/download-maestro")
def download_maestro():
    """Descarga directa del archivo maestro"""
    if not os.path.exists(MAESTRO_FILE):
        flash("No hay maestro disponible para descargar", "error")
        return redirect(url_for("index"))
    return send_file(MAESTRO_FILE, as_attachment=True, download_name="inventario_maestro.xlsx")


@app.route("/historial")
def historial():
    """Historial basado en los archivos cargados"""
    return render_template("historial.html", archivos=archivos)


@app.route("/configuracion")
def configuracion():
    """Página de configuración"""
    return render_template("config.html", config=config)


@app.route("/eliminar/<nombre_archivo>")
def eliminar(nombre_archivo):
    """Eliminar un archivo subido"""
    global archivos
    archivos = [a for a in archivos if a["nombre"] != nombre_archivo]
    guardar_json(ARCHIVOS_JSON, archivos)

    path = os.path.join(UPLOAD_FOLDER, nombre_archivo)
    if os.path.exists(path):
        os.remove(path)

    flash(f"Archivo {nombre_archivo} eliminado", "success")
    return redirect(url_for("index"))


# --- Main ---
if __name__ == "__main__":
    app.run(debug=True)

