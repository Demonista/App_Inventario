# app.py
"""
Aplicación Flask para manejo de insumos y actualización del libro maestro.

Incluye:
- Subida de archivos (múltiple)
- Listado de archivos subidos
- Integración por tipo de insumo (endpoint, personnel, tmp, da)
- Integración múltiple en batch
- Historial de acciones (persistido en JSON separado de archivos)
- Configuración de reglas de validación
- Exportación de resultados a Excel y PDF
"""

import os
import json
from datetime import datetime

from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify, flash
import pandas as pd
import openpyxl
from werkzeug.utils import secure_filename   # ✅ Import corregido

# =====================================================
# Configuración básica
# =====================================================

app = Flask(__name__)

# Carpetas principales
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")
STATIC_FOLDER = os.path.join(BASE_DIR, "static")   # solo CSS/JS/imagenes
DATA_FOLDER = os.path.join(BASE_DIR, "data")       # solo datos persistentes

# Archivos persistentes (guardados en /data)
ARCHIVOS_JSON = os.path.join(DATA_FOLDER, "archivos.json")    # Archivos subidos
HISTORIAL_JSON = os.path.join(DATA_FOLDER, "historial.json")  # Eventos y acciones
CONFIG_JSON   = os.path.join(DATA_FOLDER, "config.json")      # Configuración general

# Asegurar directorios
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(STATIC_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

# =====================================================
# Utilidades de persistencia
# =====================================================

def cargar_json(path, default):
    """Carga un archivo JSON o devuelve un valor por defecto."""
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def guardar_json(path, data):
    """Guarda un diccionario/lista en un archivo JSON."""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


# =====================================================
# Historial y registro de eventos
# =====================================================

def registrar_evento(evento, detalle=""):
    """
    Guarda en historial.json un registro con:
    - fecha
    - evento (acción realizada)
    - detalle (extra)
    """
    historial = cargar_json(HISTORIAL_JSON, [])
    historial.append({
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "evento": evento,
        "detalle": detalle
    })
    guardar_json(HISTORIAL_JSON, historial)


# =====================================================
# Manejo de archivos subidos
# =====================================================

def registrar_archivo(filename, tipo_insumo):
    """
    Registra un archivo subido en archivos.json con:
    - nombre
    - tipo_insumo (ej: maestro, endpoint, da, etc.)
    - fecha
    """
    archivos = cargar_json(ARCHIVOS_JSON, [])
    archivos.append({
        "nombre": filename,
        "tipo_insumo": tipo_insumo,
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })
    guardar_json(ARCHIVOS_JSON, archivos)
    registrar_evento("Archivo subido", f"{filename} ({tipo_insumo})")


def listar_archivos():
    """Devuelve la lista de archivos subidos registrada en archivos.json."""
    return cargar_json(ARCHIVOS_JSON, [])


# =====================================================
# Validaciones dinámicas
# =====================================================
def aplicar_reglas(df, reglas):
    """
    Aplica un conjunto de reglas sobre un DataFrame de pandas.
    Cada regla tiene:
      - etiqueta (str)
      - condiciones: lista de dicts {columna, operador, valor}
      - logic: "AND" | "OR"  (opcional, default AND)
      - hoja (opcional) -> no se usa aquí (se aplica sobre el df que se le pase)
      - columna_destino (opcional) -> si se indica, la etiqueta se escribe en esa columna
    """

    # Columna por defecto donde se escriben resultados si no se especifica en la regla
    default_dest = "ANÁLISIS VALIDACIONES"
    if default_dest not in df.columns:
        df[default_dest] = ""

    # Procesar cada regla
    for regla in reglas:
        etiqueta = regla.get("etiqueta", "") or regla.get("label", "")
        condiciones = regla.get("conditions") or regla.get("condiciones") or []
        logic = (regla.get("logic") or regla.get("operador_global") or "AND").upper()
        dest_col = regla.get("columna_destino") or regla.get("destino") or default_dest

        # Asegurar existencia de columna destino
        if dest_col not in df.columns:
            # crear columna vacía si no existe
            df[dest_col] = ""

        if not condiciones:
            continue

        masks = []
        for cond in condiciones:
            col = cond.get("columna") or cond.get("col") or ""
            op = (cond.get("operador") or cond.get("op") or "").strip()
            val = cond.get("valor", "")

            if not col or col not in df.columns:
                # intentar match flexible (buscar encabezado que contenga el texto normalizado)
                candidates = [c for c in df.columns if col.lower().strip() in str(c).lower()]
                if candidates:
                    col = candidates[0]
                else:
                    # columna no encontrada -> máscara all False
                    masks.append(pd.Series([False] * len(df)))
                    continue

            serie = df[col]

            if op == "=":
                mask = serie.astype(str).fillna("").str.strip() == str(val).strip()
            elif op == "!=":
                mask = serie.astype(str).fillna("").str.strip() != str(val).strip()
            elif op == "contiene":
                mask = serie.astype(str).str.contains(str(val), case=False, na=False)
            elif op == "no_contiene":
                mask = ~serie.astype(str).str.contains(str(val), case=False, na=False)
            elif op == ">":
                mask = pd.to_numeric(serie, errors="coerce") > pd.to_numeric(val, errors="coerce")
            elif op == "<":
                mask = pd.to_numeric(serie, errors="coerce") < pd.to_numeric(val, errors="coerce")
            elif op in ("es_vacio", "empty"):
                mask = serie.isna() | (serie.astype(str).str.strip() == "")
            elif op in ("no_vacio", "not_empty"):
                mask = ~(serie.isna() | (serie.astype(str).str.strip() == ""))
            else:
                # operador desconocido => falsa
                mask = pd.Series([False] * len(df))

            masks.append(mask)

        # combinar máscaras
        if not masks:
            continue

        if logic == "AND":
            final_mask = masks[0]
            for m in masks[1:]:
                final_mask &= m
        else:  # OR
            final_mask = masks[0]
            for m in masks[1:]:
                final_mask |= m

        # Escribir etiqueta en la columna destino (concatenar si ya hay texto)
        # si ya existe un valor, lo concatenamos separando por "; " y evitando duplicados
        existing = df.loc[final_mask, dest_col].fillna("").astype(str)
        new_vals = []
        for i, old in existing.iteritems():
            parts = [p.strip() for p in old.split(";") if p.strip()] if old else []
            if etiqueta and etiqueta not in parts:
                parts.append(etiqueta)
            new_vals.append("; ".join(parts))

        df.loc[final_mask, dest_col] = new_vals

    return df


# =============================
# Configuración de Maestro
# =============================
MAESTRO_FILE = os.path.join(OUTPUT_FOLDER, "maestro.xlsx")

# ----------------------------
# Utilidad: Leer hojas y columnas del Maestro
# ----------------------------
def get_maestro_sheets_and_columns(maestro_path=None):
    """
    Devuelve (hojas, columnas) donde:
      - hojas: lista de nombres de hoja existentes en el maestro (ordenada)
      - columnas: lista de nombres de columnas en la hoja 'General' si existe,
                  o en la primera hoja disponible.

    Si no existe maestro, retorna ([], []).
    """
    maestro_path = maestro_path or MAESTRO_FILE
    if not os.path.exists(maestro_path):
        return [], []

    try:
        xl = pd.ExcelFile(maestro_path, engine="openpyxl")
        hojas = xl.sheet_names or []
        # preferir 'General' si existe
        target = "General" if "General" in hojas else (hojas[0] if hojas else None)
        columnas = []
        if target:
            df = xl.parse(target, nrows=0)  # solo encabezados
            columnas = list(df.columns)
        return hojas, columnas
    except Exception:
        return [], []


# =============================
# Rutas principales
# =============================
@app.route("/")
def index():
    """
    Página principal:
    - Carga listado de archivos subidos desde archivos.json.
    - Lee dinámicamente hojas y columnas del Maestro.
    - Seguridad: si el JSON está dañado → retorna lista vacía.
    """
    archivos = cargar_json(ARCHIVOS_JSON, [])
    if not isinstance(archivos, list):
        archivos = []

    # Obtener hojas y columnas dinámicamente del Maestro
    hojas_disponibles, columnas_disponibles = get_maestro_sheets_and_columns()

    return render_template(
        "index.html",
        archivos=archivos,
        hojas_disponibles=hojas_disponibles,
        columnas_disponibles=columnas_disponibles
    )

@app.route("/upload", methods=["POST"])
def upload():
    """
    Subida de uno o varios archivos.
    - Guarda archivos en UPLOAD_FOLDER.
    - Registra metadatos en archivos.json.
    - Registra evento en historial.json.
    """
    if "files" not in request.files:
        flash("⚠️ No se enviaron archivos", "error")
        return redirect(url_for("index"))

    files = request.files.getlist("files")
    if not files:
        flash("⚠️ No se seleccionaron archivos", "error")
        return redirect(url_for("index"))

    archivos = cargar_json(ARCHIVOS_JSON, [])
    if not isinstance(archivos, list):
        archivos = []

    for file in files:
        if not file or file.filename.strip() == "":
            continue

        # Normalizar nombre y guardar
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)

        tipo_insumo = request.form.get("tipo_insumo", "desconocido")

        # Registrar en archivos.json
        nuevo_archivo = {
            "nombre": filename,
            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "tipo_insumo": tipo_insumo,
            "cargado": True,
            "ruta": filepath
        }
        archivos.append(nuevo_archivo)

        # Registrar en historial.json
        registrar_evento("Subida de archivo", f"{filename} como {tipo_insumo}")

    guardar_json(ARCHIVOS_JSON, archivos)
    flash("✅ Archivos subidos correctamente", "success")
    return redirect(url_for("index"))

@app.route("/eliminar/<nombre_archivo>", methods=["POST", "GET"])
def eliminar(nombre_archivo):
    """
    Elimina un archivo subido del sistema (uploads + registro en JSON).
    """
    try:
        archivo_path = os.path.join(UPLOAD_FOLDER, nombre_archivo)
        if os.path.exists(archivo_path):
            os.remove(archivo_path)

        # Eliminar del JSON de archivos
        if os.path.exists(ARCHIVOS_JSON):
            with open(ARCHIVOS_JSON, "r", encoding="utf-8") as f:
                archivos = json.load(f)
            archivos = [a for a in archivos if a.get("nombre") != nombre_archivo]
            with open(ARCHIVOS_JSON, "w", encoding="utf-8") as f:
                json.dump(archivos, f, indent=4, ensure_ascii=False)

        registrar_evento("Eliminación", f"Se eliminó el archivo {nombre_archivo}")
        flash(f"✅ Archivo {nombre_archivo} eliminado con éxito", "success")

    except Exception as e:
        flash(f"❌ Error al eliminar archivo: {str(e)}", "error")
        registrar_evento("Error", f"Fallo al eliminar archivo {nombre_archivo}: {e}")

    return redirect(url_for("index"))


# =============================
# Integración de insumos en el Maestro
# =============================
@app.route("/integrar", methods=["POST"])
def integrar():
    """
    Integra un archivo subido dentro del Maestro según el tipo de insumo.
    """
    filename = request.form.get("filename")
    tipo_insumo = request.form.get("tipo_insumo")

    if not filename:
        flash("⚠️ Debe seleccionar un archivo para integrar.", "error")
        return redirect(url_for("index"))

    filepath = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(filepath):
        flash("⚠️ El archivo no existe en el servidor.", "error")
        return redirect(url_for("index"))

    try:
        # Cargar insumo
        df_insumo = pd.read_excel(filepath, engine="openpyxl")

        # Caso 1: Definir Maestro inicial
        if tipo_insumo == "maestro":
            df_insumo.to_excel(MAESTRO_FILE, index=False, engine="openpyxl")
            flash(f"✅ El archivo {filename} se estableció como Maestro.", "success")
            registrar_evento("Integración", f"Se estableció {filename} como Maestro")

        else:
            # Caso 2: Integrar en Maestro existente
            if not os.path.exists(MAESTRO_FILE):
                flash("⚠️ No existe maestro para integrar.", "error")
                return redirect(url_for("index"))

            from openpyxl import load_workbook
            wb = load_workbook(MAESTRO_FILE)

            # Integraciones específicas
            if tipo_insumo == "general":
                replace_sheet_content(wb, "General", df_insumo)

            elif tipo_insumo in ("personal_ingresos", "personal_retiros", "personal_actualizacion"):
                update_estado_gen_usuario(wb, df_insumo, tipo_insumo)

            else:
                hoja_destino = tipo_insumo.capitalize()
                replace_sheet_content(wb, hoja_destino, df_insumo)

            # Aplicar validaciones dinámicas
            try:
                config_data = cargar_json(CONFIG_JSON, {})
                reglas = config_data.get("validaciones", [])
                if reglas:
                    apply_validations_to_general(wb, reglas)
            except Exception as e:
                flash(f"⚠️ Error aplicando validaciones dinámicas: {e}", "warning")

            # Guardar Maestro actualizado
            wb.save(MAESTRO_FILE)
            registrar_evento("Integración", f"{filename} integrado como {tipo_insumo}")

        flash(f"✅ Archivo {filename} integrado como {tipo_insumo}.", "success")

    except Exception as e:
        flash(f"❌ Error al integrar {filename}: {str(e)}", "error")
        registrar_evento("Error", f"Fallo al integrar {filename}: {e}")

    return redirect(url_for("index"))


# =============================
# Aplicar validaciones al Maestro
# =============================
@app.route("/aplicar-validaciones", methods=["POST"])
def aplicar_validaciones():
    """
    Aplica todas las reglas configuradas al Maestro.
    - Si una regla especifica 'hoja', se aplica sobre esa hoja.
    - Si no especifica, por defecto se aplica sobre 'General'.
    - Las etiquetas se escriben en la columna indicada por 'columna_destino',
      o en "ANÁLISIS VALIDACIONES" si no se indica.
    - Se sobrescribe cada hoja afectada manteniendo las demás intactas.
    """
    if not os.path.exists(MAESTRO_FILE):
        flash("⚠️ No existe maestro para aplicar validaciones", "error")
        return redirect(url_for("index"))

    try:
        # 1. Cargar reglas desde config
        config_data = cargar_json(CONFIG_JSON, {})
        reglas = config_data.get("validaciones", [])
        if not reglas:
            flash("⚠️ No hay reglas de validación configuradas", "warning")
            return redirect(url_for("configuracion_general"))

        # 2. Abrir maestro y extraer todas las hojas
        xl = pd.ExcelFile(MAESTRO_FILE, engine="openpyxl")
        sheet_names = xl.sheet_names

        # 3. Agrupar reglas por hoja objetivo
        rules_by_sheet = {}
        for r in reglas:
            sheet_target = r.get("hoja") or "General"
            rules_by_sheet.setdefault(sheet_target, []).append(r)

        # 4. Procesar hoja por hoja
        for sheet_name, rules_list in rules_by_sheet.items():
            if sheet_name not in sheet_names:
                # La hoja indicada no existe, la ignoramos
                continue

            df_sheet = xl.parse(sheet_name)

            # Asegurar columna por defecto
            if "ANÁLISIS VALIDACIONES" not in df_sheet.columns:
                df_sheet["ANÁLISIS VALIDACIONES"] = ""

            # Aplicar reglas sobre el DataFrame
            df_sheet = aplicar_reglas(df_sheet, rules_list)

            # Sobrescribir hoja en el Excel
            with pd.ExcelWriter(MAESTRO_FILE, engine="openpyxl",
                                mode="a", if_sheet_exists="replace") as writer:
                df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

        # 5. Notificar éxito
        flash("✅ Validaciones aplicadas correctamente al Maestro", "success")
        registrar_evento("Validaciones", "Se aplicaron validaciones al Maestro")

    except Exception as e:
        flash(f"❌ Error al aplicar validaciones: {str(e)}", "error")
        registrar_evento("Error", f"Fallo al aplicar validaciones: {e}")

    return redirect(url_for("index"))

# =============================
# Exportaciones (Excel / PDF / Maestro)
# =============================
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

@app.route("/exportar-excel")
def exportar_excel():
    """
    Exporta el archivo maestro completo en formato Excel (.xlsx).
    """
    if not os.path.exists(MAESTRO_FILE):
        flash("⚠️ No hay maestro disponible para exportar", "error")
        return redirect(url_for("index"))

    try:
        registrar_evento("Exportación", "Se exportó el Maestro en Excel")
        return send_file(
            MAESTRO_FILE,
            as_attachment=True,
            download_name="inventario.xlsx"
        )
    except Exception as e:
        flash(f"❌ Error al exportar a Excel: {str(e)}", "error")
        registrar_evento("Error", f"Fallo al exportar Excel: {e}")
        return redirect(url_for("index"))


@app.route("/exportar-pdf")
def exportar_pdf():
    """
    Exporta el Maestro en PDF.
    Opciones:
    - ?hoja=NombreHoja → exporta una hoja específica.
    - Si no se indica hoja: exporta 'General' (si existe) o la primera hoja.
    - ?all=true → exporta todas las filas.
    - Por defecto → exporta solo las primeras 30 filas.
    """
    if not os.path.exists(MAESTRO_FILE):
        flash("⚠️ No hay maestro disponible para exportar", "error")
        return redirect(url_for("index"))

    try:
        # ============================
        # Determinar hoja a exportar
        # ============================
        hoja_param = request.args.get("hoja", "").strip()
        xl = pd.ExcelFile(MAESTRO_FILE, engine="openpyxl")

        if hoja_param and hoja_param in xl.sheet_names:
            sheet_name = hoja_param
        elif "General" in xl.sheet_names:
            sheet_name = "General"
        else:
            sheet_name = xl.sheet_names[0]

        df = xl.parse(sheet_name)

        # ============================
        # Determinar si exportar todo o 30 filas
        # ============================
        export_all = request.args.get("all", "false").lower() == "true"
        if not export_all:
            df = df.head(30)

        os.makedirs(DATA_FOLDER, exist_ok=True)
        pdf_file = os.path.join(DATA_FOLDER, f"inventario_{sheet_name}.pdf")

        # ============================
        # Crear PDF con ReportLab
        # ============================
        doc = SimpleDocTemplate(pdf_file, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = []

        # Título
        elements.append(Paragraph(f"📦 Inventario Maestro - Hoja: {sheet_name}", styles["Heading1"]))
        elements.append(Paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", styles["Normal"]))

        # Convertir DataFrame a tabla
        data = [df.columns.tolist()] + df.values.tolist()
        table = Table(data, repeatRows=1)

        # Estilos de tabla
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B6FA4")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
        ]))

        elements.append(table)
        doc.build(elements)

        registrar_evento(
            "Exportación",
            f"Se exportó el Maestro en PDF (hoja: {sheet_name}, filas: {'todas' if export_all else '30'})"
        )
        return send_file(pdf_file, as_attachment=True, download_name=f"inventario_{sheet_name}.pdf")

    except Exception as e:
        flash(f"❌ Error al exportar a PDF: {str(e)}", "error")
        registrar_evento("Error", f"Fallo al exportar PDF: {e}")
        return redirect(url_for("index"))


@app.route("/download-maestro")
def download_maestro():
    """
    Descarga directa del Maestro en Excel.
    """
    if not os.path.exists(MAESTRO_FILE):
        flash("⚠️ No hay maestro disponible para descargar", "error")
        return redirect(url_for("index"))

    try:
        registrar_evento("Descarga", "Se descargó el Maestro en Excel")
        return send_file(
            MAESTRO_FILE,
            as_attachment=True,
            download_name="inventario_maestro.xlsx"
        )
    except Exception as e:
        flash(f"❌ Error al descargar el Maestro: {str(e)}", "error")
        registrar_evento("Error", f"Fallo al descargar Maestro: {e}")
        return redirect(url_for("index"))

# =============================
# Historial de eventos
# =============================
@app.route("/historial")
def historial():
    """
    Muestra el historial de acciones realizadas (subidas, integraciones, exportaciones, errores).
    """
    try:
        eventos = cargar_json(HISTORIAL_JSON, [])
        return render_template("historial.html", eventos=eventos)
    except Exception as e:
        flash(f"❌ Error al cargar historial: {str(e)}", "error")
        registrar_evento("Error", f"Fallo al mostrar historial: {e}")
        return redirect(url_for("index"))

# =====================================================
# Configuración de usuarios
# =====================================================
@app.route("/configuracion", methods=["GET", "POST"])
def configuracion():
    """
    Página de configuración de usuarios y opciones generales.
    """
    # Aquí cargas o guardas datos de configuración de usuarios
    try:
        if request.method == "POST":
            # Procesar datos enviados del formulario
            # Ejemplo: guardar usuarios en un JSON
            usuarios = request.form.getlist("usuarios")
            with open(CONFIG_JSON, "w", encoding="utf-8") as f:
                json.dump({"usuarios": usuarios}, f, indent=4, ensure_ascii=False)
            return redirect(url_for("configuracion"))

        # Si es GET → cargar configuración actual
        if os.path.exists(CONFIG_JSON):
            with open(CONFIG_JSON, "r", encoding="utf-8") as f:
                config_data = json.load(f)
        else:
            config_data = {"usuarios": []}

        return render_template("configuracion.html", config=config_data)

    except Exception as e:
        return f"Error en configuración de usuarios: {e}", 500



# =============================
# Configuración general (Insumo 1)
# =============================
@app.route("/configuracion_general", methods=["GET", "POST"])
def configuracion_general():
    """
    Página para configurar reglas de validación (Insumo 1).
    GET: entrega la plantilla con hojas y columnas detectadas en el Maestro.
    POST: puede usarse para guardar un JSON simple (la lógica real de guardado
          está en la ruta /guardar_configuracion_general).
    """
    # cargar configuración existente (si hay)
    config_data = cargar_json(CONFIG_JSON, {})
    reglas_guardadas = config_data.get("validaciones", [])

    # obtener hojas y columnas del maestro (para pre-poblar selects)
    hojas_disponibles, columnas_disponibles = get_maestro_sheets_and_columns()

    # si la plantilla hace POST aquí, podríamos manejarlo; pero preferimos delegar
    # al endpoint específico para guardar reglas (guardar_configuracion_general)
    if request.method == "POST":
        flash("Use el botón Guardar en el formulario de reglas para persistir.", "info")

    return render_template(
        "configuracion_general.html",
        config=config_data,
        reglas=reglas_guardadas,
        hojas_disponibles=hojas_disponibles,
        columnas_disponibles=columnas_disponibles
    )


@app.route("/guardar_configuracion_general", methods=["POST"])
def guardar_configuracion_general():
    """
    Guarda reglas compuestas (varias condiciones por etiqueta), enviadas desde
    configuracion_general.html. Soporta el formulario con campos:
      - etiquetas[]                   (por regla)
      - hojas[]                       (por regla)  --> hoja del maestro a la que aplica
      - columnas_destino[]            (por regla)  --> columna que recibirá la etiqueta
      - operadores_globales[]         (por regla)  --> AND/OR
      - condiciones[i][columna][]     (por regla i)
      - condiciones[i][operador][]    (por regla i)
      - condiciones[i][valor][]       (por regla i)
    """
    try:
        etiquetas = request.form.getlist("etiquetas[]")
        hojas = request.form.getlist("hojas[]")
        columnas_destino = request.form.getlist("columnas_destino[]")
        operadores_globales = request.form.getlist("operadores_globales[]")

        reglas = []
        n = max(len(etiquetas), len(hojas), len(columnas_destino), len(operadores_globales))

        for i in range(n):
            etiqueta = (etiquetas[i] if i < len(etiquetas) else "").strip()
            hoja = (hojas[i] if i < len(hojas) else "").strip()
            col_dest = (columnas_destino[i] if i < len(columnas_destino) else "").strip()
            logic = (operadores_globales[i] if i < len(operadores_globales) else "AND").strip().upper()
            if logic not in ("AND", "OR"):
                logic = "AND"

            # leer condiciones de la regla i usando la nomenclatura del form
            pref_col = f"condiciones[{i}][columna][]"
            pref_op = f"condiciones[{i}][operador][]"
            pref_val = f"condiciones[{i}][valor][]"

            cols = request.form.getlist(pref_col)
            ops = request.form.getlist(pref_op)
            vals = request.form.getlist(pref_val)

            condiciones = []
            m = max(len(cols), len(ops), len(vals))
            for j in range(m):
                c = cols[j].strip() if j < len(cols) else ""
                o = ops[j].strip() if j < len(ops) else ""
                v = vals[j].strip() if j < len(vals) else ""
                if not c:
                    continue
                condiciones.append({"columna": c, "operador": o, "valor": v})

            if etiqueta and condiciones:
                reglas.append({
                    "etiqueta": etiqueta,
                    "hoja": hoja,
                    "columna_destino": col_dest,
                    "logic": logic,
                    "conditions": condiciones
                })

        # Guardar en config.json
        config_data = cargar_json(CONFIG_JSON, {})
        config_data["validaciones"] = reglas
        guardar_json(CONFIG_JSON, config_data)

        flash("✅ Reglas compuestas guardadas correctamente.", "success")
        registrar_evento("Configuración general", f"Guardadas {len(reglas)} reglas")
    except Exception as e:
        flash(f"❌ Error al guardar reglas compuestas: {e}", "error")
        registrar_evento("Error", f"Fallo al guardar reglas compuestas: {e}")

    return redirect(url_for("configuracion_general"))


# =============================
# Punto de entrada principal
# =============================
if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(DATA_FOLDER, exist_ok=True)
    app.run(debug=True)
