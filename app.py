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
from typing import Optional
from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify, flash
import pandas as pd
import openpyxl
from werkzeug.utils import secure_filename   # ✅ Import corregido
# 📂 Importar funciones utilitarias desde xlsx_utils.py
from xlsx_utils import (
    integrate_personnel_to_estado,
    replace_sheet_with_df,
    get_maestro_sheets_and_columns
)


# =====================================================
# Configuración básica
# =====================================================

app = Flask(__name__)

# Clave secreta para sesiones y flash()
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "supersecreto_inventario_123")


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
def get_maestro_sheets_and_columns(maestro_path: Optional[str] = None):
    """
    Devuelve (hojas_list, hojas_columnas_map).
    - hojas_list: lista de nombres de hoja del archivo maestro (ordenada).
    - hojas_columnas_map: dict { hoja_nombre: [col1, col2, ...] } con las columnas
      detectadas en cada hoja (intenta leer solo los encabezados).
    - Si el archivo no existe o ocurre un error, retorna ([], {}).
    """
    maestro_path = maestro_path or MAESTRO_FILE
    if not os.path.exists(maestro_path):
        return [], {}

    try:
        xl = pd.ExcelFile(maestro_path, engine="openpyxl")
        hojas = xl.sheet_names or []
        hojas_columnas_map: Dict[str, List[str]] = {}

        for hoja in hojas:
            try:
                # Leer solo encabezado (nrows=0) para obtener columnas
                df_head = xl.parse(hoja, nrows=0)
                hojas_columnas_map[hoja] = list(df_head.columns)
            except Exception:
                hojas_columnas_map[hoja] = []

        return hojas, hojas_columnas_map

    except Exception:
        return [], {}
# =============================
# Rutas principales
# =============================
@app.route("/")
def index():
    """
    Página principal:
    - Carga listado de archivos subidos desde archivos.json.
    - Pasa info de hojas/columnas del maestro para JS (si existe).
    """
    archivos_raw = cargar_json(ARCHIVOS_JSON, [])
    if not isinstance(archivos_raw, list):
        archivos_raw = []

    # Normalizar estructura de cada registro para que las plantillas siempre
    # encuentren las keys: nombre, tipo, fecha, ruta
    archivos = []
    for a in archivos_raw:
        if not isinstance(a, dict):
            continue
        nombre = a.get("nombre") or a.get("filename") or a.get("file") or ""
        # permitir tanto 'tipo' como 'tipo_insumo'
        tipo = a.get("tipo") or a.get("tipo_insumo") or a.get("type") or ""
        fecha = a.get("fecha") or a.get("date") or ""
        ruta = a.get("ruta") or a.get("path") or a.get("filepath") or ""
        archivos.append({
            "nombre": nombre,
            "tipo": tipo,
            "fecha": fecha,
            "ruta": ruta
        })

    # Obtener hojas y columnas (si existe maestro)
    try:
        from xlsx_utils import get_maestro_sheets_and_columns
        hojas_disponibles, hojas_columnas_map = get_maestro_sheets_and_columns(MAESTRO_FILE)
        # columnas por defecto: preferir "General", si no la primera hoja con columnas
        columnas_disponibles = []
        if hojas_columnas_map:
            columnas_disponibles = hojas_columnas_map.get("General") or next(
                (cols for cols in hojas_columnas_map.values() if cols), []
            )
    except Exception:
        hojas_disponibles, hojas_columnas_map, columnas_disponibles = [], {}, []

    return render_template(
        "index.html",
        archivos=archivos,
        hojas_disponibles=hojas_disponibles,
        columnas_disponibles=columnas_disponibles,
        hojas_columnas_map=hojas_columnas_map
    )

@app.route("/upload", methods=["POST"])
def upload():
    """
    Subida de uno o varios archivos.
    - Soporta input name="archivo" (único) y name="files" (múltiples).
    - Normaliza y guarda metadatos en ARCHIVOS_JSON asegurando que existan
      tanto 'tipo' como 'tipo_insumo' para compatibilidad con templates.
    """
    # obtener lista de archivos (soporta ambos casos)
    files = []
    # caso: formulario con input file name="archivo" (single-file)
    single = request.files.get("archivo")
    if single and single.filename and single.filename.strip():
        files = [single]
    else:
        # caso: formulario multi-file name="files"
        files = request.files.getlist("files") or []

    if not files:
        flash("⚠️ No se enviaron archivos", "error")
        return redirect(url_for("index"))

    archivos = cargar_json(ARCHIVOS_JSON, [])
    if not isinstance(archivos, list):
        archivos = []

    tipo_insumo_from_form = request.form.get("tipo") or request.form.get("tipo_insumo") or "desconocido"

    for file in files:
        if not file or not getattr(file, "filename", "").strip():
            continue

        # Normalizar nombre y guardar
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)

        # Guardar ambos keys para compatibilidad con diferentes templates
        nuevo_archivo = {
            "nombre": filename,
            "tipo": tipo_insumo_from_form,
            "tipo_insumo": tipo_insumo_from_form,
            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "ruta": filepath,
            "cargado": True
        }
        archivos.append(nuevo_archivo)

        # Registrar en historial
        registrar_evento("Subida de archivo", f"{filename} como {tipo_insumo_from_form}")

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


@app.route("/integrar", methods=["POST"])
def integrar():
    """
    Integra un archivo subido dentro del Maestro según el tipo de insumo.
    Lógica robusta:
      - acepta campos filename / archivo / file como nombre de archivo a integrar
      - acepta tipo_insumo o tipo (según el form)
      - intenta mapear el tipo a una hoja existente (comparación robusta)
      - para 'personal' llama integrate_personnel_to_estado(MAESTRO_FILE, filepath, operacion=...)
      - para otras hojas usa replace_sheet_with_df con keep_rows configurado por hoja
    """
    # Aceptar varios nombres de campo posibles desde el form
    filename = request.form.get("filename") or request.form.get("archivo") or request.form.get("file")
    tipo_insumo = (request.form.get("tipo_insumo") or request.form.get("tipo") or "").strip()

    if not filename:
        flash("⚠️ Debe seleccionar un archivo para integrar.", "error")
        return redirect(url_for("index"))

    filepath = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(filepath):
        flash("⚠️ El archivo no existe en el servidor.", "error")
        return redirect(url_for("index"))

    try:
        # Leer insumo (si hace falta)
        # Nota: para personal pasamos la ruta al método incremental (no usar el DataFrame aquí)
        df_insumo = None
        if not tipo_insumo.lower().startswith("personal"):
            # solo cargar cuando lo necesitaremos como DataFrame
            try:
                df_insumo = pd.read_excel(filepath, engine="openpyxl")
            except Exception as e:
                # si no puede leerse con pandas, reportar
                raise RuntimeError(f"Error leyendo el insumo con pandas: {e}")

        # Caso 1: Definir Maestro inicial (sobrescribe todo)
        if tipo_insumo == "maestro":
            # sobrescribir completamente el archivo maestro
            if df_insumo is None:
                df_insumo = pd.read_excel(filepath, engine="openpyxl")
            df_insumo.to_excel(MAESTRO_FILE, index=False, engine="openpyxl")
            registrar_evento("Integración", f"Se estableció {filename} como Maestro")
            flash(f"✅ El archivo {filename} se estableció como Maestro.", "success")
            return redirect(url_for("index"))

        # Asegurar existencia del maestro
        if not os.path.exists(MAESTRO_FILE):
            flash("⚠️ No existe maestro para integrar.", "error")
            return redirect(url_for("index"))

        # Obtener hojas disponibles y mapa de columnas
        hojas, hojas_columnas_map = get_maestro_sheets_and_columns(MAESTRO_FILE)
        hojas_set = hojas or []

        # Helper: normalizar texto para matching robusto
        def _norm_key(s):
            if not s:
                return ""
            return re.sub(r'[^a-z0-9]', '', str(s).lower())

        # Intentar encontrar la hoja destino de forma flexible
        def _find_matching_sheet(candidate: str, filename_hint: str = ""):
            cand = _norm_key(candidate)
            # mapa normalizado -> original
            map_norm = { _norm_key(h): h for h in hojas_set }
            if cand in map_norm:
                return map_norm[cand]
            # tratar candidato vacío: intentar con filename hint
            if not cand and filename_hint:
                cand = _norm_key(filename_hint)
                if cand in map_norm:
                    return map_norm[cand]
            # keywords heurísticos
            keywords = {
                "antivirus": ["antivirus", "endpoint", "enpoint"],
                "estado": ["estado", "usuario", "personal", "colaborador", "cedula"],
                "useraranda_blogik": ["aranda", "useraranda", "blogik"],
                "reporte_da": ["da", "reporte", "reporte da"],
                "general": ["general"]
            }
            for target, keys in keywords.items():
                for k in keys:
                    if k in cand or k in _norm_key(filename_hint):
                        # buscar hoja que contenga la keyword k
                        for h in hojas_set:
                            if k in _norm_key(h):
                                return h
                        # mapeo directo si existe
                        mapping = {
                            "antivirus": "Antivirus",
                            "estado": "ESTADO_GEN_USUARIO",
                            "useraranda_blogik": "Useraranda_BLOGIK",
                            "reporte_da": "Reporte DA",
                            "general": "General"
                        }
                        mapped = mapping.get(target)
                        if mapped and mapped in hojas_set:
                            return mapped
            # último recurso: buscar substring en cualquier hoja
            for h in hojas_set:
                if _norm_key(cand) and (_norm_key(cand) in _norm_key(h) or _norm_key(h) in _norm_key(cand)):
                    return h
            return None

        # Si es personal -> ruta incremental (no reemplazamos hoja entera)
        if tipo_insumo.lower().startswith("personal") or tipo_insumo.lower().startswith("personal_"):
            # deducir operacion: ingresos / retiros / mixto
            oper = "mixto"
            low = tipo_insumo.lower()
            if "ingres" in low:
                oper = "ingresos"
            elif "retiro" in low or "retiros" in low or "termin" in low:
                oper = "retiros"

            resumen = integrate_personnel_to_estado(
                master_path=MAESTRO_FILE,
                personnel_path=filepath,
                operacion=oper,
                make_backup=True
            )
            registrar_evento("Integración", f"{filename} integrado como {tipo_insumo} (personal) -> {resumen}")
            flash(f"✅ Personal integrado: {resumen}", "success")
            return redirect(url_for("index"))

        # Para el resto: buscar hoja destino
        hoja_dest = _find_matching_sheet(tipo_insumo, filename)
        if not hoja_dest:
            # si no encontramos, informar y listar opciones disponibles
            texto = (
                f"La hoja destino para '{tipo_insumo or filename}' no fue encontrada en el Maestro. "
                f"Hojas disponibles: {', '.join(hojas_set) if hojas_set else 'ninguna'}"
            )
            flash(f"❌ {texto}", "error")
            registrar_evento("Error", texto)
            return redirect(url_for("index"))

        # Determinar keep_rows por hoja (para no borrar la "fila plantilla" de fórmulas)
        default_keep = {
            "ESTADO_GEN_USUARIO": 2,
            "Antivirus": 2,
            "Useraranda_BLOGIK": 2,
            "Reporte DA": 2,
            "General": 1
        }
        keep_rows = default_keep.get(hoja_dest, 2)

        # Reemplazar hoja con DataFrame (replace_sheet_with_df espera ruta maestro y df)
        try:
            res = replace_sheet_with_df(MAESTRO_FILE, hoja_dest, df_insumo, keep_rows=keep_rows)
            registrar_evento("Integración", f"{filename} integrado en hoja {hoja_dest} ({res.get('rows_written')} filas).")
            flash(f"✅ Archivo {filename} integrado en hoja '{hoja_dest}'.", "success")
        except ValueError as e:
            # hoja no existe o error específico
            flash(f"❌ Error al integrar: {e}", "error")
            registrar_evento("Error", f"Fallo integrando {filename}: {e}")
        except Exception as e:
            flash(f"❌ Error inesperado al integrar: {e}", "error")
            registrar_evento("Error", f"Fallo integrando {filename}: {e}")

        # Aplicar validaciones dinámicas (opcional, si existen reglas)
        try:
            config_data = cargar_json(CONFIG_JSON, {})
            reglas = config_data.get("validaciones", []) or []
            if reglas and "General" in hojas_set:
                xl = pd.ExcelFile(MAESTRO_FILE, engine="openpyxl")
                df_general = xl.parse("General")
                df_general = aplicar_reglas(df_general, reglas)
                with pd.ExcelWriter(MAESTRO_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_general.to_excel(writer, sheet_name="General", index=False)
        except Exception as e:
            flash(f"⚠️ Error aplicando validaciones dinámicas: {e}", "warning")

        registrar_evento("Integración", f"{filename} integrado como {tipo_insumo}")
        return redirect(url_for("index"))

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
    POST: solo muestra mensaje (el guardado real está en /guardar_configuracion_general).
    """
    # cargar configuración existente (si hay)
    config_data = cargar_json(CONFIG_JSON, {})
    reglas_guardadas = config_data.get("validaciones", [])

    # obtener hojas y columnas del maestro (para pre-poblar selects)
    try:
        # import local para evitar import circulares si los hay
        from xlsx_utils import get_maestro_sheets_and_columns
        hojas_disponibles, hojas_columnas_map = get_maestro_sheets_and_columns(MAESTRO_FILE)
    except Exception:
        hojas_disponibles, hojas_columnas_map = [], {}

    # columnas por defecto: preferir "General", si no la primera hoja con columnas
    columnas_disponibles = []
    if hojas_columnas_map:
        columnas_disponibles = hojas_columnas_map.get("General") or next(
            (cols for cols in hojas_columnas_map.values() if cols), []
        )

    if request.method == "POST":
        flash("Use el botón Guardar en el formulario de reglas para persistir.", "info")

    return render_template(
        "configuracion_general.html",
        config=config_data,
        reglas=reglas_guardadas,
        hojas_disponibles=hojas_disponibles,
        columnas_disponibles=columnas_disponibles,
        hojas_columnas_map=hojas_columnas_map
    )

@app.route("/guardar-configuracion-general", methods=["POST"])
def guardar_configuracion_general():
    data = request.form.to_dict()
    with open(CONFIG_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

    flash("Configuración guardada correctamente", "success")
    return redirect(url_for("configuracion_general"))


# =============================
# Punto de entrada principal
# =============================
if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(DATA_FOLDER, exist_ok=True)
    app.run(debug=True)
