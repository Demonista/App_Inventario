# app.py
"""
Aplicación Flask para manejo de insumos y actualización del libro maestro.
Incluye:
- Subida de archivos (con selección de tipo de insumo)
- Listado de archivos subidos
- Integración al inventario maestro según tipo de insumo
- Historial y configuración persistidos en JSON
- Exportación de inventario a Excel y PDF
- Validaciones dinámicas sobre la hoja "General"
"""

import os
import re
import json
import pandas as pd
from datetime import datetime, date, date as date_type
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# openpyxl para manipular el maestro por hoja/columnas sin perder fórmulas
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ----------------------------
# Configuración inicial
# ----------------------------
app = Flask(__name__)
app.secret_key = "clave_secreta"  # Necesario para flash messages

UPLOAD_FOLDER = "uploads"
DATA_FOLDER = "data"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

ARCHIVOS_JSON = os.path.join(DATA_FOLDER, "archivos.json")
CONFIG_JSON = os.path.join(DATA_FOLDER, "config.json")
MAESTRO_FILE = os.path.join(DATA_FOLDER, "inventario_maestro.xlsx")

# ----------------------------
# Utilidades JSON
# ----------------------------
def cargar_json(ruta, default):
    """Carga un JSON o retorna un valor por defecto si no existe o hay error."""
    if os.path.exists(ruta):
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def guardar_json(ruta, data):
    """Guarda un JSON de forma segura con indentación."""
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# Inicialización de archivos auxiliares
archivos = cargar_json(ARCHIVOS_JSON, [])
config = cargar_json(
    CONFIG_JSON,
    {
        "autor": "Sistema",
        "version": "1.0",
        "usuarios": [],        # lista de dicts: {"nombre","correo","activo"}
        "validaciones": []     # lista de reglas: {"etiqueta","columna","operador","valor"}
    }
)

# ----------------------------
# Utilidades Excel (openpyxl)
# ----------------------------
def _ensure_sheet(wb, sheet_name):
    """Devuelve una hoja por nombre; si no existe, la crea."""
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    return wb.create_sheet(title=sheet_name)


def _headers_from_sheet(ws, header_row=1):
    """Devuelve dict {columna_normalizada: índice_columna} a partir de la fila de encabezados."""
    headers = {}
    for col_idx, cell in enumerate(ws[header_row], start=1):
        name = str(cell.value).strip() if cell.value is not None else ""
        if name:
            headers[name.lower().strip()] = col_idx
    return headers


def _normalize_series(iterable):
    """Convierte nombres de columnas a versión normalizada para comparación."""
    return [str(x).strip().lower() for x in iterable]


def update_sheet_by_headers(wb, sheet_name, df, preserve_columns=None):
    """
    Actualiza una hoja usando nombres de columna como referencia.
    - Crea la hoja si no existe.
    - Mantiene las columnas en preserve_columns.
    - Limpia datos previos en columnas que se van a sobrescribir.
    """
    if preserve_columns is None:
        preserve_columns = set()

    ws = _ensure_sheet(wb, sheet_name)
    existing_headers = _headers_from_sheet(ws)
    df_headers_norm = _normalize_series(df.columns)

    # columnas comunes entre hoja y df
    common = []
    for i, col_name in enumerate(df.columns):
        key = str(col_name).strip().lower()
        if key in existing_headers and key not in preserve_columns:
            common.append((i, key, existing_headers[key]))

    # si no hay encabezados en la hoja, inicializar con los del df
    if not existing_headers:
        for j, col_name in enumerate(df.columns, start=1):
            key = str(col_name).strip().lower()
            if key not in preserve_columns:
                ws.cell(row=1, column=j, value=str(col_name))
        existing_headers = _headers_from_sheet(ws)
        common = [(i, col.lower(), existing_headers[col.lower()])
                  for i, col in enumerate(df.columns) if col.lower() in existing_headers]

    # limpiar contenido previo desde fila 2
    for _, _, col_idx in common:
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=col_idx, value=None)

    # escribir datos nuevos
    for df_row_idx, (_, row) in enumerate(df.iterrows(), start=2):
        for i_df, _, col_idx in common:
            value = row.iloc[i_df]
            ws.cell(row=df_row_idx, column=col_idx, value=None if pd.isna(value) else value)

    return ws


def replace_sheet_content(wb, sheet_name, df):
    """Reemplaza una hoja completa (encabezados y datos)."""
    ws = _ensure_sheet(wb, sheet_name)
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None

    # escribir encabezados
    for j, col in enumerate(df.columns, start=1):
        ws.cell(row=1, column=j, value=str(col))

    # escribir filas
    for i, (_, row) in enumerate(df.iterrows(), start=2):
        for j, val in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=None if pd.isna(val) else val)

    return ws


def update_estado_gen_usuario(wb, df, tipo_evento):
    """
    Actualiza la hoja ESTADO_GEN_USUARIO con base en Talento humano (ingresos/retiros/actualización).
    """
    SHEET = "ESTADO_GEN_USUARIO"
    ws = _ensure_sheet(wb, SHEET)
    headers = _headers_from_sheet(ws)

    # helper: obtener índice de columna
    def col_idx(colname):
        return headers.get(colname.lower().strip())

    idx_ced = col_idx("cedula")
    idx_nom = col_idx("nombre")
    idx_area = col_idx("área") or col_idx("area")
    idx_estado = col_idx("estado")
    idx_fecha = col_idx("ingreso/retiro") or col_idx("ingreso_retiro")

    # inicializar encabezados si no existen
    if not headers:
        base_cols = ["CEDULA", "NOMBRE", "ÁREA", "ESTADO", "INGRESO/RETIRO"]
        for j, name in enumerate(base_cols, start=1):
            ws.cell(row=1, column=j, value=name)
        headers = _headers_from_sheet(ws)
        idx_ced, idx_nom, idx_area, idx_estado, idx_fecha = (
            headers.get("cedula"),
            headers.get("nombre"),
            headers.get("área") or headers.get("area"),
            headers.get("estado"),
            headers.get("ingreso/retiro") or headers.get("ingreso_retiro"),
        )

    df_cols = _normalize_series(df.columns)

    def get_df_val(row, candidates):
        """Devuelve valor desde df según la primera coincidencia en candidates."""
        for name in candidates:
            try:
                pos = df_cols.index(name.lower().strip())
                return row.iloc[pos]
            except ValueError:
                continue
        return None

    cedula_to_row = {}
    for r in range(2, ws.max_row + 1):
        ced_val = ws.cell(row=r, column=idx_ced).value if idx_ced else None
        if ced_val:
            cedula_to_row[str(ced_val).strip()] = r

    hoy_str = date.today().strftime("%d-%m-%Y")

    for _, row in df.iterrows():
        ced = get_df_val(row, ["cedula", "cédula", "documento", "id"])
        nom = get_df_val(row, ["nombre", "nombres", "funcionario"])
        area = get_df_val(row, ["área", "area", "dependencia", "gerencia"])
        fecha_insumo = get_df_val(row, ["ingreso/retiro", "fecha", "fecha evento"])

        if not ced or str(ced).strip() == "":
            continue
        ced_str = str(ced).strip()

        dep_str = str(area).strip() if area else ""
        if tipo_evento == "personal_retiros":
            estado_final = f"RETIRADO {dep_str}".strip()
        elif tipo_evento in ("personal_ingresos", "personal_actualizacion"):
            estado_final = f"ACTIVO {dep_str}".strip()
        else:
            estado_final = str(get_df_val(row, ["estado"])) if get_df_val(row, ["estado"]) else ""

        fecha_final = (
            str(fecha_insumo).strip()
            if fecha_insumo and str(fecha_insumo).strip()
            else hoy_str
        )

        if ced_str in cedula_to_row:
            r = cedula_to_row[ced_str]
            if idx_nom: ws.cell(row=r, column=idx_nom, value=None if pd.isna(nom) else nom)
            if idx_area: ws.cell(row=r, column=idx_area, value=None if pd.isna(area) else area)
            if idx_estado: ws.cell(row=r, column=idx_estado, value=estado_final)
            if idx_fecha: ws.cell(row=r, column=idx_fecha, value=fecha_final)
        else:
            r = ws.max_row + 1
            if idx_ced: ws.cell(row=r, column=idx_ced, value=ced_str)
            if idx_nom: ws.cell(row=r, column=idx_nom, value=None if pd.isna(nom) else nom)
            if idx_area: ws.cell(row=r, column=idx_area, value=None if pd.isna(area) else area)
            if idx_estado: ws.cell(row=r, column=idx_estado, value=estado_final)
            if idx_fecha: ws.cell(row=r, column=idx_fecha, value=fecha_final)

    return ws

# ----------------------------
# Validaciones dinámicas (hoja "General")
# ----------------------------
def parse_date_like(value):
    """Convierte value en date si es posible."""
    if value is None:
        return None
    if isinstance(value, (datetime, date_type)):
        return value.date() if isinstance(value, datetime) else value
    s = str(value).strip()
    if not s:
        return None
    patterns = ["%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%d.%m.%Y", "%Y.%m.%d"]
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        pass
    for p in patterns:
        try:
            return datetime.strptime(s, p).date()
        except Exception:
            continue
    return None


def is_number_like(v):
    """Verifica si v puede convertirse en número."""
    try:
        float(v)
        return True
    except Exception:
        return False


def eval_rule(cell_value, operador, valor):
    """
    Evalúa una regla sobre cell_value.
    Operadores soportados: =, !=, contiene, regex, >, <, >=, <=, dias_mayor_que, empty, not_empty.
    """
    operador = (operador or "").strip().lower()
    raw = cell_value

    if operador in ("empty", "es_vacio", "vacío", "vacio"):
        return raw is None or (isinstance(raw, str) and raw.strip() == "")
    if operador in ("not_empty", "no_vacio", "no vacío"):
        return not (raw is None or (isinstance(raw, str) and raw.strip() == ""))

    cell_str = "" if raw is None else str(raw).strip()
    valor_str = "" if valor is None else str(valor).strip()

    if operador in ("=", "==", "igual", "equals"):
        if is_number_like(cell_str) and is_number_like(valor_str):
            return float(cell_str) == float(valor_str)
        return cell_str == valor_str
    if operador in ("!=", "diferente", "not equals"):
        if is_number_like(cell_str) and is_number_like(valor_str):
            return float(cell_str) != float(valor_str)
        return cell_str != valor_str
    if operador in ("contiene", "contains"):
        return valor_str.lower() in cell_str.lower()
    if operador in ("no contiene", "not_contains"):
        return valor_str.lower() not in cell_str.lower()
    if operador in ("regex", "re"):
        try:
            return re.search(valor, cell_str, flags=re.IGNORECASE) is not None
        except Exception:
            return False
    if operador in (">", "<", ">=", "<="):
        if is_number_like(cell_str) and is_number_like(valor_str):
            a, b = float(cell_str), float(valor_str)
            return eval(f"a {operador} b")
        cell_date = parse_date_like(raw)
        val_date = parse_date_like(valor_str)
        if cell_date and val_date:
            return eval(f"cell_date {operador} val_date")
        return eval(f"cell_str {operador} valor_str")
    if operador in ("dias_mayor_que", "dias_mayores_que", "days_gt"):
        try:
            dias = int(float(valor_str))
        except Exception:
            return False
        cell_date = parse_date_like(raw)
        return cell_date and (date_type.today() - cell_date).days > dias

    return cell_str == valor_str

def apply_validations_to_general(wb, reglas):
    """
    Recorre la hoja 'General' y aplica reglas compuestas.
    Cada regla: {"etiqueta","logic":"AND"|"OR","conditions":[{"columna","operador","valor"}, ...]}
    Escribe el resultado (concatenación de etiquetas) en 'ANALISIS VALIDACIONES'.
    """
    ws = _ensure_sheet(wb, "General")
    headers = _headers_from_sheet(ws, header_row=1)

    key_name = "analisis validaciones"
    col_valid_idx = headers.get(key_name)
    if not col_valid_idx:
        col_valid_idx = ws.max_column + 1
        ws.cell(row=1, column=col_valid_idx, value="ANALISIS VALIDACIONES")
        headers = _headers_from_sheet(ws, header_row=1)

    # mapa normalizado de headers
    normalized_headers = {k.strip().lower(): v for k, v in headers.items()}

    # normalizar reglas (asegurar estructura esperada)
    reglas_norm = []
    for r in reglas:
        etiqueta = (r.get("etiqueta") or "").strip()
        logic = (r.get("logic") or "AND").strip().upper()
        if logic not in ("AND", "OR"):
            logic = "AND"
        conds = []
        for c in r.get("conditions", []):
            col = (c.get("columna") or "").strip().lower()
            op = (c.get("operador") or "").strip().lower()
            val = c.get("valor", "")
            if col:
                conds.append({"columna": col, "operador": op, "valor": val})
        if etiqueta and conds:
            reglas_norm.append({"etiqueta": etiqueta, "logic": logic, "conditions": conds})

    # aplicar fila por fila
    for row_idx in range(2, ws.max_row + 1):
        etiquetas_fila = []
        for regla in reglas_norm:
            # evaluar cada condición
            results = []
            for cond in regla["conditions"]:
                colkey = cond["columna"]
                # index: buscar exacto o fuzzy (like contains)
                idx = normalized_headers.get(colkey)
                if not idx:
                    # fuzzy: buscar cabezera que contenga colkey
                    matches = [v for k, v in normalized_headers.items() if colkey in k]
                    idx = matches[0] if matches else None
                if not idx:
                    results.append(False)
                    continue
                cell_val = ws.cell(row=row_idx, column=idx).value
                ok = eval_rule(cell_val, cond["operador"], cond["valor"])
                results.append(bool(ok))

            # combinar según logic
            if regla["logic"] == "AND":
                cumple = all(results) if results else False
            else:
                cumple = any(results) if results else False

            if cumple:
                etiquetas_fila.append(regla["etiqueta"])

        # escribir en columna ANALISIS VALIDACIONES (usar key exacto)
        dest_idx = headers.get(key_name) or col_valid_idx
        ws.cell(row=row_idx, column=dest_idx, value="; ".join(etiquetas_fila) if etiquetas_fila else None)

    return True



# =============================
# Rutas principales
# =============================
@app.route("/")
def index():
    """
    Página principal:
    - Carga listado de archivos subidos desde archivos.json.
    - Evita depender de variables globales "archivos" que se pueden desincronizar.
    """
    archivos = cargar_json(ARCHIVOS_JSON, [])
    if not isinstance(archivos, list):
        archivos = []  # seguridad si el JSON se daña
    return render_template("index.html", archivos=archivos)


@app.route("/upload", methods=["POST"])
def upload():
    """
    Subir uno o varios archivos con su tipo de insumo.
    - Guarda físicamente los archivos en UPLOAD_FOLDER.
    - Registra metadatos en archivos.json.
    - Registra evento en historial.json (registrar_evento).
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

        # Seguridad en el nombre
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)

        tipo = request.form.get("tipo", "desconocido")

        # Registro en archivos.json
        nuevo_archivo = {
            "nombre": filename,
            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "tipo": tipo,
            "cargado": True,
            "ruta": filepath
        }
        archivos.append(nuevo_archivo)

        # Registro en historial.json
        registrar_evento("Subida de archivo", f"{filename} como {tipo}")

    guardar_json(ARCHIVOS_JSON, archivos)
    flash("✅ Archivos subidos correctamente", "success")
    return redirect(url_for("index"))

# =============================
# Integración de insumos en el Maestro
# =============================
@app.route("/integrar", methods=["POST"])
def integrar():
    """
    Integra un archivo subido dentro del Maestro según el tipo de insumo.

    Casos:
    - Si es Maestro: se establece como archivo base (MAESTRO_FILE).
    - Si es insumo de soporte: se actualiza o reemplaza la hoja correspondiente.
    - Aplica validaciones dinámicas si existen reglas guardadas en config.json.
    - Registra cada integración en el historial.
    """
    filename = request.form.get("filename")
    tipo_insumo = request.form.get("tipo_insumo")

    # ============================
    # Validaciones iniciales
    # ============================
    if not filename:
        flash("⚠️ Debe seleccionar un archivo para integrar.", "error")
        return redirect(url_for("index"))

    filepath = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(filepath):
        flash("⚠️ El archivo no existe en el servidor.", "error")
        return redirect(url_for("index"))

    try:
        # ============================
        # Cargar insumo a integrar
        # ============================
        df_insumo = pd.read_excel(filepath, engine="openpyxl")

        # ============================
        # Caso 1: Definir Maestro inicial
        # ============================
        if tipo_insumo == "maestro":
            df_insumo.to_excel(MAESTRO_FILE, index=False, engine="openpyxl")
            flash(f"✅ El archivo {filename} se estableció como Maestro.", "success")
            registrar_evento("Integración", f"Se estableció {filename} como Maestro")

        else:
            # ============================
            # Caso 2: Integración con Maestro existente
            # ============================
            if not os.path.exists(MAESTRO_FILE):
                flash("⚠️ No existe maestro para integrar.", "error")
                return redirect(url_for("index"))

            from openpyxl import load_workbook
            wb = load_workbook(MAESTRO_FILE)

            # --- Integraciones específicas por tipo ---
            if tipo_insumo == "general":
                replace_sheet_content(wb, "General", df_insumo)

            elif tipo_insumo in ("personal_ingresos", "personal_retiros", "personal_actualizacion"):
                update_estado_gen_usuario(wb, df_insumo, tipo_insumo)

            else:
                # Caso genérico: hoja nombrada con el tipo de insumo
                hoja_destino = tipo_insumo.capitalize()
                replace_sheet_content(wb, hoja_destino, df_insumo)

            # ============================
            # Aplicar validaciones dinámicas
            # ============================
            try:
                config_data = cargar_json(CONFIG_JSON, {})
                reglas = config_data.get("validaciones", [])
                if reglas:
                    apply_validations_to_general(wb, reglas)
            except Exception as e:
                flash(f"⚠️ Error aplicando validaciones dinámicas: {e}", "warning")

            # Guardar cambios en el Maestro
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
    Aplica las reglas configuradas al Maestro en la hoja 'General'.

    Operadores soportados:
    - "="           : Igual a
    - "!="          : Diferente de
    - "contiene"    : Contiene texto
    - "no_contiene" : No contiene texto
    - ">"           : Mayor que (numérico)
    - "<"           : Menor que (numérico)
    - "es_vacio"    : Celda vacía o NaN
    - "no_vacio"    : Celda con algún valor
    """
    if not os.path.exists(MAESTRO_FILE):
        flash("⚠️ No existe maestro para aplicar validaciones", "error")
        return redirect(url_for("index"))

    try:
        # ============================
        # Cargar maestro y hoja General
        # ============================
        xl = pd.ExcelFile(MAESTRO_FILE, engine="openpyxl")
        if "General" not in xl.sheet_names:
            flash("⚠️ El Maestro no contiene la hoja 'General'", "error")
            return redirect(url_for("index"))

        df_general = xl.parse("General")

        # Asegurar columna de resultados
        if "ANÁLISIS VALIDACIONES" not in df_general.columns:
            df_general["ANÁLISIS VALIDACIONES"] = ""

        # ============================
        # Cargar reglas desde configuración
        # ============================
        config_data = cargar_json(CONFIG_JSON, {})
        reglas = config_data.get("validaciones", [])
        if not reglas:
            flash("⚠️ No hay reglas de validación configuradas", "warning")
            return redirect(url_for("configuracion_general"))

        # ============================
        # Aplicar reglas a cada fila
        # ============================
        for idx, fila in df_general.iterrows():
            resultado = []
            for regla in reglas:
                columna = regla.get("columna")
                operador = regla.get("operador", "=")
                valor = regla.get("valor", "")
                etiqueta = regla.get("etiqueta", "")

                if not columna or columna not in df_general.columns:
                    continue

                # Normalizar valor de la celda
                celda = str(fila[columna]).strip().lower() if pd.notna(fila[columna]) else ""

                try:
                    # --- Evaluación de operadores ---
                    if operador == "=" and celda == valor.strip().lower():
                        resultado.append(etiqueta)

                    elif operador == "!=" and celda != valor.strip().lower():
                        resultado.append(etiqueta)

                    elif operador == "contiene" and valor.strip().lower() in celda:
                        resultado.append(etiqueta)

                    elif operador == "no_contiene" and valor.strip().lower() not in celda:
                        resultado.append(etiqueta)

                    elif operador == ">" and pd.to_numeric(fila[columna], errors="coerce") > pd.to_numeric(valor, errors="coerce"):
                        resultado.append(etiqueta)

                    elif operador == "<" and pd.to_numeric(fila[columna], errors="coerce") < pd.to_numeric(valor, errors="coerce"):
                        resultado.append(etiqueta)

                    elif operador == "es_vacio" and (celda == "" or pd.isna(fila[columna])):
                        resultado.append(etiqueta)

                    elif operador == "no_vacio" and (celda != "" and pd.notna(fila[columna])):
                        resultado.append(etiqueta)

                except Exception:
                    # Ignorar errores de comparación (ej: texto con operador numérico)
                    continue

            # Guardar etiquetas concatenadas o "OK"
            df_general.at[idx, "ANÁLISIS VALIDACIONES"] = ", ".join(resultado) if resultado else "OK"

        # ============================
        # Guardar cambios en Maestro
        # ============================
        with pd.ExcelWriter(MAESTRO_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_general.to_excel(writer, sheet_name="General", index=False)

        flash("✅ Validaciones aplicadas correctamente al Maestro", "success")
        registrar_evento("Validaciones", "Se aplicaron validaciones al Maestro")

    except Exception as e:
        flash(f"❌ Error al aplicar validaciones: {str(e)}", "error")
        registrar_evento("Error", f"Fallo al aplicar validaciones: {e}")

    return redirect(url_for("index"))

# =============================
# Exportaciones (Excel / PDF / Maestro)
# =============================
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
    Exporta un resumen del Maestro en PDF:
    - Usa la hoja 'General' si existe, de lo contrario la primera hoja.
    - Incluye como máximo 30 filas.
    - Registra evento en historial.
    """
    if not os.path.exists(MAESTRO_FILE):
        flash("⚠️ No hay maestro disponible para exportar", "error")
        return redirect(url_for("index"))

    try:
        xl = pd.ExcelFile(MAESTRO_FILE, engine="openpyxl")
        sheet_name = "General" if "General" in xl.sheet_names else xl.sheet_names[0]
        df = xl.parse(sheet_name)

        os.makedirs(DATA_FOLDER, exist_ok=True)
        pdf_file = os.path.join(DATA_FOLDER, "inventario.pdf")

        c = canvas.Canvas(pdf_file, pagesize=letter)
        width, height = letter

        # Encabezado
        c.setFont("Helvetica-Bold", 14)
        c.drawString(50, height - 50, f"📦 Inventario Maestro - Hoja: {sheet_name}")

        # Subtítulo con fecha
        c.setFont("Helvetica", 9)
        c.drawString(50, height - 65, f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

        # Contenido (máx 30 filas)
        c.setFont("Helvetica", 8)
        y = height - 90
        for _, row in df.head(30).iterrows():
            line = " | ".join([str(v) if pd.notna(v) else "" for v in row.values])
            c.drawString(50, y, line[:150])  # truncar para no desbordar
            y -= 12
            if y < 50:
                c.showPage()
                c.setFont("Helvetica", 8)
                y = height - 50

        c.save()

        registrar_evento("Exportación", f"Se exportó el Maestro en PDF (hoja: {sheet_name})")
        return send_file(pdf_file, as_attachment=True, download_name="inventario.pdf")

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
# Historial de acciones
# =============================
@app.route("/historial")
def historial():
    """
    Muestra el historial de acciones y archivos subidos.
    - Carga eventos desde archivos.json.
    - Ordena por fecha descendente (recientes primero).
    - Maneja errores sin romper la app.
    """
    try:
        historial = cargar_json(ARCHIVOS_JSON, [])
        if not isinstance(historial, list):
            historial = []

        if historial and all(isinstance(item, dict) and "fecha" in item for item in historial):
            historial = sorted(historial, key=lambda x: x.get("fecha", ""), reverse=True)

        return render_template("historial.html", historial=historial)

    except Exception as e:
        flash(f"❌ Error al cargar historial: {str(e)}", "error")
        registrar_evento("Error", f"Fallo al cargar historial: {e}")
        return redirect(url_for("index"))

# =============================
# Configuración de usuarios y sistema
# =============================
@app.route("/configuracion")
def configuracion():
    """
    Página de configuración general (usuarios, autor, versión).
    La plantilla espera 'usuarios' para iterar.
    """
    return render_template(
        "configuracion.html",
        usuarios=config.get("usuarios", []),
        config=config
    )


@app.route("/guardar_configuracion", methods=["POST"])
def guardar_configuracion():
    """Guardar cambios en la configuración general"""
    usuarios = request.form.getlist("usuarios[]")
    correos = request.form.getlist("correos[]")
    activos = request.form.getlist("activos[]")

    config["usuarios"] = []
    for i in range(len(usuarios)):
        config["usuarios"].append({
            "nombre": usuarios[i],
            "correo": correos[i] if i < len(correos) else "",
            "activo": usuarios[i] in activos
        })

    autor = request.form.get("autor")
    version = request.form.get("version")
    if autor:
        config["autor"] = autor
    if version:
        config["version"] = version

    guardar_json(CONFIG_JSON, config)
    flash("✅ Configuración general guardada correctamente.", "success")
    return redirect(url_for("configuracion"))

def registrar_evento(accion, detalles="", usuario="Sistema"):
    """
    Registra un evento en el historial (archivos.json).
    """
    evento = {
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "accion": accion,
        "usuario": usuario,
        "detalles": detalles
    }
    historial = cargar_json(ARCHIVOS_JSON, [])
    historial.append(evento)
    guardar_json(ARCHIVOS_JSON, historial)


# =============================
# Configuración de validaciones Insumo 1
# =============================
@app.route("/configuracion_general")
def configuracion_general():
    """
    Página de configuración de reglas dinámicas para Insumo 1.
    Lee las reglas guardadas en config.json y las pasa al template.
    """
    reglas = config.get("validaciones", [])
    return render_template("configuracion_general.html", reglas=reglas)


# =============================
# Guardar reglas de validaciones (Insumo 1) - soporte condiciones multiples
# =============================
@app.route("/guardar_configuracion_general", methods=["POST"])
def guardar_configuracion_general():
    """
    Guarda reglas compuestas (varias condiciones por etiqueta).
    Espera en el form:
      - etiquetas[]                : nombre/etiqueta por regla (ordenadas)
      - logic[]                    : "AND" o "OR" por regla (misma longitud)
      - cond_col_{i}[]             : lista de columnas para la regla i (i = 0,1,2...)
      - cond_op_{i}[]              : operadores para la regla i
      - cond_val_{i}[]             : valores para la regla i
    """
    try:
        etiquetas = request.form.getlist("etiquetas[]")
        logics = request.form.getlist("logic[]")  # opcional, default "AND"

        reglas = []
        n_rules = len(etiquetas)
        for i in range(n_rules):
            etiqueta = (etiquetas[i] or "").strip()
            logic = (logics[i] or "AND").strip().upper()

            # leer las condiciones para la regla i
            prefix_col = f"cond_col_{i}[]"
            prefix_op = f"cond_op_{i}[]"
            prefix_val = f"cond_val_{i}[]"

            cols = request.form.getlist(prefix_col)
            ops = request.form.getlist(prefix_op)
            vals = request.form.getlist(prefix_val)

            conditions = []
            m = max(len(cols), len(ops), len(vals))
            for j in range(m):
                col = cols[j].strip() if j < len(cols) else ""
                op = ops[j].strip() if j < len(ops) else "="
                val = vals[j].strip() if j < len(vals) else ""
                if not col:
                    continue
                conditions.append({"columna": col, "operador": op, "valor": val})

            if etiqueta and conditions:
                reglas.append({
                    "etiqueta": etiqueta,
                    "logic": logic if logic in ("AND", "OR") else "AND",
                    "conditions": conditions
                })

        # Guardar en config y persistir
        config["validaciones"] = reglas
        guardar_json(CONFIG_JSON, config)
        flash("✅ Reglas compuestas guardadas correctamente.", "success")
    except Exception as e:
        flash(f"❌ Error al guardar reglas compuestas: {str(e)}", "error")

    return redirect(url_for("configuracion_general"))

    
# =============================
# Eliminar archivo subido
# =============================
@app.route("/eliminar/<nombre_archivo>")
def eliminar(nombre_archivo):
    """
    Elimina un archivo cargado del sistema:
    - Se borra del JSON de archivos.
    - Se elimina físicamente del directorio de uploads si existe.
    - Registra la acción en historial.
    """
    try:
        archivos = cargar_json(ARCHIVOS_JSON, [])
        if not isinstance(archivos, list):
            archivos = []

        # Filtrar lista
        archivos = [a for a in archivos if a.get("nombre") != nombre_archivo]
        guardar_json(ARCHIVOS_JSON, archivos)

        # Eliminar archivo físico
        path = os.path.join(UPLOAD_FOLDER, nombre_archivo)
        if os.path.exists(path):
            os.remove(path)
            flash(f"🗑️ Archivo {nombre_archivo} eliminado correctamente", "success")
            registrar_evento("Eliminación", f"Archivo {nombre_archivo} eliminado del sistema")
        else:
            flash(f"⚠️ El archivo {nombre_archivo} no existía en el servidor", "warning")
            registrar_evento("Advertencia", f"Intento de eliminar {nombre_archivo}, no existía en disco")

    except Exception as e:
        flash(f"❌ Error al eliminar {nombre_archivo}: {str(e)}", "error")
        registrar_evento("Error", f"Fallo al eliminar {nombre_archivo}: {e}")

    return redirect(url_for("index"))



# =============================
# Main
# =============================
if __name__ == "__main__":
    app.run(debug=True)
