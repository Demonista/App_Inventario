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
from datetime import datetime, date
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Para manipular el maestro por hoja/columnas sin perder fórmulas
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# --- Configuración inicial ---
app = Flask(__name__)
app.secret_key = "clave_secreta"  # Necesario para flash messages

UPLOAD_FOLDER = "uploads"
DATA_FOLDER = "data"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

ARCHIVOS_JSON = os.path.join(DATA_FOLDER, "archivos.json")
CONFIG_JSON = os.path.join(DATA_FOLDER, "config.json")

# Archivo maestro (base del proveedor)
MAESTRO_FILE = os.path.join(DATA_FOLDER, "inventario_maestro.xlsx")


# =============================
# Utilidades JSON
# =============================
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
# Estructura de config por defecto compatible con tus plantillas
config = cargar_json(
    CONFIG_JSON,
    {
        "autor": "Sistema",
        "version": "1.0",
        "usuarios": [],                 # lista de dicts: {"nombre","correo","activo"}
        "validaciones": []              # lista de reglas: {"etiqueta","columna","operador","valor"}
    }
)


# =============================
# Utilidades de Excel (openpyxl)
# =============================
def _ensure_sheet(wb, sheet_name):
    """Devuelve una hoja por nombre; si no existe, la crea al final."""
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    return wb.create_sheet(title=sheet_name)


def _headers_from_sheet(ws, header_row=1):
    """
    Devuelve un dict {nombre_columna_normalizado: index_columna_1based}
    leyendo la fila de encabezado. Normaliza espacios y case.
    """
    headers = {}
    for col_idx, cell in enumerate(ws[header_row], start=1):
        name = str(cell.value).strip() if cell.value is not None else ""
        key = name.lower().strip()
        if key:
            headers[key] = col_idx
    return headers


def _normalize_series(iterable):
    """Convierte los valores a str normalizados (para comparar headers)."""
    return [str(x).strip().lower() for x in iterable]


def update_sheet_by_headers(wb, sheet_name, df, preserve_columns=None):
    """
    Actualiza una hoja usando nombres de columna como referencia.
    - Crea la hoja si no existe.
    - Lee encabezados existentes y solo escribe en las columnas que coinciden por nombre.
    - Si 'preserve_columns' (set de nombres normalizados) está definido,
      NO toca esas columnas (ni encabezado ni celdas).
    - Limpia datos previos (desde fila 2) en las columnas que va a escribir.
    - Escribe desde fila 2 hacia abajo.
    """
    if preserve_columns is None:
        preserve_columns = set()

    ws = _ensure_sheet(wb, sheet_name)

    # map headers existentes en la hoja
    existing_headers = _headers_from_sheet(ws, header_row=1)

    # map headers del df
    df_headers_norm = _normalize_series(df.columns)

    # Calcular columnas comunes (por nombre normalizado)
    common = []
    for i, col_name in enumerate(df.columns):
        key = str(col_name).strip().lower()
        if key in existing_headers and key not in preserve_columns:
            common.append((i, key, existing_headers[key]))

    # Si no hay encabezados comunes, intentamos escribir encabezados del df (sin romper preservadas)
    if not existing_headers:
        # Escribimos encabezados del df en la fila 1, respetando preserve_columns (por nombre)
        for j, col_name in enumerate(df.columns, start=1):
            key = str(col_name).strip().lower()
            if key in preserve_columns:
                continue  # no sobrescribir encabezados preservados
            ws.cell(row=1, column=j, value=str(col_name))
        # recalcular headers
        existing_headers = _headers_from_sheet(ws, header_row=1)
        # volver a armar common
        common = []
        for i, col_name in enumerate(df.columns):
            key = str(col_name).strip().lower()
            if key in existing_headers and key not in preserve_columns:
                common.append((i, key, existing_headers[key]))

    # Limpiar contenido previo desde fila 2 SOLO en las columnas que vamos a escribir
    max_row = ws.max_row
    for _, _, col_idx in common:
        for r in range(2, max_row + 1):
            ws.cell(row=r, column=col_idx, value=None)

    # Escribir datos (fila 2 en adelante) en columnas mapeadas
    for df_row_idx, (_, row) in enumerate(df.iterrows(), start=2):
        for i_df, key, col_idx in common:
            value = row.iloc[i_df]
            ws.cell(row=df_row_idx, column=col_idx, value=None if pd.isna(value) else value)

    return ws


def replace_sheet_content(wb, sheet_name, df):
    """
    Reemplaza el contenido de una hoja por completo (encabezados + datos).
    Se usa cuando se puede sobrescribir sin riesgo.
    """
    ws = _ensure_sheet(wb, sheet_name)
    # Borrar todo
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None

    # Escribir encabezados
    for j, col in enumerate(df.columns, start=1):
        ws.cell(row=1, column=j, value=str(col))

    # Escribir datos
    for i, (_, row) in enumerate(df.iterrows(), start=2):
        for j, val in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=None if pd.isna(val) else val)

    return ws


def update_estado_gen_usuario(wb, df, tipo_evento):
    """
    Actualiza la hoja ESTADO_GEN_USUARIO con base en df de Talento humano.
    - Busca por CEDULA (str/int).
    - Si existe: actualiza NOMBRE, ÁREA y ESTADO según 'tipo_evento'.
    - Si no existe: crea fila.
    - La columna INGRESO/RETIRO se toma del df si existe, si no usa la fecha actual (dd-mm-YYYY).
    - ESTADO = "ACTIVO {DEP}" para ingresos/actualización si hay dependencia; "RETIRADO {DEP}" para retiros.
      (Si no hay dependencia, se deja "ACTIVO" o "RETIRADO" a secas).
    Columnas objetivo (si existen en el maestro):
        "CEDULA", "NOMBRE", "ÁREA", "ESTADO", "INGRESO/RETIRO"
    """
    SHEET = "ESTADO_GEN_USUARIO"
    ws = _ensure_sheet(wb, SHEET)

    headers = _headers_from_sheet(ws, header_row=1)
    # Mapeos por nombre normalizado
    def col_idx(colname):
        return headers.get(colname.lower().strip())

    idx_ced = col_idx("cedula")
    idx_nom = col_idx("nombre")
    idx_area = col_idx("área") or col_idx("area")  # por si acaso sin tilde
    idx_estado = col_idx("estado")
    idx_fecha = col_idx("ingreso/retiro") or col_idx("ingreso_retiro")

    # Si la hoja estuviera vacía, la inicializamos con encabezados estándar
    if not headers:
        base_cols = ["CEDULA", "NOMBRE", "ÁREA", "ESTADO", "INGRESO/RETIRO"]
        for j, name in enumerate(base_cols, start=1):
            ws.cell(row=1, column=j, value=name)
        headers = _headers_from_sheet(ws, header_row=1)
        idx_ced = headers.get("cedula")
        idx_nom = headers.get("nombre")
        idx_area = headers.get("área") or headers.get("area")
        idx_estado = headers.get("estado")
        idx_fecha = headers.get("ingreso/retiro") or headers.get("ingreso_retiro")

    # Índices de columnas en el df de talento humano (flexibles por nombre)
    df_cols = _normalize_series(df.columns)

    def get_df_val(row, name_candidates):
        """Devuelve el valor de row por primera columna coincidente (normalizada)."""
        for name in name_candidates:
            try:
                pos = df_cols.index(name.lower().strip())
                return row.iloc[pos]
            except ValueError:
                continue
        return None

    # Cargar índice de CEDULA existente en hoja para búsquedas rápidas
    cedula_to_row = {}
    for r in range(2, ws.max_row + 1):
        ced_val = ws.cell(row=r, column=idx_ced).value if idx_ced else None
        if ced_val is not None and str(ced_val).strip():
            cedula_to_row[str(ced_val).strip()] = r

    hoy_str = date.today().strftime("%d-%m-%Y")

    for _, row in df.iterrows():
        ced = get_df_val(row, ["cedula", "cédula", "documento", "id"])
        nom = get_df_val(row, ["nombre", "nombres", "funcionario"])
        area = get_df_val(row, ["área", "area", "dependencia", "gerencia"])
        fecha_insumo = get_df_val(row, ["ingreso/retiro", "fecha", "fecha evento"])

        # Normalizar cedula a str
        if pd.isna(ced) or str(ced).strip() == "":
            # si no hay cédula válida, saltamos
            continue
        ced_str = str(ced).strip()

        # construir ESTADO según tipo
        dep_str = str(area).strip() if area is not None and str(area).strip() else ""
        if tipo_evento == "personal_retiros":
            estado_final = f"RETIRADO {dep_str}".strip()
        elif tipo_evento in ("personal_ingresos", "personal_actualizacion"):
            estado_final = f"ACTIVO {dep_str}".strip()
        else:
            estado_final = str(get_df_val(row, ["estado"])) if get_df_val(row, ["estado"]) else ""

        # fecha final
        fecha_final = (
            str(fecha_insumo).strip()
            if (fecha_insumo is not None and str(fecha_insumo).strip())
            else hoy_str
        )

        if ced_str in cedula_to_row:
            # Actualizar fila existente
            r = cedula_to_row[ced_str]
            if idx_nom: ws.cell(row=r, column=idx_nom, value=None if pd.isna(nom) else nom)
            if idx_area: ws.cell(row=r, column=idx_area, value=None if pd.isna(area) else area)
            if idx_estado: ws.cell(row=r, column=idx_estado, value=estado_final)
            if idx_fecha: ws.cell(row=r, column=idx_fecha, value=fecha_final)
        else:
            # Crear nueva fila
            r = ws.max_row + 1
            if idx_ced: ws.cell(row=r, column=idx_ced, value=ced_str)
            if idx_nom: ws.cell(row=r, column=idx_nom, value=None if pd.isna(nom) else nom)
            if idx_area: ws.cell(row=r, column=idx_area, value=None if pd.isna(area) else area)
            if idx_estado: ws.cell(row=r, column=idx_estado, value=estado_final)
            if idx_fecha: ws.cell(row=r, column=idx_fecha, value=fecha_final)

    return ws

import re
from datetime import datetime, date as date_type

# ----------------------------
# VALIDACIONES DINÁMICAS (HOJA "General")
# ----------------------------

def parse_date_like(value):
    """Intenta convertir value a objeto date (date_type). Devuelve None si no es parseable."""
    if value is None:
        return None
    if isinstance(value, (datetime, date_type)):
        return value.date() if isinstance(value, datetime) else value
    s = str(value).strip()
    if not s:
        return None
    # intentos de parseo comunes
    patterns = [
        "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d",
        "%d.%m.%Y", "%Y.%m.%d"
    ]
    # intento ISO / fromisoformat
    try:
        dt = datetime.fromisoformat(s)
        return dt.date()
    except Exception:
        pass
    for p in patterns:
        try:
            dt = datetime.strptime(s, p)
            return dt.date()
        except Exception:
            continue
    # intento heurístico (números como Excel timestamp no soportado aquí)
    return None

def is_number_like(v):
    try:
        if v is None: return False
        float(v)
        return True
    except Exception:
        return False

def eval_rule(cell_value, operador, valor):
    """
    Evalúa una sola regla contra cell_value.
    operador: string (ejemplo: '=', '!=', 'contiene', 'no contiene', '>', '<', '>=', '<=', 'regex', 'dias_mayor_que', 'empty', 'not_empty')
    valor: string (valor esperado según operador)
    Retorna True/False
    """
    operador = (operador or "").strip().lower()
    raw = cell_value

    # casuística emptiness
    if operador in ("empty", "es_vacio", "vacío", "vacio"):
        return raw is None or (isinstance(raw, str) and raw.strip() == "")
    if operador in ("not_empty", "no_vacio", "no vacío", "no_vacio"):
        return not (raw is None or (isinstance(raw, str) and raw.strip() == ""))

    # Normalizar a string para ciertos operadores
    cell_str = "" if raw is None else str(raw).strip()
    valor_str = "" if valor is None else str(valor).strip()

    # Igual / distinto (intenta comparación numérica si ambos son números)
    if operador in ("=", "==", "igual", "equals"):
        if is_number_like(cell_str) and is_number_like(valor_str):
            return float(cell_str) == float(valor_str)
        return cell_str == valor_str

    if operador in ("!=", "diferente", "not equals"):
        if is_number_like(cell_str) and is_number_like(valor_str):
            return float(cell_str) != float(valor_str)
        return cell_str != valor_str

    # Contiene / no contiene
    if operador in ("contiene", "contains"):
        return valor_str.lower() in cell_str.lower()
    if operador in ("no contiene", "not_contains", "not contains"):
        return valor_str.lower() not in cell_str.lower()

    # Regex
    if operador in ("regex", "re"):
        try:
            return re.search(valor, cell_str, flags=re.IGNORECASE) is not None
        except Exception:
            return False

    # Comparaciones numéricas o de fecha
    if operador in (">", "<", ">=", "<="):
        # Primero intentamos números
        if is_number_like(cell_str) and is_number_like(valor_str):
            a = float(cell_str); b = float(valor_str)
            if operador == ">": return a > b
            if operador == "<": return a < b
            if operador == ">=": return a >= b
            if operador == "<=": return a <= b
        # Intentamos fechas
        cell_date = parse_date_like(raw)
        val_date = parse_date_like(valor_str)
        if cell_date and val_date:
            if operador == ">": return cell_date > val_date
            if operador == "<": return cell_date < val_date
            if operador == ">=": return cell_date >= val_date
            if operador == "<=": return cell_date <= val_date
        # fallback: comparación lexicográfica
        if operador == ">": return cell_str > valor_str
        if operador == "<": return cell_str < valor_str
        if operador == ">=": return cell_str >= valor_str
        if operador == "<=": return cell_str <= valor_str

    # Días mayor que: asume cell_value es fecha; valor = número de días (ej: 15)
    if operador in ("dias_mayor_que", "dias_mayores_que", "days_gt"):
        try:
            dias = int(float(valor_str))
        except Exception:
            return False
        cell_date = parse_date_like(raw)
        if not cell_date:
            return False
        delta = (date_type.today() - cell_date).days
        return delta > dias

    # Si no reconocemos operador, intentar igualdad simple
    return cell_str == valor_str

def apply_validations_to_general(wb, reglas):
    """
    Recorre la hoja 'General' de wb (openpyxl workbook) y aplica las reglas.
    Cada regla debe ser dict con: 'etiqueta', 'columna', 'operador', 'valor'
    - Si la columna no existe se ignora la regla para esa fila.
    - Escribe el resultado (concatenación de etiquetas cumplidas) en 'ANALISIS VALIDACIONES'.
    """
    ws = _ensure_sheet(wb, "General")
    headers = _headers_from_sheet(ws, header_row=1)

    # Buscar/crear columna ANALISIS VALIDACIONES (texto exacto, mayúsculas permitidas)
    key_name = "analisis validaciones"
    col_valid_idx = headers.get(key_name)
    if not col_valid_idx:
        col_valid_idx = ws.max_column + 1
        ws.cell(row=1, column=col_valid_idx, value="ANALISIS VALIDACIONES")
        # actualizar headers map
        headers = _headers_from_sheet(ws, header_row=1)

    # normalizar reglas: colnames en minúsculas y trimmed
    reglas_norm = []
    for r in reglas:
        etiqueta = r.get("etiqueta", "") or ""
        columna = (r.get("columna", "") or "").strip().lower()
        operador = (r.get("operador", "") or "").strip().lower()
        valor = r.get("valor", "") or ""
        if not etiqueta or not columna:
            continue
        reglas_norm.append({
            "etiqueta": etiqueta.strip(),
            "columna": columna,
            "operador": operador,
            "valor": valor
        })

    # Recalcular headers por si se añadió la columna
    headers = _headers_from_sheet(ws, header_row=1)

    # Mapear nombres normalizados a índices (si la hoja usa nombres con tildes, la clave debe ser la normalizada)
    normalized_headers = {k.lower().strip(): v for k, v in headers.items()}

    # Recorrer filas
    for r in range(2, ws.max_row + 1):
        etiquetas_fila = []
        for regla in reglas_norm:
            colname = regla["columna"]
            idx = normalized_headers.get(colname)
            if not idx:
                # Intentar buscar cabecera que contenga la palabra (fuzzy básico)
                matches = [v for k, v in normalized_headers.items() if colname in k]
                if matches:
                    idx = matches[0]
                else:
                    continue
            cell_val = ws.cell(row=r, column=idx).value
            if eval_rule(cell_val, regla["operador"], regla["valor"]):
                etiquetas_fila.append(regla["etiqueta"])
        # Escribir resultado (si hay etiquetas, unir con '; ')
        ws.cell(row=r, column=headers.get(key_name, col_valid_idx), value="; ".join(etiquetas_fila) if etiquetas_fila else None)

    return True



# =============================
# Rutas principales
# =============================
@app.route("/")
def index():
    """Página principal con listado de archivos subidos"""
    return render_template("index.html", archivos=archivos)


@app.route("/upload", methods=["POST"])
def upload():
    """Subir archivo con tipo de insumo"""
    if "archivo" not in request.files:
        flash("No se envió archivo", "error")
        return redirect(url_for("index"))

    file = request.files["archivo"]
    tipo = request.form.get("tipo")

    if not file or file.filename == "":
        flash("Archivo no válido", "error")
        return redirect(url_for("index"))

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
    """Integra archivo subido en el maestro según tipo de insumo"""
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
        df_insumo = pd.read_excel(filepath)

        # === Caso Maestro inicial ===
        if tipo_insumo == "maestro":
            df_insumo.to_excel(MAESTRO_FILE, index=False)
            flash(f"El archivo {filename} se estableció como Maestro", "success")
        else:
            # === Integrar al Maestro existente ===
            if not os.path.exists(MAESTRO_FILE):
                flash("No existe maestro para integrar", "error")
                return redirect(url_for("index"))

            # Cargar workbook con openpyxl (para usar validaciones dinámicas)
            from openpyxl import load_workbook
            wb = load_workbook(MAESTRO_FILE)

            # Dependiendo del tipo de insumo podemos hacer integraciones diferentes
            if tipo_insumo == "general":
                # Actualizar hoja 'General' con datos de soporte
                replace_sheet_content(wb, "General", df_insumo)

            elif tipo_insumo == "personal_ingresos":
                update_estado_gen_usuario(wb, df_insumo, "personal_ingresos")

            elif tipo_insumo == "personal_retiros":
                update_estado_gen_usuario(wb, df_insumo, "personal_retiros")

            elif tipo_insumo == "personal_actualizacion":
                update_estado_gen_usuario(wb, df_insumo, "personal_actualizacion")

            else:
                # Para otros insumos que no tengan integración especial → hoja con su nombre
                replace_sheet_content(wb, tipo_insumo.capitalize(), df_insumo)

            # === Aplicar validaciones dinámicas en la hoja "General" ===
            try:
                apply_validations_to_general(wb, config.get("validaciones", []))
            except Exception as e:
                flash(f"⚠️ Error aplicando validaciones dinámicas: {e}", "warning")

            # Guardar cambios
            wb.save(MAESTRO_FILE)

        flash(f"Archivo {filename} integrado como {tipo_insumo}", "success")
    except Exception as e:
        flash(f"Error al integrar: {str(e)}", "error")

    return redirect(url_for("index"))

# =============================
# Aplicar validaciones al Maestro
# =============================
@app.route("/aplicar-validaciones", methods=["POST"])
def aplicar_validaciones():
    """
    Aplica las reglas configuradas al Maestro en la hoja 'General'.
    Soporta operadores: =, !=, contiene, no_contiene, >, <.
    """
    if not os.path.exists(MAESTRO_FILE):
        flash("⚠️ No existe maestro para aplicar validaciones", "error")
        return redirect(url_for("index"))

    try:
        # Cargar maestro con todas las hojas
        xl = pd.ExcelFile(MAESTRO_FILE)
        if "General" not in xl.sheet_names:
            flash("⚠️ El Maestro no contiene la hoja 'General'", "error")
            return redirect(url_for("index"))

        df_general = xl.parse("General")

        # Asegurar columna de resultados
        if "ANÁLISIS VALIDACIONES" not in df_general.columns:
            df_general["ANÁLISIS VALIDACIONES"] = ""

        reglas = config.get("validaciones", [])
        if not reglas:
            flash("⚠️ No hay reglas de validación configuradas", "warning")
            return redirect(url_for("configuracion_general"))

        # Aplicar reglas a cada fila
        for idx, fila in df_general.iterrows():
            resultado = []
            for regla in reglas:
                columna = regla.get("columna")
                operador = regla.get("operador", "=")
                valor = regla.get("valor", "")
                etiqueta = regla.get("etiqueta", "")

                if columna not in df_general.columns:
                    continue

                celda = str(fila[columna]).strip().lower() if pd.notna(fila[columna]) else ""

                try:
                    # ==========================
                    # Evaluación de operadores
                    # ==========================
                    if operador == "=":
                        if celda == valor.strip().lower():
                            resultado.append(etiqueta)

                    elif operador == "!=":
                        if celda != valor.strip().lower():
                            resultado.append(etiqueta)

                    elif operador == "contiene":
                        if valor.strip().lower() in celda:
                            resultado.append(etiqueta)

                    elif operador == "no_contiene":
                        if valor.strip().lower() not in celda:
                            resultado.append(etiqueta)

                    elif operador == ">":
                        if pd.to_numeric(fila[columna], errors="coerce") > pd.to_numeric(valor, errors="coerce"):
                            resultado.append(etiqueta)

                    elif operador == "<":
                        if pd.to_numeric(fila[columna], errors="coerce") < pd.to_numeric(valor, errors="coerce"):
                            resultado.append(etiqueta)

                except Exception:
                    # Si la comparación falla (ej: texto en operador >)
                    continue

            # Guardar etiquetas concatenadas o 'OK'
            df_general.at[idx, "ANÁLISIS VALIDACIONES"] = ", ".join(resultado) if resultado else "OK"

        # Reescribir el archivo con la hoja modificada
        with pd.ExcelWriter(MAESTRO_FILE, mode="a", if_sheet_exists="replace") as writer:
            df_general.to_excel(writer, sheet_name="General", index=False)

        flash("✅ Validaciones aplicadas correctamente al Maestro", "success")

    except Exception as e:
        flash(f"❌ Error al aplicar validaciones: {str(e)}", "error")

    return redirect(url_for("index"))


# =============================
# Exportaciones
# =============================
@app.route("/exportar-excel")
def exportar_excel():
    if not os.path.exists(MAESTRO_FILE):
        flash("No hay maestro disponible para exportar", "error")
        return redirect(url_for("index"))
    return send_file(MAESTRO_FILE, as_attachment=True, download_name="inventario.xlsx")


@app.route("/exportar-pdf")
def exportar_pdf():
    if not os.path.exists(MAESTRO_FILE):
        flash("No hay maestro disponible para exportar", "error")
        return redirect(url_for("index"))

    df = pd.read_excel(MAESTRO_FILE)
    pdf_file = os.path.join(DATA_FOLDER, "inventario.pdf")
    c = canvas.Canvas(pdf_file, pagesize=letter)
    width, height = letter

    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 50, "Inventario Maestro")

    c.setFont("Helvetica", 10)
    y = height - 80
    for _, row in df.head(30).iterrows():
        line = " | ".join([str(v) for v in row.values])
        c.drawString(50, y, line[:120])
        y -= 15
        if y < 50:
            c.showPage()
            y = height - 50

    c.save()
    return send_file(pdf_file, as_attachment=True, download_name="inventario.pdf")


@app.route("/download-maestro")
def download_maestro():
    if not os.path.exists(MAESTRO_FILE):
        flash("No hay maestro disponible para descargar", "error")
        return redirect(url_for("index"))
    return send_file(MAESTRO_FILE, as_attachment=True, download_name="inventario_maestro.xlsx")


@app.route("/historial")
def historial():
    historial = cargar_json(ARCHIVOS_JSON, [])
    return render_template("historial.html", historial=historial)

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
# Guardar reglas de validaciones (Insumo 1)
# =============================
@app.route("/guardar_configuracion_general", methods=["POST"])
def guardar_configuracion_general():
    """
    Guardar las reglas de validación para la hoja 'General'.
    Cada regla tiene: etiqueta, columna, operador y valor.
    """
    try:
        etiquetas = request.form.getlist("etiquetas[]")
        columnas = request.form.getlist("columnas[]")
        operadores = request.form.getlist("operadores[]")
        valores = request.form.getlist("valores[]")

        reglas = []
        n = max(len(etiquetas), len(columnas), len(valores), len(operadores))
        for i in range(n):
            etiqueta = etiquetas[i].strip() if i < len(etiquetas) else ""
            columna = columnas[i].strip() if i < len(columnas) else ""
            operador = operadores[i].strip() if i < len(operadores) else "="
            valor = valores[i].strip() if i < len(valores) else ""

            if not etiqueta or not columna:
                # Descarta filas incompletas
                continue

            reglas.append({
                "etiqueta": etiqueta,
                "columna": columna,
                "operador": operador,
                "valor": valor
            })

        # Guardar en configuración global
        config["validaciones"] = reglas
        guardar_json(CONFIG_JSON, config)

        flash("✅ Reglas de validación guardadas correctamente.", "success")

    except Exception as e:
        flash(f"❌ Error al guardar reglas: {str(e)}", "error")

    return redirect(url_for("configuracion_general"))

# =============================
# Eliminar archivo subido
# =============================
@app.route("/eliminar/<nombre_archivo>")
def eliminar(nombre_archivo):
    global archivos
    archivos = [a for a in archivos if a["nombre"] != nombre_archivo]
    guardar_json(ARCHIVOS_JSON, archivos)

    path = os.path.join(UPLOAD_FOLDER, nombre_archivo)
    if os.path.exists(path):
        os.remove(path)

    flash(f"Archivo {nombre_archivo} eliminado", "success")
    return redirect(url_for("index"))


# =============================
# Main
# =============================
if __name__ == "__main__":
    app.run(debug=True)
