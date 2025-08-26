"""xlsx_utils.py

Utilidades para integrar insumos al Libro Maestro (Insumo 1: "Inventario de proveedor").

Puntos clave:
- Cada insumo SOLO actualiza una hoja específica del maestro.
- Se debe preservar la fila de encabezados y las fórmulas.
- Estrategia recomendada: usar la fila 2 como "fila plantilla" de fórmulas.
  - Se copian sus fórmulas hacia todas las filas nuevas generadas.
  - Se pueden conservar las dos primeras filas (encabezado + plantilla) mediante el
    parámetro keep_rows=2. Si se pasa keep_rows=1, se eliminará la fila plantilla.
- Para Excel, las fórmulas se recalculan al abrir el archivo (data_only=False en openpyxl).

Este módulo implementa:
- backup_file(): resguarda el maestro antes de escribir.
- integrate_endpoint_to_antivirus(): integra Insumo 2 (Endpoint/Antivirus) en hoja "Antivirus".
- replace_sheet_with_df(): utilidad genérica (no usada directamente en Antivirus, pero disponible).

# xlsx_utils.py
# Requisitos: pandas, openpyxl
# pip install pandas openpyxl

from __future__ import annotations
"""
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Iterable
import unicodedata

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

# ---------- utilidades (idénticas a las previas) ----------

def backup_file(path: str) -> str:
    src = Path(path)
    if not src.exists():
        raise FileNotFoundError(f"No existe el archivo para backup: {path}")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = src.with_name(f"{src.stem}_backup_{ts}{src.suffix}")
    shutil.copy2(str(src), str(dst))
    return str(dst)

def _norm_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    s = re.sub(r"\s+", " ", s)
    return s

def _build_header_map(ws: Worksheet, header_row: int = 1) -> Dict[str, int]:
    headers: Dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        key = _norm_text(val)
        if key:
            headers[key] = col
    return headers

def _delete_data_rows(ws: Worksheet, keep_rows: int = 2) -> None:
    max_r = ws.max_row
    if max_r > keep_rows:
        ws.delete_rows(keep_rows + 1, max_r - keep_rows)

def _copy_formula_row(ws: Worksheet, from_row: int, to_row_start: int, to_row_end: int) -> None:
    if to_row_end < to_row_start:
        return
    for col in range(1, ws.max_column + 1):
        tmpl_val = ws.cell(row=from_row, column=col).value
        if isinstance(tmpl_val, str) and tmpl_val.startswith('='):
            for r in range(to_row_start, to_row_end + 1):
                ws.cell(row=r, column=col).value = tmpl_val

def _set_cell(ws: Worksheet, row: int, col: int, value):
    ws.cell(row=row, column=col, value=value)

def _first_match_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    norm_map = { _norm_text(c): c for c in df.columns }
    for cand in candidates:
        k = _norm_text(cand)
        if k in norm_map:
            return norm_map[k]
    return None

def _fecha_from_filename(filename: str) -> Optional[datetime]:
    base = Path(filename).stem
    m = re.search(r'(\d{4})[._-]?(\d{2})[._-]?(\d{2})', base)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            return datetime(y, mo, d)
        except Exception:
            pass
    m = re.search(r'(\d{2})[._-](\d{2})[._-](\d{4})', base)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return datetime(y, mo, d)
        except Exception:
            pass
    return None

def _clean_cedula(value) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    s = s.replace('.', '').replace(',', '')
    s = re.sub(r'\D+', '', s)
    return s or None

def _compose_nombre(df_row: pd.Series, cols: Dict[str, Optional[str]]) -> Optional[str]:
    nc = cols.get('nombre_completo')
    if nc and pd.notna(df_row.get(nc)):
        val = str(df_row[nc]).strip()
        if val:
            return val
    pa = cols.get('primer_apellido')
    sa = cols.get('segundo_apellido')
    pn = cols.get('primer_nombre')
    sn = cols.get('segundo_nombre')
    apellidos = []
    nombres = []
    for key in (pa, sa):
        if key and pd.notna(df_row.get(key)):
            txt = str(df_row[key]).strip()
            if txt:
                apellidos.append(txt)
    for key in (pn, sn):
        if key and pd.notna(df_row.get(key)):
            txt = str(df_row[key]).strip()
            if txt:
                nombres.append(txt)
    if not (apellidos or nombres):
        return None
    return f"{' '.join(apellidos)} {' '.join(nombres)}".strip()

# ==================== Integración Endpoint (idéntica) ====================
# (incluye integrate_endpoint_to_antivirus tal y como ya tenías - omito aquí por brevedad)
# --- asegúrate de mantener tu implementación anterior para integrate_endpoint_to_antivirus ---
# (Si quieres, te incluyo también la versión completa, pero aquí nos centramos en personnel.)

# ==================== Integración Personal (actualizada/incremental) ====================

def integrate_personnel_to_estado(
    master_path: str,
    personnel_path: str,
    keep_rows: int = 2,
    area: Optional[str] = None,
    operacion: Optional[str] = None,
    fecha_archivo: Optional[datetime] = None,
    cfg: Optional[Dict] = None,
    make_backup: bool = True,
) -> Dict:
    """
    Versión incremental: buscar por CEDULA en la hoja 'ESTADO_GEN_USUARIO' y:
      - si existe: actualizar campos relevantes (nombre, dependencia, area, estado, ingreso/retiro)
      - si no existe: agregar nueva fila (al final)
    No borra la tabla completa.
    Devuelve resumen: {added, updated, skipped, sheet, ...}
    """
    sheet_name = "ESTADO_GEN_USUARIO"

    if not os.path.exists(master_path):
        raise FileNotFoundError(master_path)
    if not os.path.exists(personnel_path):
        raise FileNotFoundError(personnel_path)

    # Backup opcional
    backup = None
    if make_backup:
        backup = backup_file(master_path)

    wb = load_workbook(master_path, data_only=False, keep_vba=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"La hoja '{sheet_name}' no existe en el maestro.")
    ws = wb[sheet_name]

    headers_map = _build_header_map(ws, header_row=1)

    # columnas destino (obligatoria cedula)
    dst_cedula = headers_map.get(_norm_text("CEDULA"))
    dst_nombre = headers_map.get(_norm_text("NOMBRE"))
    dst_dependencia = headers_map.get(_norm_text("DEPENDENCIA"))
    dst_area = headers_map.get(_norm_text("AREA"))
    dst_estado = headers_map.get(_norm_text("ESTADO"))
    dst_ing_ret = headers_map.get(_norm_text("INGRESO/RETIRO")) or headers_map.get(_norm_text("INGRESO")) or headers_map.get(_norm_text("FECHA"))

    if not dst_cedula:
        raise ValueError("No se encontró la columna 'CEDULA' en la hoja destino. Es obligatoria para actualizaciones incrementales.")

    # leer df de personal
    df = pd.read_excel(personnel_path, engine='openpyxl')

    # detectar columnas fuente (flexible)
    col_doc = _first_match_column(df, ["Documento", "Cédula", "Cedula", "NUMERO DOCUMENTO", "No. documento", "No documento", "Documento de identidad"])
    col_nombre_completo = _first_match_column(df, ["NOMBRE COMPLETO", "Nombre completo", "Nombres y apellidos"])
    col_primer_nombre = _first_match_column(df, ["Primer nombre", "Primer Nombre", "Nombre 1", "firstname"])
    col_segundo_nombre = _first_match_column(df, ["Segundo nombre", "Segundo Nombre", "Nombre 2"])
    col_primer_apellido = _first_match_column(df, ["Primer apellido", "Primer Apellido", "Apellido 1", "lastname"])
    col_segundo_apellido = _first_match_column(df, ["Segundo apellido", "Segundo Apellido", "Apellido 2"])
    col_dependencia = _first_match_column(df, ["DEPENDENCIA", "Dependencia", "CENTRO DE COSTOS", "Centro de costos", "Centro Costos"])
    col_area = _first_match_column(df, ["AREA", "Área", "REGIONAL", "Regional"])
    col_fec_term = _first_match_column(df, ["FECHA TERMINACIÓN", "FECHA TERMINACION", "Fecha terminación", "Fecha terminacion", "fecha fin", "fecha_terminacion"])
    col_fec_fin = _first_match_column(df, ["FECHA FIN", "Fecha fin", "fechafin"])
    col_fec_ini = _first_match_column(df, ["FECHA INICIO", "Fecha inicio", "fecha inicio", "fechaingreso"])

    # deducción global por nombre de archivo si hace falta
    filename = Path(personnel_path).name
    fname_norm = _norm_text(filename)

    def _guess_operacion() -> str:
        if operacion:
            return operacion.lower()
        if any(x in fname_norm for x in ["retiro", "retir", "terminacion", "terminación", "fin"]):
            return "retiros"
        if "ingres" in fname_norm:
            return "ingresos"
        return "mixto"

    def _guess_area() -> str:
        if area:
            return area
        if "fomag" in fname_norm:
            return "FOMAG"
        if "m.c" in fname_norm or "mc" in fname_norm:
            return "M.C"
        if "apre" in fname_norm or "pract" in fname_norm:
            return "APRE Y PRACT"
        if "mision" in fname_norm or "misión" in fname_norm:
            return "FIDU MISIÓN"
        if "planta" in fname_norm or "fidu" in fname_norm:
            return "FIDU PLANTA"
        return "FIDU PLANTA"

    def _estado_from(op: str, ar: str, row_has_term: bool, row_has_ini: bool) -> str:
        op = (op or "").lower()
        ar_up = (ar or "").upper()
        if "reti" in op or "termin" in op:
            base = "RETIRADO"
        elif "ingre" in op:
            base = "ACTIVO"
        else:
            # mixto: deducir por fila
            if row_has_term:
                base = "RETIRADO"
            elif row_has_ini:
                base = "ACTIVO"
            else:
                base = "ACTIVO"
        if "FOMAG" in ar_up:
            suf = "FOMAG"
        elif "M.C" in ar_up or "MC" in ar_up:
            suf = "M.C"
        elif "APRE" in ar_up or "PRACT" in ar_up:
            suf = "APRE Y PRACT"
        elif "MISION" in _norm_text(ar_up) or "MISIÓN" in ar_up:
            suf = "FIDU MISIÓN"
        else:
            suf = "FIDU PLANTA"
        return f"{base} {suf}"

    op_global = _guess_operacion()
    area_global = _guess_area()
    if fecha_archivo is None:
        fecha_archivo = _fecha_from_filename(filename)

    # construir índice de cédulas existentes -> fila (desde start_row hasta final)
    start_row = keep_rows + 1
    existing_map: Dict[str, int] = {}
    for r in range(start_row, ws.max_row + 1):
        cell_val = ws.cell(row=r, column=dst_cedula).value
        c = _clean_cedula(cell_val)
        if c:
            # si existen duplicados, conservamos la primera aparición
            if c not in existing_map:
                existing_map[c] = r

    added = 0
    updated = 0
    skipped = 0
    appended_rows: List[int] = []

    nombre_cols = {
        "nombre_completo": col_nombre_completo,
        "primer_apellido": col_primer_apellido,
        "segundo_apellido": col_segundo_apellido,
        "primer_nombre": col_primer_nombre,
        "segundo_nombre": col_segundo_nombre,
    }

    # Procesar filas de DF
    for _, row in df.iterrows():
        # obtener cedula entrante
        raw_doc = None
        if col_doc and pd.notna(row.get(col_doc)):
            raw_doc = row.get(col_doc)
        # si no hay cedula, intentamos saltar (no procesamos)
        ced = _clean_cedula(raw_doc)
        if not ced:
            skipped += 1
            continue

        # nombre/dependencia/area/fechas en fila de insumo
        nombre_val = _compose_nombre(row, nombre_cols)
        dep_val = None
        if col_dependencia and pd.notna(row.get(col_dependencia)):
            dep_val = str(row.get(col_dependencia)).strip()
        area_val = None
        if col_area and pd.notna(row.get(col_area)):
            area_val = str(row.get(col_area)).strip()
        else:
            area_val = area_global

        fecha_val = None
        # priorizar terminacion/fin sobre inicio
        for ccol in (col_fec_term, col_fec_fin, col_fec_ini):
            if ccol and pd.notna(row.get(ccol)):
                fecha_val = row.get(ccol)
                # si es Timestamp/str, dejar tal cual (openpyxl acepta datetime/date)
                break
        if fecha_val is None and fecha_archivo is not None:
            fecha_val = fecha_archivo.date()

        # deducir flags fila
        row_has_term = (col_fec_term and pd.notna(row.get(col_fec_term))) or (col_fec_fin and pd.notna(row.get(col_fec_fin)))
        row_has_ini = (col_fec_ini and pd.notna(row.get(col_fec_ini))) or (not row_has_term and fecha_val is not None)

        # calcular estado para esta fila
        estado_val = _estado_from(op_global, area_val, row_has_term, row_has_ini)

        if ced in existing_map:
            # actualizar fila existente
            r = existing_map[ced]
            # actualizamos campos solo si incoming no es nulo (salvo ESTADO que escribimos siempre)
            if dst_nombre and nombre_val:
                _set_cell(ws, r, dst_nombre, nombre_val)
            if dst_dependencia and dep_val:
                _set_cell(ws, r, dst_dependencia, dep_val)
            if dst_area and area_val:
                _set_cell(ws, r, dst_area, area_val)
            # fecha: actualizamos si incoming tiene fecha
            if dst_ing_ret and fecha_val is not None:
                _set_cell(ws, r, dst_ing_ret, fecha_val)
            # estado: sobreescribimos siempre con el nuevo calculado
            if dst_estado:
                _set_cell(ws, r, dst_estado, estado_val)
            updated += 1
        else:
            # agregar nueva fila (al final)
            new_row = ws.max_row + 1
            # si new_row < start_row -> colocarlo en start_row
            if new_row < start_row:
                new_row = start_row
            # escribir columnas obligatorias/destino si existen
            _set_cell(ws, new_row, dst_cedula, ced)
            if dst_nombre and nombre_val:
                _set_cell(ws, new_row, dst_nombre, nombre_val)
            if dst_dependencia and dep_val:
                _set_cell(ws, new_row, dst_dependencia, dep_val)
            if dst_area and area_val:
                _set_cell(ws, new_row, dst_area, area_val)
            if dst_ing_ret and fecha_val is not None:
                _set_cell(ws, new_row, dst_ing_ret, fecha_val)
            if dst_estado:
                _set_cell(ws, new_row, dst_estado, estado_val)
            appended_rows.append(new_row)
            # registrar en índice para evitar duplicados en el mismo batch
            existing_map[ced] = new_row
            added += 1

    # copiar fórmulas de la fila plantilla (2) a filas añadidas si procede
    if keep_rows >= 2 and appended_rows:
        min_new = min(appended_rows)
        max_new = max(appended_rows)
        # si las filas nuevas están contiguas y comienzan en start_row, se cubre el rango
        _copy_formula_row(ws, from_row=2, to_row_start=min_new, to_row_end=max_new)

    # guardar
    wb.save(master_path)

    return {
        "sheet": sheet_name,
        "added": added,
        "updated": updated,
        "skipped": skipped,
        "backup": backup,
        "start_row": start_row,
        "last_row": ws.max_row,
        "keep_rows": keep_rows,
    }

# =============================== replace_sheet_with_df (idéntico) ===============================
def replace_sheet_with_df(master_path: str, sheet_name: str, df: pd.DataFrame, keep_rows: int = 1) -> Dict:
    wb = load_workbook(master_path, data_only=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"La hoja '{sheet_name}' no existe en el maestro.")
    ws = wb[sheet_name]

    headers_map = _build_header_map(ws, header_row=1)
    _delete_data_rows(ws, keep_rows=keep_rows)

    start_row = keep_rows + 1
    rows_written = 0

    for i, (_, row) in enumerate(df.iterrows(), start=0):
        r = start_row + i
        for col_name in df.columns:
            dst_col = headers_map.get(_norm_text(col_name))
            if dst_col:
                ws.cell(row=r, column=dst_col, value=row[col_name])
        rows_written += 1

    if keep_rows >= 2 and rows_written > 0 and ws.max_row >= 2:
        _copy_formula_row(ws, from_row=2, to_row_start=start_row, to_row_end=start_row + rows_written - 1)

    wb.save(master_path)
    return {"sheet": sheet_name, "rows_written": rows_written, "start_row": start_row}
# ---------------------------
# Insumo 4 & 5: reemplazos directos de hojas
# ---------------------------

def integrate_tmp_to_useraranda(master_path, tmp_path, keep_rows=2, save_as=None, make_backup=True):
    """
    Reemplaza la hoja 'Useraranda_BLOGIK' con el contenido del archivo tmp_path (primera hoja o la que tenga datos).
    """
    if not os.path.exists(master_path):
        raise FileNotFoundError(master_path)
    if not os.path.exists(tmp_path):
        raise FileNotFoundError(tmp_path)

    df_raw = pd.read_excel(tmp_path, engine='openpyxl')
    # usamos df_raw tal cual; si tiene múltiples hojas, podríamos permitir especificar
    return replace_sheet_with_df(master_path, 'Useraranda_BLOGIK', df_raw, keep_rows=keep_rows, save_as=save_as, make_backup=make_backup)

def integrate_da_to_reporte(master_path, da_path, keep_rows=2, save_as=None, make_backup=True):
    """
    Reemplaza la hoja 'Reporte DA' con el contenido del archivo da_path.
    """
    if not os.path.exists(master_path):
        raise FileNotFoundError(master_path)
    if not os.path.exists(da_path):
        raise FileNotFoundError(da_path)

    df_raw = pd.read_excel(da_path, engine='openpyxl')
    return replace_sheet_with_df(master_path, 'Reporte DA', df_raw, keep_rows=keep_rows, save_as=save_as, make_backup=make_backup)
