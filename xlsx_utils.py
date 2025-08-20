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

Requisitos:
- openpyxl
- pandas

"""
from __future__ import annotations

import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
import unicodedata

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter


# =============================== Utilidades base ===============================

def backup_file(path: str) -> str:
    """Crea un respaldo del archivo Excel en el mismo directorio con timestamp.
    Devuelve la ruta del backup creado.
    """
    src = Path(path)
    if not src.exists():
        raise FileNotFoundError(f"No existe el archivo para backup: {path}")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = src.with_name(f"{src.stem}.backup_{ts}{src.suffix}")
    shutil.copy2(str(src), str(dst))
    return str(dst)


def _norm_text(s: str) -> str:
    """Normaliza texto para comparaciones: minúsculas, sin acentos, espacios compactados."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')  # remove accents
    s = re.sub(r"\s+", " ", s)
    return s


def _build_header_map(ws: Worksheet, header_row: int = 1) -> Dict[str, int]:
    """Crea un mapa header_normalizado -> índice de columna (1-based) leyendo una fila.
    Si una celda está vacía, se ignora.
    """
    headers: Dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        key = _norm_text(val)
        if key:
            headers[key] = col
    return headers


def _delete_data_rows(ws: Worksheet, keep_rows: int = 2) -> None:
    """Elimina todas las filas de datos preservando las primeras `keep_rows` filas.
    - Por defecto conservamos 2 filas: fila 1 (encabezados) y fila 2 (plantilla de fórmulas).
    - Si `keep_rows` == 1, se conserva solo la fila 1 (encabezados).
    """
    max_r = ws.max_row
    if max_r > keep_rows:
        ws.delete_rows(keep_rows + 1, max_r - keep_rows)


def _copy_formula_row(ws: Worksheet, from_row: int, to_row_start: int, to_row_end: int) -> None:
    """Copia las fórmulas de `from_row` a cada fila en [to_row_start, to_row_end].
    - Solo copia celdas que tengan fórmula (string que comienza con '=')
    - Las referencias relativas se ajustan automáticamente cuando Excel recalcule.
    """
    if to_row_end < to_row_start:
        return
    for col in range(1, ws.max_column + 1):
        tmpl_val = ws.cell(row=from_row, column=col).value
        if isinstance(tmpl_val, str) and tmpl_val.startswith('='):
            for r in range(to_row_start, to_row_end + 1):
                ws.cell(row=r, column=col).value = tmpl_val


def _set_cell(ws: Worksheet, row: int, col: int, value):
    ws.cell(row=row, column=col, value=value)


# =============================== Integración Antivirus ===============================

def integrate_endpoint_to_antivirus(master_path: str, endpoint_path: str, keep_rows: int = 2) -> Dict:
    """Integra el insumo Endpoint/Antivirus en la hoja "Antivirus" del maestro.

    Mapeo de columnas (Insumo 2 -> Hoja Antivirus):
      - "Endpoint name"      -> "Nombre de equipo"
      - "IP address"         -> "IP"
      - "MAC address"        -> "Mac"
      - "Last logged on user"-> "Last logged on user"
      - "Last Startup"       -> "Last Startup"
      - "Last Shutdown"      -> "Last Shutdown"
      - "Protection Manager" -> "Protection Manager"
      - "Agent Program"      -> "Agent Program"

    Reglas adicionales:
      - Columna "Estado":
          * si Protection Manager contiene "standard endpoint" (case-insensitive, "similar"),
            entonces "Antivirus Ins."; en caso contrario "NO REPORTA".
      - Eliminar filas donde "Last logged on user" esté vacío (tras escribir datos).
      - Preservar encabezados y fórmulas copiando la fila 2 (plantilla) hacia las nuevas filas.

    Parámetros:
      - keep_rows: 2 conserva encabezado+plantilla; 1 elimina la plantilla (solo encabezado).

    Devuelve un dict con resumen: { 'rows_written': int, 'sheet': 'Antivirus' }
    """
    sheet_name = "Antivirus"

    # 1) Cargar maestro y obtener hoja
    wb = load_workbook(master_path, data_only=False, keep_vba=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"La hoja '{sheet_name}' no existe en el maestro.")
    ws = wb[sheet_name]

    # 2) Leer insumo Endpoint con pandas
    df = pd.read_excel(endpoint_path)

    # Normalizamos nombres de columnas de DF para búsqueda flexible
    df_cols_norm = { _norm_text(c): c for c in df.columns }

    def get_df_col(*candidates: str) -> Optional[str]:
        """Encuentra en df.columns la primera coincidencia entre varios candidatos (normalizados)."""
        for cand in candidates:
            key = _norm_text(cand)
            if key in df_cols_norm:
                return df_cols_norm[key]
        return None

    # Mapeo de origen (df) -> destino (hoja)
    src_dst_pairs = [
        (get_df_col("Endpoint name"),      "Nombre de equipo"),
        (get_df_col("IP address"),         "IP"),
        (get_df_col("MAC address"),        "Mac"),
        (get_df_col("Last logged on user"),"Last logged on user"),
        (get_df_col("Last Startup"),       "Last Startup"),
        (get_df_col("Last Shutdown"),      "Last Shutdown"),
        (get_df_col("Protection Manager"), "Protection Manager"),
        (get_df_col("Agent Program"),      "Agent Program"),
    ]

    # Validación mínima: al menos una columna clave debe existir
    if not any(src for src, _ in src_dst_pairs):
        raise ValueError("El insumo de Endpoint no contiene ninguna de las columnas esperadas.")

    # 3) Construir mapa de headers de la hoja destino
    headers_map = _build_header_map(ws, header_row=1)

    def get_dst_col(title: str) -> Optional[int]:
        return headers_map.get(_norm_text(title))

    # 4) Limpiar datos, preservando filas según keep_rows
    _delete_data_rows(ws, keep_rows=keep_rows)

    # 5) Insertar filas nuevas a partir de start_row
    start_row = keep_rows + 1  # normalmente 3
    rows_written = 0

    # Indices de columnas de destino
    dst_col_indices: Dict[str, int] = {}
    for _, dst_name in src_dst_pairs:
        col_idx = get_dst_col(dst_name)
        if col_idx:
            dst_col_indices[dst_name] = col_idx

    # Índices para Estado y Last logged on user
    col_estado = get_dst_col("Estado")
    col_last_user = get_dst_col("Last logged on user")

    # Escribir valores fila por fila
    for i, (_, row) in enumerate(df.iterrows(), start=0):
        dest_row = start_row + i
        # Para cada par mapeado, si existe fuente y destino, escribir
        for (src_name, dst_name) in src_dst_pairs:
            if not src_name:
                continue
            dst_col = dst_col_indices.get(dst_name)
            if not dst_col:
                continue
            value = row[src_name]
            _set_cell(ws, dest_row, dst_col, value)

        # Reglas para Estado (depende de Protection Manager)
        if col_estado:
            pm_col = dst_col_indices.get("Protection Manager")
            pm_val = ws.cell(row=dest_row, column=pm_col).value if pm_col else None
            pm_norm = _norm_text(pm_val)
            if pm_norm and ("standard endpoint" in pm_norm or "standard enpoint" in pm_norm):  # tolerar typo
                estado_val = "Antivirus Ins."
            else:
                estado_val = "NO REPORTA"
            _set_cell(ws, dest_row, col_estado, estado_val)

        rows_written += 1

    end_row = start_row + rows_written - 1

    # 6) Copiar fórmulas desde la fila plantilla (2) a nuevas filas
    if keep_rows >= 2 and rows_written > 0 and ws.max_row >= 2:
        _copy_formula_row(ws, from_row=2, to_row_start=start_row, to_row_end=end_row)

    # 7) Eliminar filas sin "Last logged on user"
    if col_last_user and rows_written > 0:
        # Recorremos de abajo hacia arriba para evitar desplazamientos
        for r in range(end_row, start_row - 1, -1):
            val = ws.cell(row=r, column=col_last_user).value
            if val is None or str(val).strip() == "":
                ws.delete_rows(r, 1)
                rows_written -= 1
                end_row -= 1

    # 8) Guardar
    wb.save(master_path)

    return {
        "sheet": sheet_name,
        "rows_written": max(rows_written, 0),
        "start_row": start_row,
        "end_row": max(end_row, start_row - 1),
        "keep_rows": keep_rows,
    }


# =============================== Utilidad genérica ===============================

def replace_sheet_with_df(master_path: str, sheet_name: str, df: pd.DataFrame, keep_rows: int = 1) -> Dict:
    """Reemplaza los datos de una hoja con los de un DataFrame, preservando encabezados.
    - Mantiene la fila 1 como encabezado (y opcionalmente la fila 2 como plantilla si keep_rows>=2)
    - Copia fórmulas de la fila plantilla si existe y keep_rows>=2
    """
    wb = load_workbook(master_path, data_only=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"La hoja '{sheet_name}' no existe en el maestro.")
    ws = wb[sheet_name]

    headers_map = _build_header_map(ws, header_row=1)
    _delete_data_rows(ws, keep_rows=keep_rows)

    start_row = keep_rows + 1
    rows_written = 0

    # Escribir por coincidencia de encabezados (columna a columna)
    for i, (_, row) in enumerate(df.iterrows(), start=0):
        r = start_row + i
        for col_name in df.columns:
            dst_col = headers_map.get(_norm_text(col_name))
            if dst_col:
                ws.cell(row=r, column=dst_col, value=row[col_name])
        rows_written += 1

    # Copia de fórmulas si hay plantilla
    if keep_rows >= 2 and rows_written > 0 and ws.max_row >= 2:
        _copy_formula_row(ws, from_row=2, to_row_start=start_row, to_row_end=start_row + rows_written - 1)

    wb.save(master_path)
    return {"sheet": sheet_name, "rows_written": rows_written, "start_row": start_row}
# ---------------------------
# Insumo 3: Actualizar hoja ESTADO_GEN_USUARIO
# ---------------------------
def integrate_personnel_to_estado(master_path, personnel_path, keep_rows=2, save_as=None, make_backup=True):
    """
    Actualiza la hoja 'ESTADO_GEN_USUARIO' con la información de personal (insumo 3).
    - Busca por CEDULA; si encuentra actualiza ESTADO e INGRESO/RETIRO según la fila.
    - Si no encuentra, añade la fila nueva.
    - INTERPRETACIÓN de tipo (ingreso/retiro) basada en campo 'estado' textual o el nombre del archivo.
    """
    if not os.path.exists(master_path):
        raise FileNotFoundError(master_path)
    if not os.path.exists(personnel_path):
        raise FileNotFoundError(personnel_path)

    df_raw = pd.read_excel(personnel_path, engine='openpyxl')
    df, orig_map = _ensure_df_cols_normalized(df_raw)

    # columnas esperadas (normalized)
    need_cols = {
        'cedula': 'cedula',
        'nombre': 'nombre',
        'dependencia': 'dependencia',
        'estado': 'estado',
        'ingreso/retiro': 'ingreso/retiro'
    }
    # intentar ubicar las columnas en df
    # Si no existe 'ingreso/retirO' también aceptar 'fecha' o 'fecha ingreso' etc.
    alt_date_cols = ['ingreso', 'fecha', 'fecha ingreso', 'fecha_retiro', 'ingreso/retirO']
    # open master
    backup = None
    if make_backup:
        backup = backup_file(master_path)

    wb = load_workbook(master_path)
    ws = _find_sheet(wb, "ESTADO_GEN_USUARIO")
    if ws is None:
        raise ValueError("Hoja 'ESTADO_GEN_USUARIO' no encontrada.")

    # localizar header row y mapeo
    expected_headers = ['cedula', 'nombre', 'dependencia', 'estado', 'ingreso/retiro']
    header_row = _find_header_row(ws, expected_headers, max_scan=6)
    master_map = _map_master_headers(ws, header_row)

    # construir índice maestro por cedula (string normalized) -> row number
    master_index = {}
    ced_idx = master_map.get(_normalize('cedula'))
    if ced_idx:
        for r in range(header_row+1, ws.max_row+1):
            val = ws.cell(row=r, column=ced_idx).value
            if val is None: continue
            master_index[str(val).strip()] = r

    # proceso fila por fila del df
    rows_added = 0
    rows_updated = 0
    for _, row in df.iterrows():
        # intentar extraer cedula
        ced_val = None
        for candidate in ['cedula', 'cedula #', 'id', 'identificacion', 'numero documento']:
            if candidate in row.index:
                ced_val = row[candidate]
                break
        if ced_val is None:
            # intentar por columnas que contengan 'ced'
            for col in row.index:
                if 'ced' in col:
                    ced_val = row[col]
                    break
        if pd.isna(ced_val):
            continue
        ced_str = str(ced_val).strip()

        # determinar estado textual y fecha
        estado_val = None
        for c in ['estado', 'estado_actual', 'status']:
            if c in row.index:
                estado_val = row[c]
                break
        if estado_val is None:
            # no hay estado en esta fila -> buscar en nombre del archivo para deducir ingreso/retiro
            estado_val = ''

        # fecha
        fecha_val = None
        for alt in ['ingreso/retirO','ingreso', 'fecha', 'fecha ingreso', 'fecha_retiro', 'fecha_retiro']:
            if alt in row.index:
                fecha_val = row[alt]
                break
        # heurística: si estado contiene 'retir' => retiro; si 'activo' o 'ingres' => ingreso
        texto = '' if estado_val is None else str(estado_val).lower()
        if 'retir' in texto:
            action = 'retiro'
        elif 'activo' in texto or 'ingres' in texto or 'ingreso' in texto:
            action = 'ingreso'
        else:
            # fallback: mirar nombre del archivo
            lowerfname = os.path.basename(personnel_path).lower()
            if 'retiro' in lowerfname:
                action = 'retiro'
            elif 'ingreso' in lowerfname:
                action = 'ingreso'
            else:
                action = 'ingreso'  # por defecto ingreso

        # buscar en maestro
        if ced_str in master_index:
            rownum = master_index[ced_str]
            # actualizar ESTADO y fecha
            if _normalize('estado') in master_map:
                ws.cell(row=rownum, column=master_map[_normalize('estado')]).value = estado_val if estado_val is not None else ( 'RETIRADO' if action == 'retiro' else 'ACTIVO' )
            if _normalize('ingreso/retiro') in master_map and fecha_val is not None and not pd.isna(fecha_val):
                ws.cell(row=rownum, column=master_map[_normalize('ingreso/retiro')]).value = fecha_val
            rows_updated += 1
        else:
            # agregar nueva fila al final
            write_row = ws.max_row + 1
            # llenar columnas si existen en maestro
            for master_col_norm, col_idx in master_map.items():
                # buscar valor en row por nombres similares
                val = None
                # buscar por nombre de columna exacto normalizado
                if master_col_norm in row.index:
                    val = row[master_col_norm]
                else:
                    # intentar mapear por algunas claves
                    if 'ced' in master_col_norm:
                        val = ced_str
                    elif 'nombre' in master_col_norm:
                        # buscar cualquier columna que contenga 'nombre'
                        for c in row.index:
                            if 'nombre' in c:
                                val = row[c]
                                break
                    elif 'depend' in master_col_norm:
                        for c in row.index:
                            if 'depend' in c:
                                val = row[c]
                                break
                    elif 'estado' in master_col_norm:
                        val = estado_val if estado_val is not None else ( 'RETIRADO' if action == 'retiro' else 'ACTIVO' )
                    elif 'ingreso' in master_col_norm or 'retiro' in master_col_norm:
                        val = fecha_val
                ws.cell(row=write_row, column=col_idx).value = val
            rows_added += 1
            master_index[ced_str] = write_row

    # Forzar recálculo y guardar
    try:
        wb.calc_properties.fullCalcOnLoad = True
    except:
        try:
            wb.calculation_properties.fullCalcOnLoad = True
        except:
            pass

    out = save_as if save_as else master_path
    wb.save(out)
    return {"master": master_path, "personnel": personnel_path, "added": rows_added, "updated": rows_updated, "backup": backup, "saved_to": out}

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
