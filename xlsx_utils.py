# xlsx_utils.py
# Requisitos: pandas, openpyxl
# pip install pandas openpyxl

import os
import re
import shutil
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook

def _normalize(name):
    """Normaliza nombres para comparar (quita espacios extra, minúsculas)."""
    if name is None:
        return ''
    return re.sub(r'\s+', ' ', str(name).strip()).lower()

def _find_sheet(wb, target):
    """Busca hoja por nombre (case-insensitive). Devuelve Worksheet o None."""
    for n in wb.sheetnames:
        if _normalize(n) == _normalize(target):
            return wb[n]
    return None

def _find_header_row(ws, expected_names, max_scan=6):
    """
    Intenta ubicar la fila de encabezados buscando coincidencias entre expected_names
    en las primeras max_scan filas. Devuelve número de fila (1-based). Por defecto 1.
    """
    exp = {_normalize(x) for x in expected_names}
    best_row = 1
    best_score = 0
    for r in range(1, max_scan+1):
        row_vals = [_normalize(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column+1)]
        score = sum(1 for v in row_vals if v in exp)
        if score > best_score:
            best_score = score
            best_row = r
        if score >= max(1, len(exp)//2):
            return r
    return best_row

def _map_master_headers(ws, header_row):
    """Mapea encabezados del maestro a (normalized_name -> columna_index)."""
    mapping = {}
    for c in range(1, ws.max_column+1):
        val = ws.cell(row=header_row, column=c).value
        if val is None:
            continue
        mapping[_normalize(val)] = c
    return mapping

def backup_file(path):
    """Crea una copia de seguridad timestamped del archivo dado y devuelve la ruta."""
    dirn, base = os.path.split(path)
    name, ext = os.path.splitext(base)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = os.path.join(dirn, f"{name}_backup_{stamp}{ext}")
    shutil.copy2(path, target)
    return target

def _ensure_df_cols_normalized(df):
    """normaliza columnas del DataFrame (lower & stripped) devolviendo mapping original->normalized"""
    df = df.copy()
    orig_cols = list(df.columns)
    norm_cols = [_normalize(c) for c in orig_cols]
    df.columns = norm_cols
    return df, dict(zip(norm_cols, orig_cols))

# ---------------------------
# Función genérica: reemplazar contenido de una hoja (preservando keep_rows)
# ---------------------------
def replace_sheet_with_df(master_path, sheet_name, df, keep_rows=2, save_as=None, make_backup=True):
    """
    Reemplaza el contenido de la hoja 'sheet_name' del libro maestro con el DataFrame df.
    - keep_rows: número de filas superiores a preservar (encabezados/formulas).
    - make_backup: si True hace copia de seguridad del master antes de guardar.
    - save_as: ruta destino, si None sobrescribe master_path.
    Retorna resumen dict.
    """
    if not os.path.exists(master_path):
        raise FileNotFoundError(master_path)
    # backup
    backup = None
    if make_backup:
        backup = backup_file(master_path)

    wb = load_workbook(master_path)
    ws = _find_sheet(wb, sheet_name)
    if ws is None:
        raise ValueError(f"Hoja '{sheet_name}' no encontrada en {master_path}")

    # encontrar header row (intentar fila 1..keep_rows)
    header_row = _find_header_row(ws, df.columns.tolist(), max_scan=keep_rows+2)

    # borrar contenido debajo de keep_rows (mantener header_row y filas superiores)
    insert_row = max(keep_rows + 1, header_row + 1)
    if ws.max_row >= insert_row:
        ws.delete_rows(insert_row, ws.max_row - (insert_row-1))

    # escribir df a partir de insert_row
    # conservar orden de columnas del maestro si existieran; si no, escribir de 1..n
    master_map = _map_master_headers(ws, header_row)
    # Si maestro tiene encabezados que coinciden con df, usarlos; si no, solo escribir en columnas nuevas
    # Construir lista de target column indices en orden de df.columns
    target_cols = []
    for col in df.columns:
        norm = _normalize(col)
        if norm in master_map:
            target_cols.append(master_map[norm])
        else:
            # buscar primera columna vacía
            target_cols.append(None)

    write_row = insert_row
    for _, row in df.iterrows():
        # si target_cols has None, append at the end of existing columns
        if any(tc is None for tc in target_cols):
            # append values starting at first empty column after ws.max_column
            col_idx = ws.max_column + 1
            for i, val in enumerate(row):
                if target_cols[i] is None:
                    ws.cell(row=write_row, column=col_idx).value = val
                    col_idx += 1
                else:
                    ws.cell(row=write_row, column=target_cols[i]).value = val
        else:
            for i, val in enumerate(row):
                ws.cell(row=write_row, column=target_cols[i]).value = val
        write_row += 1

    # forzar recálculo en abrir
    try:
        wb.calc_properties.fullCalcOnLoad = True
    except Exception:
        try:
            wb.calculation_properties.fullCalcOnLoad = True
        except Exception:
            pass

    out = save_as if save_as else master_path
    wb.save(out)
    return {"master": master_path, "sheet": sheet_name, "rows_written": len(df), "backup": backup, "saved_to": out}

# ---------------------------
# Insumo 2: Endpoint -> hoja "Antivirus"
# ---------------------------
def integrate_endpoint_to_antivirus(master_path, endpoint_path, keep_rows=2, save_as=None, make_backup=True):
    """
    Toma el Endpoint excel (endpoint_path) y actualiza la hoja 'Antivirus' en master_path.
    Mapea columnas desde endpoint a las columnas del maestro según tus reglas.
    """
    if not os.path.exists(master_path):
        raise FileNotFoundError(master_path)
    if not os.path.exists(endpoint_path):
        raise FileNotFoundError(endpoint_path)

    # leer endpoint
    df_raw = pd.read_excel(endpoint_path, engine='openpyxl')
    df, orig_map = _ensure_df_cols_normalized(df_raw)

    # mapeo de columnas (normalized endpoint name -> normalized master column name)
    mapping = {
        'endpoint name': 'nombre de equipo',
        'ip address': 'ip',
        'mac address': 'mac',
        'last logged on user': 'last logged on user',
        'last startup': 'last startup',
        'last shutdown': 'last shutdown',
        'protection manager': 'protection manager',
        'agent program': 'agent program'
    }

    # construir df_out con columnas maestras (solo las columnas mapeadas que existan en endpoint)
    df_out = pd.DataFrame()
    for ep_col_norm, master_col in mapping.items():
        if ep_col_norm in df.columns:
            df_out[master_col] = df[ep_col_norm]
        else:
            # si no existe, crea columna vacía
            df_out[master_col] = pd.NA

    # añadimos columna Estado Absolute? no la escribimos, ya que es formula en maestro
    # Guardar backup
    backup = None
    if make_backup:
        backup = backup_file(master_path)

    # Abrir maestro y localizar hoja
    wb = load_workbook(master_path)
    ws = _find_sheet(wb, "Antivirus")
    if ws is None:
        raise ValueError("Hoja 'Antivirus' no encontrada en el maestro.")

    # localizar header row y mapear columnas
    expected_headers = list(df_out.columns) + ['estado', 'estado absolute']
    header_row = _find_header_row(ws, expected_headers, max_scan=6)
    master_map = _map_master_headers(ws, header_row)

    # borrar filas debajo de keep_rows
    insert_row = max(keep_rows + 1, header_row + 1)
    if ws.max_row >= insert_row:
        ws.delete_rows(insert_row, ws.max_row - (insert_row-1))

    # escribir df_out
    write_row = insert_row
    for _, row in df_out.iterrows():
        for col_name in df_out.columns:
            master_idx = master_map.get(_normalize(col_name))
            if master_idx:
                ws.cell(row=write_row, column=master_idx).value = row[col_name]
        write_row += 1

    # Actualizar columna 'Estado' basado en 'Protection Manager'
    pm_idx = master_map.get(_normalize('protection manager'))
    estado_idx = master_map.get(_normalize('estado')) or master_map.get(_normalize('estado_gen')) or None
    # si no se encuentra 'estado', tratamos de encontrar 'estado' en cualquier header que contenga 'estado'
    if estado_idx is None:
        for k, v in master_map.items():
            if 'estado' in k:
                estado_idx = v
                break

    if pm_idx and estado_idx:
        for r in range(insert_row, ws.max_row + 1):
            pm_val = ws.cell(row=r, column=pm_idx).value
            text = '' if pm_val is None else str(pm_val).lower()
            if 'standard' in text and 'endpoint' in text:
                ws.cell(row=r, column=estado_idx).value = "Antivirus Ins."
            else:
                ws.cell(row=r, column=estado_idx).value = "NO REPORTA"

    # Eliminar filas sin 'last logged on user'
    llu_idx = master_map.get(_normalize('last logged on user'))
    if llu_idx:
        for r in range(ws.max_row, insert_row-1, -1):
            v = ws.cell(row=r, column=llu_idx).value
            if v is None or (isinstance(v, str) and v.strip() == ''):
                ws.delete_rows(r, 1)

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
    return {"master": master_path, "endpoint": endpoint_path, "rows_written": len(df_out), "backup": backup, "saved_to": out}

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
