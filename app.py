# app.py
"""
Aplicación Flask para manejo de insumos y actualización del libro maestro.
Incluye:
- subida de archivos (múltiple) con renombrado para evitar sobrescribir
- listado de archivos subidos (excluye history/config cuando corresponde)
- integración por tipo de insumo (endpoint, personnel, tmp, da)
- integración múltiple en batch
- historial y configuración (persistidos en JSON dentro de uploads/)
- descarga del maestro y rutas de export (Excel/PDF)

Notas:
- Este archivo asume que tienes un módulo local `xlsx_utils.py` con las
  funciones de integración. Si no lo tienes, las rutas de integración
  devolverán errores explicativos.
- Las conversiones a PDF intentan usar LibreOffice (`soffice`) si está
  disponible en el PATH. Si no, la ruta devolverá un mensaje de error.
"""

import os
import json
import threading
import subprocess
import platform
import webbrowser
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_file,
    send_from_directory,
    abort,
)
from werkzeug.utils import secure_filename

# Intentamos importar las utilidades de xlsx. Si faltan, las rutas de
# integración mostrarán mensajes claros en vez de fallar en el import.
try:
    from xlsx_utils import (
        integrate_endpoint_to_antivirus,
        integrate_personnel_to_estado,
        integrate_tmp_to_useraranda,
        integrate_da_to_reporte,
        replace_sheet_with_df,
        backup_file,
    )
except Exception:
    integrate_endpoint_to_antivirus = None
    integrate_personnel_to_estado = None
    integrate_tmp_to_useraranda = None
    integrate_da_to_reporte = None
    replace_sheet_with_df = None
    backup_file = None


# ----------------------------- Configuración -----------------------------
app = Flask(__name__)
# Cambia esta clave por una secreta en producción
app.secret_key = os.environ.get("FLASK_SECRET", "cambiar_esta_clave_por_una_segura")

BASE_DIR = Path.cwd()
UPLOAD_FOLDER = BASE_DIR / "uploads"
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)

ALLOWED_EXTENSIONS = {"xls", "xlsx"}
HISTORY_FILE = UPLOAD_FOLDER / "history.json"
CONFIG_FILE = UPLOAD_FOLDER / "config.json"

# ----------------------------- Helpers ----------------------------------

def is_allowed(filename: str) -> bool:
    """Valida extensión permitida (case-insensitive)."""
    if not filename:
        return False
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def load_json_file(path: Path) -> List[Dict]:
    """Carga JSON desde un archivo, devolviendo lista o diccionario vacío.
    Usamos esto para history/config donde puede no existir el archivo."""
    if not path.exists():
        return []
    try:
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        app.logger.exception(f"Error leyendo JSON {path}: {e}")
        return []


def save_json_file(path: Path, data) -> None:
    try:
        with path.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    except Exception as e:
        app.logger.exception(f"Error guardando JSON {path}: {e}")


def append_history(entry: Dict) -> None:
    """Añade una entrada al inicio del historial (history.json)."""
    hist = load_json_file(HISTORY_FILE) or []
    # añadir timestamp si no viene
    if "timestamp" not in entry:
        entry["timestamp"] = datetime.now().isoformat()
    hist.insert(0, entry)
    save_json_file(HISTORY_FILE, hist)


def load_config() -> Dict:
    cfg = load_json_file(CONFIG_FILE)
    # guardamos un dict (si no existe, devolvemos dict vacío)
    return cfg if isinstance(cfg, dict) else {}


def save_config(cfg: Dict) -> None:
    save_json_file(CONFIG_FILE, cfg)


def find_master_file() -> Optional[str]:
    """
    Busca el archivo maestro en uploads por patrones usuales.
    Devuelve ruta absoluta (string) o None.
    """
    for p in UPLOAD_FOLDER.iterdir():
        if not p.is_file():
            continue
        low = p.name.lower()
        # patrones comunes que hemos usado
        if "reporte-data_inventario" in low or "reporte-data inventario" in low or "reporte-data-inventario" in low:
            return str(p)
    # fallback: buscar archivo que contenga 'inventario' en su nombre
    candidates = [p for p in UPLOAD_FOLDER.iterdir() if p.is_file() and "inventario" in p.name.lower()]
    if candidates:
        # devolver el más reciente por modificación
        candidates.sort(key=lambda pp: pp.stat().st_mtime, reverse=True)
        return str(candidates[0])
    return None


def save_uploaded_file(f) -> str:
    """Guarda un archivo subido evitando sobrescribir: si ya existe, añade sufijo con timestamp.
    Devuelve el nombre final del archivo (no la ruta completa)."""
    filename = secure_filename(f.filename)
    dest = UPLOAD_FOLDER / filename
    if dest.exists():
        base = dest.stem
        suffix = dest.suffix
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        filename = f"{base}_{timestamp}{suffix}"
        dest = UPLOAD_FOLDER / filename
    f.save(str(dest))
    return filename


def obtener_archivos_historial() -> List[Dict]:
    """Lista archivos del folder uploads con metadatos (ordenados por fecha de modificación desc).
    Se utiliza para la página de historial / listado general."""
    archivos: List[Dict] = []
    for p in UPLOAD_FOLDER.iterdir():
        if not p.is_file():
            continue
        # omitimos explícitamente los archivos internos (history/config)
        if p.name in {HISTORY_FILE.name, CONFIG_FILE.name}:
            continue
        archivos.append({
            "nombre": p.name,
            "fecha": datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            "ruta": str(p),
            "size": p.stat().st_size,
        })
    archivos.sort(key=lambda x: x["fecha"], reverse=True)
    return archivos


def abrir_navegador(url: str = "http://127.0.0.1:5000") -> None:
    """Intenta abrir el navegador de forma segura en distintos sistemas."""
    try:
        # webbrowser.open es multiplataforma y no suele bloquear
        webbrowser.open(url, new=2)
    except Exception:
        app.logger.exception("No se pudo abrir el navegador automáticamente.")


def detectar_tipo_insumo(nombre: str) -> str:
    """Función de detección textual para asignar tipo a los archivos subidos.
    Devuelve una cadena amigable para mostrar en la UI.
    """
    if not nombre:
        return "Desconocido"
    nombre = nombre.lower()
    if "inventario" in nombre or "insumo1" in nombre:
        return "Inventario Proveedor"
    if "endpoint" in nombre or "antivirus" in nombre or "insumo2" in nombre:
        return "Antivirus"
    if "retiros" in nombre or "personal" in nombre or "ingresos" in nombre or "colaboradores" in nombre or "insumo3" in nombre:
        return "Talento Humano"
    if "directorio" in nombre or "insumo5" in nombre or "directorio activo" in nombre:
        return "Directorio Activo"
    if "aranda" in nombre or "insumo4" in nombre or "tmp" in nombre:
        return "Aranda"
    if "da" in nombre or "reporte da" in nombre:
        return "Reporte DA"
    return "Desconocido"


# ----------------------------- Rutas ------------------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    """
    Página principal:
    - GET: muestra archivos subidos y formularios
    - POST: sube archivos (múltiple) y actualiza lista
    """
    # reconstruir lista desde uploads (no dependemos sólo de variable en memoria)
    archivos_cargados = []
    paths = [p for p in UPLOAD_FOLDER.iterdir() if p.is_file()]
    # ordenamos por fecha modificacion descendente
    paths.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    for p in paths:
        # mostramos sólo archivos permitidos (evitar .json internos)
        if not is_allowed(p.name):
            continue
        archivos_cargados.append({
            "nombre": p.name,
            "fecha": datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            "tipo": detectar_tipo_insumo(p.name),
            "size": p.stat().st_size,
        })

    if request.method == "POST":
        # soportar campos 'files[]' o 'files'
        files = request.files.getlist("files[]") or request.files.getlist("files") or []
        uploaded = []
        for f in files:
            if f and is_allowed(f.filename):
                filename = save_uploaded_file(f)
                uploaded.append(filename)
                append_history({
                    "accion": "upload",
                    "archivo": filename,
                    "timestamp": datetime.now().isoformat(),
                })
            else:
                app.logger.debug("Archivo no permitido o sin nombre: %r", getattr(f, 'filename', None))

        if uploaded:
            flash(f"Se subieron {len(uploaded)} archivo(s): {', '.join(uploaded)}", "success")
        else:
            flash("No se subió ningún archivo válido (extensiones permitidas: xls, xlsx).", "warning")
        return redirect(url_for("index"))

    config = load_config()
    return render_template("index.html", archivos=archivos_cargados, config=config)


@app.route("/eliminar/<path:nombre_archivo>", methods=["GET", "POST"])
def eliminar(nombre_archivo):
    """Elimina archivo físico en uploads y actualiza historial.
    Se acepta GET para compatibilidad, pero lo ideal es usar POST desde la UI.
    """
    safe_name = secure_filename(nombre_archivo)
    ruta = UPLOAD_FOLDER / safe_name
    if not ruta.exists():
        flash("Archivo no encontrado.", "warning")
        return redirect(url_for("index"))
    try:
        ruta.unlink()
        append_history({
            "accion": "delete",
            "archivo": safe_name,
            "timestamp": datetime.now().isoformat(),
        })
        flash(f"Archivo {safe_name} eliminado.", "success")
    except Exception as e:
        app.logger.exception(f"Error eliminando {ruta}: {e}")
        flash(f"Error eliminando {safe_name}: {e}", "error")
    return redirect(url_for("index"))


# ----------------- Ruta para integrar un insumo seleccionado -----------------
@app.route("/integrar", methods=["POST"])
def integrar():
    """
    Integra un archivo subido al maestro.
    Form data:
      - tipo_insumo: 'endpoint'|'personnel'|'tmp'|'da'
      - filename: nombre del archivo (de uploads)
      - keep_rows: opcional (int)
    """
    tipo = request.form.get("tipo_insumo")
    filename = request.form.get("filename")
    try:
        keep_rows = int(request.form.get("keep_rows", 2))
    except Exception:
        keep_rows = 2

    if not tipo or not filename:
        flash("Debes seleccionar tipo de insumo y archivo.", "error")
        return redirect(url_for("index"))

    src = UPLOAD_FOLDER / secure_filename(filename)
    if not src.exists():
        flash("Archivo fuente no encontrado en uploads.", "error")
        return redirect(url_for("index"))

    maestro = find_master_file()
    if maestro is None:
        flash("No se encontró el archivo maestro (Reporte-Data_Inventario). Colócalo en uploads.", "error")
        return redirect(url_for("index"))

    # hacemos copia de seguridad del maestro si existe la utilidad
    if backup_file:
        try:
            backup_file(maestro)
            append_history({"accion": "backup_maestro", "archivo": os.path.basename(maestro), "timestamp": datetime.now().isoformat()})
        except Exception:
            app.logger.exception("No se pudo crear backup del maestro (backup_file falló).")

    # ejecutar integración según tipo (comprobando que la función exista)
    resultado = None
    try:
        if tipo == "endpoint":
            if not integrate_endpoint_to_antivirus:
                raise RuntimeError("Función de integración 'integrate_endpoint_to_antivirus' no está disponible.")
            resultado = integrate_endpoint_to_antivirus(maestro, str(src), keep_rows=keep_rows)
            msg = f"Integración Antivirus completada. Resultado: {resultado}"
        elif tipo == "personnel":
            if not integrate_personnel_to_estado:
                raise RuntimeError("Función de integración 'integrate_personnel_to_estado' no está disponible.")
            resultado = integrate_personnel_to_estado(maestro, str(src), keep_rows=keep_rows)
            msg = f"Integración Personal completada. Resultado: {resultado}"
        elif tipo == "tmp":
            if not integrate_tmp_to_useraranda:
                raise RuntimeError("Función de integración 'integrate_tmp_to_useraranda' no está disponible.")
            resultado = integrate_tmp_to_useraranda(maestro, str(src), keep_rows=keep_rows)
            msg = f"Integración Useraranda completada. Resultado: {resultado}"
        elif tipo == "da":
            if not integrate_da_to_reporte:
                raise RuntimeError("Función de integración 'integrate_da_to_reporte' no está disponible.")
            resultado = integrate_da_to_reporte(maestro, str(src), keep_rows=keep_rows)
            msg = f"Integración Reporte DA completada. Resultado: {resultado}"
        else:
            flash("Tipo de insumo desconocido.", "error")
            return redirect(url_for("index"))

        # registrar en historial la integración
        append_history({"accion": "integrar", "tipo": tipo, "archivo": filename, "resultado": resultado, "timestamp": datetime.now().isoformat()})
        flash(msg, "success")
    except Exception as e:
        app.logger.exception(f"Error durante integración: {e}")
        flash(f"Error durante integración: {e}", "error")

    return redirect(url_for("index"))


# ----------------- Integrar múltiples insumos (batch) -----------------------
@app.route("/integrar-multiples", methods=["POST"])
def integrar_multiples():
    """
    Recibe listas tipo_insumo[] y filename[] que emparejan por índice.
    Ejecuta integración secuencial y devuelve un resumen agregado.
    """
    tipos = request.form.getlist("tipo_insumo[]") or request.form.getlist("tipo_insumo")
    files = request.form.getlist("filename[]") or request.form.getlist("filename")
    try:
        keep_rows = int(request.form.get("keep_rows", 2))
    except Exception:
        keep_rows = 2

    maestro = find_master_file()
    if maestro is None:
        flash("No se encontró el archivo maestro (Reporte-Data_Inventario).", "error")
        return redirect(url_for("index"))

    resultados = []
    for t, f in zip(tipos, files):
        src = UPLOAD_FOLDER / secure_filename(f)
        if not src.exists():
            resultados.append({"archivo": f, "error": "archivo fuente no encontrado"})
            continue

        try:
            if t == "endpoint":
                if not integrate_endpoint_to_antivirus:
                    raise RuntimeError("Función de integración no disponible (endpoint).")
                res = integrate_endpoint_to_antivirus(maestro, str(src), keep_rows=keep_rows)
            elif t == "personnel":
                if not integrate_personnel_to_estado:
                    raise RuntimeError("Función de integración no disponible (personnel).")
                res = integrate_personnel_to_estado(maestro, str(src), keep_rows=keep_rows)
            elif t == "tmp":
                if not integrate_tmp_to_useraranda:
                    raise RuntimeError("Función de integración no disponible (tmp).")
                res = integrate_tmp_to_useraranda(maestro, str(src), keep_rows=keep_rows)
            elif t == "da":
                if not integrate_da_to_reporte:
                    raise RuntimeError("Función de integración no disponible (da).")
                res = integrate_da_to_reporte(maestro, str(src), keep_rows=keep_rows)
            else:
                res = {"error": "tipo desconocido"}

            resultados.append({"archivo": f, "resultado": res})
            append_history({"accion": "integrar", "tipo": t, "archivo": f, "resultado": res, "timestamp": datetime.now().isoformat()})
        except Exception as e:
            app.logger.exception(f"Error integrando {f}: {e}")
            resultados.append({"archivo": f, "error": str(e)})

    # construir resumen legible y enviar un único flash
    exitos = [r for r in resultados if "error" not in r]
    errores = [r for r in resultados if "error" in r]
    msg_parts = []
    if exitos:
        msg_parts.append(f"Integrados: {', '.join([r['archivo'] for r in exitos])}")
    if errores:
        msg_parts.append(f"Errores: {', '.join([r['archivo'] + ' (' + r['error'] + ')' for r in errores])}")
    if msg_parts:
        flash("; ".join(msg_parts), "info")
    else:
        flash("No se procesaron archivos.", "warning")

    return redirect(url_for("index"))


# ----------------- Descargar maestro / archivos -------------------------
@app.route("/download-maestro")
def download_maestro():
    maestro = find_master_file()
    if not maestro or not Path(maestro).exists():
        flash("No se encontró maestro.xlsx en uploads.", "error")
        return redirect(url_for("index"))
    return send_file(maestro, as_attachment=True)


@app.route("/download/<path:filename>")
def download_file(filename: str):
    """Descarga cualquier archivo del directorio uploads de forma segura."""
    safe = secure_filename(filename)
    p = UPLOAD_FOLDER / safe
    if not p.exists():
        abort(404)
    return send_from_directory(app.config["UPLOAD_FOLDER"], safe, as_attachment=True)


# ----------------- Exportar (Excel / PDF) --------------
@app.route("/exportar-excel")
def exportar_excel():
    maestro = find_master_file()
    if not maestro:
        return ("No hay maestro para exportar", 404)
    return send_file(maestro, as_attachment=True, download_name="inventario.xlsx")


@app.route("/exportar-pdf")
def exportar_pdf():
    maestro = find_master_file()
    if not maestro:
        flash("No hay maestro para exportar.", "error")
        return redirect(url_for("index"))

    # Intento de conversión usando LibreOffice/soffice si está disponible
    soffice_cmd = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_cmd:
        flash("LibreOffice (soffice) no está disponible en el sistema. No es posible convertir a PDF.", "error")
        return redirect(url_for("index"))

    outdir = Path(tempfile.mkdtemp(prefix="maestro_pdf_"))
    try:
        # convertimos (headless)
        subprocess.run([soffice_cmd, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), maestro], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        # archivo de salida esperado
        pdf_name = Path(maestro).with_suffix(".pdf").name
        pdf_path = outdir / pdf_name
        if not pdf_path.exists():
            raise RuntimeError("Conversión completada pero no se encontró el PDF de salida.")
        return send_file(str(pdf_path), as_attachment=True, download_name=pdf_name)
    except subprocess.CalledProcessError as e:
        app.logger.exception("Error en conversión a PDF con soffice: %s", e)
        flash("Error al convertir a PDF (soffice falló). Revisa instalación/logs.", "error")
        return redirect(url_for("index"))
    except Exception as e:
        app.logger.exception("Error exportar_pdf: %s", e)
        flash(f"No fue posible exportar a PDF: {e}", "error")
        return redirect(url_for("index"))
    finally:
        # opcional: no eliminamos inmediatamente el directorio porque send_file
        # podría estar usándolo en algunos servidores; se podría limpiar con un
        # thread/cron si se desea.
        pass


# ----------------- Historial & Configuración -------------------------
@app.route("/historial")
def historial():
    """Página de historial/listado de archivos con paginación simple.
    Parámetros GET:
      - fecha: filtrar por prefijo YYYY-MM-DD (opcional)
      - page: número de página (opcional)
    """
    fecha = request.args.get("fecha")
    archivos = obtener_archivos_historial()

    if fecha:
        archivos = [a for a in archivos if a["fecha"].startswith(fecha)]

    # paginación
    try:
        pagina = int(request.args.get("page", 1))
    except Exception:
        pagina = 1
    por_pagina = 10
    total = len(archivos)
    total_paginas = (total + por_pagina - 1) // por_pagina if total > 0 else 1
    pagina = max(1, min(pagina, total_paginas))
    inicio = (pagina - 1) * por_pagina
    archivos_pagina = archivos[inicio: inicio + por_pagina]

    return render_template("historial.html", archivos=archivos_pagina, fecha_consulta=fecha, page=pagina, total_pages=total_paginas)


@app.route("/configuracion", methods=["GET", "POST"])
def configuracion():
    cfg = load_config()
    if request.method == "POST":
        cfg["empresa_default"] = request.form.get("empresa_default", cfg.get("empresa_default", ""))
        # checkbox -> boolean
        cfg["usar_fecha_archivo"] = True if request.form.get("usar_fecha_archivo") else False
        cfg["fecha_formato"] = request.form.get("fecha_formato", cfg.get("fecha_formato", "%Y-%m-%d"))
        save_config(cfg)
        flash("Configuración guardada.", "success")
        append_history({"accion": "config_update", "timestamp": datetime.now().isoformat()})
        return redirect(url_for("configuracion"))
    return render_template("configuracion.html", config=cfg)


# ----------------------------- Main -------------------------------------
if __name__ == "__main__":
    # abrir navegador en un thread para no bloquear el servidor
    threading.Thread(target=abrir_navegador, args=("http://127.0.0.1:5000",), daemon=True).start()
    # OJO: en producción quita debug=True y usa un WSGI server (gunicorn, waitress, etc.)
    app.run(host="127.0.0.1", port=5000, debug=True)
