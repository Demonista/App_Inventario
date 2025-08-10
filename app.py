import os
from flask import Flask, render_template, request, redirect, url_for
from werkzeug.utils import secure_filename
from datetime import datetime
import threading
import subprocess
import platform

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

archivos_cargados = []

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def detectar_tipo_insumo(nombre):
    nombre = nombre.lower()
    if "inventario" in nombre or "insumo1" in nombre:
        return "Inventario Proveedor"
    elif "antivirus" in nombre or "insumo2" in nombre:
        return "Antivirus"
    elif "retiros" in nombre or "personal" in nombre or "ingresos" in nombre or "colaboradores" in nombre or "insumo3" in nombre:
        return "Talento Humano"
    elif "Directorio Activo" in nombre or "insumo5" in nombre:
        return "Directorio Activo"
    elif "aranda" in nombre or "insumo4" in nombre:
        return "Aranda"
    return "Desconocido"

@app.route('/', methods=['GET', 'POST'])
def index():
    global archivos_cargados
    if request.method == 'POST':
        archivos = request.files.getlist('files[]')
        for archivo in archivos:
            if archivo and allowed_file(archivo.filename):
                filename = secure_filename(archivo.filename)
                save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                archivo.save(save_path)
                archivos_cargados.append({
                    'nombre': filename,
                    'fecha': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'tipo': detectar_tipo_insumo(filename),
                    'cargado': True
                })
    return render_template('index.html', archivos=archivos_cargados)

@app.route('/eliminar/<nombre_archivo>')
def eliminar(nombre_archivo):
    global archivos_cargados
    archivos_cargados = [a for a in archivos_cargados if a['nombre'] != nombre_archivo]
    try:
        os.remove(os.path.join(app.config['UPLOAD_FOLDER'], nombre_archivo))
    except:
        pass
    return redirect(url_for('index'))

def abrir_navegador():
    url = "http://127.0.0.1:5000"
    if platform.system() == "Windows":
        subprocess.Popen(['start', '', url], shell=True)
    else:
        subprocess.Popen(['xdg-open', url])  # Linux/macOS

if __name__ == '__main__':
    threading.Timer(1.0, abrir_navegador).start()
    app.run(debug=False)
