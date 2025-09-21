// static/js/scripts.js
// ======================================================
// Funciones de soporte para la app Flask de Inventario
// Incluye:
// - Navegación al historial (con filtro por fecha)
// - Integración de insumos con maestro
// - Exportación a Excel y PDF
// - Confirmación de eliminación de archivos
// - Dinámica de reglas y validaciones (Insumo 1)
// - Utilidades generales para interfaz
// ======================================================

// ------------------------------------------------------
// 1. Ir al historial con filtro por fecha
// ------------------------------------------------------
function irAlHistorial() {
  const fechaInput = document.getElementById('fecha_consulta');
  let url = '/historial';

  if (fechaInput && fechaInput.value) {
    url += '?fecha=' + encodeURIComponent(fechaInput.value);
  }

  window.location.href = url;
}

// ------------------------------------------------------
// 2. Integrar archivos subidos con el maestro
//    -> Envía POST a Flask
// ------------------------------------------------------
function integrarArchivos() {
  if (!confirm("🔗 ¿Deseas integrar todos los insumos al maestro?")) {
    return;
  }

  fetch('/integrar', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' }
  })
    .then(response => {
      if (!response.ok) throw new Error('Error en la integración');
      return response.json();
    })
    .then(data => {
      alert("✅ Integración completada: " + (data.mensaje || "Proceso finalizado"));
      location.reload(); // refresca resultados
    })
    .catch(err => {
      console.error("❌ Error en integración:", err);
      alert("❌ Error durante la integración de insumos.");
    });
}

// ------------------------------------------------------
// 3. Exportar inventario a Excel
// ------------------------------------------------------
function exportarExcel() {
  fetch('/exportar-excel')
    .then(response => {
      if (!response.ok) throw new Error('Error en la exportación a Excel');
      return response.blob();
    })
    .then(blob => {
      descargarArchivo(blob, "inventario.xlsx");
    })
    .catch(err => {
      console.error("❌ Error en exportación Excel:", err);
      alert("❌ No se pudo exportar a Excel. Intentando descarga directa...");
      window.open('/exportar-excel', '_blank');
    });
}

// ------------------------------------------------------
// 4. Exportar inventario a PDF
// ------------------------------------------------------
function exportarPDF() {
  fetch('/exportar-pdf')
    .then(response => {
      if (!response.ok) throw new Error('Error en la exportación a PDF');
      return response.blob();
    })
    .then(blob => {
      descargarArchivo(blob, "inventario.pdf");
    })
    .catch(err => {
      console.error("❌ Error en exportación PDF:", err);
      alert("❌ No se pudo exportar a PDF. Intentando descarga directa...");
      window.open('/exportar-pdf', '_blank');
    });
}

// ------------------------------------------------------
// 5. Utilidad genérica para descarga de archivos
// ------------------------------------------------------
function descargarArchivo(blob, nombreArchivo) {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = nombreArchivo;
  document.body.appendChild(a);
  a.click();
  a.remove();
  window.URL.revokeObjectURL(url);
}

// ------------------------------------------------------
// 6. Confirmación antes de eliminar archivo
// ------------------------------------------------------
function confirmarEliminacion(nombreArchivo) {
  return confirm(`⚠️ ¿Seguro que deseas eliminar el archivo "${nombreArchivo}"?`);
}

// ------------------------------------------------------
// 7. Debug en consola (para verificar carga de scripts)
// ------------------------------------------------------
console.log("✅ scripts.js cargado correctamente.");

// ------------------------------------------------------
// 8. Configuración de validaciones (tabla simple)
// ------------------------------------------------------
function agregarFilaValidacion() {
  const tbody = document.getElementById("tabla-config");
  if (!tbody) return;

  const fila = document.createElement("tr");
  fila.innerHTML = `
    <td><input type="text" name="etiquetas[]" placeholder="Ej: SIN ANTIVIRUS" required></td>
    <td><input type="text" name="columnas[]" placeholder="Ej: Antivirus" required></td>
    <td>
      <select name="operadores[]" onchange="actualizarCampoValor(this)" required>
        <option value="=">=</option>
        <option value="!=">≠</option>
        <option value="contiene">Contiene</option>
        <option value="no_contiene">No contiene</option>
        <option value=">">&gt;</option>
        <option value="<">&lt;</option>
        <option value="es_vacio">Está vacío</option>
        <option value="no_vacio">No está vacío</option>
      </select>
    </td>
    <td><input type="text" name="valores[]" placeholder="Ej: No instalado"></td>
    <td style="text-align: center;">
      <button type="button" class="btn rojo" onclick="eliminarFilaValidacion(this)">❌</button>
    </td>
  `;
  tbody.appendChild(fila);
}

function eliminarFilaValidacion(boton) {
  const tbody = document.getElementById("tabla-config");
  if (tbody && tbody.rows.length > 1) {
    boton.closest("tr").remove();
  } else {
    alert("Debe existir al menos una regla configurada.");
  }
}

function actualizarCampoValor(select) {
  const fila = select.closest("tr");
  const inputValor = fila.querySelector("input[name='valores[]']");
  
  const operadoresSinValor = ["es_vacio", "no_vacio"];
  if (operadoresSinValor.includes(select.value)) {
    inputValor.value = "";
    inputValor.disabled = true;
    inputValor.placeholder = "N/A";
  } else {
    inputValor.disabled = false;
    inputValor.placeholder = "Ej: No instalado";
  }
}

// Inicializar validaciones al cargar página
document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll("select[name='operadores[]']").forEach(sel => {
    actualizarCampoValor(sel);
  });
});

// ------------------------------------------------------
// 9. Reglas avanzadas (bloques dinámicos con condiciones)
// ------------------------------------------------------
let reglaCount = document.querySelectorAll('.regla-block').length || 0;

function agregarRegla() {
  const container = document.getElementById('reglas-container');
  const i = reglaCount++;
  const div = document.createElement('div');
  div.className = 'regla-block';
  div.dataset.reglaIndex = i;
  div.innerHTML = `
    <div style="display:flex; gap:8px; align-items:center;">
      <input type="text" name="etiquetas[]" placeholder="Etiqueta" required>
      <select name="logic[]">
        <option value="AND">AND</option>
        <option value="OR">OR</option>
      </select>
      <button type="button" onclick="eliminarRegla(this)">Eliminar regla</button>
    </div>
    <table class="cond-table">
      <thead><tr><th>Columna</th><th>Operador</th><th>Valor</th><th>Acción</th></tr></thead>
      <tbody>
        <tr>
          <td><input type="text" name="cond_col_${i}[]" required></td>
          <td>
            <select name="cond_op_${i}[]">
              <option value="=">=</option>
              <option value="!=">≠</option>
              <option value="contiene">Contiene</option>
              <option value="no_contiene">No contiene</option>
              <option value=">">&gt;</option>
              <option value="<">&lt;</option>
              <option value="es_vacio">Está vacío</option>
              <option value="no_vacio">No está vacío</option>
            </select>
          </td>
          <td><input type="text" name="cond_val_${i}[]"></td>
          <td><button type="button" onclick="eliminarCondicion(this)">❌</button></td>
        </tr>
      </tbody>
    </table>
    <div>
      <button type="button" onclick="agregarCondicion(${i})">➕ Agregar condición</button>
    </div>
    <hr>
  `;
  container.appendChild(div);
}

function eliminarRegla(btn) {
  const block = btn.closest('.regla-block');
  if (block) block.remove();
}

function agregarCondicion(i) {
  const block = document.querySelector(`.regla-block[data-regla-index="${i}"]`);
  if (!block) return;
  const tbody = block.querySelector('table.cond-table tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" name="cond_col_${i}[]" required></td>
    <td>
      <select name="cond_op_${i}[]">
        <option value="=">=</option>
        <option value="!=">≠</option>
        <option value="contiene">Contiene</option>
        <option value="no_contiene">No contiene</option>
        <option value=">">&gt;</option>
        <option value="<">&lt;</option>
        <option value="es_vacio">Está vacío</option>
        <option value="no_vacio">No está vacío</option>
      </select>
    </td>
    <td><input type="text" name="cond_val_${i}[]"></td>
    <td><button type="button" onclick="eliminarCondicion(this)">❌</button></td>
  `;
  tbody.appendChild(tr);
}

function eliminarCondicion(btn) {
  const tr = btn.closest('tr');
  if (tr) tr.remove();
}

// Inicializar contador de reglas si ya existen
document.addEventListener('DOMContentLoaded', function() {
  const blocks = document.querySelectorAll('.regla-block');
  reglaCount = blocks.length;
});

// ==============================
// Tema oscuro / claro
// ==============================
function toggleDarkMode() {
    document.body.classList.toggle("dark-mode");

    // Guardar preferencia en localStorage
    if (document.body.classList.contains("dark-mode")) {
        localStorage.setItem("theme", "dark");
    } else {
        localStorage.setItem("theme", "light");
    }
}

// Al cargar la página, aplicar preferencia guardada
document.addEventListener("DOMContentLoaded", () => {
    if (localStorage.getItem("theme") === "dark") {
        document.body.classList.add("dark-mode");
    }
});
