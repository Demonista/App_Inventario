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
// ------------------------------------------------------
function integrarArchivos() {
  if (!confirm("🔗 ¿Deseas integrar todos los insumos al maestro?")) return;

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
      location.reload();
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
      window.open('/exportar-excel', '_blank');
    });
}

// ------------------------------------------------------
// 4. Exportar inventario a PDF (todo o 30 filas)
// ------------------------------------------------------
function exportarPDF(todo = true) {
  // Si todo = true → exporta todas las filas
  // Si todo = false → exporta solo 30 filas
  const url = todo ? '/exportar-pdf' : '/exportar-pdf?limit=30';

  fetch(url)
    .then(response => {
      if (!response.ok) throw new Error('Error en la exportación a PDF');
      return response.blob();
    })
    .then(blob => {
      const nombreArchivo = todo ? "inventario_completo.pdf" : "inventario_30filas.pdf";
      descargarArchivo(blob, nombreArchivo);
    })
    .catch(err => {
      console.error("❌ Error en exportación PDF:", err);
      window.open(url, '_blank');
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
// 7. Debug en consola
// ------------------------------------------------------
console.log("✅ scripts.js cargado correctamente.");

// ======================================================
// Funciones de soporte para la app Flask de Inventario
// ======================================================

// ------------------------------------------------------
// 8. Configuración de validaciones dinámicas
// ------------------------------------------------------

// Contador global de reglas
let reglaCount = document.querySelectorAll('.regla-block').length || 0;

/**
 * Genera dinámicamente el HTML de una condición
 * @param {number} i - Índice de la regla a la que pertenece
 */
function filaCondicionHTML(i) {
  let opciones = "";

  // Solo cargamos columnas si existen
  if (window.columnasDisponibles && Array.isArray(window.columnasDisponibles)) {
    opciones = window.columnasDisponibles
      .map(col => `<option value="${col}">${col}</option>`)
      .join("");
  } else {
    opciones = `<option disabled>(No hay columnas disponibles)</option>`;
  }

  return `
    <tr>
      <td>
        <select name="cond_col_${i}[]" required>
          ${opciones}
        </select>
      </td>
      <td>
        <select name="cond_op_${i}[]" onchange="actualizarCampoValor(this)">
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
      <td>
        <input type="text" name="cond_val_${i}[]" placeholder="Ej: No instalado">
      </td>
      <td>
        <button type="button" class="btn rojo" onclick="eliminarCondicion(this)">❌</button>
      </td>
    </tr>
  `;
}

/**
 * Agrega una nueva regla (bloque con condiciones)
 */
function agregarRegla() {
  const container = document.getElementById('reglas-container');
  const i = reglaCount++;

  const div = document.createElement('div');
  div.className = 'regla-block';
  div.dataset.reglaIndex = i;

  div.innerHTML = `
    <div style="display:flex; gap:8px; align-items:center; margin-bottom:6px;">
      <input type="text" name="etiquetas[]" placeholder="Etiqueta (ej: Riesgo Alto)" required>
      <select name="logic[]">
        <option value="AND">AND</option>
        <option value="OR">OR</option>
      </select>
      <button type="button" class="btn rojo" onclick="eliminarRegla(this)">❌ Eliminar regla</button>
    </div>

    <table class="cond-table">
      <thead>
        <tr>
          <th>Columna</th>
          <th>Operador</th>
          <th>Valor</th>
          <th>Acción</th>
        </tr>
      </thead>
      <tbody>
        ${filaCondicionHTML(i)}
      </tbody>
    </table>

    <div style="margin-top:4px;">
      <button type="button" class="btn amarillo" onclick="agregarCondicion(${i})">➕ Añadir condición</button>
    </div>
    <hr>
  `;

  container.appendChild(div);

  // Inicializar la primera condición
  const select = div.querySelector("select[name^='cond_op_']");
  if (select) actualizarCampoValor(select);
}

/**
 * Agregar condición dentro de una regla existente
 * @param {number} i - Índice de la regla
 */
function agregarCondicion(i) {
  const block = document.querySelector(`.regla-block[data-regla-index="${i}"]`);
  if (!block) return;

  const tbody = block.querySelector('table.cond-table tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = filaCondicionHTML(i);
  tbody.appendChild(tr);

  // Inicializar campo de valor
  const select = tr.querySelector("select[name^='cond_op_']");
  if (select) actualizarCampoValor(select);
}

/**
 * Elimina una regla completa
 */
function eliminarRegla(btn) {
  const block = btn.closest('.regla-block');
  if (block) block.remove();
}

/**
 * Elimina una condición específica
 */
function eliminarCondicion(btn) {
  const tr = btn.closest('tr');
  if (tr) tr.remove();
}

/**
 * Desactiva/activa campo de valor según operador
 */
function actualizarCampoValor(select) {
  const fila = select.closest("tr");
  const inputValor = fila.querySelector("input[name^='cond_val_']");

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

// ======================================================
// 9. Tema oscuro / claro
// ======================================================
function toggleDarkMode() {
  document.body.classList.toggle("dark-mode");

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

  // Inicializar operadores existentes
  document.querySelectorAll("select[name^='cond_op_']").forEach(sel => {
    actualizarCampoValor(sel);
  });

  // Reset contador de reglas
  const blocks = document.querySelectorAll('.regla-block');
  reglaCount = blocks.length;
});

