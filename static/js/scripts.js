// static/js/scripts.js
// ======================================================
// Funciones de soporte para la app Flask de Inventario
// Incluye:
// - Navegación al historial (con filtro por fecha)
// - Integración de insumos con maestro
// - Exportación a Excel y PDF
// - Confirmación de eliminación de archivos
// - Utilidades para interfaz y mensajes
// ======================================================

// ------------------------------------------------------
// Ir al historial, con opción de filtrar por fecha
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
// Integrar archivos subidos con el maestro
// (manda petición POST al backend Flask)
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
      location.reload(); // refresca para mostrar resultados actualizados
    })
    .catch(err => {
      console.error("❌ Error en integración:", err);
      alert("❌ Error durante la integración de insumos.");
    });
}

// ------------------------------------------------------
// Exportar inventario a Excel
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
      alert("❌ No se pudo exportar a Excel.");
    });
}

// ------------------------------------------------------
// Exportar inventario a PDF
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
      alert("❌ No se pudo exportar a PDF.");
    });
}

// ------------------------------------------------------
// Utilidad para descargar archivo desde blob
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
// Confirmación antes de eliminar archivo
// (protección contra clics accidentales)
// ------------------------------------------------------
function confirmarEliminacion(nombreArchivo) {
  return confirm(`⚠️ ¿Seguro que deseas eliminar el archivo "${nombreArchivo}"?`);
}

// ------------------------------------------------------
// Mensaje en consola para depuración
// ------------------------------------------------------
console.log("✅ scripts.js cargado correctamente.");

// ------------------------------------------------------
// Dinámica de configuración de validaciones (Insumo 1)
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
  
  // Operadores que NO requieren valor manual
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

// Inicializar al cargar la página (en caso de reglas ya guardadas)
document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll("select[name='operadores[]']").forEach(sel => {
    actualizarCampoValor(sel);
  });
});

