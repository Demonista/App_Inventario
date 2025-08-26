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
