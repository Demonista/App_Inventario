// static/js/scripts.js
// ================================
// Funciones de soporte para la app Flask de Inventario
// Incluye navegación al historial con filtro de fecha
// y exportación de datos a Excel y PDF
// ================================

// Ir al historial, filtrando por fecha si se selecciona
function irAlHistorial() {
  const fecha = document.getElementById('fecha_consulta').value;
  let url = '/historial';
  if (fecha) {
    url += '?fecha=' + encodeURIComponent(fecha);
  }
  window.location.href = url;
}

// -------------------------------
// Exportar inventario a Excel
// -------------------------------
function exportarExcel() {
  fetch('/exportar-excel')
    .then(response => {
      if (!response.ok) throw new Error('Error en la exportación a Excel');
      return response.blob(); // respuesta como archivo binario
    })
    .then(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'inventario.xlsx'; // nombre sugerido
      document.body.appendChild(a);
      a.click(); // forzar descarga
      a.remove();
      window.URL.revokeObjectURL(url); // liberar memoria
    })
    .catch(err => {
      console.error(err);
      alert('❌ No se pudo exportar a Excel.');
    });
}

// -------------------------------
// Exportar inventario a PDF
// -------------------------------
function exportarPDF() {
  fetch('/exportar-pdf')
    .then(response => {
      if (!response.ok) throw new Error('Error en la exportación a PDF');
      return response.blob(); // respuesta como archivo binario
    })
    .then(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'inventario.pdf'; // nombre sugerido
      document.body.appendChild(a);
      a.click(); // forzar descarga
      a.remove();
      window.URL.revokeObjectURL(url); // liberar memoria
    })
    .catch(err => {
      console.error(err);
      alert('❌ No se pudo exportar a PDF.');
    });
}

// -------------------------------
// Confirmación antes de eliminar archivo
// (esto protege de clics accidentales en "Eliminar")
// -------------------------------
function confirmarEliminacion(nombreArchivo) {
  return confirm(`⚠️ ¿Seguro que deseas eliminar el archivo "${nombreArchivo}"?`);
}

// -------------------------------
// Mensaje de consola para depuración
// -------------------------------
console.log("✅ JS cargado correctamente.");
