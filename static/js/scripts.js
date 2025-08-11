// static/js/scripts.js
// Funciones para exportar archivos (llaman a las rutas Flask)
function exportarExcel() {
  fetch('/exportar-excel')
    .then(response => {
      if (!response.ok) throw new Error('Error en la exportación a Excel');
      return response.blob();
    })
    .then(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'inventario.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    })
    .catch(err => {
      console.error(err);
      alert('No se pudo exportar a Excel.');
    });
}

function exportarPDF() {
  fetch('/exportar-pdf')
    .then(response => {
      if (!response.ok) throw new Error('Error en la exportación a PDF');
      return response.blob();
    })
    .then(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'inventario.pdf';
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    })
    .catch(err => {
      console.error(err);
      alert('No se pudo exportar a PDF.');
    });
}

console.log("JS cargado correctamente.");
