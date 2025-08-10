/*function exportarExcel() {
    alert("Funcionalidad para exportar a Excel no implementada aún.");
}

function exportarPDF() {
    alert("Funcionalidad para exportar a PDF no implementada aún.");
}

// Script listo para extender funcionalidades
console.log("JS cargado correctamente.");*/
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
        })
        .catch(error => {
            console.error(error);
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
        })
        .catch(error => {
            console.error(error);
            alert('No se pudo exportar a PDF.');
        });
}

console.log("JS cargado correctamente.");
