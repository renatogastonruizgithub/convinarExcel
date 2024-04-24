document.addEventListener('DOMContentLoaded', function() {
    // Obtener el botón de descarga
    let downloadButton = document.getElementById('downloadButton');
    // Escuchar el evento de clic en el botón de descarga
    downloadButton.addEventListener('click', async function() {
        try {
            // Realizar una solicitud GET al servidor para descargar el archivo
            let response = await fetch('/download');
            let excelBlob = await response.blob();
            
            // Crear un objeto URL para el Blob recibido
            let url = window.URL.createObjectURL(excelBlob);
            
            // Crear un enlace temporal y asignarle el objeto URL
            let a = document.createElement('a');
            a.href = url;
            a.download = 'nuevo.xlsx'; // Nombre del archivo a descargar
            
            // Simular un clic en el enlace para iniciar la descarga
            a.click();
            
            // Liberar el objeto URL cuando ya no sea necesario
            window.URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Error al descargar el archivo:', error);
            alert('Error al descargar el archivo');
        }
    });
});
// Escuchar el evento de clic en el botón de descarga
$('#downloadButton').on('click', function() {
    // Realizar una solicitud GET al servidor para descargar el archivo
    $.ajax({
        url: '/download',
        method: 'GET',
        xhrFields: {
            responseType: 'blob' // Especificar que esperamos una respuesta de tipo Blob
        },
        success: function(data) {
            // Crear un objeto URL para el Blob recibido
            var url = window.URL.createObjectURL(data);
            // Crear un enlace temporal y asignarle el objeto URL
            var a = document.createElement('a');
            a.href = url;
            a.download = 'nuevo.xlsx'; // Nombre del archivo a descargar
            // Simular un clic en el enlace para iniciar la descarga
            a.click();
            // Liberar el objeto URL cuando ya no sea necesario
            window.URL.revokeObjectURL(url);
        },
        error: function(xhr, status, error) {
            console.error('Error al descargar el archivo:', error);
        }
    });
});

document.getElementById('copyButton').addEventListener('click', function() {
    copyTableToClipboard('dataTable')
    alert('copyButton')
});

function copyTableToClipboard(tableId) {
    const table = document.getElementById(tableId);
    const range = document.createRange();
    range.selectNode(table);
    window.getSelection().removeAllRanges();
    window.getSelection().addRange(range);
    document.execCommand('copy');
    window.getSelection().removeAllRanges();
    ajustarAlturaFilas(table); // Ajustar la altura de las filas después de pegar
}


function ajustarAlturaFilas(table) {
    // Definir el alto deseado para las filas (en píxeles)
    const nuevoAlto = '12px';

    // Iterar sobre las filas y establecer el nuevo alto
    const filas = table.querySelectorAll('tr');
    filas.forEach(fila => fila.style.height = nuevoAlto);
}

document.getElementById('volver').addEventListener('click', function() {
   
    window.location.href = '/';
});
