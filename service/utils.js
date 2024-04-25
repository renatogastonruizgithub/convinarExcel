

const XlsxPopulate = require('xlsx-populate');


 async function readExcel(file){
    const workbook = await XlsxPopulate.fromDataAsync(file.buffer);
    const sheet = workbook.sheet("REMITO");

 // Encuentra la última fila no vacía en la columna C (3)
 let lastRow = 8; // asumimos que hay al menos una fila de datos
 const columnC = sheet.column("A");
 for (let row = 8; row <= sheet.usedRange().endCell().rowNumber(); row++) {
     if (!columnC.cell(row).value() !== null) {
         lastRow = row;
     }
 }

 // Obtiene los valores desde la fila 8 hasta la última fila no vacía en la columna C
 const range = sheet.range(`A8:M${lastRow}`);

 // Obtiene los valores en el rango especificado
 const values = range.value();
 

const filteredUndefined = values.filter(subarray => subarray.some(item => item !== undefined));

// Función para determinar si un subarray contiene un número y undefined
const contieneNumeroYUndefined = subarray =>
  subarray.some(item => typeof item === 'number' && !Number.isNaN(item)) &&
  subarray.some(item => typeof item === 'undefined');

// Buscar el índice del último subarray que cumple con la condición
const indiceUltimo = filteredUndefined.findIndex(contieneNumeroYUndefined);

// Eliminar el último subarray que cumple con la condición
if (indiceUltimo !== -1) {
  filteredUndefined.splice(indiceUltimo, 1);
}

return filteredUndefined;

} 

async function readIndicadores(file){
    const workbook = await XlsxPopulate.fromDataAsync(file.buffer);
    const sheet = workbook.sheet("REMITO");

 // Encuentra la última fila no vacía en la columna C (3)
 let lastRow = 8; // asumimos que hay al menos una fila de datos
 const columnC = sheet.column("C");
 for (let row = 8; row <= sheet.usedRange().endCell().rowNumber(); row++) {
     if (!columnC.cell(row).value() !== null) {
         lastRow = row;
     }
 }

 // Obtiene los valores desde la fila 8 hasta la última fila no vacía en la columna C
 const range = sheet.range(`C8:L${lastRow}`);

 // Obtiene los valores en el rango especificado
 const values = range.value();
 
 //para indicadores
 values.forEach(value => {
  // Eliminar las posiciones 1, 2, 5 y 6
  value.splice(1, 1); // Elimina la posición 1
  value.splice(1, 1); // Elimina la posición 2 (la nueva posición 2 es la anterior posición 3)
  value.splice(3, 1); // Elimina la posición 5 (la nueva posición 3 es la anterior posición 4)
  value.splice(3, 1); // Elimina la posición 6 (la nueva posición 4 es la anterior posición 5)
 
});  

const filteredUndefined = values.filter(subarray => subarray.some(item => item !== undefined));
for (let i = 0; i < filteredUndefined.length; i++) {
    filteredUndefined[i][2] = ''; // Dejar la posición 2 en vacío
}

return filteredUndefined
} 

async function wrtiteToExcel(data){

    console.log("desde utils"+data)
    const workbook2 = await XlsxPopulate.fromBlankAsync();
    const sheet2 =workbook2.sheet(0); 
    // Escribir los valores en el nuevo documento Excel
    
    for (let i = 0; i < data.length; i++) {
        for (let j = 0; j < data[i].length; j++) {
            sheet2.cell(i + 1, j + 1).value(data[i][j]);
        }
    }
   
    const excelBlob = await workbook2.toFileAsync();

    return excelBlob;
}


module.exports = {  readExcel,  wrtiteToExcel,readIndicadores };