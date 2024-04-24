const express = require('express');
const app = express();
const multer = require('multer');
const {  readExcel , wrtiteToExcel,readIndicadores } = require('./service/utils');
const path= require("path")



const port = 3000;


app.set('views',path.join(__dirname,'views'))

app.set('view engine', 'ejs');

app.use(express.static("public"));



const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.get('/', (req, res) => {
    res.render('index');
});


app.post('/resultado', upload.array('excelFiles', 10), async (req, res) => {
    
    try {
        const files = req.files;
        let combinedData = [];
        for (const file of files) {
            const values = await readExcel(file);
            combinedData.push(...values); // Agregar datos del archivo actual al array existente
        }
        console.log(combinedData)
      
        res.render('formatoEntrega', { values: combinedData });
        combinedData = [];
       

    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Error al cargar el archivo Excel.');
    }
});

app.post('/entrega', upload.array('fileEntrega', 10), async (req, res) => {
    
    try {
        const files = req.files;
        let combinedData = [];
        for (const file of files) {
            const values = await readIndicadores(file);
            combinedData.push(...values); // Agregar datos del archivo actual al array existente
        }
        console.log(combinedData)
      
        res.render('indicadores', { values: combinedData });
        combinedData = [];
       

    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Error al cargar el archivo Excel.');
    }
});





app.get('/download', async (req, res) => {
    try {
        console.log(combinedData)
         const excelFilePath = path.join(__dirname, 'nuevo.xlsx');
        const excel= await wrtiteToExcel(combinedData);
        /* res.download(excel);   */
        res.setHeader('Content-Disposition', 'attachment; filename="nuevo.xlsx"');      
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.download(excelFilePath);
    } catch (error) {
        console.error('Error al generar el archivo Excel:', error);
        res.status(500).send('Error al generar el archivo Excel.');
    }
});

app.listen(port, () => {
    console.log(`Servidor escuchando en http://localhost:${port}`);
});