const express = require('express');
const app = express();
const multer = require('multer')
const mimeTypes = require('mime-types');
const path = require('path');
const XLSX = require('xlsx');
let route = '';


const storage = multer.diskStorage({
    destination: 'uploads/',
    filename: function (req,file,cb) {
       
       route = (__dirname + `${path.sep}uploads${path.sep}${file.originalname}`);       
       
       cb("",file.originalname); 
            
    }
    
})

function leerExcel(ruta){
    const workbook = XLSX.readFile(ruta);
    const workbookSheets = workbook.SheetNames;
    const sheet = workbookSheets[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet]);
    return data;
}


const upload = multer ({
   storage: storage 
   
})

app.get("/",(req,res)=> {
    res.sendFile(__dirname + "/views/index.html");

})

app.post('/files',upload.single('avatar'), (req,res)=> {
 
 //leerExcel(route);
 res.send(leerExcel(route));

})

app.listen(3100,() => console.log("Server iniciado"));

