const express = require('express');
const app = express();
const multer = require('multer')
const mimeTypes = require('mime-types');
const path = require('path');
const XLSX = require('xlsx');
const {config } = require('./config/config');
const getConnection = require('./libs/postgres');
const { isNumber } = require('util');
let route = '';
let filas =[];




const storage = multer.diskStorage({
    destination: 'uploads/',
    filename: function (req,file,cb) {
       
       route = (__dirname + `${path.sep}uploads${path.sep}${file.originalname}`);       
       
       cb("",file.originalname); 
            
    }
    
})



async function leerExcel(ruta){
    
    const workbook = XLSX.readFile(ruta);      
    const workbookSheets = workbook.SheetNames;    
    const sheet = workbookSheets[0];    
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {defval:'N/A'});   
    const obj = Object.keys(data[0]);
    const word = 'fecha';
    let rowdate = [];
    let strmin = '';
    
       
        
        await DeleteTable(sheet);
        
        
        let acum = '';
         obj.forEach((indice,valor) => {
                strmin = indice.toLowerCase();                

                if (strmin.includes(word)){
                   
                    rowdate.push(valor);
                }

                acum = acum + ',' + '"' + indice.split("\n").join("").trim() + '"';

        
            });
       
        acum = '(' + acum.slice(1) + ')';  
      


    
        
          let cadena = '';      
          let str = '';
          
         

        for(let i=0; i < data.length; i++) {
          
         
          const map1 = (Object.values(data[i]).map(x => x));
           
             
               for(let indice=0; indice < map1.length; indice++) {
                    

                      if (!isNaN(map1[indice])) {
                        
                          if (rowdate.length === 0) {
                            map1[indice] = `${((map1[indice]).toString())}`;
                           
                        } else { 
                           
                           if (rowdate.includes(indice)){                            
                            
                           map1[indice] = `'${new Date((map1[indice] - (25567 + 1)) * 86400 * 1000).toLocaleDateString()}'`;
                         
                           }
                        }
                         
                        
                         } else {
                          map1[indice] = map1[indice].replaceAll("'", "''");
                          map1[indice] = map1[indice].replaceAll('"', " ");
                          map1[indice] = `'${map1[indice]}'`
                         }
                           
                 }    
                 
                
                  str = (map1.join());              


              
          
              
                         
          
        


            //construiremos la cadena string insert into y el ciclo iterara por todos los valores que tenga las filas de excel
              cadena = (`Insert into comercial."${sheet}" ${acum} Values (${str})`);
              
               await incluir(cadena,i);
             
              
         };
         
        
    
    return obj;
}



async function incluir(cadena, indice){ 
    const client = await getConnection();  
   


    
   try {
    const rta= await client.query(cadena);
    

   } catch (error) {
    // aqui debe ir un arreglo con el indice // para indicar aquellas celdas que no se pudieron agregar
      console.log(cadena);     
      filas.push(indice);

   }
                 
   client.end();
  
   
}

async function DeleteTable (Tabla){  // revisa si ya se mando un mensaje al usuario aprobador (no lo puede volver a enviar)
    
    const client = await getConnection();   
   
    const rta= await client.query(`Delete FROM comercial."${Tabla}"`); 
   
    
     client.end();
     
}

const upload = multer ({
  
   storage: storage 
   
})

app.get("/",(req,res)=> {
 
    res.sendFile(__dirname + "/views/index.html");

})

app.post('/files',upload.single('avatar'), async (req,res,next)=> {

 const file = req.file

 if (!file)  {
    const error = new Error ('Please Upload a file')
    error.httpStatusCode = 400;
    res.status(400).json({
    message:'Bad request.Error 400. Seleccione un Archivo.'
    })
    } else {
      

        
      await leerExcel(route);

       if (filas.length > 0) {
        res.status(200).json({
          message:`Ok request.Status:200...Fin del proceso.Revise las siguientes filas`,
          row: filas        
        });
       } else {
         res.status(200).json({
          message:`Ok request.Status:200...Fin del proceso. Todas las filas procesadas`,
          row: filas       
        });
        filas = [];
       } 
        
              
          
        
    
       
    }

 

})

app.listen(config.port,() => console.log("Server iniciado"));

