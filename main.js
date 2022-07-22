const XlsxPopulate = require('xlsx-populate');
const path = require('path');
const fs = require('fs');
var XLSX = require("xlsx");

const DIRECTORYNAME = "" //aqui el nombre del directorio donde estan los excels
const celda = "C48" //aqui la celda a cambiar
const mensaje = "SUPERCUENTA" //aqui el mensaje que reemplaza a lo anterior

//joining path of directory 
const directoryPath = path.join(__dirname, DIRECTORYNAME);
//passsing directoryPath and callback function


//GET LIST OF SUBDIRECTORIES
let subdirectories = []
fs.readdir(directoryPath, function (err, files) {
    //handling error
    if (err) {
        return console.log('Unable to scan directory: ' + err);
    } 
    //listing all files using forEach
    subdirectories = files

    files.forEach((file)=>{
      if(checkIfIsFolder(path.join(__dirname, DIRECTORYNAME), file)){
        console.log(file, "lista de carpetas")
        let subRoute = path.join(__dirname, DIRECTORYNAME, file)
        getListOfFilesInSubdirectory(subRoute)
      }
    })


});

const checkIfIsFolder = (route, fileName) => {
  return fs.statSync(path.join(route, fileName)).isDirectory() && fileName != ".git" && fileName != "node_modules"
}


//for each subdirectory list of files

const getListOfFilesInSubdirectory = async (subdirectoryName) =>{
  let insideFiles = []
  await fs.readdir(subdirectoryName, function (err, insideFiles) {
    //handling error
    if (err) {
        return console.log('Unable to scan directory: ' + err);
    } 
    console.log("insidefiles", insideFiles)
    insideFiles.forEach((file)=>{
      let extension = file.split(".")[file.split(".").length - 1] || "nope"
      console.log("file", file, "extension", extension)
      if(extension==="xlsx"){
        console.log("modificando", file, "es XLSX")
        modifyXLSX(path.join(subdirectoryName, file))
        
      }if(extension==="xls"){
        console.log("modificando", file, "es XLS")
        modifyXLS(path.join(subdirectoryName, file))
      }
    })
  
});
}

//MODIFY EXCEL FUNCTION
const modifyXLSX = (route)=> XlsxPopulate.fromFileAsync(route)
.then(workbook => {
    // Modify the workbook.
    let thisSheet = workbook.sheet("proforma")
    thisSheet.cell(celda).value(mensaje)
    return workbook.toFileAsync(route)
})


const modifyXLS = (route) => {

  const xlsFilePath = route //a modificar
  const xlsxFilePath = path.join(__dirname, 'base.xlsx') //archivo base

  //CONSEGUIR DATOS DEL ARCHIVO A MODIFICAR 

  let workbook = XLSX.readFile(xlsFilePath)
  let proformaSheetXLS = workbook.Sheets["proforma"]
  let XLSkeys = Object.keys(proformaSheetXLS)

  //ESCRIBIR DATOS EN EL ARCHIVO BASE

  XlsxPopulate.fromFileAsync(xlsxFilePath)
  .then(XLSXworkbook => {
      // Modify the XLSXworkbook.
      let thisSheet = XLSXworkbook.sheet("proforma")

      XLSkeys.forEach((key)=>{
          if(key.length<=3 && !proformaSheetXLS[key].f){
              thisSheet.cell(key).value(proformaSheetXLS[key].v)
          }
      })

      thisSheet.cell(celda).value(mensaje)

      //GUARDAR ARCHIVO CON NOMBRE ANTIGUO ACABADO EN .XLS

      XLSXworkbook.toFileAsync(`${xlsFilePath}.xlsx`)

      // CARGARME EL ARCHIVO ANTIGUO

      eraseFile(xlsFilePath)
  })

  const eraseFile = (pathToArchivo) => {
      try {
          fs.unlinkSync(pathToArchivo)
          } catch(err) {
          console.error(err)
          }
  }


}

/* 
  ANTIGUO MODIGYXLS
  console.log(route)
  let workbook = XLSX.readFile(route)
  let worksheet = workbook.Sheets["proforma"];
  let celda = "C6"
  let mensaje = "nueva Cuenta corriente"

  console.log(worksheet['C6'].v)
  if(worksheet[celda].v){
    worksheet[celda].v = mensaje
  }else{
    XLSX.utils.sheet_add_aoa(worksheet, [[mensaje]], {origin: celda})
  }
  

  XLSX.writeFile(workbook, `${route}2.xls`);

  console.log("read")

*/