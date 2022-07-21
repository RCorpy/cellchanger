const XlsxPopulate = require('xlsx-populate');


console.log("hello")



XlsxPopulate.fromFileAsync("./nombre.xlsx")
.then(workbook => {
    // Modify the workbook.

    thisSheet.cell("C5").value("nueva Cuenta corriente")

    return workbook.toFileAsync("./nombre.xlsx")
})

/*
const path = require("path");
const express = require("express");
const app = express(); // create express app
const XlsxPopulate = require('xlsx-populate');
const openMyFile = require('open');

const userName = require("./usuario").modules.userName
const customPath = require("./usuario").modules.customPath
const listadoPath = require("./usuario").modules.listadoPath

const PRECIO_PORTES_POR_KG = 0.33

const PRIMERA_FILA_CONCEPTOS = 18
const ULTIMA_FILA_CONCEPTOS = 43
const BLUE = "0080FF"
const WHITE = "ffffff"

const texteator = {
  "epoxy brillo":{
    name: "EPOXY"
  },
  "epoxy mate":{
    name: "EPOXY MATE"
  },
  "acrilica":{
    name: "ACRILICA"
  },
  "politop":{
    name: "POLITOP"
  },
  "epoxy am":{
    name: "EPOXY AM"
  }
}


// add middlewares
app.use(express.static(path.join(__dirname, "..", "build")));
app.use(express.static("public"));
app.use(express. json())

app.post('/createexcel', (req, res)=>{

  const data = {...req.body.token}
  const thisFileName = data.thisFileName

  let hasHarina = false
  let subExtras = data.extras.map((element)=>{
    if(element.name.includes("harina")){
      hasHarina = element
    }
    else{
      return element
    }
  }).filter((ele)=>(ele))

  console.log("hi Rigga", data)
  console.log("extras", subExtras, hasHarina)

  // Load an existing workbook
  XlsxPopulate.fromFileAsync("./basePresupuestosExcel.xlsx")
    .then(workbook => {
        // Modify the workbook.
        let today = new Date()
        
        let thisSheet = workbook.sheet("proforma")

        let totalKgs = parseFloat(data.kgsImprimacion)+parseFloat(data.kgsCapas)+parseInt(data.resultData.disolvente)

        thisSheet.cell("C5").value(data.concepto.toUpperCase())
        thisSheet.cell("C8").value(data.nombre.toLowerCase().split(" ").map(s=>(s.charAt(0).toUpperCase()+s.slice(1))).join(" ")) //Churraken :D
        thisSheet.cell("E4").value(`Nº: ${thisFileName}`)
        thisSheet.cell("B15").value( `${data.m2}m²`)
        thisSheet.cell("E7").value(data.color.toUpperCase())
        thisSheet.cell("E5").value("Fecha: " + today.getDate() + "/" + `${today.getMonth()+1 <10 ? "0"+(today.getMonth()+1) : (today.getMonth()+1)}` + "/" + today.getFullYear())
        thisSheet.cell("D45").value(data.descuento*100)
        thisSheet.cell("C47").value(`Portes ${totalKgs} kgs brutos`)
        thisSheet.cell("D47").value(totalKgs)
        thisSheet.cell("E47").value(PRECIO_PORTES_POR_KG)
        thisSheet.cell("E11").value(data.telefono)

        
        //Empezar a escribir conceptos

        let currentRow = PRIMERA_FILA_CONCEPTOS

        const makeTitle = (titulo)=>{
          thisSheet.cell(`B${currentRow}`).value(titulo).style("fill", BLUE).style("fontColor", WHITE)
          currentRow++
        }

        const makeCapasConceptRows = (amount, kitSize, price) =>{

          const capasObject = {
            'epoxy brillo':{
              1:"EPOXI",
              2:"Catalizador 5 a 1",
              3:"COLOR 100% Sólidos (MANOS)"
            },
            "epoxy mate":{
              1:"EPOXY MATE",
              2:"Catalizador 10 a 1",
              3:"COLOR (MANOS)"
            },
            'acrilica':{
              1:"ENEKRIL",
              2:"",
              3:"COLOR (MANOS)"
            },
            'politop':{
              1:"POLITOP",
              2:"",
              3:"COLOR (MANOS)"
            },
            'epoxy am':{
              1:"EPOXY AM",
              2:"Catalizador 5 a 1",
              3:"COLOR (MANOS)"
            }
          }
          if(amount){
            
            thisSheet.cell(`B${currentRow}`).value((data.resina == "epoxy brillo" || data.resina =="epoxy mate" ? "KIT ": "")+kitSize+" KGS")
            thisSheet.cell(`C${currentRow}`).value(capasObject[data.resina][1])
            thisSheet.cell(`D${currentRow}`).value(parseFloat(amount))
            thisSheet.cell(`E${currentRow}`).value(parseFloat(price))
            currentRow++

            thisSheet.cell(`B${currentRow}`).value(capasObject[data.resina][2])
            thisSheet.cell(`C${currentRow}`).value(capasObject[data.resina][3].replace("COLOR", data.color.toUpperCase()).replace("MANOS", data.capas>1 ? "DOS MANOS" : "UNA MANO"))
            currentRow++
          }
        }

        const makeImprimacionConceptRows = (amount, kitSize, price) =>{

          const imprimacionObject = {
            'epoxy brillo':{
              1:"EPOXI PRIMER",
              2:"Catalizador 5 a 1",
              3:"100% SOLIDOS"
            },
            "epoxy mate":{
              1:"EPOXI PRIMER",
              2:"Catalizador 5 a 1",
              3:"100% SOLIDOS"
            },
            'acrilica':{
              1:"ENEKRIL PRIMER",
              2:"",
              3:""
            },
            'politop':{
              1:"ERROR",
              2:"ERROR",
              3:"ERROR"
            },
            'epoxy am':{
              1:"EPOXI PRIMER",
              2:"Catalizador 5 a 1",
              3:"100% SOLIDOS"
          }
        }
          if(amount){
            
            thisSheet.cell(`B${currentRow}`).value((data.resina == "epoxy brillo" || data.resina =="epoxy mate" ? "KIT ": "")+kitSize+" KGS")
            thisSheet.cell(`C${currentRow}`).value(imprimacionObject[data.resina][1])
            thisSheet.cell(`D${currentRow}`).value(parseFloat(amount))
            thisSheet.cell(`E${currentRow}`).value(parseFloat(price))
            currentRow++

            thisSheet.cell(`B${currentRow}`).value(imprimacionObject[data.resina][2])
            thisSheet.cell(`C${currentRow}`).value(imprimacionObject[data.resina][3])
            currentRow++
          }
        }

        const makeDisolventeConceptRow = (kitSize, price) =>{

          const disolventeObject = {
            'epoxy brillo': "ENESOL EPOXY",
            "epoxy mate": "ENESOL EPOXY",
            'acrilica': "AGUA",
            'politop': "ENESOL POLITOP",
            'epoxy am': "ENESOL EPOXY"
          }
          if(kitSize){
            
            thisSheet.cell(`B${currentRow}`).value(1+" LTS")
            thisSheet.cell(`C${currentRow}`).value(disolventeObject[data.resina])
            thisSheet.cell(`D${currentRow}`).value(kitSize)
            thisSheet.cell(`E${currentRow}`).value(parseFloat(price))
            currentRow++
          }
        }
        
        const makeHerramientasConceptRow = (name, amount) => {
          if(amount){
            thisSheet.cell(`B${currentRow}`).value("UNIDADES")
            thisSheet.cell(`C${currentRow}`).value(name)
            thisSheet.cell(`D${currentRow}`).value(amount)
            thisSheet.cell(`E${currentRow}`).value(0)
            currentRow++
          }
        }

        const makeExtrasConceptRow = (name, amount, price) => {
          thisSheet.cell(`B${currentRow}`).value("UNIDADES")
          thisSheet.cell(`C${currentRow}`).value(name)
          thisSheet.cell(`D${currentRow}`).value(parseFloat(amount))
          thisSheet.cell(`E${currentRow}`).value(parseFloat(price))
          currentRow++
        }
        if(data.imprimacion){
          //titulo
            makeTitle("IMPRIMACIÓN")
          //filas
          for(let i=0; i<5; i++){
            makeImprimacionConceptRows(data.resultData.imprimacion.amountOfKits[i], data.resultData.imprimacion.sizeOfKits[i], data.priceObject.imprimacion[data.resultData.imprimacion.sizeOfKits[i]])
          }

          //hueco de harina de carzo
          if(hasHarina){
            thisSheet.cell(`B${currentRow}`).value("KGS 25")
            thisSheet.cell(`C${currentRow}`).value("HARINA DE CUARZO")
            thisSheet.cell(`D${currentRow}`).value(parseFloat(hasHarina.amount))
            thisSheet.cell(`E${currentRow}`).value(parseFloat(hasHarina.value))
            currentRow++
          }
          
        }
        if(data.capas){
          //titulo
            makeTitle(`${data.capas} MANO${data.capas>1 ? "S": ""} ${texteator[data.resina].name}`)
          //capas
          for(let i=0; i<5; i++){
            makeCapasConceptRows(data.resultData.capas.amountOfKits[i], data.resultData.capas.sizeOfKits[i], data.priceObject.capas[data.resultData.capas.sizeOfKits[i]])
          }
        }
        if(data.resultData.disolvente){
          //titulo
            makeTitle("DISOLVENTE")
          //capas
            makeDisolventeConceptRow(data.resultData.disolvente, data.disolventePrice)
        }
        if(data.herramientas.rodillos || data.herramientas.basculas || data.herramientas.cubos){
          //titulo
          makeTitle("HERRAMIENTAS")
          if(data.herramientas.basculas){
            makeHerramientasConceptRow("BALANZA DE 2G A 5 KGS", data.herramientas.basculas)
          }
          if(data.herramientas.cubos){
            makeHerramientasConceptRow("CUBOS DE MEZCLA", data.herramientas.cubos)
          }
          if(data.herramientas.rodillos){
            makeHerramientasConceptRow("RODILLOS", data.herramientas.rodillos)
          }
        }
          //extras
        if(subExtras[0]){
          makeTitle("EXTRAS")
          for(let i=0; i<subExtras.length; i++){
            makeExtrasConceptRow(subExtras[i].name, subExtras[i].amount, subExtras[i].value)
          }
        }

        return workbook.toFileAsync(customPath ? customPath + `${userName} ${thisFileName}.xlsx` : `presupuestos/${userName} ${thisFileName}.xlsx`)
    })
    .then(()=>{
      console.log("trying to open", customPath ? customPath + `${userName} ${thisFileName}.xlsx` : path.join(__dirname,`/presupuestos/${userName} ${thisFileName}.xlsx`))
      openMyFile( customPath ? customPath+ `${userName} ${thisFileName}.xlsx` : path.join(__dirname,`/presupuestos/${userName} ${thisFileName}.xlsx`), {wait:true})
    })
    

  res.send("it might just work")
})





app.post('/listadotelefonos', (req, res)=>{

  const auth = req.body.token.auth
  const presupuestos = req.body.presupuestos
  const pedidos = req.body.pedidos
  const clientes = req.body.clientes

  let today = new Date()
  let dd = today.getDate();
  let mm = today.getMonth()+1; 
  let yyyy = today.getFullYear();
  if(dd<10) {
      dd='0'+dd;
  } 
  if(mm<10) {
      mm='0'+mm;
  } 
  today = yyyy+'-'+mm+'-'+dd;


let telefonosQueHanComprado = []
for(let j=0; j<pedidos.length; j++){
  if(pedidos[j][6].dato){
    telefonosQueHanComprado.push(clientes[pedidos[j][6].dato])
  }
}
  
  XlsxPopulate.fromFileAsync("./baselistado.xlsx")
  .then(workbook => {
      // Modify the workbook.      
      let thisSheet = workbook.sheet("Hoja1")
      let currentRow = 4
      let telefonosQueNoHanComprado = []

      for(let i=0; i<presupuestos.length;i++){
        if(presupuestos[i][79].dato){
          if(!telefonosQueHanComprado.includes(presupuestos[i][79].dato) && !telefonosQueNoHanComprado.includes(presupuestos[i][79].dato)){
            telefonosQueNoHanComprado.push(presupuestos[i][79].dato)

            //escribir en excel
            thisSheet.cell(`B${currentRow}`).value(presupuestos[i][2].dato)
            thisSheet.cell(`C${currentRow}`).value(presupuestos[i][79].dato)
            thisSheet.cell(`D${currentRow}`).value(presupuestos[i][61].dato)
            thisSheet.cell(`E${currentRow}`).value(presupuestos[i][3].dato.split("T")[0])
            currentRow++
          }
        }
      }

      console.warn(telefonosQueNoHanComprado)
      return workbook.toFileAsync(listadoPath ? listadoPath + `${userName} ${today}.xlsx` : `listados/${userName} ${today}.xlsx`)
  })    
  .then(()=>{
    openMyFile( listadoPath ? listadoPath+ `${userName} ${today}.xlsx` : path.join(__dirname,`/listados/${userName} ${today}.xlsx`), {wait:true})
  })

  
})


app.use((req, res, next) => {
  res.sendFile(path.join(__dirname, "..", "build", "index.html"));
});


// start express server on port 5000
app.listen(5000, () => {
  console.log("server started on port 5000");
});



        //SETTING COLOR -> cell.style("fill", "0000ff");
        //SETTING BOLD -> cell("X1").style('bold');

        // add two rich text fragments -> cell.value().add('hello ', { italic: true, bold: true }).add('world!', { fontColor: 'FF0000' });
        /*
            const richtext = new RichText();
            richtext.add('hello');
            cell.value(richtext);
            cell.value().text(); // returns 'hello'

            richtext.add(' world')
            richtext.text(); // returns 'hello world' 
            cell.value().text(); // returns 'hello'
            cell.value() === richtext; // returns false

            cell.value().add(' world');
            cell.value().text(); // returns 'hello world'
        */
