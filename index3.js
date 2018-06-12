const fs=require('fs');
const util = require('util')
const xlsx = require('node-xlsx');
const momment = require('moment');
const excel = require('excel');
const { getJsDateFromExcel } = require('excel-date-to-js');

const original = xlsx.parse(`${__dirname}/evento.xls`);
const horario1 = xlsx.parse(`${__dirname}/8.xls`);

let FilesName=momment(new Date()).format('DD_MM_Y_h_m_s');


let o=[];
 original[0].data.map(item=>{
   horario1[0].data.map(item2=>{
     if(item[8]!="Entrada T.E." && item[8]!="Salida T.E."){
       if(item[1]==item2[0]){
         o.push(item)
       }
     }
   })
 })


let test=o.map(item=>{
  let fechaCompleta=momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).add(5 ,'hours').format('D/M/Y HH:mm:ss')
  let fecha=momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).add(5 ,'hours').format('D/M/Y');
  let hora=momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).add(5 ,'hours').format('HH:mm:ss');
  limite = momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).hour(8).minutes(16).seconds(0).format('D/M/Y HH:mm:ss');
  let Tiempo="";
  
  if(item[8]=="Entrada"){
    
    if(fechaCompleta>limite){
      Tiempo="Retardo"
    }else{
      Tiempo="A tiempo"
    }
    
  }else{
    Tiempo="Salida"
  }
  
  return [[item[1]],[`${item[2]} ${item[3]}`],[`${fechaCompleta}`],[`${item[1]}`],[fecha],[hora],[item[8]],[Tiempo]];
  
})

const range = {s: {c: 0, r:0 }, e: {c:0, r:0}}; // A1:A4
const option = {'!merges': [ range ]};

let buffer2 = xlsx.build([{name: "mySheetName", data: test.sort()}], option); // Returns a buffer

  fs.writeFile(`${FilesName}_Horario8am.xls`, buffer2,'utf8', (err) => {
    if (err) throw err;
   console.log('Archivo horario 8:00 am generado');
 });




 const horario2 = xlsx.parse(`${__dirname}/9.xls`);

 let o2=[];
  original[0].data.map(item=>{
    horario2[0].data.map(item2=>{
      if(item[8]!="Entrada T.E." && item[8]!="Salida T.E."){
        if(item[1]==item2[0]){
          o2.push(item)
        }
      }
    })
  })


 let test2=o2.map(item=>{
   let fechaCompleta=momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).add(5 ,'hours').format('D/M/Y HH:mn:ss')
   let fecha=momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).add(5 ,'hours').format('D/M/Y');
   let hora=momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).add(5 ,'hours').format('HH:mm:ss');
   limite = momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).hour(9).minutes(16).seconds(0).format('D/M/Y HH:mm:ss');
   let Tiempo2="";
   
   if(item[8]=="Entrada"){
     
     if(fechaCompleta>limite){
       Tiempo2="Retardo"
     }else{
       Tiempo2="A tiempo"
     }
     
   }else{
     Tiempo2="Salida"
   }
   
   return [[item[1]],[`${item[2]} ${item[3]}`],[`${fechaCompleta}`],[`${item[1]}`],[fecha],[hora],[item[8]],[Tiempo2]];
   
 })

 const range2 = {s: {c: 0, r:0 }, e: {c:0, r:0}}; // A1:A4
 const option2 = {'!merges': [ range2 ]};

 let buffer3 = xlsx.build([{name: "mySheetName", data: test2.sort()}], option2); // Returns a buffer

   fs.writeFile(`${FilesName}_Horario9am.xls`, buffer3,'utf8', (err) => {
     if (err) throw err;
    console.log('Archivo horario 9:00 am generado');
  });


  const horario3 = xlsx.parse(`${__dirname}/10.xls`);

  let o3=[];
   original[0].data.map(item=>{
     horario3[0].data.map(item2=>{
       if(item[8]!="Entrada T.E." && item[8]!="Salida T.E."){
         if(item[1]==item2[0]){
           o3.push(item)
         }
       }
     })
   })


  let test3=o3.map(item=>{
    let fechaCompleta=momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).add(5 ,'hours').format('D/M/Y HH:mm:ss')
    let fecha=momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).add(5 ,'hours').format('D/M/Y');
    let hora=momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).add(5 ,'hours').format('HH:mm:ss');
    limite = momment(new Date((item[0] - (25567 + 2)) * 86400 * 1000)).hour(10).minutes(16).seconds(0).format('D/M/Y HH:mm:ss');
    let Tiempo3="";
    
    if(item[8]=="Entrada"){
      
      if(fechaCompleta>limite){
        Tiempo3="Retardo"
      }else{
        Tiempo3="A tiempo"
      }
      
    }else{
      Tiempo3="Salida"
    }
    
    return [[item[1]],[`${item[2]} ${item[3]}`],[`${fechaCompleta}`],[`${item[1]}`],[fecha],[hora],[item[8]],[Tiempo3]];
    
  })

  const range3 = {s: {c: 0, r:0 }, e: {c:0, r:0}}; // A1:A4
  const option3 = {'!merges': [ range3 ]};

  let buffer4 = xlsx.build([{name: "mySheetName", data: test3.sort()}], option3); // Returns a buffer

    fs.writeFile(`${FilesName}_Horario10am.xls`, buffer4,'utf8', (err) => {
      if (err) throw err;
     console.log('Archivo horario 10:00 am generado');
   });
