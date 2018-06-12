const fs=require('fs');
const util = require('util')
const xlsx = require('node-xlsx');
const momment = require('moment');

const workSheetsFromFile = xlsx.parse(`${__dirname}/evento.xls`);
const horario8 = xlsx.parse(`${__dirname}/8.xls`);

let horario1 = horario8[0].data.map(item=>{
  return item[0]
})

momment.locale('es');
let dataComplete=workSheetsFromFile[0].data.map((item,ind)=>{
    
    if(ind>0){
      return [[`${item[1]}`],[`${item[2]} ${item[3]}`],[item[8]],[`${momment(new Date((item[0] - (25567 + 1))*86400*1000).toUTCString()).format("D/M/YY h:mm:ss a")}`]];
    }
})

dataComplete.splice(0,1)

let test1=dataComplete.map((item,ind)=>{
  for(let i=0;i<=horario1.length;i++){
    if(item[0][0]==horario1[i]){            
      return item[1][0];
    }
  }
})

//console.log(test1);

let Filtro=Array.from(new Set(test1));
console.log(Filtro.sort());

//console.log(util.inspect(da, { maxArrayLength: null }))
// let buffer = xlsx.build([{name: "mySheetName", data: da.sort()}], option); // Returns a buffer
// for(let i=0;i<tam;i++){
// 
//   horario1.map(item=>{
//     console.log("hora",item);
//     console.log("da",da[i][0]);
//     console.log("===================================================");
//     if(da[i][0]==item){
//       console.log("==============================================================================================");
//       //console.log("ok");
//     }
// 
//   })
// 
// }

//console.log(util.inspect(da, { maxArrayLength: null }))
//console.log(util.inspect(n, { maxArrayLength: null }))
//console.log(util.inspect(nue, { maxArrayLength: null }))
//console.log(da[1]);
//console.log(util.inspect(da, { maxArrayLength: null }))
//console.log(util.inspect(da.sort(), { maxArrayLength: null }))
//const nueva = da.sort();
// const range = {s: {c: 0, r:0 }, e: {c:0, r:0}}; // A1:A4
// const option = {'!merges': [ range ]};
// 
// let buffer = xlsx.build([{name: "mySheetName", data: da.sort()}], option); // Returns a buffer
// 
//   fs.writeFile('message.xls', buffer, (err) => {
//     if (err) throw err;
//    console.log('The file has been saved!');
//  });
// 
// 

