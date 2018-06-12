const EXPRESS = require('express');
const APP = new EXPRESS();


APP.get('/',(req,res)=>{
  res.sendFile(__dirname+'/xlsx_2Json_SheetName_data.json')
})

APP.listen(3000,()=>{
  console.log("OK");
})
