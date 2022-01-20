const csv = require('csv-parser');
const fs = require("fs")
const {writeExcel} = require("./writeExcel");
const url = require('url');    



const readExcel = async(req, res, next) => {
  let jsonExcel = []
  let filename = req.file.filename
  let read = await fs.createReadStream(`uploads/${filename}`)
  .pipe(csv({ skipLines:1}))
  .on('data', (row) => {
    jsonExcel.push(row)
  })
  .on('end', () => {
    console.log('CSV file successfully processed');
    let direccionExcel = writeExcel(jsonExcel,filename)
    res.redirect(url.format({
      pathname:"/",
      query: {
         "direccionExcel":direccionExcel
       }
    }))
  });
  
};
module.exports = {readExcel}
