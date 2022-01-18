const csv = require('csv-parser');
const fs = require("fs")
const {writeExcel} = require("./writeExcel");


const readExcel = async(filename) => {
  let jsonExcel = []
  let read = await fs.createReadStream(`uploads/${filename}`)
  .pipe(csv({ skipLines:1}))
  .on('data', (row) => {
    jsonExcel.push(row)
  })
  .on('end', () => {
    console.log('CSV file successfully processed');
    writeExcel(jsonExcel,filename)
  });
};
module.exports = {readExcel}
