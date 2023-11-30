const csv = require("csv-parser");
const fs = require("fs");
const { writeExcel,writeExcelBancoNacion } = require("./writeExcel");
const url = require("url");

const readExcel = async (req, res, next) => {
  let jsonExcel = [];
  let filename = req.file.filename;
  let csvConfig 
  console.log(filename);
  if (!filename.toLowerCase().endsWith('.csv')) {
    return res.redirect(
      url.format({
        pathname: "/",
        query: {
          error: "No es posible convertir el excel, el tipo de archivo no es csv",
        },
      })
    );
  }
  switch (req.url) {
    case "/excel":
      csvConfig = { skipLines: 1, separator: "," } 
      break    
    case "/excelBancoNacion":
      csvConfig = { skipLines: 3, separator: ";"}     
      break;
    default:
      break;
  }
  let read = await fs
    .createReadStream(`uploads/${filename}`)
    .pipe(csv(csvConfig))
    .on("data", (row) => {
      jsonExcel.push(row);
    })
    .on("end", () => {
      console.log("CSV file successfully processed");
      console.log(jsonExcel);
      let direccionExcel 
      switch (req.url) {
        case "/excel":
          direccionExcel = writeExcel(jsonExcel, filename);
          break;
        case "/excelBancoNacion":
          direccionExcel = writeExcelBancoNacion(jsonExcel, filename);
          break;
        default:
          break;
      }
      res.redirect(
        url.format({
          pathname: "/",
          query: {
            direccionExcel: direccionExcel,
          },
        })
      );
    });
};
module.exports = { readExcel };
