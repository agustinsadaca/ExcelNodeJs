const csv = require("csv-parser");
const fs = require("fs");
const { writeExcel,writeExcelBancoNacion } = require("./writeExcel");
const url = require("url");

const readExcel = async (req, res, next) => {
  let jsonExcel = [];
  let filename = req.file.filename;
  let csvConfig 
  switch (req.url) {
    case "/excel":
      csvConfig = { skipLines: 1 }     
    case "/excelBancoNacion":
      csvConfig = { skipLines: 4, separator: ";"}     
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
      console.log(req.url);
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
