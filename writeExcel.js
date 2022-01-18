var excel = require('excel4node');

const writeExcel = (jsonExcel,filename) =>{
// Create a new instance of a Workbook class
var workbook = new excel.Workbook();
// console.log(jsonExcel[1100]['NUM.CUPON']);
// Add Worksheets to the workbook
var worksheet = workbook.addWorksheet('Sheet 1');
var worksheet2 = workbook.addWorksheet('Sheet 2');

// Create a reusable style
var style = workbook.createStyle({
  font: {
    color: '#FF0800',
    size: 12
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -'
});
const dateFormater = (date) =>{
    let dateArray = date.split("/")
    let year = dateArray[2].substring(2)
    return dateArray[1]+dateArray[0]+year
}
// Set value of cell A1 to 100 as a number type styled with paramaters of style
let headers = ['RAZONSOCIAL','NRODOCUMENTO','TIPODOCUMENTO','DIRECCIONFISCAL','PROVINCIA','LOCALIDAD','CODIGOPOSTAL','PERCIBEIVA','PERCIBEIIBB','TRATAMIENTOIMPOSITIVO','ENVIARCOMPROBANTE','MAILFACTURACION','CONTACTO','TELEFONO','CONVENIOMULTILATERAL','ENVIARBOTONPAGO','NUMEROCOMPROBANTE','FECHAHORA','TIPOCOMPROBANTE','OBSERVACIONES','DESCUENTO','PERCEPCIONIVA','PERCEPCIONIIBB','ORDENCOMPRA','REMITO','IMPORTEIMPUESTOSINTERNOS','IMPORTEPERCEPCIONESMUNIC','MONEDA','TIPODECAMBIO','CONDICIONVENTA','PRODUCTOSERVICIO','FECHAVENCIMIENTO','FECHAINICIOABONO','FECHAFINABONO','CANTIDAD','DETALLE','CODIGO','BONIFICACION','IVA','PRECIOUNITARIO','CBU','ESANULACION','TRANSFERENCIA']
for (let index = 0; index < headers.length; index++) {
    worksheet.cell(1,index+1).string(headers[index]).style(style);
}
for (let index = 0; index < jsonExcel.length; index++) {
    worksheet.cell(index+2,1).string('Sadaca Alberto Ismael').style(style);
    worksheet.cell(index+2,2).number(23213764519).style(style);
    worksheet.cell(index+2,3).string('CUIT').style(style);
    worksheet.cell(index+2,4).string('Rivadavia 798').style(style);
    worksheet.cell(index+2,5).string('Mendoza').style(style);
    worksheet.cell(index+2,6).string('Bowen').style(style);
    worksheet.cell(index+2,7).number(5634).style(style);
    worksheet.cell(index+2,10).string('RESPONSABLE INSCRIPTO').style(style);
    worksheet.cell(index+2,11).string('N').style(style);
    worksheet.cell(index+2,15).string('N').style(style);
    worksheet.cell(index+2,16).string('N').style(style);
    worksheet.cell(index+2,17).string(`0004-${jsonExcel[index]['NUM.CUPON']}`).style(style);
    worksheet.cell(index+2,18).string(`${dateFormater(jsonExcel[index]['PRESENTACION'])}`).style(style);
    worksheet.cell(index+2,19).string('FB').style(style);
    worksheet.cell(index+2,28).number(2).style(style);
    worksheet.cell(index+2,29).number(1).style(style);
    worksheet.cell(index+2,30).number(1).style(style);
    worksheet.cell(index+2,31).number(1).style(style);
    worksheet.cell(index+2,35).number(1).style(style);
    worksheet.cell(index+2,36).string('Almacen').style(style);
    worksheet.cell(index+2,39).number(21).style(style);
    worksheet.cell(index+2,40).string(jsonExcel[index]['MONTO_BRUTO']).style(style);
    worksheet.cell(index+2,42).string('FALSO').style(style);
    
    
    
}

workbook.write(`download/${filename}`);
}
module.exports = {writeExcel}