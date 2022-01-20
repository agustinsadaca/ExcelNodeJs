var excel = require('excel4node');

const writeExcel = (jsonExcel,filename) =>{
// Create a new instance of a Workbook class
var workbook = new excel.Workbook();
// Add Worksheets to the workbook
var worksheet = workbook.addWorksheet('COMPROBANTES');
// var worksheet1 = workbook.addWorksheet('ASOCIADOS');

// Create a reusable style
var style = workbook.createStyle({
  font: {
    color: '#000000',
    size: 12
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -'
});
var style2 = workbook.createStyle({
  font: {
    name:'Arial',
    color: '#000000',
    size: 10
  },
  fill:{
      type:'pattern',
      bgColor:'#C2D6EB'
  }

});
const dateFormaterDDMMYYYYxDDMMYY= (date) =>{
    let dateArray = date.split("/")
    let year = dateArray[2].substring(2)
    return dateArray[0]+'/'+dateArray[1]+'/'+year
}
const dateFormaterDDMMYYYYxMMDDYYYY = (date) =>{
    let dateArray = date.split("/")
    return dateArray[1]+'/'+dateArray[0]+'/'+dateArray[2]
}
const dateFormaterMMDDYYYYxDDMMYY= (date) =>{
    let dateArray = date.split("/")
    let year = dateArray[2].substring(2)
    return dateArray[1]+'/'+dateArray[0]+'/'+year
}
const calculateModifiedDate = (date) =>{
    const todayDate = Date.now()
    const today = new Date(todayDate);
    const currentMonth = today.getMonth() + 1
    const currentMonthWithZeros = ('0' + currentMonth).slice(-2)
    const currentDay = today.getDate()
    const currentYear = today.getFullYear()	
    const currentDateFormated = currentMonthWithZeros+'/'+currentDay +'/'+currentYear


    const excelDate = new Date(date);
    const currentDate = new Date(currentDateFormated);
    const diffTime = Math.abs(currentDate - excelDate);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
    if (diffDays > 5) {
        return currentDateFormated
    }else{return null}

}
// Set value of cell A1 to 100 as a number type styled with paramaters of style
let headers = ['RAZONSOCIAL','NRODOCUMENTO','TIPODOCUMENTO','DIRECCIONFISCAL','PROVINCIA','LOCALIDAD','CODIGOPOSTAL','PERCIBEIVA','PERCIBEIIBB','TRATAMIENTOIMPOSITIVO','ENVIARCOMPROBANTE','MAILFACTURACION','CONTACTO','TELEFONO','CONVENIOMULTILATERAL','ENVIARBOTONPAGO','NUMEROCOMPROBANTE','FECHAHORA','TIPOCOMPROBANTE','OBSERVACIONES','DESCUENTO','PERCEPCIONIVA','PERCEPCIONIIBB','ORDENCOMPRA','REMITO','IMPORTEIMPUESTOSINTERNOS','IMPORTEPERCEPCIONESMUNIC','MONEDA','TIPODECAMBIO','CONDICIONVENTA','PRODUCTOSERVICIO','FECHAVENCIMIENTO','FECHAINICIOABONO','FECHAFINABONO','CANTIDAD','DETALLE','CODIGO','BONIFICACION','IVA','PRECIOUNITARIO','CBU','ESANULACION','TRANSFERENCIA']
for (let index = 0; index < headers.length; index++) {
    worksheet.cell(1,index+1).string(headers[index]).style(style2);
}
// let headers2 = ['TIPOCOMPROBANTE','NUMEROCOMPROBANTE','TIPOCOMPROBANTEASOCIADO','NUMEROCOMPROBANTEASOCIADO','FECHACOMPROBANTEASOCIADO']
// for (let index = 0; index < headers2.length; index++) {
//     worksheet1.cell(1,index+1).string(headers2[index]).style(style);
// }
for (let index = 0; index < jsonExcel.length; index++) {
    worksheet.cell(index+2,1).string('Cliente generico de Prisma').style(style);
    worksheet.cell(index+2,2).string('1').style(style);
    worksheet.cell(index+2,3).string('DNI').style(style);
    worksheet.cell(index+2,4).string('NO ESPECIFICADA').style(style);
    worksheet.cell(index+2,5).string('Mendoza').style(style);
    worksheet.cell(index+2,6).string('Bowen').style(style);
    worksheet.cell(index+2,7).string('5634').style(style);
    worksheet.cell(index+2,10).string('CONSUMIDOR FINAL').style(style);
    worksheet.cell(index+2,11).string('N').style(style);
    worksheet.cell(index+2,15).string('N').style(style);
    worksheet.cell(index+2,16).string('N').style(style);
    worksheet.cell(index+2,17).string(`0004-${jsonExcel[index]['NUM.CUPON']}`).style(style);
    worksheet.cell(index+2,19).string('FB').style(style);
    worksheet.cell(index+2,28).string('2').style(style);
    worksheet.cell(index+2,29).string('1').style(style);
    worksheet.cell(index+2,30).string('1').style(style);
    worksheet.cell(index+2,31).string('1').style(style);
    worksheet.cell(index+2,35).string('1').style(style);
    worksheet.cell(index+2,36).string('Almacen').style(style);
    worksheet.cell(index+2,39).string('21').style(style);
    worksheet.cell(index+2,40).string(jsonExcel[index]['MONTO_BRUTO'].replace(".", ",")).style(style);
    
    let excelDate = dateFormaterDDMMYYYYxMMDDYYYY(jsonExcel[index]['PRESENTACION'])
    let modifiedDate = calculateModifiedDate(excelDate)
    let finalDate = dateFormaterMMDDYYYYxDDMMYY(modifiedDate)
  
    worksheet.cell(index+2,18).string(finalDate).style(style);
}
direccionExcel = `download/Comprobantes${filename.replace('.csv','')}.xlsx`
workbook.write(direccionExcel);
// console.log(direccionExcel);
return direccionExcel
}
module.exports = {writeExcel}