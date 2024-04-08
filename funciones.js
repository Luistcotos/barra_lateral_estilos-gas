// Variable donde guardaremos todos los estilos

function onOpen() {
  
  SpreadsheetApp.getUi().createMenu('Barra Lateral')
  .addItem('Mostrarbarralateral','mostrarBarraLateral')
  .addToUi();
}


function mostrarBarraLateral()
{
  var barra = HtmlService.createHtmlOutputFromFile('BarraLateral').setTitle('Barra lateral Aulaenlanube');
  SpreadsheetApp.getUi().showSidebar(barra)
}

function aplicarEstilo1()
{
  var hoja_actual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdas = hoja_actual.getActiveRange();

  celdas.setBackground('blue')
        .setFontColor('white')
        .setHorizontalAlignment('center')
        .setValue('Estilo 1');
}

function aplicarEstilo2()
{
  var hoja_actual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdas = hoja_actual.getActiveRange();

  celdas.setBackground('green')
        .setFontColor('white')
        .setFontWeight('bold')
        .setValue('Estilo 2');
}

function borrarestilos()
{
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear({formatOnly: true}); 
}

function borrartodo()
{
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear(); 
}