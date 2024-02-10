/** @OnlyCurrentDoc */

function ClientesOrdenarAZ() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Clientes'), true);
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort({column: 1, ascending: true});
  spreadsheet.getRange('A1').activate();
  SpreadsheetApp.flush()
};

function GuardarCliente() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Clientes'), true);
  spreadsheet.getRange('A1000').activate();
  spreadsheet.getRange('Inicio!B4:F4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('E1000').activate();
  ClientesOrdenarAZ();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Inicio'), true);
  spreadsheet.getRange('B4').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP($C$1;Clientes!$A$2:$F$1000;1;)');
  spreadsheet.getRange('C4').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP($C$1;Clientes!$A$2:$F$1000;2;)');
  spreadsheet.getRange('D4').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP($C$1;Clientes!$A$2:$F$1000;3;)');
  spreadsheet.getRange('E4').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP($C$1;Clientes!$A$2:$F$1000;4;)');
  spreadsheet.getRange('F4').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP($C$1;Clientes!$A$2:$F$1000;5;)');
  spreadsheet.getRange('G4').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('H4').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP($C$1;Clientes!$A$2:$F$1000;6;)');
  spreadsheet.getRange('C1').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  SpreadsheetApp.flush()
};

function AñadirSemana() {
  var spreadsheet = SpreadsheetApp.getActive();
  const ahora = new Date;
  const finde = new Date(ahora.getTime() + 558400000);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Pedidos'), true);
  spreadsheet.setCurrentCell(spreadsheet.getRange('A1000'));
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(2, 0).activate();
  spreadsheet.getCurrentCell().setValue("Semana del");
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(ahora).setNumberFormat("dd/mm");
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue("al");
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue(finde).setNumberFormat("dd/mm");
  spreadsheet.getCurrentCell().offset(0, -3, 1, 6).activate();
  spreadsheet.getActiveRangeList().setBackground('#d9d2e9')
  .setFontWeight('bold')
  .setHorizontalAlignment('center');
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
  SpreadsheetApp.flush()
};


function LimpiarPedido() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Inicio'), true);
  spreadsheet.getRangeList(['G7', 'F9:H9', 'E7:E14', 'E16:E33', 'G4', 'E1', 'C1']).activate()
  .clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('H7').activate();
  spreadsheet.getCurrentCell().setValue('Falta')
  spreadsheet.getRange('E15').activate();
  spreadsheet.getCurrentCell().setValue('Nos da');
  spreadsheet.getRange('E7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E8').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E9').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E10').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E11').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E12').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E13').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E14').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E16').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E17').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E18').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E19').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E20').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E21').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E22').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E23').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E24').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E25').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E26').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E27').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E28').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E29').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E30').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E30').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E31').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E32').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E33').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('C1').activate();
  SpreadsheetApp.flush()
};

function CargarPedidoChipá() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1000').activate();
  spreadsheet.getRange('Inicio!B4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('Inicio!E1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setFormula('=TRANSPOSE(Inicio!E16:E20)');
  spreadsheet.getCurrentCell().offset(0, 5).activate();
  spreadsheet.getRange('Inicio!H9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName('Chipá');
  var ultimafila = spreadsheet.getLastRow();
  var ultimacol = spreadsheet.getLastColumn();
  var ultimacolumna = ultimacol-1;
  var pedidocargado = spreadsheet.getRange(ultimafila, 1, ultimafila, ultimacolumna).getDisplayValues();
  var j = 0;
  for(var i=0;i<ultimacolumna;i++){
    if(pedidocargado[0][ultimacolumna-i] == '0' || pedidocargado[0][ultimacolumna-i] == 'Nos da'){
      j=j+1;
    }
  }
  if(j==ultimacolumna-3){
    spreadsheet.deleteRow(ultimafila);
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Chipá'), true);
    spreadsheet.getRange('A999').activate();
    spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);
    spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
    }
  else{      
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    spreadsheet.getRange('A1000:I1000').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRangeList(['A1000:I1000']).activate()
    .clear({contentsOnly: true, skipFilteredRows: true})
    }
  spreadsheet.getRange('A1').activate();
  SpreadsheetApp.flush()
};

function CargarPedidoCerveza() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1000').activate();
  spreadsheet.getRange('Inicio!B4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('Inicio!E1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setFormula('=TRANSPOSE(Inicio!E7:E15)');
  spreadsheet.getCurrentCell().offset(0, 9).activate();
  spreadsheet.getRange('Inicio!H9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName('Cerveza');
  var ultimafila = spreadsheet.getLastRow();
  var ultimacol = spreadsheet.getLastColumn();
  var ultimacolumna = ultimacol-1;
  var pedidocargado = spreadsheet.getRange(ultimafila, 1, ultimafila, ultimacolumna).getDisplayValues();
  var j = 0;
  for(var i=0;i<ultimacolumna;i++){
    if(pedidocargado[0][ultimacolumna-i] == '0' || pedidocargado[0][ultimacolumna-i] == 'Nos da'){
      j=j+1;
    }
  }
  if(j==ultimacolumna-3){
    spreadsheet.deleteRow(ultimafila);
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cerveza'), true);
    spreadsheet.getRange('A999').activate();
    spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);
    spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
    }
  else{      
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    spreadsheet.getRange('A1000:L1000').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRangeList(['A1000:L1000']).activate()
    .clear({contentsOnly: true, skipFilteredRows: true})
    }
  spreadsheet.getRange('A1').activate();
  SpreadsheetApp.flush()
};

function CargarPedidoDips() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1000').activate();
  spreadsheet.getRange('Inicio!B4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('Inicio!E1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setFormula('=TRANSPOSE(Inicio!E21:E24)');
  spreadsheet.getCurrentCell().offset(0, 4).activate();
  spreadsheet.getRange('Inicio!H9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName('Dips');
  var ultimafila = spreadsheet.getLastRow();
  var ultimacol = spreadsheet.getLastColumn();
  var ultimacolumna = ultimacol-1;
  var pedidocargado = spreadsheet.getRange(ultimafila, 1, ultimafila, ultimacolumna).getDisplayValues();
  var j = 0;
  for(var i=0;i<ultimacolumna;i++){
    if(pedidocargado[0][ultimacolumna-i] == '0' || pedidocargado[0][ultimacolumna-i] == 'Nos da'){
      j=j+1;
    }
  }
  if(j==ultimacolumna-3){
    spreadsheet.deleteRow(ultimafila);
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Dips'), true);
    spreadsheet.getRange('A999').activate();
    spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);
    spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
    }
  else{      
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    spreadsheet.getRange('A1000:I1000').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRangeList(['A1000:I1000']).activate()
    .clear({contentsOnly: true, skipFilteredRows: true})
    }
  spreadsheet.getRange('A1').activate();
  SpreadsheetApp.flush()
};

function CargarPedidoPostres() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1000').activate();
  spreadsheet.getRange('Inicio!B4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('Inicio!E1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setFormula('=TRANSPOSE(Inicio!E25:E32)');
  spreadsheet.getCurrentCell().offset(0, 8).activate();
  spreadsheet.getRange('Inicio!H9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName('Postres');
  var ultimafila = spreadsheet.getLastRow();
  var ultimacol = spreadsheet.getLastColumn();
  var ultimacolumna = ultimacol-1;
  var pedidocargado = spreadsheet.getRange(ultimafila, 1, ultimafila, ultimacolumna).getDisplayValues();
  var j = 0;
  for(var i=0;i<ultimacolumna;i++){
    if(pedidocargado[0][ultimacolumna-i] == '0' || pedidocargado[0][ultimacolumna-i] == 'Nos da'){
      j=j+1;
    }
  }
  if(j==ultimacolumna-3){
    spreadsheet.deleteRow(ultimafila);
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Postres'), true);
    spreadsheet.getRange('A999').activate();
    spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);
    spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
    }
  else{      
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
    spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    spreadsheet.getRange('A1000:L1000').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRangeList(['A1000:L1000']).activate()
    .clear({contentsOnly: true, skipFilteredRows: true})
    }
  spreadsheet.getRange('A1').activate();
  SpreadsheetApp.flush()
};

function OrdenarPedido() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Pedidos'), true);
  var v = spreadsheet.getRange('D:D').getValues();
  var l = 0;
  for(var i=v.length-1;i>=0;i--){
    if(v[0,i]=='0'||v[0,i]=='Nos da'){
      spreadsheet.deleteRow(i+1);
      l=l+1;
    }
  }
  spreadsheet.getRange('A970').activate();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), l);
  spreadsheet.getRange('A1000').activate();
  SpreadsheetApp.flush()
};

function CargarPedidoSeguro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cerveza'), true);
  CargarPedidoCerveza();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Chipá'), true);
  CargarPedidoChipá();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Dips'), true);
  CargarPedidoDips();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Postres'), true);
  CargarPedidoPostres();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Pedidos'), true);
  spreadsheet.getRange('A1000').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(2, 0).activate();
  spreadsheet.getCurrentCell().setValue('Cliente');
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.getRange('Inicio!B4:D4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getCurrentCell().offset(-1, 0, 2, 3).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRangeList().setBackground('#efefef');
  spreadsheet.getCurrentCell().offset(-1, 0, 1, 3).activate()
  .mergeAcross();
  spreadsheet.getCurrentCell().offset(0, 3).activate();
  spreadsheet.getCurrentCell().setValue('Envío');
  var envios = spreadsheet.getRange('Inicio!G4').getDisplayValues();
  if(envios == 'Envío'){
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    spreadsheet.getRange('Inicio!E4:F4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getCurrentCell().offset(0, 2).activate();
    spreadsheet.getRange('Inicio!H3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    var currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getCurrentCell().offset(-1, -2, 2, 3).activate();
    spreadsheet.setCurrentCell(currentCell);
    spreadsheet.getActiveRangeList().setBackground('#efefef');
    currentCell = spreadsheet.getCurrentCell().offset(-1, 0);
    spreadsheet.getCurrentCell().offset(-1, -2, 1, 3).activate();
    spreadsheet.setCurrentCell(currentCell);
    spreadsheet.getActiveRange().mergeAcross();
  } else{
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    spreadsheet.getRange('Inicio!G4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getCurrentCell().offset(0, 2).activate();
    var currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getCurrentCell().offset(-1, -2, 2, 3).activate();
    spreadsheet.setCurrentCell(currentCell);
    spreadsheet.getActiveRangeList().setBackground('#efefef');
    currentCell = spreadsheet.getCurrentCell().offset(-1, 0);
    spreadsheet.getCurrentCell().offset(-1, -2, 1, 3).activate();
    spreadsheet.setCurrentCell(currentCell);
    spreadsheet.getActiveRange().mergeAcross();
  }  
  spreadsheet.getCurrentCell().offset(2, 0).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
  spreadsheet.getCurrentCell().setValue('Pedido');
  spreadsheet.getCurrentCell().offset(0, 4).activate();
  spreadsheet.getCurrentCell().setValue('Total');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('Finalizado');
  spreadsheet.getCurrentCell().offset(1, -5).activate();
  spreadsheet.getRange('Inicio!B7:E33').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(27, 0).activate();
  spreadsheet.getRange('Inicio!G6:G7').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('Inicio!G8:G9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('Inicio!H6:H7').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('Inicio!H8:H9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(1, 3).activate();
  spreadsheet.getRange('PrecioStock!E3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('PrecioStock!E4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('PrecioStock!E5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('PrecioStock!E6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('PrecioStock!E7').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('PrecioStock!E8').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('PrecioStock!E9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('PrecioStock!E10').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(-1, 0).activate();
  spreadsheet.getCurrentCell().setValue('Promo');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Kit');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Postres');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Dips');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Chipá');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Envases');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Cervezas');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Envío');
  OrdenarPedido();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(-1, 0).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.getRange('Inicio!F7').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getActiveRange().setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireCheckbox('Finalizado', 'Falta algo')
  .build());
  spreadsheet.getRangeList([spreadsheet.getCurrentCell().offset(-1, -1, 2, 2).getA1Notation(),
  spreadsheet.getCurrentCell().offset(-1, -5, 1, 4).getA1Notation()]).activate()
  .setBackground('#d9d2e9');
  spreadsheet.getCurrentCell().offset(0, 6).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getCurrentCell().offset(0, 0, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#efefef');
  spreadsheet.getCurrentCell().offset(0, 1, 2, 2).activate();
  spreadsheet.getActiveRangeList().setBackground('#d5a6bd');
  spreadsheet.getCurrentCell().offset(0, 2, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#fff2cc');
  spreadsheet.getCurrentCell().offset(0, 1, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#d9ead3');
  spreadsheet.getCurrentCell().offset(0, 1, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#ea9999');
  spreadsheet.getCurrentCell().offset(0, 1, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#efefef');
  spreadsheet.getCurrentCell().offset(0, 1, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#d9d2e9');
  spreadsheet.getRange('A1').activate();
  SpreadsheetApp.flush()
  LimpiarPedido();
  SpreadsheetApp.getUi().alert("Pedido cargado! 🎉","Si lees esto, sé que puedes leer mis pensamientos muchacho: miau miau",SpreadsheetApp.getUi().ButtonSet.OK);
};

function CargarPedido(){
  SpreadsheetApp.getUi()
      .createMenu('Custom Menu')
      .addItem('Show alert', 'showAlert')
      .addToUi();
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'Cargaste todo? Segure?',
     'Lo único que puede quedar en blanco es "Pagado a:" ',
      ui.ButtonSet.YES_NO);
  if (result == ui.Button.YES) {
    CargarPedidoSeguro();
  } else {
    ui.alert('Acomodalo porfi');
  }
};

function onEdit(e){
  //Pinta gris los pedidos
  var spreadsheet= SpreadsheetApp.getActive(); 
  if (e.value == 'Finalizado'){ 
    spreadsheet.getCurrentCell().offset(-3, -5, 1, 3).activate();
    var currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
    currentCell.activateAsCurrentCell();
    currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    currentCell.activateAsCurrentCell();
    spreadsheet.getActiveRangeList().setBackground('#efefef');
    spreadsheet.getCurrentCell().activate();
  }
};

function CargarPrepedidoCervezaSeguro(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H9').activate();
  spreadsheet.getCurrentCell().setValue('PREPEDIDO');
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cerveza'), true);
  CargarPedidoCerveza();
  LimpiarPedido();
  SpreadsheetApp.getUi().alert("Prepedido cargado! 🎉","Recordá borrarlo cuando cargues el pedido real",SpreadsheetApp.getUi().ButtonSet.OK);
};


function CargarPrepedidoCerveza(){
  SpreadsheetApp.getUi()
      .createMenu('Custom Menu')
      .addItem('Show alert', 'showAlert')
      .addToUi();
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'Cargaste todo? En anotaciones pusiste que es prepedido?',
     'Lo único que puede quedar en blanco es "Pagado a:" ',
      ui.ButtonSet.YES_NO);
  if (result == ui.Button.YES) {
    CargarPrepedidoCervezaSeguro();
  } else {
    ui.alert('Acomodalo porfi');
  }
};

function CargarKitSeguro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H9').activate();
  spreadsheet.getCurrentCell().setValue('KIT');
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cerveza'), true);
  CargarPedidoCerveza();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Chipá'), true);
  CargarPedidoChipá();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Dips'), true);
  CargarPedidoDips();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Postres'), true);
  CargarPedidoPostres();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Pedidos'), true);
  spreadsheet.getRange('A1000').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(2, 0).activate();
  spreadsheet.getCurrentCell().setValue('Cliente');
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.getRange('Inicio!B4:D4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getCurrentCell().offset(-1, 0, 2, 3).activate();
  spreadsheet.setCurrentCell(currentCell);
  spreadsheet.getActiveRangeList().setBackground('#efefef');
  spreadsheet.getCurrentCell().offset(-1, 0, 1, 3).activate()
  .mergeAcross();
  spreadsheet.getCurrentCell().offset(0, 3).activate();
  spreadsheet.getCurrentCell().setValue('Envío');
  var envios = spreadsheet.getRange('Inicio!G4').getDisplayValues();
  if(envios == 'Envío'){
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    spreadsheet.getRange('Inicio!E4:F4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getCurrentCell().offset(0, 2).activate();
    spreadsheet.getRange('Inicio!H3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    var currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getCurrentCell().offset(-1, -2, 2, 3).activate();
    spreadsheet.setCurrentCell(currentCell);
    spreadsheet.getActiveRangeList().setBackground('#efefef');
    currentCell = spreadsheet.getCurrentCell().offset(-1, 0);
    spreadsheet.getCurrentCell().offset(-1, -2, 1, 3).activate();
    spreadsheet.setCurrentCell(currentCell);
    spreadsheet.getActiveRange().mergeAcross();
  } else{
    spreadsheet.getCurrentCell().offset(1, 0).activate();
    spreadsheet.getRange('Inicio!G4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getCurrentCell().offset(0, 2).activate();
    var currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getCurrentCell().offset(-1, -2, 2, 3).activate();
    spreadsheet.setCurrentCell(currentCell);
    spreadsheet.getActiveRangeList().setBackground('#efefef');
    currentCell = spreadsheet.getCurrentCell().offset(-1, 0);
    spreadsheet.getCurrentCell().offset(-1, -2, 1, 3).activate();
    spreadsheet.setCurrentCell(currentCell);
    spreadsheet.getActiveRange().mergeAcross();
  }  
  spreadsheet.getCurrentCell().offset(2, 0).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
  spreadsheet.getCurrentCell().setValue('Pedido');
  spreadsheet.getCurrentCell().offset(0, 4).activate();
  spreadsheet.getCurrentCell().setValue('Total');
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().setValue('Finalizado');
  spreadsheet.getCurrentCell().offset(1, -5).activate();
  spreadsheet.getRange('Inicio!B7:E30').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(24, 0).activate();
  spreadsheet.getRange('Inicio!G6:G7').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('Inicio!G8:G9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('Inicio!H6:H7').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('Inicio!H8:H9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getCurrentCell().offset(1, 3).activate();
  spreadsheet.getRange('PrecioStock!E3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 6).activate();
  spreadsheet.getRange('PrecioStock!E9').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getRange('PrecioStock!E10').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(-1, 0).activate();
  spreadsheet.getCurrentCell().setValue('Promo');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Kit');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Postres');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Dips');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Chipá');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Envases');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Cervezas');
  spreadsheet.getCurrentCell().offset(0, -1).activate();
  spreadsheet.getCurrentCell().setValue('Envío');
  OrdenarPedido();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(-1, 0).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();
  spreadsheet.getCurrentCell().offset(1, 0).activate();
  spreadsheet.getRange('Inicio!F7').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getCurrentCell().offset(0, 1).activate();
  spreadsheet.getActiveRange().setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireCheckbox('Finalizado', 'Falta algo')
  .build());
  spreadsheet.getRangeList([spreadsheet.getCurrentCell().offset(-1, -1, 2, 2).getA1Notation(),
  spreadsheet.getCurrentCell().offset(-1, -5, 1, 4).getA1Notation()]).activate()
  .setBackground('#d9d2e9');
  spreadsheet.getCurrentCell().offset(0, 6).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getCurrentCell().offset(0, 0, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#efefef');
  spreadsheet.getCurrentCell().offset(0, 1, 2, 2).activate();
  spreadsheet.getActiveRangeList().setBackground('#d5a6bd');
  spreadsheet.getCurrentCell().offset(0, 2, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#fff2cc');
  spreadsheet.getCurrentCell().offset(0, 1, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#d9ead3');
  spreadsheet.getCurrentCell().offset(0, 1, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#ea9999');
  spreadsheet.getCurrentCell().offset(0, 1, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#efefef');
  spreadsheet.getCurrentCell().offset(0, 1, 2, 1).activate();
  spreadsheet.getActiveRangeList().setBackground('#d9d2e9');
  spreadsheet.getRange('A1').activate();
  SpreadsheetApp.flush()
  LimpiarPedido();
  SpreadsheetApp.getUi().alert("Pedido cargado! 🎉","Hay que chequear que no se dupliquen pedidos",SpreadsheetApp.getUi().ButtonSet.OK);
};

function CargarKit(){
  SpreadsheetApp.getUi()
      .createMenu('Custom Menu')
      .addItem('Show alert', 'showAlert')
      .addToUi();
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'Cargaste todo? Segure?',
     'Lo único que puede quedar en blanco es "Pagado a:" ',
      ui.ButtonSet.YES_NO);
  if (result == ui.Button.YES) {
    CargarKitSeguro();
  } else {
    ui.alert('Acomodalo porfi');
  }
};

function ÚltimaFila() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1000').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();