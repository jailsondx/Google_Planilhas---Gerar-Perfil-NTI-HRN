//FUNÇÕES ASSISTENCIAIS/DE PESQUISA
function Setores(localizacao, quantidade, cont, row, sheet_todos, search, TITLE){

for (var col_c = sheet_todos.getRange(row,3); sheet_todos.getRange(row,3).getValue() != ""; row++){
    var valor = sheet_todos.getRange(row,3).getValue();
    if (valor.indexOf(search) > -1){
      cont++;
      }
    col_c = sheet_todos.getRange(row,3);
    }
    localizacao.setValue(TITLE + ": ");
    quantidade.setValue(cont);
}

function Setores2(localizacao, quantidade, cont, row, sheet_todos, search1, search2, TITLE){

for (var col_c = sheet_todos.getRange(row,3); sheet_todos.getRange(row,3).getValue() != ""; row++){
    var valor = sheet_todos.getRange(row,3).getValue();
    if ((valor.indexOf(search1) > -1) || (valor.indexOf(search2) > -1)){
      cont++;
      }
    col_c = sheet_todos.getRange(row,3);
    }
    localizacao.setValue(TITLE + ": ");
    quantidade.setValue(cont);
}

function Setores3(localizacao, quantidade, cont, row, sheet_todos, search1, search2, search3, TITLE){

for (var col_c = sheet_todos.getRange(row,3); sheet_todos.getRange(row,3).getValue() != ""; row++){
    var valor = sheet_todos.getRange(row,3).getValue();
    if ((valor.indexOf(search1) > -1) || (valor.indexOf(search2) > -1) || (valor.indexOf(search3) > -1)){
      cont++;
      }
    col_c = sheet_todos.getRange(row,3);
    }
    localizacao.setValue(TITLE + ": ");
    quantidade.setValue(cont);
}

//FUNÇÃO PRINCIPAL
function ContaTudo() {
  
  //DEFINIÇÃO DAS PLANILHAS
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_todos = sheet.getSheetByName("Todos");
  var sheet_localiza = sheet.getSheetByName("LOCALIZAÇÃO");
  
  //DECLARAÇÃO DE VARIAVEIS
  var localizacao;
  var quantidade;
  var TITLE;
  var search;
  var search1;
  var search2;
  var search3;
  var cont = 0;
  var row = 1;
  
  //INICIO DAS CHAMADAS DE FUNÇÕES DE BUSCA
  search = 'NUTRI';
  TITLE = 'NUTRIÇÃO';
  localizacao = sheet_localiza.getRange(2,1); //LINHA X COLUNA
  quantidade = sheet_localiza.getRange(2,2);
  Setores(localizacao, quantidade, cont, row, sheet_todos, search, TITLE);
  
  search = 'NAF';
  TITLE = 'NAF';
  localizacao = sheet_localiza.getRange(3,1); //LINHA X COLUNA
  quantidade = sheet_localiza.getRange(3,2);
  Setores(localizacao, quantidade, cont, row, sheet_todos, search, TITLE);
  
  search = 'NAC';
  TITLE = 'NAC';
  localizacao = sheet_localiza.getRange(4,1);
  quantidade = sheet_localiza.getRange(4,2);
  Setores(localizacao, quantidade, cont, row, sheet_todos, search, TITLE);
  
  search1 = 'CASRM';
  search2 = 'PARTO NORMAL';
  TITLE = 'CASRM';
  localizacao = sheet_localiza.getRange(5,1);
  quantidade = sheet_localiza.getRange(5,2);
  Setores2(localizacao, quantidade, cont, row, sheet_todos, search1, search2, TITLE);
  
  search = 'ADMIN';
  TITLE = 'ADMINISTRAÇÃO';
  localizacao = sheet_localiza.getRange(6,1);
  quantidade = sheet_localiza.getRange(6,2);
  Setores(localizacao, quantidade, cont, row, sheet_todos, search, TITLE);
  
  search1 = 'CLINICA';
  search2 = 'CLÍNICA';
  search3 = 'UCE';
  TITLE = 'CLÍNICAS/INTERNAÇÕES';
  localizacao = sheet_localiza.getRange(7,1);
  quantidade = sheet_localiza.getRange(7,2);
  Setores3(localizacao, quantidade, cont, row, sheet_todos, search1, search2, search3, TITLE);
  
  search = 'FARM';
  TITLE = 'FARMÁCIA';
  localizacao = sheet_localiza.getRange(8,1);
  quantidade = sheet_localiza.getRange(8,2);
  Setores(localizacao, quantidade, cont, row, sheet_todos, search, TITLE);
  
  search = 'SERVIÇO SOCIAL';
  TITLE = 'SERVIÇO SOCIAL';
  localizacao = sheet_localiza.getRange(9,1);
  quantidade = sheet_localiza.getRange(9,2);
  Setores(localizacao, quantidade, cont, row, sheet_todos, search, TITLE);
  
  search = 'NEONATAL';
  TITLE = 'UTI/UCI NEONATAL';
  localizacao = sheet_todos.getRange(10,1);
  quantidade = sheet_localiza.getRange(10,2);
  Setores(localizacao, quantidade, cont, row, sheet_todos, search, TITLE);
  
  search = 'CC';
  TITLE = 'CCG/CCO';
  localizacao = sheet_localiza.getRange(11,1);
  quantidade = sheet_localiza.getRange(11,2);
  Setores(localizacao, quantidade, cont, row, sheet_todos, search, TITLE);
  
  var interface = SpreadsheetApp.getUi()
  interface.alert("CONTADOR DE CHAMADOS FINALIZADO");
  SpreadsheetApp.flush();
  
  //celula.searchValue();
  //celula.searchValue("SOMA");
  
  
}
