//FUNÇÃO PRINCIPAL
function ContaTudov2() {
  
  //DEFINIÇÃO DAS PLANILHAS
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_todos = sheet.getSheetByName("Todos");
  var sheet_localiza = sheet.getSheetByName("LOCALIZAÇÃO");
  
  //DECLARAÇÃO DE VARIAVEIS
  var localizacao;
  var quantidade;
  var linha = 2;
  var cont_ADM = 0;
  var cont_NAF = 0;
  var cont_NAC = 0;
  var cont_NUTRI = 0;
  var cont_CASRM = 0;
  var cont_CLINICA = 0;
  var row = 1;
  
  for (var todos_colC = sheet_todos.getRange(row,3); sheet_todos.getRange(row,3).getValue() != ""; row++){
    var valor = sheet_todos.getRange(row,3).getValue();
    
    if (valor.indexOf("ADM") > -1){
      cont_ADM++;
      }
    if (valor.indexOf("NAF") > -1){
      cont_NAF++;
      }
    if (valor.indexOf("NAC") > -1){
      cont_NAC++;
      }
    if (valor.indexOf("NUTRI") > -1){
      cont_NUTRI++;
      }
    if ((valor.indexOf("CLINICA") > -1) || (valor.indexOf("CLÍNICA") > -1) || (valor.indexOf("UCE") > -1)){
      cont_CLINICA++;
      }
    todos_colC = sheet_todos.getRange(row,3);
    }//FIM FOR
  
  localizacao = sheet_localiza.getRange(linha,1);
  quantidade = sheet_localiza.getRange(linha,2);
  for (linha = 2; linha < 10; linha++){
    if (linha == 3){
    localizacao.setValue("ADMINISTRAÇÃO: ");
    quantidade.setValue(cont_ADM);
    }
    if (linha == 4){
    localizacao.setValue("NAF: ");
    quantidade.setValue(cont_NAF);
    }
    if (linha == 5){
    localizacao.setValue("NAC: ");
    quantidade.setValue(cont_NAC);
    }
    if (linha == 6){
    localizacao.setValue("NUTRIÇÃO: ");
    quantidade.setValue(cont_NUTRI);
    }
    if (linha == 7){
    localizacao.setValue("CLINICAS/INTERNAÇÕES: ");
    quantidade.setValue(cont_CLINICA);
    }
    localizacao = sheet_localiza.getRange(linha,1);
    quantidade = sheet_localiza.getRange(linha,2);
  } //FIM FOR
} // FIM FUNCTION