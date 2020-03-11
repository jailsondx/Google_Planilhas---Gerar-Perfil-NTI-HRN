//FUNÇÃO PRINCIPAL
function Gerar_Perfil() {
  
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
  var cont_URG = 0;
  var cont_FARM = 0;
  var cont_LAB = 0;
  var cont_CASRM = 0;
  var cont_UTI = 0;
  var cont_CC = 0;
  var cont_IMG = 0;
  var cont_SS = 0;
  var cont_OUTROS = 0;
  var row = 1;
  
  //CONTADOR DE CHAMADOS DE CADA SETOR
  for (var todos_colC = sheet_todos.getRange(row,3); sheet_todos.getRange(row,3).getValue() != ""; row++){
    var valor = sheet_todos.getRange(row,3).getValue();
    
    if (valor.indexOf("ADMINISTRAÇÃO >") > -1){
      cont_ADM++;
      //sheet_todos.getRange(row,3).setValue("ADMINISTRAÇÃO");
      }
    if (valor.indexOf("NAF") > -1){
      cont_NAF++;
      //sheet_todos.getRange(row,3).setValue("NAF");
      }
    if (valor.indexOf("NAC") > -1){
      cont_NAC++;
      //sheet_todos.getRange(row,3).setValue("NAC");
      }
    if ((valor.indexOf("NUTRI") > -1) || (valor.indexOf("BANCO DE LEITE") > -1)){
      cont_NUTRI++;
      //sheet_todos.getRange(row,3).setValue("NUTRIÇÃO");
      }
    if ((valor.indexOf("CLINICA") > -1) || (valor.indexOf("CLÍNICA") > -1) || (valor.indexOf("UCE") > -1)) {
      if ((valor.indexOf("OBST") > -1) || (valor.indexOf("FARM") > -1) || (valor.indexOf("ENGENHARIA") > -1)) {/*NADA A FAZER*/} 
      else {
        cont_CLINICA++;
        //sheet_todos.getRange(row,3).setValue("CLINICAS/INTERNAÇÕES");
        }
      }
    if ((valor.indexOf("URG") > -1) || (valor.indexOf("EMERG") > -1) || (valor.indexOf("CCA") > -1)){
      if ((valor.indexOf("FARM") > -1) || (valor.indexOf("NAC") > -1) || (valor.indexOf("CASRM")> -1) || (valor.indexOf("SOCIAL")> -1)) {/*NADA A FAZER*/} 
      else {
        cont_URG++;
        //sheet_todos.getRange(row,3).setValue("URG/EMERG");
        }
      }
    if ((valor.indexOf("FARM") > -1) || (valor.indexOf("CETIP") > -1)){
      cont_FARM++;
      //sheet_todos.getRange(row,3).setValue("FARMA");
      }
    if (valor.indexOf("LAB") > -1){
      cont_LAB++;
      //sheet_todos.getRange(row,3).setValue("LABORA");
      }
    if ((valor.indexOf("CASRM") > -1) || (valor.indexOf("PARTO NORMAL") > -1) || (valor.indexOf("NEONATAL") > -1) || (valor.indexOf("OBST") > 7)){
      if ((valor.indexOf("CCO") > -1) || (valor.indexOf("SOCIAL") > -1)) {/*NADA A FAZER*/} 
      else {
      cont_CASRM++;
      //sheet_todos.getRange(row,3).setValue("CASRM/CPN/NEO");
        }
      }
    if ((valor.indexOf("UTI AD") > -1) || (valor.indexOf("UTI PED") > -1)){
     if ((valor.indexOf("CETIP") > -1) || (valor.indexOf("NEONATAL") > -1) || (valor.indexOf("FARM") > -1)) {/*NADA A FAZER*/} 
      else {
        cont_UTI++;
        //sheet_todos.getRange(row,3).setValue("UTI AD/PED");
        }
      }
    if ((valor.indexOf("CENTRO DE IMAGEM") > -1) || (valor.indexOf("AMBULATÓRIO") > -1)){
     if ((valor.indexOf("CASRM") > -1) || (valor.indexOf("NUTRI") > -1) || (valor.indexOf("SOCIAL") > -1) || (valor.indexOf("LAB") > -1)) {/*NADA A FAZER*/} 
      else {
        cont_IMG++;
        //sheet_todos.getRange(row,3).setValue("CI/AMB");
        }
      } 
    if (valor.indexOf("CC") > -1){
      if (valor.indexOf("CCA") > -1) {/*NADA A FAZER*/} 
      else {
      cont_CC++;
      //sheet_todos.getRange(row,3).setValue("CENTROS CIRURGICOS");
        }
     }
    if ((valor.indexOf("SOCIAL") > -1) || (valor.indexOf("OUVIDORIA") > -1)){
      cont_SS++;
      }
    if ((valor.indexOf("CENTRO DE ESTUDOS") > -1) || (valor.indexOf("ENGENHARIA") > -1) || (valor.indexOf("SESMT") > -1) || (valor.indexOf("AGÊNCIA TRANSFUSIONAL") > -1) || (valor.indexOf("EQUIPAMENTOS") > -1) || (valor.indexOf("MANUTENÇÃO") > -1) || (valor.indexOf("CME") > -1)){
      cont_OUTROS++;
      //sheet_todos.getRange(row,3).setValue("OUTROS");
      }
      
    todos_colC = sheet_todos.getRange(row,3);
    }//FIM FOR
  
  //IMPRIMI O SETOR E A QUANTIDADE DE CHAMADOS DELE
  localizacao = sheet_localiza.getRange(linha,1);
  quantidade = sheet_localiza.getRange(linha,2);
  
  for (linha = 2; linha < 18; linha++){
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
    if (linha == 8){
    localizacao.setValue("URGÊNCIA/EMERGÊNCIA: ");
    quantidade.setValue(cont_URG);
    }
    if (linha == 9){
    localizacao.setValue("FARMÁCIA: ");
    quantidade.setValue(cont_FARM);
    }
    if (linha == 10){
    localizacao.setValue("LABORATÓRIO: ");
    quantidade.setValue(cont_LAB);
    }
    if (linha == 11){
    localizacao.setValue("CASRM: ");
    quantidade.setValue(cont_CASRM);
    }
    if (linha == 12){
    localizacao.setValue("UTIs: ");
    quantidade.setValue(cont_UTI);
    }
    if (linha == 13){
    localizacao.setValue("CENTRO DE IMAGEM/AMBULATÓRIO: ");
    quantidade.setValue(cont_IMG);
    }
    if (linha == 14){
    localizacao.setValue("CCG/CCO: ");
    quantidade.setValue(cont_CC);
    }
    if (linha == 15){
    localizacao.setValue("SERV. SOCIAL/OUVIDORIA: ");
    quantidade.setValue(cont_SS);
    }
    if (linha == 16){
    localizacao.setValue("OUTROS: ");
    quantidade.setValue(cont_OUTROS);
    }
    
    localizacao = sheet_localiza.getRange(linha,1);
    quantidade = sheet_localiza.getRange(linha,2);
  } //FIM FOR
  
  //SOMA TODOS OS CHAMADOS E ESCREVE ABAIXO DO ULTIMO SETOR PREENCHIDO
  localizacao = sheet_localiza.getRange(linha-1,1); //DEIXA 1 LINHA DE ESPAÇO EM BRANCO COM RELAÇÃO A ULTIMA LINHA PREENCHIDA
  quantidade = sheet_localiza.getRange(linha-1,2); //DEIXA 1 LINHA DE ESPAÇO EM BRANCO COM RELAÇÃO A ULTIMA LINHA PREENCHIDA
  localizacao.setValue("TOTAL ");
  quantidade.setValue(cont_ADM + cont_NAF + cont_NAC + cont_NUTRI + cont_CLINICA + cont_URG + cont_FARM + cont_LAB + cont_CASRM + cont_UTI + cont_IMG + cont_CC + cont_SS + cont_OUTROS);
  
  //ALERTA DE FIM DA EXECUÇÃO DO SCRIPT
  var interface = SpreadsheetApp.getUi()
  interface.alert("PERFIL GERADO");
  
  SpreadsheetApp.flush();
} // FIM FUNCTION