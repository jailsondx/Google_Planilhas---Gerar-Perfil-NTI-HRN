/*

DEV: JAILSON A DE SOUSA - TÉCNICO DE INFORMÁTICA
ENTIDADE: HOSPITAL REGIONAL NORTE - NÚCLEO DE TECNOLOGIA DA INFORMAÇÃO
Since 2020
42

*/

//FUNÇÃO PRINCIPAL
function Gerar_Perfil() {
  
  //DEFINIÇÃO DAS PLANILHAS
  var sheet = SpreadsheetApp.getActiveSpreadsheet(); //ATIVA E ATRIBUI A PLANILHA ATIVA A VARIAVEL SHEET
  var sheet_todos = sheet.getSheetByName("Todos"); //ATRIBUI A ABA "TODOS" A VARIAVEL SHEET_TODOS
  var sheet_localiza = sheet.getSheetByName("LOCALIZAÇÃO"); //ATRIBUI A ABA "LOCALIZAÇÃO" A VARIAVEL SHEET_LOCALIZA
  
  //DECLARAÇÃO DE VARIAVEIS
  var localizacao;
  var quantidade;
  var linha = 2;
  var row = 1;
  
  //CONTADOR DE CHAMADOS DE CADA SETOR
  for (var todos_colC = sheet_todos.getRange(row,3); sheet_todos.getRange(row,3).getValue() != ""; row++){ //CONTA ATÉ QUE HAJA UM VALOR VAZIO NA CÉLULA
    var valor = sheet_todos.getRange(row,3).getValue();
    var setormacro = sheet_todos.getRange(row,15); //DEFINE A COLUNA O DA ABA TODOS
    
    if (valor.indexOf("ADMINISTRAÇÃO >") > -1){
      setormacro.setValue("ADMINISTRAÇÃO");
      }
    if (valor.indexOf("NAF") > -1){
      setormacro.setValue("NAF");
      }
    if (valor.indexOf("NAC") > -1){
      setormacro.setValue("NAC");
      }
    if ((valor.indexOf("NUTRI") > -1) || (valor.indexOf("BANCO DE LEITE") > -1)){
      setormacro.setValue("NUTRIÇÃO");
      }
    if ((valor.indexOf("CLINICA") > -1) || (valor.indexOf("CLÍNICA") > -1) || (valor.indexOf("UCE") > -1)) {
      if ((valor.indexOf("OBST") > -1) || (valor.indexOf("FARM") > -1) || (valor.indexOf("ENGENHARIA") > -1) || (valor.indexOf("SOCIAL") > -1)) {/*NADA A FAZER*/} 
      else {
      setormacro.setValue("CLINICAS/INTERNAÇÕES");
        }
      }
    if ((valor.indexOf("URG") > -1) || (valor.indexOf("EMERG") > -1) || (valor.indexOf("CCA") > -1)){
      if ((valor.indexOf("FARM") > -1) || (valor.indexOf("NAC") > -1) || (valor.indexOf("CASRM")> -1) || (valor.indexOf("SOCIAL")> -1)) {/*NADA A FAZER*/} 
      else {
      setormacro.setValue("URGÊNCIA/EMERGÊNCIA");
        }
      }
    if ((valor.indexOf("FARM") > -1) || (valor.indexOf("CETIP") > -1)){
    setormacro.setValue("FARMÁCIA");
      }
    if (valor.indexOf("LAB") > -1){
    setormacro.setValue("LABORATÓRIO");
      }
    if ((valor.indexOf("CASRM") > -1) || (valor.indexOf("PARTO NORMAL") > -1) || (valor.indexOf("NEONATAL") > -1) || (valor.indexOf("OBST") > 7)){
      if ((valor.indexOf("CCO") > -1) || (valor.indexOf("SOCIAL") > -1)) {/*NADA A FAZER*/} 
      else {
      setormacro.setValue("CASRM");
        }
      }
    if ((valor.indexOf("UTI AD") > -1) || (valor.indexOf("UTI PED") > -1)){
     if ((valor.indexOf("CETIP") > -1) || (valor.indexOf("NEONATAL") > -1) || (valor.indexOf("FARM") > -1)) {/*NADA A FAZER*/} 
      else {
        setormacro.setValue("UTIs");
        }
      }
    if ((valor.indexOf("CENTRO DE IMAGEM") > -1) || (valor.indexOf("AMBULATÓRIO") > -1)){
     if ((valor.indexOf("CASRM") > -1) || (valor.indexOf("NUTRI") > -1) || (valor.indexOf("SOCIAL") > -1) || (valor.indexOf("LAB") > -1)) {/*NADA A FAZER*/} 
      else {
      setormacro.setValue("CENTRO DE IMAGEM/AMBULATÓRIO");
        }
      } 
    if (valor.indexOf("CC") > -1){
      if (valor.indexOf("CCA") > -1) {/*NADA A FAZER*/} 
      else {
      setormacro.setValue("CCG/CCO");
        }
     }
    if ((valor.indexOf("SOCIAL") > -1) || (valor.indexOf("OUVIDORIA") > -1)){
    setormacro.setValue("SERV. SOCIAL/OUVIDORIA");
      }
    if ((valor.indexOf("CENTRO DE ESTUDOS") > -1) || 
    (valor.indexOf("ENGENHARIA") > -1) || 
    (valor.indexOf("SESMT") > -1) || 
    (valor.indexOf("AGÊNCIA TRANSFUSIONAL") > -1) || 
    (valor.indexOf("EQUIPAMENTOS") > -1) || 
    (valor.indexOf("MANUTENÇÃO") > -1) || 
    (valor.indexOf("CME") > -1) || 
    (valor.indexOf("TRANSPORTE") > -1)){
      setormacro.setValue("OUTROS");
      }
      
    todos_colC = sheet_todos.getRange(row,3);
    }//FIM FOR

  //ALERTA DE FIM DA EXECUÇÃO DO SCRIPT
  var interface = SpreadsheetApp.getUi()
  interface.alert("PERFIL GERADO");
  
  SpreadsheetApp.flush();
} // FIM FUNCTION