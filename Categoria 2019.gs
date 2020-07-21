/*

DEV: JAILSON A DE SOUSA - TÉCNICO DE INFORMÁTICA
ENTIDADE: HOSPITAL REGIONAL NORTE - NÚCLEO DE TECNOLOGIA DA INFORMAÇÃO
Since 2020
42

*/

//FUNÇÃO PRINCIPAL
function Categoria_2019() {
  
  //DEFINIÇÃO DAS PLANILHAS
  var sheet = SpreadsheetApp.getActiveSpreadsheet(); //ATIVA E ATRIBUI A PLANILHA ATIVA A VARIAVEL SHEET
  var sheet_todos = sheet.getSheetByName("Todos"); //ATRIBUI A ABA "TODOS" A VARIAVEL SHEET_TODOS
  //var sheet_localiza = sheet.getSheetByName("LOCALIZAÇÃO"); //ATRIBUI A ABA "LOCALIZAÇÃO" A VARIAVEL SHEET_LOCALIZA
  
  //DECLARAÇÃO DE VARIAVEIS
  var localizacao;
  var quantidade;
  var linha = 2;
  var row = 1;
  
  //CONTADOR DE CHAMADOS DE CADA SETOR
  for (var todos_colC = sheet_todos.getRange(row,1); sheet_todos.getRange(row,1).getValue() != ""; row++){ //CONTA ATÉ QUE HAJA UM VALOR VAZIO NA CÉLULA DA COLUNA "A"
    var valor = sheet_todos.getRange(row,2).getValue(); //RECEBE O VALOR DA CÉLULA DA COLUNA "C" E ATRIBUI A VARIAVEL "VALOR" 
    var setormacro = sheet_todos.getRange(row,12); //DEFINE A COLUNA "U" DA ABA TODOS E ATRIBUI A VARIAVEL "SETORMACRO"
    
    //CATEGORIA IMPRESSORA
    if ( (valor.indexOf("IMPRESS") > -1) || 
      (valor.indexOf("MULTI") > -1) ||
      (valor.indexOf("PAPEL") > -1) || 
      (valor.indexOf("IMPRIM") > -1) ||
      (valor.indexOf("FOLHA") > -1) ||
      (valor.indexOf("PRES") > -1) ||
      (valor.indexOf("ATOLAD") > -1) )
      {
      if ( (valor.indexOf("ZEBRA") > -1) || (valor.indexOf("ETIQUETA") > -1) || (valor.indexOf("REVELADOR") > -1) || (valor.indexOf("TONNER") > -1) || (valor.indexOf("TONER") > -1) ) {/*NADA A FAZER*/} 
      else {
        setormacro.setValue("IMPRESSORA");
            }
      }
    //CATEGORIA TONNER  
    if ( (valor.indexOf("TONNER") > -1) || (valor.indexOf("TONER") > -1) || (valor.indexOf("REVELADOR") > -1) )
      {
      setormacro.setValue("TONNER");
      }
      
    //CATEGORIA COMPUTADOR  
    if ( (valor.indexOf("COMPUTADOR") > -1) || 
      (valor.indexOf("PC") > -1) || 
      (valor.indexOf("ÁREA") > -1) || 
      (valor.indexOf("AREA") > -1) ||
      (valor.indexOf("MOUSE") > -1) ||
      (valor.indexOf("TECLADO") > -1) || 
      (valor.indexOf("MONITOR") > -1) ||
      (valor.indexOf("CPU") > -1) ||
      (valor.indexOf("DRIVE") > -1) ||
      (valor.indexOf("LEITOR") > -1) ||
      (valor.indexOf("DVD") > -1) ||
      (valor.indexOf("SCAN") > -1) || 
      (valor.indexOf("MANUTENÇÃO PREVENTIVA") > -1) )
      {
      if (valor.indexOf("TEL") > -1) {/*NADA A FAZER*/} 
      else {
        setormacro.setValue("COMPUTADOR");
            }
      }
      
    //CATEGORIA SISTEMAS  
    if ( (valor.indexOf("SISTEMA") > -1) || 
      (valor.indexOf("RM") > -1) || 
      (valor.indexOf("GERCOMP") > -1) ||
      (valor.indexOf("SISAIH") > -1) ||
      (valor.indexOf("BPA") > -1) ||
      (valor.indexOf("TOTVS") > -1) ||
      (valor.indexOf("NF") > -1) ||
      (valor.indexOf("VITA") > -1) ||
      (valor.indexOf("INTERFACI") > -1) ||
      (valor.indexOf("PRONT") > -1) )
      {
      if ( (valor.indexOf("CASRM") > -1) || (valor.indexOf("IMPRESSORA") > -1) || (valor.indexOf("PRONTO") > -1) || (valor.indexOf("SCAN") > -1) || (valor.indexOf("ACESSO") > -1) || (valor.indexOf("PERMISS") > -1) ) {/*NADA A FAZER*/} 
      else {
        setormacro.setValue("SISTEMAS");
            }
      }
      
    //CATEGORIA TELEFONIA  
    if ( (valor.indexOf("TEL") > -1) || 
      (valor.indexOf("HEADSET") > -1) || 
      (valor.indexOf("DISCA") > -1) ||
      (valor.indexOf("CHAMA") > -1) ||
      (valor.indexOf("LIGA") > -1) ||
      (valor.indexOf("RAMAL") > -1) ||
      (valor.indexOf("APARELHO") > -1) )
      {
      if ( (valor.indexOf("COMPUTADOR") > -1) || (valor.indexOf("PC") > -1) ){/*NADA A FAZER*/} 
      else {
        setormacro.setValue("TELEFONIA");
            }
      }
      
      
    //CATEGORIA PERMISSÃO DE USUÁRIO  
    if ( (valor.indexOf("ACESS") > -1) || 
      (valor.indexOf("PERMIS") > -1) ||
      (valor.indexOf("ALMOX") > -1) ||
      (valor.indexOf("LIBERA") > -1) ||
      (valor.indexOf("USUARIO") > -1) || 
      (valor.indexOf("USUÁRIO") > -1) )
      {
      if ( (valor.indexOf("ÁREA") > -1) || (valor.indexOf("AREA") > -1) ){/*NADA A FAZER*/} 
      else {
        setormacro.setValue("PERMISSÃO DE USUÁRIO");
            }
      }
    
    //CATEGORIA IMPRESSORA DE ETIQUETAS
    if ( (valor.indexOf("ZEBRA") > -1) || 
      (valor.indexOf("ETIQUETA") > -1) )
      {
      if ( (valor.indexOf("PAPEL") > -1) || (valor.indexOf("PRESO") > -1) || (valor.indexOf("REVELADOR") > -1) || (valor.indexOf("TONNER") > -1) || (valor.indexOf("TONER") > -1) ) {/*NADA A FAZER*/} 
      else {
        setormacro.setValue("IMPRESSORA DE ETIQUETAS");
            }
      }
      
    //CATEGORIA INFRAESTRUTURA
    if ( (valor.indexOf("REMANEJA") > -1) || 
      (valor.indexOf("INTERNET") > -1) || 
      (valor.indexOf("REDE") > -1) ||
      (valor.indexOf("PONTO") > -1) ||
      (valor.indexOf("RELÓGIO") > -1) ||
      (valor.indexOf("RELOGIO") > -1) )
      {
      setormacro.setValue("INFRAESTRUTURA");
      }
      
    //CATEGORIA OUTROS
    if ( (valor.indexOf("CRACHÁ") > -1) ||
      (valor.indexOf("CRACHA") > -1) ||
      (valor.indexOf("CHIP") > -1) ||
      (valor.indexOf("ARTE") > -1) )
      {
      setormacro.setValue("OUTROS");
      }
      
    todos_colC = sheet_todos.getRange(row,1);
    }//FIM FOR

  //ALERTA DE FIM DA EXECUÇÃO DO SCRIPT
  var interface = SpreadsheetApp.getUi()
  interface.alert("PERFIL CATEGORIAS GERADO");
  
  SpreadsheetApp.flush();
} // FIM FUNCTION