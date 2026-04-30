"use strict";

/**
 * Módulo:    tabelasFila.gs
 * Objetivo:  Armazenar Fila de casos do sistema Bolsa Moradia
 */



/**
 * Planilha FILA
 */
const PLANILHA_FILA_ID        =  "1du1DlnwZBBAM8du3HNpA9Hpk70XtdSAGKlylGkZX2Nw";
const PLANILHA_FILA           =  SpreadsheetApp.openById(PLANILHA_FILA_ID);
const TABELA_FILA             =  PLANILHA_FILA.getSheetByName('FILA');
let BUFFER_FILA               =  TABELA_FILA.getDataRange().getDisplayValues().splice(1);
let TAMANHO_FILA              =  BUFFER_FILA.length;
const NUM_COLUNAS_TABELA_FILA =  18;

function refreshBufferFila() {
  BUFFER_FILA  =  TABELA_FILA.getDataRange().getDisplayValues().splice(1);
  TAMANHO_FILA = BUFFER_FILA.length;
}



// Posições das colunas da planilha FILA
const REFERENCIA_FAMILIAR        =  1;
const CPF_RF                     =  2;
const ORGAO_ENCAMINHADOR         =  3;
const DATA_ENCAMINHAMENTO        =  4;
const PONTUACAO                  =  5;

const IDS_PARAMETROS_CASO        =  6;
const PONTUACOES_PARAMETROS_CASO =  7;

const QUANTIDADE_CEA             =  8;
const PROBLEMAS_SAUDE            =  9;
const DATA_NASCIMENTO_RF         = 10;
const TEMPO_SITUACAO_DE_RUA      = 11;

const SITUACAO_BENEFICIO         = 12;
const DATA_ULTIMA_EVOLUCAO       = 13;
const DOC_PENDENTE               = 14;
const EMAIL_ORGAO_ENCAMINHADOR   = 15;
const DATA_LIMITE                = 16;
const JUSTIFICATIVA_ALTERACAO_DATA_LIMITE = 17;



/** 
 *  ####################################################
 *  #####                                          ##### 
 *  #####  IMPLEMENTAÇÃO DAS FUNÇÕES DESSE MÓDULO  #####
 *  #####                                          ##### 
 *  ####################################################
 */



/**
 * Função que limpa a fila, excluíndo todos os casos dela
 */
function limparFila() {

  // Caso nulo
  let casoNulo = new Array(NUM_COLUNAS_TABELA_FILA).fill("");

  // Limpa a fila
  let range;
  for( let linha=2; linha<=TAMANHO_FILA+1; ++linha ) {
    range = TABELA_FILA.getRange( linha, 1, 1, NUM_COLUNAS_TABELA_FILA );
    range.setValues([casoNulo]);
  }

  // Flush na planilha
  try {
    SpreadsheetApp.flush();
    PLANILHA_FILA.waitForAllDataExecutionsCompletion(2);      
  } catch( error ) {
    throw( "limparFila: " + error.message );
  }

} // Fim da função limparFila



/**
 * Função que carrega a fila a partir das tabelas de encaminhamentos
 */
function carregarFila() {

  let caso = [];
  let resumoCaso = new Array(NUM_COLUNAS_TABELA_FILA).fill("");
  let id = 0;

  let cpfRF = "";

  
  console.log( "\n\nCASOS EXTERNOS" );
  linhaTabela = 0;
  while( linhaTabela < NUM_LINHAS_TABELA_CASOS_EXTERNOS ) {

    caso = obterCaso( BUFFER_CASOS_EXTERNOS );    
    cpfRF = String( caso[0][UNI_CPF_RF] );

    ++id;
    resumoCaso[ID] = String( id );
    resumoCaso[REFERENCIA_FAMILIAR] = (caso[0][UNI_NOME]).trim().toUpperCase();
    resumoCaso[CPF_RF] = cpfRF.padStart(11, "0");
    resumoCaso[ORGAO_ENCAMINHADOR] = String( caso[0][UNI_ORGAO_ENCAMINHADOR] );
    
    resumoCaso[DATA_ENCAMINHAMENTO] = String( caso[0][UNI_DATA_REGISTRO_ENCAMINHAMENTO] );
    resumoCaso[PONTUACAO] = String( calcularPontuacao( caso ) );

    resumoCaso[IDS_PARAMETROS_CASO] = idsParametrosCaso.join(";");
    resumoCaso[PONTUACOES_PARAMETROS_CASO] = pontuacoesParametrosCaso.join(";");

    resumoCaso[QUANTIDADE_CEA] = String( numeroDeCEAs( caso ) );
    resumoCaso[PROBLEMAS_SAUDE] = String( numeroDeProblemasDeSaude( caso ) );
    resumoCaso[DATA_NASCIMENTO_RF] = String( caso[0][UNI_DATA_NASCIMENTO] );
    resumoCaso[TEMPO_SITUACAO_DE_RUA] = String( caso[0][UNI_TEMPO_SITUACAO_DE_RUA] );

    resumoCaso[SITUACAO_BENEFICIO] = "2"; // 2 == Habilitado e ainda não convocado

    resumoCaso[DATA_ULTIMA_EVOLUCAO] = "";

    resumoCaso[DOC_PENDENTE] = "1"; // 1 == Não

    resumoCaso[EMAIL_ORGAO_ENCAMINHADOR] = String( caso[0][UNI_EMAIL_TECNICO_ENCAMINHADOR] );
    
    TABELA_FILA.appendRow( resumoCaso );
  }  

  
  console.log( "\n\nCASOS PBH" );
  linhaTabela = 0;
  while( linhaTabela < NUM_LINHAS_TABELA_CASOS_PBH ) {

    caso = obterCaso( BUFFER_CASOS_PBH );
    cpfRF = String( caso[0][UNI_CPF_RF] );

    ++id;
    resumoCaso[ID] = String( id );
    resumoCaso[REFERENCIA_FAMILIAR] = (caso[0][UNI_NOME]).trim().toUpperCase();
    resumoCaso[CPF_RF] = cpfRF.padStart(11, "0");
    resumoCaso[ORGAO_ENCAMINHADOR] = String( caso[0][UNI_ORGAO_ENCAMINHADOR] );
    
    resumoCaso[DATA_ENCAMINHAMENTO] = String( caso[0][UNI_DATA_REGISTRO_ENCAMINHAMENTO] );
    resumoCaso[PONTUACAO] = String( calcularPontuacao( caso ) );

    resumoCaso[IDS_PARAMETROS_CASO] = idsParametrosCaso.join(";");
    resumoCaso[PONTUACOES_PARAMETROS_CASO] = pontuacoesParametrosCaso.join(";");

    resumoCaso[QUANTIDADE_CEA] = String( numeroDeCEAs( caso ) );
    resumoCaso[PROBLEMAS_SAUDE] = String( numeroDeProblemasDeSaude( caso ) );
    resumoCaso[DATA_NASCIMENTO_RF] = String( caso[0][UNI_DATA_NASCIMENTO] );
    resumoCaso[TEMPO_SITUACAO_DE_RUA] = String( caso[0][UNI_TEMPO_SITUACAO_DE_RUA] );

    resumoCaso[SITUACAO_BENEFICIO] = "2"; // 2 == Habilitado e ainda não convocado

    resumoCaso[DATA_ULTIMA_EVOLUCAO] = "";

    resumoCaso[DOC_PENDENTE] = "1"; // 1 == Não

    resumoCaso[EMAIL_ORGAO_ENCAMINHADOR] = String( caso[0][UNI_EMAIL_TECNICO_ENCAMINHADOR] );
    
    TABELA_FILA.appendRow( resumoCaso );        
  }

  // Flush na planilha
  try {
    SpreadsheetApp.flush();
    PLANILHA_FILA.waitForAllDataExecutionsCompletion(2);      
  } catch( error ) {
    throw( error.message );
  }

  refreshBufferFila();

} // Fim da função carregarFila



/**
 * Função que retorna a fila com os casos registrados no sistema
 * 
 * @return Uma fila em que cada posição contém um objeto com os dados de um caso
 */
function obterFila() {    


  // RETORNA NULL, SE TABELA DE CASOS ESTIVER VAZIA
  if( TAMANHO_FILA < 1 ) return null;
    
  
  // Obtém os casos na fila
  let fila = BUFFER_FILA.map( caso => {    
    
    return {

      id: caso[ID],

      referencia_familiar: caso[REFERENCIA_FAMILIAR],

      cpf_rf: caso[CPF_RF].padStart(11, "0"),
 
      id_orgao_encaminhador: caso[ORGAO_ENCAMINHADOR],
      
      email_orgao_encaminhador: caso[EMAIL_ORGAO_ENCAMINHADOR],
            
      ids_parametros_caso: caso[IDS_PARAMETROS_CASO] != "" ? caso[IDS_PARAMETROS_CASO].split(";") : "",

      pontuacoes_parametros_caso:  caso[PONTUACOES_PARAMETROS_CASO] != ""  ? caso[PONTUACOES_PARAMETROS_CASO].split(";") : "",

      pontuacao: (caso[SITUACAO_BENEFICIO] != "" && caso[SITUACAO_BENEFICIO] != "1") ?
                 (caso[PONTUACAO] != "" ? parseInt(caso[PONTUACAO]) : 0) :
                 0,      

      quantidade_CEA: parseInt(caso[QUANTIDADE_CEA]),

      quantidade_problemas_saude: caso[PROBLEMAS_SAUDE] != "" ? parseInt(caso[PROBLEMAS_SAUDE]) : 0,

      idade_RF: calcularIdade( caso[DATA_NASCIMENTO_RF] ),

      data_nascimento_RF: caso[DATA_NASCIMENTO_RF],
      
      id_tempo_nas_ruas: caso[TEMPO_SITUACAO_DE_RUA] != "" ? parseInt(caso[TEMPO_SITUACAO_DE_RUA]) : 0,      

      id_situacao_beneficio: caso[SITUACAO_BENEFICIO], 
        
      data_ultima_evolucao: caso[DATA_ULTIMA_EVOLUCAO],
        
      data_limite: caso[DATA_LIMITE],

      justificativa_alteracao_data_limite: caso[JUSTIFICATIVA_ALTERACAO_DATA_LIMITE],      

      id_doc_pendente: caso[DOC_PENDENTE],

      vistorias: pesquisarVistoriasPorCPF( caso[CPF_RF].padStart(11, "0") ),
           
      posicaoNaFila: 0

    };// Fim return       
  });


  // Ordena os casos pelos pontos, em ordem decrescente  
  //fila.sort( (a,b) => b.pontuacao - a.pontuacao );
  fila.sort( function(a,b) { 

    // Primeiro critério - Pontuação
    if( b.pontuacao > a.pontuacao ) {
      return 1;
    } else if(b.pontuacao < a.pontuacao) {
      return -1;
    } 

    // Segundo critétio - Quantidade C&A
    if( b.quantidade_CEA > a.quantidade_CEA ) {
      return 1;
    } else if(b.quantidade_CEA < a.quantidade_CEA) {
      return -1;
    }     

    // Terceiro critétio - Quantidade Problemas de Saúde
    if( b.quantidade_problemas_saude > a.quantidade_problemas_saude ) {
      return 1;
    } else if(b.quantidade_problemas_saude < a.quantidade_problemas_saude) {
      return -1;
    }         

    // Quarto critétio - Idade RF
    let diferencaIdade = compararDatas( b.data_nascimento_RF, a.data_nascimento_RF );
    if( diferencaIdade > 0 ) {
      return 1;
    } else if( diferencaIdade < 0 ) {
      return -1;
    }  
    
    // Quinto critétio - Tempo nas Ruas    
    if( b.id_tempo_nas_ruas > a.id_tempo_nas_ruas ) {
      return 1;
    } else if(b.id_tempo_nas_ruas < a.id_tempo_nas_ruas) {
      return -1;
    }                
    

    // Retorno Padrão
    return 0;

  });

  
  // Determina a posicao na fila da regional / fila geral
  let posicao = 1;
  fila.forEach( caso => {

    caso.posicaoNaFila = posicao;
    ++posicao; 

  });


  // Retorna a fila filtrada
  //return fila;
  return JSON.stringify( fila );

} // Fim da Função obterFila 





/** 
 *  #################################################
 *  #####                                       ##### 
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 *  #####                                       ##### 
 *  #################################################
 */




/**
 * Função para testar a função principal obterFila
 */
function teste_obterFila() {

  const fila = obterFila();

  console.log(fila);    

} // Fim da Função teste_obterFila 



/**
 * ##### FIM DO MÓDULO tabelasFila.gs #####
 */















