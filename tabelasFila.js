"use strict";

/**
 * Módulo:    tabelasFila.gs
 * Objetivo:  Armazenar Fila de casos do sistema Bolsa Moradia
 */



/**
 * Planilha FILA
 */
const PLANILHA_FILA_ID        =  PropertiesService.getScriptProperties().getProperty('PLANILHA_FILA_ID');
const PLANILHA_FILA           =  SpreadsheetApp.openById(PLANILHA_FILA_ID);
const TABELA_FILA             =  PLANILHA_FILA.getSheetByName('FILA');
let BUFFER_FILA               =  TABELA_FILA.getDataRange().getDisplayValues().splice(1);
let TAMANHO_FILA              =  BUFFER_FILA.length;
const NUM_COLUNAS_TABELA_FILA =  20;

function refreshBufferFila() {
  BUFFER_FILA  =  TABELA_FILA.getDataRange().getDisplayValues().splice(1);
  TAMANHO_FILA = BUFFER_FILA.length;
}


/**
 * Planilha FILA CASOS ANTIGOS
 */
const PLANILHA_FILA_CASOS_ANTIGOS_ID        =  PropertiesService.getScriptProperties().getProperty('PLANILHA_FILA_CASOS_ANTIGOS_ID');
const PLANILHA_FILA_CASOS_ANTIGOS           =  SpreadsheetApp.openById(PLANILHA_FILA_CASOS_ANTIGOS_ID);
const TABELA_FILA_CASOS_ANTIGOS             =  PLANILHA_FILA_CASOS_ANTIGOS.getSheetByName('FILA_CASOS_ANTIGOS');
let BUFFER_FILA_CASOS_ANTIGOS               =  TABELA_FILA_CASOS_ANTIGOS.getDataRange().getDisplayValues().splice(1);
let TAMANHO_FILA_CASOS_ANTIGOS              =  BUFFER_FILA_CASOS_ANTIGOS.length;
const NUM_COLUNAS_TABELA_FILA_CASOS_ANTIGOS =  20;

function refreshBufferFila() {
  BUFFER_FILA_CASOS_ANTIGOS  =  TABELA_FILA_CASOS_ANTIGOS.getDataRange().getDisplayValues().splice(1);
  TAMANHO_FILA_CASOS_ANTIGOS = BUFFER_FILA_CASOS_ANTIGOS.length;
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

const DATA_LIMITE                = 15;
const JUSTIFICATIVA_ALTERACAO_DATA_LIMITE = 16;
const PERFIL_COMPLETO            = 17;
const PERFIL_GENERO              = 18;
const PERFIL_ORIENTACAO_SEXUAL   = 19;



/** 
 *  ####################################################
 *  #####                                          ##### 
 *  #####  IMPLEMENTAÇÃO DAS FUNÇÕES DESSE MÓDULO  #####
 *  #####                                          ##### 
 *  ####################################################
 */



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

    let idCaso = caso[ID];
    
    let dataEncaminhamento = caso[DATA_ENCAMINHAMENTO].split("-");
    let dataEncaminhamentoFormatada = `${dataEncaminhamento[2]}/${dataEncaminhamento[1]}/${dataEncaminhamento[0]}`;

    let servicosReferencia = obterServicosReferenciaDoCaso( String( idCaso ) );
    let servicoReferenciaAtivo = obterServicoReferenciaAtivo( servicosReferencia );

    let vistoriasCaso = pesquisarVistoriasPorCPF( caso[CPF_RF].padStart(11, "0") );
    let idSituacaoVistoria = getSituacaoVistoria( vistoriasCaso );

    let idSituacaoQuestionario = getSituacaoQuestionario( idCaso );
    let idsRespostasQuestionarios = idSituacaoQuestionario == "3" ?
                                    getRespostasQuestionario( idCaso ) : "";
        
    return {

      id: idCaso,

      referencia_familiar: caso[REFERENCIA_FAMILIAR],

      cpf_rf: caso[CPF_RF].padStart(11, "0"),      
 
      id_orgao_encaminhador: caso[ORGAO_ENCAMINHADOR],
            
      email_orgao_encaminhador: BUFFER_ORGAOS_ENCAMINHADORES[ parseInt(caso[ORGAO_ENCAMINHADOR]) - 1 ][EMAIL_INSTITUICAO],      

      data_encaminhamento: dataEncaminhamentoFormatada,

      id_complexidade: BUFFER_ORGAOS_ENCAMINHADORES[ parseInt(caso[ORGAO_ENCAMINHADOR]) - 1 ][ID_COMPLEXIDADE],

      servicos_referencia: servicosReferencia,

      servico_referencia_ativo: servicoReferenciaAtivo,

      id_complexidade_servico_referencia_ativo: BUFFER_ORGAOS_ENCAMINHADORES[ parseInt(servicoReferenciaAtivo.idServico) - 1 ][ID_COMPLEXIDADE],
            
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

      vistorias: vistoriasCaso,

      id_situacao_vistoria: idSituacaoVistoria,

      id_situacao_questionario: idSituacaoQuestionario,

      ids_respostas_questionarios: idsRespostasQuestionarios,

      ids_perfil: caso[PERFIL_COMPLETO],

      ids_perfil_genero: caso[PERFIL_GENERO],

      ids_perfil_orientacao_sexual: caso[PERFIL_ORIENTACAO_SEXUAL],
           
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

    if( caso.id_situacao_beneficio != "1") {
      caso.posicaoNaFila = posicao;
      ++posicao; 
    }
  });


  // Retorna a fila   
  return fila;

} // Fim da Função obterFila 




/**
 * Função que retorna a fila com os casos antigos registrados no sistema
 * 
 * @return Uma fila em que cada posição contém um objeto com os dados de um caso antigo
 */
function obterFilaCasosAntigos() {    


  // RETORNA NULL, SE TABELA DE CASOS ESTIVER VAZIA
  if( TAMANHO_FILA_CASOS_ANTIGOS < 1 ) return null;
    
  
  // Obtém os casos na fila
  let fila = BUFFER_FILA_CASOS_ANTIGOS.map( caso => {    

    let idCaso = caso[ID];

    let servicosReferencia = obterServicosReferenciaDoCaso( "old_" + String(idCaso) );
    let servicoReferenciaAtivo = obterServicoReferenciaAtivo( servicosReferencia );

    let vistoriasCaso = "";
    let idSituacaoVistoria = "";
    if( caso[CPF_RF] != "" ) {
      vistoriasCaso = pesquisarVistoriasPorCPF( caso[CPF_RF].padStart(11, "0") );
      idSituacaoVistoria = getSituacaoVistoria( vistoriasCaso );
    }
        
    return {

      id: "old_" + caso[ID],

      referencia_familiar: caso[REFERENCIA_FAMILIAR],

      cpf_rf: caso[CPF_RF].padStart(11, "0"),      
 
      id_orgao_encaminhador: caso[ORGAO_ENCAMINHADOR],
      
      email_orgao_encaminhador: BUFFER_ORGAOS_ENCAMINHADORES[ parseInt(caso[ORGAO_ENCAMINHADOR]) - 1 ][EMAIL_INSTITUICAO],      

      data_encaminhamento: "",

      id_complexidade: caso[ORGAO_ENCAMINHADOR] != "" ?
                       BUFFER_ORGAOS_ENCAMINHADORES[ parseInt(caso[ORGAO_ENCAMINHADOR]) - 1 ][ID_COMPLEXIDADE] :
                       "",          

      servicos_referencia: servicosReferencia,

      servico_referencia_ativo: servicoReferenciaAtivo,

      id_complexidade_servico_referencia_ativo: BUFFER_ORGAOS_ENCAMINHADORES[ parseInt(servicoReferenciaAtivo.idServico) - 1 ][ID_COMPLEXIDADE],

      idade_RF: calcularIdade( caso[DATA_NASCIMENTO_RF] ),

      data_nascimento_RF: caso[DATA_NASCIMENTO_RF],      

      id_situacao_beneficio: caso[SITUACAO_BENEFICIO], 
        
      data_ultima_evolucao: caso[DATA_ULTIMA_EVOLUCAO],
        
      data_limite: caso[DATA_LIMITE],

      justificativa_alteracao_data_limite: caso[JUSTIFICATIVA_ALTERACAO_DATA_LIMITE],      

      id_doc_pendente: caso[DOC_PENDENTE],

      vistorias: vistoriasCaso,

      id_situacao_vistoria: idSituacaoVistoria,

    };// Fim return       
  });

  
  // Retorna a fila   
  return fila;

} // Fim da Função obterFilaCasosAntigos 



/**
 * Função que retorna as duas filas com os casos antigos registrados no sistema
 * 
 * @return Um objeto com as duas filas, fila processo atual e fila processo antigo
 */
function obterFilas() {

  let f = obterFila();
  let f_old = obterFilaCasosAntigos();

  const filas = {
    fila: f,
    fila_old: f_old 
  };

  return JSON.stringify( filas );   

} // Fim da Função obterFilas 




/** 
 *  #################################################
 *  #####                                       ##### 
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 *  #####                                       ##### 
 *  #################################################
 */




/**
 * Função para testar a função obterFila
 */
function teste_obterFila() {

  const fila = obterFila();

  console.log(fila);    

} // Fim da Função teste_obterFila 




/**
 * Função para testar a função obterFilaCasosAntigos
 */
function teste_obterFilaCasosAntigos() {

  const fila = obterFilaCasosAntigos();

  const arrayCasos = JSON.parse( fila );
  
  for( let i=1; i<=10; ++i ) {
    console.log( arrayCasos[i] );
  }  

} // Fim da Função teste_obterFilaCasosAntigos 




/**
 * Função para testar a função teste_obterFilaCasosAntigos
 */
function teste_obterFilas() {

  const filas = obterFilas();
  const filas_obj = JSON.parse( filas );

  let f = filas_obj.fila;
  let f_old = filas_obj.fila_old;

  console.log( "\n\nFILA\n" );
  for( let i=1; i<=5; ++i ) {
    console.log( f[i] );
  }  

  console.log( "\n\nFILA ANTIGA\n" );
  for( let i=1; i<=5; ++i ) {
    console.log( f_old[i] );
  }    

} // Fim da Função teste_obterFilas 




/**
 * ##### FIM DO MÓDULO tabelasFila.gs #####
 */




