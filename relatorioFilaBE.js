/**
 * Planilha RELATORIO
 */
const PLANILHA_RELATORIO_ID     =  "16OlzDZ0aI4lYav7_NS54859a_NP4HiO1l7pI4SU0yF4";
const PLANILHA_RELATORIO        =  SpreadsheetApp.openById(PLANILHA_RELATORIO_ID);

const TABELA_RELATORIO          =  PLANILHA_RELATORIO.getSheetByName('RELATORIO');
const TABELA_QUANTITATIVOS      =  PLANILHA_RELATORIO.getSheetByName('QUANTITATIVOS');
let BUFFER_RELATORIO            =  TABELA_RELATORIO.getDataRange().getDisplayValues().splice(1);
let NUM_RELATORIOS              =  BUFFER_RELATORIO.length;
const NUM_COLUNAS_TABELA_RELATORIO  =  14;



const RELATORIO_COLUNA_ID            =  0;
const RELATORIO_COLUNA_POSICAO_FILA  =  1;
const RELATORIO_COLUNA_NOME_RF       =  2;
const RELATORIO_COLUNA_CPF_RF        =  3;

const RELATORIO_COLUNA_ORGAO_ENCAMINHADOR  =  4;
const RELATORIO_COLUNA_SITUACAO_BENEFICIO  =  5;

const RELATORIO_COLUNA_DATA_ULTIMA_EVOLUCAO  =  6;
const RELATORIO_COLUNA_DATA_LIMITE           =  7;

const RELATORIO_COLUNA_SITUACAO_VISTORIA        =  8;
const RELATORIO_COLUNA_SITUACAO_QUESTIONARIO    =  9;
const RELATORIO_COLUNA_SITUACAO_ACOMPANHAMENTO  = 10;

const RELATORIO_COLUNA_JUSTIFICATIVA_1  = 11;
const RELATORIO_COLUNA_JUSTIFICATIVA_2  = 12;

const RELATORIO_COLUNA_OBSERVACAO  = 13;



/**
 * Função que limpa a planilha RELATORIO
 */
function limparRelatorio() {

  // Lock
  let lock;

  try {

    // TENTA PEGAR O LOCK
    lock = LockService.getScriptLock();
    lock.waitLock(10000);  

    // SE PEGAR O LOCK, PROSSEGUE A EXCLUSÃO DOS DADOS DO RELATÓRIO
    if( lock.hasLock() ) {

      console.log( "limparRelatorio - Início" );
  
      // Caso nulo
      let casoNulo = new Array(NUM_COLUNAS_TABELA_RELATORIO).fill("");
      let bufferCasosNulos = [];
    
      // Limpa a fila
      let range;
      for( let linha=2; linha<=TAMANHO_FILA+1; ++linha ) {
        bufferCasosNulos.push(casoNulo);
      }    
      
      
      // Grava o buffer de casos nulos na planilha RELATORIO
      TABELA_RELATORIO.getRange( 2, 1, bufferCasosNulos.length, NUM_COLUNAS_TABELA_RELATORIO ).setNumberFormat("@").setValues( bufferCasosNulos );  
      PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
      SpreadsheetApp.flush();  
  
      console.log( "limparRelatorio - Fim" );

    } else {

      // SE NAO CONSEGUIR PEGAR O LOCK, LANCA UMA EXCESSAO
      throw( new Error( "Nao foi possivel pegar o LOCK" ) );
    }

  } catch( error ) {

    throw( "limparRelatorio: " + error.message );

  } finally {

    // SOLTA O LOCK
    lock.releaseLock();    
  }    

} // Fim da função limparRelatorio



/**
 * Função que gera os relatórios de todos os casos da fila, 
 * gravando-os na planilha RELATORIO
 */
function gerarRelatorio() {

  // Lock
  let lock;

  try {

    // TENTA PEGAR O LOCK
    lock = LockService.getScriptLock();
    lock.waitLock(10000);  

    // SE PEGAR O LOCK, PROSSEGUE COM A GERAÇÃO DOS DADOS DO RELATÓRIO
    if( lock.hasLock() ) {
      
      console.log( "gerarRelatorio - Início" );  

      let bufferRelatorioCaso = [];      
      const bufferRelatorios  = [];      

      let nomeSituacaoVistoria
      let idsRespostasQuestionarios;

      let jsonFilaOrdenada = obterFila();
      let filaOrdenada = JSON.parse( jsonFilaOrdenada );
      console.log( "filaOrdenada" );

      let posicaoFila = 0;

      // Percorre todos os casos da fila, gerando o relatório de cada caso
      filaOrdenada.forEach( caso => {
        
        bufferRelatorioCaso = new Array(NUM_COLUNAS_TABELA_RELATORIO).fill("");

        // Id do caso no sistema de filas
        bufferRelatorioCaso[RELATORIO_COLUNA_ID] = caso.id;

        // Posição do caso na fila
        ++posicaoFila;
        bufferRelatorioCaso[RELATORIO_COLUNA_POSICAO_FILA] = posicaoFila;

        // Dados de Identificação        
        bufferRelatorioCaso[RELATORIO_COLUNA_NOME_RF] = caso.referencia_familiar;
        bufferRelatorioCaso[RELATORIO_COLUNA_CPF_RF] = caso.cpf_rf.padStart(11, "0");

        // Órgão encaminhador
        bufferRelatorioCaso[RELATORIO_COLUNA_ORGAO_ENCAMINHADOR] = idToNome( caso.id_orgao_encaminhador, "ORGAOS_ENCAMINHADORES" );                                         

        // Situação benefício
        bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_BENEFICIO] = caso.id_situacao_beneficio ? 
                                                                   idToNome( caso.id_situacao_beneficio, "SITUACOES_BENEFICIO" ) :
                                                                   "Sem Informação";

        // Data última evolução                         
        bufferRelatorioCaso[RELATORIO_COLUNA_DATA_ULTIMA_EVOLUCAO] = caso.data_ultima_evolucao != "" ?
                                                                     caso.data_ultima_evolucao  :
                                                                     "Sem Informação";

        // data limite                                 
        bufferRelatorioCaso[RELATORIO_COLUNA_DATA_LIMITE] = caso.data_limite != "" ?
                                                            caso.data_limite  :
                                                            "Sem Informação";        

        // Situação vistoria
        nomeSituacaoVistoria = idToNome( caso.id_situacao_vistoria, "SITUACOES_VISTORIA" )
        bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_VISTORIA] = nomeSituacaoVistoria != "" ?
                                                                  nomeSituacaoVistoria :
                                                                  "Sem Informação";
         
        // Situação questionário
        nomeSituacaoQuestionario = idToNome( caso.id_situacao_questionario, "SITUACOES_QUESTIONARIO" );
        bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_QUESTIONARIO] = nomeSituacaoQuestionario != "" ?
                                                                      nomeSituacaoQuestionario :
                                                                      "Sem Informação";

        // Situação acompanhamento
        idsRespostasQuestionarios = caso.ids_respostas_questionarios;
        
        idSituacaoAcompanhamento = idsRespostasQuestionarios != "" ?
                                   idsRespostasQuestionarios.q1 : 
                                   "";

        // Switch - ACOMPANHAMENTO DO CASO - (NÃO / SIM)                                   
        switch( idSituacaoAcompanhamento ) {
           
          case "1":  // CASO NÃO ACOMPANHADO
                     bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_ACOMPANHAMENTO] = "NÃO acompanhado pelo serviço"; 

                     // Justificativa de não acompanhamento
                     bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_1] = idsRespostasQuestionarios.q5 != "" ?
                                                                             idToNome( idsRespostasQuestionarios.q5, "ACOMPANHAMENTO_NAO" ) :
                                                                             "Sem Informação"; 
                     // Sem justificativa 2                                                                             
                     bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = "Sem Informação";                                                                              

                     break;


          case "2":  // CASO ACOMPANHADO
                     bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_ACOMPANHAMENTO] = "Acompanhado pelo serviço"; 

                     // Switch - ETAPAS DE ACESSO
                     switch(idsRespostasQuestionarios.q2) {

                       case "5":  // O beneficiário está impossibilitado de acessar temporariamente
                                  bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_1] = idsRespostasQuestionarios.q2 != "" ? 
                                                                                          idToNome( idsRespostasQuestionarios.q2, "ETAPA_ACESSO" ) :
                                                                                          "Sem Informação" ;
                                  // Justificativa - não acesso temporário
                                  bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = idsRespostasQuestionarios.q3 != "" ? 
                                                                                          idToNome( idsRespostasQuestionarios.q3, "IMPOSSIBILIDADE_TEMPORARIA" ) :
                                                                                          "Sem Informação" ;

                                  break;

                       case "6":  // O beneficiário não acessará o benefício de forma definitiva
                                  bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_1] = idsRespostasQuestionarios.q2 != "" ? 
                                                                                          idToNome( idsRespostasQuestionarios.q2, "ETAPA_ACESSO" ) :
                                                                                          "Sem Informação" ;
                                  // Justificativa - não acesso definitivo
                                  bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = idsRespostasQuestionarios.q4 != "" ? 
                                                                                          idToNome( idsRespostasQuestionarios.q4, "NAO_ACESSO_DEFINITIVO" ) :
                                                                                          "Sem Informação" ;

                                  break;                                 

                       default:   // Opções de 1 à 4 - sem justificativas
                                  bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_1] = idsRespostasQuestionarios.q2 != "" ? 
                                                                                          idToNome( idsRespostasQuestionarios.q2, "ETAPA_ACESSO" ) :
                                                                                          "Sem Informação" ;                                  
                                  bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = "Sem Informação";                                                                                 

                                  break;                                 

                     }  

                     break;

          default:   bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_ACOMPANHAMENTO] = "Sem Informação"; 
                     bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_1] = "Sem Informação";                                                                                        
                     bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = "Sem Informação";                                                                                        
                     break;
        }                                
        
        // Obsercvações caso
        bufferRelatorioCaso[RELATORIO_COLUNA_OBSERVACAO] = idsRespostasQuestionarios.observacoes; 

        // Acrescenta 1 caso ao buffer relatórios
        bufferRelatorios.push( bufferRelatorioCaso );        

      }); // Fim do for que percorre todos os casos da fila


      // Grava o buffer do relatório na planilha RELATORIO
      TABELA_RELATORIO.getRange( 2, 1, bufferRelatorios.length, NUM_COLUNAS_TABELA_RELATORIO ).setNumberFormat("@").setValues( bufferRelatorios );
      PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
      SpreadsheetApp.flush();  
  
      console.log( "gerarRelatorio - Fim" );

    } else {

      // SE NAO CONSEGUIR PEGAR O LOCK, LANCA UMA EXCESSAO
      throw( new Error( "Nao foi possivel pegar o LOCK" ) );
    }

  } catch( error ) {

    throw( "gerarRelatorio: " + error.message );

  } finally {

    // SOLTA O LOCK
    lock.releaseLock();    
  }   

} // Fim da função gerarRelatorio



/**
 * Função que gera os relatórios de todos os casos da fila, 
 * e o retorna em excel
 */
function getRelatorioExel() {


  // Verifica se o usuário do app tem permissão para obter o relatório
  let usuarioLogado;
  try {
    usuarioLogado = JSON.parse( autenticarUsuario() );
  } catch( error ) {
    throw( "getRelatorioExel: 1 " + error.message );
  }    

  // Gera e retorna o relatório
  try {

    console.log( "getRelatorioExel - Início" );
    limparRelatorio();  
    gerarRelatorio();    
    console.log( "getRelatorioExel - Fim" );

    return `https://docs.google.com/spreadsheets/d/${PLANILHA_RELATORIO_ID}/export?format=xlsx`;

  } catch( error ) {
    throw( "getRelatorioExel: " + error.message );
  }    

} // Fim da função getRelatorioExel



