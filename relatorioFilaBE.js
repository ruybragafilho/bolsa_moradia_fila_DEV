/**
 * Planilha RELATORIO
 */
const PLANILHA_RELATORIO_ID     =  PropertiesService.getScriptProperties().getProperty('PLANILHA_RELATORIO_ID');
const PLANILHA_RELATORIO        =  SpreadsheetApp.openById(PLANILHA_RELATORIO_ID);

const TABELA_RELATORIO          =  PLANILHA_RELATORIO.getSheetByName('HABILITADOS_2026');
const TABELA_RELATORIO_ANTIGO   =  PLANILHA_RELATORIO.getSheetByName('HABILITADOS_ANTES_DE_2026');

let BUFFER_RELATORIO            =  TABELA_RELATORIO.getDataRange().getDisplayValues().splice(1);
let BUFFER_RELATORIO_ANTIGO     =  TABELA_RELATORIO_ANTIGO.getDataRange().getDisplayValues().splice(1);

let NUM_RELATORIOS              =  BUFFER_RELATORIO.length;
let NUM_RELATORIOS_ANTIGO       =  BUFFER_RELATORIO_ANTIGO.length;

const NUM_COLUNAS_TABELA_RELATORIO         =  17;
const NUM_COLUNAS_TABELA_RELATORIO_ANTIGO  =  17;



const RELATORIO_COLUNA_ID            =  0;
const RELATORIO_COLUNA_POSICAO_FILA  =  1;
const RELATORIO_COLUNA_NOME_RF       =  2;
const RELATORIO_COLUNA_CPF_RF        =  3;

const RELATORIO_COLUNA_ORGAO_ENCAMINHADOR               =  4;
const RELATORIO_COLUNA_COMPLEXIDADE_ORGAO_ENCAMINHADOR  =  5;

const RELATORIO_COLUNA_SERVICO_REFERENCIA               =  6;
const RELATORIO_COLUNA_COMPLEXIDADE_SERVICO_REFERENCIA  =  7;

const RELATORIO_COLUNA_SITUACAO_BENEFICIO    =  8;
const RELATORIO_COLUNA_DATA_ULTIMA_EVOLUCAO  =  9;
const RELATORIO_COLUNA_DATA_LIMITE           =  10;

const RELATORIO_COLUNA_SITUACAO_VISTORIA        = 11;
const RELATORIO_COLUNA_SITUACAO_QUESTIONARIO    = 12;
const RELATORIO_COLUNA_SITUACAO_ACOMPANHAMENTO  = 13;

const RELATORIO_COLUNA_JUSTIFICATIVA_1  = 14;
const RELATORIO_COLUNA_JUSTIFICATIVA_2  = 15;

const RELATORIO_COLUNA_OBSERVACAO  = 16;



/**
 * Função que limpa a planilha RELATORIO
 */
function limparRelatorio() {

  console.log( "limparRelatorio - Início" );

  // Lock
  let lock;

  try {

    // TENTA PEGAR O LOCK
    lock = LockService.getScriptLock();
    lock.waitLock(10000);  

    // SE PEGAR O LOCK, PROSSEGUE A EXCLUSÃO DOS DADOS DO RELATÓRIO
    if( lock.hasLock() ) {
  
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

  console.log( "limparRelatorio - Fim" );  

} // Fim da função limparRelatorio



/**
 * Função que gera os relatórios de todos os casos da fila, 
 * gravando-os na planilha RELATORIO
 */
function gerarRelatorio() {

  console.log( "gerarRelatorio - Início" );  

  // Lock
  let lock;

  try {

    // TENTA PEGAR O LOCK
    lock = LockService.getScriptLock();
    lock.waitLock(10000);  

    // SE PEGAR O LOCK, PROSSEGUE COM A GERAÇÃO DOS DADOS DO RELATÓRIO
    if( lock.hasLock() ) {

      let bufferRelatorioCaso = [];      
      const bufferRelatorios  = [];      

      let nomeSituacaoVistoria
      let idsRespostasQuestionarios;

      let filaOrdenada = obterFila();      
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

        // Complexidade órgão encaminhador
        bufferRelatorioCaso[RELATORIO_COLUNA_COMPLEXIDADE_ORGAO_ENCAMINHADOR] = caso.id_complexidade ? 
                                                                                idToNome( caso.id_complexidade, "COMPLEXIDADES" ) :
                                                                                "Sem informação";                

        // Serviço referência ativo
        bufferRelatorioCaso[RELATORIO_COLUNA_SERVICO_REFERENCIA] = idToNome( caso.servico_referencia_ativo.idServico, "ORGAOS_ENCAMINHADORES" );                                         

        // Complexidade serviço referência ativo
        bufferRelatorioCaso[RELATORIO_COLUNA_COMPLEXIDADE_SERVICO_REFERENCIA] = caso.id_complexidade_servico_referencia_ativo ? 
                                                                                idToNome( caso.id_complexidade_servico_referencia_ativo, "COMPLEXIDADES" ) :
                                                                                "Sem informação";                                                                             

        // Situação benefício
        bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_BENEFICIO] = caso.id_situacao_beneficio ? 
                                                                   idToNome( caso.id_situacao_beneficio, "SITUACOES_BENEFICIO" ) :
                                                                   "Sem informação";
        // Data última evolução                         
        bufferRelatorioCaso[RELATORIO_COLUNA_DATA_ULTIMA_EVOLUCAO] = caso.data_ultima_evolucao != "" ?
                                                                     caso.data_ultima_evolucao  :
                                                                     "Sem informação";

        // data limite                                 
        bufferRelatorioCaso[RELATORIO_COLUNA_DATA_LIMITE] = caso.data_limite != "" ?
                                                            caso.data_limite  :
                                                            "Sem informação";        

        // Situação vistoria
        nomeSituacaoVistoria = idToNome( caso.id_situacao_vistoria, "SITUACOES_VISTORIA" )
        bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_VISTORIA] = nomeSituacaoVistoria != "" ?
                                                                  nomeSituacaoVistoria :
                                                                  "Sem vistoria solicitada";
         
        // Situação questionário
        nomeSituacaoQuestionario = idToNome( caso.id_situacao_questionario, "SITUACOES_QUESTIONARIO" );
        bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_QUESTIONARIO] = nomeSituacaoQuestionario != "" ?
                                                                      nomeSituacaoQuestionario :
                                                                      "Sem Informação";

        // Se caso NÃO tem questionário a responder                                                                      
        if( caso.id_situacao_questionario == "1" ) {

          bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_ACOMPANHAMENTO] = "Não se aplica"; 
          bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_1] = "Não se aplica";                                                                                        
          bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = "Não se aplica";       
          bufferRelatorioCaso[RELATORIO_COLUNA_OBSERVACAO] = "Não se aplica";                                                                                            


        // Se caso tem questionário a responder            
        } else {

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
                                                                               "Não se aplica"; 
                       // Sem justificativa 2                                                                             
                       bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = "Não se aplica";                                                                              
  
                       break;
  
  
            case "2":  // CASO ACOMPANHADO
                       bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_ACOMPANHAMENTO] = "Acompanhado pelo serviço"; 
  
                       // Switch - ETAPAS DE ACESSO
                       switch(idsRespostasQuestionarios.q2) {
  
                         case "5":  // O beneficiário está impossibilitado de acessar temporariamente
                                    bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_1] = idsRespostasQuestionarios.q2 != "" ? 
                                                                                            idToNome( idsRespostasQuestionarios.q2, "ETAPA_ACESSO" ) :
                                                                                            "Não se aplica" ;
                                    // Justificativa - não acesso temporário
                                    bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = idsRespostasQuestionarios.q3 != "" ? 
                                                                                            idToNome( idsRespostasQuestionarios.q3, "IMPOSSIBILIDADE_TEMPORARIA" ) :
                                                                                            "Não se aplica" ;
  
                                    break;
  
                         case "6":  // O beneficiário não acessará o benefício de forma definitiva
                                    bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_1] = idsRespostasQuestionarios.q2 != "" ? 
                                                                                            idToNome( idsRespostasQuestionarios.q2, "ETAPA_ACESSO" ) :
                                                                                            "Não se aplica" ;
                                    // Justificativa - não acesso definitivo
                                    bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = idsRespostasQuestionarios.q4 != "" ? 
                                                                                            idToNome( idsRespostasQuestionarios.q4, "NAO_ACESSO_DEFINITIVO" ) :
                                                                                            "Não se aplica" ;
  
                                    break;                                 
  
                         default:   // Opções de 1 à 4 - sem justificativas
                                    bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_1] = idsRespostasQuestionarios.q2 != "" ? 
                                                                                            idToNome( idsRespostasQuestionarios.q2, "ETAPA_ACESSO" ) :
                                                                                            "Não se aplica" ;  
  
                                    bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = "Não se aplica";                                                                                 
  
                                    break;                                 
  
                       }  
  
                       break;
  
            default:   bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_ACOMPANHAMENTO] = "Sem Informação"; 
                       bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_1] = "Sem Informação";                                                                                        
                       bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = "Sem Informação";                                                                                        
                       break;

          } // Fim switch-case idSituacaoAcompanhamento                               
          
          // Obsercvações caso
          bufferRelatorioCaso[RELATORIO_COLUNA_OBSERVACAO] = idsRespostasQuestionarios.observacoes;           

        } // Fim if caso.id_situacao_questionario


        // Acrescenta 1 caso ao buffer relatórios
        bufferRelatorios.push( bufferRelatorioCaso );        

      }); // Fim do for que percorre todos os casos da fila


      // Grava o buffer do relatório na planilha RELATORIO
      TABELA_RELATORIO.getRange( 2, 1, bufferRelatorios.length, NUM_COLUNAS_TABELA_RELATORIO ).setNumberFormat("@").setValues( bufferRelatorios );
      PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
      SpreadsheetApp.flush();  

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

  console.log( "gerarRelatorio - Fim" );  

} // Fim da função gerarRelatorio



/**
 * Função que limpa a planilha RELATORIO ANTIGO
 */
function limparRelatorioAntigo() {

  console.log( "limparRelatorioAntigo - Início" );

  // Lock
  let lock;

  try {

    // TENTA PEGAR O LOCK
    lock = LockService.getScriptLock();
    lock.waitLock(10000);  

    // SE PEGAR O LOCK, PROSSEGUE A EXCLUSÃO DOS DADOS DO RELATÓRIO
    if( lock.hasLock() ) {
  
      // Caso nulo
      let casoNulo = new Array(NUM_COLUNAS_TABELA_RELATORIO_ANTIGO).fill("");
      let bufferCasosNulos = [];
    
      // Limpa a fila
      let range;
      for( let linha=2; linha<=TAMANHO_FILA_CASOS_ANTIGOS+1; ++linha ) {
        bufferCasosNulos.push(casoNulo);
      }    
      
      
      // Grava o buffer de casos nulos na planilha RELATORIO
      TABELA_RELATORIO_ANTIGO.getRange( 2, 1, bufferCasosNulos.length, NUM_COLUNAS_TABELA_RELATORIO_ANTIGO ).setNumberFormat("@").setValues( bufferCasosNulos );  
      PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
      SpreadsheetApp.flush();  
  
    } else {

      // SE NAO CONSEGUIR PEGAR O LOCK, LANCA UMA EXCESSAO
      throw( new Error( "Nao foi possivel pegar o LOCK" ) );
    }

  } catch( error ) {

    throw( "limparRelatorioAntigo: " + error.message );

  } finally {

    // SOLTA O LOCK
    lock.releaseLock();    
  }    

  console.log( "limparRelatorioAntigo - Fim" );  

} // Fim da função limparRelatorioAntigo



/**
 * Função que gera os relatórios de todos os casos da fila, 
 * gravando-os na planilha RELATORIO ANTIGO
 */
function gerarRelatorioAntigo() {

  console.log( "gerarRelatorioAntigo - Início" );  

  // Lock
  let lock;

  try {

    // TENTA PEGAR O LOCK
    lock = LockService.getScriptLock();
    lock.waitLock(10000);  

    // SE PEGAR O LOCK, PROSSEGUE COM A GERAÇÃO DOS DADOS DO RELATÓRIO
    if( lock.hasLock() ) {      

      let bufferRelatorioCaso = [];      
      const bufferRelatorios  = [];      

      let nomeSituacaoVistoria;

      let fila = obterFilaCasosAntigos();      
      console.log( "fila" );

      let posicaoFila = 0;

      // Percorre todos os casos da fila, gerando o relatório de cada caso
      fila.forEach( caso => {
        
        bufferRelatorioCaso = new Array(NUM_COLUNAS_TABELA_RELATORIO_ANTIGO).fill("");

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

        // Complexidade órgão encaminhador
        bufferRelatorioCaso[RELATORIO_COLUNA_COMPLEXIDADE_ORGAO_ENCAMINHADOR] = caso.id_complexidade ? 
                                                                                idToNome( caso.id_complexidade, "COMPLEXIDADES" ) :
                                                                                "Sem informação";                

        // Serviço referência ativo
        bufferRelatorioCaso[RELATORIO_COLUNA_SERVICO_REFERENCIA] = idToNome( caso.servico_referencia_ativo.idServico, "ORGAOS_ENCAMINHADORES" );                                         

        // Complexidade serviço referência ativo
        bufferRelatorioCaso[RELATORIO_COLUNA_COMPLEXIDADE_SERVICO_REFERENCIA] = caso.id_complexidade_servico_referencia_ativo ? 
                                                                                idToNome( caso.id_complexidade_servico_referencia_ativo, "COMPLEXIDADES" ) :
                                                                                "Sem informação";                                                                             

        // Situação benefício
        bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_BENEFICIO] = caso.id_situacao_beneficio ? 
                                                                   idToNome( caso.id_situacao_beneficio, "SITUACOES_BENEFICIO" ) :
                                                                   "Sem informação";
        // Data última evolução                         
        bufferRelatorioCaso[RELATORIO_COLUNA_DATA_ULTIMA_EVOLUCAO] = caso.data_ultima_evolucao != "" ?
                                                                     caso.data_ultima_evolucao  :
                                                                     "Sem informação";

        // data limite                                 
        bufferRelatorioCaso[RELATORIO_COLUNA_DATA_LIMITE] = caso.data_limite != "" ?
                                                            caso.data_limite  :
                                                            "Sem informação";        

        // Situação vistoria
        nomeSituacaoVistoria = idToNome( caso.id_situacao_vistoria, "SITUACOES_VISTORIA" )
        bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_VISTORIA] = nomeSituacaoVistoria != "" ?
                                                                  nomeSituacaoVistoria :
                                                                  "Sem vistoria solicitada";
         
        
        // Questionário
        bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_QUESTIONARIO] = "Não se aplica";
        bufferRelatorioCaso[RELATORIO_COLUNA_SITUACAO_ACOMPANHAMENTO] = "Não se aplica"; 
        bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_1] = "Não se aplica";                                                                                        
        bufferRelatorioCaso[RELATORIO_COLUNA_JUSTIFICATIVA_2] = "Não se aplica";
        bufferRelatorioCaso[RELATORIO_COLUNA_OBSERVACAO] = "Não se aplica"; 

        // Acrescenta 1 caso ao buffer relatórios
        bufferRelatorios.push( bufferRelatorioCaso );        

      }); // Fim do for que percorre todos os casos da fila


      // Grava o buffer do relatório na planilha RELATORIO
      TABELA_RELATORIO_ANTIGO.getRange( 2, 1, bufferRelatorios.length, NUM_COLUNAS_TABELA_RELATORIO_ANTIGO ).setNumberFormat("@").setValues( bufferRelatorios );
      PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
      SpreadsheetApp.flush();  

    } else {

      // SE NAO CONSEGUIR PEGAR O LOCK, LANCA UMA EXCESSAO
      throw( new Error( "Nao foi possivel pegar o LOCK" ) );
    }

  } catch( error ) {

    throw( "gerarRelatorioAntigo: " + error.message );

  } finally {

    // SOLTA O LOCK
    lock.releaseLock();    
  }   

  console.log( "gerarRelatorioAntigo - Fim" );  

} // Fim da função gerarRelatorioAntigo



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
    limparRelatorioAntigo();  
    gerarRelatorioAntigo();        
    console.log( "getRelatorioExel - Fim" );

    return `https://docs.google.com/spreadsheets/d/${PLANILHA_RELATORIO_ID}/export?format=xlsx`;

  } catch( error ) {
    throw( "getRelatorioExel: " + error.message );
  }    

} // Fim da função getRelatorioExel



