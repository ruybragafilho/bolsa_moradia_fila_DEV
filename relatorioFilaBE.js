/**
 * Planilha RELATORIO
 */
const PLANILHA_RELATORIO_ID     =  "16OlzDZ0aI4lYav7_NS54859a_NP4HiO1l7pI4SU0yF4";
const PLANILHA_RELATORIO        =  SpreadsheetApp.openById(PLANILHA_RELATORIO_ID);

const TABELA_RELATORIO          =  PLANILHA_RELATORIO.getSheetByName('RELATORIO');
const TABELA_QUANTITATIVOS      =  PLANILHA_RELATORIO.getSheetByName('QUANTITATIVOS');
let BUFFER_RELATORIO            =  TABELA_RELATORIO.getDataRange().getDisplayValues().splice(1);
let NUM_RELATORIOS              =  BUFFER_RELATORIO.length;
const NUM_COLUNAS_TABELA_RELATORIO  =  8;



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
      for( let linha=2; linha<=NUM_RELATORIOS+1; ++linha ) {
        bufferCasosNulos.push(casoNulo);
      }    
      
      
      // Grava o buffer de casos nulos na planilha RELATORIO
      TABELA_RELATORIO.getRange( 2, 1, bufferCasosNulos.length, NUM_COLUNAS_TABELA_RELATORIO ).setValues( bufferCasosNulos );  
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
      
      console.log( "gerarRelatorioFila - Início" );
  
      let caso;

      const relatorio = [];
      let vistoriasCaso;
      let idSituacaoVistoria;
      let idSituacaoQuestionario;
      let idsRespostasQuestionarios;
  
      // Percorre todos os casos da fila, gerando o relatório de cada caso
      for( let idCaso=1; idCaso<=TAMANHO_FILA; ++idCaso ) {
      
        caso = BUFFER_FILA[idCaso - 1];
    
        bufferRelatorioCaso = new Array(NUM_COLUNAS_TABELA_RELATORIO).fill("");

        bufferRelatorioCaso[0] = caso[ ID ];
        bufferRelatorioCaso[1] = caso[ REFERENCIA_FAMILIAR ];
        bufferRelatorioCaso[2] = caso[ CPF_RF ];
        bufferRelatorioCaso[3] = idToNome( caso[SITUACAO_BENEFICIO], "SITUACOES_BENEFICIO" );

        // Situação Vistoria
        vistoriasCaso = pesquisarVistoriasPorCPF( caso[CPF_RF].padStart(11, "0") );
        idSituacaoVistoria = getSituacaoVistoria( vistoriasCaso );
        bufferRelatorioCaso[4] = idToNome( idSituacaoVistoria, "SITUACOES_VISTORIA" );

        // Situação questionário
        idSituacaoQuestionario = getSituacaoQuestionario( caso[ID] );
        bufferRelatorioCaso[5] = idToNome( idSituacaoQuestionario, "SITUACOES_QUESTIONARIO" );

        // Situação acompanhamento
        idsRespostasQuestionarios = idSituacaoQuestionario == "3" ?
                                    getRespostasQuestionario( caso[ID] ) : "";
        idSituacaoAcompanhamento = idsRespostasQuestionarios != "" ?
                                   idsRespostasQuestionarios.q1 : "";
        switch( idSituacaoAcompanhamento ) {
          case "1":  bufferRelatorioCaso[6] = "NÃO acompanhado pelo serviço"; break;
          case "2":  bufferRelatorioCaso[6] = "Acompanhado pelo serviço"; break;
          default:   bufferRelatorioCaso[6] = "Sem Informação"; break;
        }                                    

        // Órgão encaminhador
        bufferRelatorioCaso[7] = idToNome( caso[ORGAO_ENCAMINHADOR], "ORGAOS_ENCAMINHADORES" );
         
        relatorio.push( bufferRelatorioCaso );        
  
      } // Fim do for que percorre todos os casos da fila
  

      // Grava o buffer do relatório na planilha RELATORIO
      TABELA_RELATORIO.getRange( 2, 1, relatorio.length, NUM_COLUNAS_TABELA_RELATORIO ).setValues( relatorio );
      PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
      SpreadsheetApp.flush();  
  
      console.log( "gerarRelatorioFila - Fim" );

    } else {

      // SE NAO CONSEGUIR PEGAR O LOCK, LANCA UMA EXCESSAO
      throw( new Error( "Nao foi possivel pegar o LOCK" ) );
    }

  } catch( error ) {

    throw( "gerarRelatorioFila: " + error.message );

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



