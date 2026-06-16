"use strict";

/**
 * Módulo:    evoluirMultiplosCasosBE.gs
 * Objetivo:  evoluir múltiplos casos da tabela CASOS
 */


/**
 * Função que evolui múltiplos casos da tabela CASOS, com
 * o mesmo tipo de evolução
 *  
 * @param {String} idsCasos: ids dos casos de serão evoluidos
 * @param {String} idEvolucao: id da evolucao
 * @param {Data} dataLimite: data limite, caso o tipo de evolução seja Convocado par Acesso
 */
function evoluirMultiplosCasosBE( idsCasos, idEvolucao, dataLimite ) {

  // Verifica se o usuário do app tem permissão para evoluir casos
  const usuarioLogado = JSON.parse( autenticarUsuario() );
  if( usuarioLogado.instituicao != "0" || usuarioLogado.tipo != "1" ) {
    throw( new Error( "Usuário sem permissão para evoluir casos" ) );
  }    


  // TENTA PEGAR O LOCK
  const lock = LockService.getScriptLock();

  try {

    lock.waitLock(10000);  
  
    // SE PEGAR O LOCK, PROSSEGUE COM A EVOLUÇÃO
    if( lock.hasLock() ) {

      // Evolui todos os casos do array idsCasos com a evolução idEvolucao
      idsCasos.forEach( idCaso => {

        console.log( "\nchamada - evoluirCasoBE" );
        evoluirCasoBE( idCaso, idEvolucao, dataLimite );
    
      }); // Fim do forEach


      // Aguarda sincronização de dados
      PLANILHA_FILA.waitForAllDataExecutionsCompletion(3);          
      SpreadsheetApp.flush();    
      
    } else {
  
      // SE NAO CONSEGUIR PEGAR O LOCK, LANCA UMA EXCESSAO
      throw( new Error( "evoluirMultiplosCasosBE - Nao foi possivel pegar o LOCK" ) );
    } 

  } catch( error ) {

    throw( "evoluirMultiplosCasosBE - " + error.message );

  } finally {

    // Always release the lock for other waiting instances
    lock.releaseLock(); 

  }    

  return true;

} // Fim da função evoluirMultiplosCasosBE



/**
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 */



/**
 * Função para testar a função principal evoluirMultiplosCasosBE
 */
function teste_evoluirMultiplosCasosBE() {

  let idsCasos = [ "204", "205", "206", "207", "208", "209", "210" ];
  let idEvolucao = "3";   
  let dataLimite = new Date( 2026, 11, 15 );

  try {
    evoluirMultiplosCasosBE( idsCasos, idEvolucao, dataLimite );
  } catch( error ) {
    console.log( "teste_evoluirMultiplosCasosBE - " + error.message );
  }

} // Fim da função teste_evoluirMultiplosCasosBE




/**
 * ##### FIM DO MÓDULO evoluirMultiplosCasosBE.gs #####
 */




