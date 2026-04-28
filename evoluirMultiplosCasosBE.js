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

  // Evolui todos os casos do array idsCasos com a evolução idEvolucao
  idsCasos.forEach( idCaso => {

    try {
      evoluirCasoBE( idCaso, idEvolucao, dataLimite );
    } catch( error ) {
      throw( "evoluirMultiplosCasosBE - " + error.message );
    } 

  }); // Fim do forEach

} // Fim da função evoluirMultiplosCasosBE



/**
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 */



/**
 * Função para testar a função principal evoluirMultiplosCasosBE
 */
function teste_evoluirMultiplosCasosBE() {

  let idsCasos = [ "3", "5", "9", "12" ];
  let idEvolucao = "5";   
  let dataLimite = "";

  try {
    evoluirMultiplosCasosBE( idsCasos, idEvolucao, dataLimite );
  } catch( error ) {
    console.log( "teste_evoluirMultiplosCasosBE - " + error.message );
  }

} // Fim da função teste_evoluirMultiplosCasosBE




/**
 * ##### FIM DO MÓDULO evoluirMultiplosCasosBE.gs #####
 */




