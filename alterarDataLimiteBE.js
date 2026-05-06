"use strict";

/**
 * Módulo:    alterarDataLimiteBE.gs
 * Objetivo:  Alterar a data limite do caso
 */


/**
 * Função que altera a data limite do caso
 *  
 * @param {String} idCaso: id do caso de terá a data limite alterada
 * @param {Data} dataLimite: data limite do caso
 * @param {String} justificativa: Justificativa para alteração da data limite do caso
 */
function alterarDataLimiteBE( idCaso, dataLimite, justificativa ) {

  // Converte o id para Integer
  const id = parseInt(idCaso);  

  // Se id inválido, retorna uma exceção
  if( id < 1  ||  id > TAMANHO_FILA ) {
    throw( new Error( "alterarDataLimiteBE - ID Inválido" ) );
  }  

  // Verifica se o usuário do app tem permissão para registrar Documentação Pendente
  const usuarioLogado = JSON.parse( autenticarUsuario() );
  if( usuarioLogado.instituicao != "0" || usuarioLogado.tipo != "1" ) {
    throw( new Error( "Usuário sem permissão para registrar Documentação Pendente" ) );
  }  

  try {

    // TENTA PEGAR O LOCK
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);  
  
    // SE PEGAR O LOCK, PROSSEGUE COM A ALTERAÇÃO DA DATA LIMITE
    if( lock.hasLock() ) {

      // Grava a nova data limite 
      const campo_data_limite = TABELA_FILA.getRange( id+1, DATA_LIMITE+1 );
      //campo_data_limite.setValue( dataLimite );          
      let auxDataLimite = new Date( dataLimite );  
      auxDataLimite.setDate( auxDataLimite.getDate() + 1 );      
      campo_data_limite.setValue( new Date(auxDataLimite).toLocaleString("pt-BR", {dateStyle: "short"}) );          

      
      // Grava a justificativa da alteração da data limite
      let justificativaFormatada = justificativa.length > 0 ? 
                                   justificativa.trim().toUpperCase().substring(0, 512) :
                                   "SEM JUSTIFICATIVA";
      const campo_justificativa = TABELA_FILA.getRange( id+1, JUSTIFICATIVA_ALTERACAO_DATA_LIMITE+1 );
      campo_justificativa.setValue( justificativaFormatada );            

      
      // SOLTA O LOCK
      lock.releaseLock();
  
      return true;
  
    } else {
  
      // SE NAO CONSEGUIR PEGAR O LOCK, LANCA UMA EXCESSAO
      throw( new Error( "alterarDataLimiteBE - Nao foi possivel pegar o LOCK" ) );
    }

  } catch( error ) {
    throw( "alterarDataLimiteBE - " + error.message );
  }

} // Fim da função alterarDataLimiteBE



/**
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 */



/**
 * Função para testar a função principal evoluirCasoBE
 */
function teste_alterarDataLimiteBE() {

  let idCaso = "13";  
  let dataLimite = "02/05/2026";
  let justificativa = "Alteração da Data Limite";

  try {
    alterarDataLimiteBE( idCaso, dataLimite, justificativa );
  } catch( error ) {
    console.log( "teste_evoluirCasoBE - " + error.message );
  }

} // Fim da função teste_alterarDataLimiteBE




/**
 * ##### FIM DO MÓDULO alterarDataLimiteBE.gs #####
 */




