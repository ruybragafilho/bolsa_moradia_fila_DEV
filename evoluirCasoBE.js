"use strict";

/**
 * Módulo:    evoluirCasoBE.gs
 * Objetivo:  evoluir um caso da tabela CASOS
 */


/**
 * Função que evolui um caso da tabela CASOS
 *  
 * @param {String} idCaso: id do caso de será evoluido
 * @param {String} idEvolucao: id da evolucao
 * @param {Data} dataLimite: data limite, caso o tipo de evolução seja Convocado par Acesso* 
 */
function evoluirCasoBE( idCaso, idEvolucao, dataLimite ) {

  // Converte o id para Integer
  const id = parseInt(idCaso);  

  // Se id inválido, retorna uma exceção
  if( id < 1  ||  id > TAMANHO_FILA ) {
    throw( new Error( "evoluirCasoBE - ID Inválido" ) );
  }  

  try {
           
    // Gera, formata e grava a data da evolução do caso
    let dataEvolucao = new Date().toLocaleString("pt-BR", {dateStyle: "short"});    
    const campo_data = TABELA_FILA.getRange( id+1, DATA_ULTIMA_EVOLUCAO+1 );
    campo_data.setValue( dataEvolucao );    
    
    
    // Gera e grava o id da evolução do caso  
    const campo_SituacaoBeneficio = TABELA_FILA.getRange( id+1, SITUACAO_BENEFICIO+1 );
    campo_SituacaoBeneficio.setValue( idEvolucao );      


    // Grava a data limite para evolução do tipo CONVOCADO PARA ACESSO
    const campo_data_limite = TABELA_FILA.getRange( id+1, DATA_LIMITE+1 );
    let dataLimiteFormatada;
    if(dataLimite != "") {
      let auxDataLimite = new Date( dataLimite );  
      auxDataLimite.setDate( auxDataLimite.getDate() + 1 );   
      dataLimiteFormatada = new Date(auxDataLimite).toLocaleString("pt-BR", {dateStyle: "short"});          
    } else {
      dataLimiteFormatada = "";
    }
    campo_data_limite.setValue( dataLimiteFormatada );      


    // Envia email para o órgão encaminhador e para a instituição,
    // caso a evolução seja diferente de INATIVAÇÃO
    if( idEvolucao != "1" ) {

      let mensagemDataLimite;
      if( idEvolucao == "3"  &&  dataLimite != "" ) {          

        mensagemDataLimite = `<br>Atentem-se à data limite de conclusão do acesso ao benefício: <b>${dataLimiteFormatada}</b>. 
                              Após este período, o beneficiário perderá a sua reserva da vaga e só poderá ser inserido novamente
                              em um novo processo de habilitação. Caso o beneficiário já esteja com o benefício liberado, o 
                              prazo pode ser desconsiderado.<br>`;

      } else {
        mensagemDataLimite = "";
      }

      const emailOrgaoEncaminhador = BUFFER_FILA[id-1][EMAIL_ORGAO_ENCAMINHADOR];
      const cpfRFCaso = (BUFFER_FILA[id-1][CPF_RF]).padStart(11, "0");
      const nomeRFCaso = BUFFER_FILA[id-1][REFERENCIA_FAMILIAR];
      const evolucaoCaso = idToNome( idEvolucao,  "SITUACOES_BENEFICIO" );
    
      const idInstituicao = parseInt(BUFFER_FILA[id-1][ORGAO_ENCAMINHADOR]);
      const emailInstituicao = BUFFER_ORGAOS_ENCAMINHADORES[idInstituicao-1][EMAIL_INSTITUICAO];
    
      const emails = [];
      if( isEmailValidBE(emailOrgaoEncaminhador) ) { emails.push(emailOrgaoEncaminhador) }
      if( isEmailValidBE(emailInstituicao) ) { emails.push(emailInstituicao) }
            
      enviarEmailBE( emails.join(","), cpfRFCaso, nomeRFCaso, evolucaoCaso, mensagemDataLimite );    

    } // Fim if   


  } catch( error ) {
    throw( "evoluirCasoBE - " + error.message );
  }

} // Fim da função evoluirCasoBE



/**
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 */



/**
 * Função para testar a função principal evoluirCasoBE
 */
function teste_evoluirCasoBE() {

  let idCaso = "23";
  let idEvolucao = "5";   
  let dataLimite = "";

  try {
    evoluirCasoBE( idCaso, idEvolucao, dataLimite );
  } catch( error ) {
    console.log( "teste_evoluirCasoBE - " + error.message );
  }

} // Fim da função teste_evoluirCasoBE




/**
 * ##### FIM DO MÓDULO evoluirCasoBE.gs #####
 */




