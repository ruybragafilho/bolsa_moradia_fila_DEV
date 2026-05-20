"use strict";

/**
 * Módulo:    tabelasVistoria.gs
 * Objetivo:  Armazenar as tabelas do sistema Bolsa Moradia
 */


/**
 * Planilha VISTORIAS
 */
const PLANILHA_VISTORIAS_ID  =  "1FQkpBwnX181zPd2bZl1NyEsL01MGdvdAKT_uGzLMLSQ";
const PLANILHA_VISTORIAS     =  SpreadsheetApp.openById(PLANILHA_VISTORIAS_ID);
const TABELA_VISTORIAS       =  PLANILHA_VISTORIAS.getSheetByName('VISTORIAS');
const BUFFER_VISTORIAS       =  TABELA_VISTORIAS.getDataRange().getDisplayValues().splice(1);
const TAMANHO_VISTORIAS      =  BUFFER_VISTORIAS.length;



// Posições das colunas da planilha VISTORIAS
const SELO_FK                      =  1;
const CPF_VISTORIA                 =  3;
const DATA_SOLICITACAO_VISTORIA    =  6;
const DATA_VISTORIA                =  7;
const DATA_LAUDO                   =  9;
const DESCRICAO_LAUDO              = 10;
const INFORMACAO_COMPLEMENTAR      = 12;





/**
 * Função que pesquisa por todas as vistorias relacionadas a um CPF de RF
 * @param {String} cpf - CPF da RF
 * @returns Um array de objetos, onde cada objeto representa uma vistoria relacionada ao CPF da RF, 
 *          contendo as seguintes propriedades: selo, dataSolicitacao, dataVistoria, dataLaudo, 
 *          descricaoLaudo e informacaoComplementar
 */
function pesquisarVistoriasPorCPF( cpf ) {

  const vistorias = BUFFER_VISTORIAS.filter( v => v[CPF_VISTORIA].padStart(11, "0") == cpf.padStart(11, "0") )
                                    .map( v => { return { selo: v[SELO_FK],
                                                          dataSolicitacao: v[DATA_SOLICITACAO_VISTORIA],
                                                          dataVistoria: v[DATA_VISTORIA],
                                                          dataLaudo: v[DATA_LAUDO],
                                                          descricaoLaudo: v[DESCRICAO_LAUDO],
                                                          informacaoComplementar: v[INFORMACAO_COMPLEMENTAR] };
                                               });

  return vistorias;

} // Fim da função pesquisarVistoriasPorCPF()



/**
 * Função que retorn a a situação da última vistoria de um caso, a partir do histórico de vistorias do caso
 * @param {Array de vistorias} vistoriasCaso - Array com o historico de vistorias do caso
 * @returns Status da últimaVistoria do caso
 */
function getSituacaoVistoria( vistoriasCaso ) {

  if( vistoriasCaso.length == 0 ) {
    return "";  
  }

  let ultimaVistoria = vistoriasCaso[vistoriasCaso.length-1];

  let dataSolicitacao = ultimaVistoria.dataSolicitacao != "";
  let dataVistoria = ultimaVistoria.dataVistoria != "";
  let dataLaudo = ultimaVistoria.dataLaudo != "";

  if( dataLaudo ) {

    if( ultimaVistoria.descricaoLaudo.includes("Aprovado")  ) {
      return BUFFER_SITUACOES_VISTORIA[3][ID];         }
    else if( ultimaVistoria.descricaoLaudo.includes("Reprovado") ) {
      return BUFFER_SITUACOES_VISTORIA[2][ID];     
    } else if( ultimaVistoria.descricaoLaudo.includes("Passível") ) {
      return BUFFER_SITUACOES_VISTORIA[1][ID];     
    } else {
      return ""    
    }

  }

  if( dataVistoria ) {
    // Última vistoria passível de aprovação
    return BUFFER_SITUACOES_VISTORIA[1][ID];             
  }  

  if( dataSolicitacao ) {
    // Vistoria solicitada
    return BUFFER_SITUACOES_VISTORIA[0][ID];             
  }

  return "";

} // Fim da função getSituacaoVistoria



/**
 * Função para testar a função pesquisarVistoriasPorCPF() 
 * do módulo tabelasVistoria.gs
 */
function testePesquisarVistoriasPorCPF() {

  let cpf = "07434335645";

  let vistorias = pesquisarVistoriasPorCPF( cpf );

  console.log( `Vistorias encontradas para o CPF ${cpf}:` );
  console.log( vistorias );

  let situacaoVistoria = getSituacaoVistoria( vistorias );

  console.log( `Situação da última vistoria para o CPF ${cpf}: ${situacaoVistoria}` );

} // Fim da função testePesquisarVistoriasPorCPF()



/**
 * ##### FIM DO MÓDULO tabelasVistoria.gs #####
 */


