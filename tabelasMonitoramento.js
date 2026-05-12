"use strict";

/**
 * Módulo:    tabelasMonitoramento.gs
 * Objetivo:  Armazenar as tabelas do sistema Bolsa Moradia
 */


/**
 * Planilha MONITORAMENTO
 */
const PLANILHA_MONITORAMENTO_ID  =  "1hMm3u8nzAAaaLVUe3V1XA6aP48XH3i4yMsYKxeFjGgw";
const PLANILHA_MONITORAMENTO     =  SpreadsheetApp.openById(PLANILHA_MONITORAMENTO_ID);
const TABELA_MONITORAMENTO       =  PLANILHA_MONITORAMENTO.getSheetByName('Listagem de Atendidos');
const BUFFER_MONITORAMENTO       =  TABELA_MONITORAMENTO.getDataRange().getDisplayValues().splice(1);
const TAMANHO_MONITORAMENTO      =  BUFFER_MONITORAMENTO.length;



// Posições das colunas da planilha MONITORAMENTO
const SELO_PK                              =  29;
const SITUACAO_BENEFICIO_TB_MONITORAMENTO  =  46;




/**
 * Função que pesquisa por um caso usando o selo como chave de pesquisa
 * @param {String} selo - SELO do caso
 * @returns A situação do benefício
 */
function pesquisarCasoPorSelo( selo ) {

  const caso = BUFFER_MONITORAMENTO.filter( c => c[SELO_PK] == selo );
  
  if( caso.length > 0 ) {

    return new String( caso[0][SITUACAO_BENEFICIO_TB_MONITORAMENTO] );

  } else {

    return "NI";

  }  

} // Fim da função pesquisarCasoPorSelo()




function testepesquisarCasoPorSelo() {

    const selo = 292;

    const situacaoBeneficio = pesquisarCasoPorSelo( selo );

    console.log( "Situação do benefício: " + situacaoBeneficio );
}


/**
 * ##### FIM DO MÓDULO tabelasMonitoramento.gs #####
 */

