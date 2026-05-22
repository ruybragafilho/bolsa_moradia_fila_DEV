"use strict";

/**
 * Módulo:    tabelasMonitoramento.gs
 * Objetivo:  Armazenar as tabelas do sistema Bolsa Moradia
 */


/**
 * Planilha MONITORAMENTO
 */
const PLANILHA_MONITORAMENTO_ID  =  "1q7MrAuPprje6SfoQvWBqvDrOwESlY-9R5Ze-cQoJyXw";
const PLANILHA_MONITORAMENTO     =  SpreadsheetApp.openById(PLANILHA_MONITORAMENTO_ID);
const TABELA_MONITORAMENTO       =  PLANILHA_MONITORAMENTO.getSheetByName('Listagem de Atendidos');
const BUFFER_MONITORAMENTO       =  TABELA_MONITORAMENTO.getDataRange().getDisplayValues().splice(1);
const TAMANHO_MONITORAMENTO      =  BUFFER_MONITORAMENTO.length;



// Posições das colunas da planilha MONITORAMENTO
const SELO_PK                                   =  29;
const SITUACAO_BENEFICIO_TB_MONITORAMENTO       =  46;
const DATA_SITUACAO_BENEFICIO_TB_MONITORAMENTO  =  47;




/**
 * Função que pesquisa por um caso usando o selo como chave de pesquisa
 * @param {String} selo - SELO do caso
 * @returns A situação do benefício
 */
function beneficioLiberado( selo ) {

  const casos = BUFFER_MONITORAMENTO.filter( c => c[SELO_PK].replace(".", "").replace(",", "") == selo );
  
  if( casos.length > 0 ) {

      let caso = casos[0];

      if( caso[SITUACAO_BENEFICIO_TB_MONITORAMENTO].toUpperCase().includes( "LIBERADO" ) ) {
        
        return true;

      } else {
        return false;
      } 

  } else {
    return false;
  }  

} // Fim da função pesquisarCasoPorSelo()




function testebeneficioLiberado() {

    const selo = "943";

    const situacaoBeneficio = beneficioLiberado( selo );

    console.log( "Situação do benefício: " + situacaoBeneficio );
}


/**
 * ##### FIM DO MÓDULO tabelasMonitoramento.gs #####
 */


