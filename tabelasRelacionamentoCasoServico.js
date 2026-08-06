"use strict";

/**
 * Módulo:    TabelaRelacionamentoCasoServico.gs
 * Objetivo:  Armazenar a tabela com o relacionamento Caso c/ Serviços de Referência
 */


/**
 * Planilha RELACIONAMENTO CASO SERVICO
 */
const PLANILHA_RELACIONAMENTO_CASO_SERVICO_ID  =  "1WdMEldsbbbqmJYEI3Up_X6geJFgBLolFQa4H9FDpnaI";
const PLANILHA_RELACIONAMENTO_CASO_SERVICO     =  SpreadsheetApp.openById(PLANILHA_RELACIONAMENTO_CASO_SERVICO_ID);

const TABELA_RELACIONAMENTO_CASO_SERVICO       =  PLANILHA_RELACIONAMENTO_CASO_SERVICO.getSheetByName('RELACIONAMENTO');
const TABELA_INDICE_INVERTIDO_RELACIONAMENTO   =  PLANILHA_RELACIONAMENTO_CASO_SERVICO.getSheetByName('INDICE_INVERTIDO_RELACIONAMENTO');

const BUFFER_RELACIONAMENTO_CASO_SERVICO       =  TABELA_RELACIONAMENTO_CASO_SERVICO.getDataRange().getDisplayValues().splice(1);
const BUFFER_INDICE_INVERTIDO_RELACIONAMENTO   =  TABELA_INDICE_INVERTIDO_RELACIONAMENTO.getDataRange().getDisplayValues().splice(1);

const TAMANHO_RELACIONAMENTO_CASO_SERVICO      =  BUFFER_RELACIONAMENTO_CASO_SERVICO.length;
const TAMANHO_INDICE_INVERTIDO_RELACIONAMENTO  =  BUFFER_INDICE_INVERTIDO_RELACIONAMENTO.length;


// Posições das colunas da planilha RELACIONAMENTO CASO SERVICO
// const ID_CASO = 0;  Já declarado em tabelasQuestionario.js

const ID_RELACIONAMENTO          =  0;
const IDS_RELACIONAMENTOS        =  1;

const ID_CASO_RELACIONAMENTO     =  1;
const ID_SERVICO_RELACIONAMENTO  =  2;
const DATA_INFORMACAO            =  3;
const ID_RESPONSAVEL_INFORMACAO  =  4;




/** 
 *  ####################################################
 *  #####                                          ##### 
 *  #####  IMPLEMENTAÇÃO DAS FUNÇÕES DESSE MÓDULO  #####
 *  #####                                          ##### 
 *  ####################################################
 */




/**
 * Função que obtém os serviços de referência associados a um caso específico.
 * @param {String} idCaso 
 * @returns array de objetos contendo os serviços de referência do caso
 */
function obterServicosReferenciaDoCaso( idCaso ) {

  // Se id caso inválido, retorna uma exceção
  let auxIDCaso = parseInt( idCaso );  
  if( auxIDCaso < 1  ||  auxIDCaso > TAMANHO_FILA ) {
    throw( new Error( "obterServicosReferenciaDoCaso - ID Caso Inválido" ) );
  }    

  const indicesRelacionamentos = BUFFER_INDICE_INVERTIDO_RELACIONAMENTO[auxIDCaso-1][IDS_RELACIONAMENTOS].split(";").map(id => parseInt(id));

  let servicosReferencia = []  
  let relacionamento;
  indicesRelacionamentos.forEach( ir => { 
    
    if( ir != 0 ) {

      relacionamento = BUFFER_RELACIONAMENTO_CASO_SERVICO[ parseInt(ir) - 1 ];
      servicosReferencia.push(
        {         
          idServico: relacionamento[ID_SERVICO_RELACIONAMENTO],
          dataInformacao: relacionamento[DATA_INFORMACAO],
          idResponsavelInformacao: relacionamento[ID_RESPONSAVEL_INFORMACAO] 
        }
      );
    }
    
  }); // Fim do forEach

  return servicosReferencia;

} // Fim da função obterServicosReferenciaDoCaso



/**
 * Função que obtém o serviço de referência ativo associado a um caso específico.
 * @param {String} servicosReferencia 
 * @returns Serviço de referência ativo do caso
 */
function obterServicoReferenciaAtivo( servicosReferencia ) {

  const servicoAtivo = servicosReferencia[servicosReferencia.length - 1];
  
  return servicoAtivo;

} // Fim da função obterServicoReferenciaAtivo



/**
 * Função que obtém o serviço de referência ativo associado a um caso específico.
 * @param {String} servicosReferencia 
 * @returns Serviço de referência ativo do caso
 */
function adicionarServicoReferenciaAtivoParaUmCasoBE( relacionamentoServicoCaso ) {

  const obj_relacionamentoServicoCaso = JSON.parse(relacionamentoServicoCaso);

  const idCaso                   =  obj_relacionamentoServicoCaso.idCaso;
  const idServicoReferencia      =  obj_relacionamentoServicoCaso.idServicoReferencia;
  const idResponsavelInformacao  =  obj_relacionamentoServicoCaso.idResponsavelInformacao;
  

  // Se id caso inválido, retorna uma exceção
  let auxID = parseInt( idCaso );  
  if( auxID < 1  ||  auxID > TAMANHO_FILA ) {
    throw( new Error( "adicionarServicoReferenciaAtivoParaUmCaso - ID Caso Inválido" ) );
  }  

  // Se id serviço inválido, retorna uma exceção
  auxID = parseInt( idServicoReferencia );  
  if( auxID < 1  ||  auxID > NUM_ORGAOS_ENCAMINHADORES ) {
    throw( new Error( "adicionarServicoReferenciaAtivoParaUmCaso - ID Serviço Inválido" ) );
  }    

  // Se id usuário sistema inválido, retorna uma exceção
  auxID = parseInt( idResponsavelInformacao );  
  if( auxID < 1  ||  auxID > NUM_USUARIOS ) {
    throw( new Error( "adicionarServicoReferenciaAtivoParaUmCaso - ID Responsável Informação Inválido" ) );
  }      

  // Verifica se o usuário do app tem permissão para registrar Documentação Pendente
  /*const usuarioLogado = JSON.parse( autenticarUsuario() );
  if( usuarioLogado.instituicao != "0" || usuarioLogado.tipo != "1" ) {
    throw( new Error( "Usuário sem permissão para registrar nova entidade de referência para o caso" ) );
  } */ 

  try {


    // Gera o novo id do relacionamento
    // Gera da data da informação == data de hoje
    // Gera e grava a nova linha na tabela relacionamento
    const idNovoRelacionamento = parseInt(TAMANHO_RELACIONAMENTO_CASO_SERVICO) + 1;
    const dataInformacao = new Date();      
    const novoRelacionamento = [ idNovoRelacionamento, 
                                 idCaso, 
                                 idServicoReferencia, 
                                 dataInformacao.toLocaleString("pt-BR", {dateStyle: "short"}), 
                                 idResponsavelInformacao ];
    TABELA_RELACIONAMENTO_CASO_SERVICO.appendRow( novoRelacionamento );

  

    // Gera e grava a nova linha na tabela de índice invertido
    let indicesServicosReferencia = BUFFER_INDICE_INVERTIDO_RELACIONAMENTO[parseInt(idCaso)-1][IDS_RELACIONAMENTOS].split(";");
    if( indicesServicosReferencia[indicesServicosReferencia.length-1] == "0" ) {
      indicesServicosReferencia[indicesServicosReferencia.length-1] = idNovoRelacionamento;
    } else {
      indicesServicosReferencia.push(idNovoRelacionamento);
    }                          
    const campo_IndicesRelacionamentos = TABELA_INDICE_INVERTIDO_RELACIONAMENTO.getRange( parseInt(idCaso)+1, IDS_RELACIONAMENTOS+1 );
    campo_IndicesRelacionamentos.setValue( indicesServicosReferencia.join(";") );                                          


  } catch( error ) {
    throw( "adicionarServicoReferenciaAtivoParaUmCaso - " + error.message );
  }                              

} // Fim da função adicionarServicoReferenciaAtivoParaUmCasoBE



/**
 * Função auxiliar para formatar a data da informação dos relacionamentos
 */
function formatarData() {

    const NOVO_BUFFER_RELACIONAMENTO_CASO_SERVICO = [];

    BUFFER_RELACIONAMENTO_CASO_SERVICO.forEach( r => {

      // Formata a data da informação
      let dataFormatada;
      if(r[DATA_INFORMACAO].includes("-") ) {
        dataFormatada = new Date( r[DATA_INFORMACAO] );
        dataFormatada.setDate( dataFormatada.getDate() + 1 );  
        dataFormatada = dataFormatada.toLocaleString("pt-BR", {dateStyle: "short"});      
      } else {
        dataFormatada = r[DATA_INFORMACAO];
      }

      // Cria o caso com a data formatada
      let RELACIONAMENTO_CASO_SERVICO = [ r[ID_RELACIONAMENTO],
                                          r[ID_CASO_RELACIONAMENTO],
                                          r[ID_SERVICO_RELACIONAMENTO],
                                          dataFormatada,
                                          r[ID_RESPONSAVEL_INFORMACAO]
                                        ];     
                                               
      // Insere o caso no novo buffer
      NOVO_BUFFER_RELACIONAMENTO_CASO_SERVICO.push( RELACIONAMENTO_CASO_SERVICO );                                                    

    });

    // Grava o novo buffer na planilha
    TABELA_RELACIONAMENTO_CASO_SERVICO.getRange( 2, 1, BUFFER_RELACIONAMENTO_CASO_SERVICO.length, 5 ).setNumberFormat("@").setValues( NOVO_BUFFER_RELACIONAMENTO_CASO_SERVICO );
    PLANILHA_RELACIONAMENTO_CASO_SERVICO.waitForAllDataExecutionsCompletion(2);      
    SpreadsheetApp.flush();      

} // Fim da função formatarData




/** 
 *  #################################################
 *  #####                                       ##### 
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 *  #####                                       ##### 
 *  #################################################
 */




/**
 * Função de teste para a função obterServicosReferenciaDoCaso
 */
function testeObterServicosReferenciaDoCaso() {

  const idCaso = "2";

  const servicosReferencia = obterServicosReferenciaDoCaso( idCaso );

  servicosReferencia.forEach( sr => {
    console.log(`ID do Serviço: ${sr.idServico}, Data da Informação: ${sr.dataInformacao}, ID do Responsável pela Informação: ${sr.idResponsavelInformacao}`);
  });

} // Fim da função testeObterServicosReferenciaDoCaso



/**
 * Função de teste para a função obterServicoReferenciaAtivo
 */
function testeObterServicoReferenciaAtivo() {

  const idCaso = "2";

  const servicosReferencia = obterServicosReferenciaDoCaso( idCaso );

  const servicoAtivo = obterServicoReferenciaAtivo( servicosReferencia );

  console.log(`ID do Serviço Ativo: ${servicoAtivo.idServico}, Data da Informação: ${servicoAtivo.dataInformacao}, ID do Responsável pela Informação: ${servicoAtivo.idResponsavelInformacao}`);

} // Fim da função testeObterServicoReferenciaAtivo



/**
 * Função de teste para a função adicionarServicoReferenciaAtivoParaUmCaso
 */
function testeAdicionarServicoReferenciaAtivoParaUmCaso() {

  const obj_relacionamentoServicoCaso = {
    idCaso: "2",
    idServicoReferencia: "27",
    idResponsavelInformacao: "1"
  };

  const json_relacionamentoServicoCaso = JSON.stringify(obj_relacionamentoServicoCaso);

  try {
    adicionarServicoReferenciaAtivoParaUmCasoBE( json_relacionamentoServicoCaso );
  } catch( error ) {
    console.log( `testeAdicionarServicoReferenciaAtivoParaUmCaso - ${error.message}` );
  }  

} // Fim da função testeAdicionarServicoReferenciaAtivoParaUmCaso




