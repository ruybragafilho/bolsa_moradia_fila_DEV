"use strict";

/**
 * Módulo:    tabelasQuestionario.gs
 * Objetivo:  Armazenar a tabela principal do sistema Bolsa Moradia - Questionário
 */



/**
 * Planilha QUESTIONARIO
 */
const PLANILHA_QUESTIONARIO_ID  =  "1Lu69-a1VHgWcCEXRt9RKjmoVfOTmHFXBqpfM2G0pMnE";
const PLANILHA_QUESTIONARIO     =  SpreadsheetApp.openById(PLANILHA_QUESTIONARIO_ID);

const TABELA_QUESTIONARIO  =  PLANILHA_QUESTIONARIO.getSheetByName('QUESTIONARIO');
const BUFFER_QUESTIONARIO  =  TABELA_QUESTIONARIO.getDataRange().getDisplayValues().splice(1);
const NUM_QUESTIONARIO     =  BUFFER_QUESTIONARIO.length;



/**
 * Constantes que armazenam as posições das colunas da tabela QUESTIONARIO
 */
const ID_CASO           = 0;

const EMAIL_RESPONSAVEL_RESPOSTA = 1;
const DATA_ENVIO_EMAIL_QUESTIONARIO = 2;
const DATA_RESPOSTA_QUESTIONARIO = 3;
const RESPOSTAS = 4;
const OBSERVACAO = 5;




/**
 * Função backend que retorna a situação do questionário de um caso
 *    1 - Sem questionário a responder
 *    2 - Questionário enviado mas não respondido
 *    3 - Questionário enviado e respondido
 */       
function getSituacaoQuestionario( idCaso ) {

  const id = parseInt( idCaso );

  // Se id inválido, retorna uma exceção
  if( id < 1  ||  id > TAMANHO_FILA ) {
    throw( new Error( "getSituacaoQuestionario - ID Inválido" ) );
  }    


  let situacaoQuestionario;
  
  if( (BUFFER_QUESTIONARIO[id-1][DATA_ENVIO_EMAIL_QUESTIONARIO] === "") &&
      (BUFFER_QUESTIONARIO[id-1][DATA_RESPOSTA_QUESTIONARIO] === "") ) {

    // Sem questionário a responder
    situacaoQuestionario = "1";  

  } else if( (BUFFER_QUESTIONARIO[id-1][DATA_ENVIO_EMAIL_QUESTIONARIO] !== "") &&
             (BUFFER_QUESTIONARIO[id-1][DATA_RESPOSTA_QUESTIONARIO] === "") ) {

    // Questionário enviado mas não respondido
    situacaoQuestionario = "2";                  

  } else if( (BUFFER_QUESTIONARIO[id-1][DATA_ENVIO_EMAIL_QUESTIONARIO] !== "") &&
             (BUFFER_QUESTIONARIO[id-1][DATA_RESPOSTA_QUESTIONARIO] !== "") ) {

    // Questionário enviado e respondido
    situacaoQuestionario = "3";                  
    
  }


  return situacaoQuestionario;

} // Fim da função getSituacaoQuestionario




/**
 * Função backend para salvar o questionário
 */       
function salvarQuestionarioBE( jsonRespostasQuestionario ) {

  // Prepara os dados para serem gravados na planilha QUESTIONARIO

  const respostasQuestionario = JSON.parse( jsonRespostasQuestionario );

  const idCaso = parseInt( respostasQuestionario.idCaso );

  const repostas = [ respostasQuestionario.q1, 
                     respostasQuestionario.q2, 
                     respostasQuestionario.q3, 
                     respostasQuestionario.q4, 
                     respostasQuestionario.q5 ].join( ";" );

  const observacoes = respostasQuestionario.observacoes;


  // Se id inválido, retorna uma exceção
  if( idCaso < 1  ||  idCaso > TAMANHO_FILA ) {
    throw( new Error( "salvarQuestionarioBE - ID Inválido" ) );
  }      


  // Grava os dados do questionário na planilha QUESTIONARIO

  // TENTA PEGAR O LOCK
  const lock = LockService.getScriptLock();    

  try {

    lock.waitLock(10000);  
    
    
    // SE PEGAR O LOCK, PROSSEGUE COM A GRAVAÇÃO DOS DADOS DO QUESTIONÁRIO
    if( lock.hasLock() ) {


      // Gera, formata e grava a data de resposta do questionário
      let dataResposta = new Date().toLocaleString("pt-BR", {dateStyle: "short"});    
      const campo_data = TABELA_QUESTIONARIO.getRange( idCaso+1, DATA_RESPOSTA_QUESTIONARIO+1 );
      campo_data.setValue( dataResposta );    
      console.log( "\nsalvarQuestionarioBE - dataResposta" );
      
      // Grava os ids das respostas do questionário
      const campo_Respostas = TABELA_QUESTIONARIO.getRange( idCaso+1, RESPOSTAS+1 );
      campo_Respostas.setValue( repostas );      
      console.log( "\nsalvarQuestionarioBE - repostas" );
  
      // Grava as observações do questionário
      const campo_Observacoes = TABELA_QUESTIONARIO.getRange( idCaso+1, OBSERVACAO+1 );
      campo_Observacoes.setValue( observacoes );      
      console.log( "\nsalvarQuestionarioBE - observacoes" );


      // Aguarda sincronização de dados
      PLANILHA_QUESTIONARIO.waitForAllDataExecutionsCompletion(3);          
      SpreadsheetApp.flush();    
      
    } else {
  
      // SE NAO CONSEGUIR PEGAR O LOCK, LANCA UMA EXCESSAO
      throw( new Error( "salvarQuestionarioBE - Nao foi possivel pegar o LOCK" ) );
    } 


  } catch( error ) {

    console.log( "salvarQuestionarioBE - " + error.message );
    throw( "salvarQuestionarioBE - " + error.message );

  } finally {

    // Always release the lock for other waiting instances
    lock.releaseLock(); 
  }    

  return true;

} // Fim da função salvarQuestionarioBE

