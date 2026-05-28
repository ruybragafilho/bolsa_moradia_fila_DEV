"use strict";

/**
 * Módulo:    tabelasSistema.gs
 * Objetivo:  Armazenar as tabelas do sistema Bolsa Moradia
 */



/**
 * Planilha USUARIOS
 */
const PLANILHA_USUARIOS_ID        =  "1z1kH-MuamSlaQ_iKuwbEaSyDbULmku_jxkEK8CdNLzg";
const PLANILHA_USUARIOS           =  SpreadsheetApp.openById(PLANILHA_USUARIOS_ID);
const TABELA_USUARIOS             =  PLANILHA_USUARIOS.getSheetByName('USUARIOS');
const BUFFER_USUARIOS             =  TABELA_USUARIOS.getDataRange().getDisplayValues().splice(1);
const NUM_USUARIOS                =  BUFFER_USUARIOS.length;



/**
 * Planilha CODIGOS contendo as tabelas 
 */
const PLANILHA_CODIGOS_ID  =  "1eba9uB4QpIpLI7HJQWxJeUuR_pIG3bxI07deoirgN7k";
const PLANILHA_CODIGOS     =  SpreadsheetApp.openById(PLANILHA_CODIGOS_ID);

const TABELA_RESPOSTAS_SIMPLES      =  PLANILHA_CODIGOS.getSheetByName('RESPOSTAS_SIMPLES');
const TABELA_ORGAOS_ENCAMINHADORES  =  PLANILHA_CODIGOS.getSheetByName('ORGAOS_ENCAMINHADORES');
const TABELA_SITUACOES_BENEFICIO    =  PLANILHA_CODIGOS.getSheetByName('SITUACOES_BENEFICIO');
const TABELA_SITUACOES_VISTORIA     =  PLANILHA_CODIGOS.getSheetByName('SITUACOES_VISTORIA');
const TABELA_INTERVALOS_DE_TEMPO    =  PLANILHA_CODIGOS.getSheetByName('INTERVALOS_DE_TEMPO');
const TABELA_PARAMETROS             =  PLANILHA_CODIGOS.getSheetByName('PARAMETROS');
const TABELA_PERFIS                 =  PLANILHA_CODIGOS.getSheetByName('PERFIS');

const BUFFER_RESPOSTAS_SIMPLES      =  TABELA_RESPOSTAS_SIMPLES.getDataRange().getDisplayValues().splice(1);
const BUFFER_ORGAOS_ENCAMINHADORES  =  TABELA_ORGAOS_ENCAMINHADORES.getDataRange().getDisplayValues().splice(1);
const BUFFER_SITUACOES_BENEFICIO    =  TABELA_SITUACOES_BENEFICIO.getDataRange().getDisplayValues().splice(1);
const BUFFER_SITUACOES_VISTORIA     =  TABELA_SITUACOES_VISTORIA.getDataRange().getDisplayValues().splice(1);
const BUFFER_INTERVALOS_DE_TEMPO    =  TABELA_INTERVALOS_DE_TEMPO.getDataRange().getDisplayValues().splice(1);
const BUFFER_PARAMETROS             =  TABELA_PARAMETROS.getDataRange().getDisplayValues().splice(1);
const BUFFER_PERFIS                 =  TABELA_PERFIS.getDataRange().getDisplayValues().splice(1);

const NUM_RESPOSTAS_SIMPLES      =  BUFFER_RESPOSTAS_SIMPLES.length;
const NUM_ORGAOS_ENCAMINHADORES  =  BUFFER_ORGAOS_ENCAMINHADORES.length;
const NUM_SITUACOES_BENEFICIO    =  BUFFER_SITUACOES_BENEFICIO.length;
const NUM_SITUACOES_VISTORIA     =  BUFFER_SITUACOES_VISTORIA.length;
const NUM_INTERVALOS_DE_TEMPO    =  BUFFER_INTERVALOS_DE_TEMPO.length;
const NUM_PARAMETROS             =  BUFFER_PARAMETROS.length;
const NUM_PERFIS                 =  BUFFER_PERFIS.length;



/**
 * Constantes que armazenam as posições das colunas nas tabelas
 */

// Posição da coluna ID nas planilhas CODIGOS, FILA e USUARIOS
const ID = 0;


// Posições das colunas nas tabelas da planilha CODIGOS e USUARIOS
const NOME  = 1;
const ATIVO = 2;
const PESO_PARAMETRO = 3;
const PONTUACAO_PARAMETRO = 4;
const EMAIL_INSTITUICAO = 3;


// Posições das colunas da planilha USUARIOS
const EMAIL             = 1;
const NOME_USUARIO      = 3;
const ID_INSTITUICAO    = 4;
const TIPO_USUARIO      = 5;




/** 
 *  ####################################################
 *  #####                                          ##### 
 *  #####  IMPLEMENTAÇÃO DAS FUNÇÕES DESSE MÓDULO  #####
 *  #####                                          ##### 
 *  ####################################################
 */




/**
 * Função que retorna uma cópia da tabela cujo nome é passado como parâmetro.
 * É chamada pelo front-end para obter os ids e nomes das informações que
 * serão mostradas na tela
 * 
 * @param {String} nomeTabela: Nome da tabela a qual os ids se referem. 
 * 
 * return Uma cópia da tabela
 */
function obterTabelaCompleta( nomeTabela ) {
  
  // Validação de parâmetros da função
  if( !isStringValidBE(nomeTabela) ) {
    throw( new Error( "idToNome - Tabela Inválida" ) );    
  }


  // Seleciona a tabela apropriada

  let bufferTabela;

  switch( nomeTabela ) {
    case "RESPOSTAS_SIMPLES":        bufferTabela = BUFFER_RESPOSTAS_SIMPLES;
                                     break;                                                                  
    case "ORGAOS_ENCAMINHADORES":    bufferTabela = BUFFER_ORGAOS_ENCAMINHADORES;
                                     break;
    case "SITUACOES_BENEFICIO":      bufferTabela = BUFFER_SITUACOES_BENEFICIO;
                                     break;            
    case "SITUACOES_VISTORIA":       bufferTabela = BUFFER_SITUACOES_VISTORIA;
                                     break;
    case "INTERVALOS_DE_TEMPO":      bufferTabela = BUFFER_INTERVALOS_DE_TEMPO;
                                     break;
    case "PARAMETROS":               bufferTabela = BUFFER_PARAMETROS;
                                     break;
    case "PERFIS":                   bufferTabela = BUFFER_PERFIS;
                                     break;                           
    default:                         throw( new Error( "Tabela inválida" ) ); 
  }

  
  // Cria uma cópia da tabela
  let tabela = [];  
  bufferTabela.forEach( linha => tabela.push( linha ) );


  // Retorna a tabela
  return tabela;


} // Fim da função obterTabelaCompleta



/**
 * Função que recebe um id e o nome da tabela a qual o id se refere, 
 * e retorna o nome relacionado ao ID. 
 */
function idToNome( id, nomeTabela ) {
  
  // Validação de parâmetros da função
  
  if( !isStringValidBE(id) ) {
    return "";
  } 

  if( !isStringValidBE(nomeTabela) ) {
    throw( new Error( "idToNome - Tabela Inválida" ) );    
  }


  // Seleciona a tabela apropriada

  let bufferTabela;
  let tamanhoTabela;

  switch( nomeTabela ) {
    case "RESPOSTAS_SIMPLES":        bufferTabela = BUFFER_RESPOSTAS_SIMPLES;
                                     tamanhoTabela = NUM_RESPOSTAS_SIMPLES;
                                     break;                                                        
    case "ORGAOS_ENCAMINHADORES":    bufferTabela = BUFFER_ORGAOS_ENCAMINHADORES;
                                     tamanhoTabela = NUM_ORGAOS_ENCAMINHADORES;
                                     break;          
    case "SITUACOES_BENEFICIO":      bufferTabela = BUFFER_SITUACOES_BENEFICIO;
                                     tamanhoTabela = NUM_SITUACOES_BENEFICIO;
                                     break;               
    case "SITUACOES_VISTORIA":       bufferTabela = BUFFER_SITUACOES_VISTORIA;
                                     tamanhoTabela = NUM_SITUACOES_VISTORIA;
                                     break;
    case "INTERVALOS_DE_TEMPO":      bufferTabela = BUFFER_INTERVALOS_DE_TEMPO;
                                     tamanhoTabela = NUM_INTERVALOS_DE_TEMPO;
                                     break;                                                                                                                           
    case "PARAMETROS":               bufferTabela = BUFFER_PARAMETROS;
                                     tamanhoTabela = NUM_PARAMETROS;
                                     break;      
    case "PERFIS":                   bufferTabela = BUFFER_PERFIS;
                                     tamanhoTabela = NUM_PERFIS;
                                     break;                                                                                                              
    default:                         throw( new Error( "idToNome - Tabela Inválida" ) );    
  }


  // Converte o id para Integer
  const idItem = parseInt(id);
  if( !isIntegerValidBE(idItem) ) {
    throw( new Error( "idToNome - ID Inválido: " + idItem ) );      
  }

  // Se id está fora dos limites inferior ou superior, lança uma exceção
  if( idItem < 1  ||  idItem > tamanhoTabela ) {
    throw( new Error( "idToNome - ID fora do intervalo da tabela: " + idItem + " - " + tamanhoTabela ) );      
  }

  // Retorna o nome referente ao ID
  return bufferTabela[idItem-1][NOME];  


} // Fim da função idToNome




/** 
 *  #################################################
 *  #####                                       ##### 
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 *  #####                                       ##### 
 *  #################################################
 */




/**
 * Função para testar a função obterTabelaCompleta
 */
function teste_obterTabelaCompleta() {
  
  const nomeTabela = "ORGAOS_ENCAMINHADORES";

  const retorno = obterTabelaCompleta( nomeTabela );

  console.log( retorno );

} // Fim da função teste_obterTabelaCompleta



/**
 * Função para testar a função idToNome
 */
function teste_idToNome() {

  const id = "3";
  const nomeTabela = "SITUACOES_BENEFICIO";  

  const retorno = idToNome( id, nomeTabela );

  console.log( retorno );

} // Fim da função teste_idToNome




/**
 * ##### FIM DO MÓDULO tabelasSistema.gs #####
 */