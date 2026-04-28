/**
 * Módulo:    validacaoDadosBE.gs
 * Objetivo:  Módulo com funções que verificam a validade dos dados
 */


// Fumção Backend que verifica se o parâmetro passado é uma string válida
function isStringValidBE(str) {    
  return typeof str === 'string' && str.trim().length > 0;
}


// Fumção Backend que verifica se o parâmetro passado é um inteiro válido
function isIntegerValidBE( i ) {
  return Number.isInteger( i );
}  


// Fumção Backend que verifica se o parâmetro passado é uma data válida
function isDateValidBE(dateStr) {
  // Tenta criar um objeto Date
  const date = new Date(dateStr);
  
  // Verifica se o objeto Date é válido:
  // 1. date.getTime() retorna NaN se for inválida
  // 2. !isNaN(date) verifica se não é Not-a-Number
  return !isNaN(date.getTime());
}


// Função que verifica se o formato do endereço de email é válido
function isEmailValidBE(enderecoEmail) {

  // Regex padrão para validação de formato de e-mail
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  return ( isStringValidBE(enderecoEmail) && regex.test( enderecoEmail ) );

} // Fim da função validarEmail


// Função de sanitização - eliminação de caracteres html de um string
function escapeHtml(unsafe) {

  return unsafe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
    
} // Fim da função escapeHtml
    


/**
 * ##### FIM DO MÓDULO isStringValidaBE.gs #####
 */

