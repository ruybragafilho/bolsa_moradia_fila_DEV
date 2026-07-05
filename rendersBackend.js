"use strict";
    


/**
 * Gerador de Checkbox
 *
 * @param {Array} tabela - Tabela de onde os ids e os nomes das opções serão lidos
 * @param {String} idCheckBox - O id do checkbox que será gerado
 * @param {Integer} inline - 1 se o checkbox a ser gerado for inline, 0 caso contrário
 */
function gerarCheckBox( tabela, idCheckBox, inline ) {
 
  // Gera, dinamicamente, as opções do Check Box
  let checkBoxGerado = "";
  let idItem;
  let classInline = inline ? "form-check-inline" : "";

  tabela.forEach( item => {

    if( parseInt( item[ATIVO] ) ) {

      idItem = idCheckBox + item[ID];

      checkBoxGerado += 
      `<div class="form-check ${classInline}">
         <input class="form-check-input" type="checkbox" name="${idCheckBox}" id="${idItem}" value="${item[ID]}">
         <label class="form-check-label" for="${idItem}">${item[NOME]}</label>
       </div>`;

    }
  }); // Fim do tabela.forEach
      
  return checkBoxGerado;

} // Fim da função gerarCheckBox



/**
 * Gerador de Select
 *
 * @param {Array} tabela - Tabela de onde os ids e os nomes das opções serão lidos
 *
 */
function gerarSelectBackend( tabela ) {
 
  // Gera, dinamicamente, as opções do Check Box
  let selectGerado = `<option selected value="">Selecione</option>`;

  tabela.forEach( item => {

    if( parseInt( item[ATIVO] ) ) {

      selectGerado += `<option value="${item[ID]}">${item[NOME]}</option>`;

    }

  }); // Fim do tabela.forEach
      
  // Retorna o select gerado  
  return selectGerado;

} // Fim da função gerarSelectBackend    



/**
 * Gerador de Checkbox no formato de tabela
 *
 * @param {Array} tabela - Tabela de onde os ids e os nomes das opções serão lidos
 * @param {String} idCheckBox - O id do checkbox que será gerado
 * @param {Number} numColunas - Número de colunas na tabela de checkbox
 * 
 */
function gerarCheckBoxTable( tabela, idCheckBox, numColunas ) {

  
  if( numColunas < 1 ) {
    numColunas = 1;
  }

  let i = 0;
  let larguraColuna = 12 / numColunas;
  
 
  // Gera, dinamicamente, as opções do Check Box
  let checkBoxGerado = `<table class="table table-sm table-borderless">
                          <tbody>`;
  let idItem;
  
  tabela.forEach( item => {

    if( parseInt( item[ATIVO] ) ) {

      idItem = idCheckBox + item[ID];

      if( (i % numColunas) == 0 ) {
        checkBoxGerado += `<tr>`;
      }

      checkBoxGerado += 
      `<td class="col-${larguraColuna}">
         <div class="form-check">
           <input class="form-check-input" type="checkbox" name="${idCheckBox}" id="${idItem}" value="${item[ID]}">
           <label class="form-check-label" for="${idItem}">${item[NOME]}</label>
         </div>
       </td>`;

      if( (numColunas>1) && (i % numColunas) == (numColunas-1) ) {
        checkBoxGerado += `</tr>`;
      }

      ++i;

    } // Fim do if item ativo

  }); // Fim do tabela.forEach

  checkBoxGerado += `  </tbody>
                     </table>`;
      
  return checkBoxGerado;

} // Fim da função gerarCheckBoxTable



/**
 * Gerador de Radio Button
 *
 * @param {Array} tabela - Tabela de onde os ids e os nomes das opções serão lidos
 * @param {String} nameRadioButton - O name do radio button que será gerado
 * @param {Integer} inline - 1 se o radio button a ser gerado for inline, 0 caso contrário
 */
function gerarRadioButton( tabela, nameRadioButton, inline ) {
 
  // Gera, dinamicamente, as opções do Radio Button
  let radioButtonGerado = "";
  let idItem;
  let classInline = inline ? "form-check-inline" : "";

  tabela.forEach( item => {

    if( parseInt( item[ATIVO] ) ) {

      idItem = nameRadioButton + item[ID];

      radioButtonGerado += 
      `<div class="form-check ${classInline}">
         <input class="form-check-input border border-secondary-subtle" type="radio" name="${nameRadioButton}" id="${idItem}" value="${item[ID]}">
         <label class="form-check-label" for="${idItem}">${item[NOME]}</label>
       </div>`;

    }
  }); // Fim do tabela.forEach
      
  return radioButtonGerado;

} // Fim da função gerarRadioButton





