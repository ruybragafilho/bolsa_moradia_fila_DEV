"use strict";

/**
 * Módulo:    enviarEmailAlteracaoServicoBE.gs
 * Objetivo:  Envia um email para o endereço de e-mail informado
 */


/**
 * Função que envia um email para o endereço de e-mail informado
 *  
 * @param {String} enderecoEmail: enderecoEmail Endereço para onde o email será enviado
 * @param {String} cpfRFCaso: cpfRFCaso CPF do beneficiário do caso
 * @param {String} nomeRFCaso: nomeRFCaso Nome do beneficiário do caso
 * @param {String} servicoAnterior: servicoAnterior Serviço de referência anterior do caso
 * @param {String} servicoNovo: servicoNovo Novo serviço de referência do caso
 * 
 */
function enviarEmailAlteracaoServicoBE( enderecoEmail, cpfRFCaso, nomeRFCaso, servicoAnterior, servicoNovo ) {

  console.log("enderecoEmail: " + enderecoEmail);
  console.log("cpfRFCaso: " + cpfRFCaso);
  console.log("nomeRFCaso: " + nomeRFCaso);
  console.log("servicoAnterior: " + servicoAnterior);
  console.log("servicoNovo: " + servicoNovo);          

  try {
  
    MailApp.sendEmail({
  
      to: `${enderecoEmail}`,
      cc: `ruybragafilho@gmail.com`,
      subject: `Alteração serviço de referência Bolsa Moradia - ${nomeRFCaso} - ${(new Date()).toLocaleString("pt-BR")}`,
      htmlBody:  
`Prezado (a),<br><br>

Informamos que o serviço/equipamento de referência do Bolsa Moradia do(a) beneficiário(a) <b>${nomeRFCaso}</b>, CPF <b>${cpfRFCaso}</b>, foi alterado do(a) <b>${servicoAnterior}</b> para o(a) <b>${servicoNovo}</b>.<br><br>

Qualquer dúvida, procure a equipe da DPOP.<br><br><br>

Equipe Bolsa Moradia | Diretoria de Políticas para População em Situação de Rua, Migrantes e Refugiados | DPOP<br>
Secretaria Municipal de Assistência Social e Direitos Humanos | SMASDH<br>
Av. Afonso Pena, 342, 6º andar - Centro | Belo Horizonte/MG | CEP: 30130-001<br>
Telefone: (31) 3277-6373 / 3277-9994 | pbh.gov.br      <br><br> `
  
    });     

    console.log("EMAIL ENVIADO");            
  
  } catch( error ) {

    console.log( "enviarEmailAlteracaoServicoBE - " + error.message );    
    throw( "enviarEmailAlteracaoServicoBE - " + error.message );

  }

} // Fim da função enviarEmailAlteracaoServicoBE



/**
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 */

/**
 * Função para testar a função enviarEmail
 */
function teste_enviarEmailAlteracaoServico() {
  
  let enderecoEmail = "sigps@pbh.gov.br,ruy.braga@pbh.gov.br";
  let cpfRFCaso = "111.222.333-44";
  let nomeRFCaso = "João da Silva";
  let servicoAnterior = "Abrigo São Paulo";
  let servicoNovo = "Centro Pop Lagoinha";
  
  enviarEmailAlteracaoServicoBE( enderecoEmail, cpfRFCaso, nomeRFCaso, servicoAnterior, servicoNovo );
 
}




/**
 * ##### FIM DO MÓDULO enviarEmailAlteracaoServicoBE.gs #####
 */
