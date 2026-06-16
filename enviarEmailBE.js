"use strict";

/**
 * Módulo:    enviarEmailBE.gs
 * Objetivo:  Envia um email para o endereço de e-mail informado
 */


/**
 * Função que envia um email para o endereço de e-mail informado
 *  
 * @param {String} enderecoEmail: enderecoEmail Endereço para onde o email será enviado
 * @param {String} cpfRFCaso: cpfRFCaso CPF do beneficiário do caso
 * @param {String} nomeRFCaso: nomeRFCaso Nome do beneficiário do caso
 * @param {String} evolucaoCaso: evolucaoCaso Tipo de evolução do caso
 * @param {String} mensagemDataLimite: mensagemDataLimite Mensagem sobre a data limite  
 * 
 */
function enviarEmailBE( enderecoEmail, cpfRFCaso, nomeRFCaso, evolucaoCaso, mensagemDataLimite ) {

  console.log("enderecoEmail: " + enderecoEmail);
  console.log("cpfRFCaso: " + cpfRFCaso);
  console.log("nomeRFCaso: " + nomeRFCaso);
  console.log("evolucaoCaso: " + evolucaoCaso);
  console.log("mensagemDataLimite: " + mensagemDataLimite);          

  try {
  
    MailApp.sendEmail({
  
      to: `${enderecoEmail}`,
      cc: `ruybragafilho@gmail.com`,
      subject: `Atualização benefício Bolsa Moradia - ${nomeRFCaso} - ${(new Date()).toLocaleString("pt-BR")}`,
      htmlBody:  
`Prezado (a),<br><br>

informamos que houve alteração no status do benefício de <b>${nomeRFCaso}</b>, CPF <b>${cpfRFCaso}</b>, para <b>${evolucaoCaso}</b>. Pedimos que verifique as informações no sistema e comunique ao beneficiado, se necessário.<br>

${mensagemDataLimite}<br>

Qualquer dúvida, procure a equipe da DPOP.<br><br>

Equipe Bolsa Moradia | Diretoria de Políticas para População em Situação de Rua, Migrantes e Refugiados | DPOP<br>
Secretaria Municipal de Assistência Social e Direitos Humanos | SMASDH<br>
Av. Afonso Pena, 342, 6º andar - Centro | Belo Horizonte/MG | CEP: 30130-001<br>
Telefone: (31) 3277-6373 / 3277-9994 | pbh.gov.br      <br><br> `
  
    });     

    console.log("EMAIL ENVIADO");            
  
  } catch( error ) {

    console.log( "enviarEmailBE - " + error.message );    
    throw( "enviarEmailBE - " + error.message );

  }

} // Fim da função enviarEmailBE



/**
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 */

/**
 * Função para testar a função enviarEmail
 */
function teste_enviarEmail() {
  
  let enderecoEmail = "sigps@pbh.gov.br";
  let cpfRFCaso = "111.222.333-44";
  let nomeRFCaso = "João da Silva";
  let evolucaoCaso = "evolução";
  let mensagemDataLimite = "";
  
  enviarEmailBE( enderecoEmail, cpfRFCaso, nomeRFCaso, evolucaoCaso, mensagemDataLimite );
 
}


function limiteEmailDisponivel() {

  const n = MailApp.getRemainingDailyQuota();

  console.log( "Limite Email Disponivel: " + n );
}



/**
 * ##### FIM DO MÓDULO enviarEmailBE.gs #####
 */
