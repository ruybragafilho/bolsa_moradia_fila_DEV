"use strict";

/**
 * Módulo:    tabelasEncaminhamentos.gs
 * Objetivo:  Importar os dados das tabelas de encaminhamentos para a tabela interna com os resumos dos casos
 */



/**
 * Planilha CASOS EXTERNOS - TRABALHADORES EXTERNOS (casos registrados por trabalhadores externos)
 */
const PLANILHA_CASOS_EXTERNOS_ID   =  "1PVOFfOORpXJx4FsWJLQskc9x0ZanMjQMd3roJ6hib5w";
const PLANILHA_CASOS_EXTERNOS      =  SpreadsheetApp.openById(PLANILHA_CASOS_EXTERNOS_ID);

const TABELA_CASOS_EXTERNOS             =  PLANILHA_CASOS_EXTERNOS.getSheetByName('PESSOA');
const BUFFER_CASOS_EXTERNOS             =  TABELA_CASOS_EXTERNOS.getDataRange().getDisplayValues().splice(1);
const NUM_LINHAS_TABELA_CASOS_EXTERNOS  =  BUFFER_CASOS_EXTERNOS.length;


/**
 * Planilha CASOS PBH - TRABALHADORES PBH (casos registrados por trabalhadores PBH)
 */
const PLANILHA_CASOS_PBH_ID   =  "1JVmJIM8FPePpgdOeXXgd9C307ZSj1tzE99MLiajeHCc";
const PLANILHA_CASOS_PBH      =  SpreadsheetApp.openById(PLANILHA_CASOS_PBH_ID);

const TABELA_CASOS_PBH             =  PLANILHA_CASOS_PBH.getSheetByName('PESSOA');
const BUFFER_CASOS_PBH             =  TABELA_CASOS_PBH.getDataRange().getDisplayValues().splice(1);
const NUM_LINHAS_TABELA_CASOS_PBH  =  BUFFER_CASOS_PBH.length;



// Posições das colunas das planilhas de encaminhamentos
const UNI_ID = 0;
const UNI_CPF_RF = 1;
const UNI_COD_PARENTESCO = 2;
const UNI_RF_2 = 3;
const UNI_NOME = 4;
const UNI_CPF = 5;
const UNI_DATA_NASCIMENTO = 6;
const UNI_NOME_SOCIAL = 7;
const UNI_NOME_MAE = 8;
const UNI_RACA_COR = 9;
const UNI_GENERO = 10;
const UNI_ORIENTACAO_SEXUAL = 11;
const UNI_ESCOLARIDADE = 12;
const UNI_DATA_ATUALIZACAO_CADUNICO = 13;
const UNI_TEMP_RESIDENCIA_BH = 14;
const UNI_TEMPO_SITUACAO_DE_RUA = 15;
const UNI_TELEFONE = 16;
const UNI_RENDA_FAMILIAR = 17;
const UNI_PCD = 18;
const UNI_PROBLEMAS_SAUDE = 19;
const UNI_GESTANTE = 20;
const UNI_TRABALHO_FORMAL = 21;
const UNI_TRABALHO_OUTROS = 22;
const UNI_BENEFICIO = 23;
const UNI_ESTAMOS_JUNTOS = 24;
const UNI_INSTITUCIONALIZACAO = 25;
const UNI_DIAGNOSTICO = 26;
const UNI_AMEACA_CONFLITO_VIOLENCIA = 27;
const UNI_DESC_AMEACA_CONFLITO_VIOLENCIA = 28;
const UNI_PROSTITUICAO = 29;
const UNI_CEA_ACOLHIMENTO_INSTITUCIONAL = 30;
const UNI_CEA_PRIVACAO_CONVIVIO = 31;
const UNI_DESC_CEA_PRIVACAO_CONVIVIO = 32;
const UNI_CEA_TRABALHO_INFANTIL_EXPLORACAO_SEXUAL = 33;
const UNI_AUTONOMIA = 34;
const UNI_INFO_ADICIONAIS = 35;
const UNI_LOCAL_COMUM = 36;
const UNI_ORGAO_ENCAMINHADOR = 37;
const UNI_NOME_TECNICO_ENCAMINHADOR = 38;
const UNI_EMAIL_TECNICO_ENCAMINHADOR = 39;
const UNI_DATA_REGISTRO_ENCAMINHAMENTO = 40;



/** 
 *  ####################################################
 *  #####                                          ##### 
 *  #####  IMPLEMENTAÇÃO DAS FUNÇÕES DESSE MÓDULO  #####
 *  #####                                          ##### 
 *  ####################################################
 */



/**
 * Função que obtém um caso da tabela passada como parâmetro
 * @param {Array 2D} tabela onde será obtido p caso
 * @returns 1 caso
 */

let linhaTabela = 0;

function obterCaso( tabela ) {

  let caso = [];

  let numLinhas = tabela.length;

  
  caso.push( tabela[linhaTabela] );  
  ++linhaTabela;

  
  while(  ( linhaTabela < numLinhas ) &&
          ( tabela[linhaTabela][UNI_CPF_RF] == caso[0][UNI_CPF_RF] ) ) {

    caso.push( tabela[linhaTabela] );
    ++linhaTabela;
  }

  return caso;

} // Fim da função obterCaso



/**
 * Função que mostra, no console, o caso passado como parâmetro
 * @param {Array 2D} caso caso que será impresso em console 
 */
function mostrarCaso( caso ) {
  console.log( "CASO" );
  caso.forEach( c => {        
      console.log( c );  
  });
} // Fim da função mostrarCaso



function testeObterCaso() {

  let teste = obterCaso( BUFFER_CASOS_EXTERNOS );
  mostrarCaso( teste )
}



let idsParametrosCaso = [];
let pontuacoesParametrosCaso = [];

/**
 * Função que calcula a pontuação de um caso, seguindo os critérios
 * definidos pela DPOP
 * @param {Arrau 2D} caso cuja pontuação será calculada
 * @returns POntuação do caso
 */
function calcularPontuacao( caso ) {

  // Familiares do caso
  let numeroFamiliares = caso.length;
  
  if( numeroFamiliares < 1 ) {    
    return 0;
  } 

  let rf = caso[0];
  let familiar = []; 


  // Pontuação
  let peso;
  let pontuacaoCriterio;
  let pontuacaoTotal = 0;

  // Array onde cada posição armazenará um objeto contendo o nome do parâmetro
  // e a pontuação obtida devido à presença desse parâmetro
  idsParametrosCaso = [];
  pontuacoesParametrosCaso = [];


  // 1) VULNERABILIDADE ASSOCIADA A CICLOS DE VIDA E PERTENCIMENTO IDENTITÁRIO
  peso = 2;


  // 1.1) Famílias com mulheres (cis ou trans)  
  pontuacaoCriterio = 0;

  if( rf[UNI_GENERO] == "3" || rf[UNI_GENERO] == "4" ) { 
    pontuacaoCriterio = peso*2;    
    idsParametrosCaso.push(1);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);
  } else {
    for( let i=1; i<numeroFamiliares; ++i ) {
      familiar = caso[i];
      if( familiar[UNI_GENERO] == "3" || familiar[UNI_GENERO] == "4" ) {
        pontuacaoCriterio = peso*1;        
        idsParametrosCaso.push(2);
        pontuacoesParametrosCaso.push(pontuacaoCriterio);
        break;
      }
    } 
  }

  pontuacaoTotal += pontuacaoCriterio;


  // 1.2) Famílias com indivíduos cuja identidade de gênero e orientação sexual sejam
  //      diferentes da cisheterossexualidade  
  pontuacaoCriterio = 0;

  if( (rf[UNI_GENERO] == "2" || rf[UNI_GENERO] == "4" || rf[UNI_GENERO] == "5" || rf[UNI_GENERO] == "6" ) || 
      (rf[UNI_ORIENTACAO_SEXUAL] == "1" || rf[UNI_ORIENTACAO_SEXUAL] == "2" || rf[UNI_ORIENTACAO_SEXUAL] == "4" || rf[UNI_ORIENTACAO_SEXUAL] == "5") ) {     
    pontuacaoCriterio = peso*2;    
    idsParametrosCaso.push(3);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);
  } else {
    for( let i=1; i<numeroFamiliares; ++i ) {
      familiar = caso[i];
      if( (familiar[UNI_GENERO] == "2" || familiar[UNI_GENERO] == "4" || familiar[UNI_GENERO] == "5" || familiar[UNI_GENERO] == "6") || 
          (familiar[UNI_ORIENTACAO_SEXUAL] == "1" || familiar[UNI_ORIENTACAO_SEXUAL] == "2" || familiar[UNI_ORIENTACAO_SEXUAL] == "4" || familiar[UNI_ORIENTACAO_SEXUAL] == "5")  ) {
        pontuacaoCriterio = peso*1;        
        idsParametrosCaso.push(4);
        pontuacoesParametrosCaso.push(pontuacaoCriterio);
        break;
      }
    } 
  }

  pontuacaoTotal += pontuacaoCriterio;


  // 1.3) Famílias com indivíduos pretos, pardos ou indígenas  
  pontuacaoCriterio = 0;

  if( rf[UNI_RACA_COR] == "2" || rf[UNI_RACA_COR] == "4" || rf[UNI_RACA_COR] == "5" ) { 
    pontuacaoCriterio = peso*2;    
    idsParametrosCaso.push(5);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);
  } else {
    for( let i=1; i<numeroFamiliares; ++i ) {
      familiar = caso[i];
      if( familiar[UNI_RACA_COR] == "2" || familiar[UNI_RACA_COR] == "4" || familiar[UNI_RACA_COR] == "5" ) {
        pontuacaoCriterio = peso*1;        
        idsParametrosCaso.push(6);
        pontuacoesParametrosCaso.push(pontuacaoCriterio);
        break;
      }
    } 
  }

  pontuacaoTotal += pontuacaoCriterio;  


  // 1.4) Famílias com crianças e adolescentes (pontuar por criança e adolescente)
  //      compondo o núcleo familiar  
  pontuacaoCriterio = 0;

  for( let i=1; i<numeroFamiliares; ++i ) {
    familiar = caso[i];
    if( calcularIdade( familiar[UNI_DATA_NASCIMENTO] ) < 18 ) {
      pontuacaoCriterio += peso*1;      
    }
  }   

  if( pontuacaoCriterio > 0 ) {
    idsParametrosCaso.push(7);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);
  } 

  pontuacaoTotal += pontuacaoCriterio;  


  // 1.5) Idosos (acima de 60 anos, com maior prioridade para idosos com 80 anos ou
  // mais, conforme Estatuto do Idoso)  
  pontuacaoCriterio = 0;

  let idadeRF = calcularIdade( rf[UNI_DATA_NASCIMENTO] );
  let idadeFamiliar;
  if( idadeRF >= 80 ) { 
    pontuacaoCriterio = peso*4;    
    idsParametrosCaso.push(8);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  } else {
    for( let i=1; i<numeroFamiliares; ++i ) {
      familiar = caso[i];
      idadeFamiliar = calcularIdade( familiar[UNI_DATA_NASCIMENTO] );
      if( idadeFamiliar >= 80 ) {
        pontuacaoCriterio = peso*3;        
        idsParametrosCaso.push(9);
        pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
        break;
      }
    } 
  }
  if( pontuacaoCriterio == 0 ) {
    if( idadeRF >= 60 ) { 
      pontuacaoCriterio = peso*2;      
      idsParametrosCaso.push(10);
      pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
    } else {
      for( let i=1; i<numeroFamiliares; ++i ) {
        familiar = caso[i];
        idadeFamiliar = calcularIdade( familiar[UNI_DATA_NASCIMENTO] );
        if( idadeFamiliar >= 60 ) {
          pontuacaoCriterio = peso*1;          
          idsParametrosCaso.push(11);
          pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
          break;
        }
      } 
    }    
  }

  pontuacaoTotal += pontuacaoCriterio;  

  
  // 1.6) Famílias com pessoas com deficiência  
  pontuacaoCriterio = 0;

  if( rf[UNI_PCD] == "2" ) { 
    pontuacaoCriterio = peso*2;    
    idsParametrosCaso.push(12);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  } else {
    for( let i=1; i<numeroFamiliares; ++i ) {
      familiar = caso[i];
      if( familiar[UNI_PCD] == "2" ) {
        pontuacaoCriterio = peso*1;        
        idsParametrosCaso.push(13);
        pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
        break;
      }
    } 
  }

  pontuacaoTotal += pontuacaoCriterio;  


  // 1.7) Famílias com mulheres em gestação ou em fase de 
  // puerpério – até 6 meses após a gestação.  
  pontuacaoCriterio = 0;

  if( rf[UNI_GESTANTE] == "2" ) { 
    pontuacaoCriterio = peso*2;    
    idsParametrosCaso.push(14);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  } else {
    for( let i=1; i<numeroFamiliares; ++i ) {
      familiar = caso[i];
      if( familiar[UNI_GESTANTE] == "2" ) {
        pontuacaoCriterio = peso*1;                
        idsParametrosCaso.push(15);
        pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
        break;
      }
    } 
  }

  pontuacaoTotal += pontuacaoCriterio;  



  // 2) VULNERABILIDADE RELACIONADA À VIOLAÇÃO DE DIREITOS
  peso = 1;


  // 2.1) Famílias em situação de rua com crianças e/ou adolescentes com
  //      medida protetiva de acolhimento ou em processo de acolhimento, e
  //      famílias que tiveram crianças e/ou adolescentes privadas do convívio
  //      familiar em decorrência da vida nas ruas dos responsáveis  
  pontuacaoCriterio = 0;
  
  if( rf[UNI_CEA_ACOLHIMENTO_INSTITUCIONAL] == "2" ||
      rf[UNI_CEA_PRIVACAO_CONVIVIO] == "2"   ) {
    pontuacaoCriterio = peso*2;    
    idsParametrosCaso.push(16);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  }

  pontuacaoTotal += pontuacaoCriterio;


  // 2.2) Família em situação de rua com membros em situação de trabalho infantil
  //      e/ou exploração sexual de crianças/adolescentes  
  pontuacaoCriterio = 0;
  
  if( rf[UNI_CEA_TRABALHO_INFANTIL_EXPLORACAO_SEXUAL] == "2" ) {
    pontuacaoCriterio = peso*6;    
    idsParametrosCaso.push(17);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  }

  pontuacaoTotal += pontuacaoCriterio;  


  // 2.3) Família em situação de rua com membros em situação de prostituição  
  pontuacaoCriterio = 0;
  
  if( rf[UNI_PROSTITUICAO] == "2" ) {
    pontuacaoCriterio = peso*1;    
    idsParametrosCaso.push(18);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  }

  pontuacaoTotal += pontuacaoCriterio;  



  // 3) VULNERABILIDADE DE SAÚDE DA FAMÍLIA  
  peso = 1;


  // 3.1) Presença de condições de saúde que necessitem de cuidado contínuo  
  pontuacaoCriterio = 0;

  let flagParametroSaude;

  if( rf[UNI_PROBLEMAS_SAUDE] == "3" ) { 

    pontuacaoCriterio = peso*3;    
    
    flagParametroSaude = 19;
    idsParametrosCaso.push(flagParametroSaude);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             

  } else if( rf[UNI_PROBLEMAS_SAUDE] == "2" ) { 

    pontuacaoCriterio = peso*2;    

    flagParametroSaude = 21;

    for( let i=1; i<numeroFamiliares; ++i ) {
      familiar = caso[i];
      if( familiar[UNI_PROBLEMAS_SAUDE] == "3" ) {
        pontuacaoCriterio += peso*1;
        flagParametroSaude = 20;
        
        break;
      }
    } 

    idsParametrosCaso.push(flagParametroSaude);          
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                                   
    
  } else {

    flagParametroSaude = 0;

    for( let i=1; i<numeroFamiliares; ++i ) {
      familiar = caso[i];
      if( familiar[UNI_PROBLEMAS_SAUDE] == "3" ) {
        pontuacaoCriterio = peso*2;        
        flagParametroSaude = 22;
        break;
      } else if( familiar[UNI_PROBLEMAS_SAUDE] == "2" ) {
        pontuacaoCriterio = peso*1;        
        flagParametroSaude = 23;      
      }
    }
    
    if(flagParametroSaude != 0) {
      idsParametrosCaso.push(flagParametroSaude);      
      pontuacoesParametrosCaso.push(pontuacaoCriterio);
    } 

  }  

  pontuacaoTotal += pontuacaoCriterio;  


  // 3.2) Famílias com pessoas que possuem diagnóstico de sofrimento mental, uso prejudicial
  //      de álcool e drogas e/ou membros dependentes de cuidados para a vida diária  
  pontuacaoCriterio = 0;

  if( rf[UNI_DIAGNOSTICO] == "2" ) {
    pontuacaoCriterio = peso*1;    
    idsParametrosCaso.push(24);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  }

  pontuacaoTotal += pontuacaoCriterio;  



  // 4) VULNERABILIDADE EM DECORRÊNCIA DE VIDA NAS RUAS
  peso = 1;


  // 4.1) Tempo de vida nas ruas e/ou de acolhimento institucional  
  pontuacaoCriterio = 0;
  
  switch( rf[UNI_TEMPO_SITUACAO_DE_RUA] ) {
    case "1":   pontuacaoCriterio = peso*2;
                idsParametrosCaso.push(25);
                pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
                break;    
    case "2":   pontuacaoCriterio = peso*2;
                idsParametrosCaso.push(26);
                pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
                break;    
    case "3":   pontuacaoCriterio = peso*2;
                idsParametrosCaso.push(27);
                pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
                break;    
    case "4":   pontuacaoCriterio = peso*4;
                idsParametrosCaso.push(28);
                pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
                break;    
    case "5":   pontuacaoCriterio = peso*5;
                idsParametrosCaso.push(29);
                pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
                break;    
    case "6":   pontuacaoCriterio = peso*6;
                idsParametrosCaso.push(30);
                pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
                break;
    default:    pontuacaoCriterio = 0;
                break;
  }  

  pontuacaoTotal += pontuacaoCriterio;    
  

  // 4.2) Histórico de institucionalização (Clínicas psiquiátricas, comunidades terapêuticas, 
  //      sistema prisional, sistema socioeducativo, instituições asilares)  
  pontuacaoCriterio = 0;

  if( rf[UNI_INSTITUCIONALIZACAO].includes("2") || 
      rf[UNI_INSTITUCIONALIZACAO].includes("3") || 
      rf[UNI_INSTITUCIONALIZACAO].includes("4") ||      
      rf[UNI_INSTITUCIONALIZACAO].includes("5") || 
      rf[UNI_INSTITUCIONALIZACAO].includes("6") ) {
    pontuacaoCriterio = peso*1;    
    idsParametrosCaso.push(31);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  }

  pontuacaoTotal += pontuacaoCriterio; 
  

  // 4.3) Famílias ou indivíduos em situação de ameaça, conflito 
  //      territorial ou outras violências  
  pontuacaoCriterio = 0;

  if( rf[UNI_AMEACA_CONFLITO_VIOLENCIA].includes("2") ) {
    pontuacaoCriterio = peso*2;    
    idsParametrosCaso.push(32);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  } else if( rf[UNI_AMEACA_CONFLITO_VIOLENCIA].includes("3") ||
             rf[UNI_AMEACA_CONFLITO_VIOLENCIA].includes("4")   ) {
    pontuacaoCriterio = peso*1;    
    idsParametrosCaso.push(33);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  }

  pontuacaoTotal += pontuacaoCriterio;  



  // 5) ARTICULAÇÃO ENTRE MORADIA E TRABALHO
  peso = 1;

  
  // 5.1) Pessoas em situação vinculadas ao trabalho formal ou informal, 
  //      grupos produtivos, associações e cooperativas  
  pontuacaoCriterio = 0;

  if( rf[UNI_TRABALHO_FORMAL] == "2" || 
      rf[UNI_TRABALHO_OUTROS].includes("2") || 
      rf[UNI_TRABALHO_OUTROS].includes("3") ) { 
    pontuacaoCriterio = peso*2;    
    idsParametrosCaso.push(34);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  } else {
    for( let i=1; i<numeroFamiliares; ++i ) {
      familiar = caso[i];
      if( familiar[UNI_TRABALHO_FORMAL] == "2" || 
          familiar[UNI_TRABALHO_OUTROS].includes("2") || 
          familiar[UNI_TRABALHO_OUTROS].includes("3") ) {
        pontuacaoCriterio = peso*1;        
        idsParametrosCaso.push(35);
        pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
        break;
      }
    } 
  }

  pontuacaoTotal += pontuacaoCriterio;


  // 5.2) Participante do Programa Estamos Juntos, inseridos 
  //      nas frentes de trabalho ou oportunidade CLT  
  pontuacaoCriterio = 0;

  if( rf[UNI_ESTAMOS_JUNTOS] == "2" ) { 
    pontuacaoCriterio = peso*2;    
    idsParametrosCaso.push(36);
    pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
  } else {
    for( let i=1; i<numeroFamiliares; ++i ) {
      familiar = caso[i];
      if( familiar[UNI_ESTAMOS_JUNTOS] == "2" ) {
        pontuacaoCriterio = peso*1;        
        idsParametrosCaso.push(37);
        pontuacoesParametrosCaso.push(pontuacaoCriterio);                             
        break;
      }
    } 
  }

  pontuacaoTotal += pontuacaoCriterio;  



  // Retorna a pontuação total  
  return pontuacaoTotal;

} // Fim da função calcularPontuacao



function testeCalcularPontuacaoCaso() {

  let caso = obterCaso( BUFFER_CASOS_EXTERNOS );

  let pontuacaoCaso = calcularPontuacao( caso );

  console.log( "\n\nPontuação caso: " + pontuacaoCaso );
}



/**
 * Função que determina o número de criançãs e adolescentes de um caso.
 * @param {Array 2D} caso 
 * @returns Número de C&A
 */
function numeroDeCEAs( caso ) {

  // Número de crianças e adolescentes na família
  let numCEA = 0;

  // Familiares do caso
  let numeroFamiliares = caso.length;
  
  if( numeroFamiliares < 1 ) {    
    return 0;
  } 

  let familiar = []; 

  for( let i=1; i<numeroFamiliares; ++i ) {
    familiar = caso[i];
    if( calcularIdade( familiar[UNI_DATA_NASCIMENTO] ) < 18 ) {
      ++numCEA;
    }
  } 
    
  return numCEA;

} // Fim da função numeroDeCEAs



/**
 * Função que determina o número problemas de saúde de um caso.
 * @param {Array 2D} caso 
 * @returns Número de problemas de saúde
 */
function numeroDeProblemasDeSaude( caso ) {

  // Familiares do caso
  let numeroFamiliares = caso.length;
  
  if( numeroFamiliares < 1 ) {    
    return 0;
  } 

  // Número de problemas de saúde na família  
  let numProbSaude = caso[0][UNI_PROBLEMAS_SAUDE] != "" ? parseInt(caso[0][UNI_PROBLEMAS_SAUDE]) - 1 : 0;

  let familiar = []; 
  for( let i=1; i<numeroFamiliares; ++i ) {
    familiar = caso[i];
    numProbSaude += familiar[UNI_PROBLEMAS_SAUDE] != "" ? parseInt(familiar[UNI_PROBLEMAS_SAUDE]) - 1 : 0;    
  } 
    
  return numProbSaude;

} // Fim da função numeroDeProblemasDeSaude








/**
 * ##### FIM DO MÓDULO tabelasEncaminhamentos.gs #####
 */