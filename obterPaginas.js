"use strict";

/**
 * Módulo:    obterPaginas.gs
 * Objetivo:  Retorna um objeto com todas as páginas html do sistema
 */



/**
 * Função que retorna todas as tabelas de códigos do sistema
 * 
 * @return Um objeto com todas as tabelas de códigos do sistema
 */
function obterPaginas() {
  
  const paginaResumoDosCasos  = include( "resumosDosCasos" );
  const paginaVisualizarCaso  = include( "visualizarCaso" );
  

  const paginas = {

    paginaResumoDosCasos: paginaResumoDosCasos,
    paginaVisualizarCaso: paginaVisualizarCaso

  };

  return JSON.stringify( paginas );     

} // Fim da função obterPaginas



/**
 * ##### FIM DO MÓDULO obterPaginas.gs #####
 */