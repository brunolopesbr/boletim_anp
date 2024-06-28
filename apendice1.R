# Pacotes utilizados
library(readxl)
library(tidyverse)
library(rvest)
library(robotstxt)

# Verificar se podemos fazer o webscrapping

paths_allowed(
  paths  = "/anp/pt-br/centrais-de-conteudo/publicacoes/boletins-anp/boletins/boletim-mensal-da-producao-de-petroleo-e-gas-natural", 
  domain = "www.gov.br", 
  bot    = "*"
)

# Fazer o webscrapping
url <- "https://www.gov.br/anp/pt-br/centrais-de-conteudo/publicacoes/boletins-anp/boletins/boletim-mensal-da-producao-de-petroleo-e-gas-natural"
boletim_anp <- read_html(url)

# Usar o elemento "a" e o atributo "href" para obter vetor com links dentro do arquivo
links. <- boletim_anp |> 
  html_elements(".conteudo") |> 
  html_elements("a") |> 
  html_attr("href") 

# Criar vetor com as planilhas
planilhas <- dir(path = "data", pattern = "*xls*")

# Obter todas as guias
extrair_guias <- function(num){
  readxl::excel_sheets(paste0("data/", planilhas[num]))}

map(1:length(planilhas), extrair_guias)

# Importar conteúdo de cada guia
extrair_conteudo <- function(num){
  readxl::excel_sheets(paste0("data/", planilhas[num]))}