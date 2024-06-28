# Pacotes utilizados
library(readxl)
library(tidyverse)
library(rvest)
library(robotstxt)

## 1 - webscrapping e download
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
links <- boletim_anp |> 
  html_elements(".conteudo") |> 
  html_elements("a") |> 
  html_attr("href") 

# O dowloand de múltiplos arquivos exige a criação de uma função

download_arquivos <- function(x,y){download.file(url = x, destfile =y, mode = "wb")}

# Essa função foi envolvida pelas funções slowly e insistently, para evitar erros

download_lento <- slowly(download_arquivos, rate_delay(5))

download_insistente <- insistently(download_lento, rate = rate_backoff(60))

## 2 - Importação das planilhas

# Criar vetor com nome dos arquivos
planilhas <- dir(path = "data", pattern = "*xls*")

# Obter nomes de todas as guias
extrair_guias <- function(num){
  readxl::excel_sheets(paste0("data/", planilhas[num]))}

map(1:length(planilhas), extrair_guias)

# Importar conteúdo de cada guia
# Com essa função personalizada cada planilha é importada com 10 colunas e 1000 
# linhas. Essa padronização permite que elas sejam combinadas em um dataframe,
# pela função map_df

read_excel_data  <- function(x,sheet = 2, range = "A1:J1000", skip = 0, n_max = Inf)
  {read_xlsx(paste0("data/", x), 
            col_types = "text",
            sheet = sheet,
            range = range,
            skip = skip,
            col_names = FALSE,
            n_max = n_max,
            .name_repair = "unique_quiet",)}

read_excel_data(planilhas[1], sheet = 2)

# A primeira linha importa o conteúdo da guia 2 de todas as planilhas
  map_df(planilhas, read_excel_data, sheet = 2) |>
  # A segunda linha filtra todos os títulos das tabelas presentes nessas guias
    filter(if_any(2, ~ str_detect(., "Tabela"))) |> 
    # E assim mostramos apenas os títulos diferentes
    distinct()
  
library(readxl)  
read_xlsx(n_max = Inf)  
  
extrair_conteudo <- function(num){
  readxl::excel_sheets(paste0("data/", planilhas[num]))}
