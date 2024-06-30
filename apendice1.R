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

# Usar o elemento "a" e o atributo "href" para obter vetor com links dentro do elemento
links <- boletim_anp |> 
  html_elements(".conteudo") |> 
  html_elements("a") |> 
  html_attr("href") 

# A partir do vetor "links" geramos o vetor com o nome dos arquivos

# O dowloand de múltiplos arquivos exige a criação de uma função

download_arquivos <- function(x,y){download.file(url = x, destfile =y, mode = "wb")}

# Essa função foi envolvida pelas funções slowly e insistently, para evitar erros

download_lento <- slowly(download_arquivos, rate_delay(5))

download_insistente <- insistently(download_lento, rate = rate_backoff(60))

# download de todos os arquivos

map2(links, nomes_arquivos, download_lento)

# Criar vetor com nome dos arquivos
planilhas <- dir(path = "data", pattern = "*xls*")

# Nos meses seguintes, antes de fazer um download de todos os arquivos, fazemos 
# o webscrapping e verificamos se há algum arquivo novo

setdiff(planilhas, nomes_arquivos)

## 2 - Importação das planilhas


# 2.1 - Obter nomes de todas as guias
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

# A primeira linha importa o conteúdo da guia 3 de todas as planilhas
  map_df(planilhas, read_excel_data, sheet = 3) |>
  # A segunda linha filtra todos os títulos das tabelas presentes nessas guias
    filter(if_any(2, ~ str_detect(., "Tabela"))) |> 
    # E assim mostramos apenas os títulos diferentes
    distinct() |> 
    print(n = 25)
  
# Com um pouco de limpeza dos dados conseguimos saber os nomes das guias das
# planilhas. Esse procedimento deve ser repetido para as outras abas, e servirá
# de guia quando importarmos cada tabela

# 2.2 - Importar a data em que cada planilha foi gerada

# Função personalizada para extrair a data do arquivo, a partir da primeira
# tabela da guia 2
  
extrair_data_ws2 <- function(num){
    
    prov1 <- read_excel(paste0("data/", planilhas[num]), sheet = 2, .name_repair = "unique_quiet")
    
    names(prov1)[1] <- "col1"
    
    corte <- which(str_detect(prov1$col1, "Tabela [0-9]*"))
    
    maximo_linha <- corte - lag(corte)
    
    # Extrair tabela 1, na ws2
    df1 <- read_excel(paste0("data/", planilhas[num]), sheet = 2, ,skip = corte[2] + 2, n_max = maximo_linha[2] -2, .name_repair = "unique_quiet")
    
    data_ws1 <- names(df1)[length(df1)]
    
    data_ws1 <- lubridate::my(data_ws1)
    data_ws1
  }

# Extraindo data do elemento planilhas[2]
extrair_data_ws2(2)

# 2.3 - Importar todos os nomes de tabelas da guia 2, em todos os arquivos

map_df(planilhas, read_excel_data, sheet = 2) |> 
  filter(if_any(2, ~ str_detect(., "Tabela [0-9]+"))) |> 
  distinct() 

# Esse procedimento precisa ser repetido para as outras guias. Técnicas triviais 
# de transformação de dados, junto com a consulta aos dados originais, precisam 
# ser realizadas para determinar o conteúdo de cada tabela. Pelos nomes é 
# possível verificar que a Tabelas 3 e a Tabela 4 têm conteúdos diferentes,
# dependendo do arquivo do qual são extraídos

# 2.4 Importar múltiplas tabelas de uma mesma guia

# O conteúdo de cada aba podia ser importado e posteriormente limpo e
# transformado. No entanto, isso se revelou muito complexo pela quantidade de
# elementos importados e diversidade de número de linhas e colunas e classes
# de caracteres
    
# Função personalizada para extrair três tabelas, da guia 3

extrair_dados_ws3 <- function(num, ajuste = 2, ajuste2 = -2, sheet = 3){

  # A função que importa a data será usada na planilha
  data_planilha <-  extrair_data_ws2(num)  
  
  # planilha 3 é importada
  prov1 <- read_excel(paste0("data/", planilhas[num]), sheet = sheet)
  names(prov1)[1] <- "col1"
  # a partir da palavra "Tabela" é identificada a primeira linha da tabela
  corte <- which(str_detect(prov1$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o inínio da tabela seguinte
  maximo_linha <- corte - lag(corte)
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres
  df1 <- read_excel(paste0("data/", planilhas[num]), sheet = sheet, skip = corte[1] + ajuste, n_max = maximo_linha[2] + ajuste2) |> 
    mutate(data = data_planilha) 
  
  df2 <- read_excel(paste0("data/", planilhas[num]), sheet = sheet, skip = corte[1 + 1] + ajuste, n_max = maximo_linha[3] + ajuste2) |> 
    mutate(data = data_planilha)
  sheet

  df3 <- read_excel(paste0("data/", planilhas[num]), sheet = sheet, skip = corte[1 + 2] + ajuste, n_max = maximo_linha[4] + ajuste2) |> 
    mutate(data = data_planilha)
  
  
  list(df1, df2, df3)
  
}

# Importar as três primeiras tabelas da guia 3, do arquivo planilhas[1]: 
planilhas[1]
extrair_dados_ws3(2)

# Importar as três primeiras tabelas da guia 3, de todos os arquivos descritos 
# em planilhas:
worksheet3_trestabelas <- list()

i <- 1
while (i < length(planilhas) + 1) {
  worksheet3_trestabelas[[i]] <- extrair_dados_ws3(num = i)  
  i <- i + 1
}
worksheet3_trestabelas

# A primeira tabela da guia 3 tem o mesmo conteúdo em todos os arquivos 
# (Distribuição da Produção de Petróleo e Gás Natural por Estado), portanto o
# primeiro elemento de todas as listas pode ser unido em um único dataframe. 
# Essa análise precisa ser feita para as outras tabelas

do.call(bind_rows, lapply(worksheet3_trestabelas, `[[`, 1))

# Com a função bind_rows os elementos são unidos mesmo tendo nomes de colunas
# diferentes. Isso pode ser corrigido:

nomes_colunas <- colnames(as.data.frame(worksheet3_trestabelas[[2]][1]))

i <- 1
while (i < length(planilhas) + 1) {
  colnames(worksheet3_trestabelas[[i]][[1]]) <- nomes_colunas  
  i <- i + 1
}

# Estando todos os nomes de colunas iguais, a função rbind pode ser usada:
producao_por_estado <- do.call(rbind, lapply(worksheet3_trestabelas, `[[`, 1))

# Com isso, temos todos os elementos da primeira tabela em um único elemento
# Uma última etapa: fazer essa tabela obedecer aos princípios da "tidy data"

producao_por_estado |> 
  gather(key = "Tipo", value = "Produção", -c(Estado, data)) |> 
  print(n= 30)

# Os procedimentos do tópico 2.2 podem agora ser repetidos para importar cada
# tabela. A importação de cada tabela e posterior união com as de outras 
# planilhas deve considerar o conteúdo de cada tabelas, o que é descoberto 
# com as ferramentas descritas na seção 2.1 e 2.2 desse script
