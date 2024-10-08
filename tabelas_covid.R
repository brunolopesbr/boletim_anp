## ----setup, include=FALSE-----------------------------------------------------------------------------------
knitr::opts_chunk$set(echo = TRUE)


## -----------------------------------------------------------------------------------------------------------
library(tidyverse)
library(readxl)


## ----warning=FALSE------------------------------------------------------------------------------------------
# Rodar o arquivo worksheets e data 

source(knitr::purl("worksheets_data.Rmd", quiet=TRUE))



## -----------------------------------------------------------------------------------------------------------
read_excel_data  <- function(x,sheet = 2, range = "A1:J1000", skip = 0, n_max = Inf)
  {read_xlsx(paste0("data/", x), 
            col_types = "text",
            sheet = sheet,
            range = range,
            skip = skip,
            col_names = FALSE,
            n_max = n_max,
            .name_repair = "unique_quiet",)}



## -----------------------------------------------------------------------------------------------------------
extrair_tabelas_covid <- function(num, ajuste = 2, ajuste2 = -2, sheet = 5){

  # A função que importa a data será usada na planilha
  data_planilha <-  datas_arquivos$value[num]  
  
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
  
  df2 <- read_excel(paste0("data/", planilhas[num]), sheet = sheet, skip = corte[1 + 1] + ajuste) |> 
    mutate(data = data_planilha)
  sheet

  list(df1, df2)
  
}


## -----------------------------------------------------------------------------------------------------------
# planilhas
guias_covid


## -----------------------------------------------------------------------------------------------------------
map_df(planilhas[guias_covid], read_excel_data, sheet = 5) |> 
  filter(if_any(2, ~ str_detect(., "Tabela [0-9]+"))) |> 
  distinct() |> 
  # apagar colunas sem dados
  select_if(~all(!is.na(.)))


## -----------------------------------------------------------------------------------------------------------
extrair_tabelas_covid(31)


## -----------------------------------------------------------------------------------------------------------
i <- guias_covid[1]
final <- guias_covid[length(guias_covid)-1]

tabelas_covid <- list()
while (i < (final + 1)) {
  tabelas_covid[[i]] <- extrair_tabelas_covid(num = i)  
  i <- i + 1
}

 tabelas_covid[[39]] <- extrair_tabelas_covid(guias_covid[length(guias_covid)]) 
 
 tabelas_covid


## -----------------------------------------------------------------------------------------------------------

tabela_19 <- do.call(rbind, lapply(tabelas_covid, `[[`, 1))


## -----------------------------------------------------------------------------------------------------------
tabela_20 <- do.call(bind_rows, lapply(tabelas_covid, `[[`, 2))

