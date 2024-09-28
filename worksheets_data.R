## ----setup, include=FALSE-----------------------------------------------------------------------------------
knitr::opts_chunk$set(echo = TRUE) 


## -----------------------------------------------------------------------------------------------------------
library(tidyverse)
library(readxl)


## -----------------------------------------------------------------------------------------------------------
planilhas <- dir(path = "data", pattern = "*xls*")


## -----------------------------------------------------------------------------------------------------------
# 2.1 - Obter nomes de todas as guias
extrair_guias <- function(num){
  readxl::excel_sheets(paste0("data/", planilhas[num]))}


## -----------------------------------------------------------------------------------------------------------
extrair_data_ws2 <- function(num){

prov1 <- read_excel(paste0("data/", planilhas[num]), sheet = 2, .name_repair = "unique_quiet")

names(prov1)[1] <- "col1"

corte <- which(str_detect(prov1$col1, "Tabela [0-9]*"))

maximo_linha <- corte - lag(corte)

# Extrair tabela 1, na ws2
df1 <- read_excel(paste0("data/", planilhas[num]), sheet = 2, ,skip = corte[2] + 2, n_max = maximo_linha[2] -2, .name_repair = "unique_quiet")

data_ws1 <- names(df1)[length(df1)]

{
  if (str_detect(data_ws1, "/") == TRUE) {data_ws2 <- lubridate::my(data_ws1)}
  
  else {data_ws2 <- as.Date(as.numeric(data_ws1), origin = "1899-12-30")} 
}

as_tibble(data_ws2) 
}


## ----message=FALSE, warning=FALSE---------------------------------------------------------------------------
datas_arquivos <- map_df(seq_along(planilhas), extrair_data_ws2) |> 
  mutate(planilha = planilhas)

datas_arquivos


## ----message=FALSE, warning=FALSE---------------------------------------------------------------------------
# Verificando os erros

{if (datas_arquivos |> 
  filter(is.na(value))  |> 
  nrow()  == 0 ) {print("Sem erros")}
else (print("Verificar"))  } 



## -----------------------------------------------------------------------------------------------------------
guias_lista <- map(1:length(planilhas), extrair_guias)

guias <- do.call(rbind, guias_lista) |> 
  as_tibble() 

guias <- guias |> 
  mutate(V5 = str_remove_all(V5, "Lista de Tabelas")) |> 
  mutate(num_arquivo = 1:nrow(guias) )


## -----------------------------------------------------------------------------------------------------------
guias


## -----------------------------------------------------------------------------------------------------------
unique(guias$V2)



## -----------------------------------------------------------------------------------------------------------
planilhas_ws4_grupo1 <- guias |> 
  filter(V4 == unique(guias$V4)[1]) |> 
  pull(num_arquivo)

planilhas_ws4_grupo2 <- guias |> 
  filter(V4 == unique(guias$V4)[2]) |> 
  pull(num_arquivo)



## -----------------------------------------------------------------------------------------------------------
guias |> 
  filter(V5 == unique(guias$V5)[2]) |> 
  pull(num_arquivo)



## -----------------------------------------------------------------------------------------------------------
planilhas_ws4_grupo1 <- guias |> 
  filter(V5 == unique(guias$V5[1])) |> 
  pull(num_arquivo)

planilhas_ws4_grupo2 <- guias |> 
  filter(V4 == unique(guias$V4[2]))



## -----------------------------------------------------------------------------------------------------------
unique(guias$V2) |> length()
unique(guias$V3) |> length()
unique(guias$V4) |> length()
unique(guias$V5) |> length()

unique(guias$V4)
unique(guias$V5)


## -----------------------------------------------------------------------------------------------------------
unique(guias$V4)[1]
unique(guias$V4)[2]

guias_4_ranking <- guias |> 
  filter(str_detect(V4, unique(guias$V4)[1])) |> 
  pull(num_arquivo)

guias_4_ranking <- guias |> 
  filter(str_detect(V4, unique(guias$V4)[2]))



## -----------------------------------------------------------------------------------------------------------
guias_covid <- guias |> 
   filter(str_detect(V5, "Covid")) |> 
  pull(num_arquivo)


