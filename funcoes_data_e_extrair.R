## ----setup, include=FALSE------------------------------------------------------------------
knitr::opts_chunk$set(echo = TRUE)


## ------------------------------------------------------------------------------------------
library(tidyverse)
library(readxl)
  

## ------------------------------------------------------------------------------------------
planilhas <- dir(path = "data", pattern = "*xls*")


## ----eval=FALSE, include=FALSE-------------------------------------------------------------
## # planilhas_novas <- c("2024_tabela-abril.xlsm", "2024_tabela-maio.xlsm", "2024_tabela-junho.xlsm")  
## 
## 
## 
## # planilhas <- try(planilhas[!planilhas %in% planilhas_novas])


## ------------------------------------------------------------------------------------------
# 2.1 - Obter nomes de todas as guias
extrair_guias <- function(num){
  readxl::excel_sheets(paste0("data/", planilhas[num]))}


## ------------------------------------------------------------------------------------------
importar_periodo <- function(x){
    df <- read_xlsx(paste0("data/", planilhas[x]), 
    sheet = 2, range = "O8", col_names = FALSE,
    .name_repair = "unique_quiet")
  
valor <- pull(df) 

{
  if (class(valor)[1] == "character") {valor <- my(valor)}
  if (class(valor)[1] == "POSIXct") {valor <- as.Date(valor)}
}

valor
}


## ------------------------------------------------------------------------------------------
importar_periodo(1) 
importar_periodo(78)
importar_periodo(9) 

## ------------------------------------------------------------------------------------------
extrair_data_novo <- importar_periodo


## ----eval=FALSE, include=FALSE-------------------------------------------------------------
## importar_limites_tabelas <- function(nume, sheet = sheet){
## 
##   limite_tabelas <- data.frame()
## 
##   prov <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet, .name_repair = "unique_quiet")
##   names(prov)[1] <- "col1"
## 
##   linha_inicial <- which(str_detect(prov$col1, "Tabela [0-9]*")) + 2
##   # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
##   tamanho_tabela <- lead(linha_inicial) - linha_inicial - 4
## 
##   limite_tabelas <- data.frame(linha_inicial, tamanho_tabela)
## 
##    # Definir linha da última tabela da guia
##   limite_tabelas$tamanho_tabela[nrow(limite_tabelas)] <- nrow(prov) + 1  - limite_tabelas[nrow(limite_tabelas),1]
## 
##   limite_tabelas
## }


## ------------------------------------------------------------------------------------------
# planilhas[5]
# importar_limites_tabelas(nume = 5, sheet = 3)


## ------------------------------------------------------------------------------------------
importar_tabela_limpa <- function(nume, sheet = sheet){
  prov <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet, .name_repair = "unique_quiet")
  
  names(prov)[1] <- "col1"
  
  linha_inicial <- which(str_detect(prov$col1, "Tabela [0-9]*")) + 2
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  tamanho_tabela <- lead(linha_inicial) - linha_inicial - 2
  
  limite_tabelas <- data.frame(linha_inicial, tamanho_tabela)
  
   # Definir linha da última tabela da guia
  limite_tabelas$tamanho_tabela[nrow(limite_tabelas)] <- nrow(prov) + 1  - limite_tabelas[nrow(limite_tabelas),1]
  
  minha_lista <- list()
  
    if(sheet != 2) {
    mes_boletim <- importar_periodo(nume) 
    
    i <- 1
  for(i in 1:nrow(limite_tabelas)){
    minha_lista[[i]] <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = limite_tabelas$linha_inicial[i], n_max = limite_tabelas$tamanho_tabela[i], .name_repair = "unique_quiet") |> 
      mutate(mes = mes_boletim)
    i <- i + 1
  }
    
    }
  
      if(sheet == 2) {

    i <- 1
  for(i in 1:nrow(limite_tabelas)){
    minha_lista[[i]] <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = limite_tabelas$linha_inicial[i], n_max = limite_tabelas$tamanho_tabela[i], .name_repair = "unique_quiet") 
    i <- i + 1
  }
    
    }
  
  minha_lista
}




## ------------------------------------------------------------------------------------------
importar_tabela_limpa(5, sheet = 2)


## ------------------------------------------------------------------------------------------
importar_tabela_limpa(15, sheet = 3)


## ------------------------------------------------------------------------------------------
guias_lista <- map(1:length(planilhas), extrair_guias)

guias <- do.call(rbind, guias_lista) |> 
  as_tibble() 

guias <- guias |> 
  mutate(V5 = str_remove_all(V5, "Lista de Tabelas")) |> 
  mutate(num_arquivo = 1:nrow(guias) )


## ------------------------------------------------------------------------------------------
guias


## ------------------------------------------------------------------------------------------
unique(guias[,1:5])


## ------------------------------------------------------------------------------------------
guias_covid <- guias |> 
   filter(str_detect(V5, "Covid")) |> 
  pull(num_arquivo)



## ------------------------------------------------------------------------------------------
contar_tabelas <- function(nume, ajuste = 2, ajuste2 = -2, sheet = sheet){
prov <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet)

names(prov) <- "col1"

corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))

print(length(corte))
}

i <- 1
for(i in 51:60){
contar_tabelas(i, sheet = 4)  
  i <- 1 + 1
}

