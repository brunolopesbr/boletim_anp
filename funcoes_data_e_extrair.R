## ----setup, include=FALSE--------------------------------------------------------------------------------------------------------------------------------------------------------------------
knitr::opts_chunk$set(echo = TRUE)


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
library(tidyverse)
library(readxl)


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
planilhas <- dir(path = "data", pattern = "*xls*")


## ----eval=FALSE, include=FALSE---------------------------------------------------------------------------------------------------------------------------------------------------------------
## # planilhas_novas <- c("2024_tabela-abril.xlsm", "2024_tabela-maio.xlsm", "2024_tabela-junho.xlsm")
## 
## 
## 
## # planilhas <- try(planilhas[!planilhas %in% planilhas_novas])


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# 2.1 - Obter nomes de todas as guias
extrair_guias <- function(num){
  readxl::excel_sheets(paste0("data/", planilhas[num]))}


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
planilhas[1]
x2017_10 <- read_excel(paste0("data/", planilhas[1]), 
    sheet = 2, 
    range = "O8", col_names = FALSE)
x2017_10 <- pull(x2017_10) 


{
  if (str_detect(x2017_10, "/") == TRUE) {x2017_10 <- lubridate::my(x2017_10)}
  
  else {x2017_10 <- as.Date(as.numeric(x2017_10), origin = "1899-12-30")} 
}

x2017_10

## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
extrair_data_novo <- function(x){
    mes_planilha <- read_excel(paste0("data/", planilhas[x]), 
    sheet = 2, 
    range = "O8", col_names = FALSE, col_types = "text")
  
mes_planilha <- pull(mes_planilha) 

{
  if (str_detect(mes_planilha, "/") == TRUE) {mes_planilha_f <- lubridate::my(mes_planilha)}
  
  else {mes_planilha_f <- as.Date(as.numeric(mes_planilha), origin = "1899-12-30")} 
}

mes_planilha_f <- as.character(mes_planilha_f)
mes_planilha_f
}


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
extrair_data_novo(1) |> class()


## ----eval=FALSE, include=FALSE---------------------------------------------------------------------------------------------------------------------------------------------------------------
## i <- 1
## for(i in seq_along(planilhas)){
##   extrair_data_novo(i) |> print()
##   i <- 1 + i
## }
## 


## ----warning=FALSE---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

datas_boletins <- data.frame("boletim" = NA, "mês" = NA)

colunas_datas <- as.Date(NA)
i <- 1
for(i in seq_along(planilhas)){

datas_boletins[i,1] <- planilhas[i]
    
  mes_planilha <- read_excel(paste0("data/", planilhas[i]), 
    sheet = 2, 
    range = "O8", col_names = FALSE, col_types = "text")
  
mes_planilha <- pull(mes_planilha) 

{
  if (str_detect(mes_planilha, "/") == TRUE) {mes_planilha_f <- lubridate::my(mes_planilha)}
  
  else {mes_planilha_f <- as.Date(as.numeric(mes_planilha), origin = "1899-12-30")} 
}

colunas_datas[i] <- mes_planilha_f
    
    i <- 1 + i
}
datas_boletins[,2] <- colunas_datas




## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
datas_boletins <- data.frame("boletim" = NA, "mês" = NA)

colunas_datas <- as.Date(NA)
i <- 1
for(i in seq_along(planilhas)){

datas_boletins[i,1] <- planilhas[i]
    
colunas_datas[i] <- extrair_data_novo(i)
    
    i <- 1 + i
}
datas_boletins[,2] <- colunas_datas


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
colunas_datas
datas_boletins


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
guias_lista <- map(1:length(planilhas), extrair_guias)

guias <- do.call(rbind, guias_lista) |> 
  as_tibble() 

guias <- guias |> 
  mutate(V5 = str_remove_all(V5, "Lista de Tabelas")) |> 
  mutate(num_arquivo = 1:nrow(guias) )


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
guias


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
unique(guias$V2)



## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
planilhas_ws4_grupo1 <- guias |> 
  filter(V4 == unique(guias$V4)[1]) |> 
  pull(num_arquivo)

planilhas_ws4_grupo2 <- guias |> 
  filter(V4 == unique(guias$V4)[2]) |> 
  pull(num_arquivo)



## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
guias |> 
  filter(V5 == unique(guias$V5)[2]) |> 
  pull(num_arquivo)



## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
planilhas_ws4_grupo1 <- guias |> 
  filter(V5 == unique(guias$V5[1])) |> 
  pull(num_arquivo)

planilhas_ws4_grupo2 <- guias |> 
  filter(V4 == unique(guias$V4[2]))



## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
unique(guias$V2) |> length()
unique(guias$V3) |> length()
unique(guias$V4) |> length()
unique(guias$V5) |> length()

unique(guias$V4)
unique(guias$V5)


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
guias_covid <- guias |> 
   filter(str_detect(V5, "Covid")) |> 
  pull(num_arquivo)



## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
extrair_todas_tabelas <- function(nume, sheet = sheet){
  mes_boletim <- extrair_data_novo(nume) 
  
  prov <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet)
  
  names(prov)[1] <- "col1"
  
  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
  # tamanho da última tabela: número total de linhas do DF - linha da última tabela do DF
  # substituir esse valor no vetor, que aparece como "NA"
  ultima_n_max <- nrow(prov) - corte[length(corte)]
  
  maximo_linha[length(maximo_linha)] <- ultima_n_max + 2
  
  ajuste <- 2
  ajuste2 <- -2
  
  maximo_linha
  
  minha_lista <- list()
  
  
  i <- 1
  for(i in seq_along(corte)){
    minha_lista[[i]] <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[i] + 
                                     ajuste, n_max = maximo_linha[i] + ajuste2) |> 
      mutate(mes = mes_boletim)
    i <- i + 1
  }
  
  minha_lista
}



## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
extrair_todas_tabelas_novo <- function(x, nume, sheet = sheet){
  mes_boletim <- extrair_data_novo(nume)
  
  prov <- read_excel(paste0("data/", x), sheet = sheet)
  
  names(prov)[1] <- "col1"
  
  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
  # tamanho da última tabela: número total de linhas do DF - linha da última tabela do DF
  # substituir esse valor no vetor, que aparece como "NA"
  ultima_n_max <- nrow(prov) - corte[length(corte)]
  
  maximo_linha[length(maximo_linha)] <- ultima_n_max + 2
  
  ajuste <- 2
  ajuste2 <- -2
  
  maximo_linha
  
  minha_lista <- list()
  
  
  i <- 1
  for(i in seq_along(corte)){
    minha_lista[[i]] <- read_excel(paste0("data/", x), sheet = sheet, skip = corte[i] + 
                                     ajuste, n_max = maximo_linha[i] + ajuste2) |> 
      mutate(mes = mes_boletim)
    i <- i + 1
  }
  
  minha_lista
}


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
library(readxl)
library(tidyverse)
planilha_maio <- "data/2024_tabela-maio.xlsm"

primeira_importacao <- read_xlsx(path = planilha_maio, sheet = 2)
names(primeira_importacao)[1] <- "col1"
corte <- which(str_detect(primeira_importacao$col1, "Tabela [0-9]*"))
maximo_linha <- lead(corte) - corte

ultima_n_max <- nrow(primeira_importacao) - corte[length(corte)]
maximo_linha[length(maximo_linha)] <- ultima_n_max + 2

primeira_linha <- corte + 2
tamanho_tabela <- maximo_linha - 2


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
lista_exemplo <- list()

i <- 1 
for(i in seq_along(primeira_linha)) {
  lista_exemplo[[i]] <- read_excel(path = planilha_maio, sheet = 2, skip = primeira_linha[i], n_max = tamanho_tabela[i])
}

lista_exemplo


## --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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


## ----message=FALSE, warning=FALSE------------------------------------------------------------------------------------------------------------------------------------------------------------
# datas_arquivos <- map_df(seq_along(planilhas), extrair_data_ws2) |> 
#  mutate(planilha = planilhas)

# datas_arquivos


## ----message=FALSE, warning=FALSE------------------------------------------------------------------------------------------------------------------------------------------------------------
# Verificando os erros

#{if (datas_arquivos |> 
#  filter(is.na(value))  |> 
#  nrow()  == 0 ) {print("Sem erros")}
#else (print("Verificar"))  } 


