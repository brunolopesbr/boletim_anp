## ----setup, include=FALSE--------------------------------------------------------------------------------------------------------
knitr::opts_chunk$set(echo = TRUE)


## --------------------------------------------------------------------------------------------------------------------------------
library(tidyverse)
library(readxl)


## --------------------------------------------------------------------------------------------------------------------------------
planilhas <- dir(path = "data", pattern = "*xls*")

planilhas_novas <- c("2024_tabela-abril.xlsm", "2024_tabela-maio.xlsm", "2024_tabela-junho.xlsm")



planilhas <- planilhas[!planilhas %in% planilhas_novas]


## --------------------------------------------------------------------------------------------------------------------------------
# 2.1 - Obter nomes de todas as guias
extrair_guias <- function(num){
  readxl::excel_sheets(paste0("data/", planilhas[num]))}


## --------------------------------------------------------------------------------------------------------------------------------
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


## ----message=FALSE, warning=FALSE------------------------------------------------------------------------------------------------
datas_arquivos <- map_df(seq_along(planilhas), extrair_data_ws2) |> 
  mutate(planilha = planilhas)

datas_arquivos


## ----message=FALSE, warning=FALSE------------------------------------------------------------------------------------------------
# Verificando os erros

{if (datas_arquivos |> 
  filter(is.na(value))  |> 
  nrow()  == 0 ) {print("Sem erros")}
else (print("Verificar"))  } 



## --------------------------------------------------------------------------------------------------------------------------------
guias_lista <- map(1:length(planilhas), extrair_guias)

guias <- do.call(rbind, guias_lista) |> 
  as_tibble() 

guias <- guias |> 
  mutate(V5 = str_remove_all(V5, "Lista de Tabelas")) |> 
  mutate(num_arquivo = 1:nrow(guias) )


## --------------------------------------------------------------------------------------------------------------------------------
guias


## --------------------------------------------------------------------------------------------------------------------------------
unique(guias$V2)



## --------------------------------------------------------------------------------------------------------------------------------
planilhas_ws4_grupo1 <- guias |> 
  filter(V4 == unique(guias$V4)[1]) |> 
  pull(num_arquivo)

planilhas_ws4_grupo2 <- guias |> 
  filter(V4 == unique(guias$V4)[2]) |> 
  pull(num_arquivo)



## --------------------------------------------------------------------------------------------------------------------------------
guias |> 
  filter(V5 == unique(guias$V5)[2]) |> 
  pull(num_arquivo)



## --------------------------------------------------------------------------------------------------------------------------------
planilhas_ws4_grupo1 <- guias |> 
  filter(V5 == unique(guias$V5[1])) |> 
  pull(num_arquivo)

planilhas_ws4_grupo2 <- guias |> 
  filter(V4 == unique(guias$V4[2]))



## --------------------------------------------------------------------------------------------------------------------------------
unique(guias$V2) |> length()
unique(guias$V3) |> length()
unique(guias$V4) |> length()
unique(guias$V5) |> length()

unique(guias$V4)
unique(guias$V5)


## --------------------------------------------------------------------------------------------------------------------------------
guias_covid <- guias |> 
   filter(str_detect(V5, "Covid")) |> 
  pull(num_arquivo)



## --------------------------------------------------------------------------------------------------------------------------------
extrair_todas_tabelas <- function(nume, sheet = sheet){
  mes_boletim <- datas_arquivos[[nume,1]]
  
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



## --------------------------------------------------------------------------------------------------------------------------------
extrair_todas_tabelas_novo <- function(x, nume, sheet = sheet){
  mes_boletim <- datas_arquivos[[nume,1]]
  
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


## --------------------------------------------------------------------------------------------------------------------------------
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


## --------------------------------------------------------------------------------------------------------------------------------
extrair_3_tabelas <- function(nume, ajuste = 2, ajuste2 = -2, sheet = sheet){
prov <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

   
  
df1 <-  read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[1] + ajuste, n_max = maximo_linha[1] + ajuste2) |> 
     gather(key = "mes", value = "volume", -1)

df2 <-  read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[2] + ajuste, n_max = maximo_linha[2] + ajuste2) |> 
     gather(key = "mes", value = "volume", -1)
  
df3 <-  read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[3] + ajuste) |> 
     gather(key = "mes", value = "volume", -1)  
  

list(df1, df2, df3)

}  


## --------------------------------------------------------------------------------------------------------------------------------
extrair_3_tabelas_data <- function(nume, ajuste = 2, ajuste2 = -2, sheet = sheet){

mes_boletim <- datas_arquivos[[nume,1]]

  prov <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

   
  
df1 <-  read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[1] + ajuste, n_max = maximo_linha[1] + ajuste2) |> 
          mutate(mes = mes_boletim)    

df2 <-  read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[2] + ajuste, n_max = maximo_linha[2] + ajuste2) |> 
          mutate(mes = mes_boletim)    
  
df3 <-  read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[3] + ajuste) |> 
     gather(key = "mes", value = "volume", -1) |> 
          mutate(mes = mes_boletim)     
  

list(df1, df2, df3)

}  


## --------------------------------------------------------------------------------------------------------------------------------
extrair_4_tabelas <- function(nume, ajuste = 2, ajuste2 = -2, sheet = sheet){

  
  prov <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o inínio da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica ,mm
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

   
  
df1 <-  read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[1] + ajuste, n_max = maximo_linha[1] + ajuste2, col_types = "text")  |> 
     gather(key = "mes", value = "volume", -1)

df2 <-  read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[2] + ajuste, n_max = maximo_linha[2] + ajuste2, col_types = "text")  |> 
     gather(key = "mes", value = "volume", -1)

df3 <-  read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[3] + ajuste, n_max = maximo_linha[3] + ajuste2, col_types = "text") 
  
df4 <-  read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[4] + ajuste)  |> 
     gather(key = "mes", value = "volume", -1)  
  

list(df1, df2, df3, df4)

}  


## --------------------------------------------------------------------------------------------------------------------------------
extrair_4_tabelas_novo <- function(x, ajuste = 2, ajuste2 = -2, sheet = sheet){

  
  prov <- read_excel(paste0("data/", x), sheet = sheet)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o inínio da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica ,mm
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

   
  
df1 <-  read_excel(paste0("data/", x), sheet = sheet, skip = corte[1] + ajuste, n_max = maximo_linha[1] + ajuste2, col_types = "text")  |> 
     gather(key = "mes", value = "volume", -1)

df2 <-  read_excel(paste0("data/", x), sheet = sheet, skip = corte[2] + ajuste, n_max = maximo_linha[2] + ajuste2, col_types = "text")  |> 
     gather(key = "mes", value = "volume", -1)

df3 <-  read_excel(paste0("data/", x), sheet = sheet, skip = corte[3] + ajuste, n_max = maximo_linha[3] + ajuste2, col_types = "text") 
  
df4 <-  read_excel(paste0("data/", x), sheet = sheet, skip = corte[4] + ajuste)  |> 
     gather(key = "mes", value = "volume", -1)  
  

list(df1, df2, df3, df4)

}  


## --------------------------------------------------------------------------------------------------------------------------------
extrair_6_tabelas <- function(nume, ajuste = 2, ajuste2 = -2, sheet = 3){
mes_boletim <- datas_arquivos[[nume,1]]

  prov <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
ajuste <- 2
ajuste2 <- -2

my_list <- list()
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

 assign(x = paste0("dataf", 1), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[1] + 
  ajuste, n_max = maximo_linha[1] + ajuste2) |> 
          mutate(mes = mes_boletim))    

 assign(x = paste0("dataf", 2), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[2] + 
  ajuste, n_max = maximo_linha[2] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
  assign(x = paste0("dataf", 3), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[3] + 
  ajuste, n_max = maximo_linha[3] + ajuste2) |> 
          mutate(mes = mes_boletim))
  
   assign(x = paste0("dataf", 4), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[4] + 
  ajuste, n_max = maximo_linha[4] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
   assign(x = paste0("dataf", 5), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[5] + 
  ajuste, n_max = maximo_linha[5] + ajuste2) |> 
          mutate(mes = mes_boletim))
   
     
      
# i <- 1
# while(i < 13){
# assign(x = paste0("dataf", i), value = read_excel(paste0("data/", planilhas[i]), sheet = sheet, skip = corte[i], n_max = maximo_linha[i]) |> 
#          mutate(mes = mes_boletim))
#  
#  i <- i + 1
#}
   
    assign(x = paste0("dataf", 6), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[6] +  ajuste) |> 
          mutate(mes = mes_boletim))

mylist <- list(dataf1, dataf2, dataf3, dataf4, dataf5, dataf6)

mylist
}


## --------------------------------------------------------------------------------------------------------------------------------

extrair_6_tabelas_novo <- function(nume, ajuste = 2, ajuste2 = -2, sheet = 3){
mes_boletim <- data_arquivos_novo[[nume,1]]

  prov <- read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
ajuste <- 2
ajuste2 <- -2

my_list <- list()
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

 assign(x = paste0("dataf", 1), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[1] + 
  ajuste, n_max = maximo_linha[1] + ajuste2) |> 
          mutate(mes = mes_boletim))    

 assign(x = paste0("dataf", 2), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[2] + 
  ajuste, n_max = maximo_linha[2] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
  assign(x = paste0("dataf", 3), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[3] + 
  ajuste, n_max = maximo_linha[3] + ajuste2) |> 
          mutate(mes = mes_boletim))
  
   assign(x = paste0("dataf", 4), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[4] + 
  ajuste, n_max = maximo_linha[4] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
   assign(x = paste0("dataf", 5), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[5] + 
  ajuste, n_max = maximo_linha[5] + ajuste2) |> 
          mutate(mes = mes_boletim))
   
     
      
# i <- 1
# while(i < 13){
# assign(x = paste0("dataf", i), value = read_excel(paste0("data/", planilhas[i]), sheet = sheet, skip = corte[i], n_max = maximo_linha[i]) |> 
#          mutate(mes = mes_boletim))
#  
#  i <- i + 1
#}
   
    assign(x = paste0("dataf", 6), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[6] +  ajuste) |> 
          mutate(mes = mes_boletim))

mylist <- list(dataf1, dataf2, dataf3, dataf4, dataf5, dataf6)

mylist
}


## --------------------------------------------------------------------------------------------------------------------------------
extrair_7_tabelas <- function(nume, ajuste = 2, ajuste2 = -2, sheet = 3){
mes_boletim <- datas_arquivos[[nume,1]]

  prov <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
ajuste <- 2
ajuste2 <- -2

my_list <- list()
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

 assign(x = paste0("dataf", 1), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[1] + 
  ajuste, n_max = maximo_linha[1] + ajuste2) |> 
          mutate(mes = mes_boletim))    

 assign(x = paste0("dataf", 2), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[2] + 
  ajuste, n_max = maximo_linha[2] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
  assign(x = paste0("dataf", 3), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[3] + 
  ajuste, n_max = maximo_linha[3] + ajuste2) |> 
          mutate(mes = mes_boletim))
  
   assign(x = paste0("dataf", 4), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[4] + 
  ajuste, n_max = maximo_linha[4] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
   assign(x = paste0("dataf", 5), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[5] + 
  ajuste, n_max = maximo_linha[5] + ajuste2) |> 
          mutate(mes = mes_boletim))
   
      assign(x = paste0("dataf", 6), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[6] + 
  ajuste, n_max = maximo_linha[6] + ajuste2) |> 
          mutate(mes = mes_boletim))
      
# i <- 1
# while(i < 13){
# assign(x = paste0("dataf", i), value = read_excel(paste0("data/", planilhas[i]), sheet = sheet, skip = corte[i], n_max = maximo_linha[i]) |> 
#          mutate(mes = mes_boletim))
#  
#  i <- i + 1
#}

 assign(x = paste0("dataf", 7), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[7] + ajuste) |> 
          mutate(mes = mes_boletim))  

mylist <- list(dataf1, dataf2, dataf3, dataf4, dataf5, dataf6, dataf7)

mylist
}


## --------------------------------------------------------------------------------------------------------------------------------
extrair_8_tabelas <- function(nume, ajuste = 2, ajuste2 = -2, sheet = 3){
mes_boletim <- datas_arquivos[[nume,1]]

  prov <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
ajuste <- 2
ajuste2 <- -2

my_list <- list()
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

 assign(x = paste0("dataf", 1), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[1] + 
  ajuste, n_max = maximo_linha[1] + ajuste2) |> 
          mutate(mes = mes_boletim))    

 assign(x = paste0("dataf", 2), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[2] + 
  ajuste, n_max = maximo_linha[2] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
  assign(x = paste0("dataf", 3), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[3] + 
  ajuste, n_max = maximo_linha[3] + ajuste2) |> 
          mutate(mes = mes_boletim))
  
   assign(x = paste0("dataf", 4), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[4] + 
  ajuste, n_max = maximo_linha[4] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
   assign(x = paste0("dataf", 5), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[5] + 
  ajuste, n_max = maximo_linha[5] + ajuste2) |> 
          mutate(mes = mes_boletim))
   
      assign(x = paste0("dataf", 6), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[6] + 
  ajuste, n_max = maximo_linha[6] + ajuste2) |> 
          mutate(mes = mes_boletim))
      
      assign(x = paste0("dataf", 7), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[7] + 
  ajuste, n_max = maximo_linha[7] + ajuste2) |> 
          mutate(mes = mes_boletim))
      
           
# i <- 1
# while(i < 13){
# assign(x = paste0("dataf", i), value = read_excel(paste0("data/", planilhas[i]), sheet = sheet, skip = corte[i], n_max = maximo_linha[i]) |> 
#          mutate(mes = mes_boletim))
#  
#  i <- i + 1
#}

 assign(x = paste0("dataf", 8), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[8] + ajuste) |> 
          mutate(mes = mes_boletim))  

mylist <- list(dataf1, dataf2, dataf3, dataf4, dataf5, dataf6, dataf7, dataf8)

mylist
}


## --------------------------------------------------------------------------------------------------------------------------------
extrair_8_tabelas_novo <- function(nume, ajuste = 2, ajuste2 = -2, sheet = 3){
mes_boletim <- data_arquivos_novo[[nume,1]]

  prov <- read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
ajuste <- 2
ajuste2 <- -2

my_list <- list()
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

 assign(x = paste0("dataf", 1), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[1] + 
  ajuste, n_max = maximo_linha[1] + ajuste2) |> 
          mutate(mes = mes_boletim))    

 assign(x = paste0("dataf", 2), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[2] + 
  ajuste, n_max = maximo_linha[2] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
  assign(x = paste0("dataf", 3), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[3] + 
  ajuste, n_max = maximo_linha[3] + ajuste2) |> 
          mutate(mes = mes_boletim))
  
   assign(x = paste0("dataf", 4), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[4] + 
  ajuste, n_max = maximo_linha[4] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
   assign(x = paste0("dataf", 5), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[5] + 
  ajuste, n_max = maximo_linha[5] + ajuste2) |> 
          mutate(mes = mes_boletim))
   
      assign(x = paste0("dataf", 6), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[6] + 
  ajuste, n_max = maximo_linha[6] + ajuste2) |> 
          mutate(mes = mes_boletim))
      
      assign(x = paste0("dataf", 7), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[7] + 
  ajuste, n_max = maximo_linha[7] + ajuste2) |> 
          mutate(mes = mes_boletim))
      
           
# i <- 1
# while(i < 13){
# assign(x = paste0("dataf", i), value = read_excel(paste0("data/", planilhas[i]), sheet = sheet, skip = corte[i], n_max = maximo_linha[i]) |> 
#          mutate(mes = mes_boletim))
#  
#  i <- i + 1
#}

 assign(x = paste0("dataf", 8), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = sheet, skip = corte[8] + ajuste) |> 
          mutate(mes = mes_boletim))  

mylist <- list(dataf1, dataf2, dataf3, dataf4, dataf5, dataf6, dataf7, dataf8)

mylist
}


## --------------------------------------------------------------------------------------------------------------------------------
extrair_13_tabelas <- function(nume, ajuste = 2, ajuste2 = -2, sheet = 3){
mes_boletim <- datas_arquivos[[nume,1]]

  prov <- read_excel(paste0("data/", planilhas[nume]), sheet = 3)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
ajuste <- 2
ajuste2 <- -2

my_list <- list()
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

 assign(x = paste0("dataf", 1), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[1] + 
  ajuste, n_max = maximo_linha[1] + ajuste2) |> 
          mutate(mes = mes_boletim))    

 assign(x = paste0("dataf", 2), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[2] + 
  ajuste, n_max = maximo_linha[2] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
  assign(x = paste0("dataf", 3), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[3] + 
  ajuste, n_max = maximo_linha[3] + ajuste2) |> 
          mutate(mes = mes_boletim))
  
   assign(x = paste0("dataf", 4), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[4] + 
  ajuste, n_max = maximo_linha[4] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
   assign(x = paste0("dataf", 5), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[5] + 
  ajuste, n_max = maximo_linha[5] + ajuste2) |> 
          mutate(mes = mes_boletim))
   
      assign(x = paste0("dataf", 6), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[6] + 
  ajuste, n_max = maximo_linha[6] + ajuste2) |> 
          mutate(mes = mes_boletim))
      
      assign(x = paste0("dataf", 7), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[7] + 
  ajuste, n_max = maximo_linha[7] + ajuste2) |> 
          mutate(mes = mes_boletim))
      
            assign(x = paste0("dataf", 8), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[8] + 
  ajuste, n_max = maximo_linha[8] + ajuste2) |> 
          mutate(mes = mes_boletim))
   
               assign(x = paste0("dataf", 9), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[9] + 
  ajuste, n_max = maximo_linha[9] + ajuste2) |> 
          mutate(mes = mes_boletim))         
            
                  assign(x = paste0("dataf", 10), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[10] + 
  ajuste, n_max = maximo_linha[10] + ajuste2) |> 
          mutate(mes = mes_boletim))
               
                                    assign(x = paste0("dataf", 11), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[11] + 
  ajuste, n_max = maximo_linha[11] + ajuste2) |> 
          mutate(mes = mes_boletim))
                  
                                       assign(x = paste0("dataf", 12), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[12] + 
  ajuste, n_max = maximo_linha[12] + ajuste2) |> 
          mutate(mes = mes_boletim))
                                    
# i <- 1
# while(i < 13){
# assign(x = paste0("dataf", i), value = read_excel(paste0("data/", planilhas[i]), sheet = 3, skip = corte[i], n_max = maximo_linha[i]) |> 
#          mutate(mes = mes_boletim))
#  
#  i <- i + 1
#}

 assign(x = paste0("dataf", 13), value = read_excel(paste0("data/", planilhas[nume]), sheet = 3, skip = corte[13] + ajuste) |> 
          mutate(mes = mes_boletim))  

mylist <- list(dataf1, dataf2, dataf3, dataf4, dataf5, dataf6, dataf7, dataf8, dataf9, dataf10, dataf11, dataf12, dataf13)

mylist
}


## --------------------------------------------------------------------------------------------------------------------------------
extrair_13_tabelas_novo <- function(nume, ajuste = 2, ajuste2 = -2, sheet = 3){
mes_boletim <- data_arquivos_novo[[nume,1]]

  prov <- read_excel(paste0("data/", planilhas[nume]), sheet = 3)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
ajuste <- 2
ajuste2 <- -2

my_list <- list()
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

 assign(x = paste0("dataf", 1), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[1] + 
  ajuste, n_max = maximo_linha[1] + ajuste2) |> 
          mutate(mes = mes_boletim))    

 assign(x = paste0("dataf", 2), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[2] + 
  ajuste, n_max = maximo_linha[2] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
  assign(x = paste0("dataf", 3), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[3] + 
  ajuste, n_max = maximo_linha[3] + ajuste2) |> 
          mutate(mes = mes_boletim))
  
   assign(x = paste0("dataf", 4), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[4] + 
  ajuste, n_max = maximo_linha[4] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
   assign(x = paste0("dataf", 5), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[5] + 
  ajuste, n_max = maximo_linha[5] + ajuste2) |> 
          mutate(mes = mes_boletim))
   
      assign(x = paste0("dataf", 6), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[6] + 
  ajuste, n_max = maximo_linha[6] + ajuste2) |> 
          mutate(mes = mes_boletim))
      
      assign(x = paste0("dataf", 7), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[7] + 
  ajuste, n_max = maximo_linha[7] + ajuste2) |> 
          mutate(mes = mes_boletim))
      
            assign(x = paste0("dataf", 8), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[8] + 
  ajuste, n_max = maximo_linha[8] + ajuste2) |> 
          mutate(mes = mes_boletim))
   
               assign(x = paste0("dataf", 9), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[9] + 
  ajuste, n_max = maximo_linha[9] + ajuste2) |> 
          mutate(mes = mes_boletim))         
            
                  assign(x = paste0("dataf", 10), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[10] + 
  ajuste, n_max = maximo_linha[10] + ajuste2) |> 
          mutate(mes = mes_boletim))
               
                                    assign(x = paste0("dataf", 11), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[11] + 
  ajuste, n_max = maximo_linha[11] + ajuste2) |> 
          mutate(mes = mes_boletim))
                  
                                       assign(x = paste0("dataf", 12), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[12] + 
  ajuste, n_max = maximo_linha[12] + ajuste2) |> 
          mutate(mes = mes_boletim))
                                    
# i <- 1
# while(i < 13){
# assign(x = paste0("dataf", i), value = read_excel(paste0("data/", planilhas[i]), sheet = 3, skip = corte[i], n_max = maximo_linha[i]) |> 
#          mutate(mes = mes_boletim))
#  
#  i <- i + 1
#}

 assign(x = paste0("dataf", 13), value = read_excel(paste0("data/", planilhas_novas[nume]), sheet = 3, skip = corte[13] + ajuste) |> 
          mutate(mes = mes_boletim))  

mylist <- list(dataf1, dataf2, dataf3, dataf4, dataf5, dataf6, dataf7, dataf8, dataf9, dataf10, dataf11, dataf12, dataf13)

mylist
}


## --------------------------------------------------------------------------------------------------------------------------------

extrair_16_tabelas <- function(nume, ajuste = 2, ajuste2 = -2, sheet = 4){
mes_boletim <- datas_arquivos[[nume,1]]

  prov <- read_excel(paste0("data/", planilhas[nume]), sheet = sheet)

names(prov) <- "col1"

  corte <- which(str_detect(prov$col1, "Tabela [0-9]*"))
  # E a a linha seguinte com a palavra tabela indica o início da tabela seguinte
  maximo_linha <- lead(corte) - corte
  
ajuste <- 2
ajuste2 <- -2

my_list <- list()
  
  # Importação de três tabelas, usando os parâmetros descobertos anteriormente
  # nos argumentos "skip" e "n_max". Dessa maneira, a função readxl identifica
  # automaticamente o número de colunas de cada tabela e a classe dos caracteres

 assign(x = paste0("dataf", 1), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[1] + 
  ajuste, n_max = maximo_linha[1] + ajuste2) |> 
          mutate(mes = mes_boletim))    

 assign(x = paste0("dataf", 2), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[2] + 
  ajuste, n_max = maximo_linha[2] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
  assign(x = paste0("dataf", 3), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[3] + 
  ajuste, n_max = maximo_linha[3] + ajuste2) |> 
          mutate(mes = mes_boletim))
  
   assign(x = paste0("dataf", 4), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[4] + 
  ajuste, n_max = maximo_linha[4] + ajuste2) |> 
          mutate(mes = mes_boletim))
 
   assign(x = paste0("dataf", 5), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[5] + 
  ajuste, n_max = maximo_linha[5] + ajuste2) |> 
          mutate(mes = mes_boletim))
   
      assign(x = paste0("dataf", 6), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[6] + 
  ajuste, n_max = maximo_linha[6] + ajuste2) |> 
          mutate(mes = mes_boletim))
      
      assign(x = paste0("dataf", 7), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[7] + 
  ajuste, n_max = maximo_linha[7] + ajuste2) |> 
          mutate(mes = mes_boletim))
      
            assign(x = paste0("dataf", 8), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[8] + 
  ajuste, n_max = maximo_linha[8] + ajuste2) |> 
          mutate(mes = mes_boletim))
   
               assign(x = paste0("dataf", 9), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[9] + 
  ajuste, n_max = maximo_linha[9] + ajuste2) |> 
          mutate(mes = mes_boletim))         
            
                  assign(x = paste0("dataf", 10), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[10] + 
  ajuste, n_max = maximo_linha[10] + ajuste2) |> 
          mutate(mes = mes_boletim))
               
                                    assign(x = paste0("dataf", 11), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[11] + 
  ajuste, n_max = maximo_linha[11] + ajuste2) |> 
          mutate(mes = mes_boletim))
                  
                                       assign(x = paste0("dataf", 12), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[12] + 
  ajuste, n_max = maximo_linha[12] + ajuste2) |> 
          mutate(mes = mes_boletim))
                                       
                                         assign(x = paste0("dataf", 13), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[13] + 
  ajuste, n_max = maximo_linha[13] + ajuste2) |> 
          mutate(mes = mes_boletim))
                                         
                                         assign(x = paste0("dataf", 14), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[14] + 
  ajuste, n_max = maximo_linha[14] + ajuste2) |> 
          mutate(mes = mes_boletim))
                                         
                                         assign(x = paste0("dataf", 15), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[15] + 
  ajuste, n_max = maximo_linha[15] + ajuste2) |> 
          mutate(mes = mes_boletim))
                                    
# i <- 1
# while(i < 13){
# assign(x = paste0("dataf", i), value = read_excel(paste0("data/", planilhas[i]), sheet = sheet, skip = corte[i], n_max = maximo_linha[i]) |> 
#          mutate(mes = mes_boletim))
#  
#  i <- i + 1
#}

 assign(x = paste0("dataf", 16), value = read_excel(paste0("data/", planilhas[nume]), sheet = sheet, skip = corte[16] + ajuste) |> 
          mutate(mes = mes_boletim))  

mylist <- list(dataf1, dataf2, dataf3, dataf4, dataf5, dataf6, dataf7, dataf8, dataf9, dataf10, dataf11, dataf12, dataf13, dataf14, dataf15, dataf16)

mylist
}

