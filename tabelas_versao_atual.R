## ----setup, include=FALSE-----------------------------------------------------------------------------------
knitr::opts_chunk$set(echo = TRUE)


## -----------------------------------------------------------------------------------------------------------
library(tidyverse)
library(readxl)


## -----------------------------------------------------------------------------------------------------------
source(knitr::purl("funcoes_data_e_extrair.Rmd", quiet=TRUE))


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) 


## ----message=FALSE------------------------------------------------------------------------------------------
lista_guia2_versao_atual <- map(4:76, extrair_4_tabelas, sheet = 2)

## ----message=FALSE------------------------------------------------------------------------------------------
lista_guia2_versao_meno3 <- map(1:3, extrair_3_tabelas, sheet = 2)

# gather(key = "mes", value = "produto", -1)


## ----message=FALSE------------------------------------------------------------------------------------------
lista_guia4_versao_meno3 <- map(1:3, extrair_3_tabelas_data, sheet = 4)


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(1)

read_excel(paste0("data/", planilhas[4]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(1)

read_excel(paste0("data/", planilhas[1]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(1)


## -----------------------------------------------------------------------------------------------------------
Histórico_petróleo. <-   rbind(
do.call(rbind, map(lista_guia2_versao_atual, `[[`, 1)),
do.call(rbind, map(lista_guia2_versao_meno3, `[[`, 1))) |> 
  mutate(mes = str_replace(mes, "/16", "/2016")  ) |> 
  mutate(mes = str_replace(mes, "/17", "/2017")  ) |> 
  mutate(mes = str_replace(mes, "/18", "/2018")  ) 

Historico_petroleo <- rbind(
 Histórico_petróleo. |> 
  filter(!str_detect(mes, "/")) |> 
   mutate(mes = as.Date(as.numeric(mes), origin = "1899-12-30")), 
Histórico_petróleo. |> 
  filter(str_detect(mes, "/")) |>
   mutate(mes = lubridate::my(mes)))

Historico_petroleo <- Historico_petroleo |> 
  distinct() |> 
  arrange(mes)


## -----------------------------------------------------------------------------------------------------------
Historico_petroleo |> 
  write.csv2(file = "Historico_petroleo.csv")

 


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(2)

read_excel(paste0("data/", planilhas[4]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(2)

read_excel(paste0("data/", planilhas[1]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(2)



## -----------------------------------------------------------------------------------------------------------
Histórico_prod_gas. <-   rbind(
do.call(rbind, map(lista_guia2_versao_atual, `[[`, 2)),
do.call(rbind, map(lista_guia2_versao_meno3, `[[`, 2))) |> 
  mutate(mes = str_replace(mes, "/16", "/2016")  ) |> 
  mutate(mes = str_replace(mes, "/17", "/2017")  ) |> 
  mutate(mes = str_replace(mes, "/18", "/2018")  ) 

Historico_prod_gas <- rbind(
 Histórico_prod_gas. |> 
  filter(!str_detect(mes, "/")) |> 
   mutate(mes = as.Date(as.numeric(mes), origin = "1899-12-30")), 
Histórico_prod_gas. |> 
  filter(str_detect(mes, "/")) |>
   mutate(mes = lubridate::my(mes))) 

Historico_prod_gas <- Historico_prod_gas |> 
  distinct() |> 
  arrange(mes)


## -----------------------------------------------------------------------------------------------------------
Historico_prod_gas |> 
  distinct() |> 
  write.csv2(file = "Historico_prod_gas.csv")

 


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(3)

read_excel(paste0("data/", planilhas[4]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(3)

read_excel(paste0("data/", planilhas[1]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(2)


## -----------------------------------------------------------------------------------------------------------
df1 <- bind_rows( map(lista_guia2_versao_atual, `[[`, 3)) |> 
  filter(is.na(Período)) |> 
  filter(str_detect("Mês/Ano", "/")) |> 
  mutate("Mês/Ano" = my(`Mês/Ano`)) |> 
  select(!Período) |> 
  rename(Período = "Mês/Ano")

df2 <- bind_rows( map(lista_guia2_versao_atual, `[[`, 3)) |> 
  filter(!is.na(Período)) |> 
  filter(!str_detect(Período, "/")) |> 
  mutate(Período = as.Date(as.numeric(Período), origin = "1899-12-30")) |> 
  select(!`Mês/Ano`)

df3 <- bind_rows( map(lista_guia2_versao_atual, `[[`, 3)) |> 
  filter(!is.na(Período)) |> 
  filter(str_detect(Período, "/")) |> 
  mutate(Período = my(Período)) |> 
  select(!`Mês/Ano`)

MOVIMENTACAO_DE_GAS_NATURAL <- bind_rows(df1, df2, df3) |> 
  distinct()


## -----------------------------------------------------------------------------------------------------------
MOVIMENTACAO_DE_GAS_NATURAL  |> 
  write.csv2(file = "MOVIMENTACAO_DE_GAS_NATURAL_ANTIGA.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(4)

read_excel(paste0("data/", planilhas[4]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(4)

read_excel(paste0("data/", planilhas[1]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(3)



## -----------------------------------------------------------------------------------------------------------
Historico_de_Producao_de_Petroleo_Gas_Natural <-   rbind(
do.call(rbind, map(lista_guia2_versao_atual, `[[`, 4)),
do.call(rbind, map(lista_guia2_versao_meno3, `[[`, 3))) |> 
  mutate(mes = str_replace(mes, "/16", "/2016")  ) |> 
  mutate(mes = str_replace(mes, "/17", "/2017")  ) |> 
  mutate(mes = str_replace(mes, "/18", "/2018")  ) 

Historico_de_Producao_de_Petroleo_Gas_Natural |>
   mutate(mes = lubridate::my(mes)) |> 
  distinct()

# Histórico_prod_gas <- rbind(
# Histórico_prod_gas. |> 
#  filter(!str_detect(mes, "/")) |> 
#   mutate(mes = as.Date(as.numeric(mes), origin = "1899-12-30")), 
#Histórico_prod_gas. |> 
#  filter(str_detect(mes, "/")) |>
#   mutate(mes = lubridate::my(mes)))


## -----------------------------------------------------------------------------------------------------------
Historico_de_Producao_de_Petroleo_Gas_Natural |> 
  distinct() |> 
  write.csv2(file = "Historico_de_Producao_de_Petroleo_Gas_Natural.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) 


## ----message=FALSE------------------------------------------------------------------------------------------
lista_guia3_versao_atual <- map(57:76, extrair_6_tabelas, sheet = 3)


## ----message=FALSE------------------------------------------------------------------------------------------
lista_guia3_versao_menos1 <- map(4:56, extrair_7_tabelas, sheet = 3)


## -----------------------------------------------------------------------------------------------------------
lista_guia3_versao_menos3 <- map(1:3, extrair_13_tabelas, sheet = 3)


## ----message=FALSE------------------------------------------------------------------------------------------
lista_guia4_versao_menos3 <- map(1:3, extrair_3_tabelas_data, sheet = 4)


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(1)

read_excel(paste0("data/", planilhas[4]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(1)

read_excel(paste0("data/", planilhas[1]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(1)


## -----------------------------------------------------------------------------------------------------------
# Acertar nomes das colunas
nomes <- names(lista_guia3_versao_atual[[1]][[1]] )

i <- 1
while(i <= length(lista_guia3_versao_menos1)){
  names(lista_guia3_versao_menos1[[i]][[1]] ) <- nomes
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos3)){
  names(lista_guia3_versao_menos3[[i]][[1]] ) <- nomes
  i <- i + 1
}


## -----------------------------------------------------------------------------------------------------------
por_estado <- rbind(
do.call(rbind, map(lista_guia3_versao_atual, `[[`, 1)),

do.call(rbind, map(lista_guia3_versao_menos1, `[[`, 1)),

do.call(rbind, map(lista_guia3_versao_menos3, `[[`, 1))  
)


## -----------------------------------------------------------------------------------------------------------
por_estado |> 
  distinct() |> 
  write.csv2(file = "por_estado.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(2)

read_excel(paste0("data/", planilhas[4]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(2)

read_excel(paste0("data/", planilhas[1]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(2)


## -----------------------------------------------------------------------------------------------------------
# Acertar nomes das colunas
nomes <- names(lista_guia3_versao_atual[[1]][[2]] )

i <- 1
while(i <= length(lista_guia3_versao_menos1)){
  names(lista_guia3_versao_menos1[[i]][[2]] ) <- nomes
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos3)){
  names(lista_guia3_versao_menos3[[i]][[2]] ) <- nomes
  i <- i + 1
}


## -----------------------------------------------------------------------------------------------------------
por_bacia <- rbind(
do.call(rbind, map(lista_guia3_versao_atual, `[[`, 2)),

do.call(rbind, map(lista_guia3_versao_menos1, `[[`, 2)),

do.call(rbind, map(lista_guia3_versao_menos3, `[[`, 2))  
)


## -----------------------------------------------------------------------------------------------------------
por_bacia |> 
  distinct() |> 
  write.csv2(file = "por_bacia.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(3)

read_excel(paste0("data/", planilhas[4]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(3)

read_excel(paste0("data/", planilhas[1]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(3)


## ----eval=FALSE, include=FALSE------------------------------------------------------------------------------
## # Acertar nomes das colunas
## nomes <- names(lista_guia3_versao_atual[[1]][[3]] )
## 
## i <- 1
## while(i <= length(lista_guia3_versao_menos1)){
##   names(lista_guia3_versao_menos1[[i]][[3]] ) <- nomes
##   i <- i + 1
## }
## 
## i <- 1
## while(i <= length(lista_guia3_versao_menos3)){
##   names(lista_guia3_versao_menos3[[i]][[3]] ) <- nomes
##   i <- i + 1
## }


## -----------------------------------------------------------------------------------------------------------
por_operador <- rbind(
do.call(rbind, map(lista_guia3_versao_atual, `[[`, 3)),

do.call(rbind, map(lista_guia3_versao_menos1, `[[`, 3)),

do.call(rbind, map(lista_guia3_versao_menos3, `[[`, 3))  
)


## -----------------------------------------------------------------------------------------------------------
por_operador |> 
  distinct() |> 
  write.csv2(file = "por_operador.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(4)

read_excel(paste0("data/", planilhas[4]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(4)

read_excel(paste0("data/", planilhas[1]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(4)


## -----------------------------------------------------------------------------------------------------------
# Dataframe 16 tem colunas em branco, detectando e acertando isso
# i <- 1

# while(i <= length(lista_guia3_versao_atual)){
#  print(i)
#  names(lista_guia3_versao_atual[[i]][[4]])  |> 
#    print()
#
#    i <- i + 1
#}

# lista_guia3_versao_atual[[16]][[4]] <- lista_guia3_versao_atual[[16]][[4]][,c(1:5, 11)]


## -----------------------------------------------------------------------------------------------------------
# Apagar colunas que estão sobrando (do dataframe 16)
# i <- 1
# while(i <= length(lista_guia3_versao_atual)){
#  print(i)
#  print(ncol(lista_guia3_versao_atual[[i]][[4]]))
#  i <- i + 1
#}

lista_guia3_versao_atual[[16]][[4]] <- lista_guia3_versao_atual[[16]][[4]] |> 
  select(!6:10)

lista_guia3_versao_atual[[16]][[4]]


## -----------------------------------------------------------------------------------------------------------
# Acertar nomes das colunas
nomes <- names(lista_guia3_versao_atual[[1]][[4]] )

i <- 1
while(i <= length(lista_guia3_versao_atual)){
  names(lista_guia3_versao_atual[[i]][[4]] ) <- nomes
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos1)){
  names(lista_guia3_versao_menos1[[i]][[4]] ) <- nomes
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos3)){
  names(lista_guia3_versao_menos3[[i]][[4]] ) <- nomes
  i <- i + 1
}

# Transformar coluna double em character
i <- 1
while(i <= length(lista_guia3_versao_atual)){
  lista_guia3_versao_atual[[i]][[4]][,1] <- pull(lista_guia3_versao_atual[[i]][[4]][,1]) |> 
  as.character()
  i <- i + 1
}




## -----------------------------------------------------------------------------------------------------------
por_concessionario <- rbind(
do.call(rbind, map(lista_guia3_versao_atual, `[[`, 4)),

do.call(rbind, map(lista_guia3_versao_menos1, `[[`, 4)),

do.call(rbind, map(lista_guia3_versao_menos3, `[[`, 4))  
)


## -----------------------------------------------------------------------------------------------------------
por_concessionario |> 
  distinct() |> 
  write.csv2(file = "por_concessionario.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(5)

read_excel(paste0("data/", planilhas[4]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(5)

# Está na guia 5, tabela 2 
read_excel(paste0("data/", planilhas[1]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
    select(1) |> 
  slice(2)



## -----------------------------------------------------------------------------------------------------------
# Dataframe 16 tem colunas em branco, detectando e acertando isso
i <- 1

while(i <= length(lista_guia3_versao_atual)){
  print(i)
  names(lista_guia3_versao_atual[[i]][[4]])  |> 
    print()

    i <- i + 1
}

# lista_guia3_versao_atual[[16]][[4]] <- lista_guia3_versao_atual[[16]][[4]][,c(1:5, 11)]


## -----------------------------------------------------------------------------------------------------------
# Acertar nomes das colunas
nomes <- names(lista_guia3_versao_atual[[1]][[4]] )

i <- 1
while(i <= length(lista_guia3_versao_atual)){
  names(lista_guia3_versao_atual[[i]][[4]] ) <- nomes
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos1)){
  names(lista_guia3_versao_menos1[[i]][[4]] ) <- nomes
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos3)){
  names(lista_guia3_versao_menos3[[i]][[4]] ) <- nomes
  i <- i + 1
}


## -----------------------------------------------------------------------------------------------------------
MOVIMENTACAO_GAS_NATURAL_POR_DESTINO <- bind_rows(
do.call(bind_rows, map(lista_guia3_versao_atual, `[[`, 5)),

do.call(rbind, map(lista_guia3_versao_menos1, `[[`, 5)),

do.call(rbind, map(lista_guia4_versao_menos3, `[[`, 2))  
)


## -----------------------------------------------------------------------------------------------------------
MOVIMENTACAO_GAS_NATURAL_POR_DESTINO |> 
  distinct() |> 
  write.csv2(file = "MOVIMENTACAO_GAS_NATURAL_POR_DESTINO.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(6)

read_excel(paste0("data/", planilhas[4]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(6)

read_excel(paste0("data/", planilhas[1]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(6)


## -----------------------------------------------------------------------------------------------------------
# Acertar nomes das colunas
nomes <- names(lista_guia3_versao_atual[[1]][[6]] )

i <- 1
while(i <= length(lista_guia3_versao_menos1)){
  names(lista_guia3_versao_menos1[[i]][[6]] ) <- nomes
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos3)){
  names(lista_guia3_versao_menos3[[i]][[6]] ) <- nomes
  i <- i + 1
}


## -----------------------------------------------------------------------------------------------------------
i <- 1
while(i <= length(lista_guia3_versao_menos1)){
  lista_guia3_versao_menos1[[i]][[6]] <- lista_guia3_versao_menos1[[i]][[6]][-1,] 
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos3)){
 lista_guia3_versao_menos3[[i]][[6]] <- lista_guia3_versao_menos3[[i]][[6]][-1,] 
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos1)){
  lista_guia3_versao_menos1[[i]][[6]] <- lista_guia3_versao_menos1[[i]][[6]] |> 
  mutate_at(2:7, as.numeric) 
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos3)){
 lista_guia3_versao_menos3[[i]][[6]] <- lista_guia3_versao_menos3[[i]][[6]] |> 
  mutate_at(2:7, as.numeric) 
  i <- i + 1
}



## -----------------------------------------------------------------------------------------------------------
DISTRIBUICAO_DA_PRODUCAO_CAMPOS_PRE_SAL <- bind_rows(
do.call(rbind, map(lista_guia3_versao_atual, `[[`, 6)),

do.call(rbind, map(lista_guia3_versao_menos1, `[[`, 6)),



do.call(rbind, map(lista_guia3_versao_menos3, `[[`, 6))  
)


## -----------------------------------------------------------------------------------------------------------
DISTRIBUICAO_DA_PRODUCAO_CAMPOS_PRE_SAL |> 
  distinct() |> 
  write.csv2(file = "DISTRIBUICAO_DA_PRODUCAO_CAMPOS_PRE_SAL.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) 

read_excel(paste0("data/", planilhas[56]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) 

read_excel(paste0("data/", planilhas[4]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) 

# Único que aparece aqui e nos outros também é o "20 Campos que Mais Queimaram Gás Natural Neste Mês"
read_excel(paste0("data/", planilhas[1]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) 


## -----------------------------------------------------------------------------------------------------------
# Acertar nomes das colunas
nomes <- names(lista_guia3_versao_atual[[1]][[6]] )

i <- 1
while(i <= length(lista_guia3_versao_menos1)){
  names(lista_guia3_versao_menos1[[i]][[6]] ) <- nomes
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos3)){
  names(lista_guia3_versao_menos3[[i]][[6]] ) <- nomes
  i <- i + 1
}


## -----------------------------------------------------------------------------------------------------------
i <- 1
while(i <= length(lista_guia3_versao_menos1)){
  lista_guia3_versao_menos1[[i]][[6]] <- lista_guia3_versao_menos1[[i]][[6]][-1,] 
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos3)){
 lista_guia3_versao_menos3[[i]][[6]] <- lista_guia3_versao_menos3[[i]][[6]][-1,] 
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos1)){
  lista_guia3_versao_menos1[[i]][[6]] <- lista_guia3_versao_menos1[[i]][[6]] |> 
  mutate_at(2:7, as.numeric) 
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia3_versao_menos3)){
 lista_guia3_versao_menos3[[i]][[6]] <- lista_guia3_versao_menos3[[i]][[6]] |> 
  mutate_at(2:7, as.numeric) 
  i <- i + 1
}




## -----------------------------------------------------------------------------------------------------------
DISTRIBUICAO_DA_PRODUCAO_CAMPOS_PRE_SAL <- bind_rows(
do.call(rbind, map(lista_guia3_versao_atual, `[[`, 6)),

do.call(rbind, map(lista_guia3_versao_menos1, `[[`, 6)),



do.call(rbind, map(lista_guia3_versao_menos3, `[[`, 6))  
)


## -----------------------------------------------------------------------------------------------------------
DISTRIBUICAO_DA_PRODUCAO_CAMPOS_PRE_SAL |> 
  distinct() |> 
  write.csv2(file = "DISTRIBUICAO_DA_PRODUCAO_CAMPOS_PRE_SAL.csv")


## ----message=FALSE------------------------------------------------------------------------------------------
lista_guia4_versao_atual <- map(c(54, 55, 57:76), extrair_8_tabelas, sheet = 4)


## -----------------------------------------------------------------------------------------------------------
lista_guia4_versao_menos1 <- map(c(33:53, 56), extrair_16_tabelas, sheet = 4)


## ----message=FALSE------------------------------------------------------------------------------------------
lista_guia4_versao_menos2 <- map(4:32, extrair_8_tabelas, sheet = 4)


## ----message=FALSE------------------------------------------------------------------------------------------
lista_guia4_versao_menos3 <- map(1:3, extrair_3_tabelas_data, sheet = 4)


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(1)

read_excel(paste0("data/", planilhas[32]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(1)

read_excel(paste0("data/", planilhas[4]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(1)


## -----------------------------------------------------------------------------------------------------------
# Acertar nomes das colunas
nomes <- names(lista_guia4_versao_atual[[1]][[1]] )

i <- 1
while(i <= length(lista_guia4_versao_atual)){
  names(lista_guia4_versao_atual[[i]][[1]] ) <- nomes
  i <- i + 1
}

nomes <- names(lista_guia4_versao_menos2[[1]][[1]] )

i <- 1
while(i <= 13){
  names(lista_guia4_versao_menos2[[i]][[1]] ) <- nomes
  i <- i + 1
}


## -----------------------------------------------------------------------------------------------------------
maiores_pocos_de_petroleo <- bind_rows(
do.call(rbind, map(lista_guia4_versao_atual, `[[`, 1)),

do.call(rbind, map(lista_guia4_versao_menos1, `[[`, 1)),  

do.call(bind_rows, map(lista_guia4_versao_menos2, `[[`, 1))  
)
   


## -----------------------------------------------------------------------------------------------------------
maiores_pocos_de_petroleo |> 
  distinct() |> 
  write.csv2(file = "maiores_pocos_de_petroleo.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(2)

read_excel(paste0("data/", planilhas[32]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(2)

read_excel(paste0("data/", planilhas[4]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(2)


## -----------------------------------------------------------------------------------------------------------
# Acertar nomes das colunas
nomes <- names(lista_guia4_versao_atual[[1]][[2]] )

i <- 1
while(i <= length(lista_guia4_versao_atual)){
  names(lista_guia4_versao_atual[[i]][[2]] ) <- nomes
  i <- i + 1
}

nomes <- names(lista_guia4_versao_menos2[[1]][[2]] )

i <- 1
while(i <= 13){
  names(lista_guia4_versao_menos2[[i]][[2]] ) <- nomes
  i <- i + 1
}


## -----------------------------------------------------------------------------------------------------------
maiores_pocos_de_gas <- bind_rows(
do.call(rbind, map(lista_guia4_versao_atual, `[[`, 2)),

do.call(rbind, map(lista_guia4_versao_menos1, `[[`, 2)),  

do.call(bind_rows, map(lista_guia4_versao_menos2, `[[`, 2))  
)
   


## -----------------------------------------------------------------------------------------------------------
maiores_pocos_de_gas |> 
  distinct() |> 
  write.csv2(file = "maiores_pocos_de_gas.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(3)

read_excel(paste0("data/", planilhas[32]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(3)

read_excel(paste0("data/", planilhas[4]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(3)


## -----------------------------------------------------------------------------------------------------------
# Retirar colunas vazias nos dataframes 10 e 20
lista_guia4_versao_atual[[10]][[3]] <- lista_guia4_versao_atual[[10]][[3]][,c(1:6, 13)]
lista_guia4_versao_atual[[20]][[3]] <- lista_guia4_versao_atual[[20]][[3]][,c(1:6, 13)]

# Acertar nomes das colunas
nomes <- names(lista_guia4_versao_atual[[1]][[3]] )

i <- 1
while(i <= length(lista_guia4_versao_atual)){
  names(lista_guia4_versao_atual[[i]][[3]] ) <- nomes
  i <- i + 1
}

nomes <- names(lista_guia4_versao_menos2[[1]][[3]] )

i <- 1
while(i <= 13){
  names(lista_guia4_versao_menos2[[i]][[3]] ) <- nomes
  i <- i + 1
}


## -----------------------------------------------------------------------------------------------------------
# i <- 1

# for(i in 1:length(lista_guia4_versao_atual)){
#  print(i)
#  print(lista_guia4_versao_atual[[i]][[3]] |> ncol())
#i <- 1 + 1
#  }

# Erro nos dataframes 10 e 20


## -----------------------------------------------------------------------------------------------------------
maiores_pocos_producao_total <- bind_rows(
do.call(rbind, map(lista_guia4_versao_atual, `[[`, 3)),

do.call(rbind, map(lista_guia4_versao_menos1, `[[`, 3)),  

do.call(bind_rows, map(lista_guia4_versao_menos2, `[[`, 3))  
)
   


## -----------------------------------------------------------------------------------------------------------
maiores_pocos_producao_total |> 
  distinct() |> 
  write.csv2(file = "maiores_pocos_producao_total.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(4)

read_excel(paste0("data/", planilhas[32]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(4)

read_excel(paste0("data/", planilhas[4]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(4)


## -----------------------------------------------------------------------------------------------------------
# Retirar colunas vazias nos dataframes 10 e 20
#lista_guia4_versao_atual[[10]][[3]] <- lista_guia4_versao_atual[[10]][[3]][,c(1:6, 13)]
#lista_guia4_versao_atual[[20]][[3]] <- lista_guia4_versao_atual[[20]][[3]][,c(1:6, 13)]

# Acertar nomes das colunas
nomes <- names(lista_guia4_versao_atual[[1]][[4]] )

i <- 1
while(i <= length(lista_guia4_versao_atual)){
  names(lista_guia4_versao_atual[[i]][[4]] ) <- nomes
  i <- i + 1
}

nomes <- names(lista_guia4_versao_menos2[[1]][[4]] )

i <- 1
while(i <= 13){
  names(lista_guia4_versao_menos2[[i]][[4]] ) <- nomes
  i <- i + 1
}


## -----------------------------------------------------------------------------------------------------------
# i <- 1

# for(i in 1:length(lista_guia4_versao_atual)){
#  print(i)
#  print(lista_guia4_versao_atual[[i]][[3]] |> ncol())
#i <- 1 + 1
#  }

# Erro nos dataframes 10 e 20


## -----------------------------------------------------------------------------------------------------------
maiores_pocos_terrestres_petroleo <- bind_rows(
do.call(rbind, map(lista_guia4_versao_atual, `[[`, 4)),

do.call(rbind, map(lista_guia4_versao_menos1, `[[`, 4)),  

do.call(bind_rows, map(lista_guia4_versao_menos2, `[[`, 4))  
)
   


## -----------------------------------------------------------------------------------------------------------
maiores_pocos_terrestres_petroleo |> 
  distinct() |> 
  write.csv2(file = "maiores_pocos_terrestres_petroleo.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(5)

read_excel(paste0("data/", planilhas[32]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(5)

read_excel(paste0("data/", planilhas[4]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(5)


## -----------------------------------------------------------------------------------------------------------
# Retirar colunas vazias nos dataframes 10 e 20
#lista_guia4_versao_atual[[10]][[3]] <- lista_guia4_versao_atual[[10]][[3]][,c(1:6, 13)]
#lista_guia4_versao_atual[[20]][[3]] <- lista_guia4_versao_atual[[20]][[3]][,c(1:6, 13)]

# Acertar nomes das colunas
nomes <- names(lista_guia4_versao_atual[[1]][[5]] )

i <- 1
while(i <= length(lista_guia4_versao_atual)){
  names(lista_guia4_versao_atual[[i]][[5]] ) <- nomes
  i <- i + 1
}

nomes <- names(lista_guia4_versao_menos2[[1]][[5]] )

i <- 1
while(i <= 13){
  names(lista_guia4_versao_menos2[[i]][[5]] ) <- nomes
  i <- i + 1
}


## -----------------------------------------------------------------------------------------------------------
# i <- 1

# for(i in 1:length(lista_guia4_versao_atual)){
#  print(i)
#  print(lista_guia4_versao_atual[[i]][[3]] |> ncol())
#i <- 1 + 1
#  }

# Erro nos dataframes 10 e 20


## -----------------------------------------------------------------------------------------------------------
maiores_pocos_terrestres_gas_natural <- bind_rows(
do.call(rbind, map(lista_guia4_versao_atual, `[[`, 5)),

do.call(rbind, map(lista_guia4_versao_menos1, `[[`, 5)),  

do.call(bind_rows, map(lista_guia4_versao_menos2, `[[`, 5))  
)
   


## -----------------------------------------------------------------------------------------------------------
maiores_pocos_terrestres_gas_natural |> 
  distinct() |> 
  write.csv2(file = "maiores_pocos_terrestres_gas_natural.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(6)

read_excel(paste0("data/", planilhas[32]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(6)

read_excel(paste0("data/", planilhas[4]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(6)


## -----------------------------------------------------------------------------------------------------------
# Corrigir os nomes das colunas

i <- 1
while(i <= length(lista_guia4_versao_atual)){
  names(lista_guia4_versao_atual[[i]][[6]]) <- names(lista_guia4_versao_atual[[i]][[6]]) |> 
  str_replace_all("Produtores", "produtores") |>
  str_replace_all("  ", " ")
  
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia4_versao_menos1)){
  names(lista_guia4_versao_menos1[[i]][[6]]) <- names(lista_guia4_versao_menos1[[i]][[6]]) |> 
  str_replace_all("Produtores", "produtores") |>
  str_replace_all("  ", " ")
  
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia4_versao_menos2)){
  names(lista_guia4_versao_menos2[[i]][[6]]) <- names(lista_guia4_versao_menos2[[i]][[6]]) |> 
  str_replace_all("Produtores", "produtores") |>
  str_replace_all("  ", " ")
  
  i <- i + 1
}



## -----------------------------------------------------------------------------------------------------------
maiores_instalacoes <- bind_rows(
bind_rows(map(lista_guia4_versao_atual, `[[`, 6)),
bind_rows(map(lista_guia4_versao_menos1, `[[`, 6)), 
bind_rows(map(lista_guia4_versao_menos2, `[[`, 6)) 
)



## -----------------------------------------------------------------------------------------------------------
maiores_instalacoes |> 
  distinct() |> 
  write.csv2(file = "maiores_instalacoes.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(7)

read_excel(paste0("data/", planilhas[32]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(7)

read_excel(paste0("data/", planilhas[4]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(7)


## -----------------------------------------------------------------------------------------------------------
# Verificar número de colunas

# i <- 1
# while(i <= length(lista_guia4_versao_atual)){
#  print(ncol(lista_guia4_versao_atual[[i]][[7]])) 
#  i <- i + 1
#}


## -----------------------------------------------------------------------------------------------------------
lista_guia4_versao_atual[[17]][[7]] <- lista_guia4_versao_atual[[17]][[7]][,c(1:5,14)]


## -----------------------------------------------------------------------------------------------------------
# Corrigir os nomes das colunas

i <- 1
while(i <= length(lista_guia4_versao_atual)){
  names(lista_guia4_versao_atual[[i]][[7]]) <- names(lista_guia4_versao_atual[[i]][[7]]) |> 
  str_replace_all("Produtores", "produtores") |>
  str_replace_all("  ", " ")
  
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia4_versao_menos1)){
  names(lista_guia4_versao_menos1[[i]][[7]]) <- names(lista_guia4_versao_menos1[[i]][[7]]) |> 
  str_replace_all("Produtores", "produtores") |>
  str_replace_all("  ", " ")
  
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia4_versao_menos2)){
  names(lista_guia4_versao_menos2[[i]][[7]]) <- names(lista_guia4_versao_menos2[[i]][[7]]) |> 
  str_replace_all("Produtores", "produtores") |>
  str_replace_all("Plataforma", "Instalação") |>  
  str_replace_all("  ", " ")
  
  i <- i + 1
}



## -----------------------------------------------------------------------------------------------------------
maiores_instalacoes_gas <- bind_rows(
bind_rows(map(lista_guia4_versao_atual, `[[`, 7)),
bind_rows(map(lista_guia4_versao_menos1, `[[`, 7)), 
bind_rows(map(lista_guia4_versao_menos2, `[[`, 7)) 
)



## -----------------------------------------------------------------------------------------------------------
maiores_instalacoes_gas |> 
  distinct() |> 
  write.csv2(file = "maiores_instalacoes_gas.csv")


## -----------------------------------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(8)

read_excel(paste0("data/", planilhas[32]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(8)

read_excel(paste0("data/", planilhas[4]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(8)

read_excel(paste0("data/", planilhas[1]), sheet = 4, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) |> 
  slice(8)


## -----------------------------------------------------------------------------------------------------------
# Verificar número de colunas

#i <- 1
#while(i <= length(lista_guia4_versao_atual)){
#  print(ncol(lista_guia4_versao_atual[[i]][[8]])) 
#  i <- i + 1
#}


## -----------------------------------------------------------------------------------------------------------
# Corrigindo dataframe 20, com colunas sem conteúdo
lista_guia4_versao_atual[[20]][[8]] <- lista_guia4_versao_atual[[20]][[8]][,c(1:6,13)] 

## -----------------------------------------------------------------------------------------------------------
# Verificar número de colunas

# i <- 1
# while(i <= length(lista_guia4_versao_menos1)){
#  print(i)
#  print(ncol(lista_guia4_versao_menos1[[i]][[8]])) 
#  i <- i + 1
#}


## -----------------------------------------------------------------------------------------------------------
# Corrigindo dataframe 17, com colunas sem conteúdo
lista_guia4_versao_menos1[[17]][[8]] <- lista_guia4_versao_menos1[[17]][[8]][,c(1:6,12)] 


## -----------------------------------------------------------------------------------------------------------
# Corrigir os nomes das colunas

i <- 1
while(i <= length(lista_guia4_versao_atual)){
  names(lista_guia4_versao_atual[[i]][[8]]) <- names(lista_guia4_versao_atual[[i]][[8]]) |> 
  str_replace_all("Produtores", "produtores") |>
  str_replace_all("  ", " ")
  
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia4_versao_menos1)){
  names(lista_guia4_versao_menos1[[i]][[8]]) <- names(lista_guia4_versao_menos1[[i]][[8]]) |> 
  str_replace_all("Produtores", "produtores") |>
  str_replace_all("  ", " ")
  
  i <- i + 1
}

i <- 1
while(i <= length(lista_guia4_versao_menos2)){
  names(lista_guia4_versao_menos2[[i]][[8]]) <- names(lista_guia4_versao_menos2[[i]][[8]]) |> 
  str_replace_all("Produtores", "produtores") |>
  str_replace_all("Plataforma", "Instalação") |>  
  str_replace_all("  ", " ")
  
  i <- i + 1
}



## -----------------------------------------------------------------------------------------------------------
maiores_queimas <- bind_rows(
bind_rows(map(lista_guia4_versao_atual, `[[`, 8)),
bind_rows(map(lista_guia4_versao_menos1, `[[`, 8)), 
bind_rows(map(lista_guia4_versao_menos2, `[[`, 8)) 
)



## -----------------------------------------------------------------------------------------------------------
maiores_queimas |> 
  distinct() |> 
  write.csv2(file = "maiores_queimas.csv")

