## ----setup, include=FALSE----------------------------------------------------------
knitr::opts_chunk$set(echo = TRUE)


## ----------------------------------------------------------------------------------
library(tidyverse)
library(readxl)


## ----------------------------------------------------------------------------------
source(knitr::purl("funcoes_data_e_extrair.Rmd", quiet=TRUE))


## ----------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 2, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) 


## ----message=FALSE-----------------------------------------------------------------
lista_guia2_versao_atual <- map(4:76, extrair_4_tabelas, sheet = 2)

## ----message=FALSE-----------------------------------------------------------------
lista_guia2_versao_meno3 <- map(1:3, extrair_3_tabelas, sheet = 2)

# gather(key = "mes", value = "produto", -1)


## ----message=FALSE-----------------------------------------------------------------
lista_guia2_versao_meno3 <- map(1:3, extrair_3_tabelas_data, sheet = 4)


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
Historico_petroleo <- bind_rows(map(lista_guia2_versao_atual, `[[`, 1)) |> 
  distinct() |> 
  mutate(mes = if_else(condition = str_detect(mes, "/"), true = my(mes), false = as.Date(as.numeric(mes), origin = "1899-12-30"))) 


## ----------------------------------------------------------------------------------
map(lista_guia2_versao_meno3, `[[`, 1)


## ----------------------------------------------------------------------------------
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



## ----------------------------------------------------------------------------------
Historico_prod_gas <- bind_rows(map(lista_guia2_versao_atual, `[[`, 2)) |> 
  distinct() |> 
  mutate(mes = my(mes))


## ----------------------------------------------------------------------------------
map(lista_guia2_versao_meno3, `[[`, 2)


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
df1 <- bind_rows( map(lista_guia2_versao_atual, `[[`, 3))

i <- 1
for(i in 1:length(lista_guia2_versao_atual)){
  names(lista_guia2_versao_atual[[i]][[3]])[1] <- "Período"
  i <- i + 1
}

# names(lista_guia2_versao_atual[[3]][[3]])[1]


## ----------------------------------------------------------------------------------


MOVIMENTACAO_DE_GAS_NATURAL <-  rbind(bind_rows( map(lista_guia2_versao_atual, `[[`, 3)) |>
  filter(str_detect(Período, "/")) |> 
  mutate(Período = my(Período)),

bind_rows( map(lista_guia2_versao_atual, `[[`, 3)) |>
  filter(!str_detect(Período, "/")) |> 
  mutate(Período = as.Date(as.numeric(Período), origin = "1899-12-30"))
) |> 
  distinct()



## ----------------------------------------------------------------------------------
MOVIMENTACAO_DE_GAS_NATURAL <- MOVIMENTACAO_DE_GAS_NATURAL |> 
  select("Período", c(1:5,7:9)) |> 
  mutate_at(2:8, as.numeric)


## ----------------------------------------------------------------------------------
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



## ----------------------------------------------------------------------------------
prov_df <- lista_guia2_versao_atual[[28]][[4]] 



## ----------------------------------------------------------------------------------
lista_guia2_nova <- list()
i <- i 
for(i in c(3, 5, 9:20)){
  lista_guia2_nova[[i]] <- lista_guia2_versao_atual[[i]][[4]]
  i < i + 1
}

prd_boe1 <- bind_rows(lista_guia2_nova) |> 
  distinct() |> 
  mutate(mes = my(mes))



## ----------------------------------------------------------------------------------
lista_guia2_nova2 <- list()

i <- i 
for(i in c(1, 2, 4, 6:8, 21:27, 29:length(lista_guia2_versao_atual))){
  lista_guia2_nova2[[i]] <- lista_guia2_versao_atual[[i]][[4]]
  i < i + 1
}


prd_boe2 <- bind_rows(lista_guia2_nova2) |> 
  distinct() |> 
  mutate(mes = my(mes))

## ----------------------------------------------------------------------------------
Historico_de_Producao_de_Petroleo_Gas_Natural <- prd_boe1 |> 
  rbind(prd_boe2) |> 
  distinct()


## ----------------------------------------------------------------------------------
map(lista_guia2_versao_meno3, `[[`, 3)


## ----------------------------------------------------------------------------------
read_excel(paste0("data/", planilhas[76]), sheet = 3, col_types = "text",  .name_repair = "unique_quiet") |> 
  filter(if_any(1, ~ str_detect(., "Tabela"))) |> 
  select(1) 


## ----message=FALSE-----------------------------------------------------------------
lista_guia3_versao_atual <- map(57:76, extrair_6_tabelas, sheet = 3)


## ----message=FALSE-----------------------------------------------------------------
lista_guia3_versao_menos1 <- map(4:56, extrair_7_tabelas, sheet = 3)


## ----------------------------------------------------------------------------------
lista_guia3_versao_menos3 <- map(1:3, extrair_13_tabelas, sheet = 3)


## ----message=FALSE-----------------------------------------------------------------
lista_guia4_versao_menos3 <- map(1:3, extrair_3_tabelas_data, sheet = 4)


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
por_estado <- rbind(
do.call(rbind, map(lista_guia3_versao_atual, `[[`, 1)),

do.call(rbind, map(lista_guia3_versao_menos1, `[[`, 1)),

do.call(rbind, map(lista_guia3_versao_menos3, `[[`, 1))  
) |> distinct()


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
por_bacia <- rbind(
do.call(rbind, map(lista_guia3_versao_atual, `[[`, 2)),

do.call(rbind, map(lista_guia3_versao_menos1, `[[`, 2)),

do.call(rbind, map(lista_guia3_versao_menos3, `[[`, 2))  
) |> distinct()  


## ----------------------------------------------------------------------------------
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


## ----eval=FALSE, include=FALSE-----------------------------------------------------
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


## ----------------------------------------------------------------------------------
por_operador <- rbind(
do.call(rbind, map(lista_guia3_versao_atual, `[[`, 3)),

do.call(rbind, map(lista_guia3_versao_menos1, `[[`, 3)),

do.call(rbind, map(lista_guia3_versao_menos3, `[[`, 3))  
) |> distinct() 

por_operador <- por_operador[,-1]


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------

# Retirar coluna Nº, que não tem informação e só prejudica



i <- 1
for(i in seq_along(lista_guia3_versao_atual)){
lista_guia3_versao_atual[[i]][[4]] <- select(lista_guia3_versao_atual[[i]][[4]], !"Nº")
i <- i + 1
}
  
i <- 1
for(i in seq_along(lista_guia3_versao_menos1)){
lista_guia3_versao_menos1[[i]][[4]] <- select(lista_guia3_versao_menos1[[i]][[4]], !"Nº")
i <- i + 1
}


i <- 1
for(i in seq_along(lista_guia3_versao_menos3)){
lista_guia3_versao_menos3[[i]][[4]] <- select(lista_guia3_versao_menos3[[i]][[4]], !"Nº")
i <- i + 1
}



## ----------------------------------------------------------------------------------
# Apagar colunas que estão sobrando (do dataframe 16)
 i <- 1
 while(i <= length(lista_guia3_versao_atual)){
  print(ncol(lista_guia3_versao_atual[[i]][[4]]))
  i <- i + 1
}

lista_guia3_versao_atual[[17]][[4]] <- lista_guia3_versao_atual[[17]][[4]] |> 
  select(!5:9)

# Acertar nomes de colunas
names(lista_guia3_versao_atual[[17]][[4]]) <- names(lista_guia3_versao_atual[[16]][[4]])


## ----------------------------------------------------------------------------------
por_concessionario <- rbind(
map_df(lista_guia3_versao_atual, `[[`, 4),
map_df(lista_guia3_versao_menos1, `[[`, 4),
map_df(lista_guia3_versao_menos3, `[[`, 4)
)



## ----------------------------------------------------------------------------------
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



## ----------------------------------------------------------------------------------
# Dataframe 16 tem colunas em branco, detectando e acertando isso
#i <- 1

#while(i <= length(lista_guia3_versao_atual)){
#  print(i)
#  names(lista_guia3_versao_atual[[i]][[4]])  |> 
#    print()

#    i <- i + 1
#}

# lista_guia3_versao_atual[[16]][[4]] <- lista_guia3_versao_atual[[16]][[4]][,c(1:5, 11)]


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
MOVIMENTACAO_GAS_NATURAL_POR_DESTINO <- bind_rows(
do.call(bind_rows, map(lista_guia3_versao_atual, `[[`, 5)),

do.call(rbind, map(lista_guia3_versao_menos1, `[[`, 5)),

do.call(rbind, map(lista_guia4_versao_menos3, `[[`, 2))  
) |> distinct()


## ----------------------------------------------------------------------------------
MOVIMENTACAO_GAS_NATURAL_POR_DESTINO <- MOVIMENTACAO_GAS_NATURAL_POR_DESTINO |> 
  select(c(1:5, 8, 6, 9, 7))


## ----------------------------------------------------------------------------------
# names(MOVIMENTACAO_GAS_NATURAL_POR_DESTINO)
# names(tabela9)


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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



## ----------------------------------------------------------------------------------
# i <- 1
  
# for(i in seq_along(lista_guia3_versao_atual)){
#lista_guia3_versao_atual[[i]][[6]] |> 
#  ncol() |> print()
#
#i <- 1 + 1  
#}

# apagando colunas vazias
lista_guia3_versao_atual[[1]][[6]] <- lista_guia3_versao_atual[[1]][[6]] |> 
  select(c(1:7, 11))

# Colocando em todos os dfs os mesmos nomes de colunas

i <- 1
for( i in seq_along(lista_guia3_versao_atual)){
  names(lista_guia3_versao_atual[[i]][[6]]) <- names(lista_guia3_versao_atual[[20]][[6]])
  i <- 1
}



## ----------------------------------------------------------------------------------

# Se a primeira linha da primeira coluna for nula, apagar linha e converter todas as colunas para valores numéricos
i <- 1

for(i in seq_along(lista_guia3_versao_atual)) {
if(is.na(lista_guia3_versao_atual[[i]][[6]][1,1])) {lista_guia3_versao_atual[[i]][[6]] <- lista_guia3_versao_atual[[i]][[6]][-1,] |> 
  mutate_at(2:7, as.numeric)}    
  i <- 1 + i
}



## ----------------------------------------------------------------------------------

# Colocando em todos os dfs os mesmos nomes de colunas

i <- 1
for( i in seq_along(lista_guia3_versao_menos1)){
  names(lista_guia3_versao_menos1[[i]][[6]]) <- names(lista_guia3_versao_atual[[20]][[6]])
  i <- 1
}

map_df(lista_guia3_versao_menos1, `[[`, 6)


## ----------------------------------------------------------------------------------
# Colocando em todos os dfs os mesmos nomes de colunas

i <- 1
for( i in seq_along(lista_guia3_versao_menos3)){
  names(lista_guia3_versao_menos3[[i]][[6]]) <- names(lista_guia3_versao_atual[[20]][[6]])
  i <- 1
}

map_df(lista_guia3_versao_menos3, `[[`, 6)


## ----------------------------------------------------------------------------------
DISTRIBUICAO_DA_PRODUCAO_CAMPOS_PRE_SAL <- bind_rows(
map_df(lista_guia3_versao_atual, `[[`, 6),
map_df(lista_guia3_versao_menos1, `[[`, 6),
map_df(lista_guia3_versao_menos3, `[[`, 6)  
) |> distinct()


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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




## ----message=FALSE-----------------------------------------------------------------
lista_guia4_versao_atual <- map(c(55,56, 58:76), extrair_8_tabelas, sheet = 4)


## ----------------------------------------------------------------------------------
lista_guia4_versao_menos1 <- map(c(33:54, 57), extrair_16_tabelas, sheet = 4)


## ----message=FALSE-----------------------------------------------------------------
lista_guia4_versao_menos2 <- map(4:32, extrair_8_tabelas, sheet = 4)


## ----message=FALSE-----------------------------------------------------------------
lista_guia4_versao_menos3 <- map(1:3, extrair_3_tabelas_data, sheet = 4)


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
maiores_pocos_de_petroleo <- bind_rows(
do.call(rbind, map(lista_guia4_versao_atual, `[[`, 1)),

do.call(rbind, map(lista_guia4_versao_menos1, `[[`, 1)),  

do.call(bind_rows, map(lista_guia4_versao_menos2, `[[`, 1))  
) |> distinct() 


maiores_pocos_de_petroleo <- maiores_pocos_de_petroleo[,-1]


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
maiores_pocos_de_gas <- bind_rows(
do.call(rbind, map(lista_guia4_versao_atual, `[[`, 2)),

do.call(rbind, map(lista_guia4_versao_menos1, `[[`, 2)),  

do.call(bind_rows, map(lista_guia4_versao_menos2, `[[`, 2))  
) |> distinct() 

maiores_pocos_de_gas <- maiores_pocos_de_gas[,-1]   


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
# i <- 1

# for(i in 1:length(lista_guia4_versao_atual)){
#  print(i)
#  print(lista_guia4_versao_atual[[i]][[3]] |> ncol())
#i <- 1 + 1
#  }

# Erro nos dataframes 10 e 20


## ----------------------------------------------------------------------------------
maiores_pocos_producao_total <- bind_rows(
do.call(rbind, map(lista_guia4_versao_atual, `[[`, 3)),

do.call(rbind, map(lista_guia4_versao_menos1, `[[`, 3)),  

do.call(bind_rows, map(lista_guia4_versao_menos2, `[[`, 3))  
) |> distinct() 

maiores_pocos_producao_total <- maiores_pocos_producao_total[,-1]
   


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
# i <- 1

# for(i in 1:length(lista_guia4_versao_atual)){
#  print(i)
#  print(lista_guia4_versao_atual[[i]][[3]] |> ncol())
#i <- 1 + 1
#  }

# Erro nos dataframes 10 e 20


## ----------------------------------------------------------------------------------
maiores_pocos_terrestres_petroleo <- bind_rows(
do.call(rbind, map(lista_guia4_versao_atual, `[[`, 4)),

do.call(rbind, map(lista_guia4_versao_menos1, `[[`, 4)),  

do.call(bind_rows, map(lista_guia4_versao_menos2, `[[`, 4))  
) |> distinct() 

maiores_pocos_terrestres_petroleo <- maiores_pocos_terrestres_petroleo[,-1]   


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
# i <- 1

# for(i in 1:length(lista_guia4_versao_atual)){
#  print(i)
#  print(lista_guia4_versao_atual[[i]][[3]] |> ncol())
#i <- 1 + 1
#  }

# Erro nos dataframes 10 e 20


## ----------------------------------------------------------------------------------
maiores_pocos_terrestres_gas_natural <- bind_rows(
do.call(rbind, map(lista_guia4_versao_atual, `[[`, 5)),

do.call(rbind, map(lista_guia4_versao_menos1, `[[`, 5)),  

do.call(bind_rows, map(lista_guia4_versao_menos2, `[[`, 5))  
) |> distinct()

maiores_pocos_terrestres_gas_natural <- maiores_pocos_terrestres_gas_natural[,-1]
   


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
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



## ----------------------------------------------------------------------------------
#Acertando os nomes das colunas
bind_rows(map(lista_guia4_versao_menos2, `[[`, 6))


instalacao <- names(lista_guia4_versao_menos2[[1]][[6]])[1]

i <- 1
for(i in 1:length(lista_guia4_versao_menos2)){
names(lista_guia4_versao_menos2[[i]][[6]])[1]  <- instalacao
  i <- i + 1
}


## ----------------------------------------------------------------------------------
maiores_instalacoes <- bind_rows(
bind_rows(map(lista_guia4_versao_atual, `[[`, 6)),
bind_rows(map(lista_guia4_versao_menos1, `[[`, 6)), 
bind_rows(map(lista_guia4_versao_menos2, `[[`, 6)) 
) |> distinct() 

maiores_instalacoes <- maiores_instalacoes[,-1]


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
# Verificar número de colunas

# i <- 1
# while(i <= length(lista_guia4_versao_atual)){
#  print(ncol(lista_guia4_versao_atual[[i]][[7]])) 
#  i <- i + 1
#}


## ----------------------------------------------------------------------------------
lista_guia4_versao_atual[[17]][[7]] <- lista_guia4_versao_atual[[17]][[7]][,c(1:5,14)]


## ----------------------------------------------------------------------------------
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



## ----------------------------------------------------------------------------------
maiores_instalacoes_gas <- bind_rows(
bind_rows(map(lista_guia4_versao_atual, `[[`, 7)),
bind_rows(map(lista_guia4_versao_menos1, `[[`, 7)), 
bind_rows(map(lista_guia4_versao_menos2, `[[`, 7)) 
) |> distinct() 

maiores_instalacoes_gas <- maiores_instalacoes_gas[,-1]


## ----------------------------------------------------------------------------------
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


## ----------------------------------------------------------------------------------
# Verificar número de colunas

i <- 1
while(i <= length(lista_guia4_versao_atual)){
  print(ncol(lista_guia4_versao_atual[[i]][[8]])) 
  i <- i + 1
}
# corrigir o df 20

lista_guia4_versao_atual[[20]][[8]] <- lista_guia4_versao_atual[[20]][[8]] |> 
  select(!7:12)



## ----------------------------------------------------------------------------------
#i <- 1
#while(i <= length(lista_guia4_versao_menos1)){
#  print(ncol(lista_guia4_versao_menos1[[i]][[8]])) 
#  i <- i + 1
#}
# corrigindo dataframe 17
lista_guia4_versao_menos1[[18]][[8]] <- lista_guia4_versao_menos1[[18]][[8]] |> 
  select(!7:11)


## ----------------------------------------------------------------------------------
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



## ----------------------------------------------------------------------------------
maiores_queimas <- bind_rows(
bind_rows(map(lista_guia4_versao_atual, `[[`, 8)),
bind_rows(map(lista_guia4_versao_menos1, `[[`, 8)), 
bind_rows(map(lista_guia4_versao_menos2, `[[`, 8)) 
) |> distinct() |> 
  select(!`Nº`)



## ----------------------------------------------------------------------------------
Historico_petroleo |> 
  write_csv("csv/tabela01_Historico_petroleo.csv", )

Historico_prod_gas |> 
  write_csv("csv/tabela02_Historico_prod_gas.csv")

MOVIMENTACAO_DE_GAS_NATURAL |> 
  write_csv("csv/tabela03_MOVIMENTACAO_DE_GAS_NATURAL.csv")

Historico_de_Producao_de_Petroleo_Gas_Natural |> 
  write_csv("csv/tabela04_Historico_de_Producao_de_Petroleo_Gas_Natural.csv")

por_estado |> 
  write_csv("csv/tabela05_Producao_por_estado.csv")

por_bacia |> 
  write_csv("csv/tabela06_Producao_por_bacia.csv")

por_operador |> 
  write_csv("csv/tabela07_Producao_por_operador.csv")

por_concessionario |> 
  write_csv("csv/tabela08_Producao_por_concessionario.csv")

MOVIMENTACAO_GAS_NATURAL_POR_DESTINO |> 
  write_csv("csv/tabela09_MOVIMENTACAO_GAS_NATURAL_POR_DESTINO.csv")

DISTRIBUICAO_DA_PRODUCAO_CAMPOS_PRE_SAL |> 
  write_csv("csv/tabela10_DISTRIBUICAO_DA_PRODUCAO_CAMPOS_PRE_SAL.csv")

maiores_pocos_de_petroleo |> 
  write_csv("csv/tabela11_maiores_pocos_de_petroleo.csv")

maiores_pocos_de_gas |> 
  write_csv("csv/tabela12_maiores_pocos_de_gas.csv")

maiores_pocos_producao_total |> 
  write_csv("csv/tabela13_maiores_pocos_producao_total.csv")

maiores_pocos_terrestres_petroleo |> 
  write_csv("csv/tabela14_maiores_pocos_terrestres_petroleo.csv")

maiores_pocos_terrestres_gas_natural |> 
  write_csv("csv/tabela15_maiores_pocos_terrestres_gas_natural.csv")

maiores_instalacoes |> 
  write_csv("csv/tabela16_maiores_instalações.csv")
  
maiores_instalacoes_gas |> 
  write_csv("csv/tabela17_maiores_instalacoes_gas.csv")

maiores_queimas |> 
write_csv("csv/tabela18_maiores_queimas.csv")

