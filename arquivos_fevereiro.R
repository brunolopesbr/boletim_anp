## ----setup, include=FALSE-----------------------------------------------------
knitr::opts_chunk$set(echo = TRUE)

## -----------------------------------------------------------------------------
library(tidyverse)


## -----------------------------------------------------------------------------
arquivos_fev_2024 <- list.files(path = "fevereiro", pattern = "tsv")
arquivos_fev_2024 <- paste0("fevereiro/", arquivos_fev_2024)


## ----warning=FALSE------------------------------------------------------------
tabelas_fev_2024 <- lapply(arquivos_fev_2024, function(x) {read_delim(x, 
    delim = "\t", trim_ws = TRUE)})  
tabelas_fev_2024


## -----------------------------------------------------------------------------
for(i in c(1,2,4)){
tabelas_fev_2024[[i]] <-  tabelas_fev_2024[[i]] |> 
  pivot_longer(names_to = "Período", values_to = "Valor", -1) |> 
  mutate(Período = my(Período)) 

names(tabelas_fev_2024[[i]])[1] <- "Produto"  
}



## -----------------------------------------------------------------------------
tabelas_fev_2024[[3]] <- tabelas_fev_2024[[3]] |> 
  mutate(Aproveitamento = str_remove_all(Aproveitamento, "%")) |> 
  mutate(Aproveitamento = str_replace_all(Aproveitamento, ",", ".")) |> 
  mutate(Aproveitamento = as.numeric(Aproveitamento)) |> 
  mutate(Aproveitamento = Aproveitamento/100) |> 
  mutate("Mês/Ano" = my(`Mês/Ano`))


## -----------------------------------------------------------------------------
boletim_fev <- as.Date("2024-02-01")

for(i in 5:18){
tabelas_fev_2024[[i]] <- tabelas_fev_2024[[i]] |> 
  mutate(mes = as.Date(boletim_fev))
}


## -----------------------------------------------------------------------------
for(i in seq_along(tabelas_fev_2024)){
  write.csv(tabelas_fev_2024[[i]], file = paste0("fevereiro/tab_fev", i, ".csv"), row.names = FALSE )
}
 
 

