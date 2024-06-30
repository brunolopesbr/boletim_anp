datasets::AirPassengers |> 
  write.csv("AirPassengers.csv")

datasets::mtcars |> 
  write.csv("mtcars.csv")

datasets::iris |> 
  write.csv("iris.csv")

setwd("csv")
import_files <- dir(pattern = "*.csv")
map(import_files, read.csv)

guias <- excel_sheets("planilha.xlsx") 

read_planilha <- function(name) {
  read_xlsx("planilha.xlsx", sheet = name)
  }

map(guias, read_planilha)
