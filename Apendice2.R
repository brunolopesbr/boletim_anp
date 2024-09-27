#datasets::AirPassengers |> 
#  write.csv("AirPassengers.csv")

#datasets::mtcars |> 
#  write.csv("mtcars.csv")

#datasets::iris |> 
#  write.csv("iris.csv")

library(readxl)
library(tidyverse)

plan_exemplo <- paste0("exemplo/", list.files(path = "exemplo", pattern = ".xlsx"))

guias <- excel_sheets(plan_exemplo) 

lapply(guias, function(x){read_xlsx(plan_exemplo, sheet = x)})
