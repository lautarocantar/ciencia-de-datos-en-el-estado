


# Cargando paquetes a utilizar
library(tidyverse)


# ------------------------------------------------------------------------------------------
#                                  1. Funciones
# ------------------------------------------------------------------------------------------

trim <- function (x) gsub("^\\s+|\\s+$", "", x)  

# Funciones
codigo_verificado <- function(prefijo, DNI){
  
  numero <- paste0(prefijo, DNI)
  numero_individual <- str_split(numero, "")
  
  primero <- as.numeric(numero_individual[[1]][1]) * 5
  segundo <- as.numeric(numero_individual[[1]][2]) * 4
  tercero <- as.numeric(numero_individual[[1]][3]) * 3
  cuarto  <- as.numeric(numero_individual[[1]][4]) * 2
  quinto  <- as.numeric(numero_individual[[1]][5]) * 7
  sexto   <- as.numeric(numero_individual[[1]][6]) * 6
  septimo <- as.numeric(numero_individual[[1]][7]) * 5
  octavo  <- as.numeric(numero_individual[[1]][8]) * 4
  noveno  <- as.numeric(numero_individual[[1]][9]) * 3
  decimo  <- as.numeric(numero_individual[[1]][10]) * 2
  
  total <- sum(primero, segundo, tercero, cuarto, quinto, sexto, septimo, octavo, noveno, decimo)
  
  modulo <- total %% 11
  once_menos <- 11 - modulo
  
  return(once_menos)
}  

# ------------------------------------------------------------------------------------------
#                                  2. Data Manipulation
# ------------------------------------------------------------------------------------------

# Cambiando las variables del CUIT
dat$CUIT.prefijo <- dat$`Cuit/Cuil`  %>% str_extract("[0-9]{2}")
dat$DNI <- dat$`Cuit/Cuil` %>% str_sub(start = 3, end = 10)
dat$CUIT.sufijo <- dat$`Cuit/Cuil`  %>% str_sub(start = 11, end = 11)

# Asignando el Sexo
for (i in 1:nrow(academia)) {
  
  print(i)
  if (dat$CUIT.prefijo[i] == 20 | dat$CUIT.prefijo[i] == 27) {
    dat$Sexo[i] <- ifelse(dat$CUIT.prefijo[i] == 20, "Hombre",
                               ifelse(dat$CUIT.prefijo[i] == 27, "Mujer", NA))
  } else if (dat$CUIT.prefijo[i] == 23) {
    dat$Sexo[i] <- ifelse(codigo_verificado(20, dat$DNI[i]) == 10, "Hombre",
                               ifelse(codigo_verificado(27, dat$DNI[i]) == 10, "Mujer", NA))
  } else {
    dat$Sexo[i] <- NA
  }
}

