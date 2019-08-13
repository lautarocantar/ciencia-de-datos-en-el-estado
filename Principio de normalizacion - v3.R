
# Proceso de Normalización de Unidades Territoriales y Coordenadas v3

# Lauti Cantar 
# Contacto:   lcantar@mininterior.gob.ar 
#             lautaro.cantar.ar@gmail.com

# ----------------------------------------------------------------------------------
#                                 0. Cargando paquetes y data
# ----------------------------------------------------------------------------------

t0 <- Sys.time()

# Cargando los paquetes
library(readxl)
library(tidyverse)
library(leaflet)
library(sf)
library(sp)
library(georefar)
library(rgdal)



dat <- read_excel("N:/JG_Planificacion/6. BASE MENSUAL DE OBRAS/0.4. RUTINA PLAN B/03- 2019/06. 30.06/BASE MIOPyV 30.06/BaseMIOPyV-ACTIVA+TERMINADA+OBRASSIU2_30.06 - FINAL.xlsx",
                  sheet = 1,
                  col_types = c(IDCONCATENADO_STATA = "text",
                                ID_UNICO_MIOPyV = "text",
                                CARTERA = "text",
                                FechaDeProceso = "text",
                                AñoDeFinalizacion = "text",
                                OrigenDato = "text",
                                ResponsableDato = "text",
                                NIVEL1 = "text",
                                NIVEL2 = "text",
                                PROCREAR_IDProyecto = "numeric",
                                PROCREAR_IDSector = "text",
                                Conformacion_SIU = "text",
                                Tipologia_SIU = "text",
                                ID_Obra = "numeric",
                                NumeroDeObra = "text",
                                NroExpediente = "text",
                                acu = "text",
                                acu_addenda1 = "text",
                                acu_addenda2 = "text",
                                ID_FichaUnicaProyecto = "numeric",
                                Codigo_SIPPE = "text",
                                Codigo_BAPIN = "text",
                                AreaDeProyecto = "text",
                                FechaFirmaContrato = "text",
                                FechaInicioPrevista = "text",
                                FechaInicio = "text",
                                FechaActaInicio = "text",
                                FinDeObraPrevista = "text",
                                FinDeObra = "text",
                                FechaFinalizacion = "text",
                                TipoDeObra = "text",
                                NombreDeObra = "text",
                                Descripcion = "text",
                                Región = "text",
                                ID_Provincia = "numeric",
                                Provincia = "text",
                                DepartamentoPartidoComuna = "text",
                                DepartamentoPartidoComuna2 = "text",
                                LocalidadAsentamientoHumano = "text",
                                Municipio = "text",
                                Conurbano = "text",
                                EtapaDeObra = "text",
                                estado_original = "text",
                                PlazoDeObra = "numeric",
                                Viviendas = "numeric",
                                Mejoramientos = "numeric",
                                Ejecutor = "text",
                                Beneficiarios = "numeric",
                                Fuente = "text",
                                MontoTotal = "numeric",
                                MontoTotalPagado = "numeric",
                                MontoTotalPagadoConCertPagasONo = "numeric",
                                MontoTotalPagado_2016 = "numeric",
                                MontoTotalPagado_2017 = "numeric",
                                MontoTotalPagado_2018 = "numeric",
                                MontoTotalPagado_2019 = "numeric",
                                MontoTotalCertificado = "numeric",
                                FechaUltimoPago = "numeric",
                                SaldoTotal = "numeric",
                                SaldoTotal_2018 = "numeric",
                                SaldoTotal_2019 = "numeric",
                                SaldoTotal_2020 = "numeric",
                                SaldoTotal_2021 = "numeric",
                                SaldoRestoTotal = "numeric",
                                MontoTotalDevengado = "numeric",
                                SaldoTotalADevengar = "numeric",
                                Porcentaje_Avance_Fisico = "numeric",
                                Porcentaje_Avance_Financiero = "numeric",
                                Observaciones = "text",
                                SV_Observaciones_ssduv = "text",
                                RU_Observaciones = "text",
                                observaciones_ssduv = "text",
                                FechaFinalizacionEstimada = "text",
                                Inaugurable_en = "text",
                                Observaciones_sobre_Inauguracion = "text",
                                Latitud = "numeric",
                                Longitud = "numeric",
                                SV_programa = "text",
                                SSCOPF_Programa = "text",
                                RU_Cantidad = "numeric",
                                RU_MontoTotal = "numeric",
                                Alias = "text",
                                CodigoDeLocalizacion = "numeric",
                                ACUMAR = "numeric",
                                MIN11 = "numeric",
                                MIN14 = "numeric",
                                MIN22 = "numeric",
                                MTOTAL = "numeric",
                                Monto_Nacion = "numeric",
                                Monto_Provincia = "numeric",
                                Monto_Municipio = "numeric",
                                MFFFIR = "numeric",
                                MFAUSTRAL = "numeric",
                                MCPARTE = "numeric",
                                MFondoHidrico = "numeric",
                                Préstamo = "text",
                                MontoJGM2019 = "numeric",
                                MontoJGM2020 = "numeric",
                                ESTADOJGM = "text",
                                Terminada_Financiero = "text",
                                Terminada_Legal = "text",
                                FechaRecepcionProvisoria = "text",
                                FechaRecepcionDefinitiva = "text",
                                Longitud_Posicion_Exacta = "numeric",
                                Latitud_Posicion_Exacta = "numeric",
                                SVYH_glat = "text",
                                SVYH_glng = "text",
                                SVYH_gmap = "text",
                                EtapaDeObraAREA = "text"
                  ))


dat <- dat %>% 
  mutate(ID_Lauti = rep(1:nrow(dat)))

# Funcion para cambiar el formato de la fecha
cambiar.fecha <- function(fecha.excel){
  as.Date(fecha.excel, 
          origin = "1899-12-30") %>% 
    format("%d/%m/%Y")
}

# Cambiando el formato de las fehcas
dat <- dat %>%
  mutate(FechaFirmaContrato	= FechaFirmaContrato %>% as.numeric(),
         FechaInicioPrevista = FechaInicioPrevista %>% as.numeric(),	
         FechaInicio = FechaInicio %>% as.numeric(),
         FechaActaInicio = FechaActaInicio %>% as.numeric(),
         FinDeObraPrevista = FinDeObraPrevista %>% as.numeric(),
         FinDeObra = FinDeObra %>% as.numeric(),
         FechaFinalizacion= FechaFinalizacion %>% as.numeric(),
         FechaFinalizacionEstimada = FechaFinalizacionEstimada %>% as.numeric(),
         Inaugurable_en = Inaugurable_en %>% as.numeric(),
         FechaRecepcionProvisoria = FechaRecepcionProvisoria %>% as.numeric(),	
         FechaRecepcionDefinitiva = FechaRecepcionDefinitiva %>% as.numeric()) %>%
  mutate(FechaFirmaContrato = cambiar.fecha(FechaFirmaContrato),
         FechaInicioPrevista = cambiar.fecha(FechaInicioPrevista),
         FechaInicio = cambiar.fecha(FechaInicio),
         FechaActaInicio = cambiar.fecha(FechaActaInicio),
         FinDeObraPrevista = cambiar.fecha(FinDeObraPrevista),
         FinDeObra = cambiar.fecha(FinDeObra),
         FechaFinalizacion = cambiar.fecha(FechaFinalizacion),
         FechaFinalizacionEstimada = cambiar.fecha(FechaFinalizacionEstimada),
         Inaugurable_en = cambiar.fecha(Inaugurable_en),
         FechaRecepcionProvisoria = cambiar.fecha(FechaRecepcionProvisoria),
         FechaRecepcionDefinitiva = cambiar.fecha(FechaRecepcionDefinitiva)
  )

back.0 <- dat

# --------------------------------------------------------------------------------
#                   Paso 1 – Primera ronda normalización de  nombres
# --------------------------------------------------------------------------------

# La base de datos contiene las siguientes variables:
#   •	Provincia
#   •	Departamento
#   •	Localidad
#   •	Municipio

# En caso de existir datos en cada una de esas variables, se normalizan 
# utilizando la API de GeoRef. Primero se normaliza el nombre de las provincias, 
# luego el nombre de los departamentos pertenecientes a cada una de las 
# provincias y por último los nombres de las localidades.
# 
# El resultado de esta normalización se agrega a la base original como nuevas 
# columnas, para no alterar procesos posteriores que utilicen la información 
# contenida en las variables originales.


# --------------- Normalizando Provincias --------------- ####

provincias.viejas <- dat %>% 
  select(Provincia) %>% 
  filter(Provincia != "INTERPROVINCIAL") %>% 
  distinct()


for (i in 1:nrow(provincias.viejas)) {
  
  p <- provincias.viejas$Provincia[i] %>% str_to_lower()
  
  query.prov <- get_provincias(nombre = p, max = 1)
  p.norm <- as.character(query.prov["nombre"])
  p.norm.lat <- as.numeric(query.prov["centroide_lat"])
  p.norm.long <- as.numeric(query.prov["centroide_lon"])
  
  provincias.viejas$Provincia.Normalizada[i] <- p.norm
  provincias.viejas$Provincia.Lat.Centroide[i] <- p.norm.lat
  provincias.viejas$Provincia.Lon.Centroide[i] <- p.norm.long
  
  Sys.sleep(1)
  print(p)
}

# openxlsx::write.xlsx(provincias.viejas, "02.Datasets_soporte/DT. Provincias Normalizados.xlsx")
# provincias.viejas <- read_excel("02.Datasets_soporte/DT. Provincias Normalizados.xlsx")

# Haciendo un join con la base original
dat <- dat %>% 
  left_join(provincias.viejas, by = "Provincia") %>% 
  mutate(Provincia.Normalizada = case_when(Provincia == "INTERPROVINCIAL" ~ "INTERPROVINCIAL",
                                           TRUE ~ Provincia.Normalizada))

# --------------- Normalizando departamentos --------------- ####

provincias <- dat %>%
  select(Provincia.Normalizada) %>% 
  filter(Provincia.Normalizada != "INTERPROVINCIAL") %>%
  distinct()

# Creo un nuevo dataset vacio
dptos.nuevos <- data.frame()

for (i in 1:nrow(provincias)) {
  p <- provincias$Provincia.Normalizada[i]
  print(p)
  
  dptos.viejos <- dat %>% 
    select(Provincia.Normalizada, DepartamentoPartidoComuna) %>% 
    distinct(Provincia.Normalizada, DepartamentoPartidoComuna) %>% 
    mutate(DepartamentoPartidoComuna_nueva = DepartamentoPartidoComuna %>% str_to_lower()) %>% 
    filter(Provincia.Normalizada == p)
  
  
  for (j in 1:nrow(dptos.viejos)) {
    
    d <- dptos.viejos$DepartamentoPartidoComuna_nueva[j]
    print(d)
    
    query.dpto <- get_departamentos(provincia = p,
                                    nombre = d,
                                    max = 1)
    
    d.norm <- as.character(query.dpto["nombre"])
    
    dptos.viejos$Departamento.Normalizado[j] <- d.norm
    
    if (d.norm != "NULL") {
      
      d.lat <- as.numeric(query.dpto["centroide_lat"])
      d.long <- as.numeric(query.dpto["centroide_lon"])
      
      dptos.viejos$Departamento.Lat.Centroide[j] <- d.lat
      dptos.viejos$Departamento.Lon.Centroide[j] <- d.long
      
    } else{
      
      dptos.viejos$Departamento.Lat.Centroide[j] <- NA
      dptos.viejos$Departamento.Lon.Centroide[j] <- NA
      
    }
    
    
    Sys.sleep(time = 0.5) 
  }
  
  dptos.nuevos <- rbind(dptos.nuevos, dptos.viejos)
  
}

# Guardando el dataset para que en el futuro sea mas sencillo
# openxlsx::write.xlsx(dptos.nuevos, "02.Datasets_soporte/DT. Provincias y Dptos Normalizados.xlsx")
# dptos.nuevos <- read_excel("02.Datasets_soporte/DT. Provincias y Dptos Normalizados.xlsx")

# Uniendo 
dat <- dat %>% 
  left_join(dptos.nuevos %>% 
              select(Provincia.Normalizada, DepartamentoPartidoComuna, 
                     Departamento.Normalizado, Departamento.Lat.Centroide, Departamento.Lon.Centroide), 
            by = c("Provincia.Normalizada", "DepartamentoPartidoComuna"))

# --------------- Normalizando Localidades --------------- ####

departamentos <- dat %>%
  select(Provincia.Normalizada, Departamento.Normalizado) %>% 
  filter(Provincia.Normalizada != "INTERPROVINCIAL") %>%
  distinct()

# Creo un nuevo dataset vacio
localidades.nuevos <- data.frame()


for (i in 1:nrow(departamentos)) {
  p <- departamentos$Provincia.Normalizada[i]
  d <- departamentos$Departamento.Normalizado[i]
  print(p)
  print(d)
  
  localidades.viejas <- dat %>% 
    select(Provincia.Normalizada, Departamento.Normalizado, LocalidadAsentamientoHumano) %>% 
    distinct(Provincia.Normalizada, Departamento.Normalizado, LocalidadAsentamientoHumano) %>% 
    filter(!is.na(LocalidadAsentamientoHumano)) %>% 
    mutate(Localidad_nueva = LocalidadAsentamientoHumano %>% str_to_lower()) %>% 
    filter(Provincia.Normalizada == p) %>% 
    filter(Departamento.Normalizado == d)
  
  
  if(nrow(localidades.viejas) > 0){
    
    
    for (j in 1:nrow(localidades.viejas)) {
      
      l <- localidades.viejas$Localidad_nueva[j]
      print(l)
      
      query.localidad <- get_localidades(nombre = l,
                                         provincia = p,
                                         departamento = d,
                                         max = 1)
      
      l.norm <- as.character(query.localidad["nombre"])
      
      localidades.viejas$Localidad.Normalizado[j] <- l.norm
      
      if (l.norm != "NULL") {
        
        id.localidad <- as.character(query.localidad["id"])
        l.lat <- as.numeric(query.localidad["centroide_lat"])
        l.lon <- as.numeric(query.localidad["centroide_lon"])
        
        localidades.viejas$Localidad.Lat.Centroide[j] <- l.lat
        localidades.viejas$Localidad.Lon.Centroide[j] <- l.lon
        localidades.viejas$ID.Localidad[j] <- id.localidad
        
      } else {
        
        localidades.viejas$Localidad.Lat.Centroide[j] <- NA
        localidades.viejas$Localidad.Lon.Centroide[j] <- NA
        localidades.viejas$ID.Localidad[j] <- NA
        
      }
      
      Sys.sleep(time = 1) 
      
      
    }
    
    localidades.nuevos <- rbind(localidades.nuevos, localidades.viejas)
  }
}


# openxlsx::write.xlsx(localidades.nuevos, "02.Datasets_soporte/DT. Provincias, Dptos y Localiaddes Normalizados.xlsx")
# localidades.nuevos <- read_excel("02.Datasets_soporte/DT. Provincias, Dptos y Localiaddes Normalizados.xlsx")

# Uniendo 
dat <- dat %>% 
  left_join(localidades.nuevos %>% 
              select(Provincia.Normalizada, 
                     Departamento.Normalizado,
                     LocalidadAsentamientoHumano, 
                     Localidad.Normalizado,
                     ID.Localidad,
                     Localidad.Lat.Centroide,
                     Localidad.Lon.Centroide
              ), 
            by = c("Provincia.Normalizada", 
                   "Departamento.Normalizado", 
                   "LocalidadAsentamientoHumano"))


# ----- Editando valores especiales que son conocidos ----- ####

# Le cambiamos el nombre a algunos departamentos que nosotros sabemos que no
# se van a encontrar

dat <- dat %>% 
  mutate(Departamento.Normalizado = case_when(DepartamentoPartidoComuna == "Varios" | DepartamentoPartidoComuna == "Varias" ~ "Varios",
                                              DepartamentoPartidoComuna == "Sistema Agua Norte" ~ "Sistema Agua Norte",
                                              DepartamentoPartidoComuna == "Sistema Agua Sur" ~ "Sistema Agua Sur",
                                              DepartamentoPartidoComuna == "Sistema Riachuelo" ~ "Sistema Riachuelo",
                                              DepartamentoPartidoComuna == "Sistema Berazategui" ~ "Sistema Berazategui",
                                              TRUE ~ Departamento.Normalizado),
         Localidad.Normalizado = case_when(LocalidadAsentamientoHumano == "Varios" | LocalidadAsentamientoHumano == "Varias" ~ "Varios",
                                           LocalidadAsentamientoHumano == "Sistema Agua Norte" ~ "Sistema Agua Norte",
                                           LocalidadAsentamientoHumano == "Sistema Agua Sur" ~ "Sistema Agua Sur",
                                           LocalidadAsentamientoHumano == "Sistema Riachuelo" ~ "Sistema Riachuelo",
                                           LocalidadAsentamientoHumano == "Sistema Berazategui" ~ "Sistema Berazategui",
                                           TRUE ~ Localidad.Normalizado))


# ----- Finalizando la Normalizacion de Unidades Territoriales ----- ####

# Dado que para aquellas unidades territoriales (provincias, departamentos o 
# localidades) que no encuentra la API le asigna un "NULL", vamos a transformar 
# esos "NULL" en NA para simplificar los pasos posteriores

dat <- dat %>% 
  mutate(Provincia.Normalizada = case_when(Provincia.Normalizada == "NULL" ~ NA_character_,
                                           TRUE ~ Provincia.Normalizada),
         Departamento.Normalizado = case_when(Departamento.Normalizado == "NULL" ~ NA_character_,
                                              TRUE ~ Departamento.Normalizado),
         Localidad.Normalizado = case_when(Localidad.Normalizado == "NULL" ~ NA_character_,
                                           TRUE ~ Localidad.Normalizado))

# Dataset de back-up por si piso el original donde estoy trabajando
back.1 <- dat 
# dat <- back.1

# --------------------------------------------------------------------------------
#                 Paso 2 – Segunda ronda normalización de nombres DEPARTAMENTOS
# --------------------------------------------------------------------------------

# Hay casos donde el nombre del Departamentos que fue normalizado en el Paso 1 pero
# existe una localidad. Por lo tanto, en este paso se intenta normalizar la mayor 
# cantidad de Departamentos a partir del binomio “Provincia” + “Localidad”. 
# Esta normalización es un INFERENCIA, pueden existir errores al momento de asignar
# el nombre Normalizado.

# De la normalización nos vamos a quedar con el nombre de la Localidad Normalizado,
# con el Departamento Normalizado y las coordenadas de la localidad.

# Hacemos un subset de aquellos departamentos que no tienen departamento pero si localidad y
# tratamos de normalizar de esta forma

# Este es el subset que se va a normalizar
sub.dat <- dat %>% 
  filter(Provincia.Normalizada != "INTERPROVINCIAL") %>% 
  filter(is.na(Departamento.Normalizado)) %>% 
  select(ID_Lauti, Provincia.Normalizada, DepartamentoPartidoComuna,
         LocalidadAsentamientoHumano, Departamento.Normalizado,
         Localidad.Normalizado)


# Hacemos la normalizacion
for (i in 1:nrow(sub.dat)) {
  print(i)
  
  prov <- sub.dat$Provincia.Normalizada[i]
  localidad <- sub.dat$LocalidadAsentamientoHumano[i] %>% str_to_lower()
  
  print(localidad)
  
  query.localidad <- get_localidades(nombre = localidad,
                                     provincia = prov)
  
  Sys.sleep(.1)
  if (length(query.localidad) != 0) {
    
    if (nrow(query.localidad) == 1) {
      
      l.norm <- query.localidad$nombre
      d.norm <- query.localidad$departamento_nombre
      
      
      id.localidad <- query.localidad$id
      l.lat <- query.localidad$centroide_lat
      l.lon <- query.localidad$centroide_lon
      
      sub.dat$Localidad.Normalizado[i] <- l.norm
      sub.dat$Departamento.Normalizado[i] <- d.norm
      sub.dat$Localidad.Lat.Centroide[i] <- l.lat
      sub.dat$Localidad.Lon.Centroide[i] <- l.lon
      sub.dat$ID.Localidad[i] <- id.localidad
      
      
    } 
    else {
      
      sub.dat$Localidad.Normalizado[i] <- NA
      sub.dat$Departamento.Normalizado[i] <- NA
      sub.dat$Localidad.Lat.Centroide[i] <- NA
      sub.dat$Localidad.Lon.Centroide[i] <- NA
      sub.dat$ID.Localidad[i] <- NA
    }
    
  } 
  else{
    
    sub.dat$Localidad.Normalizado[i] <- NA
    sub.dat$Departamento.Normalizado[i] <- NA
    sub.dat$Localidad.Lat.Centroide[i] <- NA
    sub.dat$Localidad.Lon.Centroide[i] <- NA
    sub.dat$ID.Localidad[i] <- NA
    
  }
  
}

# openxlsx::write.xlsx(sub.dat, "02.Datasets_soporte/DT.Departamentos inferidos.xlsx")
# sub.dat <- read_excel("02.Datasets_soporte/DT.Departamentos inferidos.xlsx")

sub.dat <- sub.dat %>% 
  mutate(Flag.Normalizacion = "Departamento Normalizado por inferencia") %>% 
  rename(Departamento.Normalizado.nuevo = Departamento.Normalizado,
         Localidad.Normalizado.nuevo = Localidad.Normalizado,
         Localidad.Lat.Centroide.nuevo = Localidad.Lat.Centroide,
         Localidad.Lon.Centroide.nuevo = Localidad.Lon.Centroide,
         ID.Localidad.nuevo = ID.Localidad)

# Simplificando el dataset anterior
sub.dat2 <- sub.dat %>% 
  filter(!is.na(Departamento.Normalizado.nuevo)) %>% 
  select(ID_Lauti, 
         Departamento.Normalizado.nuevo, 
         Localidad.Normalizado.nuevo,
         Localidad.Lat.Centroide.nuevo,
         Localidad.Lon.Centroide.nuevo,
         ID.Localidad.nuevo,
         Flag.Normalizacion)

# Haciendo el merge con el dataset grande
dat <- dat %>% left_join(sub.dat2, by = "ID_Lauti")


# Reemplazando los nuevos valores encontrados 
# Esto va en el paso 2 antes de Normalizar por departamentos
dat <- dat %>% 
  mutate(Departamento.Normalizado = case_when(is.na(Departamento.Normalizado) & !is.na(Departamento.Normalizado.nuevo) ~ Departamento.Normalizado.nuevo,
                                              TRUE ~ Departamento.Normalizado),
         Localidad.Normalizado = case_when(is.na(Localidad.Normalizado) & !is.na(Localidad.Normalizado.nuevo) ~ Localidad.Normalizado.nuevo,
                                           TRUE ~ Localidad.Normalizado),
         Localidad.Lat.Centroide = case_when(is.na(Localidad.Lat.Centroide) & !is.na(Localidad.Lat.Centroide.nuevo) ~ Localidad.Lat.Centroide.nuevo,
                                             TRUE ~ Localidad.Lat.Centroide),
         Localidad.Lon.Centroide = case_when(is.na(Localidad.Lon.Centroide) & !is.na(Localidad.Lon.Centroide.nuevo) ~ Localidad.Lon.Centroide.nuevo,
                                             TRUE ~ Localidad.Lon.Centroide),
         ID.Localidad = case_when(is.na(ID.Localidad) & !is.na(ID.Localidad.nuevo) ~ ID.Localidad.nuevo,
                                  TRUE ~ ID.Localidad)
  )

# Borrando filas que no nos interesan más
dat <- dat %>% 
  select(-c(Departamento.Normalizado.nuevo, 
            Localidad.Normalizado.nuevo,
            Localidad.Lat.Centroide.nuevo,
            Localidad.Lon.Centroide.nuevo,
            ID.Localidad.nuevo))

# Dataset de back-up por si piso el original donde estoy trabajando
back.2 <- dat
# dat <- back.2

# --------------------------------------------------------------------------------
#                 Paso 3 – Segunda ronda normalización de nombres LOCALIDADES
# --------------------------------------------------------------------------------

# Una de las desventajas que tiene normalizar por la API viene dada por el archivo 
# de comparación que utiliza: no todas las localidades incluidas en BARHA están 
# incluidas en el archivo base de la API. Esto genera que muchas localidades no 
# puedan ser normalizadas aunque estén bien escritas. Para reducir este margen de 
# error, lo que se hizo fue hacer un simple join con los datos de Unidades Territoriales.

ut.localidades <- read_excel("02.Datasets_soporte/Unidades Territoriales normalizadas.xlsx")

sub.dat.loc <- dat %>% 
  # Seleccionando únicamente las variables que vamos a usar en este subset
  select(ID_Lauti, Provincia.Normalizada, Departamento.Normalizado,
         LocalidadAsentamientoHumano, Localidad.Normalizado) %>% 
  # Filtrando las variables que no nos interesan
  filter(Provincia.Normalizada != "INTERPROVINCIAL",
         !is.na(Departamento.Normalizado), 
         !is.na(LocalidadAsentamientoHumano),
         is.na(Localidad.Normalizado)) %>% 
  mutate(LocalidadAsentamientoHumano = str_to_upper(LocalidadAsentamientoHumano)) %>% 
  # Haciendo el join
  left_join(ut.localidades, by = c("Provincia.Normalizada" = "Provincia.Normalizada",
                                   "Departamento.Normalizado" = "Departamento.Normalizado",
                                   "LocalidadAsentamientoHumano" = "Nombre")) %>% 
  mutate(Localidad.Normalizado.nuevo = case_when(!is.na(Código) ~ LocalidadAsentamientoHumano,
                                                 TRUE ~ Localidad.Normalizado),
         Latitud = Latitud %>% as.numeric(),
         Longitud = Longitud %>% as.numeric()) %>% 
  filter(!is.na(Localidad.Normalizado.nuevo)) %>%
  rename(Localidad.Lat.Centroide.nuevo = Latitud,
         Localidad.Lon.Centroide.nuevo = Longitud,
         ID.Localidad.nuevo = Código) %>% 
  select(ID_Lauti, 
         Localidad.Normalizado.nuevo,
         Localidad.Lat.Centroide.nuevo,
         Localidad.Lon.Centroide.nuevo,
         ID.Localidad.nuevo)


# Haciendo el merge con el dataset grande
dat <- dat %>% left_join(sub.dat.loc, by = "ID_Lauti")


# Reemplazando los nuevos valores encontrados 
# Esto va en el paso 2 antes de Normalizar por departamentos
dat <- dat %>% 
  mutate(Localidad.Normalizado = case_when(is.na(Localidad.Normalizado) & !is.na(Localidad.Normalizado.nuevo) ~ Localidad.Normalizado.nuevo,
                                           TRUE ~ Localidad.Normalizado),
         Localidad.Lat.Centroide = case_when(is.na(Localidad.Lat.Centroide) & !is.na(Localidad.Lat.Centroide.nuevo) ~ Localidad.Lat.Centroide.nuevo,
                                             TRUE ~ Localidad.Lat.Centroide),
         Localidad.Lon.Centroide = case_when(is.na(Localidad.Lon.Centroide) & !is.na(Localidad.Lon.Centroide.nuevo) ~ Localidad.Lon.Centroide.nuevo,
                                             TRUE ~ Localidad.Lon.Centroide),
         ID.Localidad = case_when(is.na(ID.Localidad) & !is.na(ID.Localidad.nuevo) ~ ID.Localidad.nuevo,
                                  TRUE ~ ID.Localidad)
  )

# Borrando filas que no nos interesan más
dat <- dat %>% 
  select(-c(Localidad.Normalizado.nuevo,
            Localidad.Lat.Centroide.nuevo,
            Localidad.Lon.Centroide.nuevo,
            ID.Localidad.nuevo))

# Dataset de back-up por si piso el original donde estoy trabajando
back.3 <- dat
# dat <- back.3

# --------------------------------------------------------------------------------
#                 Paso 4 – Recuperar Dptos y Localidades Originales
# --------------------------------------------------------------------------------

# Dado que hay casos en los cuales ni el departamento ni la localidad han logrado 
# ser normalizados (ya sea porque contienen valores múltiples o porque no se pudo
# encontrar coincidencia entre el valor original y el de referencia), se recuperan 
# los valores originales para no perder dicha información.


# Copiando las variables viejas y sucias para no perder datos
dat <- dat %>% 
  mutate(Flag.Normalizacion = case_when(is.na(Departamento.Normalizado) & !is.na(DepartamentoPartidoComuna) ~ "Departamento Original no Normalizado",
                                        is.na(Localidad.Normalizado) & !is.na(LocalidadAsentamientoHumano) ~ "Localidad Original no Normalizada",
                                        TRUE ~ Flag.Normalizacion),
         Departamento.Normalizado = case_when(is.na(Departamento.Normalizado) & !is.na(DepartamentoPartidoComuna) ~ DepartamentoPartidoComuna,
                                              TRUE ~ Departamento.Normalizado),
         Localidad.Normalizado = case_when(is.na(Localidad.Normalizado) & !is.na(LocalidadAsentamientoHumano) ~ LocalidadAsentamientoHumano,
                                           TRUE ~ Localidad.Normalizado)
  )


# --------------------------------------------------------------------------------
#                 Paso 5 – Control de Latitud y Longitud I
# --------------------------------------------------------------------------------

# La base de datos contiene las siguientes variables:
#   •	Latitud
#   •	Longitud
# En caso de que ambos valores existan y sean diferentes a cero:
#   1.	En caso de tener valores positivos, se los transforma a negativos y se deja constancia.
#   2.	Se verifica que el par Latitud/Longitud caiga dentro del "cuadrado" de Argentina. 
#           a.	De no ser el caso, se presupone que hubo un error en la carga del dato, por 
#               lo tanto, se invierten los valores y se deja constancia.
#           b.	De ser correcto, el proceso siguen sin anomalías.


# Bounding Box
long.maxima <- -53.359917
long.minima <- -74.267376
lat.maxima <- -21.570804
lat.minima <- -55.76222

# Acomodando las latitudes y Longitudes que estaban mal asignadas
dat <- dat %>% 
  # Si Lat/Long son positivas, les cambia el signo
  mutate(Latitud = case_when(Latitud > 0 ~ -Latitud, TRUE ~ Latitud),
         Longitud = case_when(Longitud > 0 ~ -Longitud, TRUE ~ Longitud)) %>% 
  mutate(N.Latitud = case_when(Latitud > lat.maxima | Latitud < lat.minima ~ Longitud, TRUE ~ Latitud),
         N.Longitud = case_when(Longitud > long.maxima | Longitud < long.minima ~ Latitud, TRUE ~ Longitud)) %>% 
  mutate(Flag = case_when(Latitud != N.Latitud | Longitud != N.Longitud ~ "Lat/Long Invertida",
                          Latitud == 0| Longitud == 0 ~ "Lat/Long igual a 0",
                          is.na(Longitud) | is.na(Latitud) ~ "Lat/Long es NA"))

back.4 <- dat
# dat <- back.4

# ----------------------------------------------------------------------------------
#                 Paso 3 – Control de Latitud y Longitud II
# ----------------------------------------------------------------------------------

# Para aquellas obras que tengan Latitud y Longitud existentes y diferentes a 0
# se realizan los siguientes controles: 
#     1) la Latitud y Longitud dentro del Polígono de la Provincia Normalizada
#     2) la Latitud y Longitud dentro del Polígono del Departamento Normalizado
#     3) Caso contrario, se indica el error obtenido
# 
# Por ejemplo, la Latitud y Longitud de una obra en Olavarría (Provincia de Buenos 
# Aires) debería indefectiblemente caer dentro del Poligono de la Provincia de Buenos 
# Aires y, luego, dentro del poligono del Departamento de Olavarria.

# Shapefiles
prov <- st_read("../99.Shapefiles/provincias_ign/provincias.shp")
departamentos.shp <- st_read("../99.Shapefiles/departamentos/Departamentos.shp")

departamentos.shp <- departamentos.shp %>% 
  mutate(IN1 = as.character(IN1))

# ----- Controlando Provincias ----- ####

dat.shp <- dat %>% 
  filter(Latitud != 0, 
         Longitud != 0, 
         !is.na(Longitud), 
         !is.na(Latitud)) %>% 
  select(ID_Lauti, Provincia.Normalizada, Departamento.Normalizado,
         N.Longitud,N.Latitud)

# Transformando el formato del archivo
coordinates(dat.shp) <- ~ N.Longitud + N.Latitud

dat.shp <- st_as_sf(dat.shp, crs = 4326, coords = c("Longitud", "Latitud"))
st_crs(dat.shp) = 4326

nombre.provincias <- dat.shp %>% 
  as.data.frame() %>% 
  select(Provincia.Normalizada) %>% 
  filter(Provincia.Normalizada != "NULL") %>% 
  filter(Provincia.Normalizada != "INTERPROVINCIAL") %>% 
  distinct()


dat.soporte.prov <- data.frame()

for (i in 1:nrow(nombre.provincias)) {
  
  p <- nombre.provincias$Provincia.Normalizada[i]
  print(p)
  prov.dat.shp <- dat.shp %>% filter(Provincia.Normalizada == p)
  prov.shp <- prov %>% filter(NAM == p)
  
  int.prov <- st_intersection(prov.shp, prov.dat.shp)
  int.prov$En_provincia <- 1
  
  prov.dat.shp.2 <- prov.dat.shp %>% 
    as.data.frame() %>% 
    left_join(int.prov %>% select(ID_Lauti, En_provincia) %>% as.data.frame(), 
              by= "ID_Lauti" ) %>% 
    filter(is.na(En_provincia))
  
  dat.soporte.prov <- rbind(dat.soporte.prov, prov.dat.shp.2)
  
}

# dat.soporte.prov.2 <- dat.soporte.prov %>%
#   select(ID_Lauti, En_provincia)
# 
# openxlsx::write.xlsx(dat.soporte.prov.2, "02.Datasets_soporte/DT. LatLong fuera de Provincias.xlsx")


# Haciendo el join con dat
dat <- dat %>% 
  left_join(dat.soporte.prov %>% select(ID_Lauti, En_provincia) %>% mutate(En_provincia = "NO"), 
            by = "ID_Lauti") %>% 
  mutate(Flag = case_when(En_provincia == "NO" ~ "Lat/Long fuera de Provincia",
                          TRUE ~ Flag))


# ----- Controlando Departamentos ----- ####

dat.shp.departamentos <- dat %>% 
  filter(Latitud != 0, 
         Longitud != 0, 
         !is.na(Longitud), 
         !is.na(Latitud)) %>% 
  filter(Flag == "Lat/Long Invertida" | is.na(Flag)) %>%
  select(ID_Lauti, Provincia.Normalizada, Departamento.Normalizado,
         N.Longitud,N.Latitud, ID.Localidad) %>% 
  mutate(ID.Departamento = str_extract(ID.Localidad, "[0-9]{5}"))


# Transformando el formato del archivo
coordinates(dat.shp.departamentos) <- ~ N.Longitud + N.Latitud

dat.shp.departamentos <- st_as_sf(dat.shp.departamentos,
                                  crs = 4326,
                                  coords = c("Longitud", "Latitud"))

st_crs(dat.shp.departamentos)  <-  4326

nombre.provincias <- dat.shp.departamentos %>% 
  as.data.frame() %>% 
  select(Provincia.Normalizada) %>% 
  filter(Provincia.Normalizada != "NULL") %>% 
  filter(Provincia.Normalizada != "INTERPROVINCIAL") %>% 
  distinct()


dat.soporte.departamentos <- data.frame()

for (i in 1:nrow(nombre.provincias)) {
  
  p <- nombre.provincias$Provincia.Normalizada[i]
  prov.dat.shp <- dat.shp.departamentos %>% filter(Provincia.Normalizada == p)
  
  nombre.departamento <- prov.dat.shp %>% 
    as.data.frame() %>% 
    select(Departamento.Normalizado, ID.Departamento) %>% 
    filter(!is.na(ID.Departamento)) %>% 
    filter(Departamento.Normalizado != "NULL") %>% 
    distinct()
  
  for (j in 1:nrow(nombre.departamento)){
    d <- nombre.departamento$Departamento.Normalizado[j]
    id.d <- nombre.departamento$ID.Departamento[j]
    print(d)
    
    dpto.dat.shp <- prov.dat.shp %>% filter(Departamento.Normalizado == d)
    dpto.shp <- departamentos.shp %>% filter(IN1 == id.d)
    
    int.dpto <- st_intersection(dpto.shp, dpto.dat.shp)
    
    if (nrow(int.dpto) >0){ 
      int.dpto$En_departamento <- 1
      
      dpto.dat.shp.2 <- dpto.dat.shp %>% 
        as.data.frame() %>% 
        left_join(int.dpto %>% select(ID_Lauti, En_departamento) %>% as.data.frame(), 
                  by= "ID_Lauti" ) %>% 
        filter(is.na(En_departamento))
      
      dat.soporte.departamentos <- rbind(dat.soporte.departamentos, dpto.dat.shp.2)
      
    }
    
  }
  
}

# dat.soporte.departamentos.2 <- dat.soporte.departamentos %>% 
#   select(ID_Lauti, En_departamento)
# openxlsx::write.xlsx(dat.soporte.departamentos.2, "DT. LatLong fuera de Departamento.xlsx")
# dat.soporte.departamentos <- read_excel("DT. LatLong fuera de Departamento.xlsx")



dat <- dat %>% 
  left_join(dat.soporte.departamentos %>% 
              select(ID_Lauti, En_departamento) %>% 
              mutate(En_departamento = "NO"), 
            by = "ID_Lauti") %>% 
  mutate(Flag = case_when(En_departamento == "NO" ~ "Lat/Long fuera de Departamento",
                          TRUE ~ Flag))

back.4 <- dat

# ----------------------------------------------------------------------------------
#                 Paso 4 – Asignar Latitud y Longitud a faltantes
# ----------------------------------------------------------------------------------

# Para aquellas obras que tengan Latitud y Longitud NA 0 iguales a 0 se les asigna 
# la siguiente Latitud y Longitud:
#   
# 1) Si existe Localidad Normalizada, se le asigna esa Latitud y Longitud y se deja 
#     asentado
# 2) Si NO existe Localidad Normalizada pero SI existe Departamento Normalizado, 
#     se le asigna esa Latitud y Longitud y se deja asentado
# 3) Si NO existe Localidad Normalizada ni Departamento Normalizado pero SI existe 
#     Provincia Normalizada, se le asigna esa Latitud y Longitud y se deja asentado


dat <- dat %>%  
  # Lat y Long NA
  mutate(N.Latitud = case_when(Flag == "Lat/Long es NA" & !is.na(Localidad.Normalizado) ~ Localidad.Lat.Centroide,
                               Flag == "Lat/Long es NA" & !is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Departamento.Lat.Centroide,
                               Flag == "Lat/Long es NA" & !is.na(Provincia.Normalizada) & is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Provincia.Lat.Centroide,
                               TRUE ~ N.Latitud),
         N.Longitud = case_when(Flag == "Lat/Long es NA" & !is.na(Localidad.Normalizado) ~ Localidad.Lon.Centroide,
                                Flag == "Lat/Long es NA" & !is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Departamento.Lon.Centroide,
                                Flag == "Lat/Long es NA" & !is.na(Provincia.Normalizada) & is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Provincia.Lon.Centroide,
                                TRUE ~ N.Longitud),
         Flag = case_when(Flag == "Lat/Long es NA" & N.Latitud == Localidad.Lat.Centroide ~ "Lat/Long NA - Asignado a Localidad",
                          Flag == "Lat/Long es NA" & N.Latitud == Departamento.Lat.Centroide ~ "Lat/Long NA - Asignado a Departamento",
                          Flag == "Lat/Long es NA" & N.Latitud == Provincia.Lat.Centroide ~ "Lat/Long NA - Asignado a Provincia",
                          TRUE ~ Flag)) %>%
  # Lat y Long 0
  mutate(N.Latitud = case_when(Flag == "Lat/Long igual a 0" & !is.na(Localidad.Normalizado) ~ Localidad.Lat.Centroide,
                               Flag == "Lat/Long igual a 0" & !is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Departamento.Lat.Centroide,
                               Flag == "Lat/Long igual a 0" & !is.na(Provincia.Normalizada) & is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Provincia.Lat.Centroide,
                               TRUE ~ N.Latitud),
         N.Longitud = case_when(Flag == "Lat/Long igual a 0" & !is.na(Localidad.Normalizado) ~ Localidad.Lon.Centroide,
                                Flag == "Lat/Long igual a 0" & !is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Departamento.Lon.Centroide,
                                Flag == "Lat/Long igual a 0" & !is.na(Provincia.Normalizada) & is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Provincia.Lon.Centroide,
                                TRUE ~ N.Longitud),
         Flag = case_when(Flag == "Lat/Long igual a 0" & N.Latitud == Localidad.Lat.Centroide ~ "Lat/Long 0 - Asignado a Localidad",
                          Flag == "Lat/Long igual a 0" & N.Latitud == Departamento.Lat.Centroide ~ "Lat/Long 0 - Asignado a Departamento",
                          Flag == "Lat/Long igual a 0" & N.Latitud == Provincia.Lat.Centroide ~ "Lat/Long 0 - Asignado a Provincia",
                          TRUE ~ Flag)) %>%
  # Lat y Long fuera de Provincia 
  mutate(N.Latitud = case_when(Flag == "Lat/Long fuera de Provincia" & !is.na(Localidad.Normalizado) ~ Localidad.Lat.Centroide,
                               Flag == "Lat/Long fuera de Provincia" & !is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Departamento.Lat.Centroide,
                               Flag == "Lat/Long fuera de Provincia" & !is.na(Provincia.Normalizada) & is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Provincia.Lat.Centroide,
                               TRUE ~ N.Latitud),
         N.Longitud = case_when(Flag == "Lat/Long fuera de Provincia" & !is.na(Localidad.Normalizado) ~ Localidad.Lon.Centroide,
                                Flag == "Lat/Long fuera de Provincia" & !is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Departamento.Lon.Centroide,
                                Flag == "Lat/Long fuera de Provincia" & !is.na(Provincia.Normalizada) & is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Provincia.Lon.Centroide,
                                TRUE ~ N.Longitud),
         Flag = case_when(Flag == "Lat/Long fuera de Provincia" & N.Latitud == Localidad.Lat.Centroide ~ "Lat/Long fuera Provincia - Asignado a Localidad",
                          Flag == "Lat/Long fuera de Provincia" & N.Latitud == Departamento.Lat.Centroide ~ "Lat/Long fuera Provincia - Asignado a Departamento",
                          Flag == "Lat/Long fuera de Provincia" & N.Latitud == Provincia.Lat.Centroide ~ "Lat/Long fuera Provincia - Asignado a Provincia",
                          TRUE ~ Flag)) %>% 
  # Lat y Long fuera de Departamento 
  mutate(N.Latitud = case_when(Flag == "Lat/Long fuera de Departamento" & !is.na(Localidad.Normalizado) ~ Localidad.Lat.Centroide,
                               Flag == "Lat/Long fuera de Departamento" & !is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Departamento.Lat.Centroide,
                               Flag == "Lat/Long fuera de Departamento" & !is.na(Provincia.Normalizada) & is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Provincia.Lat.Centroide,
                               TRUE ~ N.Latitud),
         N.Longitud = case_when(Flag == "Lat/Long fuera de Departamento" & !is.na(Localidad.Normalizado) ~ Localidad.Lon.Centroide,
                                Flag == "Lat/Long fuera de Departamento" & !is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Departamento.Lon.Centroide,
                                Flag == "Lat/Long fuera de Departamento" & !is.na(Provincia.Normalizada) & is.na(Departamento.Normalizado) & is.na(Localidad.Normalizado) ~ Provincia.Lon.Centroide,
                                TRUE ~ N.Longitud),
         Flag = case_when(Flag == "Lat/Long fuera de Departamento" & N.Latitud == Localidad.Lat.Centroide ~ "Lat/Long fuera Departamento - Asignado a Localidad",
                          Flag == "Lat/Long fuera de Departamento" & N.Latitud == Departamento.Lat.Centroide ~ "Lat/Long fuera Departamento - Asignado a Departamento",
                          Flag == "Lat/Long fuera de Departamento" & N.Latitud == Provincia.Lat.Centroide ~ "Lat/Long fuera Departamento - Asignado a Provincia",
                          TRUE ~ Flag))



back.5 <- dat
# dat <- back.5
# ----------------------------------------------------------------------------------
#                 Paso 7 – Corrección manual de departamentos diferencias INDEC
# ----------------------------------------------------------------------------------

san.martin <- c("San Martín", 
                "General San Martin", 
                "General San Martín", 
                "San Martin", 
                "GENERAL SAN MARTIN")


dat <- dat %>% 
  mutate(Departamento.Normalizado = case_when(Departamento.Normalizado == "José M. Ezeiza" ~ "Ezeiza",
                                              Departamento.Normalizado == "General Ángel V. Peñaloza" ~ "Ángel Vicente Peñaloza",
                                              Departamento.Normalizado == "General Juan F. Quiroga" ~ "General Juan Facundo Quiroga",
                                              Provincia.Normalizada == "Buenos Aires" & DepartamentoPartidoComuna %in% san.martin ~ "General San Martín",
                                              Provincia.Normalizada == "Santa Fe" & Departamento.Normalizado == "Villa Constitución" ~ "Constitución",
                                              TRUE ~ Departamento.Normalizado))
back.6 <- dat
# dat <- back.6

# ----------------------------------------------------------------------------------
#                 Paso 7 – Agregar columnas especiales
# ----------------------------------------------------------------------------------

# Abriendo el dataset de Conurbano
pba.conurbano <- read_excel("../01.Mapas/02.Soporte_Datasets/pba.conurbano.xlsx")

pba.conurbano <- pba.conurbano %>%
  select(-Zona) %>% 
  rename(Conurbano.Nuevo = Conurbano,
         Provincia.Normalizada = Provincia,
         Departamento.Normalizado = Partido)

sistemas <- c("Sistema Agua Norte", "Sistema Agua Sur", 
              "Sistema Berazategui", "Sistema Riachuelo",
              "Ezeiza")


# Estas son columnas especiales para Castillo

dat <- dat %>% 
  left_join(pba.conurbano, by = c("Provincia.Normalizada", "Departamento.Normalizado")) %>% 
  mutate(Conurbano.Nuevo = case_when(Conurbano.Nuevo == "Conurbano" ~ Conurbano.Nuevo,
                                     Provincia.Normalizada == "Buenos Aires" & DepartamentoPartidoComuna %in% san.martin ~ "Conurbano",
                                     Departamento.Normalizado %in% sistemas ~ "Conurbano"),
         "Coordenadas originales" = "Localización aproximada",
         Latitud_mapa = case_when(!is.na(Latitud) & !is.na(N.Latitud) & Latitud == N.Latitud ~ "Coordenada original",
                                  !is.na(Latitud) & !is.na(N.Latitud) & Latitud != N.Latitud ~ "Coordenada asignada por procesamiento",
                                  TRUE ~ NA_character_),
         "Coordenadas mapa" = case_when(is.na(Flag) ~ "Posición obra",
                                        Flag == "Lat/Long 0 - Asignado a Localidad" |
                                          Flag == "Lat/Long NA - Asignado a Localidad" |
                                          Flag == "Lat/Long fuera Departamento - Asignado a Localidad" | 
                                          Flag == "Lat/Long fuera Provincia - Asignado a Localidad" ~ "Posición aprox en localidad",
                                        Flag == "Lat/Long 0 - Asignado a Departamento" |
                                          Flag == "Lat/Long NA - Asignado a Departamento" |
                                          Flag == "Lat/Long fuera Departamento - Asignado a Departamento" |
                                          Flag == "Lat/Long fuera Provincia - Asignado a Departamento" ~ "Posición aprox en departamento/partido",
                                        Flag == "Lat/Long 0 - Asignado a Provincia" |
                                          Flag == "Lat/Long fuera Provincia - Asignado a Provincia" |
                                          Flag == "Lat/Long NA - Asignado a Provincia" ~ "Posición aprox en provincia")
  )

# ----------------------------------------------------------------------------------
#                 Paso 8 – Finalizando y cerrando todo
# ----------------------------------------------------------------------------------


dat.final <- dat %>% 
  select(-c(ID_Lauti, 
            Provincia.Lat.Centroide,
            Provincia.Lon.Centroide,
            Departamento.Lat.Centroide,
            Departamento.Lon.Centroide,
            Localidad.Lat.Centroide,
            Localidad.Lon.Centroide,
            En_provincia,
            En_departamento))


openxlsx::write.xlsx(dat.final, "03.Dataset_finales/Base MIOPyV 07.2019 - Normalizada Lauti.xlsx")

openxlsx::write.xlsx(dat.final, "N:/JG_Planificacion/6. BASE MENSUAL DE OBRAS/0.4. RUTINA PLAN B/03- 2019/07. 31.07/BASE MIOPyV 31.07/BaseMIOPyV-ACTIVA+TERMINADA+OBRASSIU2_31.07.2019 - FINAL - Normaliazcion Lauti.xlsx")

t1 <- Sys.time()

t1-t0

