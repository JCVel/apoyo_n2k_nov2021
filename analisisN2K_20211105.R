################################################################################
# Archivo: analisisCNTRYES.R
# Autor: Juan Carlos VelAzquez Melero 
# Fecha: xx 2021
# VersiOn: 2021xxxx
# DescripciOn: Script elaborado para 

# Datos de entrada: 
# Datos de salida: 

################################################################################
###___ InstalaciOn de paquetes y establecimiento del directorio de trabajo
################################################################################
rm(list = ls(all = TRUE)) # Eliminar objetos previos, si los hubiera
# options(scipen=999)
# InstalaciOn de paquetes tomada de Antoine Soetewey (https://statsandr.com)
# packages <- c("sf", "readxl", "xlsx", "openxlsx", "dplyr", "RODBC")
packages <- c("sf", "RODBC", "stringr")
# IntalaciOn de paquetes
installed_packages <- packages %in% rownames(installed.packages())
if (any(installed_packages == FALSE)) {
  install.packages(packages[!installed_packages])
}
# Carga de paquetes
invisible(lapply(packages, library, character.only = TRUE))
rm(installed_packages, packages)

# Introducir aquI el directorio de trabajo (donde estAn los ficheros de entrada)
# setwd("D:/JCV_interno/19_CNTRYES_apoyoNov2021/0_datosEntrada")
# setwd("P:/Groups/Ger_Calidad_Eva_Ambiental_Bio/DpMNatural_TEC/2021_CARTOGRAFIA_HABITAT/2_R/0_analisisDatosEUNIS_LPEHT_DH/analisisClasificHab_ficheros")

################################################################################
###___  Lectura de datos
################################################################################
#########################
### Lectura de BD RNATURA_AnalisisProximidad.mdb
#########################

# (configurar R a 32bit). Para mAs detalles ver https://www.roelpeters.be/solved-importing-microsoft-access-files-accdb-mdb-in-r/
path <- "D:/JCV_interno/19_n2k_apoyoNov2021/0_datosEntrada/RNATURA_AnalisisProximidad/RNATURA_AnalisisProximidad.mdb"
con <- odbcConnectAccess(path) # Establecer conexiOn
RNProxi<- sqlFetch(con, "RNATURA_proximidad") # Leer tabla de interEs
EP <- sqlFetch(con, "ESPACIOS_PROTEGIDOS") 
odbcClose(con) # cerrar conexiOn con el mdb
rm(con)

#########################
### Lectura datos CNTRYES
#########################
CNTRYES_pibal <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "n2k_pibal_202101")
CNTRYES_pibal <- st_set_geometry(CNTRYES_pibal, NULL)
CNTRYES_pibal <- within(CNTRYES_pibal, rm(hectareas))
CNTRYES_can <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "n2k_can_202101")
CNTRYES_can <- st_set_geometry(CNTRYES_can, NULL)
CNTRYES_can <- within(CNTRYES_can, rm(HECTAREAS))

CNTRYES <- rbind(CNTRYES_pibal, CNTRYES_can) # UniOn pibal y can
rm(CNTRYES_pibal, CNTRYES_can)

################################################################################
### Depurado de datos
################################################################################
#########################
### ANadir campos desconcatenados del SITE_CODE
#########################
# FunciOn que desconcatena un id formado por dos partes separadas por "_"
desconcatenaID <- function(df, field){
  for (row in 1:nrow(df)) {
    pre <- unlist(strsplit(df[[field]][[row]], split = '_', fixed = TRUE))[1]
    field_pre <- paste0(field, "_pre")
    df[[field_pre]][[row]] <- pre

    suf <- unlist(strsplit(df[[field]][[row]], split = '_', fixed = TRUE))[2]
    field_suf <- paste0(field, "_suf")
    df[[field_suf]][[row]] <- suf
  }
  df[[field_pre]] <- unlist(df[[field_pre]])
  df[[field_suf]] <- unlist(df[[field_suf]])
  return(df)
  }

# Desconcatenar los ID de las dos tablas de la BD RNATURA_AnalisisProximidad
RNProxi <- desconcatenaID(RNProxi, "src_SITE_CODE")
EP <- desconcatenaID(EP, "SITE_CODE")

#########################
### Comprobar nUmero de registros RN2000 en BD sea igual que en CNTRYES
#########################
# Sacar solo los registros de RNATURA 2000
EPRN <- EP[EP$SITE_CODE_suf == "LIC" | EP$SITE_CODE_suf == "ZEC" |
               EP$SITE_CODE_suf == "ZEPA", ]

# Seleccionar los campos de interes para eliminar duplicados
EPRN <- data.frame(
  SITE_CODE_pre = EPRN$SITE_CODE_pre)
EPRNUniq <- unique(EPRN) # Salen 1857 sitios RN2000. EstA bien. Coincide con
# los datos de RN2000 del CNTRYES
#########################
### Añadir as API que no estAn incluidas en la BD
#########################
APIext_pibal <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/0_datosEntrada/APIs/API_ext.gdb", layer = "api_pibal")
APIext_pibal <- st_set_geometry(APIext_pibal, NULL)
APIext_can <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/0_datosEntrada/APIs/API_ext.gdb", layer = "api_can")
APIext_can <- st_set_geometry(APIext_can, NULL)

APIext <- rbind(APIext_pibal, APIext_can) # UniOn pibal y can
rm(APIext_pibal, APIext_can)

# Desconcatenar el SITE_CODE
APIext <- desconcatenaID(APIext, "SITE_CODE")
# Dejar los mismos campos que en la tabla EP de la BD
APIext <- subset(APIext, select = -c(Shape_Length, Shape_Area))
EPmod <- rbind(EP, APIext)
#########################
### En la tabla EPmod, eliminar duplicados espacios RN2000
#########################
# Primero indicar el tipo de espacio (A, B, c) a partir de los datos del CNTRYES
EPmod <- merge(EPmod, CNTRYES[ , c("SITE_CODE", "AC", "TIPO", "Shape_Area")], 
               by.x = "SITE_CODE_pre", by.y = "SITE_CODE", all.x = T)
# Comprobar que todos los tipos C estAn duplicados (y ya de paso comprobar si 
# hubiera otros duplicados) 
freq <- table(EPmod$SITE_CODE_pre)
freq <- as.data.frame(freq)
EPmod <- merge(EPmod, freq, by.x = "SITE_CODE_pre", by.y = "Var1", all.x = T)
# ConclusiOn: todos los tipos C estan duplic. y no hay mAs duplic

# Todo lo que sea LIC, ZEC, ZEPA en tabla EP pasa a ser RN2000
EPmod$FIGURA <- ifelse(EPmod$FIGURA == "LIC" | EPmod$FIGURA == "ZEC" |
                         EPmod$FIGURA == "ZEPA", "RN2000", EPmod$FIGURA )
EPmod$SITE_CODE_suf <- EPmod$FIGURA
# Copia para quitar duplicados (que solo son tipo C de RN2000)
EPmod2 <- EPmod 
EPmod2 <- subset(EPmod, select = -c(SITE_CODE, Nombre_))
EPmod2uniq <- unique(EPmod2)
EPmod2uniq$SITE_CODE_suf <- ifelse(EPmod2uniq$SITE_CODE_suf == "ENP", "CDDA", EPmod2uniq$SITE_CODE_suf)
# Unimos los nombres
EPmod2uniq <- merge(EPmod2uniq, EPmod[ , c("SITE_CODE_pre", "Nombre_")], 
                    by = "SITE_CODE_pre", all.x = T)
# Anadir un nombre que falta (ya faltaba en la tabla EP de la BD original)
EPmod2uniq$SITE_CODE_pre
EPmod2uniq[EPmod2uniq$SITE_CODE_pre == "2453", "Nombre_"] <- "Parque Nacional Marítimo-Terrestre de las Islas Atlánticas"

#########################
### En la tabla EPmod, anadir tipo de EP: ENP, RN2000, API
#########################
EPmod2uniq$jerar1 <- ifelse(EPmod2uniq$SITE_CODE_suf != "ENP" & EPmod2uniq$SITE_CODE_suf != "RN2000",
                            "API", EPmod2uniq$SITE_CODE_suf )

################################################################################
###___  En la tabla RNProxi, eliminar los duplicados de los tipos C de RN2000
################################################################################
# Meter los tipos de espacios (A, B, C) del CNTRYES
RNProxiMod <- merge(RNProxi, CNTRYES[ , c("SITE_CODE", "TIPO")], 
                    by.x = "src_SITE_CODE_pre", by.y = "SITE_CODE", all.x = T)
# Ver si hay duplicados en las combinaciones "src_SITE_CODE" y "nbr_SITE_CODE"
combinac <- RNProxiMod[ , c("src_SITE_CODE", "nbr_SITE_CODE")]
combinacUniq <- unique(combinac) # ConclusiOn: No hay duplicados en las combinaciones

cambiaARN2000 <- function(df, field){
  field2 <- paste0(field, "2")
  df[[field2]] <- ""
  for (row in 1:nrow(df)){
    if (any(grep("LIC", df[[field]][[row]])) | 
        any(grep("ZEC", df[[field]][[row]])) |
        any(grep("ZEPA", df[[field]][[row]]))){
      
      pre <- unlist(strsplit(df[[field]][[row]], "_"))[[1]]
      print(paste0(pre, "_RN2000"))
      
      df[[field2]][[row]] <- paste0(pre, "_RN2000")
    } else {
    df[[field2]][[row]] <- df[[field]][[row]] 
    }
  }
  return(df)
}

# AQUII en proceso de eliminar registros duplicados de la tabla RNProxi que
# esta modificada en combinacUniq2 Revisar esta parte. Cuando este limpio, volcar los 
# datos de la tabla RNProxi original

combinacUniq2 <- cambiaARN2000(combinacUniq, "src_SITE_CODE")
combinacUniq2 <- cambiaARN2000(combinacUniq2, "nbr_SITE_CODE")
combinacUniq2 <- combinacUniq2[ , c("src_SITE_CODE2", "nbr_SITE_CODE2")]
combinacUniq2 <- unique(combinacUniq2)
combinacUniq2$igual <- ifelse(combinacUniq2$src_SITE_CODE2 == combinacUniq2$nbr_SITE_CODE2,
                              "Igual", "Diferente")
table(RNProxiMod$TIPO)
table(combinacUniq2$igual)

combinacUniq2 <- desconcatenaID(combinacUniq2, "src_SITE_CODE2" )
combinacUniq2 <- merge(combinacUniq2, CNTRYES[ , c("SITE_CODE", "TIPO")], 
                    by.x = "src_SITE_CODE2_pre", by.y = "SITE_CODE", all.x = T)

combinacUniq2 <- combinacUniq2[!combinacUniq2$igual == "Igual", ]
################################################################################
###___  ExportaciOn de datos 
################################################################################

# write.table(CNTRYES, "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/CNTRYES.csv",
#             sep = ";" )
# write.table(EPRNUniq, "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/EPRNUniq.csv",
#             sep = ";" )
