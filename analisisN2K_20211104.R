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
options(scipen=999)
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
### Volcar datos del CNTRYES a las tablas de la BD
#########################


################################################################################
###___  ExportaciOn de datos
################################################################################

# write.table(CNTRYES, "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/CNTRYES.csv",
#             sep = ";" )
# write.table(EPRNUniq, "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/EPRNUniq.csv",
#             sep = ";" )
