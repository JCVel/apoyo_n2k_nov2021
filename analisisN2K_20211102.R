################################################################################
# Archivo: analisisN2K.R
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
# setwd("D:/JCV_interno/19_N2K_apoyoNov2021/0_datosEntrada")
# setwd("P:/Groups/Ger_Calidad_Eva_Ambiental_Bio/DpMNatural_TEC/2021_CARTOGRAFIA_HABITAT/2_R/0_analisisDatosEUNIS_LPEHT_DH/analisisClasificHab_ficheros")

################################################################################
###___ Procesado de datos iniciales (n2k y enp)
################################################################################
###____ Lectura tabla BBDD 
# (configurar R a 32bit). Para mAs detalles ver https://www.roelpeters.be/solved-importing-microsoft-access-files-accdb-mdb-in-r/
path <- "D:/JCV_interno/19_N2K_apoyoNov2021/0_datosEntrada/RNATURA_AnalisisProximidad/RNATURA_AnalisisProximidad.mdb"
con <- odbcConnectAccess(path) # Establecer conexiOn
rnatProxi<- sqlFetch(con, "RNATURA_proximidad") # Leer tabla de interEs
odbcClose(con) # cerrar conexiOn con el mdb
rm(con)


for (row in 1:nrow(rnatProxi)) {
  pre <- unlist(strsplit(rnatProxi$src_SITE_CODE[[row]], split = '_', fixed = TRUE))[1]
  rnatProxi$scr_SITE_CODE_pre[[row]] <- pre
  suf <- unlist(strsplit(rnatProxi$src_SITE_CODE[[row]], split = '_', fixed = TRUE))[2]
  rnatProxi$scr_SITE_CODE_suf[[row]] <- suf
}

rnatProxi$scr_SITE_CODE_pre <- unlist(rnatProxi$scr_SITE_CODE_pre)
rnatProxi$scr_SITE_CODE_suf <- unlist(rnatProxi$scr_SITE_CODE_suf)

rnatCuenta <- rnatProxi[ , c("scr_SITE_CODE_pre", "scr_SITE_CODE_suf")]
rnatCuenta <- unique(rnatCuenta)

rnatCuenta <- unique(rnatProxi$scr_SITE_CODE_pre)


n2k_pibal <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "n2k_pibal_202101")
n2k_pibal <- st_set_geometry(n2k_pibal, NULL)
n2k_pibal <- within(n2k_pibal, rm(hectareas))
n2k_can <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "n2k_can_202101")
n2k_can <- st_set_geometry(n2k_can, NULL)
n2k_can <- within(n2k_can, rm(HECTAREAS))

n2k <- rbind(n2k_pibal, n2k_can) # UniOn pibal y can
rm(n2k_pibal, n2k_can)


for (i in 1:nrow(n2k)){
  n2k$Presencia[i] <- ifelse(n2k$SITE_CODE[i] %in% rnatCuenta, "Presente", "Ausente")
}
n2k$Presencia <- unlist(n2k$Presencia)

# 
# cuentas$n2kCuenta[[2]] %in% rnatCuenta
# cuentas$n2kCuenta[[2]]
# "ES6320001" %in% rnatCuenta
# n2kC <- n2k[!n2k$TIPO == "C", ]
# 
# n2k$SITE_CODE[[4]]
# write.table(rnatCuenta, "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/rnatCuenta.csv", sep = ";" )
