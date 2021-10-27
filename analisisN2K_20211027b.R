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
packages <- c("sf", "readxl", "xlsx", "openxlsx", "dplyr", "RODBC")
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
############################################
###___ n2k
############################################
n2k_pibal <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "n2k_pibal_202101")
n2k_pibal <- st_set_geometry(n2k_pibal, NULL)
n2k_pibal <- within(n2k_pibal, rm(hectareas))
n2k_can <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "n2k_can_202101")
n2k_can <- st_set_geometry(n2k_can, NULL)
n2k_can <- within(n2k_can, rm(HECTAREAS))

n2k <- rbind(n2k_pibal, n2k_can) # UniOn pibal y can
# n2k_areas <- n2k[ , c("SITE_CODE", "Shape_Area")]
# names(n2k_areas) <- c("n2k_SITE_CODE", "m2")
names(n2k)[names(n2k) == "Shape_Area"] <- "n2k_m2"

############################################
###___ enp
############################################
enp_pibal <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "enp_pibal_202110")
enp_pibal <- st_set_geometry(enp_pibal, NULL)
enp_can <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "enp_can_202110")
enp_can <- st_set_geometry(enp_can, NULL)

enp <- rbind(enp_pibal, enp_can) # UniOn pibal y can

################################################################################
###___ Procesado cruce n2k_enp
################################################################################
##########################
###___ n2k_enp depuraciOn
##########################
n2k_enp_pibal <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "n2k_enp_pibal")
n2k_enp_pibal <- st_set_geometry(n2k_enp_pibal, NULL)
n2k_enp_pibal <- within(n2k_enp_pibal, rm(hectareas)) # quitar campo hectareas
n2k_enp_can <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "n2k_enp_can")
n2k_enp_can <- st_set_geometry(n2k_enp_can, NULL)
n2k_enp_can <- within(n2k_enp_can, rm(HECTAREAS)) # quitar campo hectareas

# Unificar nombres de campos para la fusiOn:
names(n2k_enp_pibal) <- gsub("_pibal_", "_", names(n2k_enp_pibal))
names(n2k_enp_can) <- gsub("_can_", "_", names(n2k_enp_can))
n2k_enp <- rbind(n2k_enp_pibal, n2k_enp_can)
rm(n2k_pibal, n2k_can, enp_pibal, enp_can, n2k_enp_pibal, n2k_enp_can)
n2k_enp <- within(n2k_enp, rm())

# Campo que identifica cada combinaciOn de espacio n2k con espacio enp
n2k_enp$n2k_enp_id <- paste0(n2k_enp$SITE_CODE, "_", n2k_enp$SITE_CODE_)
# Campo que identifica cada combinaciOn de espacio n2k con tipo de espacio enp(enp_DESIG_ABBR)
n2k_enp$n2k_enpDESIG_id <- paste0(n2k_enp$SITE_CODE, "_", n2k_enp$DESIG_ABBR)
# Campo que identifica cada combinaciOn de espacio n2k con tipo de espacio enp
# segUn ley 43/2007 (tiposENP_csv_ENP_GRAL_code)
n2k_enp$n2k_enpGRAL_id <- paste0(n2k_enp$SITE_CODE, "_", n2k_enp$tiposENP_csv_ENP_GRAL_code)

# Cambio de nombres de campos para saber el origen
n2k_enp <- n2k_enp[ , c("FID_n2k_202101", "SITE_CODE", "SITE_NAME", "AC", "TIPO",
                    "FID_enp_202110", "SITE_CODE_", "SITE_CDDA", "SITE_NAME_1",
                    "ODESIGNATE", "NUT2", "DESIG_ABBR", "tiposENP_csv_ENP_GRAL_code",
                    "tiposENP_csv_ENP_GRAL_name","n2k_enp_id", "n2k_enpDESIG_id",
                    "n2k_enpGRAL_id", "Shape_Area")]

names(n2k_enp) <- c("n2k_FID", "n2k_SITE_CODE", "n2k_SITE_NAME", "n2k_AC",
                    "n2k_TIPO", "enp_FID", "enp_SITE_CODE", "enp_SITE_CDDA", 
                    "enp_SITE_NAME", "enp_ODESIGNATE", "enp_NUT2", 
                    "enp_DESIG_ABBR",  "enp_GRAL_code", "enp_GRAL_name", 
                    "n2k_enp_id", "n2k_enpDESIG_id", "n2k_enpGRAL_id", "Shape_Area")

n2k_enp <- n2k_enp[order(n2k_enp$n2k_enp_id), ]

##########################
###___ n2k_enp nivel 1
##########################
# tabla n2k_enp_1 quita duplicados del primer cruce. Esta tabla tiene los datos
# mAs desagregados (n2k con enp) pero sin duplicados
n2k_enp_1 <- n2k_enp[ , c("n2k_enp_id", "n2k_SITE_CODE", "enp_SITE_CODE")] %>% 
  unique()
# Sacar el Area por cada combinacion n2k_enp
n2k_enp_1sum <- n2k_enp %>%                                     
  group_by(n2k_enp_id) %>%
  dplyr::summarise(n2k_enp_m2 = sum(Shape_Area)) %>% 
  as.data.frame() 
# Unir el Area anterior y el Area de cada espacio n2k
n2k_enp_1 <- merge(n2k_enp_1, n2k_enp_1sum, by = "n2k_enp_id", all.x = T)
rm(n2k_enp_1sum)
n2k_enp_1 <- merge(n2k_enp_1, n2k[ , c("SITE_CODE", "n2k_m2")], 
                   by.x = "n2k_SITE_CODE", by.y = "SITE_CODE", all.x = T)
# Sacar el porcentaje de cada espacio enp en cada sitio n2k
n2k_enp_1$porcen <- n2k_enp_1$n2k_enp_m2 / n2k_enp_1$n2k_m2 * 100
n2k_enp_1 <- n2k_enp_1[ , c("n2k_enp_id", "n2k_SITE_CODE", "enp_SITE_CODE",
                            "n2k_enp_m2",    "n2k_m2",        "porcen" )]

##########################
###___ n2k_enp nivel 2
##########################
# El nivel 2 es el que calsifica los enp por tipo de espacio ("enp_DESIG_ABBR")
n2k_enp_2 <- n2k_enp[ , c("n2k_enpDESIG_id", "n2k_SITE_CODE", "enp_DESIG_ABBR")] %>% 
  unique()
# Sacar el Area por cada combinacion n2k_enp
n2k_enp_2sum <- n2k_enp %>%                                     
  group_by(n2k_enpDESIG_id) %>%
  dplyr::summarise(n2k_enpDESIG_m2 = sum(Shape_Area)) %>% 
  as.data.frame() 
# Unir el Area anterior y el Area de cada espacio n2k
n2k_enp_2 <- merge(n2k_enp_2, n2k_enp_2sum, by = "n2k_enpDESIG_id", all.x = T)
rm(n2k_enp_2sum)
n2k_enp_2 <- merge(n2k_enp_2, n2k[ , c("SITE_CODE", "n2k_m2")], 
                   by.x = "n2k_SITE_CODE", by.y = "SITE_CODE", all.x = T)
# Sacar el porcentaje de cada espacio enp en cada sitio n2k
n2k_enp_2$porcen <- n2k_enp_2$n2k_enpDESIG_m2 / n2k_enp_2$n2k_m2 * 100
n2k_enp_2 <- n2k_enp_2[ , c("n2k_enpDESIG_id", "n2k_SITE_CODE", "enp_DESIG_ABBR",
                            "n2k_enpDESIG_m2",    "n2k_m2",        "porcen" )]

##########################
###___ n2k_enp nivel 3
##########################
# El nivel 3 es el que agrupa los tipos de enp segun ley 43/2007 ("enp_GRAL_code")
n2k_enp_3 <- n2k_enp[ , c("n2k_enpGRAL_id", "n2k_SITE_CODE", "enp_GRAL_code")] %>% 
  unique()
# Sacar el Area por cada combinacion n2k_enp
n2k_enp_3sum <- n2k_enp %>%                                     
  group_by(n2k_enpGRAL_id) %>%
  dplyr::summarise(n2k_enpGRAL_m2 = sum(Shape_Area)) %>% 
  as.data.frame() 
# Unir el Area anterior y el Area de cada espacio n2k
n2k_enp_3 <- merge(n2k_enp_3, n2k_enp_3sum, by = "n2k_enpGRAL_id", all.x = T)
rm(n2k_enp_3sum)
n2k_enp_3 <- merge(n2k_enp_3, n2k[ , c("SITE_CODE", "n2k_m2")], 
                   by.x = "n2k_SITE_CODE", by.y = "SITE_CODE", all.x = T)
# Sacar el porcentaje de cada espacio enp en cada sitio n2k
n2k_enp_3$porcen <- n2k_enp_3$n2k_enpGRAL_m2 / n2k_enp_3$n2k_m2 * 100
n2k_enp_3 <- n2k_enp_3[ , c("n2k_enpGRAL_id", "n2k_SITE_CODE", "enp_GRAL_code",
                            "n2k_enpGRAL_m2",    "n2k_m2",        "porcen" )]


################################################################################
###___ Procesado cruce n2kA_n2kB
################################################################################
n2k_AB_pibal <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "n2k_AB_pibal")
n2k_AB_pibal <- st_set_geometry(n2k_AB_pibal, NULL)
n2k_AB_pibal <- within(n2k_AB_pibal, rm(hectareas, hectareas_1))
names(n2k_AB_pibal) <- gsub("_pibal_", "_", names(n2k_AB_pibal))

n2k_AB_can <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/analisisN2K.gdb", layer = "n2k_AB_can")
n2k_AB_can <- st_set_geometry(n2k_AB_can, NULL)
n2k_AB_can <- within(n2k_AB_can, rm(HECTAREAS, HECTAREAS_1))
names(n2k_AB_can) <- gsub("_can_", "_", names(n2k_AB_can))

n2k_AB <- rbind(n2k_AB_pibal, n2k_AB_can)
rm(n2k_AB_pibal, n2k_AB_can)
names(n2k_AB) <- c("FID_n2k_A", "SITE_CODE_A", "SITE_NAME_A", "AC_A", "TIPO_A",
                   "FID_n2k_B", "SITE_CODE_B", "SITE_NAME_B", "AC_B", "TIPO_B",
                   "Shape_Length", "n2K_AB_m2")
# AQUIII --> METER EL AREA DE CADA ESPACIO A Y B Y SACAR LOS % RESPECTO A ACADA ESPACIO
# Quitar duplicados
n2k_AB$SITES_CODES_AB <- paste0("SITE_CODE_A", "_", "SITE_CODE_B")
n2k_AB1 <- n2k_AB[ , c("SITES_CODES_AB", "SITE_CODE_A", "SITE_CODE_B") ] %>% 
  unique()
# Sacar el Area por cada combinacion AB
n2k_AB1_sum <- n2k_AB %>% 
  group_by(SITES_CODES_AB) %>%
  dplyr::summarise(AB_m2 = sum(n2K_AB_m2)) %>% 
  as.data.frame()

# Unir el Area anterior y el Area de cada espacio n2k
n2k_enp_1 <- merge(n2k_enp_1, n2k_enp_1sum, by = "n2k_enp_id", all.x = T)
rm(n2k_enp_1sum)
n2k_enp_1 <- merge(n2k_enp_1, n2k[ , c("SITE_CODE", "n2k_m2")], 
                   by.x = "n2k_SITE_CODE", by.y = "SITE_CODE", all.x = T)
# Sacar el porcentaje de cada espacio enp en cada sitio n2k
n2k_enp_1$porcen <- n2k_enp_1$n2k_enp_m2 / n2k_enp_1$n2k_m2 * 100
n2k_enp_1 <- n2k_enp_1[ , c("n2k_enp_id", "n2k_SITE_CODE", "enp_SITE_CODE",
                            "n2k_enp_m2",    "n2k_m2",        "porcen" )]


################################################################################
###___ Procesado cruce n2kA_API
################################################################################




################################################################################
###___ ExportaciOn de datos
################################################################################
# Crear lista con los nombres de los objetos y con los objetos (modificar segUn necesidades)
var_list <- list(
  sheets = c("enp", "n2k", "n2k_enp", "n2k_enp_1", "n2k_enp_2", "n2k_enp_3"),
  objects = list(enp, n2k, n2k_enp, n2k_enp_1, n2k_enp_2, n2k_enp_3)
)
############################################
###___ A Excel
############################################
wb <- createWorkbook()
# Estilo de la primera fila (Cabecera)
Heading <- createStyle(textDecoration = "bold", border = "Bottom")

# Iterar sobre las dos listas de la lista anterior para exportar las hojas de excel
for (i in 1:length(var_list$sheets)){
  addWorksheet(wb, var_list$sheets[i], gridLines = TRUE)
  writeData(wb, var_list$sheet[i], var_list$objects[[i]])
  addStyle(wb, var_list$sheet[i], cols = 1:ncol(var_list$objects[[i]]), rows = 1, style = Heading)
}

saveWorkbook(wb, file = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/cruce_n2k_enp.xlsx",
             overwrite=TRUE)

############################################
###___ a Access (mdb)
############################################
db <- "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/cruce_n2k_figProtec.mdb"
con <- odbcConnectAccess(db)
for (i in 1:length(objects)){
  print(paste0("Escribiendo hoja: ", i))
  sqlSave(con, var_list$objects[[i]], tablename = var_list$sheets[[i]], rownames = F)
}
odbcClose(con)
