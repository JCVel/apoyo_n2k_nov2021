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
packages <- c("sf", "RODBC", "stringr", "dplyr")
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
### Lectura scr (datos CNTRYES) 
#########################

src_pibal <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/analisisN2K_v2_compl2_limpio1.gdb", layer = "src_n2k_pibal")
src_pibal <- st_set_geometry(src_pibal, NULL)
src_pibal <- within(src_pibal, rm(hectareas))
src_can <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/analisisN2K_v2_compl2_limpio1.gdb", layer = "src_n2k_can")
src_can <- st_set_geometry(src_can, NULL)
src_can <- within(src_can, rm(HECTAREAS))

src <- rbind(src_pibal, src_can) # UniOn pibal y can
names(src)[names(src) == "Shape_Area"] <- "SrcArea"
rm(src_pibal, src_can)

#########################
### Lectura nbr
#########################
nbr_pibal <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/analisisN2K_v2_compl2_limpio1.gdb", layer = "nbr_pibal")
nbr_pibal <- st_set_geometry(nbr_pibal, NULL)

nbr_can <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/analisisN2K_v2_compl2_limpio1.gdb", layer = "nbr_can")
nbr_can <- st_set_geometry(nbr_can, NULL)

nbr <- rbind(nbr_pibal, nbr_can) # UniOn pibal y can
names(nbr)[names(nbr) == "Shape_Area"] <- "NbrArea"
rm(nbr_pibal, nbr_can)

#########################
### Lectura cruces scr con nbr 
#########################
solapes_pibal <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/analisisN2K_v2_compl2_limpio2.gdb", layer = "src_nbr_pibal")
solapes_pibal <- st_set_geometry(solapes_pibal, NULL)
names(solapes_pibal) <- gsub("_pibal", "", names(solapes_pibal))

colin_pibal <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/analisisN2K_v2_compl2_limpio2.gdb", layer = "src_nbr_pibal_intersect")
colin_pibal <- st_set_geometry(colin_pibal, NULL)
names(colin_pibal) <- gsub("_pibal", "", names(colin_pibal))

solapes_can <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/analisisN2K_v2_compl2_limpio2.gdb", layer = "src_nbr_can")
solapes_can <- st_set_geometry(solapes_can, NULL)
names(solapes_can) <- gsub("_can", "", names(solapes_can))

colin_can <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/analisisN2K_v2_compl2_limpio2.gdb", layer = "src_nbr_can_intersect")
colin_can <- st_set_geometry(colin_can, NULL)
names(colin_can) <- gsub("_can", "", names(colin_can))


solapes_ori <- rbind(solapes_pibal, solapes_can) # UniOn pibal y can
colin_ori <- rbind(colin_pibal, colin_can) # UniOn pibal y can

rm(solapes_pibal, solapes_can, colin_pibal, colin_can)

################################################################################
### Depurado de datos SOLAPES
################################################################################
#___ Identificar los polígonos de solape (src_SITE_CODE = "")
# Nota: no hay nbr_SITE_CODE son = "" porque todos los src están en nbr
solapes <- solapes_ori
solapes$solapan <- ifelse(solapes$FID_src_n2k == -1 , "NoHaySolape", "Solape")
data.frame(table(solapes$solapan))
# Eliminar los poligonos que no representan solape src_nbr
solapes <- solapes[solapes$solapan == "Solape" , ]

#___ Identificar los polígonos que han solapado con ellos mismos (RN2000 con RN2000)
# src_SITE_CODE = nbr_SITE_CODE
solapes$identicos <- ifelse(solapes$src_SITE_CODE == solapes$nbr_SITE_CODE,
                            "Identicos", "Diferentes")
data.frame(table(solapes$identicos))
# Eliminar
solapes <- solapes[solapes$identicos == "Diferentes" , ]

#___ Borrar campos de FID del ArcGIS y ordenar
solapes <- subset(solapes, select = -c(FID_src_n2k, FID_nbr, Shape_Length))
solapes <- arrange(solapes, src_SITE_CODE, nbr_SITE_CODE)

# write.table(solapes, "D:/ww/solapes.csv", sep = ";", dec = "," )

#___ Hay situaciones duplicadas con la combinación src_SITE_CODE = nbr_SITE_CODE 
# Se debe a que hay solapes de varios polig. fuente al ejecutar la herram. union en ArcGIS. 
# Ver por ejemplo las repeticiones del scr_SITE_CODE == ES0000001
# Por tanto: Sacar el Area por cada combinacion src_SITE_CODE y nbr_SITE_CODE
solapes <- solapes %>%                                     
  group_by(src_SITE_CODE, nbr_SITE_CODE) %>%
  dplyr::summarise(AREA = sum(Shape_Area)) %>% 
  as.data.frame()

solapes <- arrange(solapes, src_SITE_CODE, nbr_SITE_CODE)

# Hay que incluir los siguientes campos
# -	src_CODE – Código de la entidad fuente (source)
# -	nbr_CODE – Código de la entidad vecina (neighbour)
# -	AREA – Área común entre entidad fuente y vecina. Cuando este valor es 0 significa que existe una relación de colindancia, cuya longitud común vendrá especificada en el campo siguiente que no podrá tomar valor de 0.
# -	LENGTH – Longitud común entre la entidad fuente y vecina, cuando existe colindancia. Cuando este valor es diferente de 0 el valor del campo AREA no puede tomar otro valor que no sea 0, porque si existe colindancia no puede existir un área común entre las dos entidades.
# -	SrcArea – Área de la entidad fuente.
# -	NbrArea – Área de entidad vecina.
# -	SrcAreaP – Porcentaje de la superficie de la entidad fuente con respecto al área común (AREA)
# -	NbrAreaP – Porcentaje de la superficie de la entidad vecina con respecto al área común (AREA)
# -	RelationNbr – Relación de proximidad existente entre la entidad fuente y vecina.

# Los 3 primeros están. LENGTH para después. Ahora los siguientes:

# NbrArea
solapes <- merge(solapes, nbr[ , c("nbr_SITE_CODE", "NbrArea")], 
                 by = "nbr_SITE_CODE", all.x = T)
# SrcArea
solapes <- merge(solapes, src[ , c("src_SITE_CODE", "SrcArea")], 
                              by = "src_SITE_CODE", all.x = T)

# SrcAreaP y NbrAreaP
solapes$SrcAreaP <- solapes$AREA / solapes$SrcArea * 100
solapes$NbrAreaP <- solapes$AREA / solapes$NbrArea * 100

# RelationNbr
solapes$RelationNbr <- ""
solapes$RelationNbr <- ifelse(solapes$SrcAreaP > 98 & solapes$NbrAreaP > 98, "COINCIDENTES",
                              ifelse(solapes$NbrAreaP > 98, "Src CONTIENE A Nbr",
                                     ifelse(solapes$SrcAreaP > 98, "Src INTEGRADO EN Nbr",
                                            ifelse(solapes$SrcAreaP < 2 & solapes$NbrAreaP < 2, "SOLAPE/COLINDANCIA POR ESCALA CARTOGRAFICA",
                                                   "SOLAPAN PARCIALMENTE"))))

data.frame(table(solapes$RelationNbr))

solapes <- arrange(solapes, src_SITE_CODE, nbr_SITE_CODE)
write.table(solapes, "D:/ww/solapes2.csv", sep = ";", dec = "," )
# AQUII mirar a ver como gestionar umbrales de solapes y colindancias.


################################################################################
### Depurado de datos COLINDANCIAS
################################################################################
colin <- colin_ori
colin <- arrange(colin, src_SITE_CODE, nbr_SITE_CODE)

#___ Quitar autocolindancias RN2000 - RN2000 (src_SITE_CODE = nbr_SITE_CODE)
colin$identicos <- ifelse(colin$src_SITE_CODE == colin$nbr_SITE_CODE,
                            "Identicos", "Diferentes")
data.frame(table(colin$identicos))
colin <- colin[colin$identicos == "Diferentes" , ] # Eliminar 


#___Hay situaciones duplicadas con la combinación src_SITE_CODE = nbr_SITE_CODE 
# Se debe a que las colindancias entre 2 poligonos en ocasiones se cortan 
# Ver por ejemplo las repeticiones del scr_SITE_CODE == ES0000001
# Por tanto: Sacar la longitud por cada combinacion src_SITE_CODE y nbr_SITE_CODE
colin <- colin %>%                                     
  group_by(src_SITE_CODE, nbr_SITE_CODE) %>%
  dplyr::summarise(LENGTH = sum(Shape_Length)) %>% 
  as.data.frame()

colin <- arrange(colin, src_SITE_CODE, nbr_SITE_CODE)

#__ Si hay solape, no hay colindancia hacia el exterior

################################################################################
### Fusión solapes y colindancias
################################################################################



#########################
# Generar tabla de proximidad
#########################



# Hay que genrar un df con estos campos
# -	src_CODE – Código de la entidad fuente (source)
# -	nbr_CODE – Código de la entidad vecina (neighbour)
# -	AREA – Área común entre entidad fuente y vecina. Cuando este valor es 0 significa que existe una relación de colindancia, cuya longitud común vendrá especificada en el campo siguiente que no podrá tomar valor de 0.
# -	LENGTH – Longitud común entre la entidad fuente y vecina, cuando existe colindancia. Cuando este valor es diferente de 0 el valor del campo AREA no puede tomar otro valor que no sea 0, porque si existe colindancia no puede existir un área común entre las dos entidades.
# -	Src_Area – Área de la entidad fuente.
# -	Nbr_Area – Área de entidad vecina.
# -	Src_AreaP – Porcentaje de la superficie de la entidad fuente con respecto al área común (AREA)
# -	Nbr_AreaP – Porcentaje de la superficie de la entidad vecina con respecto al área común (AREA)
# -	RelationNbr – Relación de proximidad existente entre la entidad fuente y vecina.


#___ Campos src_CODE y nbr_CODE estAn
#___ Campos AREA y LENGTH
# El campo AREA está en cruceAPIsPolySinDupl
# En la capa de líneas del cruce, cambiar el nombre del campo de longitud de líneas de solape
names(cruceAPIsLin)[names(cruceAPIsLin) == "Shape_Length"] <- "LENGTH"

# Unir ambas fuentes de datos (longitudes comunes y áreas comunes)
cruceAPIsPolySinDupl <- bind_rows(cruceAPIsLin, cruceAPIsPolySinDupl)

# Cambiar los NAs por 0
cruceAPIsPolySinDupl[is.na(cruceAPIsPolySinDupl)] <- 0

#___ Campos Src_Area y Nbr_Area
# Src_Area (leer las capas de RN2000 con las que se hizo el cruce)
SrcArea_pibal <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/cruce_n2k_APIext.gdb", layer = "n2k_pibal_202101")
SrcArea_pibal <- st_set_geometry(SrcArea_pibal, NULL)
SrcArea_can <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/cruce_n2k_APIext.gdb", layer = "n2k_can_202101")
SrcArea_can <- st_set_geometry(SrcArea_can, NULL)

srcArea <- rbind(SrcArea_pibal, SrcArea_can)
rm(SrcArea_pibal, SrcArea_can)

names(srcArea)[names(srcArea) == "Shape_Area"] <- "SrcArea"
cruceAPIsPolySinDupl <- merge(cruceAPIsPolySinDupl, srcArea[ , c("src_SITE_CODE", "SrcArea")], 
                   by = "src_SITE_CODE", all.x = T)

# Nbr_Area
NbrArea_pibal <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/cruce_n2k_APIext.gdb", layer = "APIext_pibal")
NbrArea_pibal <- st_set_geometry(NbrArea_pibal, NULL)
NbrArea_can <- st_read(dsn = "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/cruce_n2k_APIext.gdb", layer = "APIext_can")
NbrArea_can <- st_set_geometry(NbrArea_can, NULL)

NbrArea <- rbind(NbrArea_pibal, NbrArea_can)
rm(NbrArea_pibal, NbrArea_can)

names(NbrArea)[names(NbrArea) == "Shape_Area"] <- "NbrArea"
cruceAPIsPolySinDupl <- merge(cruceAPIsPolySinDupl, NbrArea[ , c("nbr_SITE_CODE", "NbrArea")], 
                   by = "nbr_SITE_CODE", all.x = T)

#___ Campos Src_AreaP y Nbr_AreaP
cruceAPIsPolySinDupl$SrcAreaP <- cruceAPIsPolySinDupl$AREA / cruceAPIsPolySinDupl$SrcArea * 100
cruceAPIsPolySinDupl$NbrAreaP <- cruceAPIsPolySinDupl$AREA / cruceAPIsPolySinDupl$NbrArea * 100

#___ Campo RelationNbr
df <- cruceAPIsPolySinDupl
df$RelationNbr <- ""
df$RelationNbr <- ifelse(df$LENGTH != 0, "COLINDANTES", 
                         ifelse(df$SrcAreaP > 98 & df$NbrAreaP > 98, "COINCIDENTES",
                                ifelse(df$NbrAreaP > 98, "Src CONTIENE A Nbr",
                                       ifelse(df$SrcAreaP > 98, "Src INTEGRADO EN Nbr",
                                              ifelse(df$SrcAreaP < 2 & df$NbrAreaP < 2, "SOLAPE/COLINDANCIA POR ESCALA CARTOGRAFICA",
                                              "SOLAPAN PARCIALMENTE")))))
df$NODE_COUNT <- 0
df <- df[ , c("src_SITE_CODE", "nbr_SITE_CODE", "AREA", "LENGTH", "NODE_COUNT", 
              "SrcArea", "NbrArea", "SrcAreaP", "NbrAreaP", "RelationNbr")]
df$src_SITE_CODE <- paste0(df$src_SITE_CODE,  "_RN2000")
cruceAPIsPolySinDupl <- df
rm(df)
# Comprobaciones:
# Colindantes: OK; Coincidentes: OK; Scr CONTIENE A Nbr: OK; 
# Scr INTEGRADO EN NBR: OK; "SOLAPE/COLINDANCIA POR ESCALA CARTOGRAFICA": OK;
# "SOLAPAN PARCIALMENTE": OK

# Ver resumen frecuencias de cada indicador
data.frame(table(cruceAPIsPolySinDupl$RelationNbr))

#########################
# Crear RNProxiAnt =  Unir cruces nuevos depurados (cruceAPIsPolySinDupl) a los antiguos (RNProxiComplSinDupl)
#########################
# cruceAPIsPolySinDupl pasa a ser --> RNProxiAnt (limpiando la tabla)
RNProxiAnt <- RNProxiComplSinDupl[ , c("src_SITE_CODE2", "nbr_SITE_CODE2",
                                        "AREA", "LENGTH", "NODE_COUNT", 
                                        "SrcArea", "NbrArea", 
                                        "SrcAreaP", "NbrAreaP", 
                                        "RelaNbr")]

names(RNProxiAnt)[names(RNProxiAnt) == "src_SITE_CODE2"] <- "src_SITE_CODE"
names(RNProxiAnt)[names(RNProxiAnt) == "nbr_SITE_CODE2"] <- "nbr_SITE_CODE"
names(RNProxiAnt)[names(RNProxiAnt) == "RelaNbr"] <- "RelationNbr"

# Unión
RNProxiAct <- rbind(RNProxiAnt, cruceAPIsPolySinDupl)

#########################
# Incluir los sitios RN2000 que no solapan/colindan con EPs
#########################
# Primero obtener datos
print(paste0("Número de registros de la tabla RNProxiAct: ", nrow(RNProxiAct)))
NumSitiosQueCruzan <- length(unique(RNProxiAct$src_SITE_CODE))
print(paste0("Número de sitios RN2000 que cruzan con EPs: ", NumSitiosQueCruzan))
print(paste0("Número de sitios RN2000 en el CNTRYES (enero 2021): ", nrow(CNTRYES)))
print(paste0("Número de sitios RN2000 que no solapan/colindan con EPs: ", 
             nrow(CNTRYES) - NumSitiosQueCruzan))
# vector con los sitios RN2000 que solapan/colindan con otros EP
sitiosQueCruzan <- unique(RNProxiAct$src_SITE_CODE)
sitiosQueCruzan <- gsub("_RN2000", "", sitiosQueCruzan)
# En el CNTRYES marcar los que no solapan/colindan en base al vector anterior
CNTRYESmarcado <- CNTRYES
CNTRYESmarcado$Nbr <- ifelse(CNTRYESmarcado$SITE_CODE %in% sitiosQueCruzan, "SiNbr", "NoNbr")
CNTRYESnoNbr <- CNTRYESmarcado[CNTRYESmarcado$Nbr == "NoNbr", c("SITE_CODE")]
CNTRYESnoNbr <- paste0(CNTRYESnoNbr, "_RN2000")
# crear un df con los sitios RN2000 que no solapan/colindan
RNProxiNoNbr <- data.frame(
  src_SITE_CODE = CNTRYESnoNbr,
  nbr_SITE_CODE = "Sin solape/colindancia con EP",
  RelationNbr = "Sin solape/colindancia con EP"
)

# Unir a la tabla con las relaciones de proximidad
RNProxiAct <- bind_rows(RNProxiAct,RNProxiNoNbr)
RNProxiAct[is.na(RNProxiAct)] <- 0

#########################
# Crear una tabla (RNProxiActMerge) que tenga la información de las otras dos
# EPAct en RNProxiAct (no haría falta como tal, pero puede ser útil)
#########################

RNProxiActMerge <- merge(RNProxiAct, EPAct[ , c("SITE_CODE", "Nombre_", "AC", "TIPO")],
                         by.x = "src_SITE_CODE", by.y = "SITE_CODE", all.x = T)

names(RNProxiActMerge)[names(RNProxiActMerge) %in% c("Nombre_", "AC", "TIPO")] <- c("src_Nombre", "src_AC", "src_TIPO")

RNProxiActMerge <- merge(RNProxiActMerge, EPAct[ , c("SITE_CODE", "Nombre_", "AC", "TIPO", "jerar1")],
                         by.x = "nbr_SITE_CODE", by.y = "SITE_CODE", all.x = T)
names(RNProxiActMerge)[names(RNProxiActMerge) %in% c("Nombre_", "AC", "TIPO", "jerar1")] <- c("nbr_Nombre", "nbr_AC", "nbr_TIPO", "nbr_jerar1")


#########################
# Ordenar columnas y filas de las tablas a exportar
#########################
EPAct <- EPAct[ , c("SITE_CODE", "Nombre_", "FIGURA", "DESIGNACIO", "AC", "TIPO", "jerar1")]
names(EPAct)[names(EPAct) == "Nombre_"] <- "Nombre"
EPAct$TIPO2 <- ifelse(EPAct$TIPO == "A", "A (ZEPA)",
                   ifelse(EPAct$TIPO == "B", "B (LIC)", "C (LIC y ZEPA)"))

EPAct <- arrange(EPAct, FIGURA, SITE_CODE)

RNProxiAct <- arrange(RNProxiAct, src_SITE_CODE, nbr_SITE_CODE)

RNProxiActMerge <- RNProxiActMerge[ , c("src_SITE_CODE", "nbr_SITE_CODE", "AREA",
                                        "LENGTH", "NODE_COUNT", "SrcArea", "NbrArea",
                                        "SrcAreaP", "NbrAreaP", "RelationNbr",
                                        "src_Nombre", "src_AC", "src_TIPO",
                                        "nbr_Nombre", "nbr_AC", "nbr_TIPO", "nbr_jerar1")]  

RNProxiActMerge <- arrange(RNProxiActMerge, src_SITE_CODE, nbr_SITE_CODE)


#########################
# Modificaciones de datos finales
#########################
#___ En los sitios RN2000 se quita el sufijo "_RN2000"
RNProxiAct$src_SITE_CODE <- gsub("_RN2000", "", RNProxiAct$src_SITE_CODE)
RNProxiAct$nbr_SITE_CODE <- gsub("_RN2000", "", RNProxiAct$nbr_SITE_CODE)
EPAct$SITE_CODE <- gsub("_RN2000", "", EPAct$SITE_CODE)
RNProxiActMerge$src_SITE_CODE <- gsub("_RN2000", "", RNProxiActMerge$src_SITE_CODE)
RNProxiActMerge$nbr_SITE_CODE <- gsub("_RN2000", "", RNProxiActMerge$nbr_SITE_CODE)

#___ srcArea --> la oficial de CNTRYES



#___ AREA --> multiplicando ScrAreaP * srcArea




#########################
# Crear dos DD: uno para src y otro para nbr
#########################

#____ DD para src
DD_src <- EPAct[EPAct$FIGURA == "RN2000", ]
#____ DD para nbr
# sacar todos los códigos nbr que han cruzado
nbrProxi <- unique(RNProxiAct$nbr_SITE_CODE)
# Usar el cector anterior para marcar los que estén presentes en la tabla de EPAct
DD_nbr <- EPAct
DD_nbr$nbr <- ifelse(DD_nbr$SITE_CODE %in% nbrProxi, "SiNbr", "NoNbr")
DD_nbr <- DD_nbr[DD_nbr$nbr == "SiNbr", ] # Seleccionar solo los que interesan
DD_nbr <- subset(DD_nbr, select = -c(nbr))
DD_nbr <- bind_rows(DD_nbr, )
DD_nbr[nrow(DD_nbr) + 1,] <- c("Sin solape/colindancia con EP", "", "", "", "", "","", "")



################################################################################
###___ ExportaciOn de datos
################################################################################
# Crear lista con los nombres de los objetos y con los objetos (modificar segUn necesidades)
var_list <- list(
  sheets = c("EPAct", "RNProxiAct", "RNProxiActMerge", "DD_src", "DD_nbr"),
  objects = list(EPAct, RNProxiAct, RNProxiActMerge, DD_src, DD_nbr)
)

############################################
###___ a Access (mdb)
############################################
db <- "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/RNATURA_AnalisisProximidad_int.mdb"
con <- odbcConnectAccess(db)
for (i in 1:length(var_list$objects)){
  print(paste0("Escribiendo hoja: ", var_list$sheets[[i]]))
  sqlSave(con, var_list$objects[[i]], tablename = var_list$sheets[[i]], rownames = F, safer = F)
}
odbcClose(con)
