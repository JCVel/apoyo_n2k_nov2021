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

#########################
### Lectura cruces n2k con APIs no incluidas en BD RNATURA_AnalisisProximidad.mdb
#########################
cruceAPIsPoly_pibal <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/cruce_n2k_APIext.gdb", layer = "n2k_APIext_pibal")
cruceAPIsPoly_pibal <- st_set_geometry(cruceAPIsPoly_pibal, NULL)
names(cruceAPIsPoly_pibal) <- gsub("_pibal", "", names(cruceAPIsPoly_pibal))
cruceAPIsLin_pibal <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/cruce_n2k_APIext.gdb", layer = "n2k_APIext_pibal_intersect")
cruceAPIsLin_pibal <- st_set_geometry(cruceAPIsLin_pibal, NULL)
names(cruceAPIsLin_pibal) <- gsub("_pibal", "", names(cruceAPIsLin_pibal))


cruceAPIsPoly_can <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/cruce_n2k_APIext.gdb", layer = "n2k_APIext_can")
cruceAPIsPoly_can <- st_set_geometry(cruceAPIsPoly_can, NULL)
names(cruceAPIsPoly_can) <- gsub("_can", "", names(cruceAPIsPoly_can))
cruceAPIsLin_can <- st_read(dsn = "D:/JCV_interno/19_n2k_apoyoNov2021/1_datosSalida/cruce_n2k_APIext.gdb", layer = "n2k_APIext_can_intersect")
cruceAPIsLin_can <- st_set_geometry(cruceAPIsLin_can, NULL)
names(cruceAPIsLin_can) <- gsub("_can", "", names(cruceAPIsLin_can))


cruceAPIsPoly <- rbind(cruceAPIsPoly_pibal, cruceAPIsPoly_can) # UniOn pibal y can
cruceAPIsLin <- rbind(cruceAPIsLin_pibal, cruceAPIsLin_can) # UniOn pibal y can

rm(cruceAPIsPoly_pibal, cruceAPIsPoly_can, cruceAPIsLin_pibal, cruceAPIsLin_can)

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
RNProxi <- desconcatenaID(RNProxi, "nbr_SITE_CODE")
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
### Añadir a la tabla EP las API que no estAn incluidas en la BD => EPcompl
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
# Meter las nuevas API en la tabla de EP de la BD
EPcompl <- rbind(EP, APIext)
# cuentaEP <- EPcompl$SITE_CODE_pre
# cuentaEP <- unique(cuentaEP)

#########################
### Depurar tabla EPcompl, dejAndola sin duplicados RN2000
#########################
# Primero indicar el tipo de espacio (A, B, c) a partir de los datos del CNTRYES
EPcompl <- merge(EPcompl, CNTRYES[ , c("SITE_CODE", "AC", "TIPO", "Shape_Area")], 
               by.x = "SITE_CODE_pre", by.y = "SITE_CODE", all.x = T)
# Comprobar que todos los tipos C estAn duplicados (y ya de paso comprobar si 
# hubiera otros duplicados) 
freq <- table(EPcompl$SITE_CODE_pre)
freq <- as.data.frame(freq)
EPcompl <- merge(EPcompl, freq, by.x = "SITE_CODE_pre", by.y = "Var1", all.x = T)
# ConclusiOn: todos los tipos C estan duplic. y no hay mAs duplic

# Todo lo que sea LIC, ZEC, ZEPA en tabla EP pasa a ser RN2000
EPcompl$FIGURA <- ifelse(EPcompl$FIGURA == "LIC" | EPcompl$FIGURA == "ZEC" |
                         EPcompl$FIGURA == "ZEPA", "RN2000", EPcompl$FIGURA )
EPcompl$SITE_CODE_suf <- EPcompl$FIGURA
# quitar duplicados (que solo son tipo C de RN2000). Para ello usar el campo SITE_CODE_pre, 
# puesto que el campo SITE_CODE (el original de la tabla EP que tienen prefijo y sufijo) es un artificio
# y el campo Nombre_  genera duplicados extra porque hay registros con elmismo nombre con y sin tildes
EPcomplSinDupl <- distinct(EPcompl, SITE_CODE_pre, .keep_all = TRUE)

# Cambiar la nomenclatura de los sufijos
EPcomplSinDupl$SITE_CODE_suf <- ifelse(EPcomplSinDupl$SITE_CODE_suf == "ENP", "CDDA", EPcomplSinDupl$SITE_CODE_suf)

# Anadir un nombre que falta (ya faltaba en la tabla EP de la BD original)
EPcomplSinDupl[EPcomplSinDupl$SITE_CODE_pre == "2453", "Nombre_"] <- "Parque Nacional Marítimo-Terrestre de las Islas Atlánticas"

# Indicar primer nivel jerArquico en la tabla EPcomplSinDupl (ENP, RN2000, API)
EPcomplSinDupl$jerar1 <- ifelse(EPcomplSinDupl$SITE_CODE_suf != "ENP" & EPcomplSinDupl$SITE_CODE_suf != "RN2000",
                            "API", EPcomplSinDupl$SITE_CODE_suf)

# Crear la versión actualizada con los campos adecuados
EPAct <- subset(EPcomplSinDupl, select = -c(SITE_CODE))
EPAct$SITE_CODE <- paste0(EPAct$SITE_CODE_pre, "_", EPAct$SITE_CODE_suf)
EPAct <- subset(EPAct, select = -c(SITE_CODE_pre, SITE_CODE_suf, Shape_Area, Freq))


#########################
### En la tabla RNProxi, eliminar los duplicados de los tipos C de RN2000
#########################
# Meter los tipos de espacios (A, B, C) del CNTRYES
RNProxiCompl <- merge(RNProxi, CNTRYES[ , c("SITE_CODE", "TIPO")], 
                    by.x = "src_SITE_CODE_pre", by.y = "SITE_CODE", all.x = T)

# FunciOn para cambiar los sufijos LIC, ZEC, ZEPA por _RN2000
cambiaARN2000 <- function(df, field){
  # Esta funciOn cambia los sufijos LIC, ZEC, ZEPA por _RN2000
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

# cambiar los sufijos LIC, ZEC, ZEPA por _RN2000
RNProxiCompl <- cambiaARN2000(RNProxiCompl, "src_SITE_CODE")
RNProxiCompl <- cambiaARN2000(RNProxiCompl, "nbr_SITE_CODE")

# Detectar las combinaciones src_SITE_CODE2 y nbr_SITE_CODE2 iguales (solo ocurrirAn
# en los tipos C al cruzarse consigo mismo)
RNProxiCompl$igual <- ifelse(RNProxiCompl$src_SITE_CODE2 == RNProxiCompl$nbr_SITE_CODE2,
                              "Igual", "Diferente")
# Eliminar dichas combinaciones
RNProxiComplSinDupl <- RNProxiCompl[!RNProxiCompl$igual == "Igual", ]

# Eliminar los duplicados al combinar src_SITE_CODE2, nbr_SITE_CODE2 y relationNbr (es decir, 
# los tipo C que tienen los datos por duplicado). OJO! Se incluye el campo relationNbr
# porque hay casos en los que se duplica un registro nbr al tener una relación de 
# colindancia (campo LENGTH != 0) Ver caso: src = ES0000506 y sus duplicados nbr ES000048
# uno es de solape y el otro de colindancia
RNProxiComplSinDupl <- distinct(RNProxiComplSinDupl, src_SITE_CODE2, nbr_SITE_CODE2, RelaNbr,
                                 .keep_all = TRUE)

# NOTA: RNProxiComplSinDupl es la tabla RNProxi quitando los duplicados tipo C
# Control de Calidad: Se han hecho varias comprobaciones para ver si en RNProxiComplSinDupl
# han desaparecido los duplicados de los sitios C que estaban en RNProxi
# NOTA: ahora hay que meter los datos de los cruces con el resto de API

#########################
# Depurar los datos de los cruces con los API que faltan
#########################
# Dejar solo los registros en los que hay cruce entre src y nbr
cruceAPIsPoly <- cruceAPIsPoly[cruceAPIsPoly$src_SITE_CODE != "" & cruceAPIsPoly$nbr_SITE_CODE != "", ]
# En la tabla cruceAPIsPoly hay registros con combinaciones src_SITE_CODE y nbr_SITE_CODE duplicados
# Se debe a que hay solapes de varios polig. fuente al ejecutar la herram. union en ArcGIS. 
# Ver por ejemplo las repeticiones del scr_SITE_CODE == ES0000015
# Por tanto: Sacar el Area por cada combinacion src_SITE_CODE y nbr_SITE_CODE
cruceAPIsPolySinDupl <- cruceAPIsPoly %>%                                     
  group_by(src_SITE_CODE, nbr_SITE_CODE) %>%
  dplyr::summarise(AREA = sum(Shape_Area)) %>% 
  as.data.frame()
# Comprobaciones:
# a = cruceAPIsPoly[cruceAPIsPoly$src_SITE_CODE == "ES4240017", ]
# a = a[a$nbr_SITE_CODE == "ext6_Geoparque", ]
# sum(a$Shape_Area)


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
# Meter campos de la tabla EPAct en RNProxiAct (no haría falta pero bueno) = RNProxiActMerge
#########################

RNProxiActMerge <- merge(RNProxiAct, EPAct[ , c("SITE_CODE", "Nombre_", "AC", "TIPO")], 
                         by.x = "src_SITE_CODE", by.y = "SITE_CODE", all.x = T)

names(RNProxiActMerge)[names(RNProxiActMerge) %in% c("Nombre_", "AC", "TIPO")] <- c("src_Nombre", "src_AC", "src_TIPO")

RNProxiActMerge <- merge(RNProxiActMerge, EPAct[ , c("SITE_CODE", "Nombre_", "AC", "TIPO", "jerar1")], 
                         by.x = "nbr_SITE_CODE", by.y = "SITE_CODE", all.x = T)
names(RNProxiActMerge)[names(RNProxiActMerge) %in% c("Nombre_", "AC", "TIPO", "jerar1")] <- c("nbr_Nombre", "nbr_AC", "nbr_TIPO", "nbr_jerar1")


#########################
# Ordenar campos de las tres tablas a exportar
#########################
EPAct <- EPAct[ , c("SITE_CODE", "Nombre_", "FIGURA", "DESIGNACIO", "AC", "TIPO", "jerar1")]

RNProxiActMerge <- RNProxiActMerge[ , c("src_SITE_CODE", "nbr_SITE_CODE", "AREA",
                                        "LENGTH", "NODE_COUNT", "SrcArea", "NbrArea",
                                        "SrcAreaP", "NbrAreaP", "RelationNbr",
                                        "src_Nombre", "src_AC", "src_TIPO",
                                        "nbr_Nombre", "nbr_AC", "nbr_TIPO", "nbr_jerar1")]  

################################################################################
###___ ExportaciOn de datos
################################################################################
# Crear lista con los nombres de los objetos y con los objetos (modificar segUn necesidades)
var_list <- list(
  sheets = c("EPAct", "RNProxiAct", "RNProxiActMerge"),
  objects = list(EPAct, RNProxiAct, RNProxiActMerge)
)

############################################
###___ a Access (mdb)
############################################
db <- "D:/JCV_interno/19_N2K_apoyoNov2021/1_datosSalida/RNATURA_AnalisisProximidad.mdb"
con <- odbcConnectAccess(db)
for (i in 1:length(var_list$objects)){
  print(paste0("Escribiendo hoja: ", var_list$sheets[[i]]))
  sqlSave(con, var_list$objects[[i]], tablename = var_list$sheets[[i]], rownames = F, safer = F)
}
odbcClose(con)
