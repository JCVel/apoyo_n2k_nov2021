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
packages <- c("sf", "RODBC", "stringr", "dplyr", "readxl")
# IntalaciOn de paquetes
installed_packages <- packages %in% rownames(installed.packages())
if (any(installed_packages == FALSE)) {
  install.packages(packages[!installed_packages])
}
# Carga de paquetes
invisible(lapply(packages, library, character.only = TRUE))
rm(installed_packages, packages)

# Rutas entrada y salida
pathEntrada <- "D:/3081165_Tarea1.2/entrada"
pathSalida <- "D:/3081165_Tarea1.2/salida"
################################################################################
###___ Procesado de la BD RNATURA_AnalisisProximidad.mdb
################################################################################
#########################
### Lectura de EP y RNProxi 
#########################
# Se lee a partir de BD RNATURA_AnalisisProximidad.mdb
# (configurar R a 32bit). Para mAs detalles ver https://www.roelpeters.be/solved-importing-microsoft-access-files-accdb-mdb-in-r/

pathBD <- file.path(pathEntrada, "RNATURA_AnalisisProximidad.mdb")
con <- odbcConnectAccess(pathBD) # Establecer conexiOn
RNProxi<- sqlFetch(con, "RNATURA_proximidad") # Leer tabla de interEs
EP <- sqlFetch(con, "ESPACIOS_PROTEGIDOS") 
odbcClose(con) # cerrar conexiOn con el mdb
rm(con)

#########################
### Lectura datos CNTRYES
#########################
pathGDB <- file.path(pathSalida, "RNATURA_AnalisisProximidad_Procesado.gdb")

CNTRYES_pibal <- st_read(dsn = pathGDB, layer = "src_n2k_202101_pibal")
CNTRYES_pibal <- st_set_geometry(CNTRYES_pibal, NULL)
CNTRYES_pibal <- within(CNTRYES_pibal, rm(hectareas))
CNTRYES_can <- st_read(dsn = pathGDB, layer = "src_n2k_202101_can")

CNTRYES_can <- st_set_geometry(CNTRYES_can, NULL)
CNTRYES_can <- within(CNTRYES_can, rm(HECTAREAS))

CNTRYES <- rbind(CNTRYES_pibal, CNTRYES_can) # UniOn pibal y can
rm(CNTRYES_pibal, CNTRYES_can)

#########################
### Tablas EP y RNProxi: Desconcatenar campos xxx_SITE_CODE
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
RNProxi_ori <- RNProxi
EP_ori <- EP
RNProxi <- desconcatenaID(RNProxi, "src_SITE_CODE")
RNProxi <- desconcatenaID(RNProxi, "nbr_SITE_CODE")
EP <- desconcatenaID(EP, "SITE_CODE")

#########################
### Tabla EP: Comprobar nUmero de registros RN2000 en BD sea igual que en CNTRYES
#########################
# Sacar solo los registros de RNATURA 2000
EPRN <- EP[EP$SITE_CODE_suf == "LIC" | EP$SITE_CODE_suf == "ZEC" |
             EP$SITE_CODE_suf == "ZEPA", ]

# Seleccionar los campos de interes para eliminar duplicados
EPRN <- data.frame(
  SITE_CODE_pre = EPRN$SITE_CODE_pre)
EPRNUniq <- unique(EPRN) # Salen 1857 sitios RN2000. EstA bien. Coincide con
# los datos de RN2000 del CNTRYES
rm(EPRN, EPRNUniq)
#########################
### Tabla EP: Añadir a la tabla EP las API que no estAn incluidas en la BD => EPcompl
#########################
APIext_pibal <- st_read(dsn = pathGDB, layer = "nbr_APIextra_pibal")
APIext_pibal <- st_set_geometry(APIext_pibal, NULL)
APIext_can <- st_read(dsn = pathGDB, layer = "nbr_APIextra_can")
APIext_can <- st_set_geometry(APIext_can, NULL)

APIext_ori <- rbind(APIext_pibal, APIext_can) # UniOn pibal y can
rm(APIext_pibal, APIext_can)

# Desconcatenar el SITE_CODE
APIext <- APIext_ori
APIext <- desconcatenaID(APIext, "SITE_CODE")
# Dejar los mismos campos que en la tabla EP de la BD
APIext <- subset(APIext, select = -c(Shape_Length, Shape_Area))
# Meter las nuevas API en la tabla de EP de la BD
EPcompl <- rbind(EP, APIext)
# cuentaEP <- EPcompl$SITE_CODE_pre
# cuentaEP <- unique(cuentaEP)

#########################
### Tabla EP: Depurar tabla EPcompl, dejAndola sin duplicados RN2000
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
rm(freq)
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
EPcomplSinDupl$jerar1 <- ifelse(EPcomplSinDupl$SITE_CODE_suf == "CDDA", "ENP", 
                                ifelse(EPcomplSinDupl$SITE_CODE_suf == "RN2000", "RN2000",
                                       "API"))

# Crear la versión actualizada con los campos adecuados
EPAct <- subset(EPcomplSinDupl, select = -c(SITE_CODE))
EPAct$SITE_CODE <- paste0(EPAct$SITE_CODE_pre, "_", EPAct$SITE_CODE_suf)
EPAct <- subset(EPAct, select = -c(SITE_CODE_pre, SITE_CODE_suf, Shape_Area, Freq))
rm(EP, EPcompl, EPcomplSinDupl)

#########################
### Tabla RNProxi: eliminar los duplicados de los tipos C de RN2000
#########################
# Meter los tipos de espacios (A, B, C) del CNTRYES
RNProxi <- merge(RNProxi, CNTRYES[ , c("SITE_CODE", "TIPO")], 
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
      # print(paste0(pre, "_RN2000"))
      
      df[[field2]][[row]] <- paste0(pre, "_RN2000")
    } else {
      df[[field2]][[row]] <- df[[field]][[row]] 
    }
  }
  return(df)
}

# cambiar los sufijos LIC, ZEC, ZEPA por _RN2000
RNProxi <- cambiaARN2000(RNProxi, "src_SITE_CODE")
RNProxi <- cambiaARN2000(RNProxi, "nbr_SITE_CODE")

# Detectar las combinaciones src_SITE_CODE2 y nbr_SITE_CODE2 iguales (solo ocurrirAn
# en los tipos C al cruzarse consigo mismo)
RNProxi$igual <- ifelse(RNProxi$src_SITE_CODE2 == RNProxi$nbr_SITE_CODE2,
                             "Igual", "Diferente")
# Eliminar dichas combinaciones
RNProxi <- RNProxi[!RNProxi$igual == "Igual", ]

# Meter ID 
RNProxi$sufRN2000 <- ifelse(RNProxi$src_SITE_CODE_suf == "LIC" | RNProxi$src_SITE_CODE_suf == "ZEC", "LIC/ZEC",
                            ifelse(RNProxi$src_SITE_CODE_suf == "ZEPA", "ZEPA", ""))
RNProxi$id <- paste0(RNProxi$src_SITE_CODE_pre, "_", RNProxi$SrcArea, "_", RNProxi$sufRN2000 )                            
                                        

# Marcar areas de tipos C más próximas a las del CNTRYES
# seleccionar los tipos C y quitar duplicados por SITE_CODE y area
areasTiposC <- unique(RNProxi[RNProxi$TIPO == "C" , c("id", "src_SITE_CODE_pre", "SrcArea")])
frec <- data.frame(table(areasTiposC$src_SITE_CODE_pre))
areasTiposC <- merge(areasTiposC, frec, by.x = "src_SITE_CODE_pre", by.y = "Var1", all.x = T)
rm(frec)
areasTiposC2 <- areasTiposC %>%
  group_by(src_SITE_CODE_pre) %>%
  summarize(
    AreasIguales = max(SrcArea) - min(SrcArea)
  )
areasTiposC2$Iguales <- ifelse(areasTiposC2$AreasIguales == 0, "Iguales", "Diferentes")
# Indicar qué sitios RN2000 tipo C tienen areas diferentes
areasTiposC <- merge(areasTiposC, areasTiposC2[ , c("src_SITE_CODE_pre", "Iguales")],
                     by = "src_SITE_CODE_pre", all.x = T)
rm(areasTiposC2)
areasTiposC <- areasTiposC[areasTiposC$Iguales == "Diferentes", ]

# Meter el área oficial del CNTRYES Enero 2021
areasTiposC <- merge(areasTiposC, CNTRYES[ , c("SITE_CODE", "Shape_Area")],
                     by.x = "src_SITE_CODE_pre", by.y = "SITE_CODE", all.x = T)
names(areasTiposC)[names(areasTiposC) == "Shape_Area"] <- "AreaOfi"


# OJO: Para los sitios tipo C se han duplicado los registros en LIC/ZEC y ZEPA. 
# PERO! OJO: cuando se subdividen hay casos en los que LIC/ZEC tienen una superficie (scrArea)
# diferente a la de la ZEPA (ver caso ES000048)... ¿Porqué? No se sabe. 
# PAra elegir qué área se deja en el campo srcArea se selecciona aquella que más 
# se aproxime a la oficial del CNTRYES

# Calcular las diferencias en m2 y en % absoluto entre las 2 áreas de un mismo sitio C
DifInternaSrcArea <- areasTiposC %>%
  group_by(src_SITE_CODE_pre) %>%
  summarize(
    DifInternaSrcArea = max(SrcArea) - min(SrcArea),
    DifInternaSrcAreaP = abs(((max(SrcArea) * 100) / min(SrcArea)) - 100),
  )

# Unir los datos anteriores a areasTiposC
areasTiposC <- merge(areasTiposC, DifInternaSrcArea, all.x = T )

# Calcular las diferencias en m2 y en % (P) con las áreas oficiales
areasTiposC$AreasDif <- areasTiposC$SrcArea - areasTiposC$AreaOfi
areasTiposC$AreasDifP <- ((areasTiposC$SrcArea * 100) / areasTiposC$AreaOfi) - 100
areasTiposC$AreasDifPAbs <- abs(areasTiposC$AreasDifP)

# # Sacar un ID por cada sitio y su area
# areasTiposC$AreasDifPAbs <- as.character(areasTiposC$AreasDifPAbs)
# areasTiposC$id <-  paste0(areasTiposC$src_SITE_CODE_pre, "_", areasTiposC$AreasDifPAbs)
# Obtener el área más cercana (de las dos) al datos de CNTRYES y sacar el ID
areasTiposC2 <- aggregate(AreasDifPAbs ~ src_SITE_CODE_pre, areasTiposC, function(x) min(x))

areasTiposC2$Cercano <- "MasCercano"

# Unir por AreasDifPAbs
areasTiposC <- merge(areasTiposC, areasTiposC2[, c("AreasDifPAbs", "Cercano" )], by = "AreasDifPAbs", all.x = T)
areasTiposC[is.na(areasTiposC)] <- 0
rm(areasTiposC2)

# Dejar solo los registros con las areas adecuadas
# areasTiposC <- areasTiposC[areasTiposC$Cercano == "MasCercano", ]
areasTiposC <- areasTiposC[ , c("id", "SrcArea", "DifInternaSrcArea", "DifInternaSrcAreaP", "AreaOfi", "AreasDif","AreasDifP", "AreasDifPAbs","Cercano"   )]
          
  
areasTiposCLimpio <- areasTiposC[areasTiposC$Cercano == "MasCercano", ]
areasTiposCLimpio <- areasTiposCLimpio[ , c("id", "SrcArea", "Cercano")]

# Limpiar RNProxi
RNProxi <- RNProxi[ , c("id", "src_SITE_CODE2", "TIPO", "nbr_SITE_CODE2", 
                        "AREA", "LENGTH", "NODE_COUNT", 
                        "SrcArea", "NbrArea", "SrcAreaP", "NbrAreaP", "RelaNbr")]
names(RNProxi)[names(RNProxi) == "TIPO"] <- "Src_TIPO"  
# Quitar los sufijos RN2000 (al final no ses usan)
RNProxi$src_SITE_CODE2 <- gsub("_RN2000", "", RNProxi$src_SITE_CODE2)
RNProxi$nbr_SITE_CODE2 <- gsub("_RN2000", "", RNProxi$nbr_SITE_CODE2)


# Meter las áreas más cercanas al CNTRYES para los Tipos C con 2 áreas.
tiposCaLimpiarVector <- areasTiposC$id
# Seleccionar los Tipos C con dos áreas, que hay que seleccionar
tiposCaLimpiar <- RNProxi[RNProxi$id %in% tiposCaLimpiarVector, ] 

tiposCaLimpiar <- merge(tiposCaLimpiar, areasTiposCLimpio, by = "id", all.x = T)
tiposCaLimpiar <- arrange(tiposCaLimpiar, src_SITE_CODE2, nbr_SITE_CODE2)
tiposCaLimpiar[is.na(tiposCaLimpiar)] <- 0

tiposCLimpios <- tiposCaLimpiar[tiposCaLimpiar$Cercano == "MasCercano", ]
length(unique(tiposCLimpios$src_SITE_CODE2)) #Sale 81, esta bien
tiposCLimpios <- subset(tiposCLimpios, select = -c(SrcArea.y))
names(tiposCLimpios)[names(tiposCLimpios) == "SrcArea.x"] <- "SrcArea"

# write.table(tiposCaLimpiar, "D:/ww/tiposCaLimpiar.csv", sep = ";", dec = ",")
# write.table(tiposCLimpios, "D:/ww/tiposCLimpios.csv", sep = ";", dec = ",")

# Quitar de RNProxi los tipos C con dos áreas y meter los dichos registros ya 
# limpiados, sin duplicados y con el área SrcArea más cercana al CNTRYES
tiposCLimpiosVector <- unique(tiposCLimpios$src_SITE_CODE2)

RNProxi$aLimpiarTipoC <- ifelse(RNProxi$src_SITE_CODE2 %in% tiposCLimpiosVector, "Limpiar", "Dejar")
data.frame(table(RNProxi$aLimpiarTipoC))

RNProxi <- RNProxi[RNProxi$aLimpiarTipoC == "Dejar", ] #quitar
RNProxi <- bind_rows(RNProxi, tiposCLimpios)

# Se quitan ahora los duplicados de los tipos C que tienen un área identica tanto si son LIC/ZEC como ZEPA
# Eliminar los duplicados al combinar src_SITE_CODE2, nbr_SITE_CODE2 y relationNbr (es decir, 
# los tipo C que tienen los datos por duplicado). OJO! Se incluye el campo relationNbr
# porque hay casos en los que se duplica un registro nbr al tener una relación de 
# colindancia (campo LENGTH != 0) Ver caso: src = ES0000506 y sus duplicados nbr ES000048
# uno es de solape y el otro de colindancia
RNProxi <- distinct(RNProxi, src_SITE_CODE2, nbr_SITE_CODE2, RelaNbr,
                                .keep_all = TRUE)



# NOTA: RNProxi es la tabla RNProxi_ori quitando los duplicados tipo C. 
# EN aquellos cuyo área era diferente en unción de si era LIC/ZEPA, se han dejado 
# los registros de uno de ellos (eliminando los duplicados). Los que se han dejado
# son los que el SrcArea más se aproxima al área oficial de CNTRYES
# Control de Calidad: Se han hecho varias comprobaciones para ver si en RNProxiComplSinDupl
# han desaparecido los duplicados de los sitios C que estaban en RNProxi_ori
# NOTA: ahora hay que meter los datos de los cruces con el resto de API


# NOTA: Algunas superficies RN2000 están mal... Los casos más extremos son por defecto. P.ej:
# ver ES6180009, ES1200048, ES1200023, ES6120023 donde las superficies están infravaloradas
# en un -61.18, -49.24, -39.04, -19.63 respectivamente. Por lo que parece han cortado los sitios RN2000 con las líneas autonómicas.
# Cuando se hagan los informes en access se mostrará también la oficial dle CNTRYES

rm(areasTiposC, areasTiposCLimpio, DifInternaSrcArea, tiposCaLimpiar, tiposCLimpios, tiposCaLimpiarVector, tiposCLimpiosVector)

#########################
### Tabla RNProxi: Inclusión de los cruces n2k con APIs no incluidas en BD RNATURA_AnalisisProximidad.mdb
#########################
# ___Lectura y depuración de los cruces de RN2000 con las APIs no incluidas en BD
cruceAPIs_pibal <- st_read(dsn = pathGDB, layer = "polyNeigh_pibal")
cruceAPIs_can <- st_read(dsn = pathGDB, layer = "polyNeigh_can")
cruceAPIs_ori <- rbind(cruceAPIs_pibal, cruceAPIs_can)
rm(cruceAPIs_pibal, cruceAPIs_can)

# Eliminar src que no son RN2000
cruceAPIs <- cruceAPIs_ori
cruceAPIs <- cruceAPIs[grepl('^ES', cruceAPIs$src_SITE_CODE), ]
# Eliminar dejar solo nbr que sean apis 
cruceAPIs <- cruceAPIs[grepl('^ext', cruceAPIs$nbr_SITE_CODE), ]

# ___ Se genera la tabla RNProxiAct = RNProxi + crucesAPIs. Con los sig. campos:

# -	src_CODE – Código de la entidad fuente (source)
# -	nbr_CODE – Código de la entidad vecina (neighbour)
# -	AREA – Área común entre entidad fuente y vecina. Cuando este valor es 0 significa que existe una relación de colindancia, cuya longitud común vendrá especificada en el campo siguiente que no podrá tomar valor de 0.
# -	LENGTH – Longitud común entre la entidad fuente y vecina, cuando existe colindancia. Cuando este valor es diferente de 0 el valor del campo AREA no puede tomar otro valor que no sea 0, porque si existe colindancia no puede existir un área común entre las dos entidades.
# - NODE_COUNT - 
# -	Src_Area – Área de la entidad fuente.
# -	Nbr_Area – Área de entidad vecina.
# -	Src_AreaP – Porcentaje de la superficie de la entidad fuente con respecto al área común (AREA)
# -	Nbr_AreaP – Porcentaje de la superficie de la entidad vecina con respecto al área común (AREA)
# -	RelationNbr – Relación de proximidad existente entre la entidad fuente y vecina.

# ___ Obtener dichos campos en crucesAPIs. Los 4 primeros campos están.
# Nbr_Area
cruceAPIs <- merge(cruceAPIs, APIext_ori[ , c("SITE_CODE", "Shape_Area")], 
                   by.x = "nbr_SITE_CODE", by.y = "SITE_CODE", all.x = T)
names(cruceAPIs)[names(cruceAPIs) == "Shape_Area"] <- "NbrArea"

# Src_Area 
cruceAPIs <- merge(cruceAPIs, CNTRYES[ , c("SITE_CODE", "Shape_Area")], 
                   by.x = "src_SITE_CODE", by.y = "SITE_CODE", all.x = T)
names(cruceAPIs)[names(cruceAPIs) == "Shape_Area"] <- "SrcArea"

# Src_AreaP y Nbr_AreaP
cruceAPIs$SrcAreaP <- cruceAPIs$AREA / cruceAPIs$SrcArea * 100
cruceAPIs$NbrAreaP <- cruceAPIs$AREA / cruceAPIs$NbrArea * 100

# RelationNbr OJO: estos condicionales funcionan para este pequeño set de datos.
# Es muy probable que en un cruce para toda Esp con todos los EP, haya alguna situación
# más que aquí no está contemplada. P.ej. AREA y LENGTH == 0 y NODE_COUNT != 0 (aquí no aparece)
cruceAPIs <- cruceAPIs
cruceAPIs$RelationNbr <- ""
cruceAPIs$RelationNbr <- ifelse(cruceAPIs$LENGTH != 0, "COLINDANTES", 
                         ifelse(cruceAPIs$SrcAreaP > 98 & cruceAPIs$NbrAreaP > 98, "COINCIDENTES",
                                ifelse(cruceAPIs$NbrAreaP > 98, "Src CONTIENE A Nbr",
                                       ifelse(cruceAPIs$SrcAreaP > 98, "Src INTEGRADO EN Nbr",
                                              ifelse(cruceAPIs$SrcAreaP < 2 & cruceAPIs$NbrAreaP < 2, "SOLAPE/COLINDANCIA POR ESCALA CARTOGRAFICA",
                                              "SOLAPAN PARCIALMENTE")))))

# Comprobaciones:
# Colindantes: OK; Coincidentes: OK; Scr CONTIENE A Nbr: OK; 
# Scr INTEGRADO EN NBR: OK; "SOLAPE/COLINDANCIA POR ESCALA CARTOGRAFICA": OK;
# "SOLAPAN PARCIALMENTE": OK

# Ver resumen frecuencias de cada indicador
data.frame(table(cruceAPIs$RelationNbr))

# write.table(cruceAPIs, "D:/ww/cruceAPIs.csv", sep = ";", dec = ",", row.names = F)

# ___ Modificar RNProxi para que encaje con los campos anterioes

RNProxi <- RNProxi[ , c("src_SITE_CODE2", "nbr_SITE_CODE2",
                         "AREA", "LENGTH", "NODE_COUNT", 
                         "SrcArea", "NbrArea", 
                         "SrcAreaP", "NbrAreaP", 
                         "RelaNbr")]

names(RNProxi)[names(RNProxi) == "src_SITE_CODE2"] <- "src_SITE_CODE"
names(RNProxi)[names(RNProxi) == "nbr_SITE_CODE2"] <- "nbr_SITE_CODE"
names(RNProxi)[names(RNProxi) == "RelaNbr"] <- "RelationNbr"

# ___ Unir RNProxi con crucesAPIs
# Unión
RNProxiAct <- rbind(RNProxi, cruceAPIs)

#########################
### Tabla RNProxi: Incluir los sitios RN2000 que no solapan/colindan con EPs
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
# En el CNTRYES marcar los que no solapan/colindan en base al vector anterior
CNTRYESmarcado <- CNTRYES
CNTRYESmarcado$Src <- ifelse(CNTRYESmarcado$SITE_CODE %in% sitiosQueCruzan, "SiSrc", "NoSrc")
CNTRYESnoSrc <- CNTRYESmarcado[CNTRYESmarcado$Src == "NoSrc", c("SITE_CODE")]
# crear un df con los sitios RN2000 que no solapan/colindan
RNProxiNoSrc <- data.frame(
  src_SITE_CODE = CNTRYESnoSrc,
  nbr_SITE_CODE = "Sin solape/colindancia con EP",
  RelationNbr = "Sin solape/colindancia con EP"
)

# Unir a la tabla con las relaciones de proximidad
RNProxiAct <- bind_rows(RNProxiAct,RNProxiNoSrc)
RNProxiAct[is.na(RNProxiAct)] <- 0

rm(sitiosQueCruzan, CNTRYESmarcado, CNTRYESnoSrc, RNProxiNoSrc)
#########################
# Crear una tabla (RNProxiActMerge) que tenga la información de las otras dos
# EPAct en RNProxiAct (no haría falta como tal, pero puede ser útil)
#########################
EPAct$SITE_CODE <- gsub("_RN2000", "", EPAct$SITE_CODE)

RNProxiActMerge <- merge(RNProxiAct, EPAct[ , c("SITE_CODE", "Nombre_", "AC", "TIPO")],
                         by.x = "src_SITE_CODE", by.y = "SITE_CODE", all.x = T)

names(RNProxiActMerge)[names(RNProxiActMerge) %in% c("Nombre_", "AC", "TIPO")] <- c("src_Nombre", "src_AC", "src_TIPO")

RNProxiActMerge <- merge(RNProxiActMerge, EPAct[ , c("SITE_CODE", "Nombre_", "AC", "TIPO", "jerar1")],
                         by.x = "nbr_SITE_CODE", by.y = "SITE_CODE", all.x = T)
names(RNProxiActMerge)[names(RNProxiActMerge) %in% c("Nombre_", "AC", "TIPO", "jerar1")] <- c("nbr_Nombre", "nbr_AC", "nbr_TIPO", "nbr_jerar1")

# write.table(RNProxiActMerge, "D:/ww/RNProxiActMerge.csv", sep = ";", dec = ",", row.names = F)

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

#########################
# Instrumentos planificación
#########################
# Planes de los ENP
path_nbr_planes <- file.path(pathEntrada, "ENP_Planes.xlsx")
nbr_planes <- read_excel(path_nbr_planes)
# Comprobar que los SITE_CODE de nbr_planes coinciden con SITE_CODE de EPAct
# all(nbr_planes$SITE_CODE %in% EPAct$SITE_CODE) # Sí. Sale TRUE


# Planes de RN2000
# src_planes <- 
#########################
# Presentación final de las tablas a exportar
#########################

EPAct <- EPAct[ , c("SITE_CODE", "Nombre_", "FIGURA", "DESIGNACIO", "AC", "TIPO", "jerar1")]
names(EPAct)[names(EPAct) == "Nombre_"] <- "Nombre"
EPAct$TIPO <- ifelse(EPAct$TIPO == "A", "A (ZEPA)",
                   ifelse(EPAct$TIPO == "B", "B (LIC/ZEC)", "C (LIC/ZEC y ZEPA)"))

EPAct <- arrange(EPAct, FIGURA, SITE_CODE)


RNProxiAct <- arrange(RNProxiAct, src_SITE_CODE, nbr_SITE_CODE)

RNProxiActMerge <- RNProxiActMerge[ , c("src_SITE_CODE", "nbr_SITE_CODE", "AREA",
                                        "LENGTH", "NODE_COUNT", "SrcArea", "NbrArea",
                                        "SrcAreaP", "NbrAreaP", "RelationNbr",
                                        "src_Nombre", "src_AC", "src_TIPO",
                                        "nbr_Nombre", "nbr_AC", "nbr_TIPO", "nbr_jerar1")]  

RNProxiActMerge <- arrange(RNProxiActMerge, src_SITE_CODE, nbr_SITE_CODE)

#########################
# Crear dos DD: uno para src y otro para nbr
#########################

#____ DD para src
CNTRYES$TIPO_ori <- CNTRYES$TIPO
CNTRYES$TIPO <- ifelse(CNTRYES$TIPO == "A", "A (ZEPA)",
                     ifelse(CNTRYES$TIPO == "B", "B (LIC/ZEC)", "C (LIC/ZEC y ZEPA)"))
DD_src <- CNTRYES
#____ DD para nbr
# sacar todos los códigos nbr que han cruzado
nbrProxi <- unique(RNProxiAct$nbr_SITE_CODE)
# Usar el cector anterior para marcar los que estén presentes en la tabla de EPAct
DD_nbr <- EPAct
DD_nbr$nbr <- ifelse(DD_nbr$SITE_CODE %in% nbrProxi, "SiNbr", "NoNbr")
DD_nbr <- DD_nbr[DD_nbr$nbr == "SiNbr", ] # Seleccionar solo los que interesan
DD_nbr <- subset(DD_nbr, select = -c(nbr))
DD_nbr[nrow(DD_nbr) + 1,] <- c("Sin solape/colindancia con EP", "", "", "", "", "","")
rm(nbrProxi)


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
