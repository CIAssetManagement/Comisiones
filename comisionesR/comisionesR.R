library(lubridate)
library(plyr)
library(dplyr)
library(tidyr)
library(xts)
library(xlsx) # requires java
# options(stringsAsFactors = FALSE)

# A -----------  CALCULAR LOS SALDOS PROMEDIO DE OASIS -------------------
saldo <- read.csv("tablas/posicionMes.csv", colClasses=c("character", "character", "numeric"), stringsAsFactors = FALSE)
saldo$fecha <- ymd(saldo$fecha)
diasEnMes <- days_in_month(saldo$fecha[1])
# head(saldosOasys)
# diasEnMes <- 31

NAorMin <- function(date) if(length(date)>0) {return(min(date))} else {return(NA)}
NAorMax <- function(date) if(length(date)>0) {return(max(date))} else {return(NA)}


saldoPromedio <- data.frame(contrato=sort(unique(saldo$contrato)))
saldoWide <- spread(saldo, contrato, posicion)
saldoWide <- xts(saldoWide[ ,-1], saldoWide[ ,1])
# CREATE FULL MATRIX CON SALDOS EN TODOS LOS DIAS INCLUIDO SABADOS Y DOMINGOS
firstOfMonth <- min(saldo$fecha)
maxDate <- max(saldo$fecha)
diffDate <- as.numeric(maxDate-firstOfMonth)
auxGrid <- matrix(NA, nrow=diffDate+1, ncol=ncol(saldoWide), dimnames=list(as.character(firstOfMonth + 0:diffDate), names(saldoWide)))
for(i in 1:length(index(saldoWide))) auxGrid[as.character(index(saldoWide)[i]), ] <- as.matrix(saldoWide[index(saldoWide)[i]])
for(j in 1:ncol(auxGrid)) for(i in 1:(nrow(auxGrid)-1)) if(!is.na(auxGrid[i,j])) if(is.na(auxGrid[i+1, j])) auxGrid[i+1, j] <- auxGrid[i, j]
auxGrid[is.na(auxGrid)] <- 0
promedio <- apply(auxGrid, 2, mean)

saldoWide <- xts(auxGrid, firstOfMonth + 0:diffDate)
saldoWideTabla <- data.frame(saldoWide, row.names=as.character(index(saldoWide)), check.names=FALSE)
write.csv(saldoWideTabla, "tablas/saldoDiario.csv", row.names = FALSE)
# saldosLong <- gather(data.frame(fecha=index(saldoWide), saldoWide, check.names=FALSE), fecha, contrato)

saldoPromedio$saldoPromedio <- 0
saldoPromedio$saldoPromedio[match(names(promedio), saldoPromedio$contrato)] <- promedio


# B ------------------ HACER UN MATCH CON LA CARTERA MODELO DE AM -------------------
configClientes <- read.csv("tablas/carterasModeloClientes.txt", header=FALSE, stringsAsFactors = FALSE)
names(configClientes) <- c("contrato", "carteramodelo", "RFC")
configClientes$contrato <- gsub("[^A-Za-z0-9]", "", configClientes$contrato)
# head(configClientes)
idx <- match(saldoPromedio$contrato, configClientes$contrato)
saldoPromedio$carteramodelo[!is.na(idx)] <- configClientes$carteramodelo[idx[!is.na(idx)]]
# head(saldosPromedio)


# C ------------------- CALCULAR SALDOS GRUPALES ------------------------------------
saldoPromedio$saldoPromedioGrupal <- saldoPromedio$saldoPromedio
grupo <- read.xlsx("configTables.xlsx", sheetName = "grupos", stringsAsFactors = FALSE)
uniqueGrupo <- unique(grupo$Grupo)
for(i in 1:length(uniqueGrupo)){
  sub <- grupo[grupo$Grupo==uniqueGrupo[i], ]
  uniqueContrato <- unique(sub$Titular)
  suma <- sum(saldoPromedio$saldoPromedio[saldoPromedio$contrato %in% uniqueContrato])
  saldoPromedio$saldoPromedioGrupal[saldoPromedio$contrato %in% uniqueContrato] <- suma
}
# head(saldoPromedio)


# D ------------------  CALCULAR COMISION CORRESPONDIENTE ---------------------------------
tablaComisiones <- read.xlsx("configTables.xlsx", sheetName = "tablaComisiones", stringsAsFactors = FALSE, check.names = FALSE)
uniqueCM <- unique(saldoPromedio$carteramodelo)
missingCM <- uniqueCM[!(uniqueCM %in% names(tablaComisiones))]
if(length(missingCM)>0) warning(sprintf("Las carteras modelo %s no están en la tabla de comisiones.", paste(missingCM), collapse=", "))
saldoPromedio$comisionAnual <- 0
for(i in 1:nrow(saldoPromedio)){
  monto <- saldoPromedio$saldoPromedioGrupal[i]
  count <- 1
  while(monto >= tablaComisiones$limiteInferior[count+1] & (count+1)<=nrow(tablaComisiones)) count <- count + 1
  saldoPromedio$comisionAnual[i] <- as.numeric(tablaComisiones[count, saldoPromedio$carteramodelo[i]])
}
saldoPromedio$comision <- saldoPromedio$comisionAnual/360*diasEnMes
# saldoPromedio$comision <- saldoPromedio$comisionAnual/360*30

# E ------------------- COMISIONES ESPECIALES (e.g., IMSS) ------------------------------------
especiales <- read.xlsx("configTables.xlsx", sheetName = "comisionesEspeciales", stringsAsFactors = FALSE, check.names = FALSE)
names(especiales) <- c("contrato", "comision")
for(i in 1:nrow(especiales)) saldoPromedio$comision[saldoPromedio$contrato==especiales$contrato[i]] <- especiales$comision[i]/360*diasEnMes

# F --------------------  REALIZAR DESCUENTOS EMPLEADOS -----------------------------------------
descuentos <- read.xlsx("configTables.xlsx", sheetName = "descuentos", stringsAsFactors = FALSE, check.names = FALSE)
names(descuentos) <- c("Titular", "Descuento")
carterasTestigo <- read.xlsx("configTables.xlsx", sheetName = "carterasTestigo", stringsAsFactors = FALSE, check.names = FALSE)
names(carterasTestigo) <- c("contrato")
descuentos <- rbind(descuentos, data.frame(Titular=carterasTestigo$contrato, Descuento=1))
saldoPromedio$descuento <- 0
for(i in 1:nrow(descuentos)) saldoPromedio$descuento[saldoPromedio$contrato==descuentos$Titular[i]] <- descuentos$Descuento[i]
saldoPromedio$comisionConDescuento <- saldoPromedio$comision*(1-saldoPromedio$descuento)

# G -----------------------  CALCULA MONTO POR COMISION CON Y SIN DESCUENTO ----------------------------
saldoPromedio$monto <- saldoPromedio$comision*saldoPromedio$saldoPromedio
saldoPromedio$monto <- saldoPromedio$comision*saldoPromedio$saldoPromedio
saldoPromedio$monto[saldoPromedio$monto<0] <- 0
saldoPromedio$montoConDescuento <- saldoPromedio$comisionConDescuento*saldoPromedio$saldoPromedio
saldoPromedio$montoConDescuento[saldoPromedio$montoConDescuento<0] <-  0

# H -------------------------------    IVA  ------------------------------------------------------------------
saldoPromedio$montoIVA <- saldoPromedio$montoConDescuento*1.16
# View(saldoPromedio)

# I -------------------------   BONIFICACIONES/ CARGOS MANUALES -------------------------------------------
bonificaciones <- read.xlsx("configTables.xlsx", sheetName = "bonificaciones", stringsAsFactors = FALSE, check.names = FALSE)
names(bonificaciones) <- c("contrato", "bonificacion")
saldoPromedio$bonificacion <- 0
saldoPromedio$montoNetoIVA <- saldoPromedio$montoIVA
if(nrow(bonificaciones) > 0){
  for(i in 1:nrow(bonificaciones)) {
    saldoPromedio$bonificacion[saldoPromedio$contrato==bonificaciones$contrato[i]] <- bonificaciones$bonificacion[i]
    saldoPromedio$montoNetoIVA[saldoPromedio$contrato==bonificaciones$contrato[i]] <- saldoPromedio$montoNetoIVA[saldoPromedio$contrato==bonificaciones$contrato[i]] - bonificaciones$bonificacion[i]
  }  
}


# J -----------------------  CONTRATOS SIN SALDO SUFICIENTE EN CERO  ------------------------------------------
saldoPromedio$cobroFinal <- saldoPromedio$montoNetoIVA
saldoLast <- saldo[saldo$fecha==maxDate, ]
for(i in 1:nrow(saldoPromedio)){
  lastValue <- saldoLast$posicion[saldoLast$contrato==saldoPromedio$contrato[i]]
  saldoPromedio$cobroFinal[i] <- max(min(lastValue, saldoPromedio$montoNetoIVA[i]), 0)
}

# K ---------------------------------- GUARDAR RESULTADOS -------------------------------------------------------
write.csv(saldoPromedio, "tablas/comisionesCompleto.csv")
# write.csv(saldoPromedio[ c("contrato", "saldoPromedioGrupal", "montoNetoIVA")], "tablas/comisionesResumen.csv")

# L ------------  GUARDAR RESUMEN HISTORICO MES ---------------------------------
resumenCobro <- t(data.frame(
  "dias en mes" = diasEnMes,
  "suma sin descuentos" = sum(saldoPromedio$monto),
  "suma neta sin IVA" = sum(saldoPromedio$montoNetoIVA*(1/1.16)),
  "suma neta con IVA" = sum(saldoPromedio$montoNetoIVA),
  "suma neta con IVA quitando fondos insuficientes" = sum(saldoPromedio$cobroFinal),
  check.names = FALSE
))


# K -------------- JUNTAR EN UN EXCEL -----------------
write.xlsx2(saldoPromedio[ ,c("contrato", "cobroFinal")], file=sprintf("../reporteDiario/comisiones%s.xlsx", format(maxDate, "%y%m%d")), sheet="cargos CISM", row.names=FALSE)
write.xlsx2(resumenCobro, file=sprintf("../reporteDiario/comisiones%s.xlsx", format(maxDate, "%y%m%d")), sheet="resumenCobro", append=TRUE, col.names = FALSE)
write.xlsx2(saldoPromedio, file=sprintf("../reporteDiario/comisiones%s.xlsx", format(maxDate, "%y%m%d")), sheet="completoCobro", append=TRUE, row.names=FALSE)
write.xlsx2(saldoWideTabla, file=sprintf("../reporteDiario/comisiones%s.xlsx", format(maxDate, "%y%m%d")), sheet="saldoDiarioContratosMes", append=TRUE)
write.xlsx2(grupo, file=sprintf("../reporteDiario/comisiones%s.xlsx", format(maxDate, "%y%m%d")), sheet="Grupos", append=TRUE, row.names=FALSE)
write.xlsx2(descuentos, file=sprintf("../reporteDiario/comisiones%s.xlsx", format(maxDate, "%y%m%d")), sheet="DescuentosEmpleados", append=TRUE, row.names=FALSE)
write.xlsx2(bonificaciones, file=sprintf("../reporteDiario/comisiones%s.xlsx", format(maxDate, "%y%m%d")), sheet="BonificacionesCargos", append=TRUE, row.names=FALSE)

