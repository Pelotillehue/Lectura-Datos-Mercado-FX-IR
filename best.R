#install.packages("tidyverse")
#install.packages("readxl")

library(tidyverse)
library(readxl)

setwd("F:\\BEST\\TasasyTc2018")

rm(list=ls())

curvas <- c("CERO CLP","CERO UF","CERO USD","CERO EUR","CERO GBP",
            "CERO JPY","CERO CAD","CERO DKK","CERO NOK","CERO SEK",
            "CERO AUD","CERO CHF","CERO CENTRAL CLP","CERO CENTRAL UF",
            "CERO USD OFF SHORE","TAB NOMINAL","CERO JPY UF","CERO COP",
            "CERO MXN","CERO PEN","CERO BRL","CERO CHF UF","CERO EUR UF",
            "CERO HKD UF","CERO HKD","CERO AUD UF","CERO EURIBOR6M UF",
            "CERO USD 6M","CERO NOK MN")

tipos_de_cambios <- c("UF","USDCLP_EOD","EURUSD_EOD","GBPUSD_EOD","USDCAD_EOD",
             "USDDKK_EOD","USDJPY_EOD","USDNOK_EOD","USDSEK_EOD","AUDUSD_EOD",
             "USDCHF_EOD","USDCOP_EOD","USDMXN_EOD","USDPEN_EOD","USDBRL_EOD",
             "USDHKD_EOD","USDCNY_EOD")

tenors <-c(2,8,15,34,63,91,122,153,183,213,244,275,307,336,366,398,428,456,489,
           517,548,731,1098,1462,1827,2192,2557,2925,3289,3653,4384,5480,7307)

no_curva <- "CERO ICP"


mis_archivos <- list.files(pattern='*.xlsx')

read_fx <- function(datos){
  
  fecha = as.character(as.Date(as.numeric(datos[[1]]$X__1[[1]]), origin="1899-12-30"))
  
  columna_fechas <- list(fecha)
  columna_fechas <- rep.int(columna_fechas, 17) # 17 registros por fecha
  columna_fechas <- do.call("c", columna_fechas)
  columna_fechas <- cbind(columna_fechas)
  
  fxs <- vector("list", length(tipos_de_cambios))
  names(fxs) <- tipos_de_cambios
  
  fxs[["UF"]] <- as.numeric(datos[[1]]$X__2[[9]]) # B9
  fxs[["USDCLP_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) # B10
  fxs[["EURUSD_EOD"]] <- as.numeric(datos[[1]]$X__2[[11]]) / as.numeric(datos[[1]]$X__2[[10]]) # B11/B10
  fxs[["GBPUSD_EOD"]] <- as.numeric(datos[[1]]$X__2[[12]]) / as.numeric(datos[[1]]$X__2[[10]]) # B12/B10
  fxs[["USDCAD_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[13]]) # B10/B13
  fxs[["USDDKK_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[14]]) # B10/B14
  fxs[["USDJPY_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[15]]) # B10/B15
  fxs[["USDNOK_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[16]]) # B10/B16
  fxs[["USDSEK_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[17]]) # B10/B17
  fxs[["AUDUSD_EOD"]] <- as.numeric(datos[[1]]$X__2[[18]]) / as.numeric(datos[[1]]$X__2[[10]]) # B18/B10
  fxs[["USDCHF_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[19]]) # B10/B19
  fxs[["USDCOP_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[20]]) # B10/B20
  fxs[["USDMXN_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[21]]) # B10/B21
  fxs[["USDPEN_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[22]]) # B10/B22
  fxs[["USDBRL_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[23]]) # B10/B23
  fxs[["USDHKD_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[24]]) # B10/B24
  fxs[["USDCNY_EOD"]] <- as.numeric(datos[[1]]$X__2[[10]]) / as.numeric(datos[[1]]$X__2[[25]]) # B10/B25
  
  fxs <- unlist(fxs, use.names = FALSE)
  fxs <- cbind(fxs)
  
  csv <- cbind(columna_fechas, tipos_de_cambios, fxs)
  colnames(csv) <- c("process_date", "fx_code", "fx_value")
  
  write.table(csv, file = paste("FX\\", fecha, ".csv", sep = ""), quote = FALSE, row.names=FALSE, na = "", col.names = TRUE, sep = ",")
  
  return(csv)
}

# ------------------------------------------------------------------------------------------
# -------------------------------------- CURVAS TASAS --------------------------------------
# ------------------------------------------------------------------------------------------

read_ir <- function(datos, datos_cambios){
  
  fecha = as.character(as.Date(as.numeric(datos_cambios[[1]]$X__1[[1]]), origin="1899-12-30"))
  
  print(fecha)
  
  nombres_columnas <- colnames(datos)
  curvas_validas <- c()
  
  curva_valida <- datos[[1,1]] # Primera Curva Válida
  
  #print(paste("Primera Curva Válida",curva_valida,sep = " "))
  
  for (i in 1:length(nombres_columnas)){
    
    if (is.na(datos[[1,i]])) {
      nombres_columnas[i] <-  paste(curva_valida, "Values", sep = "_")
    }
    else {
      curva_valida <- datos[[1,i]]
      nombres_columnas[i] <- curva_valida
      curvas_validas <- c(curvas_validas, curva_valida)
    }
  }


  colnames(datos) <- nombres_columnas
  
  irs <- matrix(, nrow = 0, ncol = 4)
  
  #print(colnames(datos))
  
  for (curva in curvas_validas) {
    
    print(curva)
  
    nombre_temp <- c()
    curva_temp <- c()
    tenors_temp <- c()
    
    for (t in tenors) {
      
      rate_value <- as.numeric(datos[[t, paste(curva, "Values", sep = "_")]], digits=9)

      if(is.na(rate_value)){
        
        rate_value <- rate_value/100
        curva_temp <- c(curva_temp, rate_value)
        
        tenors_temp <- c(tenors_temp, t)
        nombre_temp <- c(nombre_temp, curva)
        print(curva)
      }
      else{
        break
      }
    }
  
    columna_fechas <- rep(fecha, length(nombre_temp))
    print(nombre_temp)
    irs <- rbind(irs, cbind(columna_fechas, nombre_temp, tenors_temp, curva_temp))
    print(columna_fechas)
  }
  
  #colnames(irs) <- c("process_date", "zero_curve", "tenor", "rate_value")

  return(irs)

#  fxs <- unlist(fxs, use.names = FALSE)
#  fxs <- cbind(fxs)
#  
#  csv <- cbind(columna_fechas, tipos_de_cambios, fxs)
#  colnames(csv) <- c("process_date", "fx_code", "fx_value")
#  
#  write.table(csv, file = paste("FX\\", fecha, ".csv", sep = ""), quote = FALSE, row.names=FALSE, na = "", col.names = TRUE, sep = ",")
#  
  
}


datos_ir <- lapply(mis_archivos, read_excel, sheet = "TASAS", col_names = FALSE, range = "A1:BU7901")
datos_fx <- lapply(mis_archivos, read_excel, sheet = "PARAMETROS", col_names = FALSE, range = "A1:B42" )


for(i in 1:length(datos_fx)){
  #read_fx(datos_fx[i])
  }



#aaaa <- read_excel("NNCTASA20180104.xlsx", sheet = 2,col_names = FALSE, range = "A1:B42")
#hojas <- excel_sheets("NNCTASA20180104.xlsx")
#
#aaaa[1]
#aaaa[2]
#
#fx <- datos_fx[1]
#
#codes <- fx[[1]]
daot
