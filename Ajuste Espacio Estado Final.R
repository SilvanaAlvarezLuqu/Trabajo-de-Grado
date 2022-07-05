# DIRECTORIO DE TRABAJO
#------------------------
setwd("C:/Users/ALVARLX23/OneDrive - Abbott/Documents/Proyecto")

# PAQUETES  (solo debe hacerse 1 vez por equipo)
#-----------
install.packages("TSstudio"); install.packages("fUnitRoots")
install.packages("readxl"); install.packages("dplyr") #  %>%
install.packages("dlm"); install.packages("openxlsx")

# LIBRERIAS   (Debe correrse cada vez que se use el script)
#------------
library(TSstudio); library(fUnitRoots)
library(readxl); library(dplyr) #  %>%
library(dlm); library(openxlsx)

#--------------------------------------------------------------------
#                            INSTITUCIONES 
#--------------------------------------------------------------------

#     CARGAR DATOS 
#----------------------
BI_INS<-read_excel("Base_inicial.xlsx", sheet = "SKU INS", col_names = TRUE)
str(BI_INS)

BI_INS<-as.data.frame(BI_INS)
BI_INS$FECHA<-as.Date.character(BI_INS$FECHA)


# ESTIMACIÓN DE HIPERPARÁMETROS (S Y T)
#----------------------------------------
parametros_ins <- matrix(NA, nrow = dim(BI_INS)[2], ncol = 3)
parametros_ins[,1] <- colnames(BI_INS)

# ESTACIONALIDAD
#-----------------
max<-NULL; ciclo_ins<-NULL
x11()
par(mfrow=c(3,3))
for (i in 1:dim(BI_INS)[2] ) {
  per<-spectrum(BI_INS[,i], spans=c(3,3), log="no")
  max[i]<-which.max(per$spec)
  ciclo_ins[i]<-1/(max[i]*1/12)
}
parametros_ins[,2] <- round(ciclo_ins,0)


#   TENDENCIA
#-----------------
raices_ins<-list()

for (i in 1:dim(BI_INS)[2]){
  raices_ins[[i]] <- fUnitRoots::adfTest(ts(BI_INS[,i], start = c(2016,1), frequency = 12)) 
}

tendencia_ins <- c(NA, rep(1,5), 0, rep(1,2), 0, rep(1,3), 0, 1,
                   0, 1, 0, rep(1,5), 0, rep(1,2), rep(0,2), rep(1,12),
                   0, rep(1,4), 0, rep(1,3), 0, rep(1,5), 0, rep(1,6),
                   0, rep(1,3), rep(0,2),1, 0, 1, rep(0, 3), rep(1,2),
                   rep(0,2), 1, 0, 1, rep(0,3), rep(1,2))

parametros_ins[,3] <- tendencia_ins

# Adecuación tabla parámetros
#-------------------------------

parametros_ins_df = as.data.frame(parametros_ins)
colnames(parametros_ins_df)<-c("sku", "s", "t")
parametros_ins_df<-parametros_ins_df[-1,] # QUITAR FECHA

entrenamiento<-BI_INS[1:round(dim(BI_INS)[1]*0.9),] 
prueba<-BI_INS[(round(dim(BI_INS)[1]*0.9)+1):dim(BI_INS)[1],]


#    DIVISIÓN X CICLO
#------------------------

anual = parametros_ins_df %>% filter(s==12)

semestral = parametros_ins_df %>% filter(s==6)

cuatrimestre = parametros_ins_df %>% filter(s==4)

trimestre = parametros_ins_df %>% filter(s==3)

bimestre = parametros_ins_df %>% filter(s==2)

SinCiclo = parametros_ins_df %>% filter(s==1 | s==0)


#----------------------------
#   SERIES CON CICLO ANUAL
#----------------------------
anual_ts = entrenamiento %>%
  select(anual$sku)
anual_ts<-ts(anual_ts, start = c(2016,1), frequency = 12)

ts_plot(anual_ts)

#   AJUSTE MLE
#----------------
# INICIALIZAR
fit112 <- dlmMLE(anual_ts[,1], c(rep(1,12)), build = fn1_12, hessian = T)
(modelo_anual <- fn1_12(fit112$par))

# ESTIMACIÓN PARÁMETROS
for (i in 2:dim(anual)[1]) {
  fit112 <- dlmMLE(anual_ts[,i], c(rep(1,12)), build = fn1_12, hessian = T,method = "Nelder-Mead" )
  mod <- fn1_12(fit112$par)
  
  modelo_anual = modelo_anual %+% mod
}


# FILTRO Y FORECAST
#--------------------
filtro <- dlmFilter(anual_ts, modelo_anual)

forecast <- dlmForecast(filtro, nAhead = dim(prueba)[1] + 28, sampleNew = 1)
(forecast_anual<-forecast$f)
colnames(forecast_anual)<-anual$sku


# RESIDUALES
#------------
anual_ts_prueba = prueba %>% select(anual$sku)

(res_anual_I<- anual_ts_prueba - forecast_anual[1:dim(prueba)[1],])
mape_anual_I<-colMeans(abs(res_anual_I/anual_ts_prueba))
rmse_anual_I<-sqrt(colMeans(res_anual_I^2))

#--------------------------------
#   SERIES CON CICLO SEMESTRAL
#--------------------------------
semestre_ts = entrenamiento %>%
  select(semestral$sku)
semestre_ts<-ts(semestre_ts, start = c(2016,1), frequency = 12)

ts_plot(semestre_ts)

#   AJUSTE MLE
#----------------

# INICIALIZAR
fit16 <- dlmMLE(semestre_ts[,1], c(rep(1,2)), build = fn1_2, hessian = T)
(modelo_semestre <- fn1_2(fit16$par))


# ESTIMACIÓN PARÁMETROS
for (i in 2:6) {
  fit16 <- dlmMLE(semestre_ts[,i], c(rep(1,6)), build = fn1_6, hessian = T,method = "Nelder-Mead" )
  mod <- fn1_6(fit16$par)
  
  modelo_semestre = modelo_semestre %+% mod
}

fit16 <- dlmMLE(semestre_ts[,7], c(rep(1,2)), build = fn1_2, hessian = T,method = "Nelder-Mead" )
mod <- fn1_2(fit16$par)
modelo_semestre = modelo_semestre %+% mod

fit16 <- dlmMLE(semestre_ts[,8], c(rep(1,6)), build = fn1_6, hessian = T,method = "Nelder-Mead" )
mod <- fn1_6(fit16$par)
modelo_semestre = modelo_semestre %+% mod

# FILTRO Y FORECAST
#-------------------
filtro <- dlmFilter(semestre_ts, modelo_semestre)

forecast <- dlmForecast(filtro, nAhead = dim(prueba)[1] + 28, sampleNew = 1)
(forecast_semestral<-forecast$f)
colnames(forecast_semestral)<-semestral$sku


# RESIDUALES
#------------
semestral_ts_prueba = prueba %>% select(semestral$sku)

(res_semestral_I<- semestral_ts_prueba - forecast_semestral[1:dim(prueba)[1],])
mape_semestral_I<-colMeans(abs(res_semestral_I/semestral_ts_prueba))
rmse_semestral_I<-sqrt(colMeans(res_semestral_I^2))


#----------------------------------------
#     SERIES CON CICLO CUATRIMESTRAL
#----------------------------------------
cuatrimestre_ts = entrenamiento %>%
  select(cuatrimestre$sku)
(cuatrimestre_ts<-ts(cuatrimestre_ts, start = c(2016,1), frequency = 12))

ts_plot(cuatrimestre_ts)

#   AJUSTE MLE
#----------------

# INICIALIZAR
fit14 <- dlmMLE(cuatrimestre_ts[,1], c(rep(1,4)), build = fn1_4, hessian = T)
modelo_cuatrimestre <- fn1_4(fit14$par)

# ESTIMACIÓN PARÁMETROS
for (i in 2:dim(cuatrimestre)[1]) {
  fit14 <- dlmMLE(cuatrimestre_ts[,i], c(rep(1,4)), build = fn1_4, hessian = T,method = "Nelder-Mead" )
  mod <- fn1_4(fit14$par)
  
  modelo_cuatrimestre = modelo_cuatrimestre %+% mod
}

# FILTRO Y FORECAST
#-------------------
filtro <- dlmFilter(cuatrimestre_ts, modelo_cuatrimestre)

forecast <- dlmForecast(filtro, nAhead = dim(prueba)[1] + 28, sampleNew = 1)
(forecast_cuatrimestre<-forecast$f)
colnames(forecast_cuatrimestre)<-cuatrimestre$sku

# RESIDUALES
#-------------
cuatrimestre_ts_prueba = prueba %>% select(cuatrimestre$sku)

(res_cuatrimestre_I<- cuatrimestre_ts_prueba - forecast_cuatrimestre[1:dim(prueba)[1],])
mape_cuatrimestre_I<-colMeans(abs(res_cuatrimestre_I/cuatrimestre_ts_prueba))
rmse_cuatrimestre_I<-sqrt(colMeans(res_cuatrimestre_I^2))


#----------------------------------------
#     SERIES CON CICLO TRIMESTRAL
#----------------------------------------
trimestre_ts = entrenamiento %>%
  select(trimestre$sku)
(trimestre_ts<-ts(trimestre_ts, start = c(2016,1), frequency = 12))

ts_plot(trimestre_ts)

#   AJUSTE MLE
#----------------

# INICIALIZAR
fit13 <- dlmMLE(trimestre_ts[,1], c(rep(1,12)), build = fn1_12, hessian = T)
modelo_trimestre <- fn1_12(fit13$par)

# ESTIMACIÓN PARÁMETROS
for (i in 2:dim(trimestre)[1]) {
  fit13 <- dlmMLE(trimestre_ts[,i], c(rep(1,12)), build = fn1_12, hessian = T, method = "Nelder-Mead" ) # 
  mod <- fn1_12(fit13$par)
  
  modelo_trimestre = modelo_trimestre %+% mod
}

# FILTRO Y FORECAST
#-------------------
filtro <- dlmFilter(trimestre_ts, modelo_trimestre)

forecast <- dlmForecast(filtro, nAhead = dim(prueba)[1] + 28, sampleNew = 1)
(forecast_trimestre<-forecast$f)
colnames(forecast_trimestre)<-trimestre$sku

# RESIDUALES
#-------------
trimestre_ts_prueba = prueba %>% select(trimestre$sku)

(res_trimestre_I<- trimestre_ts_prueba - forecast_trimestre[1:dim(prueba)[1],])
mape_trimestre_I<-colMeans(abs(res_trimestre_I/trimestre_ts_prueba))
rmse_trimestre_I<-sqrt(colMeans(res_trimestre_I^2))


#----------------------------------------
#     SERIES CON CICLO BIMESTRAL
#----------------------------------------

bimestre1_ts = entrenamiento %>% select(c("ENSURE PLUS HN RPB VAINI", "PEDIALYTE 45 MEQ MANZANA 500 ML ZINC", "ENSURE ADV RPB Vainilla", "PERATIVE 8oz" ))
bimestre2_ts = entrenamiento %>% select(c("ENSURE BASE 400G Vainilla", "SIMILAC SPC 30 KCAL 2oz x 48 fcos", "SIMILAC INFANT NIPLE AND RING", "OSMOLITE LPC 1.5L",
                                          "EQUIPO FREEGO KIT BOLSA 0.5L", "ENSURE ADV RPB FRESA"))

(bimestre1_ts<-ts(bimestre1_ts, start = c(2016,1), frequency = 12))
(bimestre2_ts<-ts(bimestre2_ts, start = c(2016,1), frequency = 12))

ts_plot(cbind(bimestre1_ts, bimestre2_ts))

#   AJUSTE MLE
#----------------

# INICIALIZAR
fit12 <- dlmMLE(bimestre1_ts[,1], c(rep(1,4)), build = fn1_4, hessian = T)
modelo_bimestre1 <- fn1_4(fit12$par)

# ESTIMACIÓN PARÁMETROS
for (i in 2:dim(bimestre1_ts)[2]) {
  fit12 <- dlmMLE(bimestre1_ts[,i], c(rep(1,4)), build = fn1_4, hessian = T,method = "Nelder-Mead" )
  mod <- fn1_4(fit12$par)
  
  modelo_bimestre1 = modelo_bimestre1 %+% mod
}

# FILTRO Y FORECAST
#-------------------
filtro <- dlmFilter(bimestre1_ts, modelo_bimestre1)

forecast <- dlmForecast(filtro, nAhead = dim(prueba)[1] + 28, sampleNew = 1)
(forecast_bimestre1<-forecast$f)
colnames(forecast_bimestre1)<-c("ENSURE PLUS HN RPB VAINI", "PEDIALYTE 45 MEQ MANZANA 500 ML ZINC", "ENSURE ADV RPB Vainilla", "PERATIVE 8oz" )


# RESIDUALES
#-------------
bimestre1_ts_prueba = prueba %>% select(c("ENSURE PLUS HN RPB VAINI", "PEDIALYTE 45 MEQ MANZANA 500 ML ZINC", "ENSURE ADV RPB Vainilla", "PERATIVE 8oz" ))

(res_bimestre1_I<- bimestre1_ts_prueba - forecast_bimestre1[1:dim(prueba)[1],])
mape_bimestre1_I<-colMeans(abs(res_bimestre1_I/bimestre1_ts_prueba))
rmse_bimestre1_I<-sqrt(colMeans(res_bimestre1_I^2))


#   AJUSTE MLE
#----------------

# INICIALIZAR
fit12 <- dlmMLE(bimestre2_ts[,1], c(rep(1,2)), build = fn1_2, hessian = T)
modelo_bimestre2 <- fn1_2(fit12$par)

# ESTIMACIÓN PARÁMETROS
for (i in 2:dim(bimestre2_ts)[2]) {
  fit12 <- dlmMLE(bimestre2_ts[,i], c(rep(1,2)), build = fn1_2, hessian = T,method = "Nelder-Mead" )
  mod <- fn1_2(fit12$par)
  
  modelo_bimestre2 = modelo_bimestre2 %+% mod
}

# FILTRO Y FORECAST
#-------------------
filtro <- dlmFilter(bimestre2_ts, modelo_bimestre2)

forecast <- dlmForecast(filtro, nAhead = dim(prueba)[1] + 28, sampleNew = 1)
(forecast_bimestre2<-forecast$f)
colnames(forecast_bimestre2)<-c("ENSURE BASE 400G Vainilla", "SIMILAC SPC 30 KCAL 2oz x 48 fcos", "SIMILAC INFANT NIPLE AND RING", "OSMOLITE LPC 1.5L",
                                "EQUIPO FREEGO KIT BOLSA 0.5L", "ENSURE ADV RPB FRESA")


# RESIDUALES
#-------------
bimestre2_ts_prueba = prueba %>% select(c("ENSURE BASE 400G Vainilla", "SIMILAC SPC 30 KCAL 2oz x 48 fcos", "SIMILAC INFANT NIPLE AND RING", "OSMOLITE LPC 1.5L",
                                          "EQUIPO FREEGO KIT BOLSA 0.5L", "ENSURE ADV RPB FRESA"))

(res_bimestre2_I<- bimestre2_ts_prueba - forecast_bimestre2[1:dim(prueba)[1],])
mape_bimestre2_I<-colMeans(abs(res_bimestre2_I/bimestre2_ts_prueba))
rmse_bimestre2_I<-sqrt(colMeans(res_bimestre2_I^2))


#--------------------------------
#        SERIES SIN CICLO 
#--------------------------------

SinCiclo_ts = entrenamiento %>%
  select(SinCiclo$sku)
(SinCiclo_ts<-ts(SinCiclo_ts, start = c(2016,1), frequency = 12))

ts_plot(SinCiclo_ts)

#   AJUSTE MLE
#----------------

# INICIALIZAR
fit10 <- dlmMLE(SinCiclo_ts[,1], c(rep(1,2)), build = fn1_0, hessian = T)
modelo_sinciclo <- fn1_0(fit10$par)

# ESTIMACIÓN PARÁMETROS
for (i in 2:dim(SinCiclo)[1]) {
  fit10 <- dlmMLE(SinCiclo_ts[,i], c(rep(1,2)), build = fn1_0, hessian = T,method = "Nelder-Mead" )
  mod <- fn1_0(fit10$par)
  
  modelo_sinciclo = modelo_sinciclo %+% mod
}

# FILTRO Y FORECAST
#-------------------
filtro <- dlmFilter(SinCiclo_ts, modelo_sinciclo)

forecast <- dlmForecast(filtro, nAhead = dim(prueba)[1] + 28, sampleNew = 1)
(forecast_sinciclo<-forecast$f)
colnames(forecast_sinciclo)<-SinCiclo$sku

# RESIDUALES
#-------------
sinciclo_ts_prueba = prueba %>% select(SinCiclo$sku)

(res_sinciclo_I<- sinciclo_ts_prueba - forecast_sinciclo[1:dim(prueba)[1],])
mape_sinciclo_I<-colMeans(abs(res_sinciclo_I/sinciclo_ts_prueba))
rmse_sinciclo_I<-sqrt(colMeans(res_sinciclo_I^2))


# ADECUACIÓN EXPORTACIÓN EXCEL
#------------------------------

# MÉTRICAS
mape_ins<-c(mape_anual_I, mape_semestral_I, mape_cuatrimestre_I, mape_trimestre_I, mape_bimestre1_I, mape_bimestre2_I, mape_sinciclo_I)
rmse_ins<-c(rmse_anual_I, rmse_semestral_I, rmse_cuatrimestre_I, rmse_trimestre_I, rmse_bimestre1_I, rmse_bimestre2_I, rmse_sinciclo_I)
sku_ins<-c(anual$sku, semestral$sku, cuatrimestre$sku, trimestre$sku, colnames(bimestre1_ts), colnames(bimestre2_ts), SinCiclo$sku)

metricas_ins<-cbind(sku=sku_ins,mape=mape_ins, rmse=rmse_ins)

# FORECAST 28 meses
s<-as.Date("2022-03-01") # actualizar a 1 mes despúes del último mes que se tiene
Fecha<-as.Date(seq(from=s, by="month", length.out=28))

forecast_ins<-rbind(Fecha,t(forecast_anual[(dim(prueba)[1]+1):dim(forecast_anual)[1],]),t(forecast_semestral[(dim(prueba)[1]+1):dim(forecast_semestral)[1],]), 
                    t(forecast_cuatrimestre[(dim(prueba)[1]+1):dim(forecast_cuatrimestre)[1],]), t(forecast_trimestre[(dim(prueba)[1]+1):dim(forecast_trimestre)[1],]),
                    t(forecast_bimestre1[(dim(prueba)[1]+1):dim(forecast_bimestre1)[1],]), t(forecast_bimestre2[(dim(prueba)[1]+1):dim(forecast_bimestre2)[1],]), 
                    t(forecast_sinciclo[(dim(prueba)[1]+1):dim(forecast_sinciclo)[1],]))

colnames(forecast_ins)<-Fecha
row.names(forecast_ins)[2:86]<-sku_ins

# ESTIMACIÓN PRUEBA
est_ins<-cbind(forecast_anual[1:dim(prueba)[1],],forecast_semestral[1:dim(prueba)[1],],
               forecast_cuatrimestre[1:dim(prueba)[1],], forecast_trimestre[1:dim(prueba)[1],],
               forecast_bimestre1[1:dim(prueba)[1],], forecast_bimestre2[1:dim(prueba)[1],], 
              forecast_sinciclo[1:dim(prueba)[1],])

est<-NULL
for (i in 1:dim(est_ins)[2]) {
  ajuste<-est_ins[,i]
  est<-c(est,ajuste)
}

sku_ins<-NULL
for (i in 1:dim(est_ins)[2]) {
  nombre<-rep(colnames(est_ins)[i], dim(est_ins)[1])
  sku_ins<-c(sku_ins,nombre)
}

estimacion_ins<-data.frame(fecha=rep(prueba$FECHA,dim(est_ins)[2]),sku=sku_ins , fcst=est)

#-------------------------------------------------------------------
#                               TRADE
#-------------------------------------------------------------------
#     CARGAR DATOS 
#----------------------
BI_T<-read_excel("Base_inicial.xlsx", sheet = "SKU TRADE", col_names = TRUE)
str(BI_T)
BI_T<-as.data.frame(BI_T)
BI_T$FECHA<-as.Date.character(BI_T$FECHA)


# ESTIMACIÓN DE HIPERPARÁMETROS (S Y T)
#----------------------------------------
parametros <- matrix(NA, nrow = dim(BI_T)[2], ncol = 3)
parametros[,1] <- colnames(BI_T)

# ESTACIONALIDAD
#-----------------

max<-NULL
ciclo<-NULL
x11()
par(mfrow=c(3,3))
for (i in 1:dim(BI_T)[2] ) {
  per<-spectrum(BI_T[,i], spans=c(3,3), log="no")
  max[i]<-which.max(per$spec)
  ciclo[i]<-1/(max[i]*1/12)
}
parametros[,2] <- round(ciclo,0)


#   TENDENCIA
#-----------------

# H0: Existe raíz unitaria
# p-val > 0.05 = no rechaza --> hay raíz
#p-val < 0.05 = rechaza --> no hay raíz
raices<-list()
ndiffs

for (i in 1:dim(BI_T)[2]){
  raices[[i]] <- fUnitRoots::adfTest(ts(BI_T[,i], start = c(2016,1), frequency = 12)) 
}
# 1, 0, 0, 0, 1, 0, 0, 0, 0 --> ver si se incluye o no la tendencia

tendencia <-c(NA,rep(1,3), rep(0,2), rep(1,4), 0, 1, 0, 
              rep(1,11), 0, rep(1,4), 0, rep(1,2), 0,
              rep(1,5), 0, rep(1,6), 0, 1, rep(0,2), rep(1,10), 
              0, rep(1,23), 0, rep(1,4), rep(0,2), 1)

parametros[,3] <- tendencia

# Adecuación tabla parámetros
#-------------------------------
parametros_df = as.data.frame(parametros)
colnames(parametros_df)<-c("sku", "s", "t")
parametros_df<-parametros_df[-1,]

entrenamientoT<-BI_T[1:round(dim(BI_T)[1]*0.9),] 
pruebaT<-BI_T[(round(dim(BI_T)[1]*0.9)+1):dim(BI_T)[1],]
# La dimensión de entrenamiento y prueba dependerá de las observaciones que se tengan. En este caso va de JAN 2016 - FEB 2022
# Quedan 64 de entrenamiento JAN 2016 - JUL 2021 y
# Quedan 7 de prueba AGO 2021 - FEB 2022 

#    DIVISIÓN X CICLO
#------------------------

anualT = parametros_df %>% filter(s==12) # 68 skus

semestralT = parametros_df %>% filter(s==6) # 11 skus

cuatrimestreT = parametros_df %>% filter(s==4) # 3 skus

trimestreT = parametros_df %>% filter(s==3) # 1 sku

bimestreT = parametros_df %>% filter(s==2) # 3 skus

SinCicloT = parametros_df %>% filter(s==1 | s==0) # 4 skus


#--------------------------------
#     SERIES CON CICLO ANUAL
#--------------------------------
anualT_ts = entrenamientoT %>%
  select(anualT$sku)
anualT_ts<-ts(anualT_ts, start = c(2016,1), frequency = 12)

ts_plot(anualT_ts)

#   AJUSTE MLE
#----------------
# INICIALIZAR
fit112 <- dlmMLE(anualT_ts[,1], c(rep(1,12)), build = fn1_12, hessian = T)
(modelo_anual_T <- fn1_12(fit112$par))

# ESTIMACIÓN PARÁMETROS
for (i in 2:dim(anualT)[1]) {
  fit112 <- dlmMLE(anualT_ts[,i], c(rep(1,12)), build = fn1_12, hessian = T,method = "Nelder-Mead" )
  mod <- fn1_12(fit112$par)
  
  modelo_anual_T = modelo_anual_T %+% mod
}


# FILTRO Y FORECAST
#--------------------
filtro <- dlmFilter(anualT_ts, modelo_anual_T)

forecast <- dlmForecast(filtro, nAhead = dim(pruebaT)[1] + 28, sampleNew = 1) # 35 = 7 meses prueba + 28 estimación
(forecast_anual_T<-forecast$f)
colnames(forecast_anual_T)<-anualT$sku

# RESIDUALES
#------------
anualT_ts_prueba = pruebaT %>% select(anualT$sku)

(res_anual_T<- anualT_ts_prueba - forecast_anual_T[1:dim(pruebaT)[1],])
mape_anual_T<-colMeans(abs(res_anual_T/anualT_ts_prueba))
rmse_anual_T<-sqrt(colMeans(res_anual_T^2))


#------------------------------------
#     SERIES CON CICLO SEMESTRAL
#------------------------------------
semestreT_ts = entrenamientoT %>%
  select(semestralT$sku)
semestreT_ts<-ts(semestreT_ts, start = c(2016,1), frequency = 12)

ts_plot(semestreT_ts)

#   AJUSTE MLE
#----------------

# INICIALIZAR
fit16 <- dlmMLE(semestreT_ts[,1], c(rep(1,12)), build = fn1_12, hessian = T)
(modelo_semestre_T <- fn1_12(fit16$par))


# ESTIMACIÓN PARÁMETROS
for (i in 2:dim(semestralT)[1]) {
  fit16 <- dlmMLE(semestreT_ts[,i], c(rep(1,12)), build = fn1_12, hessian = T,method = "Nelder-Mead" )
  mod <- fn1_12(fit16$par)
  
  modelo_semestre_T = modelo_semestre_T %+% mod
}

# FILTRO Y FORECAST
#-------------------

filtro <- dlmFilter(semestreT_ts, modelo_semestre_T)

forecast <- dlmForecast(filtro, nAhead = dim(pruebaT)[1] + 28, sampleNew = 1)
(forecast_semestral_T<-forecast$f)
colnames(forecast_semestral_T)<-semestralT$sku

# RESIDUALES
#------------
semestralT_ts_prueba = pruebaT %>% select(semestralT$sku)

(res_semestral_T<- semestralT_ts_prueba - forecast_semestral_T[1:dim(pruebaT)[1],])
mape_semestral_T<-colMeans(abs(res_semestral_T/semestralT_ts_prueba))
rmse_semestral_T<-sqrt(colMeans(res_semestral_T^2))


#----------------------------------------
#     SERIES CON CICLO CUATRIMESTRAL
#----------------------------------------
cuatrimestreT_ts = entrenamientoT %>%
  select(cuatrimestreT$sku)
(cuatrimestreT_ts<-ts(cuatrimestreT_ts, start = c(2016,1), frequency = 12))

ts_plot(cuatrimestreT_ts)

#   AJUSTE MLE
#----------------

# INICIALIZAR
fit14 <- dlmMLE(cuatrimestreT_ts[,1], c(rep(1,4)), build = fn1_4, hessian = T)
modelo_cuatrimestre_T <- fn1_4(fit14$par)

# ESTIMACIÓN PARÁMETROS
for (i in 2:dim(cuatrimestreT)[1]) {
  fit14 <- dlmMLE(cuatrimestreT_ts[,i], c(rep(1,4)), build = fn1_4, hessian = T,method = "Nelder-Mead" )
  mod <- fn1_4(fit14$par)
  
  modelo_cuatrimestre_T = modelo_cuatrimestre_T %+% mod
}

# FILTRO Y FORECAST
#-------------------

filtro <- dlmFilter(cuatrimestreT_ts, modelo_cuatrimestre_T)

forecast <- dlmForecast(filtro, nAhead = dim(pruebaT)[1] + 28, sampleNew = 1)
(forecast_cuatrimestre_T<-forecast$f)
colnames(forecast_cuatrimestre_T)<-cuatrimestreT$sku

# RESIDUALES
#-------------
cuatrimestreT_ts_prueba = pruebaT %>% select(cuatrimestreT$sku)

(res_cuatrimestre_T<- cuatrimestreT_ts_prueba - forecast_cuatrimestre_T[1:dim(pruebaT)[1],])
mape_cuatrimestre_T<-colMeans(abs(res_cuatrimestre_T/cuatrimestreT_ts_prueba))
rmse_cuatrimestre_T<-sqrt(colMeans(res_cuatrimestre_T^2))


#----------------------------------------
#     SERIES CON CICLO TRIMESTRAL
#----------------------------------------
trimestreT_ts = entrenamientoT %>%
  select(trimestreT$sku)
(trimestreT_ts<-ts(trimestreT_ts, start = c(2016,1), frequency = 12))
ts_plot(trimestreT_ts)

#   AJUSTE MLE
#----------------

# ESTIMACIÓN PARÁMETROS
fit13 <- dlmMLE(trimestreT_ts, c(rep(1,3)), build = fn1_3, hessian = T)
modelo_trimestre_T <- fn1_3(fit13$par)


# FILTRO Y FORECAST
#-------------------

filtro <- dlmFilter(trimestreT_ts, modelo_trimestre_T)

forecast <- dlmForecast(filtro, nAhead = dim(pruebaT)[1] + 28, sampleNew = 1)
(forecast_trimestre_T<-forecast$f)

# RESIDUALES
#-------------
trimestreT_ts_prueba = pruebaT %>% select(trimestreT$sku)

(res_trimestre_T<- trimestreT_ts_prueba - forecast_trimestre_T[1:dim(pruebaT)[1],])
mape_trimestre_T<-colMeans(abs(res_trimestre_T/trimestreT_ts_prueba))
rmse_trimestre_T<-sqrt(colMeans(res_trimestre_T^2))



#----------------------------------------
#     SERIES CON CICLO BIMESTRAL            
#----------------------------------------

bimestreT_ts = entrenamientoT %>%
  select(bimestreT$sku)
(bimestreT_ts<-ts(bimestreT_ts, start = c(2016,1), frequency = 12))
ts_plot(bimestreT_ts)

#   AJUSTE MLE
#----------------

# INICIALIZAR
fit12 <- dlmMLE(bimestreT_ts[,1], c(rep(1,2)), build = fn1_2, hessian = T)
modelo_bimestre_T <- fn1_2(fit12$par)

# ESTIMACIÓN PARÁMETROS
for (i in 2:dim(bimestreT)[1]) {
  fit12 <- dlmMLE(bimestreT_ts[,i], c(rep(1,2)), build = fn1_2, hessian = T )
  mod <- fn1_2(fit12$par)
  
  modelo_bimestre_T = modelo_bimestre_T %+% mod
}

# FILTRO Y FORECAST
#-------------------

filtro <- dlmFilter(bimestreT_ts, modelo_bimestre_T)

forecast <- dlmForecast(filtro, nAhead = dim(pruebaT)[1] + 28, sampleNew = 1)
(forecast_bimestre_T<-forecast$f)
colnames(forecast_bimestre_T)<-bimestreT$sku

# RESIDUALES
#-------------
bimestreT_ts_prueba = pruebaT %>% select(bimestreT$sku)

(res_bimestre_T<- bimestreT_ts_prueba - forecast_bimestre_T[1:dim(pruebaT)[1],])
mape_bimestre_T<-colMeans(abs(res_bimestre_T/bimestreT_ts_prueba))
rmse_bimestre_T<-sqrt(colMeans(res_bimestre_T^2))


#--------------------------------
#        SERIES SIN CICLO 
#--------------------------------

SinCicloT_ts = entrenamientoT %>%
  select(SinCicloT$sku)
(SinCicloT_ts<-ts(SinCicloT_ts, start = c(2016,1), frequency = 12))

ts_plot(SinCicloT_ts)

#   AJUSTE MLE
#----------------

# INICIALIZAR
fit10 <- dlmMLE(SinCicloT_ts[,1], c(rep(1,2)), build = fn1_2, hessian = T)
modelo_sinciclo_T <- fn1_2(fit10$par)

# ESTIMACIÓN PARÁMETROS
for (i in 2:dim(SinCicloT)[1]) {
  fit10 <- dlmMLE(SinCicloT_ts[,i], c(rep(1,2)), build = fn1_2, hessian = T,method = "Nelder-Mead" )
  mod <- fn1_2(fit10$par)
  
  modelo_sinciclo_T = modelo_sinciclo_T %+% mod
}

# FILTRO Y FORECAST
#-------------------

filtro <- dlmFilter(SinCicloT_ts, modelo_sinciclo_T)

forecast <- dlmForecast(filtro, nAhead = dim(pruebaT)[1] + 28, sampleNew = 1)
(forecast_sinciclo_T<-forecast$f)
colnames(forecast_sinciclo_T)<-SinCicloT$sku


# RESIDUALES
#-------------
sincicloT_ts_prueba = pruebaT %>% select(SinCicloT$sku)

(res_sinciclo_T<- sincicloT_ts_prueba - forecast_sinciclo_T[1:dim(pruebaT)[1],])
mape_sinciclo_T<-colMeans(abs(res_sinciclo_T/sincicloT_ts_prueba))
rmse_sinciclo_T<-sqrt(colMeans(res_sinciclo_T^2))


#   ADECUACIÓN EXPORTACIÓN EXCEL
#----------------------------------

# MÉTRICAS
mape_tra<-c(mape_anual_T, mape_semestral_T, mape_cuatrimestre_T, mape_trimestre_T, mape_bimestre_T, mape_sinciclo_T)
rmse_tra<-c(rmse_anual_T, rmse_semestral_T, rmse_cuatrimestre_T, rmse_trimestre_T, rmse_bimestre_T, rmse_sinciclo_T)
sku_tra<-c(anualT$sku, semestralT$sku, cuatrimestreT$sku, trimestreT$sku, bimestreT$sku, SinCicloT$sku)

metricas_tra<-cbind(sku=sku_tra,mape=mape_tra, rmse=rmse_tra)

# FORECAST 28 meses
forecast_tra<-rbind(t(forecast_anual_T[(dim(pruebaT)[1]+1):dim(forecast_anual_T)[1],]),t(forecast_semestral_T[(dim(pruebaT)[1]+1):dim(forecast_semestral_T)[1],]),
                    t(forecast_cuatrimestre_T[(dim(pruebaT)[1]+1):dim(forecast_cuatrimestre_T)[1],]), t(forecast_trimestre_T[(dim(pruebaT)[1]+1):dim(forecast_trimestre_T)[1],]), 
                    t(forecast_bimestre_T[(dim(pruebaT)[1]+1):dim(forecast_bimestre_T)[1],]),t(forecast_sinciclo_T[(dim(pruebaT)[1]+1):dim(forecast_sinciclo_T)[1],])) 
                    
s<-as.Date("2022-03-01") # actualizar a 1 mes despúes del último mes que se tiene
colnames(forecast_tra)<-as.Date(seq(from=s, by="month", length.out=28))
row.names(forecast_tra)<-c(anualT$sku, semestralT$sku, cuatrimestreT$sku,
                           trimestreT$sku, bimestreT$sku,  SinCicloT$sku)

# ESTIMACIÓN COMPARACIÓN
est_tra<-cbind(forecast_anual_T[1:dim(pruebaT)[1],], forecast_semestral_T[1:dim(pruebaT)[1],],
               forecast_cuatrimestre_T[1:dim(pruebaT)[1],], forecast_trimestre_T[1:dim(pruebaT)[1],],
               forecast_bimestre_T[1:dim(pruebaT)[1],], forecast_sinciclo_T[1:dim(pruebaT)[1],])

est<-NULL
for (i in 1:dim(est_tra)[2]) {
  ajuste<-est_tra[,i]
  est<-c(est,ajuste)
}

sku_tra<-NULL
for (i in 1:dim(est_tra)[2]) {
  nombre<-rep(colnames(est_tra)[i], dim(est_tra)[1])
  sku_tra<-c(sku_tra,nombre)
}

estimacion_tra<-data.frame(fecha=rep(pruebaT$FECHA,dim(est_tra)[2]),sku=sku_tra , fcst=est)

#------------------------------------------------
#                Exportar a Excel
#------------------------------------------------

wb <- createWorkbook()

# crear el número de hojas y sus nombres

addWorksheet(wb, "Estimación TRADE")
addWorksheet(wb, "Métricas TRADE")
addWorksheet(wb, "Fcst 28 TRADE")
addWorksheet(wb, "Estimación INST")
addWorksheet(wb, "Métricas INST")
addWorksheet(wb, "Fcst 28 INST")

# Poner los datos en cada hojas
writeData(wb, 1, estimacion_tra)
writeData(wb, 2, metricas_tra)
writeData(wb, 3, forecast_tra, rowNames = T, colNames = T)
writeData(wb, 4, estimacion_ins)
writeData(wb, 5, metricas_ins)
writeData(wb, 6, forecast_ins, rowNames = T, colNames = T)


# Guardar el workbook en un File
# Es importante que en la ruta espScificada el archivo NO exista antes
saveWorkbook(wb, "C:\\Users\\ALVARLX23\\OneDrive - Abbott\\Documents\\Proyecto\\Modelos Estructurales\\forecast_ESPACIO_ESTADO.xlsx")




