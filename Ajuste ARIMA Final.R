setwd("C:/Users/ALVARLX23/OneDrive - Abbott/Documents/Proyecto")
rm(list=ls()) # borrar todos los elementos del enviroment


# LIBRERIAS
#-----------

# Instalar paquetes necesarios (solo debe hacerse 1 vez por equipo)
install.packages("readxl") # Leer archivos excel
install.packages("FitAR") ; install.packages("TSstudio"); install.packages("fUnitRoots") 
install.packages('Metrics') # medición del modelo


# Cargar los paquetes instalados (debe hacerse cada vez que se utilice el script)
library(readxl); library(openxlsx)
library(FitAR); library(TSstudio); library(forecast)
library(Metrics); library(fUnitRoots)

# DATOS
#-------
# Toca tener cuidado con la ruta (definida en la línea 1 con "setwd")
# y con los nombres del archivo y las respectivas hojas.
# Los datos se deben actualizar en el Excel.

# TRADE
BI_T<-read_excel("Base_inicial.xlsx", sheet = "SKU TRADE", col_names = TRUE)
str(BI_T)
BI_T<-as.data.frame(BI_T)
BI_T$FECHA<-as.Date.character(BI_T$FECHA)
dim(BI_T)[2] # número de sku´s

# INSTITUCIONES
BI_INS<-read_excel("Base_inicial.xlsx", sheet = "SKU INS", col_names = TRUE)
str(BI_INS)
BI_INS<-as.data.frame(BI_INS)
BI_INS$FECHA<-as.Date.character(BI_INS$FECHA)
dim(BI_INS)[2] # número de sku´s


#                       AUTOARIMA TRADE
#-----------------------------------------------------------

# Dividir entrenamiento - prueba
#---------------------------------

# Se toma el 90% de los datos de entrenamiento y el 10% de prueba
entrenamiento<-BI_T[1:round(dim(BI_T)[1]*0.9),]
prueba<-BI_T[(round(dim(BI_T)[1]*0.9)+1):dim(BI_T)[1],]

auto_fit<-matrix(NA,dim(entrenamiento)[1],dim(entrenamiento)[2])
auto_fcst<-matrix(NA,dim(prueba)[1],dim(entrenamiento)[2])
AIC<-NULL; BIC <-NULL # vamos con MAPE
futuro<-matrix(NA,28,dim(entrenamiento)[2]) # estimación 28 meses adelante


# Ajustar Modelo
#-----------------
# Tener cuidado cual es la primera fecha que se está tomando: 
# start = c(2016,1) representa el 1 mes del 2016
# ARIMA es la familia de modelos que se están ajustando

for (i in 2:dim(entrenamiento)[2]) {
  af<-auto.arima(ts(entrenamiento[,i], start = c(2016,1), frequency = 12), biasadj = T)
  fcs<-forecast(af,dim(prueba)[1]) # forecast meses prueba 
  auto_fit[,i]<-af$fitted
  auto_fcst[,i]<-fcs$mean
  
  AIC[i] <- AIC(af); BIC[i] <- BIC(af)
  futuro[,i]<-forecast(af,dim(prueba)[1]+28)$mean[8:(dim(prueba)[1]+28)] # 7 prueba + 28 proyección
}

# Ajuste del data frame de los datos ajustados 
colnames(auto_fcst)<-colnames(BI_T)
colnames(auto_fit)<-colnames(BI_T)

auto_fit<-as.data.frame(auto_fit)
auto_fit$FECHA <- BI_T$FECHA[1:round(dim(BI_T)[1]*0.9)]
head(auto_fit)

# Adecuación forecast 28 meses adelante
futuro<-futuro[,-1]
futuro<-as.data.frame(t(futuro))  

s<-as.Date("2022-03-01") # actualizar a 1 mes despúes del último mes real que se tiene
colnames(futuro)<-as.Date(seq(from=s, by="month", length.out=28))
row.names(futuro)<-colnames(BI_T)[2:dim(BI_T)[2]]

# GRÁFICOS
#-----------
# Gráficos reales vs. ajustados del conjunto de entrenamiento y de pruena

x11()
par(mfrow=c(4,3))
for (i in 2:13) {
  plot(BI_T$FECHA, BI_T[,i], type = "l", main = colnames(BI_T)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_T$FECHA, c(auto_fit[,i],auto_fcst[,i]) , col="red")
}

x11()
par(mfrow=c(4,3))
for (i in 14:25) {
  plot(BI_T$FECHA, BI_T[,i], type = "l", main = colnames(BI_T)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_T$FECHA, c(auto_fit[,i],auto_fcst[,i]) , col="red")
}

x11()
par(mfrow=c(4,3))
for (i in 26:37) {
  plot(BI_T$FECHA, BI_T[,i], type = "l", main = colnames(BI_T)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_T$FECHA, c(auto_fit[,i],auto_fcst[,i]) , col="red")
}

x11()
par(mfrow=c(4,3))
for (i in 38:49) {
  plot(BI_T$FECHA, BI_T[,i], type = "l", main = colnames(BI_T)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_T$FECHA, c(auto_fit[,i],auto_fcst[,i]) , col="red")
}

x11()
par(mfrow=c(4,3))
for (i in 50:61) {
  plot(BI_T$FECHA, BI_T[,i], type = "l", main = colnames(BI_T)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_T$FECHA, c(auto_fit[,i],auto_fcst[,i]) , col="red")
}

x11()
par(mfrow=c(3,3))
for (i in 62:70) {
  plot(BI_T$FECHA, BI_T[,i], type = "l", main = colnames(BI_T)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_T$FECHA, c(auto_fit[,i],auto_fcst[,i]) , col="red")
}

x11()
par(mfrow=c(3,2))
for (i in 71:79) {
  plot(BI_T$FECHA, BI_T[,i], type = "l", main = colnames(BI_T)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_T$FECHA, c(auto_fit[,i],auto_fcst[,i]) , col="red")
}

# EVALUACIÓN DE MODELO
#-----------------------

# Aquí se obtienen los resultados de las métricas para comparar los modelos
# con los diferentes modelos disponibles en Abbott

rmse<-NULL; mape<-NULL; mae<-NULL # mae no

for (i in 2:dim(prueba)[2]) {
  rmse[i] <- Metrics::rmse(prueba[,i], auto_fcst[,i])
  mape[i] <- Metrics::mape(prueba[,i], auto_fcst[,i])
  mae[i] <- Metrics::mae(prueba[,i], auto_fcst[,i])
}
rmse<-rmse[-1]; mape<- mape[-1]; mae<-mae[-1]
AIC <- AIC[-1]; BIC <- BIC[-1]


metricas <- cbind(sku=colnames(BI_T)[2:dim(BI_T)[2]],
                      mape=mape, rmse = rmse,AIC= AIC)


# ADECUACIÓN EXPORTACIÓN EXCEL
#------------------------------

est<-NULL; sku<-NULL
for (i in 2:dim(auto_fcst)[2]) {
  ajuste<-c(auto_fit[,i], auto_fcst[,i])
  est<-c(est,ajuste) }

for (i in 2:dim(auto_fcst)[2]) {
  nombre<-rep(colnames(BI_T)[i], dim(BI_T)[1])
  sku<-c(sku,nombre) }

estimacion_trade <- data.frame(fecha=rep(BI_T$FECHA,(dim(BI_T)[2])-1),sku=sku , fcst=est)


#-----------------------------------------------------------
#                   AUTOARIMA INSTITUCIONES
#-----------------------------------------------------------

# Dividir entrenamiento - prueba
#---------------------------------
entrenamiento_ins<-BI_INS[1:round(dim(BI_INS)[1]*0.9),]
prueba_ins<-BI_INS[(round(dim(BI_INS)[1]*0.9)+1):dim(BI_INS)[1],]

auto_fit_ins<-matrix(NA,dim(entrenamiento_ins)[1],dim(entrenamiento_ins)[2])
auto_fcst_ins<-matrix(NA,dim(prueba_ins)[1],dim(entrenamiento_ins)[2])
AIC_ins <- NULL; BIC_ins <- NULL
futuro_ins<-matrix(NA,28,dim(entrenamiento_ins)[2]) # estimación 28 meses adelante

# Ajustar Modelo
#------------------

for (i in 2:dim(entrenamiento_ins)[2]) {
  af<-auto.arima(ts(entrenamiento_ins[,i], start = c(2016,1), frequency = 12))
  fcs<-forecast(af,7)
  auto_fit_ins[,i]<-af$fitted
  auto_fcst_ins[,i]<-fcs$mean
  AIC_ins[i]<-AIC(af)
  BIC_ins[i]<-BIC(af)
  futuro_ins[,i]<-forecast(af,dim(prueba_ins)[1]+28)$mean[8:(dim(prueba_ins)[1]+28)]
}


# Adecuación data frame datos ajustados
colnames(auto_fcst_ins)<-colnames(BI_INS)
colnames(auto_fit_ins)<-colnames(BI_INS)
auto_fit_ins<-as.data.frame(auto_fit_ins)
auto_fit_ins$FECHA <- BI_INS$FECHA[1:round(dim(BI_INS)[1]*0.9)]
head(auto_fit_ins)

# Adecuación forecast 28 meses adelante

futuro_ins<-futuro_ins[,-1]
futuro_ins<-as.data.frame(t(futuro_ins))  

s<-as.Date("2022-03-01") # actualizar a 1 mes despúes del último mes que se tiene
colnames(futuro_ins)<-as.Date(seq(from=s, by="month", length.out=28))
row.names(futuro_ins)<-colnames(BI_INS)[2:dim(BI_INS)[2]]

# GRÁFICOS (MESES CON INFORMACIÓN)
#-----------------------------------

x11()
par(mfrow=c(4,3))
for (i in 2:13) {
  plot(BI_INS$FECHA, BI_INS[,i], type = "l", main = colnames(BI_INS)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_INS$FECHA, c(auto_fit_ins[,i],auto_fcst_ins[,i]) , col="red")
}

x11()
par(mfrow=c(4,3))
for (i in 14:25) {
  plot(BI_INS$FECHA, BI_INS[,i], type = "l", main = colnames(BI_INS)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_INS$FECHA, c(auto_fit_ins[,i],auto_fcst_ins[,i]) , col="red")
}

x11()
par(mfrow=c(4,3))
for (i in 26:37) {
  plot(BI_INS$FECHA, BI_INS[,i], type = "l", main = colnames(BI_INS)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_INS$FECHA, c(auto_fit_ins[,i],auto_fcst_ins[,i]) , col="red")
}

x11()
par(mfrow=c(4,3))
for (i in 38:49) {
  plot(BI_INS$FECHA, BI_INS[,i], type = "l", main = colnames(BI_INS)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_INS$FECHA, c(auto_fit_ins[,i],auto_fcst_ins[,i]) , col="red")
}

x11()
par(mfrow=c(4,3))
for (i in 50:61) {
  plot(BI_INS$FECHA, BI_INS[,i], type = "l", main = colnames(BI_INS)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_INS$FECHA, c(auto_fit_ins[,i],auto_fcst_ins[,i]) , col="red")
}

x11()
par(mfrow=c(3,3))
for (i in 62:70) {
  plot(BI_INS$FECHA, BI_INS[,i], type = "l", main = colnames(BI_INS)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_INS$FECHA, c(auto_fit_ins[,i],auto_fcst_ins[,i]) , col="red")
}

x11()
par(mfrow=c(3,2))
for (i in 71:75) {
  plot(BI_INS$FECHA, BI_INS[,i], type = "l", main = colnames(BI_INS)[i],
       xlab = "Fecha", ylab = "unidades")
  lines(BI_INS$FECHA, c(auto_fit_ins[,i],auto_fcst_ins[,i]) , col="red")
}


# EVALUACIÓN DE MODELO
#-----------------------

# Aquí se obtienen los resultados de las métricas para comparar
# con los diferentes modelos disponibles en Abbott

rmse_ins<-NULL; mape_ins<-NULL; mae_ins<-NULL

for (i in 2:dim(prueba_ins)[2]) {
  rmse_ins[i] <- Metrics::rmse(prueba_ins[,i], auto_fcst_ins[,i])
  mape_ins[i] <- Metrics::mape(prueba_ins[,i], auto_fcst_ins[,i])
  mae_ins[i] <- Metrics::mae(prueba_ins[,i], auto_fcst_ins[,i])
}
rmse_ins<-rmse_ins[-1]; mape_ins<- mape_ins[-1]; mae_ins<-mae_ins[-1]
AIC_ins <- AIC_ins[-1]; BIC_ins <- BIC_ins[-1]

metricas_ins <- cbind(sku=colnames(BI_INS)[2:86],
                      mape=mape_ins, rmse = rmse_ins,AIC= AIC_ins)


# ADECUACIÓN EXPORTACIÓN EXCEL
#------------------------------
est_ins<-NULL
for (i in 2:dim(auto_fcst_ins)[2]) {
  ajuste<-c(auto_fit_ins[,i], auto_fcst_ins[,i])
  est_ins<-c(est_ins,ajuste)
}

sku_ins<-NULL
for (i in 2:dim(auto_fcst_ins)[2]) {
  nombre<-rep(colnames(BI_INS)[i], dim(BI_INS)[1])
  sku_ins<-c(sku_ins,nombre)
}

estimacion_ins<-data.frame(fecha=rep(BI_INS$FECHA,(dim(BI_INS)[2])-1),sku=sku_ins , fcst=est_ins)

# Exportar a Excel
#------------------

wb <- createWorkbook()

# crear el número de hojas y sus nombres
addWorksheet(wb, "Fcst 28 TRADE")
addWorksheet(wb, "Métricas TRADE")
addWorksheet(wb, "Estimación TRADE")
addWorksheet(wb, "Fcst 28 INST")
addWorksheet(wb, "Métricas INST")
addWorksheet(wb, "Estimación INST")

# Poner los datos en cada hoja
writeData(wb, 1, futuro, rowNames = T)
writeData(wb, 2, metricas)
writeData(wb, 3, estimacion_trade)
writeData(wb, 4, futuro_ins, rowNames = T)
writeData(wb, 5, metricas_ins)
writeData(wb, 6, estimacion_ins)

# Guardar el workbook en un File
# Es importante que en la ruta especificada el archivo NO exista antes
saveWorkbook(wb, "C:\\Users\\ALVARLX23\\OneDrive - Abbott\\Documents\\Proyecto\\forecast_ARIMA.xlsx")
