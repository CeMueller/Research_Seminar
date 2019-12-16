# Set Working Directory for Research Seminar
setwd("C:/Users/Cedric/Desktop/Uni/19_FS/Research Seminar/R_WD")

library(openxlsx)
# library(ggplot2)
library(fUnitRoots)
library(forecast)


# Read in excel file 
USD_CHF <- read.xlsx("USDCHF.xlsx")

# Rename column name of data.frame
colnames(USD_CHF) <- c("date", "fxrate")

# Drop duplicate values, i.e. weekend rates
USD_CHF <- USD_CHF[!duplicated(USD_CHF$fxrate),]

# Get major descriptive statistics
summary(USD_CHF)

#need to add mean, Variance, skewness, kurtosis for returns


# Change Excel date format to R date format and fxrate to numeric
USD_CHF$date <- as.Date(USD_CHF$date, format = "%d.%m.%Y")
USD_CHF$fxrate <- as.numeric(USD_CHF$fxrate)

# Create plot to visually  see if stationary or not
windows()
plot(USD_CHF$date, USD_CHF$fxrate, type="l", xlab="Year", ylab="USD/CHF")
grid(nx = NA, ny = NULL)

# Split sample up in "training" and "holdout" sample
n_rows <- nrow(USD_CHF)
USD_CHF_Copy <- USD_CHF
USD_CHF <- USD_CHF[0:(n_rows-62),]

# Prepare for ADF test
mRate <- ar(USD_CHF$fxrate, method= 'mle')
n1 <- mRate$order

# P Value of 0.4431  and DF ~= -0.5436 -> Cannot reject null hypothesis, 
# hence unit root is present
adfTest( USD_CHF[,2], lags=n1)

# Prep for ADF test with differenced TS
mRate2 <- ar(diff(USD_CHF$fxrate), method= 'mle')
n2 <- mRate2$order

# P Value of 0.01 and DF ~= -35.7492 -> Reject null hpothesis, 
# hence unit root not present anymore
adfTest( diff(USD_CHF[,2]), lags=n2)

# Create 1 time differenced ts
USD_CHF$diff_fxrate<- c(0,diff(USD_CHF$fxrate, lag = 1))

# Plot the differenced ts
windows()
plot(USD_CHF$date, USD_CHF$diff_fxrate, type="l", xlab="Year", ylab="USD/CHF Differenced (1)")
grid(nx = NA, ny = NULL)

# Create ACF plot to check for autocorrelation
windows()
acf(USD_CHF$diff_fxrate, type="correlation", lag=20, main="", ylab="ACF USD/CHF", ylim = c(-0.2, 0.2))

# Create PACF plot to check for partial autocorrelation
windows()
acf(USD_CHF$diff_fxrate, type="partial", lag=20, main="",  ylab="PACF USD/CHF", ylim = c(-0.2, 0.2))

# Create Box-Ljung Test --> p-value indicates that there is 
# autocorrelation in the last 20 lag terms
Box.test(USD_CHF$diff_fxrate, lag=10, type="Ljung")

# Check what model AIC would create
#  (p,d,q) = (3,1,3)
fit_aic <- auto.arima(USD_CHF$fxrate, approximation=FALSE, trace=T, stepwise=FALSE, max.order=15, ic="aic")
arimaorder(fit_aic)

# Check what model BIC would create
#  (p,d,q) = (0,1,0)
fit_bic <- auto.arima(USD_CHF$fxrate, approximation=FALSE, trace=T, stepwise=FALSE, max.order=15, ic="bic")
arimaorder(fit_bic)

# fit the ARIMA model with results from AIC
fit = arima(USD_CHF$fxrate, order=c(3,1,3), method="ML")

# Create forecast
fcast <- forecast(fit, h=5)
fcast <- as.numeric(fcast[["mean"]])
fcast <- tail(fcast,1)

# number of iterations manually applied to this time series
for (j in 61:6) {
  
  USD_CHF_Copy_ <- USD_CHF_Copy[0:(n_rows-j),]
  fit = arima(USD_CHF_Copy_$fxrate, order=c(3,1,3), method="ML")
  fcast_ <- forecast(fit, h=5)
  fcast_num <- as.numeric(fcast_[["mean"]])
  point_fcast <- tail(fcast_num,1)
  fcast <- rbind(fcast, point_fcast)
  
} 

rownames(fcast) <- c()
USD_CHF_fc_ARIMA <- as.numeric(fcast)
USD_CHF_2 <- USD_CHF_Copy[(n_rows-56):n_rows,]
date <- as.Date(USD_CHF_2$date, format = "%d.%m.%Y")
USD_CHF_FC_Output <- data.frame(date, fcast)
colnames(USD_CHF_FC_Output) <- c("date", "Arima")
write.xlsx(USD_CHF_FC_Output, "C:/Users/Cedric/Desktop/Uni/19_FS/Research Seminar/Final/Rolling_5D_Arima_USD_CHF.xlsx", sheetName="Rolling_5D_Arima_USD_CHF",  col.names=TRUE,  append=FALSE)






