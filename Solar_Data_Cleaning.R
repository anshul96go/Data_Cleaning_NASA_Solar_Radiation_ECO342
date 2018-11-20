#created a vecor of statements for read
c1[c(1:15)] <- "E:/solar_data_R/"
c2[c(1:15)] <- 2000:2014
c3 <- character(15)
c3[c(1:15)] <- ".xlsx"
c4 <- character(15)
c2_char <- paste(c2)
for(i in 1:15) {
  c4[i] <- paste(c1[i],c2_char[i],c3[i],sep="")
}

#c2_char is for write

x <- 2000

for(i in 1:15)
{
  #start: reading
  data <- read.xlsx(c4[i],sheetIndex = 1)
  #end
  
  
  #start: no. of days in year
  year = data$Year[1]
  if((year %% 4) == 0) {
    if((year %% 100) == 0){
      if ((year %% 400) == 0) {
        no_days <- 366
      } else {
        no_days <- 365
      }
    } else {
      no_days <- 366
    }
  } else {
    no_days <= 365
  }
  #end  
  
  
  #start: declaring a vector which is storing daily average radiation
  ##dhi
  daily_dhi <- c(1:no_days)
  daily_dhi[c(1:no_days)] <- 0
  monthly_dhi <- c(1:12)
  monthly_dhi[c(1:12)] <- 0
  ##dni
  daily_dni <- c(1:no_days)
  daily_dni[c(1:no_days)] <- 0
  monthly_dni <- c(1:12)
  monthly_dni[c(1:12)] <- 0
  ##ghi
  daily_ghi <- c(1:no_days)
  daily_ghi[c(1:no_days)] <- 0
  monthly_ghi <- c(1:12)
  monthly_ghi[c(1:12)] <- 0
  ##dewpoint
  daily_dew <- c(1:no_days)
  daily_dew[c(1:no_days)] <- 0
  monthly_dew <- c(1:12)
  monthly_dew[c(1:12)] <- 0
  ##temperature
  daily_tem <- c(1:no_days)
  daily_tem[c(1:no_days)] <- 0
  monthly_tem <- c(1:12)
  monthly_tem[c(1:12)] <- 0
  ##pressure
  daily_pre <- c(1:no_days)
  daily_pre[c(1:no_days)] <- 0
  monthly_pre <- c(1:12)
  monthly_pre[c(1:12)] <- 0
  ##windspeed
  daily_wsd <- c(1:no_days)
  daily_wsd[c(1:no_days)] <- 0
  monthly_wsd <- c(1:12)
  monthly_wsd[c(1:12)] <- 0
  #end
  
  
  #start: variable declaration
  n <- length(data$Year)
  i <- 1
  day <- 1
  month <- 1
  k <- 0
  
  #main function
  for (i in 1:n) {
    if(i==n) {
      #
      daily_dhi[day] = daily_dhi[day] + data$DHI[i]
      daily_dhi[day] = daily_dhi[day]/24
      daily_dni[day] = daily_dni[day] + data$DNI[i]
      daily_dni[day] = daily_dni[day]/24
      daily_ghi[day] = daily_ghi[day] + data$GHI[i]
      daily_ghi[day] = daily_ghi[day]/24
      daily_dew[day] = daily_dew[day] + data$Dew.Point[i]
      daily_dew[day] = daily_dew[day]/24
      daily_tem[day] = daily_tem[day] + data$Temperature[i]
      daily_tem[day] = daily_tem[day]/24
      daily_pre[day] = daily_pre[day] + data$Pressure[i]
      daily_pre[day] = daily_pre[day]/24
      daily_wsd[day] = daily_wsd[day] + data$Wind.Speed[i]
      daily_wsd[day] = daily_wsd[day]/24
      #
    } else if(k==24) {
      #
      daily_dhi[day] = daily_dhi[day]/24
      daily_dni[day] = daily_dni[day]/24
      daily_ghi[day] = daily_ghi[day]/24
      daily_dew[day] = daily_dew[day]/24
      daily_tem[day] = daily_tem[day]/24
      daily_pre[day] = daily_pre[day]/24
      daily_wsd[day] = daily_wsd[day]/24
      #
      day=day+1
      #
      daily_dhi[day] = data$DHI[i]
      daily_dni[day] = data$DNI[i]
      daily_ghi[day] = data$GHI[i]
      daily_dew[day] = data$Dew.Point[i]
      daily_tem[day] = data$Temperature[i]
      daily_pre[day] = data$Pressure[i]
      daily_wsd[day] = data$Wind.Speed[i]
      #
      k=1
    } else {
      #
      daily_dhi[day] = daily_dhi[day] + data$DHI[i]
      daily_dni[day] = daily_dni[day] + data$DNI[i]
      daily_ghi[day] = daily_ghi[day] + data$GHI[i]
      daily_dew[day] = daily_dew[day] + data$Dew.Point[i]
      daily_tem[day] = daily_tem[day] + data$Temperature[i]
      daily_pre[day] = daily_pre[day] + data$Pressure[i]
      daily_wsd[day] = daily_wsd[day] + data$Wind.Speed[i]
      #
      k=k+1
    } 
  }
  #working for leap year (monthly average)
  if(no_days == 366) {
    #
    ##
    monthly_dhi[1] = mean(daily_dhi[c(1:31)])
    monthly_dhi[2] = mean(daily_dhi[c(32:59)])
    monthly_dhi[3] = mean(daily_dhi[c(60:90)])
    monthly_dhi[4] = mean(daily_dhi[c(91:120)])
    monthly_dhi[5] = mean(daily_dhi[c(121:151)])
    monthly_dhi[6] = mean(daily_dhi[c(152:181)])
    monthly_dhi[7] = mean(daily_dhi[c(182:212)])
    monthly_dhi[8] = mean(daily_dhi[c(213:243)])
    monthly_dhi[9] = mean(daily_dhi[c(244:273)])
    monthly_dhi[10] = mean(daily_dhi[c(274:304)])
    monthly_dhi[11] = mean(daily_dhi[c(305:334)])
    monthly_dhi[12] = mean(daily_dhi[c(335:365)])
    ##
    monthly_dni[1] = mean(daily_dni[c(1:31)])
    monthly_dni[2] = mean(daily_dni[c(32:60)])
    monthly_dni[3] = mean(daily_dni[c(61:91)])
    monthly_dni[4] = mean(daily_dni[c(92:121)])
    monthly_dni[5] = mean(daily_dni[c(122:152)])
    monthly_dni[6] = mean(daily_dni[c(153:182)])
    monthly_dni[7] = mean(daily_dni[c(183:213)])
    monthly_dni[8] = mean(daily_dni[c(214:244)])
    monthly_dni[9] = mean(daily_dni[c(245:274)])
    monthly_dni[10] = mean(daily_dni[c(275:305)])
    monthly_dni[11] = mean(daily_dni[c(306:335)])
    monthly_dni[12] = mean(daily_dni[c(336:366)])
    ##
    monthly_ghi[1] = mean(daily_ghi[c(1:31)])
    monthly_ghi[2] = mean(daily_ghi[c(32:60)])
    monthly_ghi[3] = mean(daily_ghi[c(61:91)])
    monthly_ghi[4] = mean(daily_ghi[c(92:121)])
    monthly_ghi[5] = mean(daily_ghi[c(122:152)])
    monthly_ghi[6] = mean(daily_ghi[c(153:182)])
    monthly_ghi[7] = mean(daily_ghi[c(183:213)])
    monthly_ghi[8] = mean(daily_ghi[c(214:244)])
    monthly_ghi[9] = mean(daily_ghi[c(245:274)])
    monthly_ghi[10] = mean(daily_ghi[c(275:305)])
    monthly_ghi[11] = mean(daily_ghi[c(306:335)])
    monthly_ghi[12] = mean(daily_ghi[c(336:366)])
    ##
    monthly_dew[1] = mean(daily_dew[c(1:31)])
    monthly_dew[2] = mean(daily_dew[c(32:60)])
    monthly_dew[3] = mean(daily_dew[c(61:91)])
    monthly_dew[4] = mean(daily_dew[c(92:121)])
    monthly_dew[5] = mean(daily_dew[c(122:152)])
    monthly_dew[6] = mean(daily_dew[c(153:182)])
    monthly_dew[7] = mean(daily_dew[c(183:213)])
    monthly_dew[8] = mean(daily_dew[c(214:244)])
    monthly_dew[9] = mean(daily_dew[c(245:274)])
    monthly_dew[10] = mean(daily_dew[c(275:305)])
    monthly_dew[11] = mean(daily_dew[c(306:335)])
    monthly_dew[12] = mean(daily_dew[c(336:366)])
    ##
    monthly_tem[1] = mean(daily_tem[c(1:31)])
    monthly_tem[2] = mean(daily_tem[c(32:60)])
    monthly_tem[3] = mean(daily_tem[c(61:91)])
    monthly_tem[4] = mean(daily_tem[c(92:121)])
    monthly_tem[5] = mean(daily_tem[c(122:152)])
    monthly_tem[6] = mean(daily_tem[c(153:182)])
    monthly_tem[7] = mean(daily_tem[c(183:213)])
    monthly_tem[8] = mean(daily_tem[c(214:244)])
    monthly_tem[9] = mean(daily_tem[c(245:274)])
    monthly_tem[10] = mean(daily_tem[c(275:305)])
    monthly_tem[11] = mean(daily_tem[c(306:335)])
    monthly_tem[12] = mean(daily_tem[c(336:366)])
    ##
    monthly_pre[1] = mean(daily_pre[c(1:31)])
    monthly_pre[2] = mean(daily_pre[c(32:60)])
    monthly_pre[3] = mean(daily_pre[c(61:91)])
    monthly_pre[4] = mean(daily_pre[c(92:121)])
    monthly_pre[5] = mean(daily_pre[c(122:152)])
    monthly_pre[6] = mean(daily_pre[c(153:182)])
    monthly_pre[7] = mean(daily_pre[c(183:213)])
    monthly_pre[8] = mean(daily_pre[c(214:244)])
    monthly_pre[9] = mean(daily_pre[c(245:274)])
    monthly_pre[10] = mean(daily_pre[c(275:305)])
    monthly_pre[11] = mean(daily_pre[c(306:335)])
    monthly_pre[12] = mean(daily_pre[c(336:366)])
    ##
    monthly_wsd[1] = mean(daily_wsd[c(1:31)])
    monthly_wsd[2] = mean(daily_wsd[c(32:60)])
    monthly_wsd[3] = mean(daily_wsd[c(61:91)])
    monthly_wsd[4] = mean(daily_wsd[c(92:121)])
    monthly_wsd[5] = mean(daily_wsd[c(122:152)])
    monthly_wsd[6] = mean(daily_wsd[c(153:182)])
    monthly_wsd[7] = mean(daily_wsd[c(183:213)])
    monthly_wsd[8] = mean(daily_wsd[c(214:244)])
    monthly_wsd[9] = mean(daily_wsd[c(245:274)])
    monthly_wsd[10] = mean(daily_wsd[c(275:305)])
    monthly_wsd[11] = mean(daily_wsd[c(306:335)])
    monthly_wsd[12] = mean(daily_wsd[c(336:366)])
    #
  } else {
    #
    ##
    monthly_dhi[1] = mean(daily_dhi[c(1:31)])
    monthly_dhi[2] = mean(daily_dhi[c(32:59)])
    monthly_dhi[3] = mean(daily_dhi[c(60:90)])
    monthly_dhi[4] = mean(daily_dhi[c(91:120)])
    monthly_dhi[5] = mean(daily_dhi[c(121:151)])
    monthly_dhi[6] = mean(daily_dhi[c(152:181)])
    monthly_dhi[7] = mean(daily_dhi[c(182:212)])
    monthly_dhi[8] = mean(daily_dhi[c(213:243)])
    monthly_dhi[9] = mean(daily_dhi[c(244:273)])
    monthly_dhi[10] = mean(daily_dhi[c(274:304)])
    monthly_dhi[11] = mean(daily_dhi[c(305:334)])
    monthly_dhi[12] = mean(daily_dhi[c(335:365)])
    ##
    monthly_dni[1] = mean(daily_dni[c(1:31)])
    monthly_dni[2] = mean(daily_dni[c(32:60)])
    monthly_dni[3] = mean(daily_dni[c(60:90)])
    monthly_dni[4] = mean(daily_dni[c(91:120)])
    monthly_dni[5] = mean(daily_dni[c(121:151)])
    monthly_dni[6] = mean(daily_dni[c(152:181)])
    monthly_dni[7] = mean(daily_dni[c(182:212)])
    monthly_dni[8] = mean(daily_dni[c(213:243)])
    monthly_dni[9] = mean(daily_dni[c(244:273)])
    monthly_dni[10] = mean(daily_dni[c(274:304)])
    monthly_dni[11] = mean(daily_dni[c(305:334)])
    monthly_dni[12] = mean(daily_dni[c(335:365)])
    ##
    monthly_ghi[1] = mean(daily_ghi[c(1:31)])
    monthly_ghi[2] = mean(daily_ghi[c(32:60)])
    monthly_ghi[3] = mean(daily_ghi[c(60:90)])
    monthly_ghi[4] = mean(daily_ghi[c(91:120)])
    monthly_ghi[5] = mean(daily_ghi[c(121:151)])
    monthly_ghi[6] = mean(daily_ghi[c(152:181)])
    monthly_ghi[7] = mean(daily_ghi[c(182:212)])
    monthly_ghi[8] = mean(daily_ghi[c(213:243)])
    monthly_ghi[9] = mean(daily_ghi[c(244:273)])
    monthly_ghi[10] = mean(daily_ghi[c(274:304)])
    monthly_ghi[11] = mean(daily_ghi[c(305:334)])
    monthly_ghi[12] = mean(daily_ghi[c(335:365)])
    ##
    monthly_dew[1] = mean(daily_dew[c(1:31)])
    monthly_dew[2] = mean(daily_dew[c(32:60)])
    monthly_dew[3] = mean(daily_dew[c(60:90)])
    monthly_dew[4] = mean(daily_dew[c(91:120)])
    monthly_dew[5] = mean(daily_dew[c(121:151)])
    monthly_dew[6] = mean(daily_dew[c(152:181)])
    monthly_dew[7] = mean(daily_dew[c(182:212)])
    monthly_dew[8] = mean(daily_dew[c(213:243)])
    monthly_dew[9] = mean(daily_dew[c(244:273)])
    monthly_dew[10] = mean(daily_dew[c(274:304)])
    monthly_dew[11] = mean(daily_dew[c(305:334)])
    monthly_dew[12] = mean(daily_dew[c(335:365)])
    ##
    monthly_tem[1] = mean(daily_tem[c(1:31)])
    monthly_tem[2] = mean(daily_tem[c(32:60)])
    monthly_tem[3] = mean(daily_tem[c(60:90)])
    monthly_tem[4] = mean(daily_tem[c(91:120)])
    monthly_tem[5] = mean(daily_tem[c(121:151)])
    monthly_tem[6] = mean(daily_tem[c(152:181)])
    monthly_tem[7] = mean(daily_tem[c(182:212)])
    monthly_tem[8] = mean(daily_tem[c(213:243)])
    monthly_tem[9] = mean(daily_tem[c(244:273)])
    monthly_tem[10] = mean(daily_tem[c(274:304)])
    monthly_tem[11] = mean(daily_tem[c(305:334)])
    monthly_tem[12] = mean(daily_tem[c(335:365)])
    ##
    monthly_pre[1] = mean(daily_pre[c(1:31)])
    monthly_pre[2] = mean(daily_pre[c(32:60)])
    monthly_pre[3] = mean(daily_pre[c(60:90)])
    monthly_pre[4] = mean(daily_pre[c(91:120)])
    monthly_pre[5] = mean(daily_pre[c(121:151)])
    monthly_pre[6] = mean(daily_pre[c(152:181)])
    monthly_pre[7] = mean(daily_pre[c(182:212)])
    monthly_pre[8] = mean(daily_pre[c(213:243)])
    monthly_pre[9] = mean(daily_pre[c(244:273)])
    monthly_pre[10] = mean(daily_pre[c(274:304)])
    monthly_pre[11] = mean(daily_pre[c(305:334)])
    monthly_pre[12] = mean(daily_pre[c(335:365)])
    ##
    monthly_wsd[1] = mean(daily_wsd[c(1:31)])
    monthly_wsd[2] = mean(daily_wsd[c(32:60)])
    monthly_wsd[3] = mean(daily_wsd[c(60:90)])
    monthly_wsd[4] = mean(daily_wsd[c(91:120)])
    monthly_wsd[5] = mean(daily_wsd[c(121:151)])
    monthly_wsd[6] = mean(daily_wsd[c(152:181)])
    monthly_wsd[7] = mean(daily_wsd[c(182:212)])
    monthly_wsd[8] = mean(daily_wsd[c(213:243)])
    monthly_wsd[9] = mean(daily_wsd[c(244:273)])
    monthly_wsd[10] = mean(daily_wsd[c(274:304)])
    monthly_wsd[11] = mean(daily_wsd[c(305:334)])
    monthly_wsd[12] = mean(daily_wsd[c(335:365)])
    #
  }
  
  #creating dataframe
  new_data <- data.frame(daily_dew,daily_dhi,daily_dni,daily_ghi,daily_pre,daily_tem,daily_wsd)
  #exporting data
  print(c4[i])
  typeof(c4[i])
  write.xlsx(new_data,"e:/data_delhi.xlsx",sheetName = as.character(x), append = TRUE)
  x <- x+1
}