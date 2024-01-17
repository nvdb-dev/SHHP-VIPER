################################################################################
#
# generateAQI.R
#
# Purpose: Generate Air Quality Index based on values from EPA Guidelines:
#          2018 U.S. Environmental Protection Agency Guidelines for the 
#          Reporting of Daily Air Quality – the Air Quality Index (AQI); 
#          EPA454/B-18-007
#
# Requirements: xlsx file with first row containing names each of the pollutants
#               as described: NO2; CO; PM2.5; PM10; SO2.
#               Subsequent rows in this csv file should contain the raw values
#               (i.e., microgram/mcubed, ppm, ppb, etc.) of the pollutants as 
#               sampled for any number of timepoints.
#
# Author:  Nicholas van den Berg, nic.vandenberg@outlook.com, affiliated with:
#          - CIUSS NIM Simonelli Lab for Sleep Health and Human Performance,
#          Hopital Sacre Coeur, Montreal
#          - VIPER Research Unit, Royal Military Academy, Brussels, Belgium
# 
# Created: 3 March, 2023
#
# Modifications: 
# 16 January, 2024 - Changed Low Breakpoint values to always be a continuous 
#                    decimal, in other words 0.1 values higher than the previous
#                    category's high breakpoint, and not a whole integer higher.
#                    This strays from the EPA guidelines by a maximum of one raw
#                    integer, but it avoids catastrophic exaggerations!
#                   
# Copyright Nicholas van den Berg 2024
################################################################################

#Preamble: Get libraries, load excel data (GUI pop, use xlsx), create empty vars
print("Did you remember to change the filename, on the last line?")
setwd("C:\\Users\\nicva\\OneDrive\\ONR-G\\Finland\\2023\\FinlandData\\Analysis\\AirQuality\\AQI")
file_path <- file.choose()
library(readxl)
library(dplyr)
library(tibble)
myAQI <- read_xlsx(file_path)

#Preallocate variables...sort of
myAQI['AQI_NO2']<-0; myAQI['AQI_CO']<-0; myAQI['AQI_PM2.5']<-0; myAQI['AQI_PM10']<-0; myAQI['AQI_SO2']<-0

## Sets up Hard coded Table of Breakpoints, from EPA Air Quality Guidelines
AQI_bp <- matrix(1:140, ncol = 28,dimnames = list(c("NO2","CO","PM2.5","PM10","SO2"),c("GoodLoBP","GoodHiBP","ModLoBP","ModHiBP","UnHealthSensLoBP","UnHealthSensHiBP","UnHealthLoBP","UnHealthHiBP","VUnHealthLoBP","VUnHealthHiBP","HazLoBP","HazHiBP","VHazLoBP","VHazHiBP","GoodLoAQI","GoodHiAQI","ModLoAQI","ModHiAQI","UnHealthSensLoAQI","UnHealthSensHiAQI","UnHealthLoAQI","UnHealthHiAQI","VUnHealthLoAQI","VUnHealthHiAQI","HazLoAQI","HazHiAQI","VHazLoAQI","VHazHiAQI")))

#Hard Code AQI value Breakpoints based on EPA guidelines
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("GoodLoAQI","GoodHiAQI")]<-c(0,0,0,0,0,50,50,50,50,50)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("ModLoAQI","ModHiAQI")]<-c(51,51,51,51,51,100,100,100,100,100)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("UnHealthSensLoAQI","UnHealthSensHiAQI")]<-c(101,101,101,101,101,150,150,150,150,150)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("UnHealthLoAQI","UnHealthHiAQI")]<-c(151,151,151,151,151,200,200,200,200,200)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("VUnHealthLoAQI","VUnHealthHiAQI")]<-c(201,201,201,201,201,300,300,300,300,300)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("HazLoAQI","HazHiAQI")]<-c(301,301,301,301,301,400,400,400,400,400)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("VHazLoAQI","VHazHiAQI")]<-c(401,401,401,401,401,500,500,500,500,500)
#Hard Code Breakpoints based on EPA guidelines
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("GoodLoBP","GoodHiBP")]<-c(0,0,0,0,0,53,4.4,12,54,35)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("ModLoBP","ModHiBP")]<-c(53.1,4.5,12.1,54.1,35.1,100,9.4,35.4,154,75)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("UnHealthSensLoBP","UnHealthSensHiBP")]<-c(100.1,9.5,35.5,154.1,75.1,360,12.4,55.4,254,185)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("UnHealthLoBP","UnHealthHiBP")]<-c(360.1,12.5,55.5,254.1,185.1,649,15.4,150.4,354,304)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("VUnHealthLoBP","VUnHealthHiBP")]<-c(649.1,15.5,150.5,354.1,304.1,1249,30.4,250.4,424,604)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("HazLoBP","HazHiBP")]<-c(1249.1,30.5,250.5,424.1,604.1,1649,40.4,350.4,504,804)
AQI_bp[c("NO2","CO","PM2.5","PM10","SO2"),c("VHazLoBP","VHazHiBP")]<-c(1649.1,40.5,350.5,504.1,804.1,2049,50.4,500.4,604,1004)


## Now complete the AQI indices for each value, given our values, relative to EPA guidelines

#NO2
myAQI$AQI_NO2 <-ifelse(AQI_bp["NO2","GoodHiBP"]>= myAQI$NO2 & myAQI$NO2 >= AQI_bp["NO2","GoodLoBP"],((AQI_bp["NO2","GoodHiAQI"]-AQI_bp["NO2","GoodLoAQI"])/(AQI_bp["NO2","GoodHiBP"]-AQI_bp["NO2","GoodLoBP"])*(myAQI$NO2-AQI_bp["NO2","GoodLoBP"])+AQI_bp["NO2","GoodLoAQI"]),
                       ifelse(AQI_bp["NO2","ModHiBP"]>= myAQI$NO2 & myAQI$NO2 >= AQI_bp["NO2","ModLoBP"],((AQI_bp["NO2","ModHiAQI"]-AQI_bp["NO2","ModLoAQI"])/(AQI_bp["NO2","ModHiBP"]-AQI_bp["NO2","ModLoBP"])*(myAQI$NO2-AQI_bp["NO2","ModLoBP"])+AQI_bp["NO2","ModLoAQI"]),
                              ifelse(AQI_bp["NO2","UnHealthSensHiBP"]>= myAQI$NO2 & myAQI$NO2 >= AQI_bp["NO2","UnHealthSensLoBP"],((AQI_bp["NO2","UnHealthSensHiAQI"]-AQI_bp["NO2","UnHealthSensLoAQI"])/(AQI_bp["NO2","UnHealthSensHiBP"]-AQI_bp["NO2","UnHealthSensLoBP"])*(myAQI$NO2-AQI_bp["NO2","UnHealthSensLoBP"])+AQI_bp["NO2","UnHealthSensLoAQI"]),
                                     ifelse(AQI_bp["NO2","UnHealthHiBP"]>= myAQI$NO2 & myAQI$NO2 >= AQI_bp["NO2","UnHealthLoBP"],((AQI_bp["NO2","UnHealthHiAQI"]-AQI_bp["NO2","UnHealthLoAQI"])/(AQI_bp["NO2","UnHealthHiBP"]-AQI_bp["NO2","UnHealthLoBP"])*(myAQI$NO2-AQI_bp["NO2","UnHealthLoBP"])+AQI_bp["NO2","UnHealthLoAQI"]),
                                            ifelse(AQI_bp["NO2","VUnHealthHiBP"]>= myAQI$NO2 & myAQI$NO2 >= AQI_bp["NO2","VUnHealthLoBP"],((AQI_bp["NO2","VUnHealthHiAQI"]-AQI_bp["NO2","VUnHealthLoAQI"])/(AQI_bp["NO2","VUnHealthHiBP"]-AQI_bp["NO2","VUnHealthLoBP"])*(myAQI$NO2-AQI_bp["NO2","VUnHealthLoBP"])+AQI_bp["NO2","VUnHealthLoAQI"]),
                                                   ifelse(AQI_bp["NO2","HazHiBP"]>= myAQI$NO2 & myAQI$NO2 >= AQI_bp["NO2","HazLoBP"],((AQI_bp["NO2","HazHiAQI"]-AQI_bp["NO2","HazLoAQI"])/(AQI_bp["NO2","HazHiBP"]-AQI_bp["NO2","HazLoBP"])*(myAQI$NO2-AQI_bp["NO2","HazLoBP"])+AQI_bp["NO2","HazLoAQI"]),
                                                          ifelse(myAQI$NO2 >= AQI_bp["NO2","VHazLoBP"],((AQI_bp["NO2","VHazHiAQI"]-AQI_bp["NO2","VHazLoAQI"])/(AQI_bp["NO2","VHazHiBP"]-AQI_bp["NO2","VHazLoBP"])*(myAQI$NO2-AQI_bp["NO2","VHazLoBP"])+AQI_bp["NO2","VHazLoAQI"]),"500")))))))

#CO
myAQI$AQI_CO <-ifelse(AQI_bp["CO","GoodHiBP"]>= myAQI$CO & myAQI$CO >= AQI_bp["CO","GoodLoBP"],((AQI_bp["CO","GoodHiAQI"]-AQI_bp["CO","GoodLoAQI"])/(AQI_bp["CO","GoodHiBP"]-AQI_bp["CO","GoodLoBP"])*(myAQI$CO-AQI_bp["CO","GoodLoBP"])+AQI_bp["CO","GoodLoAQI"]),
                       ifelse(AQI_bp["CO","ModHiBP"]>= myAQI$CO & myAQI$CO >= AQI_bp["CO","ModLoBP"],((AQI_bp["CO","ModHiAQI"]-AQI_bp["CO","ModLoAQI"])/(AQI_bp["CO","ModHiBP"]-AQI_bp["CO","ModLoBP"])*(myAQI$CO-AQI_bp["CO","ModLoBP"])+AQI_bp["CO","ModLoAQI"]),
                              ifelse(AQI_bp["CO","UnHealthSensHiBP"]>= myAQI$CO & myAQI$CO >= AQI_bp["CO","UnHealthSensLoBP"],((AQI_bp["CO","UnHealthSensHiAQI"]-AQI_bp["CO","UnHealthSensLoAQI"])/(AQI_bp["CO","UnHealthSensHiBP"]-AQI_bp["CO","UnHealthSensLoBP"])*(myAQI$CO-AQI_bp["CO","UnHealthSensLoBP"])+AQI_bp["CO","UnHealthSensLoAQI"]),
                                     ifelse(AQI_bp["CO","UnHealthHiBP"]>= myAQI$CO & myAQI$CO >= AQI_bp["CO","UnHealthLoBP"],((AQI_bp["CO","UnHealthHiAQI"]-AQI_bp["CO","UnHealthLoAQI"])/(AQI_bp["CO","UnHealthHiBP"]-AQI_bp["CO","UnHealthLoBP"])*(myAQI$CO-AQI_bp["CO","UnHealthLoBP"])+AQI_bp["CO","UnHealthLoAQI"]),
                                            ifelse(AQI_bp["CO","VUnHealthHiBP"]>= myAQI$CO & myAQI$CO >= AQI_bp["CO","VUnHealthLoBP"],((AQI_bp["CO","VUnHealthHiAQI"]-AQI_bp["CO","VUnHealthLoAQI"])/(AQI_bp["CO","VUnHealthHiBP"]-AQI_bp["CO","VUnHealthLoBP"])*(myAQI$CO-AQI_bp["CO","VUnHealthLoBP"])+AQI_bp["CO","VUnHealthLoAQI"]),
                                                   ifelse(AQI_bp["CO","HazHiBP"]>= myAQI$CO & myAQI$CO >= AQI_bp["CO","HazLoBP"],((AQI_bp["CO","HazHiAQI"]-AQI_bp["CO","HazLoAQI"])/(AQI_bp["CO","HazHiBP"]-AQI_bp["CO","HazLoBP"])*(myAQI$CO-AQI_bp["CO","HazLoBP"])+AQI_bp["CO","HazLoAQI"]),
                                                          ifelse(myAQI$CO >= AQI_bp["CO","VHazLoBP"],((AQI_bp["CO","VHazHiAQI"]-AQI_bp["CO","VHazLoAQI"])/(AQI_bp["CO","VHazHiBP"]-AQI_bp["CO","VHazLoBP"])*(myAQI$CO-AQI_bp["CO","VHazLoBP"])+AQI_bp["CO","VHazLoAQI"]),"500")))))))

#PM2.5
myAQI$AQI_PM2.5 <-ifelse(AQI_bp["PM2.5","GoodHiBP"]>= myAQI$PM2.5 & myAQI$PM2.5 >= AQI_bp["PM2.5","GoodLoBP"],((AQI_bp["PM2.5","GoodHiAQI"]-AQI_bp["PM2.5","GoodLoAQI"])/(AQI_bp["PM2.5","GoodHiBP"]-AQI_bp["PM2.5","GoodLoBP"])*(myAQI$PM2.5-AQI_bp["PM2.5","GoodLoBP"])+AQI_bp["PM2.5","GoodLoAQI"]),
                       ifelse(AQI_bp["PM2.5","ModHiBP"]>= myAQI$PM2.5 & myAQI$PM2.5 >= AQI_bp["PM2.5","ModLoBP"],((AQI_bp["PM2.5","ModHiAQI"]-AQI_bp["PM2.5","ModLoAQI"])/(AQI_bp["PM2.5","ModHiBP"]-AQI_bp["PM2.5","ModLoBP"])*(myAQI$PM2.5-AQI_bp["PM2.5","ModLoBP"])+AQI_bp["PM2.5","ModLoAQI"]),
                              ifelse(AQI_bp["PM2.5","UnHealthSensHiBP"]>= myAQI$PM2.5 & myAQI$PM2.5 >= AQI_bp["PM2.5","UnHealthSensLoBP"],((AQI_bp["PM2.5","UnHealthSensHiAQI"]-AQI_bp["PM2.5","UnHealthSensLoAQI"])/(AQI_bp["PM2.5","UnHealthSensHiBP"]-AQI_bp["PM2.5","UnHealthSensLoBP"])*(myAQI$PM2.5-AQI_bp["PM2.5","UnHealthSensLoBP"])+AQI_bp["PM2.5","UnHealthSensLoAQI"]),
                                     ifelse(AQI_bp["PM2.5","UnHealthHiBP"]>= myAQI$PM2.5 & myAQI$PM2.5 >= AQI_bp["PM2.5","UnHealthLoBP"],((AQI_bp["PM2.5","UnHealthHiAQI"]-AQI_bp["PM2.5","UnHealthLoAQI"])/(AQI_bp["PM2.5","UnHealthHiBP"]-AQI_bp["PM2.5","UnHealthLoBP"])*(myAQI$PM2.5-AQI_bp["PM2.5","UnHealthLoBP"])+AQI_bp["PM2.5","UnHealthLoAQI"]),
                                            ifelse(AQI_bp["PM2.5","VUnHealthHiBP"]>= myAQI$PM2.5 & myAQI$PM2.5 >= AQI_bp["PM2.5","VUnHealthLoBP"],((AQI_bp["PM2.5","VUnHealthHiAQI"]-AQI_bp["PM2.5","VUnHealthLoAQI"])/(AQI_bp["PM2.5","VUnHealthHiBP"]-AQI_bp["PM2.5","VUnHealthLoBP"])*(myAQI$PM2.5-AQI_bp["PM2.5","VUnHealthLoBP"])+AQI_bp["PM2.5","VUnHealthLoAQI"]),
                                                   ifelse(AQI_bp["PM2.5","HazHiBP"]>= myAQI$PM2.5 & myAQI$PM2.5 >= AQI_bp["PM2.5","HazLoBP"],((AQI_bp["PM2.5","HazHiAQI"]-AQI_bp["PM2.5","HazLoAQI"])/(AQI_bp["PM2.5","HazHiBP"]-AQI_bp["PM2.5","HazLoBP"])*(myAQI$PM2.5-AQI_bp["PM2.5","HazLoBP"])+AQI_bp["PM2.5","HazLoAQI"]),
                                                          ifelse(myAQI$PM2.5 >= AQI_bp["PM2.5","VHazLoBP"],((AQI_bp["PM2.5","VHazHiAQI"]-AQI_bp["PM2.5","VHazLoAQI"])/(AQI_bp["PM2.5","VHazHiBP"]-AQI_bp["PM2.5","VHazLoBP"])*(myAQI$PM2.5-AQI_bp["PM2.5","VHazLoBP"])+AQI_bp["PM2.5","VHazLoAQI"]),"500")))))))

#PM10
myAQI$AQI_PM10 <-ifelse(AQI_bp["PM10","GoodHiBP"]>= myAQI$PM10 & myAQI$PM10 >= AQI_bp["PM10","GoodLoBP"],((AQI_bp["PM10","GoodHiAQI"]-AQI_bp["PM10","GoodLoAQI"])/(AQI_bp["PM10","GoodHiBP"]-AQI_bp["PM10","GoodLoBP"])*(myAQI$PM10-AQI_bp["PM10","GoodLoBP"])+AQI_bp["PM10","GoodLoAQI"]),
                       ifelse(AQI_bp["PM10","ModHiBP"]>= myAQI$PM10 & myAQI$PM10 >= AQI_bp["PM10","ModLoBP"],((AQI_bp["PM10","ModHiAQI"]-AQI_bp["PM10","ModLoAQI"])/(AQI_bp["PM10","ModHiBP"]-AQI_bp["PM10","ModLoBP"])*(myAQI$PM10-AQI_bp["PM10","ModLoBP"])+AQI_bp["PM10","ModLoAQI"]),
                              ifelse(AQI_bp["PM10","UnHealthSensHiBP"]>= myAQI$PM10 & myAQI$PM10 >= AQI_bp["PM10","UnHealthSensLoBP"],((AQI_bp["PM10","UnHealthSensHiAQI"]-AQI_bp["PM10","UnHealthSensLoAQI"])/(AQI_bp["PM10","UnHealthSensHiBP"]-AQI_bp["PM10","UnHealthSensLoBP"])*(myAQI$PM10-AQI_bp["PM10","UnHealthSensLoBP"])+AQI_bp["PM10","UnHealthSensLoAQI"]),
                                     ifelse(AQI_bp["PM10","UnHealthHiBP"]>= myAQI$PM10 & myAQI$PM10 >= AQI_bp["PM10","UnHealthLoBP"],((AQI_bp["PM10","UnHealthHiAQI"]-AQI_bp["PM10","UnHealthLoAQI"])/(AQI_bp["PM10","UnHealthHiBP"]-AQI_bp["PM10","UnHealthLoBP"])*(myAQI$PM10-AQI_bp["PM10","UnHealthLoBP"])+AQI_bp["PM10","UnHealthLoAQI"]),
                                            ifelse(AQI_bp["PM10","VUnHealthHiBP"]>= myAQI$PM10 & myAQI$PM10 >= AQI_bp["PM10","VUnHealthLoBP"],((AQI_bp["PM10","VUnHealthHiAQI"]-AQI_bp["PM10","VUnHealthLoAQI"])/(AQI_bp["PM10","VUnHealthHiBP"]-AQI_bp["PM10","VUnHealthLoBP"])*(myAQI$PM10-AQI_bp["PM10","VUnHealthLoBP"])+AQI_bp["PM10","VUnHealthLoAQI"]),
                                                   ifelse(AQI_bp["PM10","HazHiBP"]>= myAQI$PM10 & myAQI$PM10 >= AQI_bp["PM10","HazLoBP"],((AQI_bp["PM10","HazHiAQI"]-AQI_bp["PM10","HazLoAQI"])/(AQI_bp["PM10","HazHiBP"]-AQI_bp["PM10","HazLoBP"])*(myAQI$PM10-AQI_bp["PM10","HazLoBP"])+AQI_bp["PM10","HazLoAQI"]),
                                                          ifelse(myAQI$PM10 >= AQI_bp["PM10","VHazLoBP"],((AQI_bp["PM10","VHazHiAQI"]-AQI_bp["PM10","VHazLoAQI"])/(AQI_bp["PM10","VHazHiBP"]-AQI_bp["PM10","VHazLoBP"])*(myAQI$PM10-AQI_bp["PM10","VHazLoBP"])+AQI_bp["PM10","VHazLoAQI"]),"500")))))))

#SO2
myAQI$AQI_SO2 <-ifelse(AQI_bp["SO2","GoodHiBP"]>= myAQI$SO2 & myAQI$SO2 >= AQI_bp["SO2","GoodLoBP"],((AQI_bp["SO2","GoodHiAQI"]-AQI_bp["SO2","GoodLoAQI"])/(AQI_bp["SO2","GoodHiBP"]-AQI_bp["SO2","GoodLoBP"])*(myAQI$SO2-AQI_bp["SO2","GoodLoBP"])+AQI_bp["SO2","GoodLoAQI"]),
                       ifelse(AQI_bp["SO2","ModHiBP"]>= myAQI$SO2 & myAQI$SO2 >= AQI_bp["SO2","ModLoBP"],((AQI_bp["SO2","ModHiAQI"]-AQI_bp["SO2","ModLoAQI"])/(AQI_bp["SO2","ModHiBP"]-AQI_bp["SO2","ModLoBP"])*(myAQI$SO2-AQI_bp["SO2","ModLoBP"])+AQI_bp["SO2","ModLoAQI"]),
                              ifelse(AQI_bp["SO2","UnHealthSensHiBP"]>= myAQI$SO2 & myAQI$SO2 >= AQI_bp["SO2","UnHealthSensLoBP"],((AQI_bp["SO2","UnHealthSensHiAQI"]-AQI_bp["SO2","UnHealthSensLoAQI"])/(AQI_bp["SO2","UnHealthSensHiBP"]-AQI_bp["SO2","UnHealthSensLoBP"])*(myAQI$SO2-AQI_bp["SO2","UnHealthSensLoBP"])+AQI_bp["SO2","UnHealthSensLoAQI"]),
                                     ifelse(AQI_bp["SO2","UnHealthHiBP"]>= myAQI$SO2 & myAQI$SO2 >= AQI_bp["SO2","UnHealthLoBP"],((AQI_bp["SO2","UnHealthHiAQI"]-AQI_bp["SO2","UnHealthLoAQI"])/(AQI_bp["SO2","UnHealthHiBP"]-AQI_bp["SO2","UnHealthLoBP"])*(myAQI$SO2-AQI_bp["SO2","UnHealthLoBP"])+AQI_bp["SO2","UnHealthLoAQI"]),
                                            ifelse(AQI_bp["SO2","VUnHealthHiBP"]>= myAQI$SO2 & myAQI$SO2 >= AQI_bp["SO2","VUnHealthLoBP"],((AQI_bp["SO2","VUnHealthHiAQI"]-AQI_bp["SO2","VUnHealthLoAQI"])/(AQI_bp["SO2","VUnHealthHiBP"]-AQI_bp["SO2","VUnHealthLoBP"])*(myAQI$SO2-AQI_bp["SO2","VUnHealthLoBP"])+AQI_bp["SO2","VUnHealthLoAQI"]),
                                                   ifelse(AQI_bp["SO2","HazHiBP"]>= myAQI$SO2 & myAQI$SO2 >= AQI_bp["SO2","HazLoBP"],((AQI_bp["SO2","HazHiAQI"]-AQI_bp["SO2","HazLoAQI"])/(AQI_bp["SO2","HazHiBP"]-AQI_bp["SO2","HazLoBP"])*(myAQI$SO2-AQI_bp["SO2","HazLoBP"])+AQI_bp["SO2","HazLoAQI"]),
                                                          ifelse(myAQI$SO2 >= AQI_bp["SO2","VHazLoBP"],((AQI_bp["SO2","VHazHiAQI"]-AQI_bp["SO2","VHazLoAQI"])/(AQI_bp["SO2","VHazHiBP"]-AQI_bp["SO2","VHazLoBP"])*(myAQI$SO2-AQI_bp["SO2","VHazLoBP"])+AQI_bp["SO2","VHazLoAQI"]),"500")))))))

##Now, save your new output to xlsx
library("writexl")
write_xlsx(myAQI,"C:\\Users\\nicva\\OneDrive\\ONR-G\\Finland\\2023\\FinlandData\\Analysis\\AirQuality\\AQI\\AQI_BarracksSensor1_V2.xlsx")
print("All Done!")