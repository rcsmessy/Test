#NON-ISFE MONTHLY FILE SCRIPT----------------------------------------------------------------------------
#NOTES


# 1 - You will need to updates lines with correct file paths to your directories (check lines 29-39)
#     Check that the number of rows is still being covered by : Range =  "A14:BA215"
#
# 2 - You will need to ensure that the OrgHier.xls file has a list of every CCG Codes for which you 
#     are importing (any CCG files that do not have a matching CCG Code in your OrgHier table will be exluded)
#
# 3 - All column formats must match - else error when bind.  If there is an error at line 42 you will need 
#      to check the formats of the column causing the error and correct in the xlsm source file
#
#----------------------------------------------------------------------------------------------------
#install required packages
list.of.packages <- c("tidyverse","plyr","stringr","readxl","writexl","dplyr")


new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[, "Package"])]
if(length(new.packages)>0) install.packages(new.packages)
#load required packages
for (i in 1:length(list.of.packages)) {
  library(list.of.packages[i], character.only = TRUE)
}

#clean temporary package variables
rm(list = c("i","list.of.packages","new.packages"))

#----------------------------------------------------------------------------------------------------
#load files in RAW format

setwd("C:/Users/RMassam/Desktop/QIPP/201920/M02")   ### update end path with required month
Directory = "C:/Users/RMassam/Desktop/QIPP/201920/M02" ### path of files to load
filenames <- list.files(pattern='*xlsm') ### creates list if files to load

T1 <- lapply(filenames, read_excel, sheet = "CCG QIPP", range =  "A14:BA215") #imported tables
OrgHier <- read_excel("C:/Users/RMassam/Desktop/QIPP/201920/OrgHier.xlsx") # imports organisation hierarchy

rm(list = c("Directory","filenames"))

#----------------------------------------------------------------------------------------------------
#Combine Tables and discard unnecessary rows

T2 <- bind_rows(T1) # union all tables

#add headers
names(T2) <- c(
  "UniqueID",
  "Blank1",
  "CCG EFFICIENCY PLANS",
  "Area of Spend - Determines mapping on CCG Detail variance analysis",
  "Area of Spend - 10PP (fixed from mapping or select from dropdown for new schemes)",
  "National/Local Schemes",
  "Transactional / Transformational",
  "Scheme Part of RightCare? - Scheme Part of RightCare?",
  "Right Care Category - Right Care Category",
  "Blank2",
  "TOTAL QIPP - YTD - Plan",
  "TOTAL QIPP - YTD - Actual",
  "TOTAL QIPP - YTD - Variance",
  "TOTAL QIPP - YTD - % Achieved",
  "TOTAL QIPP - Forecast - Plan",
  "TOTAL QIPP - Forecast - Actual",
  "TOTAL QIPP - Forecast - Variance",
  "TOTAL QIPP - Forecast - % Achieved",
  "TOTAL QIPP",
  "NON RECURRENT - Forecast - Plan",
  "NON RECURRENT - Forecast - Actual",
  "Blank3",
  "Risk of Slippage",
  "Mitigation (QIPP extension)",
  "Blank4",
  "CCG Narrative",
  "Blank5",
  "Month 1 - Plan",
  "Month 1 - Actual",
  "Month 2 - Plan",
  "Month 2 - Actual",
  "Month 3 - Plan",
  "Month 3 - Actual",
  "Month 4 - Plan",
  "Month 4 - Actual",
  "Month 5 - Plan",
  "Month 5 - Actual",
  "Month 6 - Plan",
  "Month 6 - Actual",
  "Month 7 - Plan",
  "In-Month Actual - Month 7 - Actual",
  "Month 8 - Plan",
  "In-Month Forecast - Month 8 - Actual",
  "Month 9 - Plan",
  "In-Month Forecast - Month 9 - Actual",
  "Month 10 - Plan",
  "In-Month Forecast - Month 10 - Actual",
  "Month 11 - Plan",
  "In-Month Forecast - Month 11 - Actual",
  "Month 12 - Plan",
  "In-Month Forecast - Month 12 - Actual",
  "Blank6",
  "VALIDATIONS")


T2$Scheme <- !is.na(T2[1]) # flag nulls (where row is not a scheme)

# remove nulls - keep schemes only
T3 <- select(filter(T2, `Scheme` == TRUE),c(
  "UniqueID",
  "CCG EFFICIENCY PLANS",
  "Area of Spend - Determines mapping on CCG Detail variance analysis",
  "Area of Spend - 10PP (fixed from mapping or select from dropdown for new schemes)",
  "Transactional / Transformational",
  "Scheme Part of RightCare? - Scheme Part of RightCare?",
  "Right Care Category - Right Care Category",
  "TOTAL QIPP - Forecast - Plan",
  "TOTAL QIPP - Forecast - Actual",
  "TOTAL QIPP - Forecast - Variance",
  "TOTAL QIPP - Forecast - % Achieved",
  "TOTAL QIPP",
  "NON RECURRENT - Forecast - Plan",
  "NON RECURRENT - Forecast - Actual",
  "Risk of Slippage",
  "Mitigation (QIPP extension)",
  "CCG Narrative",
  "Scheme"))


T3$CCG <- substr(T3$UniqueID,1,3) # generates CCG Column

#----------------------------------------------------------------------------------------------------

# List of columns which need to be convert from Character to Numeric
cols.num <- c(
  "TOTAL QIPP - Forecast - Plan",
  "TOTAL QIPP - Forecast - Actual",
  "TOTAL QIPP - Forecast - Variance",
  "TOTAL QIPP - Forecast - % Achieved")

# convert rounded to 8 DP
T3[cols.num] <- round(sapply(T3[cols.num],as.numeric),8) 


#----------------------------------------------------------------------------------------------------
#Finalise

T4 <- merge(T3,OrgHier, by.x = "CCG", by.y = "Org Code") # adds CCG STP and Region info

# you can run this to checK whether any CCGs are dropping out after merge with OrgHeir
length(unique(T3$CCG)) # Num CCGs after merge with OrgHeir
length(unique(T4$CCG)) # Num CCGs after merge with OrgHeir
setdiff(T3unique,T4unique)

write_xlsx(T4, path = "qipp_Raw_Num.xlsx", col_names = TRUE) #Export to Excel - location as working directory

#----------------------------------------------------------------------------------------------------
#Run Validation Checks

#Create RC Subset
RC<-select(filter(T4, `Scheme Part of RightCare? - Scheme Part of RightCare?` == "Yes"),c(
  "TOTAL QIPP - Forecast - Plan","TOTAL QIPP - Forecast - Actual","Region","Scheme"))

#STP Total QIPP Values
CheckTot<-aggregate(T4$`TOTAL QIPP - Forecast - Plan`, by=list(T4$Scheme), sum, na.rm=true)
names(CheckTot)<-c("Type","Value")
CheckTot$Type[1]<-"Total Qipp Plan"

#STP Total RC Values
CheckRC<-aggregate(RC$`TOTAL QIPP - Forecast - Plan`, by=list(RC$Scheme), sum, na.rm=true)
names(CheckRC)<-c("Type","Value")
CheckRC$Type[1]<-"Total RC Plan"

rbind(CheckRC,CheckTot) # Values should be around 58.4 and 668.8

# checKs whether any CCGs are dropping out after merge with OrgHeir
length(unique(T3$CCG)) # Num CCGs before merge with OrgHeir
length(unique(T4$CCG)) # Num CCGs after merge with OrgHeir
setdiff(unique(T3$CCG),unique(T4$CCG)) # Any CCGs missing will show here

rm(list = c("cols.num","OrgHier","T1","T2","RC","CheckRC","CheckTot"))


# a good check is to open the "qipp_Raw_Num.xlsx" file and ensure that the correct number of 'TOTAL Unidentified' 
# records is the same as the total number of CCGs files you have loaded
# This should be the final row in every table for every CCG


# This needs to be here


