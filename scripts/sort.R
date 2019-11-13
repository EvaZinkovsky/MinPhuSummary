library(tidyverse)
library(readxl)
library(lubridate)



# find the names of the sheets contained in the excel dataset        
excel_sheets("data/MinhPhuSales_09112019.xlsx")

# create a table of the internal sales for 2018
internal2018 <- read_excel("data/MinhPhuSales_09112019.xlsx", 1)

# remove all the spaces in the column names
internal2018 <- select(internal2018, "DateOfShipment" = "Date of Shipment", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")

# create a table of the internal sales for 2019
internal2019 <- read_excel("data/MinhPhuSales_09112019.xlsx", 4)

# remove all the spaces in the column names
internal2019 <- select(internal2019, "DateOfShipment" = "Date of Shipment", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")

# combine internal2018 and internal2019 into internal1819
internal1819 <- bind_rows(internal2018, internal2019)

# create a new column which has the month and year of each sale only
internal1819$ShipmentMonth <- format(as.Date(internal1819$DateOfShipment), "%Y-%m") 

# find the quantity of each product sold each month in each destination 
internalbymonth <- internal1819 %>% 
  group_by(ShipmentMonth, FirstLevelExportMarket, SecondProductCategory)%>%
  summarise(QuantityKG = sum(QuantityKG))

# find the quantity of each product sold each month in each destination 
internalbymonth2 <- internal1819 %>% 
  group_by(ShipmentMonth, SecondProductCategory, FirstLevelExportMarket)%>%
  summarise(QuantityKG = sum(QuantityKG))

# replace all occurrences of Tôm in the second product category
internalbymonth$SecondProductCategory[grepl("*Tôm\\s*", internalbymonth$SecondProductCategory)] <- "Penaeus sp"
internalbymonth2$SecondProductCategory[grepl("*Tôm\\s*", internalbymonth2$SecondProductCategory)] <- "Penaeus sp"

# Write CSV in R
write.csv(internalbymonth, file = "data/InternalByMonth.csv",row.names=FALSE)
write.csv(internalbymonth2, file = "data/InternalByMonth2.csv",row.names=FALSE)

# create a table of the domestic sales for 2018            
domestic2018 <- read_excel("data/MinhPhuSales_09112019.xlsx", 2)

# remove all the spaces in the column names
domestic2018 <- select(domestic2018, "DateOfShipment" = "Date of Shipment", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")

# create a table of the domestic sales for 2019
domestic2019 <- read_excel("data/MinhPhuSales_09112019.xlsx", 5)

# remove all the spaces in the column names
domestic2019 <- select(domestic2019, "DateOfShipment" = "Date of Shipment", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")

# combine domestic2018 and domestic2019 into domestic1819
domestic1819 <- bind_rows(domestic2018, domestic2019)

# create a new column which has the month and year of each sale only
domestic1819$ShipmentMonth <- format(as.Date(domestic1819$DateOfShipment), "%Y-%m") 

# find the quantity of each product sold each month in each destination
domesticbymonth <- domestic1819 %>% 
  group_by(ShipmentMonth, FirstLevelExportMarket, SecondProductCategory)%>%
  summarise(QuantityKG = sum(QuantityKG))

# find the quantity of each product sold each month in each destination 
domesticbymonth2 <- domestic1819 %>% 
  group_by(ShipmentMonth, SecondProductCategory, FirstLevelExportMarket)%>%
  summarise(QuantityKG = sum(QuantityKG))

# replace all occurrences of Tôm in the second product category
domesticbymonth$SecondProductCategory[grepl("*Tôm\\s*", domesticbymonth$SecondProductCategory)] <- "Penaeus sp"
domesticbymonth2$SecondProductCategory[grepl("*Tôm\\s*", domesticbymonth2$SecondProductCategory)] <- "Penaeus sp"

# Write CSV in R
write.csv(domesticbymonth, file = "data/DomesticByMonth.csv",row.names=FALSE)
write.csv(domesticbymonth2, file = "data/DomesticByMonth2.csv",row.names=FALSE)

# create a table of the export sales for 2018
export2018 <- read_excel("data/MinhPhuSales_09112019.xlsx", 3)

# remove all the spaces in the column names
export2018 <- select(export2018, "DateOfShipment" = "Date of Shipment", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")

# create a table of the export sales for 2019 
export2019 <- read_excel("data/MinhPhuSales_09112019.xlsx", 6)

# remove all the spaces in the column names
export2019 <- select(export2019, "DateOfShipment" = "Date of Shipment", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")

# combine export2018 and export2019 into export1819
export1819 <- bind_rows(export2018, export2019)

# create a new column which has the month and year of each sale only
export1819$ShipmentMonth <- format(as.Date(export1819$DateOfShipment), "%Y-%m") 

# find the quantity of each product sold each month in each destination
exportbymonth <- export1819 %>% 
  group_by(ShipmentMonth, FirstLevelExportMarket, SecondProductCategory)%>%
  summarise(QuantityKG = sum(QuantityKG))

# find the quantity of each product sold each month in each destination
exportbymonth2 <- export1819 %>% 
  group_by(ShipmentMonth, SecondProductCategory, FirstLevelExportMarket)%>%
  summarise(QuantityKG = sum(QuantityKG))

# replace all occurrences of Tôm in the second product category
exportbymonth$SecondProductCategory[grepl("*Tôm\\s*", exportbymonth$SecondProductCategory)] <- "Penaeus sp"
# replace all occurrences of Tôm in the second product category
exportbymonth2$SecondProductCategory[grepl("*Tôm\\s*", exportbymonth2$SecondProductCategory)] <- "Penaeus sp"

# Write CSV in R
write.csv(exportbymonth, file = "data/ExportByMonth.csv",row.names=FALSE)
write.csv(exportbymonth, file = "data/ExportByMonth2.csv",row.names=FALSE)