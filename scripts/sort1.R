library(tidyverse)
library(readxl)
library(lubridate)



# find the names of the sheets contained in the excel dataset        
excel_sheets("data/MinhPhuSales_09112019.xlsx")

internal2018 <- read_excel("data/MinhPhuSales_09112019.xlsx", 1)
internal2018 <- select(internal2018, "DateOfShipment" = "Date of Shipment", "Destination", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")
internal2019 <- read_excel("data/MinhPhuSales_09112019.xlsx", 4)
internal2019 <- select(internal2019, "DateOfShipment" = "Date of Shipment", "Destination", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")
internal1819 <- bind_rows(internal2018, internal2019)

internal1819$ShipmentMonth <- format(as.Date(internal1819$DateOfShipment), "%Y-%m") 

internalbymonth <- internal1819 %>% 
  group_by(ShipmentMonth, Destination, SecondProductCategory)%>%
  summarise(QuantityKG = sum(QuantityKG))
             
domestic2018 <- read_excel("data/MinhPhuSales_09112019.xlsx", 2)
domestic2018 <- select(domestic2018, "DateOfShipment" = "Date of Shipment", "Destination", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")
domestic2019 <- read_excel("data/MinhPhuSales_09112019.xlsx", 5)
domestic2019 <- select(domestic2019, "DateOfShipment" = "Date of Shipment", "Destination", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")
domestic1819 <- bind_rows(domestic2018, domestic2019)

domestic1819$ShipmentMonth <- format(as.Date(domestic1819$DateOfShipment), "%Y-%m") 

domesticbymonth <- domestic1819 %>% 
  group_by(ShipmentMonth, Destination, SecondProductCategory)%>%
  summarise(QuantityKG = sum(QuantityKG))

export2018 <- read_excel("data/MinhPhuSales_09112019.xlsx", 3)
export2018 <- select(export2018, "DateOfShipment" = "Date of Shipment", "Destination", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")
export2019 <- read_excel("data/MinhPhuSales_09112019.xlsx", 6)
export2019 <- select(export2019, "DateOfShipment" = "Date of Shipment", "Destination", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")
export1819 <- bind_rows(export2018, export2019)

export1819$ShipmentMonth <- format(as.Date(export1819$DateOfShipment), "%Y-%m") 

exportbymonth <- internal1819 %>% 
  group_by(ShipmentMonth, Destination, SecondProductCategory)%>%
  summarise(QuantityKG = sum(QuantityKG))

