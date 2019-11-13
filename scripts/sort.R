library(tidyverse)
library(readxl)
library(lubridate)



# find the names of the sheets contained in the excel dataset        
excel_sheets("data/MinhPhuSales_09112019.xlsx")

# create a table of the internal sales for 2018
internal2018 <- read_excel("data/MinhPhuSales_09112019.xlsx", 1)

# remove all the spaces in the column names
internal2018 <- select(internal2018, "DateOfShipment" = "Date of Shipment", "Destination", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")

# create a table of the internal sales for 2019
internal2019 <- read_excel("data/MinhPhuSales_09112019.xlsx", 4)

# remove all the spaces in the column names
internal2019 <- select(internal2019, "DateOfShipment" = "Date of Shipment", "Destination", "SecondProductCategory" = "2nd product category", "ThirdProductCategory" = "3rd product category", "QuantityPackages" = "Quantity (Packages)", "QuantityKG" = "Quanity (KG)","Plant", "PlantName" = "Plant Name", "Certificate", "FirstLevelExportMarket" = "1st level export market", "SecondLevelExportMarket" = "2nd level export market")

# combine internal2018 and internal2019 into internal1819
internal1819 <- bind_rows(internal2018, internal2019)

# create a new column which has the month and year of each sale onle
internal1819$ShipmentMonth <- format(as.Date(internal1819$DateOfShipment), "%Y-%m") 

# find the quantity of each product sold each month in each destination 
internalbymonth <- internal1819 %>% 
  group_by(ShipmentMonth, Destination, SecondProductCategory)%>%
  summarise(QuantityKG = sum(QuantityKG))
             

