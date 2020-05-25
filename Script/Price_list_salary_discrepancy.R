library(tidyverse)
library(readxl)
library(googledrive)
library(openxlsx)
drive_auth(email = "beacontechstc@gmail.com")
options(digits = 2)

# Reading the Price List
PL <- read_xlsx("C:\\Users\\USER\\Desktop\\WORK CURRENT\\HRM Transfer\\Monthly Invoice\\Monthly invoice- Basic Sheet v5.xlsx", skip = 3, range = "B3:C96")
PL <- PL %>% mutate(PL = read_xlsx("C:\\Users\\USER\\Desktop\\WORK CURRENT\\HRM Transfer\\Monthly Invoice\\Monthly invoice- Basic Sheet v5.xlsx", 
                             range = cell_cols("B4:Q96")) %>% slice(3:95) %>% pull(17))
colnames(PL) <- c("Job_Group","Job_Location","PL")
PL <- PL %>% mutate(Helper = paste(trimws(Job_Group),Job_Location,sep = "-"))

#reading 2B data
Profile_2B <- read_xlsx("C:\\Users\\USER\\Desktop\\WORK CURRENT\\HRM Transfer\\Profiles\\2B EMPLOYEE PROFILE v3.xlsx")
colnames(Profile_2B) <- str_replace_all(colnames(Profile_2B)," ","_")
Profile_2B <- Profile_2B %>% select(ID,Job_Group,Job_Location)
Profile_2B <- Profile_2B %>% mutate(Helper = paste(Job_Group,Job_Location,sep = "-"))

drive_download("2B Employee Master Sheet", path = file.path(getwd(),"DATA","2B Employee Master Sheet.xlsx"),overwrite = TRUE)
TH_2B <- read_xlsx(file.path(getwd(),"DATA","2B Employee Master Sheet.xlsx"), skip = 1) %>% slice(3:n()) %>% select(ID, `Take Home Pay`)
colnames(TH_2B) <- str_replace_all(colnames(TH_2B)," ","_")

#outputting discrepancies 2B
Discrepancy_2B <- Profile_2B %>% left_join(PL) %>% left_join(TH_2B)

Discrepancy_2B %>% mutate(Disc = as.numeric(PL) - as.numeric(Take_Home_Pay)) %>% filter(abs(Disc) > 1) %>% write_csv("C:\\Users\\USER\\Desktop\\R modifications\\PL discrepancies 2B.csv")
Discrepancy_2B %>% mutate(Disc = as.numeric(PL) - as.numeric(Take_Home_Pay)) %>% filter(abs(Disc) > 1) %>% unique()


#reading Bapco data
drive_download("BAPCO Employee Master sheet", path = file.path(getwd(),"DATA","BAPCO Employee Master Sheet.xlsx"), overwrite = TRUE)
TH_BAPCO <- read_xlsx(file.path(getwd(),"DATA","BAPCO Employee Master Sheet.xlsx"), skip = 1) %>% slice(3:n()) %>% select(ID, `Take Home Pay`)
colnames(TH_BAPCO) <- str_replace_all(colnames(TH_2B)," ","_")

Profile_BAPCO <- read_xlsx("C:\\Users\\USER\\Desktop\\WORK CURRENT\\HRM Transfer\\Profiles\\BAPCO EMPLOYEE PROFILE v3.xlsx")
colnames(Profile_BAPCO) <- str_replace_all(colnames(Profile_BAPCO)," ","_")
Profile_BAPCO <- Profile_BAPCO %>% select(ID,Job_Group,Job_Location)
Profile_BAPCO <- Profile_BAPCO %>% mutate(Helper = paste(trimws(Job_Group),Job_Location,sep = "-"))

#outputting discrepancies BAPCO
Discrepancy_BAPCO <- Profile_BAPCO %>% left_join(PL) %>% left_join(TH_BAPCO)
Discrepancy_BAPCO %>% mutate(Disc = as.numeric(PL) - as.numeric(Take_Home_Pay)) %>% filter(abs(Disc) > 1) %>% write_csv("C:\\Users\\USER\\Desktop\\R modifications\\PL discrepancies BAPCO.csv")
