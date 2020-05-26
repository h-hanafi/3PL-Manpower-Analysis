library(tidyverse)
library(readxl)
library(googledrive)
library(openxlsx)
drive_auth(email = "beacontechstc@gmail.com")
options(digits = 2)

# Reading the Price List
PL <- read_xlsx("C:\\Users\\USER\\Desktop\\WORK CURRENT\\HRM Transfer\\Monthly Invoice\\Monthly invoice- Basic Sheet v5.xlsx", skip = 3, range = "B3:C97")
PL <- PL %>% mutate(PL = read_xlsx("C:\\Users\\USER\\Desktop\\WORK CURRENT\\HRM Transfer\\Monthly Invoice\\Monthly invoice- Basic Sheet v5.xlsx", 
                             range = cell_cols("B4:Q96")) %>% slice(3:96) %>% pull(17))
colnames(PL) <- c("Job_Group","Job_Location","PL")

#Output File Column names
CNames <- c("ID","Name","Job Group","Job Location","Take Home as per Price List",
            "Take Home as per Master Sheet","Discrepancy in Take Home (Price List - Master Sheet)"
            ,"Location Change fixes Discrepancy")
#reading 2B data
Profile_2B <- read_xlsx("C:\\Users\\USER\\Desktop\\WORK CURRENT\\HRM Transfer\\Profiles\\2B EMPLOYEE PROFILE v3.xlsx")
colnames(Profile_2B) <- str_replace_all(colnames(Profile_2B)," ","_")
Profile_2B <- Profile_2B %>% select(ID,Name_in_Arabic, Job_Group,Job_Location)

drive_download("2B Employee Master Sheet", path = file.path(getwd(),"DATA","2B Employee Master Sheet.xlsx"),overwrite = TRUE)
TH_2B <- read_xlsx(file.path(getwd(),"DATA","2B Employee Master Sheet.xlsx"), skip = 1) %>% slice(3:n()) %>% select(ID, `Take Home Pay`)
colnames(TH_2B) <- str_replace_all(colnames(TH_2B)," ","_")

#outputting discrepancies 2B
Discrepancy_2B <- Profile_2B %>% left_join(PL) %>% left_join(TH_2B)
Discrepancy_2B <- Discrepancy_2B %>% mutate(Disc = as.numeric(PL) - as.numeric(Take_Home_Pay))

Location_2B <- Discrepancy_2B %>% filter(abs(Disc) >1) %>% mutate(Job_Location_Orig = Job_Location,
                                                                        Job_Location = if_else(Job_Location == "KRT","FIELD","KRT"),
                                                                        PL_Orig = PL) %>% select(-PL) %>% left_join(PL) %>% 
  mutate(Disc_2 = as.numeric(PL) - as.numeric(Take_Home_Pay)) %>% filter(abs(Disc_2) < 1) %>% pull(ID)

Discrepancy_2B %>% filter(abs(Disc) > 1) %>% mutate(Loc_Disc = ID %in% Location_2B) %>% rename_at(colnames(.),~ CNames) %>%
  write_excel_csv("C:\\Users\\USER\\Desktop\\R modifications\\PL discrepancies 2B.csv")

#reading Bapco data
drive_download("BAPCO Employee Master sheet", path = file.path(getwd(),"DATA","BAPCO Employee Master Sheet.xlsx"), overwrite = TRUE)
TH_BAPCO <- read_xlsx(file.path(getwd(),"DATA","BAPCO Employee Master Sheet.xlsx"), skip = 1) %>% slice(3:n()) %>% select(ID, `Take Home Pay`)
colnames(TH_BAPCO) <- str_replace_all(colnames(TH_BAPCO)," ","_")

Profile_BAPCO <- read_xlsx("C:\\Users\\USER\\Desktop\\WORK CURRENT\\HRM Transfer\\Profiles\\BAPCO EMPLOYEE PROFILE v3.xlsx")
colnames(Profile_BAPCO) <- str_replace_all(colnames(Profile_BAPCO)," ","_")
Profile_BAPCO <- Profile_BAPCO %>% select(ID,Name_in_Arabic,Job_Group,Job_Location)

#outputting discrepancies BAPCO
Discrepancy_BAPCO <- Profile_BAPCO %>% left_join(PL) %>% left_join(TH_BAPCO)
Discrepancy_BAPCO <- Discrepancy_BAPCO %>% mutate(Disc = as.numeric(PL) - as.numeric(Take_Home_Pay))

Location_BAPCO <- Discrepancy_BAPCO %>% filter(abs(Disc) >1) %>% mutate(Job_Location_Orig = Job_Location,
                                                      Job_Location = if_else(Job_Location == "KRT","FIELD","KRT"),
                                                      PL_Orig = PL) %>% select(-PL) %>% left_join(PL) %>% 
  mutate(Disc_2 = as.numeric(PL) - as.numeric(Take_Home_Pay)) %>% filter(abs(Disc_2) < 1) %>% pull(ID)

Discrepancy_BAPCO %>% filter(abs(Disc) > 1) %>% mutate(Loc_Disc = ID %in% Location_BAPCO) %>% rename_at(colnames(.),~ CNames) %>%
  write_excel_csv("C:\\Users\\USER\\Desktop\\R modifications\\PL discrepancies BAPCO.csv")



# reading PETCO data
drive_download("PETCO Employee Master Sheet", path = file.path(getwd(),"DATA","PETCO Employee Master Sheet.xlsx"), overwrite = TRUE)
TH_PETCO <- read_xlsx(file.path(getwd(),"DATA","PETCO Employee Master Sheet.xlsx"), skip = 1) %>% slice(3:n()) %>% select(ID, `Take Home Pay`)
colnames(TH_PETCO) <- str_replace_all(colnames(TH_PETCO)," ","_")

Profile_PETCO <- read_xlsx("C:\\Users\\USER\\Desktop\\WORK CURRENT\\HRM Transfer\\Profiles\\PETCO EMPLOYEE PROFILE v3.xlsx")
colnames(Profile_PETCO) <- str_replace_all(colnames(Profile_PETCO)," ","_")
Profile_PETCO <- Profile_PETCO %>% select(ID,Name_in_Arabic,Job_Group,Job_Location)

#outputting discrepancies PETCO
Discrepancy_PETCO <- Profile_PETCO %>% left_join(PL) %>% left_join(TH_PETCO)
Discrepancy_PETCO <- Discrepancy_PETCO %>% mutate(Disc = as.numeric(PL) - as.numeric(Take_Home_Pay))

Location_PETCO <- Discrepancy_PETCO %>% filter(abs(Disc) >1) %>% mutate(Job_Location_Orig = Job_Location,
                                                                        Job_Location = if_else(Job_Location == "KRT","FIELD","KRT"),
                                                                        PL_Orig = PL) %>% select(-PL) %>% left_join(PL) %>% 
  mutate(Disc_2 = as.numeric(PL) - as.numeric(Take_Home_Pay)) %>% filter(abs(Disc_2) < 1) %>% pull(ID)

Discrepancy_PETCO %>% filter(abs(Disc) > 1) %>% mutate(Loc_Disc = ID %in% Location_PETCO) %>% rename_at(colnames(.),~ CNames) %>%
  write_excel_csv("C:\\Users\\USER\\Desktop\\R modifications\\PL discrepancies PETCO.csv")

# reading RPOC data
drive_download("RPOC Employee Master Sheet", path = file.path(getwd(),"DATA","RPOC Employee Master Sheet.xlsx"), overwrite = TRUE)
TH_RPOC <- read_xlsx(file.path(getwd(),"DATA","RPOC Employee Master Sheet.xlsx"), skip = 1) %>% slice(3:n()) %>% select(ID, `Take Home Pay`)
colnames(TH_RPOC) <- str_replace_all(colnames(TH_RPOC)," ","_")

Profile_RPOC <- read_xlsx("C:\\Users\\USER\\Desktop\\WORK CURRENT\\HRM Transfer\\Profiles\\RPOC EMPLOYEE PROFILE v3.xlsx")
colnames(Profile_RPOC) <- str_replace_all(colnames(Profile_RPOC)," ","_")
Profile_RPOC <- Profile_RPOC %>% select(ID,Name_in_Arabic,Job_Group,Job_Location)

#outputting discrepancies RPOC
Discrepancy_RPOC <- Profile_RPOC %>% left_join(PL) %>% left_join(TH_RPOC)
Discrepancy_RPOC <- Discrepancy_RPOC %>% mutate(Disc = as.numeric(PL) - as.numeric(Take_Home_Pay))

Location_RPOC <- Discrepancy_RPOC %>% filter(abs(Disc) >1) %>% mutate(Job_Location_Orig = Job_Location,
                                                                        Job_Location = if_else(Job_Location == "KRT","FIELD","KRT"),
                                                                        PL_Orig = PL) %>% select(-PL) %>% left_join(PL) %>% 
  mutate(Disc_2 = as.numeric(PL) - as.numeric(Take_Home_Pay)) %>% filter(abs(Disc_2) < 1) %>% pull(ID)

Discrepancy_RPOC %>% filter(abs(Disc) > 1) %>% mutate(Loc_Disc = ID %in% Location_RPOC) %>% rename_at(colnames(.),~ CNames) %>%
  write_excel_csv("C:\\Users\\USER\\Desktop\\R modifications\\PL discrepancies RPOC.csv")

# reading SHPOC data
drive_download("SHPOC Employee Master Sheet", path = file.path(getwd(),"DATA","SHPOC Employee Master Sheet.xlsx"), overwrite = TRUE)
TH_SHPOC <- read_xlsx(file.path(getwd(),"DATA","SHPOC Employee Master Sheet.xlsx"), skip = 1) %>% slice(3:n()) %>% select(ID, `Take Home Pay`)
colnames(TH_SHPOC) <- str_replace_all(colnames(TH_SHPOC)," ","_")

Profile_SHPOC <- read_xlsx("C:\\Users\\USER\\Desktop\\WORK CURRENT\\HRM Transfer\\Profiles\\SHPOC EMPLOYEE PROFILE v3.xlsx")
colnames(Profile_SHPOC) <- str_replace_all(colnames(Profile_SHPOC)," ","_")
Profile_SHPOC <- Profile_SHPOC %>% select(ID,Name_in_Arabic,Job_Group,Job_Location)

#outputting discrepancies SHPOC
Discrepancy_SHPOC <- Profile_SHPOC %>% left_join(PL) %>% left_join(TH_SHPOC)
Discrepancy_SHPOC <- Discrepancy_SHPOC %>% mutate(Disc = as.numeric(PL) - as.numeric(Take_Home_Pay))

Location_SHPOC <- Discrepancy_SHPOC %>% filter(abs(Disc) >1) %>% mutate(Job_Location_Orig = Job_Location,
                                                                      Job_Location = if_else(Job_Location == "KRT","FIELD","KRT"),
                                                                      PL_Orig = PL) %>% select(-PL) %>% left_join(PL) %>% 
  mutate(Disc_2 = as.numeric(PL) - as.numeric(Take_Home_Pay)) %>% filter(abs(Disc_2) < 1) %>% pull(ID)

Discrepancy_SHPOC %>% filter(abs(Disc) > 1) %>% mutate(Loc_Disc = ID %in% Location_SHPOC) %>% rename_at(colnames(.),~ CNames) %>%
  write_excel_csv("C:\\Users\\USER\\Desktop\\R modifications\\PL discrepancies SHPOC.csv")


#
Discrepancy_2B %>% filter(is.na(PL))
Discrepancy_BAPCO %>% filter(is.na(PL))
Discrepancy_RPOC %>% filter(is.na(PL))
Discrepancy_SHPOC %>% filter(is.na(PL))
Discrepancy_PETCO %>% filter(is.na(PL))
