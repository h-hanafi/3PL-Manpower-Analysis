library(tidyverse)
library(readxl)


read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 1.xlsx",
           sheet = "INS",col_types = rep("text",7),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C","Comment"),skip = 1) %>%
  mutate(Comment_2 = NA)

read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 2.xlsx",
           sheet = "INS",col_types = rep("text",7),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C","Comment"), skip = 1) %>% 
  mutate(Comment_2 = NA)

read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 3.xlsx",
           sheet = "INS",col_types = rep("text",6),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C"),skip = 1) %>% 
  mutate(Comment = NA, Comment_2 = NA)
  

read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 4.xlsx",
           sheet = "INS",col_types = rep("text",6),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C"),skip = 1) %>% 
  mutate(Comment = NA, Comment_2 = NA)

read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 5.xlsx",
           sheet = "INS",col_types = rep("text",6),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C"), skip = 1) %>% 
  mutate(Comment = NA, Comment_2 = NA)
  
read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 6.xlsx",
           sheet = "INS",col_types = rep("text",8),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C","Comment","Comment_2"),skip = 1)

read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 7.xlsx",
           sheet = "INS",col_types = rep("text",7),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C","Comment"), skip = 1) %>%
  mutate(Comment_2 = NA)

read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 8.xlsx",
           sheet = "INS",col_types = rep("text",10),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C","Comment","Comment_2","",""), skip = 1) %>% 
  select(1:8)

#############
INS_FAM <- 
  read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 1.xlsx",
             sheet = "INS",col_types = rep("text",7),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C","Comment"),skip = 1) %>%
  mutate(Comment_2 = NA) %>%
  rbind(
    read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 2.xlsx",
               sheet = "INS",col_types = rep("text",7),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C","Comment"), skip = 1) %>% 
      mutate(Comment_2 = NA)) %>%
  rbind(
    read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 3.xlsx",
               sheet = "INS",col_types = rep("text",6),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C"),skip = 1) %>% 
      mutate(Comment = NA, Comment_2 = NA)) %>%
  rbind(
    read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 4.xlsx",
           sheet = "INS",col_types = rep("text",6),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C"),skip = 1) %>% 
  mutate(Comment = NA, Comment_2 = NA)) %>%
  rbind(
    read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 5.xlsx",
               sheet = "INS",col_types = rep("text",6),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C"), skip = 1) %>% 
      mutate(Comment = NA, Comment_2 = NA)) %>%
  rbind(
    read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 6.xlsx",
           sheet = "INS",col_types = rep("text",8),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C","Comment","Comment_2"),skip = 1)) %>%
  rbind(
    read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 7.xlsx",
           sheet = "INS",col_types = rep("text",7),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C","Comment"), skip = 1) %>%
  mutate(Comment_2 = NA)) %>%
  rbind(
    read_excel(path = "C:\\Users\\USER\\Desktop\\WORK CURRENT\\Insurance - THE REINSURING\\Family Invoice Check\\Incoming - Family - Check - 8.xlsx",
           sheet = "INS",col_types = rep("text",10),col_names = c("Card_No","Name","Relation_O","Batch","ID","Relation_C","Comment","Comment_2","",""), skip = 1) %>% 
  select(1:8))

INS_FAM <- INS_FAM %>% fill(ID)
INS_FAM <- INS_FAM %>% mutate(OPCO = if_else(str_detect(ID,"2B"),"2B",
                                             if_else(str_detect(ID,"BAP"),"BAPCO",
                                                     if_else(str_detect(ID,"PET"),"PETCO",
                                                             if_else(str_detect(ID,"RPC"),"RPOC",
                                                                     if_else(str_detect(ID,"SHP"), "SHARIF",
                                                                             if_else(str_detect(ID,"SHL"),"SHPOC-L","Unidintified")))))))
# Analysis
INS_FAM %>% distinct(Card_No, .keep_all = T)

DUP_CARD <- INS_FAM %>% group_by(Card_No) %>% summarise(n = n()) %>% filter(n > 1) %>% pull(Card_No)

DUP <- INS_FAM %>% distinct(Card_No, .keep_all = T) %>% 
  filter(Relation_C == "Correct Primary") %>% group_by(ID) %>% summarise(n = n()) %>% filter(n > 1) %>% pull(ID)

INS_FAM %>% filter(Card_No %in% DUP_CARD) %>% write_excel_csv("C:\\Users\\USER\\Desktop\\R modifications\\MMI Duplicate Card no.csv")


INS_FAM %>% 
  filter(Relation_C == "Correct Primary") %>% filter(ID %in% DUP) %>% arrange(Batch,ID) %>% 
  write_excel_csv("C:\\Users\\USER\\Desktop\\R modifications\\MMI Duplicated Primaries.csv")

INS_FAM %>% filter(!is.na(Comment)) %>%  write_excel_csv("C:\\Users\\USER\\Desktop\\R modifications\\MMI Incorrect Relations.csv")

INS_FAM %>% filter(ID == "X") %>% write_excel_csv("C:\\Users\\USER\\Desktop\\R modifications\\MMI Unkown Issues.csv")

INS_FAM %>% group_by(OPCO,Relation_C) %>% summarise(Count = n()) %>% write_excel_csv("C:\\Users\\USER\\Desktop\\R modifications\\MMI Relation Count.csv")

INS_FAM %>% distinct(Relation_C)
INS_FAM %>% distinct(Comment) %>% pull(Comment)
INS_FAM %>% mutate(Relation_C = if_else(str_detect(Relation_C,"Primary"),"Primary",Relation_C)) %>% 
  group_by(Card_No,ID) %>% summarize(n = n()) %>% filter(n >1)
  group_by(ID) %>% summarise(n = n()) %>% filter(n > 1) %>% pull(ID)
INS_FAM %>% filter(ID %in% DUP)

INS_FAM %>% filter(ID %in% c("BAP0524","PET0324","2B00670","SHP0080","2B00328","2B00716")) %>%  mutate(Relation_C = if_else(str_detect(Relation_C,"Primary"),"Primary",Relation_C)) %>% 
  filter(Relation_C == "Primary") 

DUP_PRIMARIES_FIXED <- INS_FAM %>% 
  filter(Relation_C == "Correct Primary") %>% filter(ID %in% DUP) %>% arrange(ID)

INS_FAM %>% filter(ID == "X", str_detect(Relation_C,"Primary")

                   