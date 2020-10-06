# Match_REPO_WE
# Takes the files REPO_WE and POS_REQ.
# POS_REQ is prepare with the IRFFG codes that the user wants to check 
# REPO_WE (Merged Homere and OCG files IRFFG work experiences codes)
# the output is in REPO_POS_POS1
# Other output files are REPO_JF_JF1, REPO_LEV_LEV1 and REPO_PG_PG1. matching  Job Family, Level and Professional Group 
# # Libraries
library(readxl)
library(writexl)
library(tidyr)
library(tibble)
library(dplyr) # mensajes
library(stringr)
library(tidyverse) # not here but in the code 
library(xlsx)
library(hablar)
library(lubridate)
library(hablar)

# Directory "C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository/Sept_2020"

# Read Work Experience and required positions
REPO_WE <- read_excel("C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository/Sept_2020/REPO_WE.xlsx")

REPO_WE <- REPO_WE[,c(1,5:11,2)] # took off dates columns, move column to 2 to last
# POS_REQ does not have empty value
POS_REQ <- read_excel("C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository/Sept_2020/POS_REQ.xlsx") 
POS_REQ[is.na(POS_REQ)] <- "Empty" # change NA to empty
# Getting the records per job family, Level, Proffesional_Group and Position from POS_REQ(The Project Needs)
JF1 <- unique(POS_REQ$Job_family) 
JF1 <- subset(JF1, JF1 != "Empty") 
LEV1 <- unique(POS_REQ$Level)
LEV1 <- subset(LEV1, LEV1 != "Empty")
POS1 <- unique(POS_REQ$Position)
POS1 <- subset(POS1, POS1 != "Empty")
PRG1 <- unique(POS_REQ$Professional_Group)
PRG1 <- subset(PRG1 != "Empty")
# Seeking the above requirements in REPO_WE, the file with the Merged IRFFG Work Experiences
REPO_JF <- subset(REPO_WE, REPO_WE$Job_family != "Empty")
REPO_JF_JF1 <- subset(REPO_WE, REPO_WE$Job_family %in% JF1) # Job Families
REPO_PG <- subset(REPO_WE, REPO_WE$Professional_Group != "Empty")
REPO_PG_PG1 <- subset(REPO_WE, REPO_WE$Professional_Group %in% PRG1) # Proffesional_Group
REPO_LEV <- subset(REPO_WE, REPO_WE$Level != "Empty")
REPO_LEV_LEV1 <- subset(REPO_WE, REPO_WE$Level %in% LEV1) # Levels
REPO_POS <- subset(REPO_WE, REPO_WE$Position != "Empty")
REPO_POS_POS1 <-  subset(REPO_WE, REPO_WE$Position %in% POS1) # Positions

# REPO_JF_JF1 the records that match job family from POS_REQ
# REPO_LEV_LEV1 the records that match Level from POS_REQ
# REPO_POS_POS1 the records that match Position from POS_REQ
# REPO_PG_PG the records that match Professional Group from POS_REQ 
# POS_REQ , Positions in CAPITAL as IRFFG proxima version lower or upper case 
# 

write_xlsx(REPO_JF_JF1,"C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository/Sept_2020/REPO_JF_JF1.xlsx")
write_xlsx(REPO_LEV_LEV1,"C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository/Sept_2020/REPO_LEV_LEV1.xlsx")
write_xlsx(REPO_PG_PG1,"C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository/Sept_2020/REPO_PG_PG1.xlsx")
write_xlsx(REPO_POS_POS1,"C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository/Sept_2020/REPO_POS_POS1.xlsx")
