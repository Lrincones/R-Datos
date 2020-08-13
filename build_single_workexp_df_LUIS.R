library(flexdashboard)
library(readxl)
library(tidyverse)
library(plotly)
library(ggmap)
library(png)
library(knitr)
library(ggplot2)
library(hrbrthemes)
library(gtable)
library(gt)
library(glue)
library(gtsummary)
library(knitr)
library(magick)
library(rvest)
library(stringr)
library(lubridate)
library(writexl)
#
# to search for the characters [], inside them is the IRFFG code. This for Italy and extract the code
# (?<=\[).*(?=\])  
#
# a = str_extract(string = b, pattern = "\\[(?<=).*(?=])")
# first make b the text from ITA_WE[i,9]
# then applied the str_extract but did not how to start after [
# will delete using  a = str_remove(a, "\\[")
#
# Leer work experience cAN IT y OCG armar una sola tabla
#
# The first character of the code indicates the job family
# the next version may incorporate the Professional Group and the level
#
#
#
# Directory "C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository"
setwd("C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository")

# For the proof of concept

CAN_WE <- read_xlsx("CANHeroWorkExperience.xlsx") 

ITA_WE <- read_xlsx("ITHeroWorkExperience.xlsx")

OCG_WE <- read_xlsx("OCGDynamicsWorkExperience.xlsx")

# OCG_HOM_WE_PW <- read_xlsx("OCGHomerePast_Work_Experience.xlsx") # OCG-Homere WE Past WE no MSF
# 
# OCG_HOM_WE_MSF <- read_xlsx("OCGHomereMSFExperience.xlsx")

OCG_IRFFG <- read_xlsx("OCG_IRFFG.xlsx")

# Changing the NA values to "Empty"


OCG_IRFFG = OCG_IRFFG %>% replace_na(list(Code_MP="Empty",Code_LS = "Empty", Code_HF = "Empty", Code_O ="Empty", `Medical_&_Paramedical`= "Empty",
                               `Logistics_&_Supply`= "Empty",`HR_&_FIN`= "Empty", Operations = "Empty"))




# Build a unique table with all Work Experience from Repository


REPO_WE <- data.frame(Office = character(),
                      Legacy_ID = character(),
                      Start_date = as_datetime(character()),
                      End_date = as_datetime(character()),
                      Country= character(), 
                      Company= character(),
                      IRFFG = character(),
                      Job_family = character(),
                      Position = character(),
                      Professional_Group = character(),
                      Level = character(),
                      stringsAsFactors=FALSE)


for(i in 1:nrow(ITA_WE)){
  REPO_WE[i,1] = "Italy"
  REPO_WE[i,2] = ITA_WE[i,3] # Field LegacyPersonnelnumber
  REPO_WE[i,3] = ITA_WE[i,1] # Field Start Date
  REPO_WE[i,4] = ITA_WE[i,2] # Field End Date
  REPO_WE[i,5] = ITA_WE[i,4] # Field Country Key
  REPO_WE[i,6] = ITA_WE[i,7] # Field Company
  REPO_WE[i,7] = ITA_WE[i,9] # Field Description
  REPO_WE[i,8] = ITA_WE[i,4] # Field populate with country to avoid NA but thr value is generated blow based on the IRFFG code
  REPO_WE[i,9] = ITA_WE[i,6] # Field Position
}

#
# Extracting the code for column IRFFG
#
for(i in 1 :nrow(ITA_WE)){
  b = REPO_WE[i, 7]
  a = str_extract(string = b, pattern = "\\[(?<=).*(?=])")
  REPO_WE[i,7] = str_remove(a, "\\[")
}
 
#
# VE SI SE PUEDE HACER CON UN LAZO EMPEZANDO DESDE EL ULTIMO REPOID EN VEZ DE ESTE ENREDO SIMILAR A LO QUE HICISTE EN REPO-SKILLS FOR HOMERE
#

# Filas en REPO_WE se controlan por estas variables en este caso OCG solo se proceso una oficina antes
low_A = nrow(ITA_WE) +1
top_A = low_A + nrow(OCG_WE) -1
# filas en OCG se controlan dentro del indice


for(i in low_A : top_A ){
  REPO_WE[i,1] = "Switzerland"
  REPO_WE[i,2] =OCG_WE[i-nrow(ITA_WE),1] # Field Personnel_number
  REPO_WE[i,3] =OCG_WE[i-nrow(ITA_WE),4] # Field Start_date
  REPO_WE[i,4] =OCG_WE[i-nrow(ITA_WE),5] # Field End_date
  REPO_WE[i,5] =OCG_WE[i-nrow(ITA_WE),9] # Field Country/region
  REPO_WE[i,6] =OCG_WE[i-nrow(ITA_WE),6] # Field Employer
  REPO_WE[i,7] =OCG_WE[i-nrow(ITA_WE),11]# Field Metier
  REPO_WE[i,8] =OCG_WE[i-nrow(ITA_WE),12]# Field Job_family
  REPO_WE[i,9] =OCG_WE[i-nrow(ITA_WE),7] # Field Position
  
}

# Filas en REPO_WE por estas variables en este caso Canada se procesaron dos oficinas antes
low_A = nrow(ITA_WE) + nrow(OCG_WE) + 1
top_A = low_A + nrow(CAN_WE) -1
# filas en CAN se controlan dentro del indice

for(i in low_A : top_A ){
  REPO_WE[i,1] = "Canada"
  REPO_WE[i,2] =CAN_WE[i-low_A+1,1]  # Field ID
  REPO_WE[i,3] =CAN_WE[i-low_A+1,11] # Field Start date
  REPO_WE[i,4] =CAN_WE[i-low_A+1,5]  # Field End date
  REPO_WE[i,5] =CAN_WE[i-low_A+1,3]  # Field Country
  REPO_WE[i,6] =CAN_WE[i-low_A+1,4]  # Field Employer name
  REPO_WE[i,7] =CAN_WE[i-low_A+1,7]  # Field Matching IRP2 Code
  #REPO_WE[i,8] =CAN_WE[i-low_A+1,7]  # Field Matching IRP2 Code
  REPO_WE[i,9] =CAN_WE[i-low_A+1,10] # Field Position
  
}

#
# working REPO_WE and IRFFG adding informatio to REPO_WE from IRFFG
# for all the offices loaded
#

# Populating the  column Job_family as per IRFFG 
# check for NA switzerland OCG has problem with IRFFG Empty check if it is empty and skip and no process it
# clean NA# for code consistency use this if possible
# df = dfG %>% replace_na(list(column ="Empty")) in the DF for a specific column
# REPO_WE[is.na(REPO_WE)] <- "Empty" # Try with one word or that doesnot get in the jobfamily or just blank as it is now
#


for (i in 1:nrow(REPO_WE)) {
  fchar = substr(REPO_WE$IRFFG[i], 1,1 )
  if (fchar == "M"){
    REPO_WE$Job_family[i] = "Medical_&_Paramedical"
    fchar = " "
  }
  if (fchar == "L"){
    REPO_WE$Job_family[i] = "Logistics_&_Supply"
    fchar = " "
  }
  if (fchar == "A"){
    REPO_WE$Job_family[i] = "HR_&_Fin"
    fchar = " "
  }
  if (fchar == "O"){
    REPO_WE$Job_family[i] = "Operations"
    fchar = " "
  }
}

# populating the columns Professional_Group and Level
# depends on the IRFFG
# for all the offices loaded

for (i in 1:nrow(REPO_WE)) {
  IRFFG_C = REPO_WE$IRFFG[i]
  Job_FA = REPO_WE$Job_family[i]
  if (Job_FA == "Operations") {
    for (j in 1:nrow(OCG_IRFFG)) {
      if (IRFFG_C == OCG_IRFFG$Code_O[j]){
        REPO_WE$Professional_Group[i] = OCG_IRFFG$Professional_Group[j]
        REPO_WE$Level[i] = OCG_IRFFG$Level[j]
      }
    }
  }
  
  if (Job_FA == "Medical_&_Paramedical") {
    for (j in 1:nrow(OCG_IRFFG)) {
      if (IRFFG_C == OCG_IRFFG$Code_MP[j]){
        REPO_WE$Professional_Group[i] = OCG_IRFFG$Professional_Group[j]
        REPO_WE$Level[i] = OCG_IRFFG$Level[j]
      }
    }
  }
  
  if (Job_FA == "HR_&_Fin") {
    for (j in 1:nrow(OCG_IRFFG)) {
      if (IRFFG_C == OCG_IRFFG$Code_HF[j]){
        REPO_WE$Professional_Group[i] = OCG_IRFFG$Professional_Group[j]
        REPO_WE$Level[i] = OCG_IRFFG$Level[j]
      }
    }
  }
  
  if (Job_FA == "Logistics_&_Supply") {
    for (j in 1:nrow(OCG_IRFFG)) {
      if (IRFFG_C == OCG_IRFFG$Code_LS[j]){
        REPO_WE$Professional_Group[i] = OCG_IRFFG$Professional_Group[j]
        REPO_WE$Level[i] = OCG_IRFFG$Level[j]
      }
    }
  }
  
}

# El Codigo a continuacion es prueba de cross tabulacion se comentara para la prueba de concepto

# Prueba de Cross Tabs entre Oficina y IRFFG
# REPO_WE_TAB <- table(REPO_WE$Office, REPO_WE$IRFFG)
# margin.table(REPO_WE_TAB,1) # Row Marginal
# margin.table(REPO_WE_TAB,2) # Column Marginal
# round(prop.table(REPO_WE_TAB),2)# Cell %
# round(prop.table(REPO_WE_TAB,1),2)# Row %
# round(prop.table(REPO_WE_TAB,2),2) # Column%
# chisq.test(REPO_WE_TAB)

# Prueba de Cross Tabs entre Oficina y LEVEL
# REPO_WE_TAB <- table(REPO_WE$Office, REPO_WE$Level)
# margin.table(REPO_WE_TAB,1) # Row Marginal
# margin.table(REPO_WE_TAB,2) # Column Marginal
# round(prop.table(REPO_WE_TAB),2)# Cell %
# round(prop.table(REPO_WE_TAB,1),2)# Row %
# round(prop.table(REPO_WE_TAB,2),2) # Column%
# chisq.test(REPO_WE_TAB)

# Prueba de Cross Tabs entre Oficina y Job Family
# REPO_WE_TAB <- table(REPO_WE$Office, REPO_WE$Job_family)
# margin.table(REPO_WE_TAB,1) # Row Marginal
# margin.table(REPO_WE_TAB,2) # Column Marginal
# round(prop.table(REPO_WE_TAB),2)# Cell %
# round(prop.table(REPO_WE_TAB,1),2)# Row %
# round(prop.table(REPO_WE_TAB,2),2) # Column%
# chisq.test(REPO_WE_TAB)

# Se ve cierta relacion entre Oficinas y Job Family pero es fake data


# save REPO_WE for future usage Formatos en Excel y RDS la lectura de prueba se comenta

saveRDS(REPO_WE, "REPO_WE.rds")

# REPO_WE_1 <- readRDS("REPO_WE.rds")

write_xlsx(REPO_WE,"REPO_WE.xlsx")
