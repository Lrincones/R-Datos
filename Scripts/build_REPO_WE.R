#

#
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
library(xlsx) 
# library(janitor) did not work the date conversion do more research

#
# Directory "C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository/Sept_2020"

# For the proof of concept September 2020

HOM_WE <- read_xlsx("C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/Homere/homere_assignments.xlsx")
# HOM_WE[is.na(HOM_WE)] <- "Empty" # change NA to empty
HOM_WE$Pos <- "Empty"

OCG_IRFFG <- read_xlsx("C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository/Sept_2020/OCG_IRFFG.xlsx")

# Changing the NA values to "Empty"

OCG_IRFFG = OCG_IRFFG %>% replace_na(list(Code_MP="Empty",Code_LS = "Empty", Code_HF = "Empty", Code_O ="Empty", `Medical_&_Paramedical`= "Empty",
                                          `Logistics_&_Supply`= "Empty",`HR_&_FIN`= "Empty", Operations = "Empty"))


for(i in 1:nrow(HOM_WE)){
  for(j in 1: nrow(OCG_IRFFG)){
    if(HOM_WE$fonction[i] == OCG_IRFFG$Code_MP[j]){
      HOM_WE$Pos[i] = OCG_IRFFG$`Medical_&_Paramedical`[j]    }
    if(HOM_WE$fonction[i] == OCG_IRFFG$Code_LS[j]){
      HOM_WE$Pos[i] = OCG_IRFFG$`Logistics_&_Supply`[j]    }
    if(HOM_WE$fonction[i] == OCG_IRFFG$Code_HF[j]){
      HOM_WE$Pos[i] = OCG_IRFFG$`HR_&_FIN`[j]    }
    if(HOM_WE$fonction[i] == OCG_IRFFG$Code_O[j]){
      HOM_WE$Pos[i] = OCG_IRFFG$Operations[j]    }
  }
}

OCG_WE <- read_xlsx("C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository/Sept_2020/OCGDynamicsWorkingExperience.xlsx")
# OCG_WE[is.na(OCG_WE[,c(1,2,4:15)])] <- "Empty" # change NA to empty
# Update the format date for date of Scaling
OCG_WE[,3] <-   as.Date(as.numeric(unlist(OCG_WE[,3])), origin = "1899-12-30" )



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

# To update after issue with date formats is solve las fechas de Homere resolver formato 
# Start_date = as_datetime(character()),
# End_date = as_datetime(character()),


for(i in 1:nrow(HOM_WE)){
  REPO_WE[i,1] = "Homere"
  REPO_WE[i,2] = HOM_WE[i,1] # Field LegacyPersonnelnumber
  REPO_WE[i,3] = HOM_WE[i,5] # Field Start Date
  REPO_WE[i,4] = HOM_WE[i,6] # Field End Date
  REPO_WE[i,6] = "MSF" # Field Company
  REPO_WE[i,7] = HOM_WE[i,2] # Field IRFFG
  REPO_WE[i,9] = HOM_WE[i,12] # Field Position
}
# REPO_WE$Start_date  <-  convertToDate(REPO_WE$Start_date)
# REPO_WE$End_date  <-  convertToDate(REPO_WE$End_date)
 
#
# VE SI SE PUEDE HACER CON UN LAZO EMPEZANDO DESDE EL ULTIMO REPOID EN VEZ DE ESTE ENREDO SIMILAR A LO QUE HICISTE EN REPO-SKILLS FOR HOMERE
#

# Filas en REPO_WE se controlan por estas variables en este caso OCG solo se proceso una oficina antes
low_A = nrow(HOM_WE) +1
top_A = low_A + nrow(OCG_WE) -1
# filas en OCG se controlan dentro del indice


for(i in low_A : top_A ){
  REPO_WE[i,1] = "OCG"
  REPO_WE[i,2] =OCG_WE[i-nrow(HOM_WE),1] # Field Personnel_number
  REPO_WE[i,3] =OCG_WE[i-nrow(HOM_WE),15] # Field Start_date
  REPO_WE[i,4] =OCG_WE[i-nrow(HOM_WE),7] # Field End_date
  REPO_WE[i,5] =OCG_WE[i-nrow(HOM_WE),2] # Field Country/region
  REPO_WE[i,6] =OCG_WE[i-nrow(HOM_WE),5] # Field Employer
  REPO_WE[i,7] =OCG_WE[i-nrow(HOM_WE),12]# Field Metier
  REPO_WE[i,8] =OCG_WE[i-nrow(HOM_WE),10]# Field Job_family
  REPO_WE[i,9] =OCG_WE[i-nrow(HOM_WE),13] # Field Position
  
}


# REPO_WE[is.na(REPO_WE[,])] <- "Empty" # change NA to empty

# OCG_WE[is.na(OCG_WE[,c(1,2,4:15)])] <- "Empty"


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
  if (!is.na(REPO_WE$IRFFG[i])) {
    fchar = substr(REPO_WE$IRFFG[i], 1, 1)
      if (fchar == "M") {
        REPO_WE$Job_family[i] = "Medical_&_Paramedical"
        # fchar = " "
      }
      if (fchar == "L") {
        REPO_WE$Job_family[i] = "Logistics_&_Supply"
        # fchar = " "
      }
      if (fchar == "A") {
        REPO_WE$Job_family[i] = "HR_&_Fin"
        # fchar = " "
      }
      if (fchar == "O") {
        REPO_WE$Job_family[i] = "Operations"
        # fchar = " "
      }
    }
  fchar = " "
}

# populating the columns Professional_Group and Level
# depends on the IRFFG
# for all the offices loaded

for (i in 1:nrow(REPO_WE)) {
  IRFFG_C = REPO_WE$IRFFG[i]
  Job_FA = REPO_WE$Job_family[i]
  if (!is.na(Job_FA)) {
    if (Job_FA == "Operations") {
      for (j in 1:nrow(OCG_IRFFG)) {
        if (IRFFG_C == OCG_IRFFG$Code_O[j]) {
          REPO_WE$Professional_Group[i] = OCG_IRFFG$Professional_Group[j]
          REPO_WE$Level[i] = OCG_IRFFG$Level[j]
        }
      }
    }
    
    if (Job_FA == "Medical_&_Paramedical") {
      for (j in 1:nrow(OCG_IRFFG)) {
        if (IRFFG_C == OCG_IRFFG$Code_MP[j]) {
          REPO_WE$Professional_Group[i] = OCG_IRFFG$Professional_Group[j]
          REPO_WE$Level[i] = OCG_IRFFG$Level[j]
        }
      }
    }
    
    if (Job_FA == "HR_&_Fin") {
      for (j in 1:nrow(OCG_IRFFG)) {
        if (IRFFG_C == OCG_IRFFG$Code_HF[j]) {
          REPO_WE$Professional_Group[i] = OCG_IRFFG$Professional_Group[j]
          REPO_WE$Level[i] = OCG_IRFFG$Level[j]
        }
      }
    }
    
    if (Job_FA == "Logistics_&_Supply") {
      for (j in 1:nrow(OCG_IRFFG)) {
        if (IRFFG_C == OCG_IRFFG$Code_LS[j]) {
          REPO_WE$Professional_Group[i] = OCG_IRFFG$Professional_Group[j]
          REPO_WE$Level[i] = OCG_IRFFG$Level[j]
        }
      }
    }
  }
}




# save REPO_WE for script Match_REPO_WE-POS_RE 
write_xlsx(REPO_WE,"C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository/Sept_2020/REPO_WE.xlsx")
