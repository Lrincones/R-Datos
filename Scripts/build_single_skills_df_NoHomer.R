#
# Revisar el codigo para eliminar lo de Homere
# solo OCs, PS, Branchs files
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
library(writexl)
library(xlsx)
#
# Directory "C:/Users/luis.rincones/OneDrive - MSF/Documents/0_OM_Repository"
setwd("C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/Output")
#
# Leer skills IT y OCG armar una sola tabla Only Italy and OCG have data for skills
#
ITA_SKILLs <- read_xlsx("ITHeroSkills.xlsx")

OCG_SKILLS <- read_xlsx("OCGDynamicsSkills.xlsx")

OCG_Homere_Skills <- read_xlsx("OCGHomereSkills.xlsx")

# Build a unique table with all skills from Repository this table is a draft based in the availble data needs more work

REPO_SKILLS <- data.frame(Office = character(),
                          Legacy_ID =character(),
                          Skill=character(), 
                          Scale_ID=character(),
                          Name_Scale = character(),
                          stringsAsFactors=FALSE)

# Mapping for Italy need to document it the excel file 
# "C:\Users\luis.rincones\OneDrive - MSF\Documents\0_OM_Repository\References\Mapping_Repository_Souk_Maps.xlsx"
# has the basics of the mapping need to scale it for more flies and code the mapping using a mapping input

for(i in 1:nrow(ITA_SKILLs)){
  REPO_SKILLS[i,1] = "Italy"
  REPO_SKILLS[i,2] = ITA_SKILLs[i,3] # Field LegacyPersonnelnumber
  REPO_SKILLS[i,3] = ITA_SKILLs[i,4] # Field Qualification_Name
  REPO_SKILLS[i,4] = ITA_SKILLs[i,5] # Field Scale_ID
  REPO_SKILLS[i,5] = ITA_SKILLs[i,6] # Field Name_of_the_Scale_ID
}

# Filas en Repo_Skills se controlan por estas variables only two offices 
# Italy was first next OCG needs more control
low_A = nrow(ITA_SKILLs) +1 # Next available position
top_A = low_A + nrow(OCG_SKILLS) -1 # Max number of records to be added
OCG_Sk_row = nrow(OCG_SKILLS)
delta_r = top_A - OCG_Sk_row
# filas en OCG se controlan dentro del indice

for(i in low_A : top_A ){
  REPO_SKILLS[i,1] = "Switzerland"
  REPO_SKILLS[i,2] = OCG_SKILLS[i-delta_r,1] # Field Party_ID
  REPO_SKILLS[i,3] = OCG_SKILLS[i-delta_r,3]  # Field Skill
  REPO_SKILLS[i,4] = OCG_SKILLS[i-delta_r,4]  # Level
  REPO_SKILLS[i,5] = OCG_SKILLS[i-delta_r,5]  # Level_type
}

# # Filas en Repo_Skills se controlan por estas variables only two offices 
# # Italy was first next OCG  now Homere needs more control
# low_B = nrow(ITA_SKILLs)+  nrow(OCG_SKILLS) +1 # Next available position
# top_B = low_B + nrow(OCG_Homere_Skills) -1 # Max number of records to be added
# OCG_Hom_row = nrow(OCG_Homere_Skills)
# delta_r = top_B - OCG_Hom_row
# filas en OCG-Homere se controlan dentro del indice
skill_a = " "
level_a = " "
level_t = " "
for(i in low_B : top_B ){
  if(OCG_Homere_Skills[i-delta_r,3] == "Y"){ skill_a == "English"}
  REPO_SKILLS[i,1] = "OCG_Homere"
  REPO_SKILLS[i,2] = OCG_Homere_Skills[i-delta_r,1] # Field Party_ID
  REPO_SKILLS[i,3] = skill_a  # Field Skill
  REPO_SKILLS[i,4] = level_a  # Level
  REPO_SKILLS[i,5] = level_t  # Level_type
}
for(i in low_B +1 : top_B + 1){
if(OCG_Homere_Skills[i-delta_r,4] == "Y"){ skill_a == "French"}
  REPO_SKILLS[i,1] = "OCG_Homere"
  REPO_SKILLS[i,2] = OCG_Homere_Skills[i-delta_r,1] # Field Party_ID
  REPO_SKILLS[i,3] = skill_a  # Field Skill
  REPO_SKILLS[i,4] = level_a  # Level
  REPO_SKILLS[i,5] = level_t  # Level_type
}
for(i in low_B +2 : top_B + 2){
  if(OCG_Homere_Skills[i-delta_r,5] == "Y"){ skill_a == "Spanish"}
  REPO_SKILLS[i,1] = "OCG_Homere"
  REPO_SKILLS[i,2] = OCG_Homere_Skills[i-delta_r,1] # Field Party_ID
  REPO_SKILLS[i,3] = skill_a  # Field Skill
  REPO_SKILLS[i,4] = level_a  # Level
  REPO_SKILLS[i,5] = level_t  # Level_type

}


# save REPO_SKILLS for future usage in Excel and RDS format reading added for verification

write_xlsx(REPO_SKILLS,"REPO_SKILLS.xlsx")

saveRDS(REPO_SKILLS, "REPO_SKILLS.rds")



