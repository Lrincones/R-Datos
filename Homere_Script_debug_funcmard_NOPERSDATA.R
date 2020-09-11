# Ingestion first step converting all fields
# One script for each office and code sections for each data group in each office
# once it is done create functions for the sections
#
#
# Libraries
#
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
# Start with Homere Personal Data

# Function to clear the names blanks for "_"
colum_name_subs <- function(obj_load){
  object_names <- colnames(obj_load)
  for (x in 1 : length(object_names)){
    object_names[x] = chartr(" ", "_", object_names[x])
  }
  for (x  in 1 :length(object_names)) {
    colnames(obj_load)[x] = object_names[x]
  }
  return(obj_load)
}
# function to prepare the file for the office ID loaded example "OCGHomerePerData"
# ####################### Needs to improve, simplify 
# To prepare the data for the ID files from the source data given
# the information is from 3 fields from the metadata, TAB1
# the actual values come from the file with the data
# recall that the position depends on the given file, the position for the ID column in the data
# the write_offc_IDSnumber four"4" here should be a parameter, improvement for next version
#
#
#prepare_offc_IDS(offc_datos, offc_datos_id,df_IDS, col_ID_number, repo_id_old)
#
prepare_offc_IDS <- function(office_file , ID_datos, data_fram_1, col_param, old_repo_id) { # each file is from an office and group data
  # prepare the parameters for number of rows and to control while loop
  # prepare to check if the record existed in the repository col_ids_v
  n_row = 1 # to control records read from office_file
  n_row_w = 1 # to control records to write
  top_row = as.integer(nrow(office_file)) # Max value used to control the While loop
  # 
  # check in parameter repo_id_old to see if it was loaded
  offc_loaded <- old_repo_id$office # Just the office
  offc_loaded <- unlist(offc_loaded)
  check_office = ID_datos$Office[1] %in% offc_loaded # if false load all if true check the group in the source field
  perdat_load = subset(old_repo_id, old_repo_id$office == as.character(ID_datos[1,1]), select = c(source)) #los grupos cargados
  perdat_load = unlist(perdat_load)
  check_perdat = "PersonalData" %in% perdat_load # If true PersonalData was loaded for the office
  check_group = ID_datos$Source[1] %in% perdat_load # check the group # if false is a new group full load

  # If the office and personal data was loaded load the data verifiyng repo_id from PersonalData
  # Case 1 first time loaded
  #
  # Add an If here or in the while below
    # while (n_row  < top_row & !check_office) { # first load while there are records to be read load them
    #   # check record is in repository, if it is do not process it.
    #   # add functionality if the records are new only as an example
    #   if(check_office == FALSE || check_group == TRUE){ # if the record does not exits process it 
    #     if(office_file[n_row,col_param] == office_file[(n_row+1),col_param]){ # if the next ID is the same, prepare to read next after proccesing this
    #       n_row = n_row + 1
    #     }
    #     else { # if the next record is not the same ID prepare the data to create the records to be written
    #       data_fram_1[n_row_w,'office'] = ID_datos[1]
    #       data_fram_1[n_row_w,'source'] = ID_datos[2]
    #       data_fram_1[n_row_w,'system'] = ID_datos[3]
    #       data_fram_1[n_row_w,'key'] = ID_datos[4]  # the 4 should be a parameter pased in the function call
    #       data_fram_1[n_row_w,'value'] = office_file[n_row,col_param]
    #       n_row = n_row + 1 # must increase to get next record
    #       n_row_w = n_row_w + 1 # must increase to populate next record
    #     } # if the next record has the same id populate the records to be written
    #     data_fram_1[n_row_w,'office'] = ID_datos[1]
    #     data_fram_1[n_row_w,'source'] = ID_datos[2]
    #     data_fram_1[n_row_w,'system'] = ID_datos[3]
    #     data_fram_1[n_row_w,'key'] = ID_datos[4]
    #     data_fram_1[n_row_w,'value'] = office_file[top_row,col_param]
    #   }
    # }
  
  # Case 2 the office was loaded before need to check the records
  #
  n_row_ID = nrow(ID_datos)
  offc <- ID_datos$Office[1] # Tab1 de la carga
  col_ids <- subset(old_repo_id, office == offc, select=c(repoid))
  col_ids_v = unlist(col_ids) # generate the vector to check it was loaded or not ## que pasa if na no hay cargas anteriores??
  col_ids_ID <- subset(old_repo_id, office == offc, select=c(Legacy_ID))
  col_ids_ID_v = unlist(col_ids_ID)
  while (n_row  <= top_row & check_office & check_perdat) { # while there are records to be read the office was loaded with PersonalData
    # check record is in repository, if it is do not process it.
    # add functionality if the records are new only as an example
    check_record = office_file[n_row, col_param] %in%  col_ids_ID_v 
    if(check_record == FALSE || check_group == FALSE ){ # if the record does not exits process it
        data_fram_1[n_row_w,'office'] = ID_datos[1]
        data_fram_1[n_row_w,'source'] = ID_datos[2]
        data_fram_1[n_row_w,'system'] = ID_datos[3]
        data_fram_1[n_row_w,'key'] = ID_datos[4]  # the 4 should be a parameter pased in the function call
        data_fram_1[n_row_w,'value'] = office_file[n_row,col_param]
        n_row = n_row + 1 # must increase to get next record
        n_row_w = n_row_w + 1 # must increase to populate next record
      }
  
    else{ n_row = n_row + 1} # if record was loaded before go the next row
  } # End while Loop
  
  return(data_fram_1) 
} # End of function

#
# End of Functions Sections so far TWO FUNCTIONS  colum_name_subs and prepare_offc_IDS
#

#
# Create the list of files to be proccessed
# the input files will have two tabs the first tab is basic metadata
# Column Names First Tabs, same per every source; Office	Source	System	key
# Example Office: OCG	
# Example Source: Address	
# Example System: Dynamics	
# Example Key: Personnel number: 
#
# the second tab the data as received. May need to do some colum names adjustments
# Get the lis of files to be read in "archivos"
path_1 <-  "C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/Input"
archivos <- list.files(path = path_1,recursive = F , 
                       all.files = FALSE, full.names = TRUE, include.dirs = FALSE, pattern = "*Homere*") #  Check lower and upper case for testing I used ",pattern = "*Homere*""
# How many archivos to be read
num_arch = length(archivos)
arch_read = 1
while (arch_read <= num_arch){
  offc_datos_id <- read_excel(archivos[arch_read],sheet = 1)
  offc_datos <- read_excel(archivos[arch_read],sheet = 2)
  
  # converting to characters
  list_cols <- colnames(offc_datos)
  offc_datos <- offc_datos %>%
    convert(chr(all_of(list_cols)))
  # Replacing NA values 
  offc_datos[is.na(offc_datos)] <- "Empty"
  
  repo_id_old <- read_excel(
    "C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/NoInput/REPO_IDS.xlsx"
    ) # creating  need to convert numeric field to character
  list_cols <- colnames(repo_id_old)
  repo_id_old <- repo_id_old %>%
    convert(chr(all_of(list_cols)))
  
  # verify if the file to be load has been loaded before, if yes then is a maintenance of the data group
  # in this proof of concept we include only the case for new records
  #
  # Read the "Loaded_Summary.xlsx" file  all fields are character
  #
  Loaded_Summary_Old =read_excel(
    "C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/NoInput/Loaded_Summary.xlsx"
     ) # creating  need to convert numeric field to character
  list_cols <- colnames(Loaded_Summary_Old)
  Loaded_Summary_Old <- Loaded_Summary_Old %>%
    convert(chr(all_of(list_cols)))
  
  # Create the general data frame for the IDS using exceltemplate "df_id_template" all fields are character
  df_IDS <- read_excel(
    "C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/NoInput/df_id_template.xlsx"
    ) # all fields are character no need to convert
  
  # Find the position for the legacy ID
  col_ID_Name  = offc_datos_id$key
  col_ID_number = grep(col_ID_Name, colnames(offc_datos))
  
  # call the function to prepare IDs 5 Parameters Tab2, Tab1, Template for IDS File, Column for IDs, Previous loaded files
  IDS_file <- prepare_offc_IDS(offc_datos, offc_datos_id,df_IDS, col_ID_number, repo_id_old) #Feed the data frame
  
  #preparing the file name and writing the file in the working directory
  #for the source ID
  # getting the parts of the names, to concatenate and for the name
  
  a1 <- offc_datos_id[1] # office
  a2 <- offc_datos_id[4] # name of the field with legacy ID
  a3 <- offc_datos_id[3] # legacy system
  a4 <- offc_datos_id[2] # data Group to be load
  ID_filewr <- paste(a1,a2,a3,a4,".xlsx", sep = "", collapse = NULL)
  ID_filewr = paste("C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/Output/", ID_filewr, sep = "")
  write_xlsx(IDS_file, col_names = TRUE, format_headers = TRUE,path =  ID_filewr)
  
  #
  # Actualizar REPO_ID con lo procesado  OJO ARREGLAR EL DIRECTORIO DONDE SE ESCRIBE IDS_FILE
  # 
  
  # Determine the next repoid to use
  last_key <- as.integer(nrow(repo_id_old)) # record for the last key 
  last_key_val <- as.integer(repo_id_old[last_key,2]) # next number is the next repoid to use
  
  # column name for the legacy id in offc_datos. 
  col_id <-  offc_datos_id[4]
  c = grep(col_id, colnames(offc_datos))
  

  n_row  = 1 # to control records read from office_file
  n_roww = 1 # to control records to write to update repo_id # Should be the last position
  top_row = as.integer(as.character(nrow(offc_datos)))


  
  # delete after test communications load
  # IDS to check if record was loaded  OJO se esta haciendo vs los de personal data  
  #
  # col_ids_per = subset(repo_id_old, source == "PersonalData", select=c(Legacy_ID))  # EL LEGACY ID EN repo_id_old
  # col_ids_per_v = unlist(col_ids_per)
  # 
  # col_ID_Name  = offc_datos_id$key
  # col_ID_number = grep(col_ID_Name, colnames(offc_datos))
  
  # PersonalData Repoid corresponding to Legacy_ID
  col_ids_leg_repo = subset(repo_id_old, repo_id_old$office == as.character(offc_datos_id[1,1]), 
                            repo_id_old$source == "PersonalData", select=c(Legacy_ID))
  col_ids_leg_repo_v = unlist(col_ids_leg_repo)
  
  # check if record is in repository if it is process Personal data OK
  # the office for the file in process
  # off_c = as.character(offc_datos_id[1,1])
  # check_office = off_c %in% repo_id_old$office
  
  ofc_OCG <- subset(repo_id_old, office == "OCG")
  col_ofc_OCG = ofc_OCG[,6]
  col_ofc_OCG_v = unlist(col_ofc_OCG)
  
  # Check the system
  system_loaded <- repo_id_old$system
  check_system = offc_datos_id$System[1] %in% system_loaded # if false load all case 1 first time load
  
  # Check the SOurce for Homere Case if the source was not loaded we add to the repo_id_old
  # we do not create new repo_ids. The check is for system Homere office OCG
  source_loaded <- subset(repo_id_old, office == "OCG" & system == "Homere", select = source)
  check_source = offc_datos_id$Source[1] %in% source_loaded # if false load all case 1 first time load
  
  # For the case of a new system, in this instance Homere 
  if(!check_system){
    while (n_row <= top_row) {
      repo_id_old[n_roww + last_key,'office'] = offc_datos_id[1]
      repo_id_old[n_roww + last_key,'source'] = offc_datos_id[2]
      repo_id_old[n_roww + last_key,'system'] = offc_datos_id[3]
      repo_id_old[n_roww + last_key,'key'] = offc_datos_id[4]
      repo_id_old[n_roww + last_key,'repoid'] = as.character(n_roww + last_key_val) # ID for Repo_Id is an integer sequentialrepo_id_old[n_roww + last_key, 'Legacy_ID'] = offc_datos[n_row,c] # c has the column name for the ID
      repo_id_old[n_roww + last_key,'Legacy_ID'] = offc_datos[n_row,c]
      n_row = n_row + 1
      n_roww = n_roww + 1
    }
  } # End the while loop ---> MOVER LUEGO DE ESTAS ASIGNACIONES
  
  # For the case of a new source, for Homere, but this source was not loaded before 
  if(!check_source){
    while (n_row <= top_row) {
      repo_id_old[n_roww + last_key,'office'] = offc_datos_id[1]
      repo_id_old[n_roww + last_key,'source'] = offc_datos_id[2]
      repo_id_old[n_roww + last_key,'system'] = offc_datos_id[3]
      repo_id_old[n_roww + last_key,'key'] = offc_datos_id[4]
      repo_id_old[n_roww + last_key,'repoid'] = as.character(n_roww + last_key_val) # ID for Repo_Id is an integer sequentialrepo_id_old[n_roww + last_key, 'Legacy_ID'] = offc_datos[n_row,c] # c has the column name for the ID
      repo_id_old[n_roww + last_key,'Legacy_ID'] = offc_datos[n_row,c]
      n_row = n_row + 1
      n_roww = n_roww + 1
    }
  } 
  
  
  
  # El Sistema ha sido cargado antes y el source tambien
  if(check_system){
    while (n_row <= top_row) {
      check_record = offc_datos[n_row, c] %in% col_ofc_OCG_v # check if record legacy id is in repository if it is dont process it
      if(check_record){ # SKIP record
        n_row = n_row + 1 # Skip record
      } # if not skipped add it process it
      else { # write after last record (row + last row repo_id_old)
        
        repo_id_old[n_roww + last_key,'office'] = offc_datos_id[1]
        repo_id_old[n_roww + last_key,'source'] = offc_datos_id[2]
        repo_id_old[n_roww + last_key,'system'] = offc_datos_id[3]
        repo_id_old[n_roww + last_key,'key'] = offc_datos_id[4]
        repo_id_old[n_roww + last_key,'repoid'] = as.character(n_roww + last_key_val) # ID for Repo_Id is an integer sequentialrepo_id_old[n_roww + last_key, 'Legacy_ID'] = offc_datos[n_row,c] # c has the column name for the ID
        repo_id_old[n_roww + last_key,'Legacy_ID'] = offc_datos[n_row,2]
        n_row = n_row + 1
        n_roww = n_roww + 1
      }
    } # End the while loop ---> MOVER LUEGO DE ESTAS ASIGNACIONES
  }
  
  write_xlsx(repo_id_old, "C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/NoInput/REPO_IDS.xlsx", 
             col_names=TRUE, format_headers=TRUE)
  
  #
  # Code for creating the data in the repository from the source
  # to write the second tab from the object offc_datos
  #
  ID_filewrd1 <- paste(a1,a3,a4,".xlsx", sep = "", collapse = NULL)
  ID_filewrd1 <- str_replace(ID_filewrd1, " ", "")
  ID_filewrd1 = paste("C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/Output/", ID_filewrd1, sep = "")
  write_xlsx(offc_datos, col_names=TRUE, format_headers=TRUE, path = ID_filewrd1)
  
  arch_read = arch_read + 1 # incrementar los archivos leidos
  
  if(arch_read > num_arch){
    print("Termino la carga")
  }
} # end of while loop for reading archivos
# 
# Codigo para el sumario update Loaded_Summary
#
# read latest repo_id_old from above code
setwd("C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/NoInput/")
repo_id_old <- read_excel("REPO_IDS.xlsx")
# Ordenar por Office, System, Source and Key y crear un nuevo objeto
repo_id_order <- repo_id_old[order(repo_id_old$office, repo_id_old$system, repo_id_old$source, repo_id_old$key),]
# glimpse(repo_id_order)
# Elimino la columna que no se usa en Loaded_Summary y luego elimini las filas repetidas
repo_id_order <- select(repo_id_order, -repoid, -Legacy_ID)
repo_id_order = repo_id_order %>% distinct
# now to write the Loaded_Summary
# first lets reorder the columnss as per the Loaded_Summary
# office, Source, System, key
repo_id_order = repo_id_order %>% select(office, source, system, key)
write_xlsx(repo_id_order, col_names=TRUE, format_headers=TRUE, path = "Loaded_Summary.xlsx" )
#
