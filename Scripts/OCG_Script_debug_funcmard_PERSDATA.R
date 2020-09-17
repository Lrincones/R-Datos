# Personal Data for Dynamics - OCG
# Using the general case to adapt for the OCG Case
# Ingestion first step converting all fields
#
# Libraries
#
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
#
# Functions
#
# Function to clear the names changing blanks for "_"
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
# function to prepare the file for the office ID loaded example "OCGDynamicsPersonalData"
# To prepare the data for the ID files from the source data given
# the information is from 3 fields at the metadata, TAB1, the tab is created manually when receiving the file
# the raw actual values are in the second tab
#
# Next line,  the parameters passed to the function
# prepare_offc_IDS(offc_datos, offc_datos_id,df_IDS, col_ID_number, repo_id_old)
# offc_datos -> raw data(tab2), offc_datos_id(tab1 metadata), df_IDS the template to populate, col_ID_number is the column in Raw data
# for the Legacy ID, repo_id_old is the previously loaded IDs to avoid creating duplicates repo ids and to proper document
# the new IDs
#
prepare_offc_IDS <- function(office_file , ID_datos, data_fram_1, col_param, old_repo_id) { 
  # this code is for the personal data file from OCG
  # prepare the parameters for number of rows and to control while loop
  # prepare to check if the record existed in the repository col_ids_v to avoid duplication
  n_row = 1 # to control records read from office_file
  n_next = 2
  n_row_w = 1 # to control records to write
  top_row = as.integer(nrow(office_file)) # Max value used to control the While loop
  # Which Office as per parameter ID_datos column 1 - In this case OCG
  # check in parameter repo_id_old to see if it was loaded
  offc_loaded <- repo_id_old$office # Offices previously loaded
  check_office = ID_datos$Office[1] %in% offc_loaded # if false load all - Have OCG Personal Data being loaded
  #
  # Case 1 first time loaded  VERIFY THE CODE AFTER THE ELSE IT MAY NOT APPLY FOR PERSONAL DATA
  #
  while (n_row  < top_row & !check_office) { # first load while there are records to be read load them
    # check record is in repository, if it is do not process it.
    # add functionality if the records are new only as an example
    if(check_office == FALSE){ # if the office record does not exits process it, in this case OCG 
      if(office_file[n_row,col_param] == office_file[(n_next),col_param]){ # if the next ID is the same, prepare to read next after proccesing this
        n_row = n_row + 1
      } # For OCG Case each row should be a single record
      else { # if the next record is not the same ID prepare the data to create the records to be written
        data_fram_1[n_row_w,'office'] = ID_datos[1]
        data_fram_1[n_row_w,'source'] = ID_datos[2]
        data_fram_1[n_row_w,'system'] = ID_datos[3]
        data_fram_1[n_row_w,'key'] = ID_datos[4]  
        data_fram_1[n_row_w,'value'] = office_file[n_row,col_param]
        n_row = n_row + 1 # must increase to get next record
        n_row_w = n_row_w + 1 # must increase to populate next record
        if(n_next < top_row) {n_next = n_next + 1}    
      } # if the next record has the same id populate the records to be written
      if(n_row == top_row){
        data_fram_1[n_row_w,'office'] = ID_datos[1]
        data_fram_1[n_row_w,'source'] = ID_datos[2]
        data_fram_1[n_row_w,'system'] = ID_datos[3]
        data_fram_1[n_row_w,'key'] = ID_datos[4]
        data_fram_1[n_row_w,'value'] = office_file[n_row,col_param]
      }
    }
  }
  
  #
  # Case 2 the office was loaded before need to check the records
  #
  # Create the IDs loaded, "col_ids_ID_v"  contains the IDs loaded
  #
  if(check_office){
    offc <- ID_datos$Office[1] # Tab1 de la carga
    col_ids <- subset(old_repo_id, office == offc, select=c(repoid))
    col_ids_v = unlist(col_ids) # generate the vector to check it was loaded or not ## que pasa if na no hay cargas anteriores??
    col_ids_ID <- subset(old_repo_id, office == offc, select=c(Legacy_ID))
    col_ids_ID_v = unlist(col_ids_ID) # IDs loaded
  }
  while (n_row  <= top_row & check_office) { # while there are records to be read
    # check record is in repository, if it is, don't process it.
    check_record = office_file[n_row, col_param] %in%  col_ids_ID_v 
    if(check_record == FALSE){ # if the record does not exist process it
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
# Example Source: PersonalData	
# Example System: Dynamics	
# Example Key: Personnel number: 
#
# the second tab the data as received. May need to do some colum names adjustments
# Get the lisT of files to be read in "archivos"
# For Production, the pattern to obtain the list of files, should be a parameter indicating the MSF section'
# As part of the logs registering the production workflow
#
# Files processed for the Data Ingestion.

# One input file with 2 tabs, this file is generated after receiving the single file from the section.
# A tab is added for the metadata (first tab) the second tab is the original content from the section.
# Four files are generated:
# One file with the Data as received from the section (Raw data)
# One file with the legacy IDs processed. In the case of personal data, is the base to identify the IDs to be ingested, if a person does not have the personal data, we do not create other data groups.
# The data ingestion generates the four repository files.
# 1- Data File (Raw Data)
# 2 - ID file (Legacy IDs)
# 3 - Repository IDs file (office, repoid (unique ID in the repository), key (variable name for the Legacy IDs), source (data group), system, Legacy_ID)
#     Each person has a unique record per section-source-system. The legacy ID is linked to the repository unique ID.
# 4 - Summary file (Section, source (data group), system, key (variable name for the Legacy IDs)
              
path_1 <-  "C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/Input/" # Update the path accordingly 
archivos <- list.files(path = path_1,recursive = F , 
                       all.files = FALSE, full.names = TRUE, pattern = "*OCG*") # for production this needs to be a parameter
# How many archivos to be read  For this case only one Personal Data from OCG
num_arch = length(archivos)
arch_read = 1
while (arch_read <= num_arch){
  # offc_datos_id <- read_excel(archivos[arch_read],sheet = 1)
  # offc_datos <- read_excel(archivos[arch_read],sheet = 2)
  passwrd_ocg = "#Only4Symphony"
  offc_datos_id <- xlsx::read.xlsx(archivos[arch_read], sheetIndex = 1, password = passwrd_ocg )
  offc_datos <- xlsx::read.xlsx(archivos[arch_read], sheetIndex = 2, password = passwrd_ocg )
  
  
  
  # converting to characters
  list_cols <- colnames(offc_datos)
  offc_datos <- offc_datos %>%
    convert(chr(all_of(list_cols)))
  
  repo_id_old <- read_excel(
    "C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/NoInput/REPO_IDS.xlsx"
    ) # creating.  Converting numeric field to character
  list_cols <- colnames(repo_id_old)
  repo_id_old <- repo_id_old %>%
    convert(chr(all_of(list_cols)))
  
  #
  # verify if the file to be load has been loaded before, if yes then is a maintenance of the data group
  # in this proof of concept we include only the case for new records. 
  # Next version include updates need a rule for key fields determining if data needs maintenance or not
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
  
  a1 <- offc_datos_id[1,1] # office
  a2 <- offc_datos_id[1,4] # name of the field with legacy ID
  a3 <- offc_datos_id[1,3] # legacy system
  a4 <- offc_datos_id[1,2] # data Group to be load
  ID_filewr <- paste(a1,a2,a3,a4,".xlsx", sep = "", collapse = NULL)
  write_xlsx(IDS_file, col_names = TRUE, format_headers = TRUE,path =  ID_filewr)
  
  #
  # Actualizar REPO_ID con lo procesado  OJO ARREGLAR EL DIRECTORIO DONDE SE ESCRIBE IDS_FILE
  # 
  last_key <- nrow(repo_id_old) # record for the last key 
  last_key_val <- as.integer(repo_id_old[last_key,2]) # next record is the next number
  
  
  col_id <-  offc_datos_id[4]
  c = colnames(offc_datos[col_id[1,1]])
  valor = as.integer(as.character(last_key_val[1]))
  n_ids = as.integer(as.character(nrow(offc_datos)))
  

  n_row  = 1 # to control records read from office_file
  n_roww = 1 # to control records to write
  top_row = as.integer(n_ids)
  delta = last_key + 1
  
  ofc_1 <-  as.character(offc_datos_id[1,1])
  ofc_load <- subset(repo_id_old, office == ofc_1 ) 
  col_ofc_load = ofc_load[,"Legacy_ID"] 
  col_ofc_load_v = unlist(col_ofc_load)

  
  #
  # IDS to check if record was loaded  OJO se esta haciendo en  col_ofc_load_V son los de persdata
  #
  col_ids = repo_id_old[,"Legacy_ID"]  # EL LEGACY ID EN repo_id_old
  col_ids_v = unlist(col_ids)
  
  col_ID_Name  = offc_datos_id$key
  col_ID_number = grep(col_ID_Name, colnames(offc_datos))
  
  # check if record is in repository if it is process Personal data OK
  
  while (n_row <= top_row) {
    check_record = offc_datos[n_row, c] %in% col_ofc_load_v # check if record legacy id is in repository if it is dont process it
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
  
  write_xlsx(repo_id_old, "C:/Users/luis.rincones/OneDrive - MSF/Documents/0-DataRep/0-July_2020-Repository/NoInput/REPO_IDS.xlsx", 
             col_names=TRUE, format_headers=TRUE)
  
  #
  # Code for creating the data in the repository from the source
  # to write the second tab from the object offc_datos

  #
  ID_filewrd1 <- paste(a1,a3,a4,".xlsx", sep = "", collapse = NULL)
  ID_filewrd1 <- str_replace(ID_filewrd1, " ", "")
  write_xlsx(offc_datos,ID_filewrd1,col_names=TRUE, format_headers = TRUE)
  
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
# Elimino la columna que no se usa en Loaded_Summary y luego elimino las filas repetidas
repo_id_order <- select(repo_id_order, -repoid, -Legacy_ID)
repo_id_order = repo_id_order %>% distinct
# now to write the Loaded_Summary
# first lets reorder the columnss as per the Loaded_Summary
# office, Source, System, key
repo_id_order = repo_id_order %>% select(office, source, system, key)
write_xlsx(repo_id_order, col_names=TRUE, format_headers=TRUE, path = "Loaded_Summary.xlsx" )
#
