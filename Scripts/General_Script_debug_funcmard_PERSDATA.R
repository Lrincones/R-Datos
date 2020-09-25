# Ingestion first step converting all fields
# The order is staring with Personal data
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
# 

# function to prepare the file for the office ID loaded example "OCGHomerePerData" 
# for production version check performance for the productive version 
# To prepare the data for the ID files from the source data given
# the information is from 3 fields from the metadata, TAB1
# the actual values come from the file with the data
# recall that the filed's position depends on the given file, the position for the ID column in the data
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
  # it is the Office as per parameter ID_datos column 1
  # check in parameter repo_id_old to see if it was loaded
  load_office = ID_datos$Office[1]
  load_source = ID_datos$Source[1]
  load_system = ID_datos$System[1]
  offc_loaded <- old_repo_id$Office
  check_office = load_office %in% offc_loaded # if false load all
  system_loaded <- old_repo_id$System
  check_system = load_system %in% system_loaded # if false load all
  sys_source = which((old_repo_id$System == load_system))
  sys_source1 = which(old_repo_id$Source == load_source)
  check_source = sys_source1 %in% sys_source
  if(is_empty(check_source)) {check_source = FALSE}
  # Case 1 first time loaded
  while (n_row  < top_row &
         (!check_office | !check_system |
          !check_source)) {
           # first load while there are records to  read load them
           # check record is in repository, if it is do not process it.
           # add functionality if the records are new only as an example
           if (check_office == FALSE |
               check_system == FALSE | check_source == FALSE) {
             # if the record does not exits process it
             if (office_file[n_row, col_param] == office_file[(n_row + 1), col_param]) {
               # if the next ID is the same, prepare to read next after proccesing this
               n_row = n_row + 1
             }
             else {
               # if the next record is not the same ID prepare the data to create the records to be written
               data_fram_1[n_row_w, 'Office'] = ID_datos[1]
               data_fram_1[n_row_w, 'Source'] = ID_datos[2]
               data_fram_1[n_row_w, 'System'] = ID_datos[3]
               data_fram_1[n_row_w, 'key'] = ID_datos[4]  # the 4 should be a parameter pased in the function call
               data_fram_1[n_row_w, 'value'] = office_file[n_row, col_param]
               n_row = n_row + 1 # must increase to get next record
               n_row_w = n_row_w + 1 # must increase to populate next record
             } # if the next record has the same id populate the records to be written
             if (n_row == top_row) {
               data_fram_1[n_row_w, 'Office'] = ID_datos[1]
               data_fram_1[n_row_w, 'Source'] = ID_datos[2]
               data_fram_1[n_row_w, 'System'] = ID_datos[3]
               data_fram_1[n_row_w, 'key'] = ID_datos[4]
               data_fram_1[n_row_w, 'value'] = office_file[top_row, col_param]
             }
           }
         }

  
  # Case 2 the office was loaded before need to check the records
  #
  n_row_ID = nrow(ID_datos)
  offc <- ID_datos$Office[1] # Tab1 de la carga
  col_ids <- subset(old_repo_id, Office == offc, select=c(repoid))
  col_ids_v = unlist(col_ids) # generate the vector to check it was loaded or not 
  col_ids_ID <- subset(old_repo_id, Office == offc, select=c(Legacy_ID))
  col_ids_ID_v = unlist(col_ids_ID)
  while (n_row  <= top_row & check_office) { # while there are records to  read
    # check record is in repository, if it is do not process it.
    # add functionality if the records are new only as an example
    check_record = office_file[n_row, col_param] %in%  col_ids_ID_v 
    if(check_record == FALSE){ # if the record does not exits process it
        data_fram_1[n_row_w,'Office'] = ID_datos[1]
        data_fram_1[n_row_w,'Source'] = ID_datos[2]
        data_fram_1[n_row_w,'System'] = ID_datos[3]
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
# End of Functions Sections 
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
# Get the lis of files to  read in "archivos"
path_1 <-  "C:/Users/luis.rincones/OneDrive - MSF/Documents/Input/"
archivos <- list.files(path = path_1,recursive = F , 
                       all.files = FALSE, full.names = TRUE, pattern = "*xlsx*")
# How many files to  read
num_arch = length(archivos)
arch_read = 1
while (arch_read <= num_arch){
  offc_datos_id <- read_excel(archivos[arch_read],sheet = 1)
  offc_datos <- read_excel(archivos[arch_read],sheet = 2)
  # converting to characters
  list_cols <- colnames(offc_datos)
  offc_datos <- offc_datos %>%
    convert(chr(all_of(list_cols)))
  
  repo_id_old <- read_excel(
    "C:/Users/luis.rincones/OneDrive - MSF/Documents/NoInput/REPO_IDS.xlsx"
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
    "C:/Users/luis.rincones/OneDrive - MSF/Documents/NoInput/Loaded_Summary.xlsx"
     ) # creating  need to convert numeric field to character
  list_cols <- colnames(Loaded_Summary_Old)
  Loaded_Summary_Old <- Loaded_Summary_Old %>%
    convert(chr(all_of(list_cols)))
  
  # Create the general data frame for the IDS using exceltemplate "df_id_template" all fields are character
  df_IDS <- read_excel(
    "C:/Users/luis.rincones/OneDrive - MSF/Documents/NoInput/df_id_template.xlsx"
    ) # all fields are character no need to convert
  
  # Find the position for the legacy ID
  col_ID_Name  = offc_datos_id$key
  col_ID_number = grep(col_ID_Name, colnames(offc_datos))
  
  # call the function to prepare IDs 5 Parameters Tab2, Tab1, Template for IDS File, Column for IDs, Previous loaded files
  IDS_file <- prepare_offc_IDS(offc_datos, offc_datos_id,df_IDS, col_ID_number, repo_id_old) #Feed the data frame
  
  #preparing the file name and writing the file in the working directory
  #for the source ID
  # getting the parts of the names, to concatenate and for the name
  df_IDS_conca <- paste(df_IDS$key[1], df_IDS$value[1],df_IDS$Office[1], df_IDS$Source[1], df_IDS$System[1], sep="", collapse =NULL)
  
  IDS_file_conca <- paste(IDS_file$key[1], IDS_file$value[1],IDS_file$Office[1], IDS_file$Source[1], IDS_file$System[1], sep="", collapse =NULL)
  
if( nrow(IDS_file) >1  & (IDS_file_conca != df_IDS_conca)){
  a1 <- offc_datos_id[1] # office
  a2 <- offc_datos_id[4] # name of the field with legacy ID
  a3 <- offc_datos_id[3] # legacy system
  a4 <- offc_datos_id[2] # data Group to be load
  ID_filewr <- paste(a1,a2,a3,a4,".xlsx", sep = "", collapse = NULL)
  ID_filewr <- paste("C:/Users/luis.rincones/OneDrive - MSF/Documents/Output/", ID_filewr, sep = "")
  write_xlsx(IDS_file, col_names = TRUE, format_headers = TRUE,path =  ID_filewr)
} 
  #
  # Actualizar REPO_ID con lo procesado  OJO ARREGLAR EL DIRECTORIO DONDE SE ESCRIBE IDS_FILE
  # 
  last_key <- nrow(repo_id_old) # record for the last key 
  last_key_val <- as.integer(repo_id_old[last_key,2]) # next record is the next number
  
  
  col_id <-  offc_datos_id[4]
  c = grep(col_id, colnames(offc_datos))
  valor = as.integer(as.character(last_key_val[1]))
  n_ids = as.integer(as.character(nrow(offc_datos)))
  
 
  n_row  = 1 # to control records read from office_file
  n_roww = 1 # to control records to write
  top_row = as.integer(n_ids)
  delta = last_key + 1
  
  #
  # Office to check if office from the record was loaded
  #
  ofc_1 <-  as.character(offc_datos_id[1,"Office"]) 
  ofc_arch <- subset(repo_id_old, Office == ofc_1)
  col_ofc_arch = ofc_arch[,6]
  col_ofc_arch_v = unlist(col_ofc_arch)

  
  #
  # IDS to check if record was loaded  
  #
  col_ids = repo_id_old[,6]  # EL LEGACY ID EN repo_id_old
  col_ids_v = unlist(col_ids)
  
  col_ID_Name  = offc_datos_id$key
  col_ID_number = grep(col_ID_Name, colnames(offc_datos))
  
  # Check the system
  system_loaded <- repo_id_old$System
  check_system = offc_datos_id$System[1] %in% system_loaded # if false load all case 1 first time load
  
  # check if record is in repository if it is process, Production version needs data quality report after loads
  # check the system
  
  # For the case of a new system
  if(!check_system){
  while (n_row <= top_row) {
      repo_id_old[n_roww + last_key,'Office'] = offc_datos_id[1]
      repo_id_old[n_roww + last_key,'Source'] = offc_datos_id[2]
      repo_id_old[n_roww + last_key,'System'] = offc_datos_id[3]
      repo_id_old[n_roww + last_key,'key'] = offc_datos_id[4]
      repo_id_old[n_roww + last_key,'repoid'] = as.character(n_roww + last_key_val) # ID for Repo_Id is an integer sequentialrepo_id_old[n_roww + last_key, 'Legacy_ID'] = offc_datos[n_row,c] # c has the column name for the ID
      repo_id_old[n_roww + last_key,'Legacy_ID'] = offc_datos[n_row,c]
      n_row = n_row + 1
      n_roww = n_roww + 1
    }
  } # End the while loop 
   
  
  #
  # El Sistema ha sido cargado antes
  if(check_system){
  while (n_row <= top_row) {
    check_record = offc_datos[n_row, c] %in% col_ofc_arch_v # check if record legacy id is in repository if it is dont process it
    if(check_record){ # SKIP record
      n_row = n_row + 1 # Skip record
    } # if not skipped add it process it
    else { # write after last record (row + last row repo_id_old)
      
      repo_id_old[n_roww + last_key,'Office'] = offc_datos_id[1]
      repo_id_old[n_roww + last_key,'Source'] = offc_datos_id[2]
      repo_id_old[n_roww + last_key,'System'] = offc_datos_id[3]
      repo_id_old[n_roww + last_key,'key'] = offc_datos_id[4]
      repo_id_old[n_roww + last_key,'repoid'] = as.character(n_roww + last_key_val) # ID for Repo_ID, c has the column name for the ID
      repo_id_old[n_roww + last_key,'Legacy_ID'] = offc_datos[n_row,c]
      n_row = n_row + 1
      n_roww = n_roww + 1
    }
  } # End the while loop ---> MOVER LUEGO DE ESTAS ASIGNACIONES
  }
  
  write_xlsx(repo_id_old, "C:/Users/luis.rincones/OneDrive - MSF/Documents/NoInput/REPO_IDS.xlsx", 
             col_names=TRUE, format_headers=TRUE)
  
  # 
  # Code for creating the data in the repository from the source
  # to write the second tab from the object offc_datos
  #
  a1 <- offc_datos_id[1] # office
  a3 <- offc_datos_id[3] # legacy system
  a4 <- offc_datos_id[2] # data Group to be load
  ID_filewrd1 <- paste(a1,a3,a4,".xlsx", sep = "", collapse = NULL)
  ID_filewrd1 <- paste("C:/Users/luis.rincones/OneDrive - MSF/Documents/Output/", ID_filewrd1, sep = "")
  write_xlsx(offc_datos, col_names=TRUE, format_headers=TRUE, path = ID_filewrd1)
  
  arch_read = arch_read + 1 # incrementar los archivos leidos
  
  
  # Update Loaded summary for the file
  
  loaded_sum_old <- read_excel("C:/Users/luis.rincones/OneDrive - MSF/Documents/NoInput/Loaded_Summary.xlsx")
  a1 <- loaded_sum_old$Office
  a2 <- loaded_sum_old$Source
  a3 <- loaded_sum_old$System
  a4 <- loaded_sum_old$key
  loaded_sum_old$conca <- paste(a1,a2,a3,a4, sep = "", collapse = NULL)
  a1 <- offc_datos_id$Office
  a2 <- offc_datos_id$Source
  a3 <- offc_datos_id$System
  a4 <- offc_datos_id$key
  offc_datos_id_con <- paste(a1,a2,a3,a4, sep = "", collapse = NULL)
  check_loaded_sum = offc_datos_id_con%in% loaded_sum_old$conca
  if(!check_loaded_sum){
    loaded_sum_old <- loaded_sum_old %>% add_row(Office = a1, Source = a2, System = a3, key = a4, conca = paste(a1,a2,a3,a4, sep = "", collapse = NULL) )
  }
  write_xlsx(loaded_sum_old[,1:4], col_names=TRUE, format_headers=TRUE, 
             path = "C:/Users/luis.rincones/OneDrive - MSF/Documents/NoInput/Loaded_Summary.xlsx" )  
  # Loaded summary Updated  for the file
  
  if(arch_read > num_arch){
    print("Load batch ended")
  }
} # end of while loop for reading archivos

