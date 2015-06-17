setwd("~/git/sorm")                                                    # change this directory setting to suit your preference
                                                                       # for a base directory; should be same as this script file
baseDir <- getwd()                                                     # many functions depend on base directory variable 
                                                                       

#--- various functions needed for the task of creating a tidy dataset for this project:
libraries <- c("XLConnect", "stringr", "dplyr", "rvest"); lapply(libraries, require, character.only=T)
source(paste(baseDir,"create_tidy_data_functions.R",sep="/"))                      

#--- scrape SORM web site for raw assessment workbook URLs, and cleanup URLs:
getSORMworkbookDf ()                                                                                                                 
cleanupURLs()
get_initial_wb_urls()
get_final_wb_urls()

#--- download SORM workbooks and save to subdirectories for further processing:
if(!paste(baseDir, "workbooks", sep="/") %in% list.dirs(baseDir)){
        dir.create(paste(baseDir, "workbooks", sep="/"))
}
get_initial_wbs()
get_final_wbs()

#--- make copies and rename the copies of SORM workbooks to make coding easier and safer:
make_copies_initial_wbs()
make_copies_final_wbs()

#--- extract the workbooks; copy selected columns (variables) and create new workbooks one worksheet per workbook:
extract_initial_wbs() 
extract_final_wbs() 

#--- combine workbooks for fiscal years and save in processed_files subdirectory:
library(xlsx)
if(!paste(baseDir, "processed_files", sep="/") %in% list.dirs(baseDir)){
        dir.create(paste(baseDir, "processed_files", sep="/"))
}
regularize_2004_final_extract()
combine_same_fy_workbooks()                        # This may take a couple of minutes, please wait...

#--- assemble tidyRawSORM as a dataframe and save as an R object before any code/name replacements or mergers are added:
assemble_raw_tidy_dataset()

#--- replace clients as needed; these replacement function calls update the instance of tidyRawSORM present in the .Global environment:
load(paste(baseDir, "tidyRawSORM.Rdata", sep="/"))
replace_client("C009","C185","Parmer")                                         # replaces instances of Bailey with Parmer     
replace_client("C128","C007","Atascosa")                                       # replaces instances of Karnes with Atascosa
replace_client("A527","A542","Cancer Prevention and Research Institute")       # replaces instances of Cancer Council with Cancer Prevention and Research Institute
replace_client("A325","A326","Texas Emergency Services Retirement System")     # replaces instances of Fire Fighter's Pension Commission with Texas Emergency Services Retirement System

#--- merge two predecessor clients together with new client code and new client name:
merge_and_replace_juve_clients(c("A694","A665"),"A644","Texas Juvenile Justice Department")

#--- save the modified tidyRawSORM dataframe as the tidySORM dataframe:
tidySORM <- tidyRawSORM
if(sum(duplicated(tidySORM$code.year)) == 0){
        save(tidySORM,file = paste(baseDir, "tidySORM.Rdata", sep="/"))
        
}else{
        print("WARNING: There are duplicates in the tidySORM$code.year dataframe that need to be fixed. No file written.")
}
rm("tidyRawSORM")                                                             # removes modified tidyRawSORM data frame from the workspace; unmodified tidyRawSORM is still saved on disk


        
        