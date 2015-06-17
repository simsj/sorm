# ----------------------------------------------------- Scrape SORM web site and cleanup URLs:
getSORMworkbookDf <- function(){
        library(rvest)
        url <- "http://www.sorm.state.tx.us/legislative-reports/assessments"  #correct as of 16 Jun 2015
        
        theTitles <- url %>%
                html()  %>%
                html_nodes("td a") %>%                                         #use Google Chrome SelectorGadget to find best tag to use                
                html_text()                                     
        
        theLinks <- url %>%                          
                html()  %>%
                html_nodes("td a")  %>%                                        
                html_attr("href")             
        
        theDescriptions <- url %>%
                html()  %>%
                html_nodes("td:nth-child(2)")  %>%                            #use Google Chrome SelectorGadget to find best tag to use 
                html_text() 
        
        temp <- as.data.frame(cbind(theTitles,theLinks,theDescriptions ))
        colnames(temp) <- c("fiscalyear", "url","description")
        temp$url <- as.character(temp$url); temp$description <- as.character(temp$description)
        
        urlFront <- "https://www.sorm.state.tx.us"                            
        for(i in 1:nrow(temp)){                                               
                if(substr(temp$url[i], start=1, stop=1) == "/"){
                        temp$url[i] <- paste(urlFront,temp$url[i], sep="")  
                }
        }
        urls_SORM_workbooks_df <<- temp
}
cleanupURLs <- function(){
        temp <- urls_SORM_workbooks_df
        temp$url <- as.character(temp$url)
        for(i in 1:length(temp$url)){
                if(!grepl("%20",temp$url[i])){
                        #this string has spaces rather than "%20" characters
                        theWords <- unlist(str_split(temp$url[i]," "))
                        fixedURL <- theWords[1]
                        if(length(theWords) > 1){
                                for(j in 2:length(theWords)){
                                        fixedURL <- paste(fixedURL, "%20",theWords[j], sep="") 
                                }
                        }
                        temp$url[i] <-  fixedURL
                }
        }
        
        cleaned_urls <<- temp[order(temp$fiscalyear),]
}
get_initial_wb_urls <- function(){
        solo_wbs <- filter(cleaned_urls, fiscalyear %in% c("FY2002","FY2003","FY2005"))
        
        init_wbs <- cleaned_urls[grep("initial", cleaned_urls$url, ignore.case=TRUE), ]
        
        init2004 <- filter(cleaned_urls, cleaned_urls$fiscalyear == "FY2004")
        init2004 <- init2004[!grepl("final", init2004$url,  ignore.case=TRUE),]
        
        init2006 <- filter(cleaned_urls, cleaned_urls$fiscalyear == "FY2006")
        init2006 <- init2006[!grepl("May", init2006$url,  ignore.case=TRUE),]
        
        init2007 <- filter(cleaned_urls, cleaned_urls$fiscalyear == "FY2007")
        init2007 <- init2007[!grepl("adjustment", init2007$url,  ignore.case=TRUE),]
        
        init2009 <- filter(cleaned_urls, cleaned_urls$fiscalyear == "FY2009")
        init2009 <- init2009[!grepl("midyear", init2009$url,  ignore.case=TRUE),]
        
        initial_wb_urls <<- rbind(solo_wbs, init2004, init2006, init2007, init2009, init_wbs)
}
get_final_wb_urls <- function(){
        
        fin2007 <- filter(cleaned_urls, cleaned_urls$fiscalyear == "FY2007")
        fin2007 <- fin2007[grepl("adjustment", fin2007$url,  ignore.case=TRUE),]
        final_wb_urls <- fin2007
        
        theyears <- c("2004", "2006")
        for(i in 1:length(theyears)){
                temp <- cleaned_urls[cleaned_urls$fiscalyear == paste("FY",theyears[i],sep=""),]
                temp <- temp[grepl("final", temp$url,  ignore.case=TRUE),]
                final_wb_urls <- rbind(final_wb_urls, temp)      
        }
        theyears <- c("2008", "2009","2010","2011","2012","2013","2014")
        for(i in 1:length(theyears)){
                temp <- cleaned_urls[cleaned_urls$fiscalyear == paste("FY",theyears[i],sep=""),]
                temp <- temp[grepl("midyear", temp$url,  ignore.case=TRUE),]
                final_wb_urls <- rbind(final_wb_urls, temp)
        }
        theyears <- c("2015")
        for(i in 1:length(theyears)){
                temp <- cleaned_urls[cleaned_urls$fiscalyear == paste("FY",theyears[i],sep=""),]
                temp <- temp[grepl("mid-year", temp$url,  ignore.case=TRUE),]
                final_wb_urls <- rbind(final_wb_urls, temp)
        }
        final_wb_urls <<- arrange(final_wb_urls, fiscalyear)
}
# ----------------------------------------------------- Download SORM workbooks:
get_initial_wbs  <- function(){ 
        if(!paste(baseDir, "workbooks", "initial",sep="/") %in% list.dirs(baseDir)){
                dir.create(paste(baseDir, "workbooks", "initial",sep="/"))
        }
        theURLs <- initial_wb_urls$url
        for(i in 1:length(theURLs)){
                temp <- unlist(strsplit(theURLs[i], "/"))
                destination <- paste(baseDir,"workbooks","initial", temp[length(temp)],sep="/")
                download.file(theURLs[i], destfile = destination, method = "curl")
        }
}
get_final_wbs <- function(){
        if(!paste(baseDir, "workbooks", "final",sep="/") %in% list.dirs(baseDir)){
                dir.create(paste(baseDir, "workbooks", "final",sep="/"))
        }
        theURLs <- final_wb_urls$url
        for(i in 1:length(theURLs)){
                temp <- unlist(strsplit(theURLs[i], "/"))
                destination <- paste(baseDir,"workbooks","final", temp[length(temp)],sep="/")
                download.file(theURLs[i], destfile = destination, method = "curl")
        }
}
# ----------------------------------------------------- Create copies of downloaded workbooks and work on the copies, not the originals:
make_copies_final_wbs <- function(){
        if(!paste(baseDir, "workbooks/final_copies", sep="/") %in% list.dirs(baseDir)){
                dir.create(paste(baseDir, "workbooks/final_copies", sep="/"))
        } 
        theFileList <- list.files(paste(baseDir, "workbooks/final", sep="/")) 
        for(i in 1:length(theFileList)){
                file.copy(paste(baseDir, "workbooks/final",theFileList[i],sep="/"), paste(baseDir, "workbooks/final_copies",theFileList[i],sep="/"))
                if(grepl("xlsx",theFileList[i]) | grepl("XLSX",theFileList[i])){
                        new_name <- paste("20", substr(theFileList[i], start = 3, stop= 4),"_final.xlsx",sep="")
                }else{
                        new_name <- paste("20", substr(theFileList[i], start = 3, stop= 4),"_final.xls",sep="")
                } 
                file.rename(paste(baseDir, "workbooks/final_copies",theFileList[i],sep="/"), paste(baseDir, "workbooks/final_copies", new_name,sep="/"))
                
        }
}
make_copies_initial_wbs <- function(){
        if(!paste(baseDir, "workbooks/initial_copies", sep="/") %in% list.dirs(baseDir)){
                dir.create(paste(baseDir, "workbooks/initial_copies", sep="/"))
        } 
        theFileList <- list.files(paste(baseDir, "workbooks/initial", sep="/")) 
        for(i in 1:length(theFileList)){
                file.copy(paste(baseDir, "workbooks/initial",theFileList[i],sep="/"), paste(baseDir, "workbooks/initial_copies",theFileList[i],sep="/"))
                if(grepl("xlsx",theFileList[i]) | grepl("XLSX",theFileList[i])){

                        new_name <- paste("20", substr(theFileList[i], start = 3, stop= 4),"_initial.xlsx",sep="")
                }else{

                        new_name <- paste("20", substr(theFileList[i], start = 3, stop= 4),"_initial.xls",sep="")
                } 
                file.rename(paste(baseDir, "workbooks/initial_copies",theFileList[i],sep="/"), paste(baseDir, "workbooks/initial_copies", new_name,sep="/"))
                
        }
}
# ----------------------------------------------------- Extract copies of downloaded workbooks:
extract_initial_wbs <- function(){
        if(!paste(baseDir, "workbooks/initial_extracts", sep="/") %in% list.dirs(baseDir)){
                dir.create(paste(baseDir, "workbooks/initial_extracts", sep="/"))
        } 
        extract_initial_2003()                                                                        # there was only one workbook for 2003
        extract_initial_2004()
        extract_initial_2005()                                                                        # there was only one workbook for 2005 
        extract_2006_2007("initial") 
        extract_2008("initial")  
        extract_initial_2009()  
        extract_initial_2010()  
        extract_initial_2011_2015() 
}
extract_final_wbs <- function(){
        if(!paste(baseDir, "workbooks/final_extracts", sep="/") %in% list.dirs(baseDir)){
                dir.create(paste(baseDir, "workbooks/final_extracts", sep="/"))
        } 
        extract_final_2004()
        extract_2006_2007("final") 
        extract_2008("final") 
        extract_final_2009_2015()
} 
extract_initial_2003 <- function(){  
        library(XLConnect)
        theYears <-2003
        theYearRowEnd <- 290
        i <- 1
        ExtractFilePath <- paste(baseDir,"/workbooks/initial_extracts/",theYears[i],"_initial_extract.xls", sep = "")
        theExtractWorkbook <- loadWorkbook(ExtractFilePath, create = TRUE)                                     # create a new empty file, if file does not exist
        theSheetName <- paste(theYears[i],"_data", sep="")
        createSheet(theExtractWorkbook,name = theSheetName) 
        theSourceFile <- paste(baseDir,"/workbooks/initial_copies/",theYears[i],"_initial.xls", sep = "") 
        theSourceWorkbook <- loadWorkbook(theSourceFile, create = FALSE)                                      # load in the SORM workbook for further reading
        ##
        # agency code, name, assessment   Pattern3 = assessment(1,2,10) 
        agencyCodeAndName <- readWorksheet(theSourceWorkbook, sheet="assessment", startRow = 5, startCol=1, endRow = theYearRowEnd[i], endCol = 2, header=FALSE)
        colnames(agencyCodeAndName) <- c("code", "name")
        writeWorksheet(theExtractWorkbook, agencyCodeAndName, sheet=theSheetName, startRow = 1, startCol = 1) ## fills columns 1, 2
        theAssessment <- readWorksheet(theSourceWorkbook, sheet="assessment", startRow=5, startCol=10, endRow = theYearRowEnd[i], endCol = 10, header=FALSE)
        colnames(theAssessment) <- c(paste("assessmentdollarsfy", theYears[i], sep="")) ## edited from original version
        writeWorksheet(theExtractWorkbook, theAssessment, sheet= theSheetName, startRow=1, startCol=3) ## fills column 3
        ##
        # payroll_4, payroll_3, payroll_2 Pattern1 = payroll(4,5,6) 
        payrollData <- readWorksheet(theSourceWorkbook, sheet="payroll", startRow=5, startCol=4, endRow= theYearRowEnd[i], endCol=6, header=FALSE)
        colnames(payrollData)<- c(paste("payrolldollarsfy", theYears[i]-4,"per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, payrollData, sheet = theSheetName, startRow = 1, startCol=4) ## fills in columns 4, 5 and 6
        ##
        # fte_4, fte_3, fte_2 Pattern2 = IFR(7,12,17)  
        fte_4 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=7, endRow = theYearRowEnd[i], endCol=7, header=FALSE)
        colnames(fte_4) <- c(paste("ftefy", theYears[i]-4, "per", theYears[i], sep=""))
        fte_3 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=12, endRow = theYearRowEnd[i], endCol=12, header=FALSE)
        colnames(fte_3) <- c(paste("ftefy", theYears[i]-3,"per", theYears[i], sep=""))
        fte_2 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=17, endRow = theYearRowEnd[i], endCol=17, header=FALSE)
        colnames(fte_2) <- c(paste("ftefy", theYears[i]-2, "per", theYears[i], sep=""))
        ##
        # claims_4, claims_3, claims_2 Pattern2 = IFR(22,23,24) 
        theClaims <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=22, endRow = theYearRowEnd[i], endCol=24, header=FALSE)
        colnames(theClaims) <- c(paste("claimsfy", theYears[i]-4,"per", theYears[i], sep=""), paste("claimsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-2, "per", theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, fte_4, sheet= theSheetName, startRow = 1, startCol=7) ## fills in col 7
        writeWorksheet(theExtractWorkbook, fte_3 , sheet= theSheetName, startRow = 1, startCol=8) ## fills in col 8
        writeWorksheet(theExtractWorkbook, fte_2, sheet = theSheetName, startRow = 1, startCol=9) ## fills in col 9
        writeWorksheet(theExtractWorkbook, theClaims, sheet= theSheetName, startRow=1, startCol=10) ## fills in col 10,11,12
        ##
        # costs_4, costs_3, costs_2 Pattern = costs(5,6,7)
        costData <- readWorksheet(theSourceWorkbook, sheet="costs", startRow=5, startCol=5, endRow= theYearRowEnd[i], endCol=7, header=FALSE)
        colnames(costData) <- c(paste("costdollarsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, costData, sheet= theSheetName, startRow=1, startCol=13) ## fills in col 13, 14, 15
        ##
        # SORM cap on assessment based on numbers of claims; this is the column named difference as found in column Z
        capOnNumOfClaims <- readWorksheet(theSourceWorkbook, sheet="claims", startRow = 5, startCol=26, endRow = theYearRowEnd[i], endCol = 26, header=FALSE)
        colnames(capOnNumOfClaims) <- c(paste("capnumclaimsfy",theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, capOnNumOfClaims, sheet=theSheetName, startRow = 1, startCol = 16) ## fills column 16
        ##
        # SORM cap on assessment based on claims cost; this is the column named difference as found in column V
        capOnClamsCost <- readWorksheet(theSourceWorkbook, sheet="costs", startRow = 5, startCol=22, endRow = theYearRowEnd[i], endCol = 22, header=FALSE)
        colnames(capOnClamsCost) <- c(paste("capclaimscostfy",theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, capOnClamsCost, sheet=theSheetName, startRow = 1, startCol = 17) ## fills column 17
        ##
        saveWorkbook(theExtractWorkbook)
        
}
extract_initial_2004 <- function(){  
        library(XLConnect)
        theYears <-2004
        theYearRowEnd <- 290
        i <- 1
        ExtractFilePath <- paste(baseDir,"/workbooks/initial_extracts/",theYears[i],"_initial_extract.xls", sep = "")
        theExtractWorkbook <- loadWorkbook(ExtractFilePath, create = TRUE)                                     # create a new empty file, if file does not exist
        theSheetName <- paste(theYears[i],"_data", sep="")
        createSheet(theExtractWorkbook,name = theSheetName) 
        theSourceFile <- paste(baseDir,"/workbooks/initial_copies/",theYears[i],"_initial.xls", sep = "") 
        theSourceWorkbook <- loadWorkbook(theSourceFile, create = FALSE)                                      # load in the SORM workbook for further reading
        ##
        # agency code, name, assessment   Pattern3 = assessment(1,2,10) 
        agencyCodeAndName <- readWorksheet(theSourceWorkbook, sheet="assessment", startRow = 5, startCol=1, endRow = theYearRowEnd[i], endCol = 2, header=FALSE)
        colnames(agencyCodeAndName) <- c("code", "name")
        writeWorksheet(theExtractWorkbook, agencyCodeAndName, sheet=theSheetName, startRow = 1, startCol = 1) ## fills columns 1, 2
        theAssessment <- readWorksheet(theSourceWorkbook, sheet="assessment", startRow=5, startCol=10, endRow = theYearRowEnd[i], endCol = 10, header=FALSE)
        colnames(theAssessment) <- c(paste("assessmentdollarsfy", theYears[i], sep="")) ## edited from original version
        writeWorksheet(theExtractWorkbook, theAssessment, sheet= theSheetName, startRow=1, startCol=3) ## fills column 3
        ##
        # payroll_4, payroll_3, payroll_2 Pattern1 = payroll(3,4,5) 
        payrollData <- readWorksheet(theSourceWorkbook, sheet="payroll", startRow=5, startCol=3, endRow= theYearRowEnd[i], endCol=5, header=FALSE)
        colnames(payrollData)<- c(paste("payrolldollarsfy", theYears[i]-4,"per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, payrollData, sheet = theSheetName, startRow = 1, startCol=4) ## fills in columns 4, 5 and 6
        ##
        # fte_4, fte_3, fte_2 Pattern2 = IFR(7,12,17)  
        fte_4 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=7, endRow = theYearRowEnd[i], endCol=7, header=FALSE)
        colnames(fte_4) <- c(paste("ftefy", theYears[i]-4, "per", theYears[i], sep=""))
        fte_3 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=12, endRow = theYearRowEnd[i], endCol=12, header=FALSE)
        colnames(fte_3) <- c(paste("ftefy", theYears[i]-3,"per", theYears[i], sep=""))
        fte_2 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=17, endRow = theYearRowEnd[i], endCol=17, header=FALSE)
        colnames(fte_2) <- c(paste("ftefy", theYears[i]-2, "per", theYears[i], sep=""))
        ##
        # claims_4, claims_3, claims_2 Pattern2 = IFR(22,23,24) 
        theClaims <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=22, endRow = theYearRowEnd[i], endCol=24, header=FALSE)
        colnames(theClaims) <- c(paste("claimsfy", theYears[i]-4,"per", theYears[i], sep=""), paste("claimsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-2, "per", theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, fte_4, sheet= theSheetName, startRow = 1, startCol=7) ## fills in col 7
        writeWorksheet(theExtractWorkbook, fte_3 , sheet= theSheetName, startRow = 1, startCol=8) ## fills in col 8
        writeWorksheet(theExtractWorkbook, fte_2, sheet = theSheetName, startRow = 1, startCol=9) ## fills in col 9
        writeWorksheet(theExtractWorkbook, theClaims, sheet= theSheetName, startRow=1, startCol=10) ## fills in col 10,11,12
        ##
        # costs_4, costs_3, costs_2 Pattern = costs(3,4,5)
        costData <- readWorksheet(theSourceWorkbook, sheet="costs", startRow=5, startCol=3, endRow= theYearRowEnd[i], endCol=5, header=FALSE)
        colnames(costData) <- c(paste("costdollarsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, costData, sheet= theSheetName, startRow=1, startCol=13) ## fills in col 13, 14, 15
        ##
        # SORM cap on assessment based on numbers of claims; this is the column named difference as found in column Z
        capOnNumOfClaims <- readWorksheet(theSourceWorkbook, sheet="claims", startRow = 5, startCol=26, endRow = theYearRowEnd[i], endCol = 26, header=FALSE)
        colnames(capOnNumOfClaims) <- c(paste("capnumclaimsfy",theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, capOnNumOfClaims, sheet=theSheetName, startRow = 1, startCol = 16) ## fills column 16
        ##
        # SORM cap on assessment based on claims cost; this is the column named difference as found in column V
        capOnClamsCost <- readWorksheet(theSourceWorkbook, sheet="costs", startRow = 5, startCol=20, endRow = theYearRowEnd[i], endCol = 20, header=FALSE)
        colnames(capOnClamsCost) <- c(paste("capclaimscostfy",theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, capOnClamsCost, sheet=theSheetName, startRow = 1, startCol = 17) ## fills column 17
        ##
        saveWorkbook(theExtractWorkbook)
        
}
extract_initial_2005 <- function(){  
        theYears <- c(2005)
        theYearRowEnd <- c(298)
        for(i in 1:length(theYears)){
                theExtractFile <- paste(baseDir,"/workbooks/initial_extracts/",theYears[i],"_initial_extract.xls", sep = "")
                theSourceFile <- paste(baseDir,"/workbooks/initial_copies/",theYears[i],"_initial.xls", sep = "")        
                ##
                theExtractWorkbook <- loadWorkbook(theExtractFile, create = TRUE) ## create a new empty file, if file does not exist
                theSheetName <- paste(theYears[i],"_data", sep="")
                createSheet(theExtractWorkbook,name = theSheetName)
                theSourceWorkbook <- loadWorkbook(theSourceFile, create = FALSE) ## load in the SORM workbook for further reading
                ##
                # agency code, name, assessment   Pattern3 = assessment(2,3,11)
                agencyCodeAndName <- readWorksheet(theSourceWorkbook, sheet="assessment", startRow = 5, startCol=2, endRow = theYearRowEnd[i], endCol = 3, header=FALSE)
                colnames(agencyCodeAndName) <- c("code", "name")
                writeWorksheet(theExtractWorkbook, agencyCodeAndName, sheet=theSheetName, startRow = 1, startCol = 1) ## fills columns 1, 2
                theAssessment <- readWorksheet(theSourceWorkbook, sheet="assessment", startRow=5, startCol=11, endRow = theYearRowEnd[i], endCol = 11, header=FALSE)
                colnames(theAssessment) <- c(paste("assessmentdollars",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, theAssessment, sheet= theSheetName, startRow=1, startCol=3) ## fills column 3
                ##
                # payroll_4, payroll_3, payroll_2 Pattern1 = payroll(4,5,6)
                payrollData <- readWorksheet(theSourceWorkbook, sheet="payroll", startRow=5, startCol=4, endRow= theYearRowEnd[i], endCol=6, header=FALSE)
                colnames(payrollData)<- c(paste("payrolldollarsfy", theYears[i]-4,"per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, payrollData, sheet = theSheetName, startRow = 1, startCol=4) ## fills in columns 4, 5 and 6
                ##
                ##
                # fte_4, fte_3, fte_2 Pattern2 = IFR(8,13,18)  
                fte_4 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=8, endRow = theYearRowEnd[i], endCol=8, header=FALSE)
                colnames(fte_4) <- c(paste("ftefy", theYears[i]-4, "per", theYears[i], sep=""))
                fte_3 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=13, endRow = theYearRowEnd[i], endCol=13, header=FALSE)
                colnames(fte_3) <- c(paste("ftefy", theYears[i]-3,"per", theYears[i], sep=""))
                fte_2 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=18, endRow = theYearRowEnd[i], endCol=18, header=FALSE)
                colnames(fte_2) <- c(paste("ftefy", theYears[i]-2, "per", theYears[i], sep=""))
                ##
                # claims_4, claims_3, claims_2 Pattern2 = IFR(23,24,25)
                theClaims <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=23, endRow = theYearRowEnd[i], endCol=25, header=FALSE)
                colnames(theClaims) <- c(paste("claimsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, fte_4, sheet= theSheetName, startRow = 1, startCol=7) ## fills in col 7
                writeWorksheet(theExtractWorkbook, fte_3 , sheet= theSheetName, startRow = 1, startCol=8) ## fills in col 8
                writeWorksheet(theExtractWorkbook, fte_2, sheet = theSheetName, startRow = 1, startCol=9) ## fills in col 9
                writeWorksheet(theExtractWorkbook, theClaims, sheet= theSheetName, startRow=1, startCol=10) ## fills in col 10,11,12
                ##
                # costs_4, costs_3, costs_2 Pattern = costs(4,5,6)  
                costData <- readWorksheet(theSourceWorkbook, sheet="costs", startRow=5, startCol=4, endRow= theYearRowEnd[i], endCol=6, header=FALSE)
                colnames(costData) <- c(paste("costdollarsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, costData, sheet= theSheetName, startRow=1, startCol=13) ## fills in col 13
                ##
                # SORM cap on assessment based on numbers of claims; this is the column named difference as found in column AA
                capOnNumOfClaims <- readWorksheet(theSourceWorkbook, sheet="claims", startRow = 5, startCol=27, endRow = theYearRowEnd[i], endCol = 27, header=FALSE)
                colnames(capOnNumOfClaims) <- c(paste("capnumclaimsfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnNumOfClaims, sheet=theSheetName, startRow = 1, startCol = 16) ## fills column 16
                ##
                # SORM cap on assessment based on claims cost; this is the column named difference as found in column U
                capOnClamsCost <- readWorksheet(theSourceWorkbook, sheet="costs", startRow = 5, startCol=21, endRow = theYearRowEnd[i], endCol = 21, header=FALSE)
                colnames(capOnClamsCost) <- c(paste("capclaimscostfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnClamsCost, sheet=theSheetName, startRow = 1, startCol = 17) ## fills column 17
                ##
                # 
                saveWorkbook(theExtractWorkbook)
        }        
}
extract_2006_2007 <- function(whichone){  
        theYears <- 2006:2007
        theYearRowEnd <- c(282,277)
        for(i in 1:length(theYears)){
                if(whichone == "final"){
                        theExtractFile <- paste(baseDir,"/workbooks/final_extracts/",theYears[i],"_final_extract.xls", sep = "")
                        theSourceFile <- paste(baseDir,"/workbooks/final_copies/",theYears[i],"_final.xls", sep = "") 
                }
                if(whichone == "initial"){
                        theExtractFile <- paste(baseDir,"/workbooks/initial_extracts/",theYears[i],"_initial_extract.xls", sep = "")
                        theSourceFile <- paste(baseDir,"/workbooks/initial_copies/",theYears[i],"_initial.xls", sep = "") 
                }
                
                ##
                theExtractWorkbook <- loadWorkbook(theExtractFile, create = TRUE) ## create a new empty file, if file does not exist
                theSheetName <- paste(theYears[i],"_data", sep="")
                createSheet(theExtractWorkbook,name = theSheetName)
                theSourceWorkbook <- loadWorkbook(theSourceFile, create = FALSE) ## load in the SORM workbook for further reading
                ##
                # agency code, name, assessment   Pattern3 = assessment(2,3,11)
                agencyCodeAndName <- readWorksheet(theSourceWorkbook, sheet="assessment", startRow = 5, startCol=2, endRow = theYearRowEnd[i], endCol = 3, header=FALSE)
                colnames(agencyCodeAndName) <- c("code", "name")
                writeWorksheet(theExtractWorkbook, agencyCodeAndName, sheet=theSheetName, startRow = 1, startCol = 1) ## fills columns 1, 2
                theAssessment <- readWorksheet(theSourceWorkbook, sheet="assessment", startRow=5, startCol=11, endRow = theYearRowEnd[i], endCol = 11, header=FALSE)
                colnames(theAssessment) <- c(paste("assessmentdollars",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, theAssessment, sheet= theSheetName, startRow=1, startCol=3) ## fills column 3
                ##
                # payroll_4, payroll_3, payroll_2 Pattern1 = payroll(4,5,6)
                payrollData <- readWorksheet(theSourceWorkbook, sheet="payroll", startRow=5, startCol=4, endRow= theYearRowEnd[i], endCol=6, header=FALSE)
                colnames(payrollData)<- c(paste("payrolldollarsfy", theYears[i]-4,"per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, payrollData, sheet = theSheetName, startRow = 1, startCol=4) ## fills in columns 4, 5 and 6
                ##
                ##
                # fte_4, fte_3, fte_2 Pattern2 = IFR(8,13,18)  
                fte_4 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=8, endRow = theYearRowEnd[i], endCol=8, header=FALSE)
                colnames(fte_4) <- c(paste("ftefy", theYears[i]-4, "per", theYears[i], sep=""))
                fte_3 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=13, endRow = theYearRowEnd[i], endCol=13, header=FALSE)
                colnames(fte_3) <- c(paste("ftefy", theYears[i]-3,"per", theYears[i], sep=""))
                fte_2 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=18, endRow = theYearRowEnd[i], endCol=18, header=FALSE)
                colnames(fte_2) <- c(paste("ftefy", theYears[i]-2, "per", theYears[i], sep=""))
                ##
                # claims_4, claims_3, claims_2 Pattern2 = IFR(23,24,25)
                theClaims <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=23, endRow = theYearRowEnd[i], endCol=25, header=FALSE)
                colnames(theClaims) <- c(paste("claimsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, fte_4, sheet= theSheetName, startRow = 1, startCol=7) ## fills in col 7
                writeWorksheet(theExtractWorkbook, fte_3 , sheet= theSheetName, startRow = 1, startCol=8) ## fills in col 8
                writeWorksheet(theExtractWorkbook, fte_2, sheet = theSheetName, startRow = 1, startCol=9) ## fills in col 9
                writeWorksheet(theExtractWorkbook, theClaims, sheet= theSheetName, startRow=1, startCol=10) ## fills in col 10,11,12
                ##
                # costs_4, costs_3, costs_2 Pattern = costs(4,5,6)  
                costData <- readWorksheet(theSourceWorkbook, sheet="costs", startRow=5, startCol=4, endRow= theYearRowEnd[i], endCol=6, header=FALSE)
                colnames(costData) <- c(paste("costdollarsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, costData, sheet= theSheetName, startRow=1, startCol=13) ## fills in col 13
                ##
                # SORM cap on assessment based on numbers of claims; this is the column named difference as found in column AA
                capOnNumOfClaims <- readWorksheet(theSourceWorkbook, sheet="claims", startRow = 5, startCol=27, endRow = theYearRowEnd[i], endCol = 27, header=FALSE)
                colnames(capOnNumOfClaims) <- c(paste("capnumclaimsfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnNumOfClaims, sheet=theSheetName, startRow = 1, startCol = 16) ## fills column 16
                ##
                # SORM cap on assessment based on claims cost; this is the column named difference as found in column U
                capOnClamsCost <- readWorksheet(theSourceWorkbook, sheet="costs", startRow = 5, startCol=21, endRow = theYearRowEnd[i], endCol = 21, header=FALSE)
                colnames(capOnClamsCost) <- c(paste("capclaimscostfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnClamsCost, sheet=theSheetName, startRow = 1, startCol = 17) ## fills column 17
                ##
                # 
                saveWorkbook(theExtractWorkbook)
        }        
}
extract_2008 <- function(whichone){ 
        ## library(XLConnect)
        theYears <- 2008
        theYearRowEnd <- c(264)
        i <- 1
        if(whichone == "final"){
                theExtractFile <- paste(baseDir,"/workbooks/final_extracts/",theYears[i],"_final_extract.xls", sep = "")
                theSourceFile <- paste(baseDir,"/workbooks/final_copies/",theYears[i],"_final.xls", sep = "") 
        }
        if(whichone == "initial"){
                theExtractFile <- paste(baseDir,"/workbooks/initial_extracts/",theYears[i],"_initial_extract.xls", sep = "")
                theSourceFile <- paste(baseDir,"/workbooks/initial_copies/",theYears[i],"_initial.xls", sep = "") 
        }
        
        theExtractWorkbook <- loadWorkbook(theExtractFile, create = TRUE) ## create a new empty file, if file does not exist
        theSheetName <- paste(theYears[i],"_data", sep="")
        createSheet(theExtractWorkbook,name = theSheetName) 
        theSourceWorkbook <- loadWorkbook(theSourceFile, create = FALSE) ## load in the SORM workbook for further reading
        ##
        # agency code, name, assessment   Pattern2 = Invoice(2,3,7) 
        if (whichone == "final"){
                agencyCodeAndName <- readWorksheet(theSourceWorkbook, sheet="invoices", startRow = 5, startCol=2, endRow = theYearRowEnd[i], endCol = 3, header=FALSE)
                theAssessment <- readWorksheet(theSourceWorkbook, sheet="invoices", startRow=5, startCol=7, endRow = theYearRowEnd[i], endCol = 7, header=FALSE)
                
        }else{
                agencyCodeAndName <- readWorksheet(theSourceWorkbook, sheet="initial invoices", startRow = 5, startCol=2, endRow = theYearRowEnd[i], endCol = 3, header=FALSE)
                theAssessment <- readWorksheet(theSourceWorkbook, sheet="initial invoices", startRow=5, startCol=7, endRow = theYearRowEnd[i], endCol = 7, header=FALSE)
                
        }
        colnames(agencyCodeAndName) <- c("code", "name")
        writeWorksheet(theExtractWorkbook, agencyCodeAndName, sheet=theSheetName, startRow = 1, startCol = 1) ## filles columns 1, 2
        colnames(theAssessment) <- c(paste("assessmentdollarsfy",theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, theAssessment, sheet= theSheetName, startRow=1, startCol=3) ## fills column 3
        ##
        # payroll_4, payroll_3, payroll_2 Pattern1 = payroll(4,5,6) 
        payrollData <- readWorksheet(theSourceWorkbook, sheet="payroll", startRow=5, startCol=4, endRow= theYearRowEnd[i], endCol=6, header=FALSE)
        colnames(payrollData)<- c(paste("payrolldollarsfy", theYears[i]-4,"per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, payrollData, sheet = theSheetName, startRow = 1, startCol=4) ## fills in columns 4, 5 and 6
        ##
        # fte_4, fte_3, fte_2 Pattern2 = IFR(8,13,18) 
        fte_4 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=8, endRow = theYearRowEnd[i], endCol=8, header=FALSE)
        colnames(fte_4) <- c(paste("ftefy", theYears[i]-4, "per", theYears[i], sep=""))
        fte_3 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=13, endRow = theYearRowEnd[i], endCol=13, header=FALSE)
        colnames(fte_3) <- c(paste("ftefy", theYears[i]-3,"per", theYears[i], sep=""))
        fte_2 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=18, endRow = theYearRowEnd[i], endCol=18, header=FALSE)
        colnames(fte_2) <- c(paste("ftefy", theYears[i]-2, "per", theYears[i], sep=""))
        ##
        # claims_4, claims_3, claims_2 Pattern2 = IFR(23,24,25) 
        theClaims <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=23, endRow = theYearRowEnd[i], endCol=25, header=FALSE)
        colnames(theClaims) <- c(paste("claimsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-2, "per", theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, fte_4, sheet= theSheetName, startRow = 1, startCol=7) ## fills in col 7
        writeWorksheet(theExtractWorkbook, fte_3 , sheet= theSheetName, startRow = 1, startCol=8) ## fills in col 8
        writeWorksheet(theExtractWorkbook, fte_2, sheet = theSheetName, startRow = 1, startCol=9) ## fills in col 9
        writeWorksheet(theExtractWorkbook, theClaims, sheet= theSheetName, startRow=1, startCol=10) ## fills in col 10,11,12
        ##
        # costs_4, costs_3, costs_2 Pattern = costs(4,5,6)
        costData <- readWorksheet(theSourceWorkbook, sheet="costs", startRow=5, startCol=4, endRow= theYearRowEnd[i], endCol=6, header=FALSE)
        colnames(costData) <- c(paste("costdollarsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, costData, sheet= theSheetName, startRow=1, startCol=13) ## fills in col 13, 14, 15
        ##
        # SORM cap on assessment based on numbers of claims; this is the column named difference as found in column AA
        capOnNumOfClaims <- readWorksheet(theSourceWorkbook, sheet="claims", startRow = 5, startCol=27, endRow = theYearRowEnd[i], endCol = 27, header=FALSE)
        colnames(capOnNumOfClaims) <- c(paste("capnumclaimsfy",theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, capOnNumOfClaims, sheet=theSheetName, startRow = 1, startCol = 16) ## fills column 16
        ##
        # SORM cap on assessment based on claims cost; this is the column named difference as found in column U
        capOnClamsCost <- readWorksheet(theSourceWorkbook, sheet="costs", startRow = 5, startCol=21, endRow = theYearRowEnd[i], endCol = 21, header=FALSE)
        colnames(capOnClamsCost) <- c(paste("capclaimscostfy",theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, capOnClamsCost, sheet=theSheetName, startRow = 1, startCol = 17) ## fills column 17
        ##
        #
        saveWorkbook(theExtractWorkbook)
}
extract_initial_2009 <- function(){  
        theYears <-2009
        theYearRowEnd <- c(268)
        for(i in 1:length(theYears)){
                theExtractFile <- paste(baseDir,"/workbooks/initial_extracts/",theYears[i],"_initial_extract.xls", sep = "")
                theSourceFile <- paste(baseDir,"/workbooks/initial_copies/",theYears[i],"_initial.xls", sep = "") 
                
                theExtractWorkbook <- loadWorkbook(theExtractFile, create = TRUE) ## create a new empty file
                theSheetName <- paste(theYears[i],"_data", sep="")
                createSheet(theExtractWorkbook,name = theSheetName) 
                theSourceWorkbook <- loadWorkbook(theSourceFile, create = FALSE) ## load in the SORM workbook for further reading
                ##
                agencyCodeAndName <- readWorksheet(theSourceWorkbook, sheet="invoices", startRow = 5, startCol=2, endRow = theYearRowEnd[i], endCol = 3, header=FALSE)
                colnames(agencyCodeAndName) <- c("code", "name")
                writeWorksheet(theExtractWorkbook, agencyCodeAndName, sheet=theSheetName, startRow = 1, startCol = 1)   # fills in col 1 and 2
                theAssessment <- readWorksheet(theSourceWorkbook, sheet="invoices", startRow=5, startCol=7, endRow = theYearRowEnd[i], endCol = 7, header=FALSE)
                colnames(theAssessment) <- c(paste("assessmentdollars",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, theAssessment, sheet= theSheetName, startRow=1, startCol=3)          # fills in col 3
                ##
                payrollData <- readWorksheet(theSourceWorkbook, sheet="payroll", startRow=5, startCol=4, endRow= theYearRowEnd[i], endCol=6, header=FALSE)
                colnames(payrollData)<- c(paste("payrolldollarsfy", theYears[i]-4,"per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, payrollData, sheet = theSheetName, startRow = 1, startCol=4)        # fills in columns 4, 5 and 6
                ##
                fte_4 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=7, endRow = theYearRowEnd[i], endCol=8, header=FALSE)
                colnames(fte_4) <- c(paste("ftefy", theYears[i]-4, "per", theYears[i], sep=""))
                fte_3 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=12, endRow = theYearRowEnd[i], endCol=13, header=FALSE)
                colnames(fte_3) <- c(paste("ftefy", theYears[i]-3,"per", theYears[i], sep=""))
                fte_2 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=17, endRow = theYearRowEnd[i], endCol=18, header=FALSE)
                colnames(fte_2) <- c(paste("ftefy", theYears[i]-2, "per", theYears[i], sep=""))
                ##
                theClaims <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=23, endRow = theYearRowEnd[i], endCol=25, header=FALSE)
                colnames(theClaims) <- c(paste("claimsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, fte_4, sheet= theSheetName, startRow = 1, startCol=7) ## fills in col 7
                writeWorksheet(theExtractWorkbook, fte_3 , sheet= theSheetName, startRow = 1, startCol=8) ## fills in col 8
                writeWorksheet(theExtractWorkbook, fte_2, sheet = theSheetName, startRow = 1, startCol=9) ## fills in col 9
                writeWorksheet(theExtractWorkbook, theClaims, sheet= theSheetName, startRow=1, startCol=10) ## fills in col 10,11,12
                ##
                costData <- readWorksheet(theSourceWorkbook, sheet="costs", startRow=5, startCol=4, endRow= theYearRowEnd[i], endCol=6, header=FALSE)
                colnames(costData) <- c(paste("costdollarsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, costData, sheet= theSheetName, startRow=1, startCol=13) ## fills in col 13,14,15
                ##
                capOnNumOfClaims <- readWorksheet(theSourceWorkbook, sheet="claims", startRow = 5, startCol=27, endRow = theYearRowEnd[i], endCol = 27, header=FALSE)
                colnames(capOnNumOfClaims) <- c(paste("capnumclaimsfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnNumOfClaims, sheet=theSheetName, startRow = 1, startCol = 16) ## fills column 16
                ##
                capOnClamsCost <- readWorksheet(theSourceWorkbook, sheet="costs", startRow = 5, startCol=21, endRow = theYearRowEnd[i], endCol = 21, header=FALSE)
                colnames(capOnClamsCost) <- c(paste("capclaimscostfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnClamsCost, sheet=theSheetName, startRow = 1, startCol = 17) ## fills column 17
                ## ##
                saveWorkbook(theExtractWorkbook)
        }
}
extract_initial_2010 <- function(){  
        theYears <-2010
        theYearRowEnd <- c(269)
        for(i in 1:length(theYears)){
                theExtractFile <- paste(baseDir,"/workbooks/initial_extracts/",theYears[i],"_initial_extract.xls", sep = "")
                theSourceFile <- paste(baseDir,"/workbooks/initial_copies/",theYears[i],"_initial.xls", sep = "") 
                
                theExtractWorkbook <- loadWorkbook(theExtractFile, create = TRUE) ## create a new empty file
                theSheetName <- paste(theYears[i],"_data", sep="")
                createSheet(theExtractWorkbook,name = theSheetName) 
                theSourceWorkbook <- loadWorkbook(theSourceFile, create = FALSE) ## load in the SORM workbook for further reading
                ##
                agencyCodeAndName <- readWorksheet(theSourceWorkbook, sheet="invoices", startRow = 5, startCol=1, endRow = theYearRowEnd[i], endCol = 2, header=FALSE)
                colnames(agencyCodeAndName) <- c("code", "name")
                writeWorksheet(theExtractWorkbook, agencyCodeAndName, sheet=theSheetName, startRow = 1, startCol = 1)   # fills in col 1 and 2
                theAssessment <- readWorksheet(theSourceWorkbook, sheet="invoices", startRow=5, startCol=6, endRow = theYearRowEnd[i], endCol = 6, header=FALSE)
                colnames(theAssessment) <- c(paste("assessmentdollars",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, theAssessment, sheet= theSheetName, startRow=1, startCol=3)          # fills in col 3
                ##
                payrollData <- readWorksheet(theSourceWorkbook, sheet="payroll", startRow=5, startCol=3, endRow= theYearRowEnd[i], endCol=5, header=FALSE)
                colnames(payrollData)<- c(paste("payrolldollarsfy", theYears[i]-4,"per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, payrollData, sheet = theSheetName, startRow = 1, startCol=4)        # fills in columns 4, 5 and 6
                ##
                fte_4 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=7, endRow = theYearRowEnd[i], endCol=7, header=FALSE)
                colnames(fte_4) <- c(paste("ftefy", theYears[i]-4, "per", theYears[i], sep=""))
                fte_3 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=12, endRow = theYearRowEnd[i], endCol=12, header=FALSE)
                colnames(fte_3) <- c(paste("ftefy", theYears[i]-3,"per", theYears[i], sep=""))
                fte_2 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=17, endRow = theYearRowEnd[i], endCol=17, header=FALSE)
                colnames(fte_2) <- c(paste("ftefy", theYears[i]-2, "per", theYears[i], sep=""))
                ##
                theClaims <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=22, endRow = theYearRowEnd[i], endCol=24, header=FALSE)
                colnames(theClaims) <- c(paste("claimsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, fte_4, sheet= theSheetName, startRow = 1, startCol=7) ## fills in col 7
                writeWorksheet(theExtractWorkbook, fte_3 , sheet= theSheetName, startRow = 1, startCol=8) ## fills in col 8
                writeWorksheet(theExtractWorkbook, fte_2, sheet = theSheetName, startRow = 1, startCol=9) ## fills in col 9
                writeWorksheet(theExtractWorkbook, theClaims, sheet= theSheetName, startRow=1, startCol=10) ## fills in col 10,11,12
                ##
                costData <- readWorksheet(theSourceWorkbook, sheet="costs", startRow=5, startCol=3, endRow= theYearRowEnd[i], endCol=5, header=FALSE)
                colnames(costData) <- c(paste("costdollarsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, costData, sheet= theSheetName, startRow=1, startCol=13) ## fills in col 13,14,15
                ##
                capOnNumOfClaims <- readWorksheet(theSourceWorkbook, sheet="claims", startRow = 5, startCol=26, endRow = theYearRowEnd[i], endCol = 26, header=FALSE)
                colnames(capOnNumOfClaims) <- c(paste("capnumclaimsfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnNumOfClaims, sheet=theSheetName, startRow = 1, startCol = 16) ## fills column 16
                ##
                capOnClamsCost <- readWorksheet(theSourceWorkbook, sheet="costs", startRow = 5, startCol=20, endRow = theYearRowEnd[i], endCol = 20, header=FALSE)
                colnames(capOnClamsCost) <- c(paste("capclaimscostfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnClamsCost, sheet=theSheetName, startRow = 1, startCol = 17) ## fills column 17
                ## ##
                saveWorkbook(theExtractWorkbook)
        }
}
extract_initial_2011_2015 <- function(){  
        theYears <-2011:2015
        theYearRowEnd <- c(270, 265, 262, 262, 263)
        for(i in 1:length(theYears)){
                if( theYears[i]  == 2011 | theYears[i] == 2012){
                        theExtractFile <- paste(baseDir,"/workbooks/initial_extracts/",theYears[i],"_initial_extract.xls", sep = "")
                        theSourceFile <- paste(baseDir,"/workbooks/initial_copies/",theYears[i],"_initial.xls", sep = "") 
                }
                
                if(theYears[i] == 2013 | theYears[i] == 2014 | theYears[i] == 2015){
                        theExtractFile <- paste(baseDir,"/workbooks/initial_extracts/",theYears[i],"_initial_extract.xlsx", sep = "")
                        theSourceFile <- paste(baseDir,"/workbooks/initial_copies/",theYears[i],"_initial.xlsx", sep = "") 
                }
                theExtractWorkbook <- loadWorkbook(theExtractFile, create = TRUE) ## create a new empty file
                theSheetName <- paste(theYears[i],"_data", sep="")
                createSheet(theExtractWorkbook,name = theSheetName) 
                theSourceWorkbook <- loadWorkbook(theSourceFile, create = FALSE) ## load in the SORM workbook for further reading
                ##
                agencyCodeAndName <- readWorksheet(theSourceWorkbook, sheet="invoices", startRow = 5, startCol=1, endRow = theYearRowEnd[i], endCol = 2, header=FALSE)
                colnames(agencyCodeAndName) <- c("code", "name")
                writeWorksheet(theExtractWorkbook, agencyCodeAndName, sheet=theSheetName, startRow = 1, startCol = 1)   # fills in col 1 and 2
                theAssessment <- readWorksheet(theSourceWorkbook, sheet="invoices", startRow=5, startCol=6, endRow = theYearRowEnd[i], endCol = 6, header=FALSE)
                colnames(theAssessment) <- c(paste("assessmentdollars",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, theAssessment, sheet= theSheetName, startRow=1, startCol=3)          # fills in col 3
                ##
                payrollData <- readWorksheet(theSourceWorkbook, sheet="payroll", startRow=5, startCol=3, endRow= theYearRowEnd[i], endCol=5, header=FALSE)
                colnames(payrollData)<- c(paste("payrolldollarsfy", theYears[i]-4,"per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, payrollData, sheet = theSheetName, startRow = 1, startCol=4)        # fills in columns 4, 5 and 6
                ##
                fte_4 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=7, endRow = theYearRowEnd[i], endCol=7, header=FALSE)
                colnames(fte_4) <- c(paste("ftefy", theYears[i]-4, "per", theYears[i], sep=""))
                fte_3 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=12, endRow = theYearRowEnd[i], endCol=12, header=FALSE)
                colnames(fte_3) <- c(paste("ftefy", theYears[i]-3,"per", theYears[i], sep=""))
                fte_2 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=17, endRow = theYearRowEnd[i], endCol=17, header=FALSE)
                colnames(fte_2) <- c(paste("ftefy", theYears[i]-2, "per", theYears[i], sep=""))
                ##
                theClaims <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=22, endRow = theYearRowEnd[i], endCol=24, header=FALSE)
                colnames(theClaims) <- c(paste("claimsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, fte_4, sheet= theSheetName, startRow = 1, startCol=7) ## fills in col 7
                writeWorksheet(theExtractWorkbook, fte_3 , sheet= theSheetName, startRow = 1, startCol=8) ## fills in col 8
                writeWorksheet(theExtractWorkbook, fte_2, sheet = theSheetName, startRow = 1, startCol=9) ## fills in col 9
                writeWorksheet(theExtractWorkbook, theClaims, sheet= theSheetName, startRow=1, startCol=10) ## fills in col 10,11,12
                ##
                costData <- readWorksheet(theSourceWorkbook, sheet="costs", startRow=5, startCol=3, endRow= theYearRowEnd[i], endCol=5, header=FALSE)
                colnames(costData) <- c(paste("costdollarsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, costData, sheet= theSheetName, startRow=1, startCol=13) ## fills in col 13,14,15
                ##
                capOnNumOfClaims <- readWorksheet(theSourceWorkbook, sheet="claims", startRow = 5, startCol=26, endRow = theYearRowEnd[i], endCol = 26, header=FALSE)
                colnames(capOnNumOfClaims) <- c(paste("capnumclaimsfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnNumOfClaims, sheet=theSheetName, startRow = 1, startCol = 16) ## fills column 16
                ##
                capOnClamsCost <- readWorksheet(theSourceWorkbook, sheet="costs", startRow = 5, startCol=20, endRow = theYearRowEnd[i], endCol = 20, header=FALSE)
                colnames(capOnClamsCost) <- c(paste("capclaimscostfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnClamsCost, sheet=theSheetName, startRow = 1, startCol = 17) ## fills column 17
                ## ##
                saveWorkbook(theExtractWorkbook)
        }
}
extract_final_2004 <- function(){ 
        theYears <-2004
        theYearRowEnd <- 1168
        i <- 1
        ## this workbook data starts on row 10
        ## agency code is non-standard
        ##
        theExtractFile <- paste(baseDir,"/workbooks/final_extracts/",theYears[i],"_final_extract.xls", sep = "")
        theExtractWorkbook <- loadWorkbook(theExtractFile, create = TRUE) ## create a new empty file, if file does not exist
        theSheetName <- paste(theYears[i],"_data", sep="")
        createSheet(theExtractWorkbook,name = theSheetName) 
        theSourceFile <- paste(baseDir,"/workbooks/final_copies/",theYears[i],"_final.xls", sep = "") 
        theSourceWorkbook <- loadWorkbook(theSourceFile, create = FALSE) ## load in the SORM workbook for further reading
        ##
        agencyCodeAndName <- readWorksheet(theSourceWorkbook, sheet="consolidated", startRow = 10, startCol=1, endRow = theYearRowEnd[i], endCol = 2, header=FALSE)
        colnames(agencyCodeAndName) <- c("code", "name")
        writeWorksheet(theExtractWorkbook, agencyCodeAndName, sheet=theSheetName, startRow = 1, startCol = 1) ## fills columns 1, 2
        ##
        theAssessment <- readWorksheet(theSourceWorkbook, sheet="consolidated", startRow=10, startCol=12, endRow = theYearRowEnd[i], endCol = 12, header=FALSE)
        colnames(theAssessment) <- c(paste("assessmentdollars", theYears[i], sep=""))
        writeWorksheet(theExtractWorkbook, theAssessment, sheet= theSheetName, startRow=1, startCol=3) ## fills column 3
        saveWorkbook(theExtractWorkbook)
}
extract_final_2009_2015 <- function(){ 
        theYears <-2009:2015
        theYearRowEnd <- c(268, 266, 266, 266, 262, 262,264)
        for(i in 1:length(theYears)){
                if(theYears[i] == 2012 | theYears[i] == 2013 | theYears[i] == 2014 | theYears[i] == 2015){
                        theExtractFile <- paste(baseDir,"/workbooks/final_extracts/",theYears[i],"_final_extract.xlsx", sep = "")
                        theSourceFile <- paste(baseDir,"/workbooks/final_copies/",theYears[i],"_final.xlsx", sep = "") 
                        
                }
                if(theYears[i] == 2009 | theYears[i]  == 2010 | theYears[i]  == 2011){
                        theExtractFile <- paste(baseDir,"/workbooks/final_extracts/",theYears[i],"_final_extract.xls", sep = "")
                        theSourceFile <- paste(baseDir,"/workbooks/final_copies/",theYears[i],"_final.xls", sep = "") 
                }
                theExtractWorkbook <- loadWorkbook(theExtractFile, create = TRUE) ## create a new empty file
                theSheetName <- paste(theYears[i],"_data", sep="")
                createSheet(theExtractWorkbook,name = theSheetName) 
                
                theSourceWorkbook <- loadWorkbook(theSourceFile, create = FALSE) ## load in the SORM workbook for further reading
                # agency code, name, assessment   Pattern1 = Invoice(1,2,6)
                agencyCodeAndName <- readWorksheet(theSourceWorkbook, sheet="invoices", startRow = 5, startCol=1, endRow = theYearRowEnd[i], endCol = 2, header=FALSE)
                colnames(agencyCodeAndName) <- c("code", "name")
                writeWorksheet(theExtractWorkbook, agencyCodeAndName, sheet=theSheetName, startRow = 1, startCol = 1)
                theAssessment <- readWorksheet(theSourceWorkbook, sheet="invoices", startRow=5, startCol=6, endRow = theYearRowEnd[i], endCol = 6, header=FALSE)
                colnames(theAssessment) <- c(paste("assessmentdollars",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, theAssessment, sheet= theSheetName, startRow=1, startCol=3)
                ##
                ### # payroll_4, payroll_3, payroll_2 Pattern1 = payroll(3,4,5)
                payrollData <- readWorksheet(theSourceWorkbook, sheet="payroll", startRow=5, startCol=3, endRow= theYearRowEnd[i], endCol=5, header=FALSE)
                colnames(payrollData)<- c(paste("payrolldollarsfy", theYears[i]-4,"per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("payrolldollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, payrollData, sheet = theSheetName, startRow = 1, startCol=4) ## fills in columns 4, 5 and 6
                ##
                # fte_4, fte_3, fte_2 Pattern1 = IFR(7,12,17)
                # claims_4, claims_3, claims_2 Pattern1 = IFR(22,23,24)  
                fte_4 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=7, endRow = theYearRowEnd[i], endCol=7, header=FALSE)
                colnames(fte_4) <- c(paste("ftefy", theYears[i]-4, "per", theYears[i], sep=""))
                fte_3 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=12, endRow = theYearRowEnd[i], endCol=12, header=FALSE)
                colnames(fte_3) <- c(paste("ftefy", theYears[i]-3,"per", theYears[i], sep=""))
                fte_2 <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=17, endRow = theYearRowEnd[i], endCol=17, header=FALSE)
                colnames(fte_2) <- c(paste("ftefy", theYears[i]-2, "per", theYears[i], sep=""))
                theClaims <- readWorksheet(theSourceWorkbook, sheet="IFR", startRow=5, startCol=22, endRow = theYearRowEnd[i], endCol=24, header=FALSE)
                colnames(theClaims) <- c(paste("claimsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("claimsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, fte_4, sheet= theSheetName, startRow = 1, startCol=7) ## fills in col 7
                writeWorksheet(theExtractWorkbook, fte_3 , sheet= theSheetName, startRow = 1, startCol=8) ## fills in col 8
                writeWorksheet(theExtractWorkbook, fte_2, sheet = theSheetName, startRow = 1, startCol=9) ## fills in col 9
                writeWorksheet(theExtractWorkbook, theClaims, sheet= theSheetName, startRow=1, startCol=10) ## fills in col 10,11,12
                ##
                # costs_4, costs_3, costs_2 Pattern = costs(3,4,5)
                costData <- readWorksheet(theSourceWorkbook, sheet="costs", startRow=5, startCol=3, endRow= theYearRowEnd[i], endCol=5, header=FALSE)
                colnames(costData) <- c(paste("costdollarsfy", theYears[i]-4, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-3, "per", theYears[i], sep=""), paste("costdollarsfy", theYears[i]-2, "per", theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, costData, sheet= theSheetName, startRow=1, startCol=13) ## fills in col 13
                ##
                # SORM cap on assessment based on numbers of claims; this is the column named difference as found in column Z
                capOnNumOfClaims <- readWorksheet(theSourceWorkbook, sheet="claims", startRow = 5, startCol=26, endRow = theYearRowEnd[i], endCol = 26, header=FALSE)
                colnames(capOnNumOfClaims) <- c(paste("capnumclaimsfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnNumOfClaims, sheet=theSheetName, startRow = 1, startCol = 16) ## fills column 16
                ##
                # SORM cap on assessment based on claims cost; this is the column named difference as found in column T
                capOnClamsCost <- readWorksheet(theSourceWorkbook, sheet="costs", startRow = 5, startCol=20, endRow = theYearRowEnd[i], endCol = 20, header=FALSE)
                colnames(capOnClamsCost) <- c(paste("capclaimscostfy",theYears[i], sep=""))
                writeWorksheet(theExtractWorkbook, capOnClamsCost, sheet=theSheetName, startRow = 1, startCol = 17) ## fills column 17
                ## ##
                saveWorkbook(theExtractWorkbook)
        }
}
regularize_2004_final_extract <- function(){
        theFile <- paste(baseDir,"/workbooks/final_extracts/", "2004_final_extract.xls", sep="")
        assessment_final_2004_df <- read.xlsx(theFile, sheetName="2004_data", 
                                              startRow=1, endRow=1160, 
                                              as.data.frame=TRUE, 
                                              header=TRUE)
        assessment_final_2004_df$code <- as.character(assessment_final_2004_df$code); assessment_final_2004_df$name <- as.character(assessment_final_2004_df$name); assessment_final_2004_df$assessmentdollars2004 <- as.numeric(assessment_final_2004_df$assessmentdollars2004) 
        for(i in 1:nrow(assessment_final_2004_df)){
                if(!is.na(assessment_final_2004_df$code[i])){
                        theCode <- assessment_final_2004_df$code[i]
                        theName <- assessment_final_2004_df$name[i]
                }else{
                        assessment_final_2004_df$code[i] <- theCode
                        assessment_final_2004_df$name[i]  <- theName
                }
        }
        assessment_final_2004_df$code <- paste("A", assessment_final_2004_df$code,sep="")
        assessment_final_2004_df$code <- as.factor(assessment_final_2004_df$code); assessment_final_2004_df$name <- as.factor(assessment_final_2004_df$name); assessment_final_2004_df$assessmentdollars2004 <- as.numeric(assessment_final_2004_df$assessmentdollars2004)
        assessment_final_2004_df <- assessment_final_2004_df[complete.cases(assessment_final_2004_df),]
        theSummary <- group_by(assessment_final_2004_df,code,name) %>%
                summarise(max(assessmentdollars2004))
        colnames(theSummary) <- c("code","name","assessmentdollars2004")
        write.csv(theSummary,paste(baseDir,"/processed_files/", "assessment_final_2004.csv", sep=""))
}
# ----------------------------------------------------- Combine workbooks:
combine_same_fy_workbooks <- function(){
        for(i in 2006:2015){
                if( i > 2011){
                        iniFile <- paste(baseDir, "/workbooks", "/initial_extracts/", i, "_initial_extract.xlsx",sep="")
                        finFile <- paste(baseDir, "/workbooks", "/final_extracts/", i, "_final_extract.xlsx",sep="")
                }else{
                        iniFile <- paste(baseDir, "/workbooks", "/initial_extracts/", i, "_initial_extract.xls",sep="")
                        finFile <- paste(baseDir, "/workbooks", "/final_extracts/", i, "_final_extract.xls",sep="")
                }
                if( i == 2012){
                        iniFile <- paste(baseDir, "/workbooks", "/initial_extracts/", i, "_initial_extract.xls",sep="")
                        finFile <- paste(baseDir, "/workbooks", "/final_extracts/", i, "_final_extract.xlsx",sep="")
                }
                
                ini_year <- read.xlsx(iniFile, sheetName=paste(i,"_data", sep=""), 
                                      #startRow=1, endRow=1160, 
                                      as.data.frame=TRUE, 
                                      header=TRUE)
                fin_year <- read.xlsx(finFile, sheetName=paste(i,"_data", sep=""), 
                                      #startRow=1, endRow=1160, 
                                      as.data.frame=TRUE, 
                                      header=TRUE)
                fin_year$wbhistory <- "final"; fin_year$wbhistory <- as.character(fin_year$wbhistory) # record history of row
                fin_year$code <- as.character(fin_year$code); ini_year$code <- as.character(ini_year$code)
                if(TRUE %in% is.na(fin_year$code)){
                        fin_year <- fin_year[-which(is.na(fin_year$code)), ]   # need to remove any blank lines from dataframe here
                }
                for(j in 1:nrow(fin_year)){
                        if(fin_year$code[j] %in% ini_year$code){
                                fin_year$wbhistory[j]  <- "updated"
                        }
                }
                onlyInInitial <- ini_year[!ini_year$code %in% fin_year$code,]
                if(TRUE %in% is.na(onlyInInitial$code)){
                        onlyInInitial <- onlyInInitial[-which(is.na(onlyInInitial$code)), ]   # need to remove any blank lines from dataframe here
                }
                if(nrow(onlyInInitial) > 0){
                        # need to copy info from ini_year  to fin_year  dataframe
                        onlyInInitial$wbhistory <- "initial"; onlyInInitial$wbhistory <- as.character(onlyInInitial$wbhistory)
                        fin_year <- rbind(fin_year,onlyInInitial)
                        fin_year <- arrange(fin_year,code)
                }
                if(i == 2006){
                        # put code here to delete duplicated client rows for "A452" "A551" "A582"; its the 2nd ocurrance that should be deleted
                        whichduplicatedcodes <- c("A452", "A551", "A582")
                        for(q in 1:length(whichduplicatedcodes)){
                                fin_year <- fin_year[ -max(which(fin_year$code == whichduplicatedcodes[q],)),]   
                        }       
                }
                if(i == 2007){
                        # put code here to delete duplicated client rows for  "A551" "A582"; its the 2nd ocurrance that should be deleted
                        whichduplicatedcodes <- c("A551","A582")
                        for(q in 1:length(whichduplicatedcodes)){
                                fin_year <- fin_year[ -max(which(fin_year$code == whichduplicatedcodes[q],)),]   
                        }
                }
                if(i == 2009){
                        # put code here to delete duplicated client rows for  "A551"; its the 2nd ocurrance that should be deleted
                        whichduplicatedcodes <- c("A551")
                        for(q in 1:length(whichduplicatedcodes)){
                                fin_year <- fin_year[ -max(which(fin_year$code == whichduplicatedcodes[q],)),]   
                        }
                }
                write.csv(fin_year, paste(baseDir,"/processed_files/",i,"_combined.csv",sep=""))
        }
        
        single_years <- c(2003,2005)  # there is only one workbook for these fiscal years
        for( i in 1:length(single_years)){
                iniFile <- paste(baseDir, "/workbooks", "/initial_extracts/", single_years[i], "_initial_extract.xls",sep="")
                ini_year <- read.xlsx(iniFile, sheetName=paste(single_years[i],"_data", sep=""),      
                                      #startRow=1, endRow=1160, 
                                      as.data.frame=TRUE, 
                                      header=TRUE)
                ini_year$wbhistory <- "initial"; ini_year$wbhistory <- as.character(ini_year$wbhistory) # record history of row
                ini_year$code <- as.character(ini_year$code)
                if(TRUE %in% is.na(ini_year$code)){
                        ini_year <- ini_year[-which(is.na(ini_year$code)), ]           # need to remove any blank lines from dataframe here
                }
                #-- insert code here to deal with duplicated client codes:
                if (single_years[i] == 2005){
                        fixThisCode <- c("A212", "A320" ,"A405")
                        for(q in 1:length(fixThisCode)){
                                # the second occurrance is the row to delete
                                ini_year <- ini_year[-max(which(ini_year$code == fixThisCode[q])), ]
                        }
                }
                write.csv(ini_year, paste(baseDir,"/processed_files/",single_years[i],"_combined.csv",sep=""))
        }
        
        i <- 2004       # need to combine the inital workbook with the data in assessment_final_2004.csv, which has no history
        iniFile <- paste(baseDir, "/workbooks", "/initial_extracts/", i, "_initial_extract.xls",sep="")
        ini_year <- read.xlsx(iniFile, sheetName=paste(i,"_data", sep=""),      
                              #startRow=1, endRow=1160, 
                              as.data.frame=TRUE, 
                              header=TRUE)
        ini_year$wbhistory <- "initial"; ini_year$wbhistory <- as.character(ini_year$wbhistory) # record history of row
        ini_year$code <- as.character(ini_year$code)
        fin_year <- read.csv(paste(baseDir, "/processed_files/", "assessment_final_2004.csv",sep=""))
        fin_year$code <- as.character(fin_year$code)
        if(TRUE %in% is.na(fin_year$code)){
                fin_year <- fin_year[-which(is.na(fin_year$code)), ]   # need to remove any blank lines from dataframe here, if any
        }
        if(TRUE %in% is.na(ini_year$code)){
                ini_year <- ini_year[-which(is.na(ini_year$code)), ]   # need to remove any blank lines from dataframe here, if any
        }
        for(j in 1:nrow(ini_year)){
                if(fin_year$code[j] %in% ini_year$code){
                        ini_year$wbhistory[ini_year$code == fin_year$code[j]] <- "updated"
                        ini_year$assessmentdollarsfy2004[ini_year$code  == fin_year$code[j]] <- fin_year$assessmentdollars2004[j]      
                }
        }
        onlyInFinal <- fin_year[!fin_year$code %in% ini_year$code,]
        if(TRUE %in% is.na(onlyInFinal$code)){
                onlyInFinal <- onlyInFinal[-which(is.na(onlyInFinal$code)), ]   # need to remove any blank lines from dataframe here
        }
        if(nrow(onlyInFinal) > 0){
                # need to copy info from fin_year to ini_year  to dataframe
                onlyInFinal$wbhistory <- "final"; onlyInFinal$wbhistory <- as.character(onlyInFinal$wbhistory)
                ini_year <- rbind(ini_year,onlyInFinal)
                ini_year <- arrange(ini_year,code)
        }
        write.csv(ini_year, paste(baseDir,"/processed_files/",i,"_combined.csv",sep=""))
}
# ----------------------------------------------------- Assemble raw tidy dataset from combined workbooks:
assemble_raw_tidy_dataset <- function(){
        data_2003 <-  read.csv(paste(baseDir, "/processed_files/","2003","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2004 <-  read.csv(paste(baseDir, "/processed_files/","2004","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2005 <-  read.csv(paste(baseDir, "/processed_files/","2005","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2006 <-  read.csv(paste(baseDir, "/processed_files/","2006","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2007 <-  read.csv(paste(baseDir, "/processed_files/","2007","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2008 <-  read.csv(paste(baseDir, "/processed_files/","2008","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2009 <-  read.csv(paste(baseDir, "/processed_files/","2009","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2010 <-  read.csv(paste(baseDir, "/processed_files/","2010","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2011 <-  read.csv(paste(baseDir, "/processed_files/","2011","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2012 <-  read.csv(paste(baseDir, "/processed_files/","2012","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2013 <-  read.csv(paste(baseDir, "/processed_files/","2013","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2014 <-  read.csv(paste(baseDir, "/processed_files/","2014","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_2015 <-  read.csv(paste(baseDir, "/processed_files/","2015","_combined.csv",sep=""),stringsAsFactors = FALSE)
        data_list <- list(data_2003,data_2004,data_2005,data_2006,data_2007,data_2008,data_2009,data_2010,data_2011,data_2012,data_2013,data_2014,data_2015)
        for(i in 1:(length(data_list)-2)){   
                tidy_data_one_fiscalyear <- data_list[[i]]
                drops <- c(1,5:16)       
                tidy_data_one_fiscalyear <- tidy_data_one_fiscalyear[,-drops]  # discard variables that are not needed
                tidy_data_one_fiscalyear$fiscalyear <- NA ; tidy_data_one_fiscalyear$fiscalyear <- as.numeric( tidy_data_one_fiscalyear$fiscalyear)
                colnames(tidy_data_one_fiscalyear) <- c("code", "clientname", "assessmentdollars", "capnumclaims", "capclaimscost","wbhistory","fiscalyear")
                
                # add new historical variables to dataframe; use NA as initial values to detect missing data later:
                whichYear <- 2002 + i                                          # this would need to be changed if the first fiscal year was not 2003
                tidy_data_one_fiscalyear$fiscalyear <- whichYear
                tidy_data_one_fiscalyear$fte <- NA;tidy_data_one_fiscalyear$fte <- as.numeric(tidy_data_one_fiscalyear$fte)
                tidy_data_one_fiscalyear$payrolldollars <- NA;tidy_data_one_fiscalyear$payrolldollars <- as.numeric(tidy_data_one_fiscalyear$payrolldollars)
                tidy_data_one_fiscalyear$numberclaims <- NA;tidy_data_one_fiscalyear$numberclaims <- as.numeric(tidy_data_one_fiscalyear$numberclaims)
                tidy_data_one_fiscalyear$costdollars <- NA;tidy_data_one_fiscalyear$costdollars <- as.numeric(tidy_data_one_fiscalyear$costdollars)
                
                # create a dataframe with the appropriate historical values of fte, payrolldollars, numberclaims and costdollars from a future year + 2 workbook:
                tempHistory <- data_list[[i+2]]
                tempHistory <- tempHistory[,-1]        # discard unneeded variables
                tempHistory <- tempHistory[,-(16:18)]  #   "        "        "
                tempHistory <- tempHistory[,-(13:14)]  #   "        "        "
                tempHistory <- tempHistory[,-(10:11)]  #   "        "        "
                tempHistory <- tempHistory[,-(7:8)]    #   "        "        "
                tempHistory <- tempHistory[,-(3:5)]    #   "        "        "
                colnames(tempHistory) <- c("code","clientname", "payrolldollars", "fte", "numberclaims", "costdollars") # rename variables for easy stacking of rows & easy coding
                
                # copy historical values for payroll, number of claims, claims costs and fte to the dataframe containing the assessment info
                for(t in 1:nrow(tidy_data_one_fiscalyear)){
                        if(tidy_data_one_fiscalyear$code[t] %in% tempHistory$code){
                                tidy_data_one_fiscalyear$payrolldollars[t] <- tempHistory$payrolldollars[tempHistory$code == tidy_data_one_fiscalyear$code[t]]
                                tidy_data_one_fiscalyear$numberclaims[t] <- tempHistory$numberclaims[tempHistory$code == tidy_data_one_fiscalyear$code[t]]
                                tidy_data_one_fiscalyear$costdollars[t] <- tempHistory$costdollars[tempHistory$code == tidy_data_one_fiscalyear$code[t]]
                                tidy_data_one_fiscalyear$fte[t] <- tempHistory$fte[tempHistory$code == tidy_data_one_fiscalyear$code[t]]
                        }
                }
                if(i == 1){
                        tidyRawSORM <- tidy_data_one_fiscalyear
                }else{
                        tidyRawSORM <- rbind(tidyRawSORM,tidy_data_one_fiscalyear)
                }  
        }
        
        #create the calculated variables of code.year, costperfte and ifr:
        tidyRawSORM$code.year <- paste(tidyRawSORM$code,".",tidyRawSORM$fiscalyear,sep="")
        tidyRawSORM$costperfte <- tidyRawSORM$costdollars / tidyRawSORM$fte
        tidyRawSORM$ifr <- (tidyRawSORM$numberclaims / tidyRawSORM$fte) * 100   
        tidyRawSORM <-  tidyRawSORM[,c(1,7,2,8,9,11,10,14,13,3,4,5,6,12)]       # reorder the columns for better viewing
        
        #save data to disk as an R dataframe object keeping variable classes intact:
        if(sum(duplicated(tidyRawSORM$code.year)) == 0){
                save(tidyRawSORM,file = paste(baseDir, "tidyRawSORM.Rdata", sep="/"))
                
        }else{
                print("WARNING: There are duplicates in the tidyRawSORM$code.year dataframe that need to be fixed. No file written.")
        }
}
# ----------------------------------------------------- Perform client replacements and mergers:
replace_client <- function(fromCode,toCode,newName){
        if(!exists("tidyRawSORM")){
                load(tidyRawSORM,file = paste(baseDir, "tidyRawSORM.Rdata", sep="/"))
        }else{
                has_old_data <- filter(tidyRawSORM, code %in% fromCode)
                # update code and clientname values in tidyRawSORM based on code.year 
                tidyRawSORM$code[tidyRawSORM$code.year %in% has_old_data$code.year] <<- toCode
                tidyRawSORM$clientname[tidyRawSORM$code.year %in% has_old_data$code.year] <<- newName
                needsComplexUpdate <- has_old_data[!complete.cases(has_old_data),]                   # which rows need updated historical values for some variables
                for(v in 1:nrow(needsComplexUpdate)){
                        futureYear <- needsComplexUpdate$fiscalyear[v] + 2
                        aFile <- paste(futureYear,"_combined.csv",sep="")
                        futureData <- read.csv(paste(baseDir,"processed_files",aFile,sep="/"),stringsAsFactors = FALSE)
                        futureData <- futureData[,-c(1,4:6,8:9,11:12,14:15,17:19)]  # remove unneeded variables
                        colnames(futureData) <- c("code","clientname", "payrolldollars", "fte", "numberclaims", "costdollars") # rename variables for easier coding
                        whichFutureData <- filter(futureData, code == toCode)
                        tidyRawSORM$payrolldollars[tidyRawSORM$code.year == needsComplexUpdate$code.year[v]] <<- whichFutureData$payrolldollars
                        tidyRawSORM$fte[tidyRawSORM$code.year == needsComplexUpdate$code.year[v]] <<- whichFutureData$fte
                        tidyRawSORM$numberclaims[tidyRawSORM$code.year == needsComplexUpdate$code.year[v]] <<- whichFutureData$numberclaims
                        tidyRawSORM$costdollars[tidyRawSORM$code.year == needsComplexUpdate$code.year[v]] <<- whichFutureData$costdollars
                        tidyRawSORM$ifr[tidyRawSORM$code.year == needsComplexUpdate$code.year[v]] <<- whichFutureData$numberclaims / whichFutureData$fte
                        tidyRawSORM$costperfte[tidyRawSORM$code.year == needsComplexUpdate$code.year[v]] <<- whichFutureData$costdollars / whichFutureData$fte
                }
        }
        tidyRawSORM$code.year <<- paste(tidyRawSORM$code,".",tidyRawSORM$fiscalyear,sep="")  # update code.year variable to indicate correct values, discarding out-of-date values
        return("OK,done.")
}
merge_and_replace_juve_clients <- function(fromList,toCode,newName){
        mergeTheseObvs <- filter(tidyRawSORM,code %in% fromList)     # filter(tidyRawSORM,code %in% fromList)
        whichAreSimple <- complete.cases(mergeTheseObvs)
        mergeSimple <- mergeTheseObvs[whichAreSimple,]
        yearList <- unique(mergeSimple$fiscalyear)
        for(w in 1:length(yearList) ){
                single_year_obvs <- filter(mergeSimple, fiscalyear == yearList[w])
                if(nrow(single_year_obvs) == 2){
                        tidyRawSORM$fte[tidyRawSORM$code.year == single_year_obvs$code.year[1]] <<- sum(single_year_obvs$fte)
                        tidyRawSORM$payrolldollars[tidyRawSORM$code.year == single_year_obvs$code.year[1]] <<- sum(single_year_obvs$payrolldollars)
                        tidyRawSORM$numberclaims[tidyRawSORM$code.year == single_year_obvs$code.year[1]] <<- sum(single_year_obvs$numberclaims)
                        tidyRawSORM$costdollars[tidyRawSORM$code.year == single_year_obvs$code.year[1]] <<- sum(single_year_obvs$costdollars)
                        tidyRawSORM$assessmentdollars[tidyRawSORM$code.year == single_year_obvs$code.year[1]] <<- sum(single_year_obvs$assessmentdollars)
                        tidyRawSORM$ifr[tidyRawSORM$code.year == single_year_obvs$code.year[1]] <<- (tidyRawSORM$numberclaims[tidyRawSORM$code.year == single_year_obvs$code.year[1]] / tidyRawSORM$fte[tidyRawSORM$code.year == single_year_obvs$code.year[1]] ) *100
                        tidyRawSORM$costperfte[tidyRawSORM$code.year == single_year_obvs$code.year[1]] <<- tidyRawSORM$costdollars[tidyRawSORM$code.year == single_year_obvs$code.year[1]] / tidyRawSORM$fte[tidyRawSORM$code.year == single_year_obvs$code.year[1]] 
                        tidyRawSORM$code[tidyRawSORM$code.year == single_year_obvs$code.year[1]] <<- toCode
                        tidyRawSORM$clientname[tidyRawSORM$code.year == single_year_obvs$code.year[1]] <<- newName
                        tidyRawSORM <<- tidyRawSORM[-which(tidyRawSORM$code.year == single_year_obvs$code.year[2]),]
                        tidyRawSORM$code.year <<- paste(tidyRawSORM$code,".",tidyRawSORM$fiscalyear,sep="")          # overwrites old code.year entry
                }
        }
        if(nrow(mergeTheseObvs) > nrow(mergeSimple)){
                # there are rows with missing values that need to be updated
                theComplexObvs <- mergeTheseObvs[!whichAreSimple,]
                theComplexObvs <- theComplexObvs[!theComplexObvs$fiscalyear == 2012,]   # there is already an observation for Texas Juvenile Justice Dept for fiscal 2012
                yearList <- unique(theComplexObvs$fiscalyear)
                for(w in 1:length(yearList)){
                        if(nrow(theComplexObvs) == 2){
                                futureYear <- yearList[w] + 2
                                aFile <- paste(futureYear,"_combined.csv",sep="")
                                futureData <- read.csv(paste(baseDir,"processed_files",aFile,sep="/"),stringsAsFactors = FALSE)
                                futureData <- futureData[,-c(1,4:6,8:9,11:12,14:15,17:19)]  # remove unneeded variables
                                colnames(futureData) <- c("code","clientname", "payrolldollars", "fte", "numberclaims", "costdollars") # rename variables for easier coding
                                whichFutureData <- filter(futureData, code == toCode)
                        }
                        tidyRawSORM$fte[tidyRawSORM$code.year == theComplexObvs$code.year[1]] <<- whichFutureData$fte[1]
                        tidyRawSORM$payrolldollars[tidyRawSORM$code.year == theComplexObvs$code.year[1]] <<- whichFutureData$payrolldollars[1]
                        tidyRawSORM$costdollars[tidyRawSORM$code.year == theComplexObvs$code.year[1]] <<- whichFutureData$costdollars[1]
                        tidyRawSORM$numberclaims[tidyRawSORM$code.year == theComplexObvs$code.year[1]] <<- whichFutureData$numberclaims[1]
                        tidyRawSORM$assessmentdollars[tidyRawSORM$code.year == theComplexObvs$code.year[1]] <<- sum(theComplexObvs$assessmentdollars)
                        tidyRawSORM$clientname[tidyRawSORM$code.year == theComplexObvs$code.year[1]] <<- whichFutureData$clientname[1]
                        tidyRawSORM$code[tidyRawSORM$code.year == theComplexObvs$code.year[1]] <<- whichFutureData$code[1]
                        tidyRawSORM$ifr[tidyRawSORM$code.year == theComplexObvs$code.year[1]] <<- (tidyRawSORM$numberclaims[tidyRawSORM$code.year == theComplexObvs$code.year[1]] / tidyRawSORM$fte[tidyRawSORM$code.year == theComplexObvs$code.year[1]]) * 100
                        tidyRawSORM$costperfte[tidyRawSORM$code.year == theComplexObvs$code.year[1]] <<- tidyRawSORM$costdollars[tidyRawSORM$code.year == theComplexObvs$code.year[1]] / tidyRawSORM$fte[tidyRawSORM$code.year == theComplexObvs$code.year[1]]
                        tidyRawSORM <<- tidyRawSORM[-which(tidyRawSORM$code.year == theComplexObvs$code.year[2]),]   # delete unneed observation 2011
                        tidyRawSORM <<- tidyRawSORM[-which(tidyRawSORM$code.year %in% c("A665.2012","A694.2012" )),] # delete unneed observatios 2012
                        tidyRawSORM$code.year <<- paste(tidyRawSORM$code,".",tidyRawSORM$fiscalyear,sep="")          # overwrites old code.year entry
                }
        }
        return("Ok, merge complete.")
}
