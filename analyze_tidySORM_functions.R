load_tidy_data <- function(){
        if(file.exists(paste(baseDir,"/tidySORM.Rdata",sep=""))){
                load(paste(baseDir,"/tidySORM.Rdata",sep=""),envir = .GlobalEnv)
        }else{
                print("WARNING: expected tidySORM.Rdata but did not find the file.  Set your working directory and the baseDir variable and try again.")
        }
}
getName <- function(which){
        if(which %in% tidySORM$code){
                temp <- tidySORM[grepl(which,tidySORM$code),]
                temp <- temp[,c(1,3)]
                temp[match(unique(temp$code), temp$code),]
        }else{
                print(paste("No data for client code: ",which,". Check your entry and try again.", sep=""))
        }
}
getCode <- function(which){
        temp <- tidySORM[grepl(which,tidySORM$clientname),]
        temp <- temp[,1:3]
        temp$code <- as.character(temp$code)
        temp[match(unique(temp$code), temp$code),]
}
getData <- function(which,begin = 2003, end = 2013){
        tidySORM$code  <- as.character(tidySORM$code)
        filter(tidySORM,code == which,fiscalyear >= begin, fiscalyear <= end)
}
chartIFR <- function(whichcode){
        if(whichcode %in% tidySORM$code){
                temp1 <- filter(tidySORM,code == whichcode)
                if(max(temp1$ifr) < 3.5){
                        plot(temp1$fiscalyear, temp1$ifr, ylim = c(0,4.5), type="b", col="blue", main = paste("IFR by Fiscal Year for ",whichcode,sep=""),xlab="Fiscal Year", ylab="IFR (injuries per 100 FTE per year)")
                        abline(h=3.5,col= "red", lty = 2); abline(h=7.5,col= "red", lty = 2)
                }
                if(max(temp1$ifr) > 3.5 & max(temp1$ifr) < 7.5){
                        plot(temp1$fiscalyear, temp1$ifr, ylim = c(0,8), type="b", col="blue", main = paste("IFR by Fiscal Year for ",whichcode,sep=""),xlab="Fiscal Year", ylab="IFR (injuries per 100 FTE per year)")
                        abline(h=3.5,col= "red", lty = 2); abline(h=7.5,col= "red", lty = 2)
                        
                }
                if(max(temp1$ifr) > 7.5){
                        newmax <- 1 + max(temp1$ifr)
                        plot(temp1$fiscalyear, temp1$ifr, ylim = c(0,newmax), type="b", col="blue", main = paste("IFR by Fiscal Year for ",whichcode,sep=""),xlab="Fiscal Year", ylab="IFR (injuries per 100 FTE per year)")
                        abline(h=3.5,col= "red", lty = 2); abline(h=7.5,col= "red", lty = 2)
                        
                }
        }else{
                print("The code you supplied was not found in the tidySORM dataset.")  
        }
}
chartAE <- function(theCode){
        theTemp <- filter(tidySORM, code == theCode, fiscalyear > 2002 & fiscalyear < 2014)  %>%
                arrange(fiscalyear)
        theTemp <- mutate(theTemp,difference = assessmentdollars - costdollars)
        theTemp <- mutate(theTemp,rolling_dif = 0)
        for(i in 1:nrow(theTemp)){
                if(i==1){
                        theTemp$rolling_dif[i] <-  theTemp$difference[i]  
                }
                if(i > 1){
                        theTemp$rolling_dif[i] <- theTemp$difference[i] + theTemp$rolling_dif[i-1]
                }
        }
        barplot(theTemp$rolling_dif/10^6, main = paste("Assessment Effect for",theCode,sep=" ") ,ylim = c(-18,8), xlab = "Fiscal Year", ylab = "Cummulative Under & Over-assessment ($ millions)",names.arg = theTemp$year,col="darkblue"  )
}
getTopCost <- function(howmany,begin,finish){
        theData <- filter(tidySORM,fiscalyear >= begin, fiscalyear <= finish) %>%
                group_by(code) %>%
                summarise(sum(costdollars)) 
        colnames(theData) <- c("code", "totalclaimscost")          # fix the column names for futher code
        costTotals <- arrange(theData,desc(totalclaimscost))
        if(howmany > 0){
                costTotals <-  costTotals[1:howmany,]
                costTotals$code <- as.character(costTotals$code)
                costTotals$clientname <- NA; costTotals$clientname <- as.character(costTotals$clientname)
                costTotals$clientname <- tidySORM$clientname[match(costTotals$code, tidySORM$code)]
                costTotals <- costTotals[,c(1,3,2)] 
        }
        topCostClients <<- costTotals
        topCostClients
}
getTopIFR <- function(howmany,low,high, begin,finish){
        theData <- filter(tidySORM,fiscalyear >= begin, fiscalyear <= finish) %>%
                group_by(code) %>%
                summarise(mean(ifr)) 
        colnames(theData) <- c("code", "meanIFR")          
        theData <- arrange(theData,desc(meanIFR))
        theData <- theData[theData$meanIFR >= low & theData$meanIFR <= high,]
        if(howmany > 0){
                if(howmany > nrow(theData)){
                        
                }else{
                        theData <-  theData[1:howmany,]
                }
                
        }
        keepers <- complete.cases(theData); theData <- theData[keepers,]
        theData$clientname <- tidySORM$clientname[match(theData$code, tidySORM$code)]; theData <- theData[,c(1,3,2)]
        topIFRClients <<- theData
        topIFRClients
}
getFiscal2001_costDistribution <- function(){
        library(xlsx)
        theInfo <- read.xlsx(paste(baseDir, "/workbooks/initial_extracts/","2003_initial_extract.xls", sep=""),1)
        the2001CostDollars <<- theInfo$costdollarsfy2001per2003/10^6
        hist(the2001CostDollars,xlab="Millions of Dollars per Client", ylab="Number of Clients", main="Claims Cost Distribution Fiscal 2001") 
}
compareifrs <- function(theDF){
        colnames(theDF) <- c("code", "clientname", "ifrpercentchange"); theDF$ifrpercentchange <- NA
        for(s in 1:nrow(theDF)){
                aResult <- tidySORM[tidySORM$code == theDF$code[s],]
                aResult <- arrange(aResult,fiscalyear)
                if(nrow(aResult) >= 6){
                        first3years <- mean(aResult$ifr[1:3])
                        last3years <- mean(aResult$ifr[(nrow(aResult)-3):nrow(aResult)])
                        theDF$ifrpercentchange[s] <- - ((first3years - last3years) / first3years) * 100
                }
        }
        theDF
}