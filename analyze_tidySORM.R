setwd("~/git/sorm")                                                    # change this directory setting to suit your preference
                                                                       # for a base directory; should be same as this file
baseDir <- getwd()                                                     # many functions depend on base directory variable 

library(dplyr)                                                         # load needed library
options(dplyr.print_max = 50)                                          # set dplyr options to display up to fifty rows of a local data frame
source(paste(baseDir,"analyze_tidySORM_functions.R",sep="/"))                   # load needed functions from R script file
load_tidy_data()                                                       # load tidySORM dataset from .Rdata file


#----------------Examples of calling a few functions, which are part of this project, for the interactive exploration of the tidySORM dataset:

chartAE("A737")                                                        # example of plotting the cummulative over and under assessment for a single client

chartIFR("A737")                                                       # example of plotting IFR versus fiscal year for a single client

getCode("Angelo")                                                      # example of looking up the code for Angelo State Univeristy by searching the tidySORM dataset by a partial client name

getData("A737",2003,2005)                                              # example of displaying data in the console for a single client by specifiying the client code and fiscal years to display

getName("A737")                                                        # example of looking up the name of a client in the tidySORM dataset by specifying the full code

getTopCost(10,2003,2013)                                               # example of finding the top ten clients with the highest claims cost for a range of fiscal years (inclusive)
                                                                       # sorted from high to low; see topCostClients dataframe

getTopIFR(5,7.6,25,2003,2013)                                          # example of finding the top five clients with mean injury frequency rates (IFR) in the range 7.6 to 25 inclusive,
                                                                       # for given range of fiscal years, sorted high to low; see topIFRClients dataframe




#-------------                                An Analysis of Eleven Years of Summary Workers’ Compensation Data for the
#-------------                             Texas State Office of Risk Management’s Risk Pool and Risk Pool Finance Method
#-------------                                                                      by
#-------------                                                                James L. Sims

#-------------                                                                "First Look"

#------------------------------------------ Cacluations Supporting Introduction:

#--- Figure 1. Histogram of claims cost distribution in 2001 for clients in 2003:

getFiscal2001_costDistribution()                                                         # display a histogram of claims cost distribution for fiscal year 2001 for clients that were still clients in fiscal 2003
                                                                                         # REQUIRES workbooks/initial_extracts/2003_initial_extract.xls"; returns a required R vector for Figure 1. legend calculations
#--- Figure 1. legend:
        length(the2001CostDollars)                                                       # number of clients = N
        numnocost <- length(the2001CostDollars[the2001CostDollars == 0])
        numnocost                                                                        # number of clients with no claims cost fiscal 2001
        (numnocost/  length(the2001CostDollars)) * 100                                   # percent of clients with no claims cost fiscal 2001
        length(the2001CostDollars[the2001CostDollars > 0 & the2001CostDollars < 1])      # number of clients with cost less than 1 million U.S. dollars excluding those with zero costs for fiscal 2001
        length(the2001CostDollars[the2001CostDollars > 1 & the2001CostDollars < 10])     # number of clients with more than 1 million US dollars but less than 10 million in claims cost for fiscal 2001
        length(the2001CostDollars[the2001CostDollars > 10])                              # number of clients with more than 10 million US dollars in claims cost


#------------------------------------------ Cacluations Supporting Results:

#--- Claim #1 Assessment will hold clients more accountable for their claims:

nrow(filter(tidySORM, numberclaims == 0, costdollars == 0, fiscalyear > 2002, fiscalyear < 2014))    # number of entries in tidySORM with no claims and no claims cost 2003-2013
(nrow(filter(tidySORM, numberclaims == 0, costdollars == 0, fiscalyear > 2002, 
             fiscalyear < 2014)) / nrow(tidySORM))*100                                               # percent of entries in tidySORM with no claims and no claims costs
filter(tidySORM, numberclaims == 0, costdollars == 0, fiscalyear > 2002, 
       fiscalyear < 2014) %>%
        summarise(sum(assessmentdollars))                                                            # total dollars collected in assessments from clients with no claims and no claims cost in a given fiscal year
zeroTemp <- filter(tidySORM, numberclaims == 0, costdollars == 0, 
                   fiscalyear > 2002, fiscalyear < 2014)
range(zeroTemp$assessmentdollars)                                                                    # range of assessment for clients with no claims history
zeroTemp[zeroTemp$assessmentdollars == max(zeroTemp$assessmentdollars),]                             # which client(s) with a perfect history assessed the most dollars
zeroTemp[zeroTemp$assessmentdollars == min(zeroTemp$assessmentdollars),]                             # which client(s) with a perfect history assessed the least dollars
thesum <- (filter(tidySORM, numberclaims == 0, costdollars == 0, 
                  fiscalyear > 2002, fiscalyear < 2014) %>%
                   summarise(sum(assessmentdollars))  )
thesum                                                                                               # total assessment of clients with perfect records
round(thesum/ nrow(filter(tidySORM, numberclaims == 0, costdollars == 0, 
                          fiscalyear > 2002, fiscalyear < 2014)),digits=0)                           # rounded average assessment for clients with perfect records


#---  Claim #2 Assessment will incentivize clients will to lower injury frequency rates:

#---- Table 1:
getTopCost(20,2003,2013)                                                                            # top 20 clients by total claims cost 2003-2013 sorted high to low
top20costclients <- topCostClients
sum(top20costclients$totalclaimscost[1:5])                                                          # total claims cost top 5
sum(top20costclients$totalclaimscost[6:10])                                                         # total claims cost top ranked 6 through 10
sum(top20costclients$totalclaimscost[11:15])                                                        # total claims cost top ranked 11 through 15
sum(top20costclients$totalclaimscost[16:20])                                                        # total claims cost top ranked 16 through 20
# write.csv(top20costclients,paste(baseDir,"/processed_files/","top20_bycost.csv",sep=""))

#---- Chart Group 1:
for(r in 1:5){
        chartIFR(top20costclients$code[r])                                                         # plots IFR data, one plot per client; diplayed in Viewer tab of RStudio
}
theDF <- top20costclients[1:5,]
compareifrs(theDF)                                                                                 # percent change un-weighted IFR mean for first 3 years compared to last 3 years Chart Group 1                                                                                           

# --- Chart Group 2:
for(r in 6:10){
        chartIFR(top20costclients$code[r])
}
theDF <- top20costclients[6:10,]
compareifrs(theDF)                                                                                # percent change un-weighted IFR mean for first 3 years compared to last 3 years Chart Group 2                                                                        

# --- Chart Group 3:
for(r in 11:15){
        chartIFR(top20costclients$code[r])
}
theDF <- top20costclients[11:15,]
compareifrs(theDF)                                                                               # percent change un-weighted IFR mean for first 3 years compared to last 3 years Chart Group 3

# --- Chart Group 4:
for(r in 16:20){
        chartIFR(top20costclients$code[r])
}
theDF <- top20costclients[16:20,]
compareifrs(theDF)                                                                               # percent change un-weighted IFR mean for first 3 years compared to last 3 years Chart Group 4
       

# --- Calculations to generalize IFR incentive findings to the entire risk pool:

 ifrHighLow <- filter(tidySORM, fiscalyear > 2002, fiscalyear < 2014) %>%                     
        group_by(code) %>%
        summarise(max(ifr),min(ifr),sum(costdollars),
                  sum(assessmentdollars),min(fiscalyear),max(fiscalyear))  
nrow(ifrHighLow)                                                                                                    # number of unique clients by code

ifrHighLow <- ifrHighLow[complete.cases(ifrHighLow), ]                                                              # remove clients with missing data 
nrow(ifrHighLow)                                                                                                    # number of unique clients by code with complete data
colnames(ifrHighLow) <- c("code", "maxifr", "minifr",
                          "totalclaimsclost","totalassessment","firstyear","lastyear")
ifrHighLow <- ifrHighLow[-which(ifrHighLow$maxifr == 0 | ifrHighLow$minifr == 0), ]                                 # remove all the entries that have zero ifrs; 
nrow(ifrHighLow)                                                                                                    # number of clients that have claims every year of the study
sum(ifrHighLow$maxifr >= 3.5 & ifrHighLow$maxifr < 7.5 & ifrHighLow$minifr < 3.5)                                   # number of clients crossing only the 3.5 IFR incentive threshold
(sum(ifrHighLow$maxifr >= 3.5 & ifrHighLow$maxifr < 7.5 &
             ifrHighLow$minifr < 3.5) /nrow(ifrHighLow) ) *100                                                      # percent of clients with claims every fiscal year crossing 3.5 IFR
whichEntries <- (ifrHighLow$maxifr >= 3.5 & ifrHighLow$maxifr < 7.5 & ifrHighLow$minifr < 3.5) 
ifrHighLow[whichEntries,] ; thecodes <- ifrHighLow[whichEntries,]                                                   # identities of clients crossing only the 3.5 IFR threshold

for(g in 1:nrow(thecodes)){
        chartIFR(thecodes$code[g])                                                                                  # visualize how these clients started and ended the study by IFR
}

sum(ifrHighLow$maxifr >= 7.5 & ifrHighLow$minifr >= 3.5 & ifrHighLow$minifr < 7.5)                                  # number of clients crossing only the 7.5 IFR incentive threshold; 3 observations
(sum(ifrHighLow$maxifr >= 7.5 & ifrHighLow$minifr >= 3.5 &
             ifrHighLow$minifr < 7.5) / nrow(ifrHighLow) ) *100                                                     # percent of clients crossing only 7.5 IFR incentive threshold
whichEntries1 <- (ifrHighLow$maxifr >= 7.5 & ifrHighLow$minifr >= 3.5 & ifrHighLow$minifr < 7.5)
ifrHighLow[whichEntries1,]  ; someclients <-  ifrHighLow[whichEntries1,]                                            # identities of clients crossing upper IFR threshold only


# --- Chart Group 6: client crossing only the 7.5 IFR threshold:
for (g in 1:nrow(someclients)){
        chartIFR(someclients$code[g])                                                                               # plot IFR vs Fiscal Year from this group of clients
}
sum(someclients$totalclaimsclost)                                                                                   # total cost of claims for this group all fiscal years          

sum((ifrHighLow$maxifr >= 7.5 & ifrHighLow$minifr <= 3.5))                                                          # number of clients crossing both IFR incentive thresholds
(sum((ifrHighLow$maxifr >= 7.5 & ifrHighLow$minifr <= 3.5)) /nrow(ifrHighLow) ) * 100                               # percent of clients crossing both IFR incentive thresholds
whichEntries2 <- (ifrHighLow$maxifr >= 7.5 & ifrHighLow$minifr <= 3.5)
ifrHighLow[whichEntries2,]; someclients2 <-  ifrHighLow[whichEntries2,]                                             # identities of clients crossing two IFR thresholds

# --- Chart Group 5: clients that crossed both IFR incentive thresholds
for (g in 1:nrow(someclients2)){
        chartIFR(someclients2$code[g])                                                                              # plot IFR vs Fiscal Year from this group of clients
}

# --- Explain Chart Group 5 for clients crossing both IFR thresholds
filter(tidySORM,code == "A809") %>%
        summarize(max(fte), min(numberclaims), max(numberclaims))                                                   # maximum FTE, max and min number of claims

filter(tidySORM,code == "C161") %>%
        summarize(max(fte), min(numberclaims), max(numberclaims)) 

filter(tidySORM,code == "C178") %>%
        summarize(max(fte), min(numberclaims), max(numberclaims))  

atable <- filter(tidySORM, code %in% ifrHighLow$code[whichEntries2]) %>%
        group_by(code) %>%
        summarise(sum(costdollars))
sum(atable[,2])                                                                                                    # total claims cost these three clients



#---  Claim #3 Assessment will provide a smoother ride compared to paying own claims every year:

# --- Calculations for Table 2:
top20smootherRide <- top20costclients[,1:2]
top20smootherRide$sdcost <- NA; top20smootherRide$sdcost <- as.numeric(top20smootherRide$sdcost)
top20smootherRide$sdassessment  <- NA
top20smootherRide$sdassessment <- as.numeric(top20smootherRide$sdassessment)
top20smootherRide$overpaid  <- NA
for(v in 1:nrow(top20smootherRide)){
        tempclient <- filter(tidySORM,code == top20smootherRide$code[v], fiscalyear > 2002, fiscalyear < 2014)  
        top20smootherRide$sdcost[v] <- round(sd(tempclient$costdollars), digits = 0)
        top20smootherRide$sdassessment[v] <- round(sd(tempclient$assessmentdollars), digits = 0)
        top20smootherRide$overpaid[v] <- sum(tempclient$assessmentdollars) > sum(tempclient$costdollars) 
}
top20smootherRide$smoother <- top20smootherRide$sdassessment < top20smootherRide$sdcost
top20smootherRide                                                                                               # Table 2 entries
# write.csv(top20smootherRide,paste(baseDir,"/processed_files/","top20_bycost_smoother.csv",sep=""))

# --- Description of Table 2:
sum(top20smootherRide$overpaid); sum(top20smootherRide$overpaid) / nrow(top20smootherRide)                      # how many of top 20 overpaid their claims; fraction of those overpaying
sum(top20smootherRide$smoother); sum(top20smootherRide$smoother) / nrow(top20smootherRide)                      # how many of top 20 had a smoother ride by assessment; fraction of those with smoother ride


# --- Generalize Assessment Effect to entire tidy dataset:
clients <- filter(tidySORM, fiscalyear > 2002, fiscalyear < 2014)  %>%                                            # select observations in the study period
        group_by(code) %>%                                                                                        # and calculate SD for both variables
        summarise(sd(assessmentdollars), sd(costdollars))                                                         
colnames(clients) <- c("code", "sdassessment", "sdclaimscost" )

clients <- filter(clients, sdclaimscost > 0) %>%                                                                  # exclude clients with no claims cost
        mutate(smoother = (sdassessment > sdclaimscost))                                                          # calculate the logical value of the smoother variable
sum(clients$smoother)                                                                                             # number of clients with assessments smoother than paying claims
(sum(clients$smoother) /nrow(clients)) * 100                                                                      # percent of clients having a smoother experience with assessment

# --- Calculations for Table 3:
compareTotalClaimscostAndAssessments <- filter(tidySORM, fiscalyear > 2002, fiscalyear < 2014) %>%                 
        group_by(code) %>%
        summarise(sum(costdollars),sum(assessmentdollars))                                                        # get the total claims cost & total assessment for each client
colnames(compareTotalClaimscostAndAssessments) <- c("code", "totalclaimscost","totalassessmentcost")
nrow(filter(compareTotalClaimscostAndAssessments,totalclaimscost > 0, totalassessmentcost > totalclaimscost))     # number of clients assessed more than their claims cost excluding agencies with no costs                           
nrow(filter(compareTotalClaimscostAndAssessments,totalclaimscost > 0, totalassessmentcost < totalclaimscost))     # number of clients assessed less than their claims
underassessed <- filter(compareTotalClaimscostAndAssessments,totalclaimscost > 0, 
                        totalassessmentcost < totalclaimscost) %>%
        arrange(desc(totalclaimscost)) %>%
        mutate(underassessment = totalclaimscost - totalassessmentcost )  %>%
        mutate(underassessmentpercentage = (underassessment/totalclaimscost)*100)
underassessed$clientname <- tidySORM$clientname[match(underassessed$code, tidySORM$code)]
underassessed <- underassessed[,c(1,6,2,3,4,5)]
underassessed$totalclaimscost <- round(underassessed$totalclaimscost, digits = 0)
underassessed$totalassessmentcost <- round(underassessed$totalassessmentcost, digits = 0)
underassessed$underassessment <- round(underassessed$underassessment, digits = 0)
underassessed$underassessmentpercentage <- round(underassessed$underassessmentpercentage, digits = 2)
underassessed                                                                                                     # Table 3: clients under-assessed, sorted by total claims cost
# write.csv(underassessed,paste(baseDir,"/processed_files/","underassessed_03_12.csv",sep=""))

# --- Legend for Table 3:
sum(underassessed$underassessment)                                                                                # total under-assessment dollars

(sum(underassessed$underassessment[1:3]) / sum(underassessed$underassessment) ) *100                              # percent of under-assessment captured by top 20 clients


# --- Chart Group 7: How Three Subsidized Clients Experienced the Subsidy Year-By-Year:
for(f in 1:3){
        chartAE(underassessed$code[f])
}

# --- Chart Group 8: How Top Five Clients by Claims Cost Experienced the Assessment Effect Year-By-Year:
for(f in 1:5){
        chartAE(top20costclients$code[f])
}


# --- Calculations for Table 4: which clients overpaid the most:
overassessed <- filter(compareTotalClaimscostAndAssessments,totalclaimscost > 0, totalassessmentcost > totalclaimscost) %>%
        arrange(desc(totalassessmentcost)) %>%
        mutate(overassessed = totalassessmentcost - totalclaimscost  )  %>%
        mutate(overassessedpercentage = (overassessed/totalassessmentcost)*100)

overassessed$clientname <- tidySORM$clientname[match(overassessed$code, tidySORM$code)]
overassessed <- overassessed[,c(1,6,2,3,4,5)]
overassessed$totalclaimscost <- round(overassessed$totalclaimscost, digits = 0)
overassessed$totalassessmentcost <- round(overassessed$totalassessmentcost, digits = 0)
overassessed$overassessed <- round(overassessed$overassessed, digits = 0)
overassessed$overassessedpercentage <- round(overassessed$overassessedpercentage, digits = 2)
overassessed <- arrange(overassessed,desc(overassessed))
top20overassessed <- overassessed[1:20,]
top20overassessed                                                                                              # Table 4, sorted by amount of over-assessment compared to claims cost
# write.csv(top20overassessed,paste(baseDir,"/processed_files/","top20overassessed_03_12.csv",sep=""))


# --- Calculations for one-half of the Charts in Chart Group 10, the cummulative assessment effect: 
for(f in 1:5){
        chartAE(top20overassessed$code[f])
}


#---  Claim #4  Small Clients will be Protected from Catestrophic Losses:
round(sum(tidySORM$capnumclaims), digits = 0)                                                                  # number of clients bumping against cap on assessment due to number of claims
round(sum(tidySORM$capclaimscost), digits = 0)                                                                 # number of clients bumping against 


# --- Supplimental Table 1:  clients with missing data:
theCompleteCases <- complete.cases(tidySORM)
obsWithMissingData <- tidySORM[!theCompleteCases,]
obsWithMissingData <- group_by(obsWithMissingData,code) %>%
        summarise(sum(assessmentdollars)) 
colnames(obsWithMissingData) <- c("code","assessment")       
obsWithMissingData <- mutate(obsWithMissingData,clientname = 0)
obsWithMissingData$clientname <-  tidySORM$clientname[match(obsWithMissingData$code, tidySORM$code)]
obsWithMissingData <- obsWithMissingData[,c(1,3,2)]
obsWithMissingData$assessment <- round(obsWithMissingData$assessment, digits = 0)
obsWithMissingData <- arrange(obsWithMissingData,desc(assessment))
obsWithMissingData                                                                                             # Suppliental Table 1
sum(obsWithMissingData$assessment)                                                                             # Dollars of assessments associated with observations with missing data
# write.csv(obsWithMissingData,paste(baseDir,"/processed_files/","obsWithMissingData_03_12.csv",sep=""))