  An Analysis of Eleven Years of Summary Workers’ Compensation Data for the
Texas State Office of Risk Management’s Risk Pool and Risk Pool Finance Method
                                by
                           James L. Sims

                           "First Look"
                                                
When the Texas State Office of Risk Management (SORM) introduced a new risk 
pool finance method in fiscal 2001-2002 to assess SORM clients for the 
anticipated costs of workers' compensation and the delivery of certain other 
risk management services, SORM made certain claims about the benefits of the 
new method:

(1) clients would be held more accountable for their claims;
(2) clients would have incentives to lower their injury frequency rates;
(3) clients would have a smoother process when budgeting for workers' 
    compensation rather than paying their claims directly;
(4) small clients would be protected from catastrophic losses
  
This project was undertaken to determine if SORM clients actually experienced 
the benefits SORM claimed by analyzing eleven fiscal years (2003-2013) of client
data.  During fiscal 2003-2013, the weighting of the risk pool finance 
assessment factors did not change.

A tidySORM dataset was constructed from a collection of annual assessment 
Excel workbooks posted on the SORM web site.  The code for downloading the 
assessment workbooks, extracting relevant historical data from the workbooks 
and the assembly of a tidy data set can be found in the 
create_tidy_data.R script file, which calls the source function on the file 
create_tidy_data_functions.R.

The results of running the code in create_tidy_data.R on June 16, 2015 can be 
found by loading the dataframe stored in tidySORM.Rdata in the project repo.

The GitHub repo for this project also contains a file called 
2003_initial_extract.xls in the workbooks/initial_extracts/ subdirectory.
This file is needed when running a few lines of code in the 
analyze_tidySORM.R script file related to the distribution of claims costs 
for fiscal 2001, the year before the initial implementation of the new risk pool
finance method.  It is provided in the repo for those who do not want to go 
through the process of downloading assessment workbooks and assembling the tidy 
data set on their own.

The comments in analyze_tidySORM.R are intended to guide the reader of a report 
based on this project as they match up the calculations with the text, charts 
and tables in the report.  The report was sent to the Executive Director of 
SORM in June 2015, and copied to risk managers for the SORM clients that 
overpaid the most for their claims, those clients being those listed in Table 4 
of the report.  The report is not posted in the project repo on GitHub.

The code for this project is provided to others by the author under a 
Creative Commons Attribution License 4.0. Authors of required R packages needed 
for this project retain their individual copyrights as stated in their package 
documentation.


