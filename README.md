# VBA_challenge
VBA wall-street homework
# A new repository was created in GitHub
#VBA_codes for analysis.
First step was defining the variables (Ticker, yearly_change, percent_change and Total_volume).
Then for the code to run through each worksheet, ws was defined.
To keep track of the data a summary row was created.
initial values for open_price, summary, and Totalvolume were noted.
For the loop to run through all cells till the last cell containing data, the LastRow was defined.
Since the analysis was required to run through each cell, For loop was initiated.
Also the conditional code was written to calculate the appropriate values for each variable.
All the values were printed in the assigned rows and columns.
Conditional formatting for colourindex was done for the values of Yearly-Change.
Also the number format (0.00%)for Percent_change was done.
Another conditional code was written so as to pick the open_price with '0'.
Summary row was incremented by 1 so that it will not be overwritten.
Same was done by resetting Total_volume and open price inside the for loop.
summary was now reset to get values in appropriate rows in the next worksheet.
