# VBA_challenge


download the file from module VBA_challenge

we have the data of stock market 
we need to calculte 
The ticker symbol

Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.


first we create the loop all over the worksheet 
will go through all worksheets(Q1,Q2,Q3,Q4)

To go through each worksheet we do for each ws in activeworkbook.worksheets


then we activate the ws
later we got the last row


afterthat we put output headers

ticker
Quaterly change
Percent change
Total stock volume


then we got the open price that is in row number 3

we loop through all the ticker from row 2 to last row because our values start from row 2 
if values are not equal to original value then name of ticker will come


we set the the ticker name in column 9

later we will look into close price that is in column 6

we calculate the quaterly change value by subtratcing the open price at the beginning of quarter  from close price at end of the quarter
close price -open price=quaterly change 

percent change= quaterly cahnge/open price
I did the percent change in Number format="0.00"%

intial we took the value of volume 0
calculate the total volume by  volume+ value of volume in 7 column
to print to next row went through loop 
cells volume in 12 row 

row= row+1

endif 
nexti


calculate the last row of ticker column



later on set the colors

j=2 to quaterly change last row
put conditional statement if value is more than 0 then the color the cell green otherwise color it red


after that do endifand next j

calculate the ticker value greatest %age
increase the %age
decrease the %age
total volume '

To find the highest value of ticker
maximum value and minimum value

mext x 
next ws

