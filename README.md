# VBA-challenge
VBA Challange
I needed to create a scrit that loops through all the stocks for one year and outputs the ticker symbol, 
yearly change from the opening price at the beginning of a given year to the closing price at the ene of that year,
the percentage change from the opening price at the beginning of a given year to the closing price at the end if that year and the total stock volume.

I first defined the variables then looped through all the worksheets
then find the last row of data in the worksheets
looped through all the rows in th eworksheet
set the closing price
calculated the yearly change
claculated the percent change
added the yearly change
fornatted the percent chsnge as a percentage
incremented the summary row
reset the variables for the next ticker symbol
added stock volume to the total volume
checked if the opening price has been set

I did this in stages so next I added in the positive change ti green and negative to red using conditionak formatting 

Then i needed to add functionality to return the stock with the greatest % increase, greatest % decrease and greatest total volume to my script so it would show up on each spreadsheet
**This is where I couldnt figure out how to calculate it properly so I did two seperate modules with different ways o do it to see if it would work but 
**module 1 would only calculate the Greatest total volume and module 2 would only calculate the greatest % increase and the greatest % decrease. Couldnt figure it out. As you can see on my screenshots
