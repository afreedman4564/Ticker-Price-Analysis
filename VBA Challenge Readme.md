# Story board...

## number of records in file
Trying to identify the number of unique brands of credit card companies
Go through the various rows, starting in row 2 and column 1
Iterate from i = 2 to n, where n is the last row. Use XUP to identify the last row and then assign as n

## Close Price and ticker symbol id
Now that n has been defined, run an iterative loop and pick up unique brand values looking for when the value is no longer the same and at the same time pick up the close price since the data is sorted by ticker and trade date

## open price on first trading day of year
Create a counter at the ticker level. Once the values are produced identify the value for the first. Should be 1. Use this to grab the opening price.

## aggregate the trade volume
leverage the if statements created to grab the open price and close price, and possibly create additional as needed to create a counter that serves as an aggregator

## calculate the percent yearly change and # change, comparing open and close price from beginning of the year and end of year respectively
add calculations and formatting to cells displaying the metrics sought

## bonus
create a loop to run through the consolidated table at the ticker level for each sheet to id the ticker(s) with the greatest percent increase and decrease, along with the ticker that has the greatest volume increase

once calculated for each sheet and all other loops and conditionals are complete, build a loop that goes across all worksheets and identify the tickers with the greatest percent increase and decrease, along with the ticker that has the greatest volume increase regardless of worksheet

code has to run across all worksheets

## considerations
vet the data to make sure aggregators and counters working properly, as well as only one year in each worksheet
account for nulls, missing and zero values...
use approaches taught in class with google-fu if needed to 'creatively' overcome obstacle(s)
use test data file shared so don't mess up larger data set
make sure script runs same/yields same output on each worksheet


### completeion checklist
create readme 
from class readme with instructions:
Create a script that will loop through all the stocks for one year and output the following information:

  * The ticker symbol. 

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year. 

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year. 

  * The total stock volume of the stock. 

* You should also have conditional formatting that will highlight positive change in green and negative change in red. 

Bonus, also from class readme with instructions:
* Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". 

* Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once. 

### submission checklist
1. Create a new repository for this project called `VBA-challenge`. **Do not add this homework to an existing repository**. 

2. Inside the new repository that you just created, add any VBA files you use for this assignment. These will be the main scripts to run for each analysis.

* To submit please upload the following to GitHub:

  * A screen shot for each year of your results on the Multi Year Stock Data.

  * VBA scripts as separate files.

* Ensure you commit regularly to your repository and it contains a README.md file.

* Upload all of your files to your GitHub repository which should also contain a README.md file.

