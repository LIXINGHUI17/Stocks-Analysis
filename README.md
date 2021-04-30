# Stocks-Analysis
[VBA_Challenge.xlsm](VBA_Challenge.xlsm)

# Overview of Project
## The purpose and background are well defined.
  ### In this Project, I use the VBA solution code to analyze the Stock Price. For Refactoring I just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. 

# Results
## The analysis is well described with screenshots and code.

### 1. The tickerIndex is set equal to zero before looping over the rows.
#### Step 1a: Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.
![1a](/Resources/1a.png)

### 2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
#### Step 1b: Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
##### The tickerVolumes array should be a Long data type.
##### The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.
![1b](/Resources/1b.png)

### 3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.
#### Step 2a: Create a for loop to initialize the tickerVolumes to zero.
![2a](/Resources/2a.png)
#### Step 2b: Create a for loop that will loop over all the rows in the spreadsheet.
![2b](/Resources/2b.png)

### 4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
#### Step 3a: Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
##### Use the tickerIndex variable as the index.
![3a](/Resources/3a.png)
#### Step 3b: Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable.
![3b](/Resources/3b.png)
#### Step 3c: Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.
![3c](/Resources/3c.png)
#### Step 3d: Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.
![3d](/Resources/3d.png)
#### Step 4: Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.
![4](/Resources/4.png)

### 5. Code for formatting the cells in the spreadsheet is working.
![formatting](/Resources/formatting.png)
### 6. There are comments to explain the purpose of the code.
### 7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module.

#### Dataset Provided
![2017Provided](/Resources/2017Provided.png)
![2018Provided](/Resources/2018Provided.png)

#### VBA Solution
![2017](/Resources/2017.png)
![2018](/Resources/2018.png)

### 8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png.
![VBA_Challenge_2017](/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](/Resources/VBA_Challenge_2018.png)

# Summary
## There is a detailed statement on the advantages and disadvantages of refactoring code in general.
## There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script.
