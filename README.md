# VBA of Wall Street

## Overview of Project

We have a database with two worksheets for the years 2017 and 2018, for this analisys we want to find the sum of yearly volume and yearly return (the percentage difference in price from the beginning of the year to the end of the year) for each green energy stock. We have 12 different stocks (AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP and VSLR) and we need to determine which stocks performed well and which ones did not for each year.

As we need to have the greatest amount of information available to make a decision, it is likely that more stocks and years will be added in the future to the database. With this idea in mind it becomes necessary to refactor the code taking into account taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

## Results

### Stock performance comparison between 2017 and 2018

From the results obtained, we can deduce that 2017 was positive for green energy stocks, 11 of the 12 stocks had a positive return and four of them (DQ, ENPH, FSLR and SEDG) obtained a return above 100%. 

However, during 2018, 10 of the 12 stocks had a negative return, only two (ENPH ​​and RUN) obtained a positive return. Also we can observe that the average of the Total Daily Volume is higher in 2018 than in 2017, which could indicate that 2018 was a negative year for all investments in general even despite having greater yearly volume. 

(imagen)

### Execution times comparison between original script and the refactored script.

**First Code**

In the first code, we ran the first loop through each of 12 tickers, that is, 12 times the second loop was run (`For j = 2 To RowCount`).

    For i = 0 To 11
     ticker = tickers(i)
     totalVolume = 0
    Next i

This last loop is the one that contains the 3 conditionals that generate the greatest load of the code.

    For j = 2 To RowCount

           If Cells(j, 1).Value = ticker Then
               totalVolume = totalVolume + Cells(j, 8).Value
           End If

           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               startingprice = Cells(j, 6).Value
           End If

           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               endingprice = Cells(j, 6).Value
          End If
          
    Next j

Our initial code took 0.69 seconds to run for 2017 and 0.67 seconds for 2018.

(imagen de los tiempos)

In the refactored code we reduce the load using a unique loop with a index variable (tickerIndex) for iterating through all the rows. This unique loop gets the volume and the starting and ending price for each ticker. 

    tickerIndex = 0
      
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
            
    For j = 2 To RowCount   
         
               tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value       
            
           If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
               tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
           End If
                        
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
               tickerEndingPrices(tickerIndex) = Cells(j, 6).Value                 
               tickerIndex = tickerIndex + 1
            End If
   
    Next j
    
After we refactored the code the run time was reduced to only 0.086 seconds for 2017 and 0.088 seconds for 2018, generating a much faster code.

(imagen de los tiempos)

## Summary

- What are the advantages or disadvantages of refactoring code?

The main advantage is that refactoring allows us to develop code with better performance and less execution time. It also allows us to improve the usability of previously created codes so that we can adapt them to larger databases.

On the other hand, the main disadvantage is the time it can take to find a better way to rethink the code, improving the logic of the code, seeking to reduce the code and use less memory.

- How do these pros and cons apply to refactoring the original VBA script?

In our refactored code we change the logic of thinking of a loop within another loop by creating a variable (tickerIndex) that allows us to iterate within a single loop, which generated a code with better performance and less execution time.

Refactoring the code also allows us to better understand the code and identify certain functions that we could eliminate, finally the time consumed by refactoring could be to much, because sometime it is a trial and error process. However, the time invested can force us to find new ways of doing things and to improve our programming skills.
