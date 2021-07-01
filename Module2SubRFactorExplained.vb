Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    YearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + YearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    
     Sheets(YearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Created a tickerIndex and initialized it to equal "0" for later access
    'Also combining the tickers array with the arrays created below in section 1B.
            
        tickerindex = 0
        
    '1b) Created 3 arrays using "Dimensions" command to declare the variable and the type of variable.  
    'These output arrays will be used to store the data we assign to the appropriate Dimension (in sections 3A - 3C).

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Created a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        
        tickerVolumes(i) = 0
       
    Next i
    'We used "Next i" here, to end the loop, before we move onto the next loop, instead of nesting it within the next loop below
    'which should eliminate processing more iterations in section 2B.

    '2b) Loop over all the rows in the spreadsheet.
    '(In this section, we continue eliminating unnecessary iterations to further "refactor" and repurpose the code.
    'By using the "tickerIndex" we are able to store the data in separate outputs arrays(section 1B) with it's own loop(2A), in which we are able to access
    'upon command or indication.)
        
           For i = 2 To RowCount 'As shown here, there is only one "For Loop" and no other "nested For Loops" below before ending "Next i as opposed to the orignal"
                            
                '3a) Increase volume for the current ticker
                
                If Cells(i, 1).Value = tickers(tickerindex) Then
                
                tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
                
                End If
                
        '3b) Check if the current row is the first row with the selected tickerIndex.
    
                 If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
                 
                 tickerStartingPrices(tickerindex) = Cells(i, 6).Value

                 End If
        
        '3c) Check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.

                 If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then

                 tickerEndingPrices(tickerindex) = Cells(i, 6).Value
                        
            '3d Increase the tickerIndex.
        
                tickerindex = tickerindex + 1
            
            End If
            
    Next i

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return. 
    'This is where we output the data collected onto a separate sheet for our client to view the data more clearly

Worksheets("All Stocks Analysis").Activate

    For i = 0 To 11
           
            Cells(i + 4, 1).Value = tickers(i)
            Cells(i + 4, 2).Value = tickerVolumes(i)
            Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
            
    Next i
    
    'Formatting: We are using condtional formatting with font and colors, which help our client indicate what stocks had a positive returns versus negative returns. 
    
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For b = dataRowStart To dataRowEnd
        
        If Cells(b, 3) > 0 Then
            
            Cells(b, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(b, 3).Interior.Color = vbRed
            
        End If
        
    Next b
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (YearValue)

End Sub
