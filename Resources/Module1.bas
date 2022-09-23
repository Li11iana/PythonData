Attribute VB_Name = "Module1"

Sub DQAnalysis()

    Worksheets("DQAnalysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"


    Worksheets("2018").Activate
    
    'Set initial volume to zero
    totalVolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'Establish the number of rows to loop over
        rowStart = 2
    'Row count fromhttps://stackoverflow.com/questions/18088729/row-count-where-data-exists
        rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
 
  
       

    'loop over all the rows
    
    For i = rowStart To rowEnd
    
    
    
        If Cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value

        End If
        
        
        'Checks if the previous row is NOT a DQ data row, but the current row IS a DQ data row meaning its the start of DQ data.
        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            startingPrice = Cells(i, 6).Value

        End If

        'Checks that current row is DQ data row but the next is NOT, meaning is the end od DQ data
        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If
        
        
    Next i
    
    'MsgBox (totalVolume)

    Worksheets("DQAnalysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = endingPrice / startingPrice - 1
    

End Sub



Sub AllStocksAnalysis()

'0. Choose year to analyze & start timer
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

'1. Format the output sheet on the "All Stocks Analysis" worksheet.

    Worksheets("AllStocksAnalysis").Activate
    
    Cells(1, 1).Value = "All Stocks (" + yearValue + ")"

    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'2. Initialize an array of all tickers.
    
    Dim tickers(11) As String
    
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
 
 
'3.1. Initialize variables for the starting price and ending price.

    Dim startingPrice As Single
    Dim endingPrice As Single
    
'3.2. Activate the data worksheet.

    Worksheets(yearValue).Activate
    
'3.3. Find the number of rows to loop over.
  
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

'4. Loop through the tickers

    
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
        
'5.Loop through rows in the data.
        
        Worksheets(yearValue).Activate
            For j = 2 To rowEnd
        
'5.1. Find the total volume for the current ticker.

                If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
                End If
        
'5.2. Find the starting price for the current ticker.
        
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
                End If


'5.3 Find the ending price for the current ticker.
            
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
                End If
                
            Next j
            

'6. Output the data for the current ticker.
            
   
    Worksheets("AllStocksAnalysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    
'7. Formatting analysis results
     Worksheets("AllStocksAnalysis").Activate
    
'' Bold headers
    Range("A3:C3").Font.Bold = True
    
'' Header Bottom edge border
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
'' Font size tittle
    Range("A1").Font.Size = 14
    
'' Number display format
    Range("B4:B15").NumberFormat = "#,##0.00"
    
'' Return in porcentage
    Range("C4:C15").NumberFormat = "0.00%"
    
'' Auto-fit column width to data

    Columns("B").AutoFit
    
''Conditional formating  Loop through Returns
    
    returnStart = 4
    returnEnd = 15

    For i = returnStart To returnEnd
    
''   x>0 green, x = 0 clear, x<0 red

        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
    
        Else
            Cells(i, 3).Interior.Color = xlNone
    
        End If
    
    Next i
    
'8.End timer and display results
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub


'Formatting

Sub formatAllStocksAnalysisTable()

    Worksheets("AllStocksAnalysis").Activate
    
'1. Bold headers
    Range("A3:C3").Font.Bold = True
    
'2. Header Bottom edge border
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
'3. Font size tittle
    Range("A1").Font.Size = 14
    
'4. Number display format
    Range("B4:B15").NumberFormat = "#,##0.00"
    
'5. Return in porcentage
    Range("C4:C15").NumberFormat = "0.00%"
    
'6. Auto-fit column width to data

    Columns("B").AutoFit
    
'7.Conditional formating
'7.1 Loop through Returns
    
    returnStart = 4
    returnEnd = 15

    For i = returnStart To returnEnd
    
'7.2   x>0 green, x = 0 clear, x<0 red

        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
    
        Else
            Cells(i, 3).Interior.Color = xlNone
    
        End If
    
    Next i
    
End Sub



