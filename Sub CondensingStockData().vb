Sub CondensingStockData()
    
    'Define all variables
    '--------------------------------------------------------
    'Define StockCount variable for table 2 rows
    Dim StockCount As Integer
    
    'Define Rowcount variable
    Dim RowCount As Long
    
    'Define Opening Value for first day of Year for stock
    Dim FirstOpen As Double
    
    'Define Closing Value for last day of the year for stock
    Dim LastClose As Double
    
    'Define Yearly change for stock
    Dim YearlyChange As Double
    
    'Define Percent change for stock
    Dim PercentChange As Double
    
    'Define Total volume for stock
    Dim TotalVolume As Double
    
    'Define Ticker symbol for current row
    Dim Ticker As String
    
    'Define Ticker symbol for next row
    Dim NextTicker As String
    
    'Define variable for greatest % Increase
    Dim MaxIncrease As Double
    
    'Define variable for max increase ticker
    Dim MaxIncreaseTicker As String
    
    'Define variable for greatest % decrease
    Dim MaxDecrease As Double
    
    'Define variable for max decrease ticker
    Dim MaxDecreaseTicker As String
    
    'Define variable for greatest total volume
    Dim MaxTotalVolume As Double
    
    'Define variable for max total volume ticker
    Dim MaxTotalVolumeTicker As String
    '----------------------------------------------------------
   
    'Loop through each worksheet
    '----------------------------------------------------------
    For Each ws In Worksheets
    
        'Reset values for new worksheet
        '------------------------------------------------------
        'Find the row number for the last row
        RowCount = ws.UsedRange.Rows.Count

    
        'Set first opening value for first stock to C2
        FirstOpen = ws.Cells(2, 3).Value
        
        'Set value for first row of data in new table
        StockCount = 2
        '------------------------------------------------------
        
        'Add Table 2 Column Headers
        '------------------------------------------------------
        'Add Column Header "Ticker"
        ws.Range("I1").Value = "Ticker"
        
        'Add Column Header "Yearly Change"
        ws.Range("J1").Value = "Yearly Change"
        
        'Add Column Header "Percent Change"
        ws.Range("K1").Value = "Percent Change"
        
        'Add Column Header "Total Volume"
        ws.Range("L1").Value = "Total Volume"
        '-------------------------------------------------------
        
        'Add new table 3 row labels and column labels
        '-------------------------------------------------------
        'Add row label "Greatest % Increase"
        ws.Range("O2").Value = "Greatest % Increase"
        
        'Add row label "Greatest % Decrease"
        ws.Range("O3").Value = "Greatest % Decrease"
        
        'Add row label "Greatest Total Volume"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Add column label "Ticker"
        ws.Range("P1").Value = "Ticker"
        
        'Add column label "Value"
        ws.Range("Q1").Value = "Value"
        '-------------------------------------------------------
    
        'Loop through each row in table 1 in current worksheet
        '-------------------------------------------------------
        For i = 2 To RowCount
        
            'Set all values for current row
            '---------------------------------------------------
            'Set Ticker equal to the value in cell i,3
            Ticker = ws.Cells(i, 1).Value
            
            'Set NextTicker equal to the value in cell i+1,3
            NextTicker = ws.Cells(i + 1, 1).Value
            '---------------------------------------------------
            
            'Conditional to find if the next row is the same stock
            '---------------------------------------------------
            If Ticker <> NextTicker Then
            
                'Find unique Ticker symbols and place into new table
                '-----------------------------------------------
                'Place Ticker symbol for current row onto new table
                ws.Cells(StockCount, 9).Value = Ticker
                '-----------------------------------------------
            
                'Find yearly change and put it into new table
                '-----------------------------------------------
                'Set Last Close equal to the value in cell i,6
                LastClose = ws.Cells(i, 6).Value
                
                'Subtract last close from first open to find yearly change
                YearlyChange = LastClose - FirstOpen
                
                'Place Yearly change into table at cell stockcount,10
                ws.Cells(StockCount, 10).Value = YearlyChange
                
                'Format yearly change cell with 2 decimals
                ws.Cells(StockCount, 10).NumberFormat = "0.00"
                
                'Conditional formatting to color yearly change cells red or Green
                '------------------------------------------------
                If YearlyChange >= 0 Then
                
                    'Format cells in green
                    ws.Cells(StockCount, 10).Interior.ColorIndex = 4
                    
                Else
                    
                    'Format cells in red
                    ws.Cells(StockCount, 10).Interior.ColorIndex = 3
                
                End If
                '-----------------------------------------------
                '-----------------------------------------------
                
                'Find Percent Change and Place and format in new table
                '-----------------------------------------------
                'Divide Yearly change by first open to find percent change
                PercentChange = YearlyChange / FirstOpen
                
                'Place percent change into table at cell stockcount, 11
                ws.Cells(StockCount, 11).Value = PercentChange
                
                'Format cell as percent
                ws.Cells(StockCount, 11).NumberFormat = "0.00%"
                '------------------------------------------------
                
                '-----------------------------------------------
                'Find Total Volume and put into table
                '-----------------------------------------------
                'Add last row of stock's volume to the total
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
                'Place totalvolume into table
                ws.Cells(StockCount, 12).Value = TotalVolume
                '-----------------------------------------------
                
                '-----------------------------------------------
                'Set values as needed to get ready for the next stock
                '-----------------------------------------------
                'Set Firstopen for the next stock equal to the cell value in the next row
                FirstOpen = ws.Cells(i + 1, 3).Value
                
                'Increase Stock count to move down to the next row in new table
                StockCount = StockCount + 1
                
                'Reset Total Volume to zero to get ready for next stock
                TotalVolume = 0
                '-----------------------------------------------
                
            Else
            
                'Add volume of current row to total volume for stock
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
            End If
            '----------------------------------------------------
            
        Next i
        '--------------------------------------------------------
        
        'set values for min, max and total volume to 0
        '--------------------------------------------------------
        'Set max increase equal to 0
        MaxIncrease = 0
            
        'Set max decrease equal to 0
        MaxDecrease = 0
        
        'Set Max total volume to 0
        MaxTotalVolume = 0
        '---------------------------------------------------------
        
        'Loop through table 2 to find Greatest % Increase, Greatest % decrease, and greatest total volume
        '---------------------------------------------------------
        For i = 2 To StockCount
            
            'Conditional to find if the percent on this row is bigger than current maxincrease
            '-----------------------------------------------------
            If ws.Cells(i, 11).Value >= MaxIncrease Then
            
                'when current row's percent change is bigger change max increase to that value
                MaxIncrease = ws.Cells(i, 11).Value
                
                'When bigger change the MaxIncreaseticker to the current row's ticker
                MaxIncreaseTicker = ws.Cells(i, 9).Value
                
            End If
            '------------------------------------------------------
            
            'Conditional to find if percent on this row is less than current maxdecrease
            '------------------------------------------------------
            If ws.Cells(i, 11).Value <= MaxDecrease Then
            
                'When current row's percent change is smaller change max decrease to that value
                MaxDecrease = ws.Cells(i, 11).Value
                
                'When smaller change the maxdecreaseticker to current row's ticker
                MaxDecreaseTicker = ws.Cells(i, 9).Value
            
            End If
            '-------------------------------------------------------
            
            'Conditional to find if max total volume on this row is greater than maxtotalvolume
            '-------------------------------------------------------
            If ws.Cells(i, 12).Value >= MaxTotalVolume Then
            
                'When current row's total volume is greater than maxtotalvolume change maxtotalvolume to that value
                MaxTotalVolume = ws.Cells(i, 12).Value
                
                'When higher change maxtotalvolumeticker to current row's ticker
                MaxTotalVolumeTicker = ws.Cells(i, 9).Value
                
            End If
            '-------------------------------------------------------
            
        Next i
        '-----------------------------------------------------------
        
        'Place values into table
        '-----------------------------------------------------------
        
        'place maxincreaseticker into table
        ws.Range("P2").Value = MaxIncreaseTicker
        
        'place maxincrease into table
        ws.Range("Q2").Value = MaxIncrease
        
        'Format maxincrease cell as percent
        ws.Range("Q2").NumberFormat = "0.00%"
        
        'place maxdecreaseticker into table
        ws.Range("P3").Value = MaxDecreaseTicker
        
        'place maxdecrease into table
        ws.Range("Q3").Value = MaxDecrease
        
        'Format maxdecrease cell as percent
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'place maxtotalvolumeticker into table
        ws.Range("P4").Value = MaxTotalVolumeTicker
        
        'place maxtotalvolume into table
        ws.Range("Q4").Value = MaxTotalVolume
        
        '------------------------------------------------------------
        
        'Format Column widths to fit data
        '------------------------------------------------------------
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("O:Q").EntireColumn.AutoFit
        '------------------------------------------------------------
        
    Next ws
    '----------------------------------------------------------------
            
End Sub
