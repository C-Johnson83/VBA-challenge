Sub vbachallenge()

    'turning these off speeds up the macro
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
     
        For Each Sheet In Sheets
        Sheet.Select
        
    'assigns the variables
            Dim lastrow As Double
            Dim vol As Double
            Dim irow As Double
            Dim yrclose As Double
            Dim yropen As Double
            Dim pinc As Double
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            vol = 0
            irow = 2
        
    ' inserts headers into designated cells
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Changes"
            Cells(1, 12).Value = "Total Stock Volume"
                  
    'loops through sheet,finds unique tickers, and places them in column 9
            For i = 2 To lastrow
                ticker = Cells(i, 1).Value
                tickers = Cells(i - 1, 1).Value
                
            If ticker <> tickers Then
                  Cells(irow, 9).Value = ticker
                  irow = irow + 1
                  
            End If
               
        Next i
               
    'loops through sheet, gets the combined total volume of each unique ticker, and places them in column 12
            irow = 2
            For i = 2 To lastrow
                ticker = Cells(i, 1).Value
                tickers = Cells(i - 1, 1).Value
                
            If ticker = tickers And i > 2 Then
                vol = vol + Cells(i, 7).Value
            ElseIf i > 2 Then
                Cells(irow, 12).Value = vol
                irow = irow + 1
                  vol = 0
            Else
                  vol = vol + Cells(i, 7).Value
          
            End If
              
        Next i
                
    'Loops through the sheet, finds the open to close difference and places it in colum 10, and finds the_
    'percentage change of that difference, and places it in colum 11
       
            irow = 2
            For i = 2 To lastrow
              
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                  yrclose = Cells(i, 6).Value
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                  yropen = Cells(i, 3).Value
                  
            End If
                
            If yropen > 0 And yrclose > 0 Then
                diff = yrclose - yropen
                pinc = diff / yropen
                Cells(irow, 10).Value = diff
                Cells(irow, 11).Value = FormatPercent(pinc)
                yrclose = 0
                yropen = 0
                irow = irow + 1
                  
            End If
                
        Next i
                   
    'applies conditional formatting for the percent change difference. If a positive difference, then change to green_
    'if a negative difference, then change to red
            For i = 2 To lastrow
             
            If IsEmpty(Cells(i, 10).Value) Then Exit For
            If Cells(i, 10).Value >= 0 Then
                   Cells(i, 10).Interior.Color = vbGreen
            Else:
                   Cells(i, 10).Interior.Color = vbRed
                   
            End If
                
        Next i
              
    ' Sets variables and range for min and max functions
            mxdiff = WorksheetFunction.Max(Columns("K"))
            mndiff = WorksheetFunction.Min(Columns("K"))
            mxvol = WorksheetFunction.Max(Columns("L"))
              
    'Inserts data labels for min and max table, places min and max values in column L_
    'and formats min and max cells to percentages.
            Range("O2").Value = "Greatest % increase"
            Range("O3").Value = "Greatest % decrease"
            Range("o4").Value = "Greatest total volume"
            Range("Q2").Value = FormatPercent(mxdiff)
            Range("Q3").Value = FormatPercent(mndiff)
            Range("Q4").Value = mxvol
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
              
    'loops through and assigns the Ticker associated with the min, max, and total values
            For i = 2 To lastrow
            
            If mxdiff = Cells(i, 11).Value Then
                  Range("P2").Value = Cells(i, 9).Value
            ElseIf mndiff = Cells(i, 11).Value Then
                  Range("P3").Value = Cells(i, 9).Value
            ElseIf mxvol = Cells(i, 12).Value Then
                  Range("P4").Value = Cells(i, 9).Value
                  
            End If
             
        Next i
              
    'adjust columns to be able to see the full header
            Range("I:Q").EntireColumn.AutoFit
               
        Next Sheet
         
    'goes back to the first sheet
        Sheets("2018").Select
         
    ' turns the applications back on
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub
