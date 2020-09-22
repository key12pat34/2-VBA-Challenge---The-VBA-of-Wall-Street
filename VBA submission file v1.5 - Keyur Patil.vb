Sub VBAHomework()
    For Each ws In Worksheets
        ws.Activate
        Call CalculateSummary
        
    Next ws
End Sub


Sub SetTitle()
    Range("I:Q").Value = ""
    Range("I:Q").Interior.ColorIndex = 0
    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    'this is for challenge only
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("I:O").Columns.AutoFit
End Sub


Sub CalculateSummary()
    Call SetTitle
    Call conditionalformat
    
    ''''''''''''''''''''''''Variables'''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lastrow As Long
        lastrow = ActiveSheet.UsedRange.Rows.Count
        
    Dim tickercount As Integer
       tickercount = 2
        
    Dim yearopening As Double
    Dim yearclosing As Double
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For currentrow = 2 To lastrow
        
        'initilizing yearopening and first ticker
        If currentrow = 2 Then
            yearopening = Cells(currentrow, 3).Value
            Cells(currentrow, 9).Value = Cells(currentrow, 1).Value
        End If
        
        'adding total stock volume
        If (Cells(currentrow, 1).Value = Cells(currentrow + 1, 1).Value) Then
            Cells(tickercount, 12).Value = (Cells(tickercount, 12).Value) + (Cells(currentrow, 7).Value)
        End If
      
            
        'checks <ticker> column for changes
        If (Cells(currentrow, 1).Value <> Cells(currentrow + 1, 1).Value) Then
            Cells(tickercount, 12).Value = (Cells(tickercount, 12).Value) + (Cells(currentrow, 7).Value)
            
            yearclosing = Cells(currentrow, 6).Value
            
            'yearly change
            Cells(tickercount, 10).Value = yearclosing - yearopening
            
            'percent change
            'If (Cells(tickercount, 10) <> 0) Then
            If (yearopening = 0) Then
                Cells(tickercount, 11).Value = Format(0, "0%")
            Else
                Cells(tickercount, 11).Value = FormatPercent((((Cells(tickercount, 10).Value) / yearopening)))
            End If
            
            '''''''''''''''''''''''''''''''''''''''''''''
            tickercount = tickercount + 1
            Cells(tickercount, 9).Value = Cells(currentrow + 1, 1).Value
            yearopening = Cells(currentrow + 1, 3).Value
            
        End If
       
    Next currentrow
    
    ''''''''''''''''''''''''''''Bonus'''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim maxval As Double, minval As Double, totalval As Double
    Dim maxid As String, minid As String, volid As String
    
    maxval = 0
    minval = 0
    totalval = 0
    
    For tickerrow = 2 To lastrow
        If (Cells(tickerrow, 11).Value > maxval) Then
            maxval = Cells(tickerrow, 11).Value
            maxid = Cells(tickerrow, 9).Value
            
        End If
        
        If (Cells(tickerrow, 11).Value < minval) Then
            minval = Cells(tickerrow, 11).Value
            minid = Cells(tickerrow, 9).Value
        End If
        
        If (Cells(tickerrow, 12).Value > totalval) Then
            totalval = Cells(tickerrow, 12).Value
            volid = Cells(tickerrow, 9).Value
        End If
    Next tickerrow
    
        Range("Q2") = FormatPercent(maxval)
        Range("P2") = maxid
        
        Range("Q3") = FormatPercent(minval)
        Range("P3") = minid

        Range("Q4") = totalval
        Range("P4") = volid
 

End Sub


Sub conditionalformat()

   'clears any formatting in J column
   Range("j:j").FormatConditions.Delete
    
    'conditional formatting for J column
    With Range("j:j").FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Interior.Color = 5296274
    End With

    With Range("j:j").FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
         .Interior.Color = 255
    End With
    Range("j1").FormatConditions.Delete
End Sub










