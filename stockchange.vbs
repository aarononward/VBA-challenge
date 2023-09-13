Sub StockChange()
    
    'Defines Variables
    Dim i As Long
    Dim table As Integer
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim percentChange As String
    Dim TSV As Double
    Dim openrow As Long
    
    'Runs Macro over each worksheet
    For Each ws In Worksheets
    
    'Set titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greastest Total Volume"
    
    
  
    
        'creates the summary table
        table = 2
        openrow = 2
        'Counts rows for summary table
        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        lastrow2 = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
            For i = 2 To lastrow

                'Checks for different value in subsequent cell
                If ws.Cells(i + 1, 1) <> ws.Cells(i, 1).Value Then
                    
                   
                    'Sets Ticker variable  equal to the Ticker name in Coulumn A
                    Ticker = ws.Cells(i, 1).Value
        
                    'Fills cells in Column I with Ticker names
                    ws.Range("I" & table).Value = Ticker
        
                    '_________________________________________
                    '
                    'For calulating Yearly Change
                    
                    
                    
                    'Chooses Year Start Value
                    OpenValue = ws.Cells(openrow, 3).Value
                    
                    'Chooses Year End Value
                    CloseValue = ws.Cells(i, 6).Value
                    
                    'Calculates Yearly Change
                    YearlyChange = CloseValue - OpenValue
                                        
                    'Puts change Yearly change Value into table
                    ws.Range("J" & table).Value = YearlyChange
                    
                    
                    '___________________________________________
                    '
                    
                    'Calculates the percent change for the year
                    percentChange = (YearlyChange / OpenValue)
                                     
                    'Changes Column format to percentage
                    percentChange = FormatPercent(percentChange)
                    
                    'Fills Percent Change Column
                    ws.Range("K" & table).Value = percentChange
                    
                    '___________________________________________
                    '
                    
                    
                    'Calculates Total Stock Volume
                    TSV = ws.Cells(i, 7) + TSV
                    ws.Range("L" & table).Value = TSV
                    
                    
                    'Moves to next line in Summary Table
                    table = table + 1
                    openrow = i + 1
                    TSV = 0
                Else:
                    
                    TSV = ws.Cells(i, 7) + TSV
           End If
           
        
            'Adds Green for positive change
            '_______________________________
            If Not IsEmpty(ws.Cells(i, 10)) Then
                If ws.Cells(i, 10).Value > 0 Then
                    With ws.Cells(i, 10).Interior
                          .ColorIndex = 4
                    End With
                            
                ElseIf ws.Cells(i, 10).Value = 0 Then
                    With ws.Cells(i, 10).Interior
                            .ColorIndex = 6
                    End With
                            
                ElseIf ws.Cells(i, 10).Value < 0 Then
                    With ws.Cells(i, 10).Interior
                            .ColorIndex = 3
                    End With
                            
                End If
            End If
        
        
        Next i
         
         'Inserts Value Column
         ws.Range("Q1").Value = "Value"
        
        '_______________________________________________
        
         
        'Finds Greatest % Increase
        Max = Application.WorksheetFunction.Max(ws.Range("K:K"))
        ws.Range("Q2") = Max
        ws.Range("Q2") = FormatPercent(ws.Range("Q2"))
        
        'Adds name of greatest increased stock to second Ticker column
        ws.Range("P2").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K:K"), 0))
        
        '_______________________________________________
 
        
        'Finds Greatest % Decreased
        Min = Application.WorksheetFunction.Min(ws.Range("K:K"))
        ws.Range("Q3") = Min
        ws.Range("Q3") = FormatPercent(ws.Range("Q3"))
        
        'Adds name of greatest decreased stock to second ticker column
        ws.Range("P3").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K:K"), 0))
        '_______________________________________________
        
         
        'Finds Greatest Total Volume
        TSV = Application.WorksheetFunction.Max(ws.Range("L:L"))
        ws.Range("Q4") = TSV
        
        'Adds name of stock with greatest total volume to second ticker column
        
        ws.Range("P4").Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L:L"), 0))

         
        'Ensures full cell value fits into cell for new rows
         ws.Columns("I:Q").AutoFit
                                                  

    Next ws
End Sub
