Sub Calc()
      
    'Variables declaration
    Dim ws As Worksheet
    Dim Ticker As String
    Dim yearly_change As Double
    Dim open_price As Double
    Dim Close_price As Double
    Dim percent_change As Double
    Dim Total_stock_volume As Double
    Dim Fill_counter As Long
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total As Double
    
    
    'Looping through worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        'counter for filling data in new rows
        Fill_counter = 2
        
                
        'Find the last non-blank cell in column A(1)
        last_row_index = Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'formating
        
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("N1").Value = "Open Price"
        ws.Range("O1").Value = "Close price"
        
        
        'Initial values
        open_price = ws.Range("c2").Value
        Close_price = ws.Range("f2").Value
        Total_stock_volume = ws.Range("g2")
        
        For i = 2 To last_row_index
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              
                'Calculations
                yearly_change = Close_price - open_price
                percent_change = yearly_change / open_price
                
                'Filling the Cells
                ws.Cells(Fill_counter, 10).Value = Ticker
                ws.Cells(Fill_counter, 11).Value = yearly_change
                ws.Cells(Fill_counter, 12).Value = FormatPercent(percent_change)
                ws.Cells(Fill_counter, 13).Value = Total_stock_volume
                
                'the columns of these values will be hidden in the worksheets
                ws.Cells(Fill_counter, 14).Value = open_price
                ws.Cells(Fill_counter, 15).Value = Close_price
                
                'Conditional Formatting
                If yearly_change >= 0 Then
                    ws.Cells(Fill_counter, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(Fill_counter, 11).Interior.ColorIndex = 3
                End If
                
                                
                Fill_counter = Fill_counter + 1
                
                'Reset Values
                yearly_change = 0
                percent_change = 0
                Total_stock_volume = ws.Cells(i + 1, 7).Value
                open_price = ws.Cells(i + 1, 3).Value
                                
            Else
            
                Ticker = ws.Cells(i, 1).Value
                Total_stock_volume = Total_stock_volume + ws.Cells(i + 1, 7).Value
                Close_price = ws.Cells(i + 1, 6).Value
            
            End If
            
        
        Next i
    
        'Finding the last row number in column K(11)
        last_row_index2 = Cells(ws.Rows.Count, 11).End(xlUp).Row

        
        'Finding the Greatest % Increase
        Greatest_Increase = WorksheetFunction.Max(ws.Range("k2" & ":" & "k" & last_row_index2))
        ws.Range("t2").Value = Greatest_Increase
        ws.Range("r2").Value = "Greatest%Increase"
        
        
        'Finding the Greatest % Decrease
        
        Greatest_Decrease = WorksheetFunction.Min(ws.Range("k2" & ":" & "k" & last_row_index2))
        ws.Range("t3").Value = Greatest_Decrease
        ws.Range("r3").Value = "Greatest%Decrease"
        
        
        'Finding the Greatest Total Volume
        Greatest_Total = WorksheetFunction.Max(ws.Range("m2" & ":" & "m" & last_row_index2))
        ws.Range("t4").Value = Greatest_Total
        ws.Range("r4").Value = "Greatest Total Volume"
        
        
        'link between the value and Ticker
        For j = 2 To last_row_index2
        
            If ws.Cells(j, 11).Value = Greatest_Increase Then
            
                ws.Range("s2").Value = ws.Cells(j, 10).Value
                
            ElseIf ws.Cells(j, 11).Value = Greatest_Decrease Then
            
                ws.Range("s3").Value = ws.Cells(j, 10).Value
                
            ElseIf ws.Cells(j, 13).Value = Greatest_Total Then
            
                ws.Range("s4").Value = ws.Cells(j, 10).Value
                
            End If
        
        Next j
    
    
    Next ws

End Sub