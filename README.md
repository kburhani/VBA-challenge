# VBA-challenge
Assignment 2, VBA Challenge

VBA script:

Sub VBAChallenge()

For Each ws In Worksheets

    Dim total As Double
    Dim summaryrow As Integer
    Dim ticker_count As Long
    
    'Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Bonus headers :'D
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest total volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    total = 0
    summaryrow = 2
    
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    'instruct to loop through rows
    For i = 2 To lastrow
        'is next cell same or not
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            total = total + Cells(i, 7).Value
            'ticker column entry
            Range("I" & 2 + summaryrow).Value = Cells(i, 1).Value
            'total stock column entry
            Range("L" & 2 + summaryrow).Value = total
            summaryrow = summaryrow + 1
            total = 0
            
        Else
            total = total + Cells(i, 7).Value
        
                'Calc and write Yearly Change in column J/10
                'Cells(ticker_count, 10).Value = Cells(i, 6).Value - Cells(j, 3).Value
                
                    'Conditional formating
                    'If Cells(ticker_count, 10).Value < 0 Then
                
                    'Set cell background color to red
                    'Cells(ticker_count, 10).Interior.ColorIndex = 3
                
                    'Else
                
                    'Set cell background color to green
                    'Cells(ticker_count, 10).Interior.ColorIndex = 4
                
                    'End If
        
        End If
        
        
    Next i
    
Next ws
End Sub
