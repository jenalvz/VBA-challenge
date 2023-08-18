Sub Stock_Market()

    Dim Total, Start, openprice, closeprice As Double
    
    Total = 0
    Start = 2
    openprice = Cells(2, "C").Value
    
    ' hard coding top row names
    Cells(1, "I").Value = "Ticker"
    
    
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Total = Total + Cells(i, "G").Value
    
        Ticker = Cells(i, 1).Value
    
        'checks to see if ticker on the line underneath is different from the one that we are on
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'populates tickers
            Cells(Start, "I").Value = Ticker
        
            'this gathers yearly close price
            closeprice = Cells(i, "F").Value
            
            'populates yearly change
            Cells(Start, "J").Value = closeprice - openprice
        
            If Cells(Start, "J").Value > 0 Then
                Cells(Start, "J").Interior.ColorIndex = 4
             
            Else
                Cells(Start, "J").Interior.ColorIndex = 3
             
            End If
             
            
            If openprice <> 0 Then
                Cells(Start, "K").Value = FormatPercent((closeprice - openprice) / openprice, 2)
                
            Else
                Cells(Start, "K").Value = Null
                
            End If
        
            Cells(Start, "L").Value = Total
        
            ' these reset the variables
            openprice = Cells(i + 1, "C").Value
            Start = Start + 1
            Total = 0
        
    
        End If
    Next i



End Sub
