Sub Stock_insights()

'create variables for ticker, open price, close price, total stock volume and ticker counter

Dim tickSym As String
Dim oPrice As Double
Dim cPrice As Double
Dim totStockVol As Double
Dim tickCount As Long

'create column headers

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'get the symbol and opening price for the year of the first symbol
tickCount = 2


'Loop through all sheets

 For Each ws In Worksheets
 
  tickSym = ws.Cells(2, 1).Value
  oPrice = ws.Cells(2, 3).Value
  

'loop through ticker symbols

    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        
        'begin testing for the next symbol
        
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         
           cPrice = ws.Cells(i, 6).Value
         
          'write out ticker, yearly change, percent change, and total stock volume
          
         
          Cells(tickCount, 9).Value = tickSym
         
          Cells(tickCount, 10).Value = cPrice - oPrice
          
          'color the cell red if negative, green if positive
          
            If (cPrice - oPrice < 0) Then
            
            
              Cells(tickCount, 10).Interior.ColorIndex = 3
            
            Else
              
              Cells(tickCount, 10).Interior.ColorIndex = 4
            
            End If
            
         
          Cells(tickCount, 11).Value = (cPrice - oPrice) / oPrice
          Cells(tickCount, 11).NumberFormat = "0.00%"
         
          Cells(tickCount, 12).Value = totStockVol
         
         'increment tickCount and change tickSym to the next symbol and zero out total stock volume
         
         'get next open price
         
          oPrice = ws.Cells(i + 1, 3).Value
         
         tickCount = tickCount + 1
         tickSym = ws.Cells(i + 1, 1).Value
         totStockVol = 0
         
         Else
         
         totStockVol = totStockVol + ws.Cells(i, 7).Value
         
         End If
         
        Next i
        
    Next ws
    

End Sub


