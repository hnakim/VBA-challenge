Sub stock_market_data()

'set initial variable for ticker
Dim ticker As String

'set initial variable for volume and summary table row using j
Dim volume As Double
Dim j As Long
volume = 0
j = 2
'setting initial variable to store row count for opening value
k = 2

'set initial variables for opening and closing prices
Dim opening As Double
Dim closing As Double
Dim change As Double
Dim percentage As Double

'set lastrow
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'set columns for summary table
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'loop through all the stock data
For i = 2 To lastrow

    'if ticker is same
    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    
        'add volume total
        volume = volume + Cells(i, 7).Value
        
    'if ticker changes
    Else
  
        'add to volume total
        volume = volume + Cells(i, 7).Value
        
        'print to summary table
        ticker = Cells(i, 1).Value
        Range("I" & j).Value = ticker
        Range("L" & j).Value = volume

            'determine opening price
            opening = Cells(k, 3).Value
            
            'determine closing price
            closing = Cells(i, 6).Value
             
            'calculate difference to determine yearly change
            change = closing - opening
            
            'calculate percent change
            percentage = change / opening
            
            'format percentage
            formatpercentage = FormatPercent(percentage)
          
            'print to table
            Range("J" & j).Value = change
            Range("K" & j).Value = formatpercentage
            
            'color condition
            If change > 0 Then
            Range("J" & j).Interior.ColorIndex = 4
            
            Else
            Range("J" & j).Interior.ColorIndex = 3
            
            End If
            
            'add one to the summary table
            j = j + 1
            
            'reset the volume total
            volume = 0
            
            'store row number
            k = i + 1
                  
            
    End If
    
Next i

End Sub

