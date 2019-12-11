Attribute VB_Name = "Module1"
  Sub vbahomework()
  
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
 
    Dim tickersymbol As String
    
    Dim yearlychange As Double
    
    Dim percentchange As Double
    
    Dim totalstockvolume As Double
    
    Dim summarytablerow As Integer
    
    summarytablerow = 2
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Dim openv As Double

Dim closev As Double

'Sets opening value before For Loop is initiated
openv = Cells(2, 3).Value

totalstockvolume = 0


For i = 2 To lastrow

'Sets closing value to last cell encountered in For Loop before a change is detected in column A
closev = Cells(i, 6).Value


    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        tickersymbol = Cells(i, 1).Value
        Range("I" & summarytablerow).Value = tickersymbol
        
        yearlychange = openv - closev
        Range("J" & summarytablerow).Value = yearlychange
        
        percentchange = (1 - (closev / openv)) * 100
        Range("K" & summarytablerow).Value = percentchange
        
        totalstockvolume = stockvolume + Cells(i, 7).Value
        Range("L" & summarytablerow).Value = totalstockvolume
        
        
        'Nested if for conditional formatting
        If yearlychange < 0 Then
            Range("J" & summarytablerow).Interior.ColorIndex = 3
        Else
            Range("J" & summarytablerow).Interior.ColorIndex = 4
        End If
        
        summarytablerow = summarytablerow + 1
        
        
        
        
    yearlychange = 0
        
    openv = Cells(i + 1, 3).Value
    
    
    Else
        
        totalstockvolume = totalstockvolume + Cells(i, 7).Value
    
        
    End If
    

Next i


End Sub
