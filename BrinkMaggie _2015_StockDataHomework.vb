Sub Stock2015()
'create variable, and incremetor..then make the increment equal to 0.
'set up end table by creating a variable that will keep track of the rows and set that equal to the first row

Dim Ticker_Type As String
Dim Total_Volume As Double

Total_Volume = 0

Dim Final_Table_Row As Integer
Final_Table_Row = 2

Dim StartPrice As Double

Dim ClosedPrice As Double

Dim Yearly_Change As Double
Yearly_Change = 0

Dim Percent_Change As Double
Percent_Change = 0

Dim Max As Double

StartPrice = Cells(2, 3).Value

'Coloration section

For i = 2 To 760192

For j = 11 To 11

If (Cells(i, j) > 0) Then
Cells(i, j).Interior.ColorIndex = 4

Else

Cells(i, j).Interior.ColorIndex = 3

End If

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker_Type = Cells(i, 1).Value
        
        Total_Volume = Total_Volume + Cells(i, 7).Value
        
        ClosedPrice = Cells(i, 6).Value
        
        Yearly_Change = ClosedPrice - StartPrice
        
        If StartPrice = 0 Then
        
        Percent_Change = 0
        
        Else
        
        Percent_Change = ((Yearly_Change) / StartPrice) * 100
        
        End If
        
        StartPrice = Cells(i + 1, 3).Value
        
        Range("K" & Final_Table_Row).Value = Yearly_Change
        
        Range("I" & Final_Table_Row).Value = Ticker_Type
        
        Range("J" & Final_Table_Row).Value = Total_Volume
        
        Range("L" & Final_Table_Row).Value = Percent_Change
        
        Final_Table_Row = Final_Table_Row + 1
        
        Total_Volume = 0
    
    Else
        Total_Volume = Total_Volume + Cells(i, 7).Value
        
    End If
    
   
Next j
Next i

End Sub

