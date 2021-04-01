VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Stocks():

Dim Ticker As String
Dim Change As Double
Dim Percent_Change As Double
Dim Volume As Double
Dim Row As Integer
Dim Initial As Double
Dim P_CH As Double
Dim LastRow As String
Dim Color As Integer
Dim max As Double
Dim min As Double

    'Sets Initial as first Opeing stock price
    
Initial = Cells(2, 3).Value
    
    'Sets Row for the first row used
    
Row = 2
    
    'For Loop to start on second row and run to end of data
For i = 2 To 787
    
    'When column 1 value is not equal then statement
    
    If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        
    'Perform these when column 1 value changes
    
        Closed = Cells(i, 6).Value
        Ticker = Cells(i, 1)
        Change = Cells(i, 6).Value - Initial
        Percent_Change = Cells(i, 6).Value / Initial - 1
        Volume = Volume + Cells(i, 7).Value
                
                
        Range("I" & Row).Value = Ticker
        Range("J" & Row).Value = Change
        Range("K" & Row).Value = Percent_Change
        Range("L" & Row).Value = Volume
        
        Row = Row + 1
        Volume = 0
        Initial = Cells(i + 1, 3).Value
        
     'Column 1 still the same then add volume
     
    Else
        Volume = Volume + Cells(i, 7).Value
        
        
    End If
    
    Next i

'-------------------------------------------------------------------

For j = 2 To 4
    
    'coloring cells with positive and negative values
    
    If Cells(j, 11).Value > 0 Then
    
        Color = 4
        Cells(j, 11).Interior.ColorIndex = 4 'Green
    
    Else
    
        Cells(j, 11).Interior.ColorIndex = 3 'Red
        
    End If
    
    Next j

'--------------------------------------------------------------------

For x = 2 To 4

    'finding greatest percentage increase
    
    If Cells(x, 11).Value > max Then
    
        max = Cells(x, 11).Value
        Ticker = Cells(x, 9)
        
        Cells(2, 16) = max
        Cells(2, 15) = Ticker
    
    End If
    
    Next x
    
'---------------------------------------------------------------------

For y = 2 To 4

    'finding greatest percentage decrease

    If Cells(y, 11).Value < min Then
    
        min = Cells(y, 11).Value
        Ticker = Cells(y, 9).Value
        
        Cells(3, 16) = min
        Cells(3, 15) = Ticker
        
    End If
    
    Next y
        


End Sub
