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

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate

Dim Ticker As String
Dim Change As Double
Dim Percent_Change As Double
Dim Volume As Double
Dim Initial As Double
Dim P_CH As Double
Dim LastRow As Long
Dim LastRowB As Integer
Dim LastRowC As Integer
Dim Color As Integer
Dim max As Double
Dim min As Double
Dim Line As Integer

min = 0
max = 0
LastRow = Cells(Rows.Count, 1).End(xlUp).Row



Range("I1") = "Ticker"
Range("I1").Font.Bold = True
Range("J1") = "Yearly Change"
Range("J1").Font.Bold = True
Columns("J").ColumnWidth = 12
Range("K1") = "Percent Change"
Range("K1").Font.Bold = True
Columns("K").ColumnWidth = 15
Range("L1") = "Total Volume"
Range("L1").Font.Bold = True
Columns("L").ColumnWidth = 20
Range("N2") = "Greatest % Increase"
Range("N2").Font.Bold = True
Range("N3") = "Greatest % Decrease"
Range("N3").Font.Bold = True
Range("N4") = "Greatest Total Volume"
Range("N4").Font.Bold = True
Columns("N").ColumnWidth = 20
Range("O1") = "Ticker"
Range("O1").Font.Bold = True
Range("P1") = "Value"
Range("P1").Font.Bold = True
Columns("P").ColumnWidth = 20
Columns("M").ColumnWidth = 5

LastRowB = Cells(Rows.Count, 9).End(xlUp).Row
LastRowC = Cells(Rows.Count, 9).End(xlUp).Row

Range("P2:P3").NumberFormat = "0.00%"

  'Sets Initial as first Opening stock price
    
Initial = Cells(2, 3).Value

    
    
    
    'Sets Row for the first row used
    
Line = 2
    
    'For Loop to start on second row and run to end of data
      
For i = 2 To LastRow
    'When column 1 value is not equal then statement
    
    If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        
    'Perform these when column 1 value changes
    
        Closed = Cells(i, 6).Value
        Ticker = Cells(i, 1)
        Change = Cells(i, 6).Value - Initial
        'Percent_Change = Cells(i, 6).Value / Initial - 1
        Volume = Volume + Cells(i, 7).Value
                
                
        Range("I" & Line).Value = Ticker
        Range("J" & Line).Value = Change
        Range("K" & Line).Value = Percent_Change
        Range("L" & Line).Value = Volume
        Range("K" & Line).NumberFormat = "0.00%"
        
        Line = Line + 1
        Volume = 0
        Initial = Cells(i + 1, 3).Value
        
     'Column 1 still the same then add volume
     
    Else
        Volume = Volume + Cells(i, 7).Value
        
        
    End If
    
    Next i



'-------------------------------------------------------------------


For j = 2 To LastRowB
    
    'coloring cells with positive and negative values
    
    If Cells(j, 11).Value <= 0 Then
    
        Cells(j, 11).Interior.ColorIndex = 3 'Red
    
    Else
    
        Cells(j, 11).Interior.ColorIndex = 4 'Green
        
    End If
    
    Next j

'--------------------------------------------------------------------

For x = 2 To LastRowB



    'finding greatest percentage increase
    
    If Cells(x, 11).Value > max Then
    
        max = Cells(x, 11).Value
        Ticker = Cells(x, 9).Value
        
        Cells(2, 16) = max
        Cells(2, 15) = Ticker
    
    End If
    
    Next x
  max = 0
'---------------------------------------------------------------------

For y = 2 To LastRowB

    'finding greatest percentage decrease

    If Cells(y, 11).Value <= min Then
    
        min = Cells(y, 11).Value
        Ticker = Cells(y, 9).Value
        
        Cells(3, 16) = min
        Cells(3, 15) = Ticker
        
    End If
    
    Next y
  min = 0
'-----------------------------------------------------------------------

For Z = 2 To LastRowC

    'finding greatest total volume
    
    If Cells(Z, 12).Value > max Then
    
        max = Cells(Z, 12).Value
        Ticker = Cells(Z, 9).Value
        
        Cells(4, 16) = max
        Cells(4, 15) = Ticker
    
    End If
    
    Next Z



Next ws

End Sub

