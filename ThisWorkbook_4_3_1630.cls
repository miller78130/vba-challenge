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



LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row



ws.Range("I1") = "Ticker"
ws.Range("I1").Font.Bold = True
ws.Range("J1") = "Yearly Change"
ws.Range("J1").Font.Bold = True
ws.Columns("J").ColumnWidth = 12
ws.Range("K1") = "Percent Change"
ws.Range("K1").Font.Bold = True
ws.Columns("K").ColumnWidth = 15
ws.Range("L1") = "Total Volume"
ws.Range("L1").Font.Bold = True
ws.Columns("L").ColumnWidth = 20
ws.Range("N2") = "Greatest % Increase"
ws.Range("N2").Font.Bold = True
ws.Range("N3") = "Greatest % Decrease"
ws.Range("N3").Font.Bold = True
ws.Range("N4") = "Greatest Total Volume"
ws.Range("N4").Font.Bold = True
ws.Columns("N").ColumnWidth = 20
ws.Range("O1") = "Ticker"
ws.Range("O1").Font.Bold = True
ws.Range("P1") = "Value"
ws.Range("P1").Font.Bold = True
ws.Columns("P").ColumnWidth = 20
ws.Columns("M").ColumnWidth = 5

ws.Range("P2:P3").NumberFormat = "0.00%"

    'Sets Initial as first Opening stock price
    
Initial = ws.Cells(2, 3).Value
    
    'Sets Row for the first row used
    
Line = 2
    
    'For Loop to start on second row and run to end of data
For i = 2 To LastRow
    
    'When column 1 value is not equal then statement
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        
    'Perform these when column 1 value changes
    
        Closed = ws.Cells(i, 6).Value
        Ticker = ws.Cells(i, 1)
        Change = Closed - Initial
        Percent_Change = Closed / Initial - 1
        Volume = Volume + ws.Cells(i, 7).Value
                
                
        ws.Range("I" & Line).Value = Ticker
        ws.Range("J" & Line).Value = Change
        ws.Range("K" & Line).Value = Percent_Change
        ws.Range("K" & Line).NumberFormat = "0.00%"
        ws.Range("L" & Line).Value = Volume
        
        Line = Line + 1
        Volume = 0
        Initial = ws.Cells(i + 1, 3).Value
        
     'Column 1 still the same then add volume
     
    Else
        Volume = Volume + ws.Cells(i, 7).Value
        
        
    End If
    
    Next i

'-------------------------------------------------------------------

LastRowB = ws.Cells(Rows.Count, 9).End(xlUp).Row
LastRowC = ws.Cells(Rows.Count, 9).End(xlUp).Row


For j = 2 To LastRowB
    
    'coloring cells with positive and negative values
    
    If ws.Cells(j, 11).Value > 0 Then
    
        Color = 4
        ws.Cells(j, 11).Interior.ColorIndex = 4 'Green
    
    Else
    
        ws.Cells(j, 11).Interior.ColorIndex = 3 'Red
        
    End If
    
    Next j

'--------------------------------------------------------------------

For x = 2 To LastRowB



    'finding greatest percentage increase
        max = 0
    If ws.Cells(x, 11).Value > max Then
    
        max = ws.Cells(x, 11).Value
        Ticker = ws.Cells(x, 9).Value
        
        Cells(2, 16) = max
        Cells(2, 15) = Ticker
     
    End If
    
    Next x
    
'---------------------------------------------------------------------

For y = 2 To LastRowB

    'finding greatest percentage decrease
        min = 0
    If ws.Cells(y, 11).Value < min Then
    
        min = ws.Cells(y, 11).Value
        Ticker = ws.Cells(y, 9).Value
        
        Cells(3, 16) = min
        Cells(3, 15) = Ticker
        
    End If
    
    Next y
   
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





