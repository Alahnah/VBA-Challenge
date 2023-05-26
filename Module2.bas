Attribute VB_Name = "Module1"

Sub Hw2()

For Each ws In Worksheets

' Setting Ticker and stock volume
Dim Worksheetname As String
Dim Ticker As String
Dim Stock_Volume As Double
Stock_Volume = 0
Dim Beg_Price As Double
Dim End_Price As Double
End_Price = 0
Dim Yearly_Change As Double
Yearly_Change = 0
Dim Percent_Change As Double
Dim Summary_Row As Double
Summary_Row = 2
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Beg_Price = Cells(2, 3).Value

Worksheetname = ws.Name

'Column Headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To LastRow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker = ws.Cells(i, 1).Value
    End_Price = ws.Cells(i, 6).Value
    
    Yearly_Change = End_Price - Beg_Price
    Percent_Change = (Yearly_Change / Beg_Price)

    
    ' Add to the Stock Total
    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

    ws.Range("I" & Summary_Row).Value = Ticker
    ws.Range("L" & Summary_Row).Value = Stock_Volume
    ws.Range("J" & Summary_Row).Value = Yearly_Change
    ws.Range("K" & Summary_Row).Value = Percent_Change

    ' Update to next ticker
    Summary_Row = Summary_Row + 1
    Stock_Volume = 0
    Beg_Price = ws.Cells(i + 1, 3).Value
    
    
    Else
    
    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    End If
    
     'Conditional Formatting
     
     If ws.Cells(i, "J").Value < 0 Then
     ws.Cells(i, "J").Interior.ColorIndex = 3
    
     Else: ws.Cells(i, "J").Interior.ColorIndex = 4
     
     End If
    
   ws.Cells(i, "K").Value = Format(Percent_Change, "Percent")
    
Next i

'% Summary
    
Greatest_Increase = ws.Cells(2, "K").Value
Greatest_Decrease = ws.Cells(2, "K").Value
Greatest_Total_Volume = ws.Cells(2, "L").Value

For i = 2 To LastRow

    If ws.Cells(i, "K").Value > Greatest_Increase Then
    Greatest_Increase = ws.Cells(i, "K").Value

    ws.Cells(2, 15).Value = ws.Cells(i, "I").Value

    Else

    Greatest_Increase = Greatest_Increase

    End If

    If ws.Cells(i, "K").Value < Greatest_Decrease Then
    Greatest_Decrease = ws.Cells(i, "K").Value
    
    ws.Cells(3, 15).Value = ws.Cells(i, "I").Value

    Else: Greatest_Decrease = Greatest_Decrease

    End If

    If ws.Cells(i, "L") > Greatest_Total_Volume Then
    Greatest_Total_Volume = ws.Cells(i, "L").Value

    ws.Cells(4, 15).Value = ws.Cells(i, "I").Value

    End If


ws.Cells(2, 16).Value = Format(Greatest_Increase, "Percent")
ws.Cells(3, 16).Value = Format(Greatest_Decrease, "percent")
ws.Cells(4, 16).Value = Format(Greatest_Total_Volume, "Scientific")


Next i

'Headers

ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"


Next ws


End Sub


