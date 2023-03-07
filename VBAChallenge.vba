Attribute VB_Name = "Module1"
Sub Stock_Analysis():

' Loop through all sheets
For Each ws In Worksheets
ws.Select

' Set variable for Ticker
Dim Ticker As String

' Set variables for open and close value and yearly change
Dim Open_Value As Double
Dim Close_Value As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double

' Set an initial variable for holding the total stock volume per ticker
Dim StockTotal As Double
StockTotal = 0

' Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
' Last Row
LR = Cells(Rows.Count, 1).End(xlUp).Row

' Add the Column Headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
 
    ' Loop through all stocks
    For i = 2 To LR

        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            Open_Value = Cells(i, 3).Value
            
         End If
        ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            ' Set the Ticker
            Ticker = Cells(i, 1).Value
            
            ' Add to the Stock Volume Total
            StockTotal = StockTotal + Cells(i, 7).Value
            
            ' Print the Ticker in the Summary Table
            Range("I" & Summary_Table_Row).Value = Ticker
            
            ' Set close value
            Close_Value = Cells(i, 6)
            'MsgBox (Close_Value)
            Yearly_Change = Close_Value - Open_Value
            Percent_Change = Yearly_Change / Open_Value
            
            'Print yearly change to summary table
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            'Print percent change to summary table
            Range("K" & Summary_Table_Row).Value = Format(Percent_Change, "0.00%")
            
            ' Print the Stock Volume Amount to the Summary Table
            Range("L" & Summary_Table_Row).Value = StockTotal
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the Stock Total & Yearly Change
            StockTotal = 0
            ' Yearly_Change = 0
    
        ' If the cell immediately following a row is the same Ticker...
        Else
        
            ' Add to the Stock Volume Total
            StockTotal = StockTotal + Cells(i, 7).Value
        End If
        
    Next i
    
                 ' Formatting
                 Columns("J:J").Select
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
                        Formula1:="=0.001", Formula2:="=100"
                    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                    With Selection.FormatConditions(1).Interior
                        .PatternColorIndex = xlAutomatic
                        .Color = 5296274
                        .TintAndShade = 0
                    End With
                    Selection.FormatConditions(1).StopIfTrue = False
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
                        Formula1:="=-0.001", Formula2:="=-100"
                    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                    With Selection.FormatConditions(1).Interior
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                    End With
                    Selection.FormatConditions(1).StopIfTrue = False
                    Range("P12").Select
    
    ' Autofit to display data
    ws.Columns("I:L").AutoFit
    
Next ws
End Sub
