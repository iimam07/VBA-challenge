VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBA_challenge_ticker()

Dim WS As Worksheet
    On Error Resume Next
    For Each WS In ThisWorkbook.Worksheets
        
        ' find the last row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' summary definitions
        WS.Cells(1, "I").Value = "Ticker"
        WS.Cells(1, "J").Value = "Yearly Change"
        WS.Cells(1, "K").Value = "Percent Change"
        WS.Cells(1, "L").Value = "Total Stock Volume"
        
        ' variable values
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Integer
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        ' find initial price
        Open_Price = WS.Cells(2, Column + 2).Value

        
        For i = 2 To LastRow
            If WS.Cells(i + 1, Column).Value <> WS.Cells(i, Column).Value Then
                ' det ticker name
                Ticker_Name = WS.Cells(i, Column).Value
                WS.Cells(Row, Column + 8).Value = Ticker_Name
                ' det close price
                Close_Price = WS.Cells(i, Column + 5).Value
                ' det yearly change
                Yearly_Change = Close_Price - Open_Price
                WS.Cells(Row, Column + 9).Value = Yearly_Change
                ' det percent change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    WS.Cells(Row, Column + 10).Value = Percent_Change
                    WS.Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                ' det total volume
                Volume = Volume + WS.Cells(i, Column + 6).Value
                WS.Cells(Row, Column + 11).Value = Volume
                ' Add one to the summary table row
                Row = Row + 1
                ' reset the Open Price
                Open_Price = WS.Cells(i + 1, Column + 2)
                ' reset volumn total
                Volume = 0
            
            'if cells are the same ticker
            Else
                Volume = Volume + WS.Cells(i, Column + 6).Value
            End If
        Next i
        
        ' det law row of yearly change per ws
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        ' set cell colors:
        For j = 2 To YCLastRow
            If (WS.Cells(j, Column + 9).Value > 0 Or WS.Cells(j, Column + 9).Value = 0) Then
                WS.Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf WS.Cells(j, Column + 9).Value < 0 Then
                WS.Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
       
    Next WS
        
End Sub

