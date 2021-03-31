Attribute VB_Name = "Module1"
Sub Attempt3()
' loop through worksheets

    For Each ws In Worksheets

'Define new titles

    Dim Title1 As String
    Title1 = "Ticker"
    ws.Range("I1").Value = Title1
    
    Dim Title2 As String
    Title2 = "Yearly Change"
    ws.Range("J1").Value = Title2
    
    Dim Title3 As String
    Title3 = "Percent Change"
    ws.Range("K1").Value = Title3
    
    Dim Title4 As String
    Title4 = "Total Stock Volume"
    ws.Range("L1") = Title4


    ws.Range("I1:L1").Font.Bold = True
    
'Find last row for Column A and Column I

    Dim LastRowLong As Long
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim LastRowShort As Long
    
        LastRowShort = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
'Define Variables

    Dim Ticker_ID As String
    
    Dim Yearly_Change As Double
        
        Yearly_Change = 0
    
    Dim Percent_Change As Long
        
        Percent_Change = 0
    
    
    Dim Total_Stock_Volume As Double
        
        Total_Stock_Volume = 0
    
    Dim Summary_Table_Row As Integer
        
        Summary_Table_Row = 2
        
    Dim Starting_number As Long
        Starting_number = 2

'Set up Sheet Loop
For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        Ticker_ID = ws.Cells(i, 1).Value
        
        Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(Starting_number, 3).Value
        
        ' Percent_Change = (ws.Cells(i, 6).Value / ws.Cells(Starting_number, 3).Value) - 1
        
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        ws.Range("I" & Summary_Table_Row).Value = Ticker_ID
        
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        Yearly_Change = 0
        
        Percent_Change = 0
        
        
        Total_Stock_Volume = 0
        
        Starting_number = i + 1
        
    
 Else
    
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
    
    End If
    
  Next i
    
'Loop for pulled data

For j = 2 To LastRowShort

    If ws.Cells(j, 10).Value < 0 Then
        
        ws.Cells(j, 10).Interior.ColorIndex = 3
        
    Else
    
        ws.Cells(j, 10).Interior.ColorIndex = 4
        
    End If
    
        
Next j
    
Next ws

End Sub
