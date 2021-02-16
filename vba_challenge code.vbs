Sub Stock_Total()
    Dim ws As Worksheet
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim i As Double
        Dim Ticker_Total As Double
        Ticker_Total = 0
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        Dim Ticker_Name As String
        Dim Ticker_Open As Single
        Ticker_Open = ws.Cells(2, 3).Value
        Dim Ticker_Close As Single
        Dim Ticker_Percent As Double
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            Ticker_Name = ws.Cells(i, 1).Value
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            Ticker_Close = ws.Cells(i, 6).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            ws.Range("N" & Summary_Table_Row).Value = Ticker_Total
            ws.Range("J" & Summary_Table_Row).Value = Format(Ticker_Open, "0.00")
            ws.Range("K" & Summary_Table_Row).Value = Format(Ticker_Close, "0.00")
            ws.Range("L" & Summary_Table_Row).Value = Format((Ticker_Close - Ticker_Open), "0.00")
                    If ws.Range("J" & Summary_Table_Row).Value = 0 Then
                    ws.Range("M" & Summary_Table_Row).Value = Null
                        Else
                        Ticker_Percent = ws.Range("L" & Summary_Table_Row).Value / ws.Range("J" & Summary_Table_Row).Value
                        ws.Range("M" & Summary_Table_Row).Value = Format((Ticker_Percent), "0.00%")
                    End If
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Ticker Open"
            ws.Range("K1").Value = "Ticker Close"
            ws.Range("L1").Value = "Yearly Change"
            ws.Range("M1").Value = "Percent Change"
            ws.Range("N1").Value = "Total Stock Volume"
            Summary_Table_Row = Summary_Table_Row + 1
            Ticker_Total = 0
            Ticker_Open = ws.Cells(i + 1, 3).Value
            Else
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            End If
    Next i
    Dim MaxVal As Double
    Dim MinVal As Double
    Dim RowNum1, RowNum2, RowNum3 As Integer
    Dim Maxvol As Double
    MaxVal = Application.WorksheetFunction.max(ws.Range("M:M"))
    ws.Range("R2").Value = Format(MaxVal, "0.00%")
    'Find row number of Max
    RowNum1 = Application.Match(MaxVal, ws.Range("M:M"), 0)
    ws.Range("Q2").Value = Range("I" & RowNum1)
    MinVal = Application.WorksheetFunction.Min(ws.Range("M:M"))
    ws.Range("R3").Value = Format(MinVal, "0.00%")
    RowNum2 = Application.Match(MinVal, ws.Range("M:M"), 0)
    ws.Range("Q3").Value = Range("I" & RowNum2)
    Maxvol = Application.WorksheetFunction.max(ws.Range("N:N"))
    ws.Range("R4").Value = Maxvol
    RowNum3 = Application.Match(Maxvol, ws.Range("N:N"), 0)
    ws.Range("Q4").Value = Range("I" & RowNum3)
    ws.Range("P1").Value = ws.Name
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Columns("I:R").AutoFit
    Next ws
End Sub