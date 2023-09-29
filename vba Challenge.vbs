Attribute VB_Name = "Module1"
Sub StockCalcu()
    For Each ws In Worksheets
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'declare ticker var
    Dim Ticker As String
    
'Declare yearly change
    Dim Year As Double
    Year = 0
    Dim OYear As Double
    OYear = 0
    Dim yearStored As Boolean
    Dim CYear As Double
    CYear = 0
    
'Declare Percentage Change
    Dim Percent As Double
    Percent = 0
    
'delare Total Volume
    Dim TotalVolume As Double
    TotalVolume = 0
    
'which row do we start at
    Dim SummaryTicker As Integer
    SummaryTicker = 2
    
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
'Just in case the macro is called twice
    If ws.Cells(1, 9).Value <> "Ticker" Then
        ws.Cells(1, 9).Value = "Ticker"
        ws.Range("J1").ColumnWidth = 13
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Range("K1").ColumnWidth = 14
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Range("L1").ColumnWidth = 18
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Range("O1").ColumnWidth = 21
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Range("Q1").ColumnWidth = 25
        ws.Cells(1, 17).Value = "Value"
    End If
    
'Loop all
    For i = 2 To LastRow
        If yearStored = False Then
            OYear = ws.Cells(i, 3).Value
            yearStored = True
        End If
             
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            CYear = ws.Cells(i, 6).Value
            Year = CYear - OYear
            Percent = Year / OYear
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            ws.Range("I" & SummaryTicker).Value = Ticker
            ws.Range("J" & SummaryTicker).Value = Year
            ws.Range("K" & SummaryTicker).Value = Percent
            ws.Range("L" & SummaryTicker).Value = TotalVolume
            SummaryTicker = SummaryTicker + 1
            TotalVolume = 0
            yearStored = False
        Else
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        End If
        
    Next i
    
    
    ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Columns("K"))
    ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Columns("K"))
    ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Columns("L"))
    
    Dim High As Double
    Dim Low As Double
    Dim Vol As Double

    High = ws.Cells(2, 17).Value
    Low = ws.Cells(3, 17).Value
    Vol = ws.Cells(4, 17).Value
    
    Dim TickH As String
    Dim TickL As String
    Dim TickV As String
    
    
    LastPercent = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    For c = 2 To LastPercent
    If ws.Cells(c, 10).Value > 0 Then
            ws.Cells(c, 10).Interior.ColorIndex = 4
        ElseIf IsEmpty(ws.Cells(c, 10)) Then
            GoTo ColorCheck
        Else
            ws.Cells(c, 10).Interior.ColorIndex = 3
        End If
ColorCheck:
    Next c
    
    For q = 2 To LastPercent
        If ws.Cells(q, 11).Value = High Then
            TickH = ws.Cells(q, 9).Value
            ws.Cells(2, 16).Value = TickH
            'MsgBox ("Got em")
            GoTo HighTick
        End If
    Next q
HighTick:

    For w = 2 To LastPercent
        If ws.Cells(w, 11).Value = Low Then
            TickL = ws.Cells(w, 9).Value
            ws.Cells(3, 16).Value = TickL
            'MsgBox ("Got em")
            GoTo LowTick
        End If
    Next w
LowTick:

    For e = 2 To LastPercent
        If ws.Cells(e, 12).Value = Vol Then
            TickV = ws.Cells(e, 9).Value
            ws.Cells(4, 16).Value = TickV
            'MsgBox ("Got em")
            GoTo VolTick
        End If
    Next e
VolTick:
    
    

      
Next ws

MsgBox ("Done!")

End Sub
