Sub final()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ProcessWorksheet ws
    Next ws
End Sub

Sub ProcessWorksheet(ws As Worksheet)
    ' Save the current state of Excel settings
    Dim screenUpdateState As Boolean, statusBarState As Boolean, calcState As XlCalculation, eventsState As Boolean, displayPageBreakState As Boolean
    screenUpdateState = Application.ScreenUpdating
    statusBarState = Application.DisplayStatusBar
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    displayPageBreakState = ws.DisplayPageBreaks
    
    ' Disable screen updating and calculation
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ws.DisplayPageBreaks = False
    
    ' Finding the ticker symbol
    Dim Tickersymbols As Range
    Set Tickersymbols = ws.Range("I1:I" & ws.Rows.Count)
    Tickersymbols.ClearContents
    ws.Range("A1:A" & ws.Rows.Count).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Tickersymbols, Unique:=True
    
    ' Set Dates to numeric values
    With ws.Range("B:B")
        .NumberFormat = "0"
        .Value = .Value
    End With
    
    ' Get the beginning of the year and end of year date and percent change
    Dim Dates As Range
    Dim firstofyear As Long, lastofyear As Long
    Set Dates = ws.Range("B:B")
    firstofyear = WorksheetFunction.Min(Dates)
    lastofyear = WorksheetFunction.Max(Dates)
    
    ' Read data into arrays for net change
    Dim SymbolcolumnA() As Variant, fulldate() As Variant, OpenPrice() As Variant, ClosePrice() As Variant, Volume() As Variant, FinalData() As Variant, Netchange() As Variant
    Dim NetChangeData() As Double
    Dim TickerColumnA As Long, TickerColumnI As Long
    TickerColumnA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    TickerColumnI = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    SymbolcolumnA = ws.Range("A1:A" & TickerColumnA).Value
    fulldate = ws.Range("B1:B" & TickerColumnA).Value
    OpenPrice = ws.Range("C1:C" & TickerColumnA).Value
    ClosePrice = ws.Range("F1:F" & TickerColumnA).Value
    Volume = ws.Range("G1:G" & TickerColumnA).Value
    FinalData = ws.Range("I1:I" & TickerColumnI).Value
    Netchange = ws.Range("J1:J" & TickerColumnI).Value

    ' Perform calculations for net change
    Dim Firstnetchangeloop As Long, Secondnetchangeloop As Long, Color As Long
    Dim OpenPrices As Double, ClosePrices As Double
    For Firstnetchangeloop = 1 To TickerColumnI
        For Secondnetchangeloop = 1 To TickerColumnA
            If FinalData(Firstnetchangeloop, 1) = SymbolcolumnA(Secondnetchangeloop, 1) And fulldate(Secondnetchangeloop, 1) = firstofyear Then
                OpenPrices = OpenPrice(Secondnetchangeloop, 1)
            End If
            If FinalData(Firstnetchangeloop, 1) = SymbolcolumnA(Secondnetchangeloop, 1) And fulldate(Secondnetchangeloop, 1) = lastofyear Then
                ClosePrices = ClosePrice(Secondnetchangeloop, 1)
                Exit For
            End If
        Next Secondnetchangeloop
        ws.Cells(Firstnetchangeloop, 10) = ClosePrices - OpenPrices
    Next Firstnetchangeloop
        
    For Firstnetchangeloop = 1 To TickerColumnI
        If ws.Cells(Firstnetchangeloop, 10) > 0 Then
            ws.Cells(Firstnetchangeloop, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(Firstnetchangeloop, 10) < 0 Then
            ws.Cells(Firstnetchangeloop, 10).Interior.ColorIndex = 3
        End If
    Next Firstnetchangeloop
    
    ' Perform calculations for percent change
    For Firstnetchangeloop = 2 To TickerColumnI
        For Secondnetchangeloop = 2 To TickerColumnA
            If FinalData(Firstnetchangeloop, 1) = SymbolcolumnA(Secondnetchangeloop, 1) And fulldate(Secondnetchangeloop, 1) = firstofyear Then
                OpenPrices = OpenPrice(Secondnetchangeloop, 1)
                Exit For
            End If
        Next Secondnetchangeloop
        If OpenPrices = 0 Then
            ws.Cells(Firstnetchangeloop, 11) = 0
            Else
            ws.Cells(Firstnetchangeloop, 11) = Netchange(Firstnetchangeloop, 1) / OpenPrices
        End If
    Next Firstnetchangeloop
    
    ws.Range("K2").Resize(TickerColumnI).NumberFormat = "0.00%"
    
    ' Get total volume
    Dim totalvolume As Long
    For totalvolume = 2 To TickerColumnI
        ws.Cells(totalvolume, 12) = WorksheetFunction.SumIf(ws.Range("A:A"), ws.Cells(totalvolume, 9), ws.Range("G:G"))
    Next totalvolume
    
    ws.Range("L2").Resize(TickerColumnI).NumberFormat = "0,000"
    
    'Bonus
    Dim totalvolumedata As Variant
    Dim maxdata As Double, mindata As Double
    Dim maxvolume As Long
    totalvolumedata = ws.Range("L1:L" & TickerColumnI).Value
    maxdata = WorksheetFunction.Max(ws.Range("J:J"))
    mindata = WorksheetFunction.Min(ws.Range("J:J"))
    maxvolumedata = WorksheetFunction.Max(ws.Range("L:L"))
    
    For Firstnetchangeloop = 1 To TickerColumnI
        If Netchange(Firstnetchangeloop, 1) = maxdata Then
        ws.Cells(2, 16) = FinalData(Firstnetchangeloop, 1)
        Exit For
        End If
    Next Firstnetchangeloop
    
    For Firstnetchangeloop = 1 To TickerColumnI
        If Netchange(Firstnetchangeloop, 1) = mindata Then
        ws.Cells(3, 16) = FinalData(Firstnetchangeloop, 1)
        Exit For
        End If
    Next Firstnetchangeloop
    
    For Firstnetchangeloop = 1 To TickerColumnI
        If totalvolumedata(Firstnetchangeloop, 1) = maxvolumedata Then
        ws.Cells(4, 16) = FinalData(Firstnetchangeloop, 1)
        Exit For
        End If
    Next Firstnetchangeloop
    
    For Firstnetchangeloop = 2 To 4
        For Secondnetchangeloop = 1 To TickerColumnA
            If Cells(Firstnetchangeloop, 16) = SymbolcolumnA(Secondnetchangeloop, 1) And fulldate(Secondnetchangeloop, 1) = firstofyear Then
            ClosePrices = ClosePrice(Secondnetchangeloop, 1)
            Exit For
            End If
        Next Secondnetchangeloop
        ws.Cells(Firstnetchangeloop, 17) = ClosePrices
    Next Firstnetchangeloop

    
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
  
    ' Auto-fit columns
    ws.Cells.EntireColumn.AutoFit
    
    ' Restore Excel settings to the original state
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
    ws.DisplayPageBreaks = displayPageBreakState
End Sub





