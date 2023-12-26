Attribute VB_Name = "Module1"
Sub stock_analysis()

'Declare data types
Dim TickerSymbol As String
Dim TickerTotal As Double
Dim TickerOpen As Double
Dim TickerClose As Double
Dim YearlyChange As Double
Dim SummaryTableRow As Integer
Dim Increase As Variant
Dim PercentIncrease As Variant

'Count number of worksheets
SheetCount = ThisWorkbook.Worksheets.Count

'Loop through all worksheets
For j = 1 To SheetCount

'Set worksheet
Sheets(j).Activate

'Declare first row in the summary table
 SummaryTableRow = 2

'Set the total to 0
TickerTotal = 0

'Add the headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest total volume"

'Determine last row of data
 LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Set TickerTotal to 0 and pull out opening price on day one for first company
TickerTotal = 0
TickerOpen = Cells(2, 3).Value

'Loop through all data on sheet 1
For i = 2 To LastRow

    'If we're still in the same ticker
    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    
        'add to the TickerTotal and move on
        TickerTotal = TickerTotal + Cells(i, 7).Value
    
    'At last line, finalize total and grab closing price
    Else
        TickerSymbol = Cells(i, 1).Value
        TickerTotal = TickerTotal + Cells(i, 7).Value
        TickerClose = Cells(i, 6).Value
        
        'Determine change in price
        TickerChange = TickerClose - TickerOpen
        
        'print info in summary table
        Cells(SummaryTableRow, 9).Value = TickerSymbol
        Cells(SummaryTableRow, 12).Value = TickerTotal
        Cells(SummaryTableRow, 10).Value = TickerChange
        Cells(SummaryTableRow, 11).Value = FormatPercent(TickerChange / TickerOpen)
        
        'apply conditional formatting
        If TickerChange > 0 Then
            Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
            
        Else
            Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
            
        End If
        
        'Go to next line of summary table
        SummaryTableRow = SummaryTableRow + 1
        
        'Reset tickertotal
        TickerTotal = 0
        
        'Get next opening price
        TickerOpen = Cells(i + 1, 3).Value
    
    End If

Next i

'Find the greatest percent increase, decrease, and max volume
MaxIncrease = WorksheetFunction.Max(ActiveSheet.Columns("k"))
MaxDecrease = WorksheetFunction.Min(ActiveSheet.Columns("k"))
MaxVolume = WorksheetFunction.Max(ActiveSheet.Columns("l"))

'Find the ticker symbols for max increase, decrease, and volume
MaxIncRow = WorksheetFunction.Match(MaxIncrease, ActiveSheet.Columns("k"), 0)
MaxIncSymbol = Cells(MaxIncRow, 9).Value
MaxDecRow = WorksheetFunction.Match(MaxDecrease, ActiveSheet.Columns("k"), 0)
MaxDecSymbol = Cells(MaxDecRow, 9).Value
MaxVolRow = WorksheetFunction.Match(MaxVolume, ActiveSheet.Columns("l"), 0)
MaxVolSymbol = Cells(MaxVolRow, 9).Value

'Print in new summary table
Cells(2, 17).Value = FormatPercent(MaxIncrease)
Cells(3, 17).Value = FormatPercent(MaxDecrease)
Cells(4, 17).Value = MaxVolume
Cells(2, 16).Value = MaxIncSymbol
Cells(3, 16).Value = MaxDecSymbol
Cells(4, 16).Value = MaxVolSymbol

'Move on to next sheet
Next j

End Sub

