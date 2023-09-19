# VBA-challenge
assign 2



Sub Stock_Data_calculation():

    'loop through all worksheets in the workbook
      For Each ws In Worksheets

    'define variables for worksheet-specific calculations
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TickCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double

    'get the name of the current worksheet
        WorksheetName = ws.Name

    'create column headers for data
        ws.Cells(1, 9).Value = "<Ticker>"
        ws.Cells(1, 10).Value = "Yearly_Change($)"
        ws.Cells(1, 11).Value = "Percent_Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "<Ticker>"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest Increase %"
        ws.Cells(3, 15).Value = "Greatest Decrease %"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
     'initialize counters and trackers
        TickerCount = 2
        j = 2
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'loop through rows of data
        For i = 2 To LastRowA

        'check if the ticker symbol changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
        'ticker symbol in the Ticker column
            ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value

        'calculation yearly_change
            ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

        'apply color formatting based on positive or negative yearly change
            If ws.Cells(TickerCount, 10).Value < 0 Then
            ws.Cells(TickerCount, 10).Interior.ColorIndex = 3 'Red
                
            Else
            ws.Cells(TickerCount, 10).Interior.ColorIndex = 4 'Green
            End If

        'Calculation percent_change
            If ws.Cells(j, 3).Value <> 0 Then
            PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
            ws.Cells(TickerCount, 11).Value = Format(PerChange, "Percent")
            
            Else
            ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
            End If

        'Calculate total stock volume
            ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

        'Move to the next ticker symbol
            TickerCount = TickerCount + 1

                'the start row of the next ticker block
                j = i + 1
            End If
            Next i

        'Find the last cell in the Ticker column
            LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'variables for summary calculations
            GreatVolume = ws.Cells(2, 12).Value
            GreatIncrease = ws.Cells(2, 11).Value
            GreatDecrease = ws.Cells(2, 11).Value

        'Loop through ticker symbols for summary calculations
            For i = 2 To LastRowI

        'greatest total volume
            If ws.Cells(i, 12).Value > GreatVolume Then
            GreatVolume = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If

        'greatest increase
            If ws.Cells(i, 11).Value > GreatIncrease Then
                GreatIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If

            'greatest decrease
            If ws.Cells(i, 11).Value < GreatDecrease Then
                GreatDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If

        Next i

        'return the stock value summary
            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")

        'Auto adjust and update column width
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
        Next ws

    End Sub

