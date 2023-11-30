Sub StockAnalyser():

    'Loop through all rows in a worksheet, in all worksheets
    
    'Get ticker symbol and create a new Ticker column in "I"
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets

        'Create variable to hold Ticker
        Dim Ticker As String

        'Create variable for holding the total volume for each ticker
        Dim Ticker_Volume As Double
        Ticker_Volume = 0

        'Create variable for price
        Dim Ticker_Price As Double

        'Track location of unique tickers in the summary
        Dim Ticker_Summary_Row As Integer
        Ticker_Summary_Row = 2

        'Create variable for year open
        Dim Year_Open As Double

        'Create variable for year close
        Dim Year_Close As Double

        'Create variable for yearly change
        Dim Year_Change As Double

        'Determine last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Create Ticker column
        ws.Cells(1, 9).Value = "Ticker"

        'Create Yearly Change Column
        ws.Cells(1, 10).Value = "Yearly Change"

        'Create Percent Change Column
        ws.Cells(1, 11).Value = "Percent Change"

        'Create Total Stock Volume Column
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Dim start As Long
        start = 2

        'Loop through all ticker transactions
        For i = 2 To LastRow
        
        'Year Open
        Year_Open = ws.Cells(start, 3).Value

            'Check if still within the same stock ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set Ticker Name
                Ticker = ws.Cells(i, 1).Value

                'Add to Ticker Volume
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

                'Print the Ticker in the summary
                ws.Range("I" & Ticker_Summary_Row).Value = Ticker

                'Print Ticker Volume in the summary
                ws.Range("L" & Ticker_Summary_Row).Value = Ticker_Volume

                'Add another summary row
                Ticker_Summary_Row = Ticker_Summary_Row + 1

                'Reset Ticker Volume
                Ticker_Volume = 0

                'Year Close = column f
                Year_Close = ws.Cells(i, 6).Value
                'Test Year Close
                'ws.Range("O" & Ticker_Summary_Row - 1).Value = Year_Close
                
                'Year Change
                Year_Change = Year_Close - Year_Open

                'Print the Year Change in summary column "J"
                ws.Range("J" & Ticker_Summary_Row - 1).Value = (Year_Change)
                
                'Percentage Change column "K"
                ws.Range("K" & Ticker_Summary_Row - 1).Value = WorksheetFunction.Round(((Year_Change) / Year_Close) * 100, 2) & "%"
                
                'Colour code Yearly Change cells
                If Year_Change >= 0 Then
                    ws.Range("J" & Ticker_Summary_Row - 1).Interior.ColorIndex = 4
                
                Else
                    ws.Range("J" & Ticker_Summary_Row - 1).Interior.ColorIndex = 3
                    
                End If
                
                
                'start of the next stock ticker
                start = i + 1
                
                'Update Year Open ... before moving to the next row
                'Year_Open = Cells(i, 3).Value
                'Test Year Open
                'ws.Range("N" & Ticker_Summary_Row - 1).Value = Year_Open

            'If the ticker is the same in the next row as the current row...
            Else

                'Add to Ticker Volume
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

                'Move Year Close to the next row

            End If



        Next i
    
    'ouput the yearly change in value from the opening price to the closing price. Colour the cells in green or red as appropriate.
    'output the total volume of stock
    'Bonus: Output the Greatest % increase, Greatest % decrease and greatest total volume.
    
    'Run on all worksheets (years) from the single sub
    
    'green = Cells(, ).Interior.ColorIndex = 4
    'red = Cells(, ).Interior.ColorIndex = 3
    Next ws

End Sub
