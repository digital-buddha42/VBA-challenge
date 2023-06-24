Sub Stonks():
    
    'Set an intial variable for holding ticker name
    Dim ticker As String
    
    'Create a variable to find the last row in column 1, ticker
    Dim lastrow As Long

    'Create summary table row counter
    Dim summary_table_row As Long
    
    'Create variables to find year open and close and calculate change
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    'Set an intial variable to count the volume totals per ticker
    Dim ticker_vol As Double
    
    'Define current row value placeholders
    Dim greatest_percent_increase As Double
    Dim greatest_percent_decrease As Double
    Dim greatest_total_vol As Double

    'Define max value variables
    Dim max_decrease As Double
    Dim max_volume As Double
    
    Dim ws As Worksheet
    
     For Each ws In ThisWorkbook.Sheets
     
        
        'Intitialize values
        summary_table_row = 2
        ticker_vol = 0
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set current yearly change value as the first value in first row
        greatest_percent_increase = 0
        greatest_per_increase_ticker = ""
        
        greatest_percent_decrease = 9999999
        greatest_per_decrease_ticker = ""
        
        greatest_total_vol = 0
        greatest_total_vol_ticker = ""
        
        year_open = ws.Cells(2, 3).Value
        
        'Summary labels
        ws.Cells(1, 9).Value = ("Ticker")
        ws.Cells(1, 10).Value = ("Yearly Change")
        ws.Cells(1, 11).Value = ("Percent Change")
        ws.Cells(1, 12).Value = ("Total Volume")
        
        ws.Cells(2, 15).Value = ("Greatest % Increase")
        ws.Cells(3, 15).Value = ("Greatest % Decrease")
        ws.Cells(4, 15).Value = ("Greatest Total Volume")
        
        ws.Cells(1, 16).Value = ("Ticker")
        ws.Cells(1, 17).Value = ("Value")
        
        For i = 2 To lastrow
            
            'ticker volume adds the vol number to its previous total
            ticker_vol = ticker_vol + ws.Cells(i, 7).Value
            
            'This checks the current cell against the next cell to see if the ticker changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            'Stick the ending ticker value in the first avilable cell under "ticker"
                ws.Cells(summary_table_row, 9).Value = ws.Cells(i, 1).Value
                
                'Store the year_close variable
                year_close = ws.Cells(i, 6).Value
                
                'Define yearly_change with formula
                yearly_change = year_close - year_open
                
                'Define percent_change with formula
                percent_change = yearly_change / year_open
                
                'Pull the next tickers year_open value before looping through again
                year_open = ws.Cells(i + 1, 3).Value
                
                'Output yearly_change to desired cell per summary_table_row counter
                ws.Cells(summary_table_row, 10).Value = yearly_change
                
                If yearly_change > 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                    
                ElseIf yearly_change < 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                
                End If
                
                
                'Output percent_change to desired cell per summary_table_row counter
                ws.Cells(summary_table_row, 11).Value = percent_change
                
                'Convert percent_change to percent format
                ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
                
                'Output ticker volume to desired cell per summary_table_row counter
                ws.Cells(summary_table_row, 12).Value = ticker_vol
            
                'This checks the current cell against the next cell and resets the max_increase to the higher of the two
                If greatest_percent_increase < percent_change Then
                
                    greatest_percent_increase = percent_change
                    
                    'Once greatest_percent_increase is found, set ticker value to desired cell in max_table.
                    greatest_per_increase_ticker = ws.Cells(summary_table_row, 9).Value
                
                End If
                
                If greatest_percent_decrease > percent_change Then
                
                    greatest_percent_decrease = percent_change
                    
                    'Once greatest_percent_increase is found, set ticker value to desired cell in max_table.
                    greatest_per_decrease_ticker = ws.Cells(summary_table_row, 9).Value
                
                End If
                
                If ticker_vol > greatest_total_vol Then
                
                    greatest_total_vol = ticker_vol
                    
                    'Once greatest_percent_increase is found, set ticker value to desired cell in max_table.
                    greatest_total_vol_ticker = ws.Cells(summary_table_row, 9).Value
                
                End If
                
                'Reset yearly_change to 0 for new ticker
                yearly_change = 0
                
                'Add 1 to summary table row count
                summary_table_row = summary_table_row + 1
                
                'Reset ticker_vol to 0
                ticker_vol = 0
            
            End If
            
        Next i
        
        ws.Cells(4, 16).Value = greatest_total_vol_ticker
        ws.Cells(4, 17).Value = greatest_total_vol
        
        ws.Cells(2, 16).Value = greatest_per_increase_ticker
        ws.Cells(2, 17).Value = greatest_percent_increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = greatest_per_decrease_ticker
        ws.Cells(3, 17).Value = greatest_percent_decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Columns("A:Q").AutoFit
        
    Next ws
    
End Sub


