Sub tickeranalysis():

'Repeat code across all worksheets
For Each ws In Worksheets

    'Creating headers for summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    

    
    'Define variable for last row value
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Define variable to track unique ticker names
    Dim ticker_name As String
    

    'Define variable for holding total stock volume - double because long doesn't seem to work for some reason
    Dim ticker_total As Double
    ticker_total = 0
    
    'Define opening price at start of year
    Dim open_startyear As Double
    
    'Define ending price at end of year
    Dim close_endyear As Double
    
    'Keep track of first row for each unique ticker name
    Dim ticker_firstrow As Integer
    
    'Keep track of location of unique ticker brands in final table
    Dim ticker_tracker As Integer
    ticker_tracker = 2
    
    'Keep track of firstrow of unique name for start of year, set initial value to 4
    Dim row_startyear As Integer
    row_startyear = 4
    
    
    
    'Loop through all ticker rows dynamically
    For i = 2 To LastRow
        
        'Check if is not the same ticker brand
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Set the brand name/put in 'clipboard'
            ticker_name = ws.Cells(i, 1).Value
            
            'Add to total volume
            ticker_total = ticker_total + ws.Cells(i, 7)
            
            'Print ticker name to summary table
            ws.Range("I" & ticker_tracker).Value = ticker_name
            
            'Print ticker volume to summary table
            ws.Range("L" & ticker_tracker).Value = ticker_total
            
        
            'Conditional to make sure that calculating yearly change/percent change works for first ticker name
            If row_startyear > i Then
                'Grab opening price at start of the year for this ticker
                open_startyear = ws.Cells(row_startyear - i, 3).Value
            
                'Grab closing price at end of the year for this ticker
                close_endyear = ws.Cells(i, 6).Value
                
                'Define variable to track yearly change -> input into summary table
                Dim yearly_change As Double
                yearly_change = close_endyear - open_startyear
                ws.Range("J" & ticker_tracker).Value = yearly_change
                
                'Conditional formatting for yearly change - green for positive, yellow for no change, red for negative
                If yearly_change > 0 Then
                    ws.Range("J" & ticker_tracker).Interior.ColorIndex = 4
                
                    ElseIf yearly_change = 0 Then
                        ws.Range("J" & ticker_tracker).Interior.ColorIndex = 6
                        
                    Else
                        ws.Range("J" & ticker_tracker).Interior.ColorIndex = 3
                
                End If
                
                'Define variable to track percent change -> input into summary table -> change format to percentage
                Dim percent_change As Double
                percent_change = ((close_endyear - open_startyear) / open_startyear)
                ws.Range("K" & ticker_tracker).Value = percent_change
                ws.Range("K" & ticker_tracker).NumberFormat = "0.00%"
                
            'For every other unique ticker name after the first - same code essentially
            Else
                'subtraction reversed because counter reset
                open_startyear = ws.Cells(i - row_startyear, 3).Value
                
                close_endyear = ws.Cells(i, 6).Value
                
                yearly_change = close_endyear - open_startyear
                ws.Range("J" & ticker_tracker).Value = yearly_change
                
                If yearly_change > 0 Then
                    ws.Range("J" & ticker_tracker).Interior.ColorIndex = 4
                
                    ElseIf yearly_change = 0 Then
                        ws.Range("J" & ticker_tracker).Interior.ColorIndex = 6
                        
                    Else
                        ws.Range("J" & ticker_tracker).Interior.ColorIndex = 3
                
                End If
                
                percent_change = ((close_endyear - open_startyear) / open_startyear)
                ws.Range("K" & ticker_tracker).Value = percent_change
                ws.Range("K" & ticker_tracker).NumberFormat = "0.00%"
                
                
            End If
            
            'Set up for the next row in summary table
            ticker_tracker = ticker_tracker + 1
            
            'Reset ticker volume total for the next unique ticker name
            ticker_total = 0
            
            'Reset counter for startyear variable
            row_startyear = 0
        
        'if same ticker brand = add to total volume and start year
        Else
            ticker_total = ticker_total + ws.Cells(i, 7)
            row_startyear = row_startyear + 1
        
           
        
        End If
        
    Next i
    
    'Create bonus table:
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest total Volume"
    
    'Count how many ticker names in summary table
    summarytablecount = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Keep track of each variable
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    'Go through summary table
    For j = 2 To summarytablecount
    
     'if next value to greater than current count, replace and take note of ticker name
        If ws.Cells(j + 1, 11).Value > greatestIncrease Then
            greatestIncrease = ws.Cells(j + 1, 11).Value
            greatestIncreaseN = ws.Cells(j + 1, 9).Value
        
        'Greatest decrease, keep iterating below 0 to find largest negative number
        ElseIf ws.Cells(j + 1, 11).Value < greatestDecrease Then
            greatestDecrease = ws.Cells(j + 1, 11).Value
            greatestDecreaseN = ws.Cells(j + 1, 9).Value
            
        End If
        
        'Separate If statement for the greatest total volume (different column), same logic
        
        If ws.Cells(j + 1, 12).Value > greatestVolume Then
            greatestVolume = ws.Cells(j + 1, 12).Value
            greatestVolumeN = ws.Cells(j + 1, 9).Value
        
        End If
    Next j
        
       
    
    'print 'greatest' summary table numbers
    ws.Cells(2, 16).Value = greatestIncreaseN
    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(3, 16).Value = greatestDecreaseN
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(4, 16).Value = greatestVolumeN
    ws.Cells(4, 17).Value = greatestVolume
    
    'Format percentage values in 'greatest' summary table
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    'Auto adjust certain columns - total volume + yearly change
    ws.Columns(10).AutoFit
    ws.Columns(11).AutoFit
    ws.Columns(12).AutoFit
    ws.Columns(15).AutoFit
    ws.Columns(17).AutoFit
    
Next ws
End Sub



