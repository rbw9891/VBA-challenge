Sub ticktock():
    
    'Variable for worksheet object
    Dim current As Worksheet
    
    'for each loop to go through worksheets in workbook one at a time
    For Each current In Worksheets
        
        'calling the worksheet method
        current.Activate
        
        'add Ticker heading to column I
        current.Cells(1, 9).Value = "Ticker"
        'add Yearly Change heading to column J
        current.Cells(1, 10).Value = "Yearly Change"
        'add Percent Change heading to column K
        current.Cells(1, 11).Value = "Percent Change"
        'add Total Stock volumn heading to column L
        current.Cells(1, 12).Value = "Total Stock Volume"
        'add Ticker heading to column P
        current.Cells(1, 16).Value = "Ticker"
        'add Value heading to column Q
        current.Cells(1, 17).Value = "Value"
        'add greatest % increase heading
        current.Cells(2, 15).Value = "Greatest % Increase"
        'add greatest % decrease heading
        current.Cells(3, 15).Value = "Greatest % Decrease"
        'add greatest total volume heading
        current.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Variable for last row and define
        Dim lastrow As Long
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (Str(lastrow))
        
        'Variable for ticker name
        Dim tickernm As String
        
        'Variable for location of tickernm in summary table
        Dim sum_table_row As Integer
        sum_table_row = 2
        
        'Variable for row counter
        Dim rowcounter As Integer
        rowcounter = 1
        
        'Variable for opening price
        Dim openingprice As Double
        
        'Variable for closing price
        Dim closeprice As Double
        
        'Variable for yearly change
        Dim yearlychange As Variant
        
        'Variable for yearly change column
        Dim change_row As Integer
        change_row = 2
        
        'Variable for percent change
        Dim percentchange As Variant
        
        'Variable for percent change column
        Dim percent_row As Integer
        percent_row = 2
        
        'Variable for volume by ticker
        Dim ticker_total As Double
        ticker_total = 0
        
        'Variable for volume by ticker column
        Dim ticker_volumn_row As Integer
        ticker_volumn_row = 2
    
       
        
            'For loop on column 1
            For i = 2 To lastrow
            
                'Conditional to find change in cell value
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                    'set tickernm
                    tickernm = Cells(i, 1).Value
                
                    'generate tickernm in column I
                    Range("I" & sum_table_row).Value = tickernm
                
                    'add one to the sum_table_row
                    sum_table_row = sum_table_row + 1
                    
                    'generate closing price
                    closeprice = Cells(i, 6).Value
                    
                    'generate opening price
                    openingprice = Cells((i) - (rowcounter - 1), 3).Value
                    
                    'rowcounter back to zero
                    rowcounter = 0
                    
                    'set yearly change
                    yearlychange = closeprice - openingprice
                    
                    'generate yearly change
                    Range("J" & change_row).Value = yearlychange
                        
                        'Conditional to set fill color dependant on positive or negative yearly change
                        If yearlychange > 0 Then
                            
                            Range("J" & change_row).Interior.ColorIndex = 4
                            
                        Else
                        
                            Range("J" & change_row).Interior.ColorIndex = 3
                            
                        End If
                    
                    'add one to the change_row
                    change_row = change_row + 1
                    
                    'set percentage change
                    percentchange = ((closeprice - openingprice) / openingprice)
                    
                    'generate percent change
                    Range("K" & percent_row).Value = percentchange
                    
                    'add one to percent row
                    percent_row = percent_row + 1
                    
                    'add to the ticker_total
                    ticker_total = ticker_total + Cells(i, 7).Value
                    
                    'generate volumn total
                    Range("L" & ticker_volumn_row).Value = ticker_total
                    
                    'add one to the ticker volumn row
                    ticker_volumn_row = ticker_volumn_row + 1
                    
                    'reset ticker_total to zero
                    ticker_total = 0
                    
                Else
                
                    'add to ticker_total
                    ticker_total = ticker_total + Cells(i, 7).Value
                    
                End If
                
                'add one to row counter
                rowcounter = rowcounter + 1
                
            Next i
            
        'Variable and definition for last row of newly generated summary table
        Dim lastrow_new As Long
        lastrow_new = Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox (Str(lastrow))
        
        'variable for ticker name to go into greatest table
        Dim tickernm_new As String
        'variable for greatest percent increase
        Dim percent_max As Double
        'variable for greatest percent decrease
        Dim percent_min As Double
        'variable for greatest total volume
        Dim volume_max As Double
        'variable for range of values in column K
        Dim range_perc As Range
        'variable for range of values in column L
        Dim range_volume As Range
        
        'setting the range to be find max /min metrics on percent change column
        Set range_perc = Range("K2:K" & Rows.Count)
        'setting range to find total metric on stock volumn column
        Set range_volume = Range("L2:L" & Rows.Count)
        'generate percent_max
        percent_max = Application.WorksheetFunction.Max(range_perc)
        'generate percent_min
        percent_min = Application.WorksheetFunction.Min(range_perc)
        'generate volume_max
        volume_max = Application.WorksheetFunction.Max(range_volume)
            
           
            'for loop over newly created ticker column I
            For j = 2 To lastrow_new
                
                'Conditional to find greatest % increase in column K and generate value for greatest table
                If Cells(j, 11).Value = percent_max Then
                    
                    tickernm_new = Cells(j, 9).Value
                    
                    Range("P2").Value = tickernm_new
                    
                    Range("Q2").Value = percent_max
                    
                'Conditional to find greatest % decrease in column K and generate value for greatest table
                ElseIf Cells(j, 11).Value = percent_min Then
                    
                    tickernm_new = Cells(j, 9).Value
                    
                    Range("P3").Value = tickernm_new
                    
                    Range("Q3").Value = percent_min
                    
                'Conditional to find greatest total volume in column L and generate value for greatest table
                ElseIf Cells(j, 12).Value = volume_max Then
                    
                    tickernm_new = Cells(j, 9).Value
                    
                    Range("P4").Value = tickernm_new
                    
                    Range("Q4").Value = volume_max
                    
                End If
            
            Next j
            
        'format percentage change values and autofit column widths of new stats for readability
        Range("K2:K" & lastrow_new).NumberFormat = "0.00%"
        Range("Q2:Q3").NumberFormat = "0.00%"
        Columns("I:Q").AutoFit
        
    Next current
        
    
End Sub
