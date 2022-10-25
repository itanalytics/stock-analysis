Sub stockstats()

'Create variables
Dim ws As Worksheet
Dim i As Long
Dim stockcount As Integer
Dim currentstock As String
Dim openprice, closeprice, volume As Double
Dim top_inc, top_dec, top_vol As Double
Dim inc_name, dec_name, vol_name As String

'Loop through each worksheet in the workbook
For Each ws In Worksheets
    'Activate the currect worksheet
    ws.Activate
    
    'Create and format column titles for statistics
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("I1").Font.Bold = "True"
    Range("J1").Font.Bold = "True"
    Range("K1").Font.Bold = "True"
    Range("L1").Font.Bold = "True"
    Range("P1").Font.Bold = "True"
    Range("Q1").Font.Bold = "True"
    
    'initialize values for variables
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    stockcount = 1
    openprice = 0
    closeprice = 0
    volume = 0
    
    'Loop through each row of stock data
    For i = 2 To lastrow
    
        'Check for first entry of a stock and store values in variables
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            currentstock = Cells(i, 1).Value
            openprice = Cells(i, 3).Value
            volume = volume + Cells(i, 7).Value
            stockcount = stockcount + 1
            
        'Add volume for all subsequent entries
        ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            volume = volume + Cells(i, 7).Value
        
        'At final entry, calculate statistics and print to summary list
        Else
            closeprice = Cells(i, 6).Value
            volume = volume + Cells(i, 7).Value
            Cells(stockcount, 9).Value = currentstock
            Cells(stockcount, 10).Value = closeprice - openprice
            
                'Color Format for Positive or Negative Change
                If Cells(stockcount, 10).Value >= 0 Then
                    Cells(stockcount, 10).Interior.ColorIndex = 4
                Else
                    Cells(stockcount, 10).Interior.ColorIndex = 3
                End If
            
            Cells(stockcount, 11).Value = (closeprice - openprice) / openprice
            Cells(stockcount, 11).NumberFormat = "0.00%"
            Cells(stockcount, 12).Value = volume
            
            'Reset variables for next stock
            openprice = 0
            closeprice = 0
            volume = 0
            
        End If
        
    Next i
    
    'Initialize variables for Top Values List
    top_inc = 0
    top_dec = 0
    top_vol = 0
    
    'Loop through summary statistics table
    For j = 2 To stockcount
        
        'Compare entries against current Top Value, if greater, store new Top Value
        If Cells(j, 11).Value > top_inc Then
            top_inc = Cells(j, 11).Value
            inc_name = Cells(j, 9).Value
        End If
        If Cells(j, 11).Value < top_dec Then
            top_dec = Cells(j, 11).Value
            dec_name = Cells(j, 9).Value
        End If
        If Cells(j, 12).Value > top_vol Then
            top_vol = Cells(j, 12).Value
            vol_name = Cells(j, 9).Value
        End If

    Next j
    
    'Print Top Values to table and format
    Range("O2").Value = "Greatest % Increase"
    Range("P2").Value = inc_name
    Range("Q2").Value = top_inc
    Range("Q2").NumberFormat = "0.00%"
    
    Range("O3").Value = "Greatest % Decrease"
    Range("P3").Value = dec_name
    Range("Q3").Value = top_dec
    Range("Q3").NumberFormat = "0.00%"
    
    Range("O4").Value = "Greatest Total Volume"
    Range("P4").Value = vol_name
    Range("Q4").Value = top_vol


Next ws


End Sub
