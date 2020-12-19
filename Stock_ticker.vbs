Sub stock_ticker()
'Set up a variable for specifying the column of interest
'Need new columns Ticker, Change, percent change, total stock volume
'Create a script that will loop through all the stocks for one year and output the following information.
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.




'You should also have conditional formatting that will highlight positive change in green and negative change in red.

    Dim column As Integer
    Dim newrow As Integer
    Dim opencol As Integer
    Dim closecol As Integer
    Dim stockcol As Integer
    Dim totaltickercol As Integer
    Dim totalchangecol As Integer
    Dim percentchangecol As Integer
    Dim totalvolcol As Integer
    Dim Total As Double
    Dim totalStock As Double
    Dim NumRows As Long
    Dim curopen As Double
    Dim curclose As Double
    'Dim sheet1 As Worksheet
    Dim greatperinc As Integer
    Dim greatperdec As Integer
    Dim greatvol As Integer
    Dim greattitlecol As Integer
    Dim greattickercol As Integer
    Dim greatvaluecol As Integer
    Dim percentcalc As Double
    
    'Initialize Variables
    totalStock = 0
    column = 1
    opencol = 3
    closecol = 6
    stockcol = 7
    totaltickercol = 9
    totalchangecol = 10
    percentchangecol = 11
    totalvolcol = 12
    greattitlecol = 14
    greattickercol = 15
    greatvaluecol = 16
    
    For Each ws In Worksheets
        'Initialize worksheet variables
        Total = 0
        curopen = 0
        curclose = 0
        greatperinc = 0
        greatperdec = 0
        greatvol = 0
        newrow = 2
        NumRows = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Put labels in
        ws.Cells(newrow - 1, totaltickercol) = "Ticker"
        ws.Cells(newrow - 1, totalchangecol) = "Yearly Change"
        ws.Cells(newrow - 1, percentchangecol) = "Percent Change"
        ws.Cells(newrow - 1, totalvolcol) = "Total Stock Volume"
    
        'First row with have the open for first list
        curopen = ws.Cells(newrow, opencol)
        
        'Loop through the rows of the column
        For i = 2 To NumRows
            'Searchs for when the volue of the next cell is different from that of current cell
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                'Message Box the value
                'MsgBox (Cells(i, column).Value)
                'Add to the total
                    Total = Total + ws.Cells(i, stockcol).Value
                'Store the close price
                    curclose = ws.Cells(i, closecol)
                'Put the ticker in
                    ws.Cells(newrow, totaltickercol) = ws.Cells(i, column)
                'Calculate difference between open and close
                    ws.Cells(newrow, totalchangecol) = curclose - curopen
                    'If positive green, negative red
                    If ws.Cells(newrow, totalchangecol) >= 0 Then
                        ws.Cells(newrow, totalchangecol).Interior.ColorIndex = 4
                    Else
                        ws.Cells(newrow, totalchangecol).Interior.ColorIndex = 3
                    End If
            
                'Calculate percent change
                    If curopen = 0 Then
                        ws.Cells(newrow, percentchangecol) = "0%"
                    Else
                        percentcalc = (curclose - curopen) / curopen
                        ws.Cells(newrow, percentchangecol) = percentcalc
                        'ws.Cells(newrow, percentchangecol).Style = "Percent"
                    End If
                'Save the total to the summary columns
                    ws.Cells(newrow, totaltickercol) = ws.Cells(i, column)
                    ws.Cells(newrow, totalvolcol) = Total
                    
                'check for greatest changes
                    If greatperinc = 0 Then
                        greatperinc = newrow
                    ElseIf Cells(newrow, percentchangecol) > Cells(greatperinc, percentchangecol) Then
                        greatperinc = newrow
                    End If
                    If greatperdec = 0 Then
                        greatperdec = newrow
                    ElseIf Cells(newrow, percentchangecol) < Cells(greatperdec, percentchangecol) Then
                        greatperdec = newrow
                    End If
                    If greatvol = 0 Then
                        greatvol = newrow
                    ElseIf Cells(newrow, totalvolcol) > Cells(greatvol, totalvolcol) Then
                        greatvol = newrow
                    End If
                    
                'Increment row
                    newrow = newrow + 1
                'Reset total to zero
                    Total = 0
                'Reset curopen
                    curopen = ws.Cells(i + 1, opencol)
            Else
                'Add to the total
                Total = Total + ws.Cells(i, stockcol).Value
        
            End If
            
        Next i
        
        'Write out the greatest values
         ws.Cells(2, greattitlecol) = "Greatest % Increase"
         ws.Cells(3, greattitlecol) = "Greatest % Decrease"
         ws.Cells(4, greattitlecol) = "Greatest Total Volume"
         ws.Cells(1, greattickercol) = "Ticker"
         ws.Cells(1, greatvaluecol) = "Value"
         ws.Cells(2, greatvaluecol) = Cells(greatperinc, percentchangecol) * 100
         ws.Cells(2, greattickercol) = Cells(greatperinc, totaltickercol)
         ws.Cells(3, greatvaluecol) = Cells(greatperdec, percentchangecol) * 100
         ws.Cells(3, greattickercol) = Cells(greatperdec, totaltickercol)
         ws.Cells(4, greatvaluecol) = Cells(greatvol, totalvolcol)
         ws.Cells(4, greattickercol) = Cells(greatvol, totaltickercol)
        
    Next ws


End Sub


