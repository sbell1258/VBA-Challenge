Sub CalcStocks()

'  Description :  subroutine takes an alphabetized stock listing by year (sorted in ticker and date order) and generates an annual summary
'   listing the total volume for the stock for the specific year, dollar value of stock price change from beginning of the year to end of 
'   the year, and provides that change in percentage value.  Note: $ changes of 0 or greater reflected with green fill, otherwise negative
'   values reflected with red fill.
'
'   Finally, after processing all entries the annual summary info is used to determine the stock with the highest volume, highest percentage
'   increase and highest percentage decrease.



Dim sheet As Worksheet
Dim volsum, volmax As Single
Dim i, j, k As Single
Dim tick_begin, tick_begin_1st, tick_end As Single
Dim ticker, tick_inc, tick_dec, tick_vol As String
Dim stock_range As Range
Dim perc_inc, perc_dec As Double

Set sheet = ActiveSheet

'  Set Column headings for Calculations by Stock
'  sheet.Cells(row_number, column_number)


sheet.Cells(1, 9).Value = "Ticker"
sheet.Cells(1, 10).Value = "Yearly Change"
sheet.Cells(1, 11).Value = "Percent Change"
sheet.Cells(1, 12).Value = "Total Stock Volume"

'  Determine begin coordinates for all ticker data
'   & establish 1st ticker as default


ticker = sheet.Cells(2, 1).Value
j = 2
k = 2
tick_begin = 0
tick_begin_1st = 2


'  Calc begin/end prices and total volumes for each Stock


While ticker > ""

    volsum = 0

'  Find beginning and end rows for each stock for accumulative processing.  
'   (1) set range that contains a stocks information 
'   (2) process that information
'       (a) Calc total volume for stock
'       (b) Calc year end $ value difference from beginning of year to end of year closing
'       (c) Calc percentage increase/decrease from beginning of year to close of the year


    While sheet.Cells(j, 1).Value = ticker
        If sheet.Cells(j, 3) > 0 And tick_begin = 0 Then
            tick_begin = j
        End If

        j = j + 1
    Wend
    
    Set stock_range = sheet.Range(sheet.Cells(tick_begin_1st, 7), sheet.Cells(j - 1, 7))
    volsum = Application.WorksheetFunction.Subtotal(109, stock_range)

    sheet.Cells(k, 9).Value = ticker
    If tick_begin = 0 Then
        sheet.Cells(k, 10).Value = 0
    Else
        sheet.Cells(k, 10).Value = sheet.Cells(j - 1, 6).Value - sheet.Cells(tick_begin, 3).Value
    End If
    
    If (sheet.Cells(j - 1, 6).Value = 0 And tick_begin = 0) Then
        sheet.Cells(k, 11).Value = 0
    Else
        sheet.Cells(k, 11).Value = (sheet.Cells(j - 1, 6).Value - sheet.Cells(tick_begin, 3).Value) / sheet.Cells(tick_begin, 3).Value
    End If
    
    sheet.Cells(k, 11).NumberFormat = "0.00%"
    sheet.Cells(k, 12).Value = volsum

    If sheet.Cells(k, 10).Value < 0 Then
        sheet.Cells(k, 10).Interior.Color = vbRed
    Else
        sheet.Cells(k, 10).Interior.Color = vbGreen
    End If

'  Reset values for new ticker processing

    tick_begin = 0
    tick_begin_1st = j
    ticker = sheet.Cells(j, 1).Value
    k = k + 1

Wend

'  Determine Tickers with (1) greatest % increase, (2) greatest % decrease and (3) total volume.
'   Set Column headings for Calculations greatest % increase/decrease and total volume


sheet.Cells(1, 16).Value = "Ticker"
sheet.Cells(1, 17).Value = "Value"
sheet.Cells(2, 15).Value = "Greatest % Increase"
sheet.Cells(3, 15).Value = "Greatest % Decrease"
sheet.Cells(4, 15).Value = "Greatest Total Volume"


'  Determine last row of newly created Stock Summary Info
'   and set 1st entry as default for stocks w/ largest increase or decrease and highest volume.


tick_end = sheet.Cells(Rows.Count, 9).End(xlUp).Row


tick_vol = sheet.Cells(2, 9).Value
tick_inc = sheet.Cells(2, 9).Value
tick_dec = sheet.Cells(2, 9).Value

perc_inc = sheet.Cells(2, 11).Value
perc_dec = sheet.Cells(2, 11).Value
volmax = sheet.Cells(2, 12).Value


' Determine worksheet Maximums
'   (1) stock with greatest percentage increase
'   (2) stock with greatest percentage decrease
'   (3) stock with greatest volume total


For i = 3 To tick_end

    If (perc_inc < sheet.Cells(i, 11).Value) Then
        perc_inc = sheet.Cells(i, 11).Value
        tick_inc = sheet.Cells(i, 9).Value
    End If

    If (perc_dec > sheet.Cells(i, 11).Value) Then
        perc_dec = sheet.Cells(i, 11).Value
        tick_dec = sheet.Cells(i, 9).Value
    End If

    If (volmax < sheet.Cells(i, 12).Value) Then
        volmax = sheet.Cells(i, 12).Value
        tick_vol = sheet.Cells(i, 9).Value
    End If

Next i

'  Add worksheet Maximums to xls


sheet.Cells(2, 16).Value = tick_inc
sheet.Cells(3, 16).Value = tick_dec
sheet.Cells(4, 16).Value = tick_vol

sheet.Cells(2, 17).Value = perc_inc
sheet.Cells(2, 17).NumberFormat = "0.00%"
sheet.Cells(3, 17).Value = perc_dec
sheet.Cells(3, 17).NumberFormat = "0.00%"
sheet.Cells(4, 17).Value = volmax

End Sub