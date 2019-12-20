Attribute VB_Name = "Module1"
Sub Ticker():

    For Each ws In Worksheets
   
'======================================================================================================
'[0] Cast templates for summmary tables
'======================================================================================================

    'Insert primary summary column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    'Insert secondary summary column headers
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Insert secondary summary row headers
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
'======================================================================================================
'[1] Calculate & write primary summary values
'======================================================================================================
        
    'Set dimensions
    Dim dblYearOpen As Double
    Dim dblYearClose As Double
    Dim dblYearChg As Double
    Dim dblYearPctChg As Double
    Dim curTtlVol As Currency
    Dim intSummaryRow As Integer
    Dim dblGreatPctInc As Double
    Dim dblGreatPctDec As Double
    Dim curGreatTtlVol As Currency
    Dim strGreatPctIncTicker As String
    Dim strGreatPctDecTicker As String
    Dim strGreatTtlVolTicker As String
        
    'Set default values ahead of the loops
    dblYearOpen = 0
    dblYearClose = 0
    dblYearChg = 0
    dblYearPctChg = 0
    curTtlVol = 0
    intSummaryRow = 2
    dblGreatPctInc = 0
    dblGreatPctDec = 0
    curGreatTtlVol = 0
    strGreatPctIncTicker = ""
    strGreatPctDecTicker = ""
    strGreatTtlVolTicker = ""
        
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Store the YearOpen value
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            dblYearOpen = ws.Cells(i, 3).Value
        End If
        
        'Trigger summary calculations
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'Write ticker symbol to the summary table
            ws.Cells(intSummaryRow, 9).Value = ws.Cells(i, 1).Value
            'Store YearClose value
            dblYearClose = ws.Cells(i, 6)
            'Calculate & write YearChg
            dblYearChg = dblYearClose - dblYearOpen
            ws.Cells(intSummaryRow, 10).Value = dblYearChg
            ws.Cells(intSummaryRow, 10).NumberFormat = "0.00"
                If ws.Cells(intSummaryRow, 10).Value > 0 Then
                    ws.Cells(intSummaryRow, 10).Interior.ColorIndex = 50
                Else
                    ws.Cells(intSummaryRow, 10).Interior.ColorIndex = 53
                End If
            'Calculate & write YearPctChg
                If dblYearOpen <> 0 And dblYearChg <> 0 Then
                    dblYearPctChg = dblYearChg / dblYearOpen
                Else
                    dblYearPctChg = 0
                End If
            ws.Cells(intSummaryRow, 11).Value = dblYearPctChg
            ws.Cells(intSummaryRow, 11).NumberFormat = "0.00%"
            'Calculate & write TtlVol
            curTtlVol = curTtlVol + ws.Cells(i, 7).Value
            ws.Cells(intSummaryRow, 12).Value = curTtlVol
            ws.Cells(intSummaryRow, 12).NumberFormat = "0_);(0)"
            'Tee up the next loop
            intSummaryRow = intSummaryRow + 1
            curTtlVol = 0
        Else
            'Not time to summarize yet.  Keep tallying up the TtlVol
            curTtlVol = curTtlVol + ws.Cells(i, 7).Value
        End If
    Next i

'======================================================================================================
'[2] Throw me a fricken bonus-section here
'======================================================================================================
            
    'Calculate & write Greatest % Increase
    For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        If ws.Cells(i, 11).Value > dblGreatPctInc Then
            dblGreatPctInc = ws.Cells(i, 11).Value
            strGreatPctIncTicker = ws.Cells(i, 9).Value
        End If
    Next i
            
    ws.Cells(2, 16).Value = strGreatPctIncTicker
    ws.Cells(2, 17).Value = dblGreatPctInc
    ws.Cells(2, 17).NumberFormat = "0.00%"
                        
    'Calculate & write Greatest % Decrease
    For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        If ws.Cells(i, 11).Value < dblGreatPctDec Then
            dblGreatPctDec = ws.Cells(i, 11).Value
            strGreatPctDecTicker = ws.Cells(i, 9).Value
        End If
    Next i
            
    ws.Cells(3, 16).Value = strGreatPctDecTicker
    ws.Cells(3, 17).Value = dblGreatPctDec
    ws.Cells(3, 17).NumberFormat = "0.00%"
            
    'Calculate & write Greatest Total Volume
    For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        If ws.Cells(i, 12).Value > curGreatTtlVol Then
            curGreatTtlVol = ws.Cells(i, 12).Value
            strGreatTtlVolTicker = ws.Cells(i, 9).Value
        End If
    Next i
            
    ws.Cells(4, 16).Value = strGreatTtlVolTicker
    ws.Cells(4, 17).Value = curGreatTtlVol
    ws.Cells(4, 17).NumberFormat = "0_);(0)"

'======================================================================================================
'[3] Tidy up the formatting (nobody likes scrunchy column widths)
'======================================================================================================
    ws.Columns("I:Q").AutoFit
    Next ws
End Sub
