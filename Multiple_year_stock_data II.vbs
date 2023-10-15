Sub stocks()
    For Each ws In Worksheets

    'Set function for each variable
        Dim tickername As String
        Dim yrchange As Double
        Dim percentchange As Double
        Dim stockvolume As LongLong
        Dim tickerincrease As String
        Dim tickerdecrease As String
        Dim volumeticker As String
        Dim stockvaluedecrease As Double
        Dim stockvalueincrease As Double
        Dim stockvaluevolume As Double
        Dim worksheetname As String
        Dim ticker As String
        Dim value As Double
        
        ws.Activate
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        
'   set column headers for added columns of data created in script

        Range("J1").value = "Ticker"
        Range("K1").value = "Yearly_Change"
        Range("L1").value = "Percent_Change"
        Range("M1").value = "Total_Stock_Volume"
        Range("p1").value = "Ticker"
        Range("q1").value = "Value"
        Cells(2, 15).value = "Greatest % increase"
        Cells(3, 15).value = "Greatest % decrease"
        Cells(4, 15).value = "Greatest Total Volume"
                
                
    'setting the initial values and formating
        Range("L2:L" & lastrow).NumberFormat = "0.00%"
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(4, 17).NumberFormat = "General"
        stockvolume = 0
        openingprice = Cells(2, 3).value
        
    'set the table that is created from gathering all the data under one ticker name
        Dim summary_ticketyr_row As Integer
        summary_ticketyr_row = 2
        
        
    'Set loop to group ticker name, yr change, percent change, & sum total stock volume
        For i = 2 To lastrow
            
            If Cells(i + 1, 1).value <> Cells(i, 1).value Then
                tickername = Cells(i, 1).value
                stockvolume = stockvolume + Cells(i, 7).value
                yrchange = Cells(i, 6).value - openingprice
                percentchange = (yrchange / openingprice)
                
                Range("j" & summary_ticketyr_row).value = tickername
                Range("m" & summary_ticketyr_row).value = stockvolume
                Range("k" & summary_ticketyr_row).value = yrchange
                Range("l" & summary_ticketyr_row).value = percentchange
                
                ' The conditional formating of column k (yearly change)

                If yrchange < 0 Then
                    Cells(summary_ticketyr_row, 11).Interior.ColorIndex = 3
                Else
                    Cells(summary_ticketyr_row, 11).Interior.ColorIndex = 4
                End If
                
                summary_ticketyr_row = summary_ticketyr_row + 1
                openingprice = Cells(i + 1, 3)
                
                stockvolume = 0
                
            Else
                stockvolume = stockvolume + Cells(i, 7).value
                
                
            End If
            
'   finding the max, min, and greatest values from the summary ticket row
        Next i

        lastrow = Cells(Rows.Count, 10).End(xlUp).Row
            stockvalueincrease = Cells(2, 12).value
            stockvaluedecrease = Cells(2, 12).value
            stockvaluevolume = Cells(2, 13).value
            For i = 2 To lastrow
                If Cells(i, 12).value > stockvalueincrease Then
                    tickerincrease = Cells(i, 10).value
                    stockvalueincrease = Cells(i, 12).value
                    Cells(2, 16).value = tickerincrease
                    Cells(2, 17).value = stockvalueincrease
                End If
                If Cells(i, 12).value < stockvaluedecrease Then
                    tickerdecrease = Cells(i, 10).value
                    stockvaluedecrease = Cells(i, 12).value
                    Cells(3, 16).value = tickerdecrease
                    Cells(3, 17).value = stockvaluedecrease
                End If
                If Cells(i, 13).value > stockvaluevolume Then
                    volumeticker = Cells(i, 10).value
                    stockvaluevolume = Cells(i, 13).value
                    Cells(4, 16).value = volumeticker
                    Cells(4, 17).value = stockvaluevolume
                End If
            
            Next i
    Next ws
    
End Sub
