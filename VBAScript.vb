Sub StockAnalysis()

    ' Variable Declaration
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Variant
    Dim uniqueTickers As Integer
    Dim rowNumbers As Long
    Dim tickerName As String
    Dim previousTickerName As String
    Dim stockVolume As Double
    

    ' Looping through each worksheet
    
    For Each ws In Worksheets

        ' Add headers
        ws.Range("I1") = "Ticker"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        
        ' Sorting the worksheet by Column A and B
        With ws.Sort
             .SortFields.Add Key:=Range("A1"), Order:=xlAscending
             .SortFields.Add Key:=Range("B1"), Order:=xlAscending
             .SetRange Columns("A:G")
             .Header = xlYes
             .Apply
        End With
        
        ' Declaring a new 2D array of large size with four columns
        
        ReDim stockVolume_array(9999, 3) As Variant
        
        ' Variables initialization

        uniqueTickers = 0
        rowNumbers = 2
        tickerName = ws.Range("A" & rowNumbers)
        stockVolume_array(0, 0) = tickerName
        stockVolume = CDbl(ws.Range("G" & rowNumbers).Value)  ' CDbl() returns double data type
        openPrice = CDbl(ws.Range("C" & rowNumbers).Value)
        
        ' Initializing previousTickerName with first tickerName
        
        previousTickerName = tickerName
         
         ' Looping through rows in Column A

        Do While (tickerName <> "")
            
            If tickerName = previousTickerName Then
                
                ' Add volume of row to total stock volume
                stockVolume_array(uniqueTickers, 1) = stockVolume_array(uniqueTickers, 1) + stockVolume
            
            Else
            
                 ' Get close price of previous row
                closePrice = ws.Range("F" & rowNumbers - 1)
                
                ' Calculate and store yearly change
                yearlyChange = closePrice - openPrice
                
                ' Store yearly change in array
                stockVolume_array(uniqueTickers, 2) = yearlyChange
                
                ' Calculate percent change
                If openPrice <> 0 Then
                    percentageChange = (closePrice - openPrice) / openPrice
                
                Else
                    percentageChange = "n/a"
                End If
                
                ' Store percent change in array
                 stockVolume_array(uniqueTickers, 3) = percentageChange
               
                ' Increment uniqueTickers
                uniqueTickers = uniqueTickers + 1
                
                ' Store tickerName and stockVolume
                stockVolume_array(uniqueTickers, 0) = tickerName
                stockVolume_array(uniqueTickers, 1) = stockVolume
                
                ' Store new open price
                openPrice = ws.Range("C" & rowNumbers)
                
            End If
            

            ' Updating the previousTickerName to current tickerName for next loop
            previousTickerName = tickerName
            
            ' Updating the rowNumbers to the next row in Column A and then the tickerName and stockVolume
            rowNumbers = rowNumbers + 1
            tickerName = ws.Range("A" & rowNumbers) ' new ticker name
            stockVolume = CDbl(ws.Range("G" & rowNumbers).Value)

        Loop
        

        ' Looping through stockVolume_array and print out the results

        For i = 0 To uniqueTickers - 1
            
            ' add ticker name
            tickerName = stockVolume_array(i, 0)
            ws.Range("I" & i + 2) = tickerName
            
            ' add total volume of stock
            stockVolume = stockVolume_array(i, 1)
            ws.Range("L" & i + 2) = stockVolume
            
            ' add yearly change
            yearlyChange = stockVolume_array(i, 2)
            ws.Range("J" & i + 2) = yearlyChange
            
            
            ' Conditional formating: positive change in green and negative change in red in Column J
            
            If yearlyChange > 0 Then
                ws.Range("J" & i + 2).Interior.Color = vbGreen
            Else
                ws.Range("J" & i + 2).Interior.Color = vbRed
            End If
            
            ' Print percent change
            ws.Range("K" & i + 2) = stockVolume_array(i, 3)
        Next i

    Next ws
    
End Sub




