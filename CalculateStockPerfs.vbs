Sub main()
    ' Define variables here
        Dim ticker, nextRowTicker As String
    Dim currentSummaryRow, lastRowNum As Long
    Dim openPrice, closePrice As Single
    ' Stupid VBA Integer can handle only up to 32767 for Integer
    ' so gotta use Long or LongLong
    Dim volume As Long
    Dim totalVolume As LongLong
    Dim yearlyChange As Single


    For Each ws In Worksheets
        Debug.Print ("Starting Tab " & ws.Name)

         'Calculate the last row number in the current Worksheet
        lastRowNum = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Reset total numbers per ticker counters
        totalVolume = 0
        currentSummaryRow = 2

        ' Write summary stats table Header
         ws.Range("I1").Value = "Ticker"
         ws.Range("J1").Value = "Yearly Change"
         ws.Range("K1").Value = "Percent Change"
         ws.Range("L1").Value = "Total Stock Volume"

         'Set summary stats column format
         ws.Columns("J").NumberFormat = "0.00"
         ws.Columns("K").NumberFormat = "#.##%"


        'Loop through each stock performance entry in the table
        For i = 2 To lastRowNum
            'I hate dealing with cell/range funcions.  Put them into more human friendly var names.
            ticker = ws.Range("A" & i).Value
            nextRowTicker = ws.Range("A" & i + 1).Value

            ' For first data row only, record the opening price
            If i = 2 Then
                openPrice = ws.Range("C" & i).Value
            End If

            volume = ws.Range("G" & i).Value
            totalVolume = totalVolume + volume

            ' Next row has a new ticker.
            ' Calculate and post stats of the current ticker in currentSummaryRow
            ' then prepare for the next ticker stats counting
            If ticker <> nextRowTicker Then
                closePrice = ws.Range("F" & i).Value

                Debug.Print ("======" & ticker & "======")
                Debug.Print ("Year openPrice: " & openPrice & " closePrice: " & closePrice)
                Debug.Print ("Total Vol: " & totalVolume)

                ws.Range("I" & currentSummaryRow).Value = ticker

                ' YearlyChange Column.  Set font color for loss.
                yearlyChange = closePrice - openPrice
                ws.Range("J" & currentSummaryRow).Value = yearlyChange
                If yearlyChange < 0 Then
                    ws.Range("J" & currentSummaryRow).Interior.ColorIndex = vbRed
                Else
                    ws.Range("J" & currentSummaryRow).Interior.ColorIndex = vbGreen
                End If

                ' Percent Change Column
                ' To prevent div by 0 error.
                If openPrice = 0 Then
                    ws.Range("K" & currentSummaryRow).Value = 0
                Else
                    ws.Range("K" & currentSummaryRow).Value = (closePrice - openPrice) / openPrice
                End If

                ' Total Stock Volume Column
                ws.Range("L" & currentSummaryRow).Value = totalVolume


                Debug.Print ("Resetting counter for ticket: " & nextRowTicker)
                currentSummaryRow = currentSummaryRow + 1
                ' Record the first day of year price of the next ticker.
                openPrice = ws.Range("C" & i + 1).Value

                totalVolume = 0
            End If

        Next i
        ' Exit clause to process the first tab only.  Comment out for quick debugging for 1 tab run
        ' Exit For 'For Each ws In Worksheets
    Next ws 'For Each ws In Worksheets
End Sub
