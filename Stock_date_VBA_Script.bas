Attribute VB_Name = "Module1"
Sub stock_data()
    Dim Ticker As Variant 'do i even need this?
    Dim Qtly_Change As Variant
    Dim currentTickerOpen As Variant
    Dim currentTickerClose As Variant
    Dim Percent_Change As Variant
    Dim Total_Stock_Vol As Variant
    Dim LastRow As Variant
    Dim ws As Worksheet
    Dim newTickerValue As Variant
    Dim lastRowWithCondition As Variant
    Dim currenttickerstartrow As Variant
    Dim rng As Variant
Dim minVal As Variant
Dim maxVal As Variant
Dim rngVol As Variant
Dim maxValVol As Variant

    'Dim FirstQty_Change As Variant
    'Dim Firsttickeropen As Variant
    'Dim firsttickerclose As Variant
    
        
    For Each ws In Worksheets
        ' Add column titles
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Find the last row for each worksheet
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row


        tickernewrow = 1

        ' Loop through rows starting from the second row. if condition not met, start new row in column I that has new ticker ID
        For i = 2 To LastRow
            
            If ws.Cells(i, "A").Value <> ws.Cells(i - 1, "A").Value Then
                tickernewrow = tickernewrow + 1
                ws.Cells(tickernewrow, "I").Value = ws.Cells(i, "A").Value
                
                    'update current ticker value showing to be what is in the condition breaking row (i)
                    newTickerValue = ws.Cells(i, "A").Value
                    
                    lastRowWithCondition = i - 1 'row before condition breaks (condition breaks in i)
                   
                    For j = lastRowWithCondition To 2 Step -1
                        'if value w/in LadtRowWithCondition is not same as one above it
                        If ws.Cells(j, "A").Value <> ws.Cells(j - 1, "A").Value Then
                        ' Calculate Quarterly Change based on the condition
                        currenttickerstartrow = j
                        currentTickerOpen = ws.Cells(currenttickerstartrow, "C").Value
                        currentTickerClose = ws.Cells(lastRowWithCondition, "F").Value
                        Qtly_Change = currentTickerClose - currentTickerOpen
                        ws.Cells(tickernewrow, "J").Value = Qtly_Change
                            If Qtly_Change > 0 Then
                            ws.Cells(tickernewrow, "J").Interior.Color = RGB(0, 255, 0)
                            ElseIf Qtly_Change < 0 Then
                            ws.Cells(tickernewrow, "J").Interior.Color = RGB(255, 0, 0)
                            Else: ws.Cells(tickernewrow, "J").Interior.Color = xlNone
                            
                            
                            End If
                            
                        'do color code for qtly_change value dep if over(green), red(under), or same as 0(blank)
                        
                        'calc percent change and populate it
                        Percent_Change = (Qtly_Change / currentTickerOpen)
                        ws.Cells(tickernewrow, "K").Value = Percent_Change
                        
                        'calc total stock vol for quarter for each ticker. populate it
                        
                        Total_Stock_Vol = WorksheetFunction.Sum(ws.Range("G" & currenttickerstartrow & ":G" & lastRowWithCondition))
                        ws.Cells(tickernewrow, "L").Value = Total_Stock_Vol
                        End If
                    Next j
                     
                     'add titles
                        ws.Cells(2, "N").Value = "Greatest Percent Increase"
                        ws.Cells(3, "N").Value = "Greatest Percent Decrease"
                        ws.Cells(4, "N").Value = "Greatest Total Volume"
                        ws.Cells(1, "O").Value = "Ticker"
                        ws.Cells(1, "P").Value = "Value"
                        
                    'calculate Greatest Percent Decrease
                    ' Set the range to the column where you want to find the lowest value
                        Set rng = ws.Range("K2:K" & LastRow)
                            
                            ' Initialize minVal with the first cell value in the range
                            minVal = rng.Cells(1).Value
                            
                            ' Loop through each cell in the range to find the lowest value
                            For Each cell In rng
                                If cell.Value < minVal Then
                                    minVal = cell.Value
                                End If
                            Next cell
                        ws.Cells(2, "P").Value = minVal
                        
                        'calculate Greatest Percent Increase
                            ' Initialize maxVal with the first cell value in the range
                            maxVal = rng.Cells(1).Value
                            
                            ' Loop through each cell in the range to find the highest value
                            For Each cell In rng
                                If cell.Value > maxVal Then
                                    maxVal = cell.Value
                                End If
                            Next cell
                        ws.Cells(3, "P").Value = maxVal
                        
                        
                        'calculate Greatest Total Volume
                        ' Set the range to the column where you want to find the highest value of vol
                            Set rngVol = ws.Range("L2:L" & LastRow)
                            
                            ' Initialize maxValVol with the first cell value in the range
                            maxValVol = rngVol.Cells(1).Value
                            
                            ' Loop through each cell in the range to find the lowest value
                            For Each cell In rngVol
                                If cell.Value < maxValVol Then
                                    maxValVol = cell.Value
                                End If
                            Next cell
                        ws.Cells(4, "P").Value = minVal

                     
                    ' Final calculations for the last set of rows
                    'For h = lastRowWithCondition To LastRow
                        'Firsttickeropen = ws.Cells(2, "C").Value
                        'firsttickerclose = ws.Cells((i - 1), "F").Value
                        '[] -[] = FirstQty_Change
                        'FirstQty_Change = ws.Cells().Value
                       
                    'Next h
      
       End If
        Next i
    Next ws
End Sub
