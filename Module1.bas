Attribute VB_Name = "Module1"
Sub TSA_prediction()

Dim ticker

ticker = "NVDA"
Worksheets(ticker).Activate

cal_month_mean
cal_DMA
cal_MMA
cal_CMMA
cal_RMA
cal_SI

next_year_date

'the remaining part of TSA_predictional calculates the price of a stock using Time Series Analysis
Columns("Q").Clear
Range("Q1").Value = "Prediction"

'first create an array containing the row number of the 4 seasonal index
Dim dataRange As Range
Dim lastRow As Long
Dim rowNumbers() As Long
Dim rowCount As Long
Dim row_num As Long
    
lastRow = Cells(Rows.count, 9).End(xlUp).Row
    
ReDim rowNumbers(0 To 0)
rowCount = 0
    
For row_num = 2 To lastRow
    If Not IsEmpty(Cells(row_num, 14).Value) Then
        ' Resize the array to accommodate the new row number
        ReDim Preserve rowNumbers(0 To rowCount)
        rowNumbers(rowCount) = row_num
        rowCount = rowCount + 1
    End If
Next row_num

'calculation: prediction = DMA one year ago x seasonal index
Dim count As Long
count = 0
Dim si

Dim predictionStartRow As Long
predictionStartRow = 5
For j = predictionStartRow To lastRow
    If j > rowNumbers(count) Then
        count = count + 1
    End If
    si = Cells(rowNumbers(count), 14).Value
    Cells(j, 17).Value = Cells(j, 10).Value * si

Next j

'scaling for continuity of stock price crossing years
Dim diff
diff = Cells(predictionStartRow, 17).Value - Cells(lastRow, 2).Value
For i = predictionStartRow To lastRow
    Cells(i, 17).Value = Cells(i, 17) - diff
Next i

Call PlotPricePredictionGraph(ticker, lastRow)

Columns("C:G").Hidden = True
Columns("I:O").Hidden = True

End Sub

Function cal_month_mean()
    Columns(9).Clear
    Range("I1").Value = "Mean Value of the Month"
    Dim dataRange As Range
    Dim lastRow As Long
    
    ' Find the last used row in column A
    lastRow = Cells(Rows.count, "A").End(xlUp).Row
    
    ' Set the data range
    Set dataRange = Range("A1:B" & lastRow)
    
    Dim counter As Long
    Dim month As Integer
    Dim acc_value As Double
    Dim date_value As Date
    Dim curr_month As Integer
    Dim row_num As Long
    Dim month_means() As Double
    
    ReDim month_means(1 To 12)
    
    row_num = 2
    Do While dataRange.Cells(row_num, 1).Value <> ""
        date_value = dataRange.Cells(row_num, 1).Value
        curr_month = DatePart("m", date_value)
        If curr_month = month Then
            counter = counter + 1
            acc_value = acc_value + dataRange.Cells(row_num, 2).Value
        Else
            If counter > 0 Then
                month_means(month) = acc_value / counter
                Cells(row_num - 1, 9).Value = month_means(month) ' Modify this line to write the value to the cell
            End If
            acc_value = dataRange.Cells(row_num, 2).Value
            counter = 1
            month = curr_month
        End If
        row_num = row_num + 1
    Loop
    
    If counter > 0 Then
        month_means(month) = acc_value / counter
        Cells(row_num - 1, 9).Value = month_means(month) ' Modify this line to write the value to the cell
    End If
    
    ' No need to return the month_means() array since it's already written to the cells
End Function
Function cal_DMA()
    Columns(10).Clear
    Range("J1").Value = "Daily Moving Average (DMA)"
    Dim dataRange As Range
    Dim lastRow As Long
    
    ' Find the last used row in the worksheet
    lastRow = Cells(Rows.count, "A").End(xlUp).Row
    
    ' Set the data range starting from row 5
    Set dataRange = Range("A5:D" & lastRow)
    
    Dim row_num As Long
    Dim moving_avg As Double
    Dim sum As Double
    
    ' Loop through the data range starting from row5
    For row_num = 5 To lastRow
        ' Calculate the moving average
        sum = Cells(row_num - 3, 2).Value + Cells(row_num - 2, 2).Value + Cells(row_num - 1, 2).Value + Cells(row_num, 2).Value
        moving_avg = sum / 4
        
        ' Write the moving average to the 10th column
        dataRange.Cells(row_num - 4, 10).Value = moving_avg
    Next row_num
    
    calc_moving_average = "Moving average calculated and written to column J."
End Function

Function cal_MMA()
    Columns(11).Clear
    Range("K1").Value = "Monthly Moving Average (MMA)"
    Dim dataRange As Range
    Dim lastRow As Long
    Dim rowNumbers() As Long
    Dim rowCount As Long
    Dim row_num As Long
    
    lastRow = Cells(Rows.count, 9).End(xlUp).Row
    
    ReDim rowNumbers(0 To 0)
    rowCount = 0
    
    For row_num = 2 To lastRow
        If Not IsEmpty(Cells(row_num, 9).Value) Then
            ' Resize the array to accommodate the new row number
            ReDim Preserve rowNumbers(0 To rowCount)
            rowNumbers(rowCount) = row_num
            rowCount = rowCount + 1
        End If
    Next row_num
    
    For i = 3 To UBound(rowNumbers)
    Dim sum As Double
    Dim j As Long
    
    ' Calculate the sum of the current cell and the previous 3 cells
    sum = Cells(rowNumbers(i), 9) + Cells(rowNumbers(i - 1), 9) + Cells(rowNumbers(i - 2), 9) + Cells(rowNumbers(i - 3), 9)
    ' Calculate the moving average and store it in the movingAverage array
    Cells(rowNumbers(i), 11) = sum / 4
Next i
    
End Function

Function cal_CMMA()
    Columns(12).Clear
    Range("L1").Value = "Center Monthly Moving Average (CMMA)"
    Dim dataRange As Range
    Dim lastRow As Long
    Dim rowNumbers() As Long
    Dim rowCount As Long
    Dim row_num As Long
    
    lastRow = Cells(Rows.count, 9).End(xlUp).Row
    
    ReDim rowNumbers(0 To 0)
    rowCount = 0
    
    For row_num = 2 To lastRow
        If Not IsEmpty(Cells(row_num, 11).Value) Then
            ' Resize the array to accommodate the new row number
            ReDim Preserve rowNumbers(0 To rowCount)
            rowNumbers(rowCount) = row_num
            rowCount = rowCount + 1
        End If
    Next row_num
    
    For i = 1 To UBound(rowNumbers)
    Dim sum As Double
    Dim j As Long
    
    ' Calculate the sum of the current cell and the previous 1 cell
    sum = Cells(rowNumbers(i), 11) + Cells(rowNumbers(i - 1), 11)
    
    Cells(rowNumbers(i), 12) = sum / 2
Next i

End Function

Function cal_RMA()
    Columns(13).Clear
    Range("M1").Value = "Ratio to Moving Average (RMA)"
    Dim dataRange As Range
    Dim lastRow As Long
    Dim rowCount As Long
    Dim row_num As Long
    
    lastRow = Cells(Rows.count, 9).End(xlUp).Row
    
    ReDim rowNumbers(0 To 0)
    rowCount = 0
    
    For row_num = 2 To lastRow
        If Not IsEmpty(Cells(row_num, 12).Value) Then
            Cells(row_num, 13).Value = Cells(row_num, 9) / Cells(row_num, 12)
        End If
    Next row_num


End Function

Function cal_SI()
    Columns(14).Clear
    Range("N1").Value = "Seasonal Index"
    Dim dataRange As Range
    Dim lastRow As Long
    Dim rowNumbers() As Long
    Dim rowCount As Long
    Dim row_num As Long
    
    lastRow = Cells(Rows.count, 9).End(xlUp).Row
    
    ReDim rowNumbers(0 To 0)
    rowCount = 0
    
    For row_num = 2 To lastRow
        If Not IsEmpty(Cells(row_num, 13).Value) Then
            ' Resize the array to accommodate the new row number
            ReDim Preserve rowNumbers(0 To rowCount)
            rowNumbers(rowCount) = row_num
            rowCount = rowCount + 1
        End If
    Next row_num
    
    
    
    For i = 1 To UBound(rowNumbers)
        Dim sum As Double
        Dim j As Long
        
        ' Calculate the sum of the current cell and the previous 1 cell
        sum = Cells(rowNumbers(i), 13) + Cells(rowNumbers(i - 1), 13)
        
        Cells(rowNumbers(i), 14) = sum / 2
    Next i
    
   For i = 1 To UBound(rowNumbers)
    If i Mod 3 <> 1 Then
        Cells(rowNumbers(i), 14).Clear
    End If
Next i

End Function

Function next_year_date()
    Columns(16).Clear
    Range("P1").Value = "Date"
    lastRow = Range("A" & Rows.count).End(xlUp).Row

    For i = 2 To lastRow
        
        Dim currentDate As Date
        currentDate = Range("A" & i).Value

        Dim nextDate As Date
        nextDate = DateAdd("yyyy", 1, currentDate)

        Range("P" & i).Value = nextDate
        j = j + 1
    Next i
End Function


Sub PlotPricePredictionGraph(ticker, lastRow)
    Dim cht As ChartObject
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets(ticker)
    
    ' Remove any existing chart on the worksheet
    On Error Resume Next
    For Each cht In ws.ChartObjects
        cht.Delete
    Next cht
    On Error GoTo 0

    ' Define the ranges for data and dates
    Set realData = ws.Range("B2:B" & lastRow)
    Set realDates = ws.Range("A2:A" & lastRow)
    Set predictData = ws.Range("Q2:Q" & lastRow)
    Set predictDates = ws.Range("P2:P" & lastRow)
    Set blankData = ws.Range("H2:H" & lastRow)

    ' Create a new chart object on the same worksheet
    Set cht = ws.ChartObjects.Add(Left:=ws.Cells(2, 19).Left, _
                                  Width:=1000, _
                                  Top:=ws.Cells(2, 19).Top, _
                                  Height:=400)
    
    ' Set the chart type to line chart
    cht.Chart.ChartType = xlLine
    
    ' Set the chart data
    cht.Chart.SeriesCollection.NewSeries
    cht.Chart.SeriesCollection(1).XValues = Union(realDates, predictDates)
    cht.Chart.SeriesCollection(1).Values = Union(realData, blankData)
    cht.Chart.SeriesCollection(1).Name = "Actual Price"

    cht.Chart.SeriesCollection.NewSeries
    cht.Chart.SeriesCollection(2).XValues = Union(realDates, predictDates)
    cht.Chart.SeriesCollection(2).Values = Union(blankData, predictData)
    cht.Chart.SeriesCollection(2).Name = "Predicted Price"

    ' Set the chart title
    cht.Chart.HasTitle = True
    cht.Chart.ChartTitle.Text = ticker & _
        " Prices (" & Year(realDates.Cells(1, 1).Value) & ") " & _
        "and Prediction (" & Year(predictDates.Cells(1, 1).Value) & ")"
    
    ' Set the X-axis title
    With cht.Chart.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Date"
    End With
    
    ' Set the Y-axis title
    With cht.Chart.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Price"
    End With
End Sub


