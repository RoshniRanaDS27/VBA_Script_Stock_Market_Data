Attribute VB_Name = "Module1"
Sub LoopThroughSheetsAndCalculateYearlyChanges()
    Dim ws As Worksheet
    
    ' Loop through each specified sheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "2018" Or ws.Name = "2019" Or ws.Name = "2020" Then
            ' Call the CalculateYearlyChanges subroutine for each sheet
            CalculateYearlyChanges ws
        End If
    Next ws
End Sub

Sub CalculateYearlyChanges(ws As Worksheet)
    Dim lastRow As Long
    Dim row As Long
    Dim i As Long
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim summaryRow As Integer
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim rng As Range
    Dim firstStock As Boolean

    ' Find the last row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Add headers for the new columns
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Yearly Change ($)"
        ws.Cells(1, 11).Value = "Percent Change (%)"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest(%)increase"
        ws.Cells(3, 15).Value = "Greatest(%)Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
  ' Initialize summary row
        summaryRow = 2
        
          ' Initialize variables for tracking greatest increase, decrease, and volume
    greatestIncrease = -999999999 'initializing with a very low value
    greatestDecrease = 999999999 'initializing with a very high value
    greatestVolume = 0
    firstStock = True
    
    ' Loop through each row in the worksheet
    For i = 2 To lastRow
        
        ' Check if the next row contains a new ticker symbol
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        
        ' Calculate and output the yearly change and percent change
                If totalVolume <> 0 Then
                    ws.Cells(summaryRow, 9).Value = ticker
                    ws.Cells(summaryRow, 10).Value = closePrice - openPrice 'yearly change
                    If openPrice <> 0 Then
                    ws.Cells(summaryRow, 11).Value = ((closePrice - openPrice) / openPrice) * 100 'percentage change
                    Else
                    ws.Cells(summaryRow, 11).Value = "N/A" ' Avoid division by zero
                    End If
                    
                      ' Update percentageChange
    
    If openPrice <> 0 Then
        percentageChange = ((closePrice - openPrice) / openPrice) * 100
    Else
        percentageChange = 0 ' Avoid division by zero
    End If
            
              ' Update greatest increase and decrease

            
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    greatestIncreaseTicker = ws.Cells(i, 1).Value
                End If
                
                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    greatestDecreaseTicker = ws.Cells(i, 1).Value
                End If
                
                    ws.Cells(summaryRow, 12).Value = totalVolume
                    summaryRow = summaryRow + 1
           End If
                
                ' Reset variables for the new stock
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0
            End If
            
            ' Accumulate total volume and update closing price
            totalVolume = totalVolume + ws.Cells(i, 7).Value
             closePrice = ws.Cells(i, 6).Value
             
             ' Update greatest increase Volume
             
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ws.Cells(i, 1).Value
                End If
              
            
    Next i
    
    ' Output the stock with greatest increase, decrease, and volume
    ws.Cells(2, 16).Value = greatestIncreaseTicker
    ws.Cells(3, 16).Value = greatestDecreaseTicker
    ws.Cells(4, 16).Value = greatestVolumeTicker
    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(4, 17).Value = greatestVolume
    
   
    ' Define the range for column J 'as per the instruction mentioned to match with given image
    
    Set rng = ws.Range("j2:K" & lastRow) 'if we want to apply for column K too, then update range "("j2:k" & lastRow)" instead
    
    ' Apply conditional formatting to the range for negative and positive values
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 0, 0) ' Set the color as red for negative values
    End With
    
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(0, 255, 0) ' Set the color as green for positive values
    End With
End Sub

