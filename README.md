Here is my code:
Sub StockData()

    Dim ws As Worksheet
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    
    ' Loop through each worksheet (each quarter)
    For Each ws In ThisWorkbook.Worksheets
    
        ' Set headers for output
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Set initial values
        j = 0
        total = 0
        change = 0
        start = 2
        
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To rowCount
        
            'If ticker changes then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Results in variables
                total = total + ws.Cells(i, 7).Value
                
                'Handle zero total volume
                If total = 0 Then
                    'print the results
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0
                    
                Else
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                    
                    'Calculate change
                    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percentChange = change / ws.Cells(start, 3)
                    
                    'next stock ticker
                    start = i + 1
                    
                    'print results
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = change
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    ws.Range("K" & 2 + j).Value = percentChange
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + j).Value = total
                    
                    'colors: positives green and negative red
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3 'added back green by changing 2 to 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                End If
                
                total = 0
                change = 0
                j = j + 1
                days = 0
                
            Else
                total = toal + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        'max and min
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
        
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        
        'final part
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
        
    Next ws
    
  MsgBox "Stock data analysis complete!"
  
End Sub


I utilized the Xpert Learning tool to help begin my code on Developer.
Here is the original code that I received from the Xpert Learning tool:
Sub AnalyzeQuarterlyStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim outputRow As Long
    Dim outputSheet As Worksheet

    ' Create or set the output sheet
    On Error Resume Next
    Set outputSheet = ThisWorkbook.Sheets("StockAnalysis")
    On Error GoTo 0
    
    If outputSheet Is Nothing Then
        Set outputSheet = ThisWorkbook.Sheets.Add
        outputSheet.Name = "StockAnalysis"
    Else
        outputSheet.Cells.Clear ' Clear previous data
    End If

    ' Set headers for output
    outputSheet.Cells(1, 1).Value = "Ticker Symbol"
    outputSheet.Cells(1, 2).Value = "Quarterly Change"
    outputSheet.Cells(1, 3).Value = "Percent Change"
    outputSheet.Cells(1, 4).Value = "Total Volume"

    ' Initialize output row
    outputRow = 2 ' Start output from row 2

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Skip the output sheet
        If ws.Name <> "StockAnalysis" Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Assuming ticker symbol is in column A

            ' Loop through each row of stock data
            For i = 2 To lastRow ' Assuming headers are in row 1
                ticker = ws.Cells(i, 1).Value ' Ticker symbol in column A
                openPrice = ws.Cells(i, 2).Value ' Open price in column B
                closePrice = ws.Cells(i, 3).Value ' Close price in column C
                volume = ws.Cells(i, 4).Value ' Volume in column D

                ' Calculate quarterly change and percentage change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If

                ' Output the results to the output sheet
                outputSheet.Cells(outputRow, 1).Value = ticker ' Output ticker symbol
                outputSheet.Cells(outputRow, 2).Value = quarterlyChange ' Output quarterly change
                outputSheet.Cells(outputRow, 3).Value = percentChange ' Output percentage change
                outputSheet.Cells(outputRow, 4).Value = volume ' Output total volume

                ' Move to the next output row
                outputRow = outputRow + 1
            Next i
        End If
    Next ws

    ' Apply conditional formatting for positive and negative changes in the output sheet
    With outputSheet.Range("B2:B" & outputRow - 1)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive change
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative change
    End With

    MsgBox "Stock data analysis complete!"
End Sub


After going through this code, I realized that it created a separate sheet that looped through all the sheets to conduct one output, which is not what I wanted. I removed the "outputSheet" and replaced those with "ws" however, I was still missing some things.
I referenced our in class learnging to change the colors of the cells formula to line up with what we have been taught.
After this point, I decided to attend office hours to edit my code. I found I had additional "Dim" declarations that weren't necessary and needed other variables defined. The instructor and I were able to reconstruct my code.
Following one of our classes, I and some other classmates joined in a zoom call to discuss what information we had for our homework. We asked each other questions with formulating this in accordance to the instructions provided.
I found that utilizing "Range" and "Cell" was a minor difference, but for this assignment, I leaned towards using "Range".

I reused Xpert Learning to confirm that I was properly looping through all the spreadsheets getting separate outputs, as my Greatest % Increase/Decrease were the same on each sheet and I needed to adjust my code.
Here is what Xpert Learning provided me with:
Sub LoopThroughSheets()

    Dim ws As Worksheet

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Perform actions on each worksheet
        ws.Cells.Font.Bold = True ' Example action: make all text bold
        ws.Cells.HorizontalAlignment = xlCenter ' Center align all text
    Next ws

End Sub

After outsourcing all the help I could get with this assignment, I can finally say I am confident with the code I have created!

Thank you.



