Attribute VB_Name = "Module5"
Sub tickerStockanalysis()

 ' creating the loop for worksheet
 'loop will go through all worksheets(Q1,Q2,Q3,Q4)
 

     Dim ws As Worksheet
     For Each ws In ActiveWorkbook.Worksheets

    ws.Activate

        ' Calculate the last row of the table by using ws as worksheet and in which cell with row counts in column 1
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).row

        ' Add output headers
        'settin the output headers to fix
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quaterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        Dim open_price As Double
        Dim close_price As Double
        Dim quaterly_change As Double
        Dim ticker As String
        Dim percent_change As Double

        Dim volume As Double
        Dim row As Double
        Dim column As Integer

        volume = 0
        row = 2
      





        ' Set the open price
        open_price = Cells(2, 3).Value

          ' Loop through all ticker
        For i = 2 To last_row


        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

         ' Set the name of the ticker
                ticker = Cells(i, 1).Value
                Cells(row, 9).Value = ticker
                
        ' set close price
                close_price = Cells(i, 6).Value

        ' Calculate  the quaterly change
                quaterly_change = close_price - open_price
                Cells(row, 10).Value = quaterly_change

        ' Calculate the percent change
                    percent_change = quaterly_change / open_price
                    Cells(row, 11).Value = percent_change
                    Cells(row, 11).NumberFormat = "0.00%"


        ' Calculate total volume per quater
                volume = volume + Cells(i, 7).Value
                Cells(row, 12).Value = volume

            ' Loop to the next row
                row = row + 1

            '  Open price to next ticker
                open_price = Cells(i + 1, 3)

            ' Volume for next ticker
                volume = 0

            Else
                volume = volume + Cells(i, 7).Value
            End If
        Next i


        ' Calculate last row of ticker column
        quaterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row

        ' Set the Cell Colors
        For j = 2 To quaterly_change_last_row
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 10
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

        ' Set Ticker, Value, Greatest %, Increase, % Decrease, and Total volume headers
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"


        ' Find the highest value of each ticker
        For x = 2 To quaterly_change_last_row
            If Cells(x, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quaterly_change_last_row)) Then
                Cells(2, 16).Value = Cells(x, 9).Value
                Cells(2, 17).Value = Cells(x, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(x, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & quaterly_change_last_row)) Then
                Cells(3, 16).Value = Cells(x, 9).Value
                Cells(3, 17).Value = Cells(x, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(x, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & quaterly_change_last_row)) Then
                Cells(4, 16).Value = Cells(x, 9).Value
                Cells(4, 17).Value = Cells(x, 12).Value
            End If
        Next x

        ActiveSheet.Range("I:Q").Font.Bold = True
        ActiveSheet.Range("I:Q").EntireColumn.AutoFit
        Worksheets("Q1").Select

    Next ws

End Sub

