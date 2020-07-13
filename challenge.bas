Attribute VB_Name = "Module1"
Sub stock_homework()

Dim total_volume As Double
Dim year_change As Double
Dim percent_change As Double
Dim open_date As Double
Dim close_date As Double
Dim summary_table_row As Integer
Dim last_row As Long
Dim ws As Worksheet
    'Loop through sheets

    For Each ws In Worksheets

        'add headers
         ws.Range("I1").Value = "Ticker"
         ws.Range("J1").Value = "Yearly Change"
         ws.Range("K1").Value = "Percent Change"
         ws.Range("L1").Value = "Total Volume"

        'beginning values
        total_volume = 0
        open_date = 0
        close_date = 0
        yearly_change = 0
        percent_change = 0
        summary_table_row = 2
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row




        For i = 2 To last_row

            'calculate total volume
            total_volume = total_volume + ws.Cells(i, 7)

            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                open_date = ws.Cells(i, 3).Value
            End If

            'Conditional to find ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(summary_table_row, 12).Value = total_volume
                ws.Cells(summary_table_row, 9).Value = ws.Cells(i, 1).Value

                'calculate price change
                close_date = ws.Cells(i, 6).Value
                yearly_change = close_date - open_date
                ws.Cells(summary_table_row, 10).Value = yearly_change

                'conditional color formatting
                If yearly_change >= 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                End If

        'cclculate percent change from yearly change
        If open_date = 0 And close_date = 0 Then
            percent_change = 0
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
        ElseIf date_open = 0 Then
            Dim new_stock As String
            new_stock = "New Stock"
            ws.Cells(summary_table_row, 11).Value = new_stock
        Else
            percent_change = year_change / open_date
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
            End If
            End If
        Next i
    Next ws
End Sub