Attribute VB_Name = "Module1"
Sub ticker()

For Page = 1 To ThisWorkbook.Worksheets.Count
    If Sheets(Page).Name = 2018 Or Sheets(Page).Name = 2019 Or Sheets(Page).Name = 2020 Then
        ThisWorkbook.Worksheets(Page).Activate
    End If
    Dim ticker As String
    Dim num As Integer
    Dim total As Double
    Dim beginning As Double
    Dim ending As Double
    Dim difference As Double
    Dim percent_change As Double
    Dim increase As Double
    Dim decrease As Double
    Dim greatest As Double

    num = 2
    total = 0
    beginning = Cells(2, 3).Value
    increase = -9999
    decrease = 9999
    greatest = -9999
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow
        ticker = Cells(i, 1).Value
        total = total + Cells(i, 7).Value
        If Cells(i + 1, 1).Value <> ticker Then
            ending = Cells(i, 6).Value
            difference = ending - beginning
            percent_change = (difference / beginning)
            beginning = Cells(i + 1, 3)
            Cells(num, 9).Value = ticker
            Cells(num, 10).Value = difference
            If difference >= 0 Then
                Cells(num, 10).Interior.ColorIndex = 4
            Else
                Cells(num, 10).Interior.ColorIndex = 3
            End If
            Cells(num, 11).Value = percent_change
            Cells(num, 12).Value = total
            num = num + 1
            total = 0
        End If
    Next i

    For i = 2 To lastrow
        If Cells(i, 11).Value > increase Then
            increase = Cells(i, 11).Value
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = increase
        End If
        If Cells(i, 11).Value < decrease Then
            decrease = Cells(i, 11).Value
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = decrease
        End If
        If Cells(i, 12).Value > greatest Then
            greatest = Cells(i, 12).Value
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = greatest
        End If
    Next i
Next Page
End Sub

