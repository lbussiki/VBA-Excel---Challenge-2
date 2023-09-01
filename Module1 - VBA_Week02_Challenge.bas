Attribute VB_Name = "Module1"


Sub Ticker()
Dim totalstockvolume As Double
totalstockvolume = 0
Dim Row As Long
Dim i As Long
Dim ws As Worksheet

    For Each ws In Worksheets
    ws.Range("i1") = "Ticker"
    ws.Range("j1") = "Yearly Change"
    ws.Range("k1") = "Percent Change"
    ws.Range("l1") = "Total Stock Volume"
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Total Volume"


j = 0
Row = 2


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Range("I" & 2 + j) = ws.Cells(i, 1)
        totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
        
       Yearlychange = (ws.Cells(i, 6) - ws.Cells(Row, 3))
       Percentchange = Yearlychange / ws.Cells(Row, 3)
       ws.Range("K" & 2 + j) = Percentchange
        
        Row = i + 1

        ws.Range("j" & 2 + j) = Yearlychange
        If Yearlychange > 0 Then
        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf Yearlychange < 0 Then
        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
        End If
    
    
        
        
        
        
       ws.Range("l" & 2 + j) = totalstockvolume
        
        
        totalstockvolume = 0
        Yearlychange = 0
        
        j = j + 1
        



    Else
        totalstockvolume = totalstockvolume + Cells(i, 7)


    End If
Next i

Dim Greatest_increase As Double
Dim Greatest_increase_ticker As String
Dim greatest_decrease_ticker As String
Dim greatest_decrease As Double
Dim greatest_totalvolume As Double
Dim greatest_totalvolume_ticker As String

lastrow_greatest = ws.Cells(Rows.Count, 11).End(xlUp).Row

' loop over the percent change column in summary table
For k = 2 To lastrow_greatest

    If ws.Cells(k, 11).Value > Greatest_increase Then
        Greatest_increase = ws.Cells(k, 11).Value
        Greatest_increase_ticker = ws.Cells(k, 9).Value
    ElseIf ws.Cells(k, 11).Value < greatest_decrease Then
        greatest_decrease = ws.Cells(k, 11).Value
        greatest_decrease_ticker = ws.Cells(k, 9).Value
    End If
       
    If ws.Cells(k, 12).Value > greatest_totalvolume Then
        greatest_totalvolume = ws.Cells(k, 12).Value
        greatest_totalvolume_ticker = ws.Cells(k, 9).Value
    End If
        
Next k


ws.Range("O2").Value = Greatest_increase_ticker
ws.Range("P2").Value = Greatest_increase

ws.Range("O3").Value = greatest_decrease_ticker
ws.Range("P3").Value = greatest_decrease

ws.Range("P4").Value = greatest_totalvolume
ws.Range("O4").Value = greatest_totalvolume_ticker



Next ws



End Sub


