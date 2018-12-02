' run in module

Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stocktotalvolume
    Next
    Application.ScreenUpdating = True

End Sub

Sub stocktotalvolume()

numrows = Range("A1", Range("A1").End(xlDown)).Rows.Count

   volume_total = 0
    sum_row = 2
    open_sum_row = 2
    ticker = ""
    close_sum_row = 2
    open_price = 0
    close_price = 0


For i = 2 To numrows
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        ticker = Cells(i, 1).Value
        volume_total = volume_total + Cells(i, 7).Value
        close_price = Cells(i, 6).Value
        Range("N" & sum_row).Value = close_price
        Range("I" & sum_row).Value = ticker
        Range("L" & sum_row).Value = volume_total
        sum_row = sum_row + 1
        volume_total = 0
    Else
        volume_total = volume_total + Cells(i, 7).Value
    End If

Next i


For k = 2 To numrows

    If Cells(k, 1).Value <> Cells(k - 1, 1).Value Then
        open_price = Cells(k, 3).Value
        Range("M" & open_sum_row).Value = open_price
        open_sum_row = open_sum_row + 1
    End If
    
    Next k
    
sumrows = Range("I1", Range("I1").End(xlDown)).Rows.Count

For m = 2 To sumrows

    Cells(m, 10).Value = Cells(m, 14).Value - Cells(m, 13).Value
    
 Next m
 
For n = 2 To sumrows
    Range("K2:K" & sumrows).NumberFormat = "0.00%"
    If Cells(n, 10).Value = 0 Then
        Cells(n, 10).Value = 1
    End If

    If Cells(n, 13).Value = 0 Then
        Cells(n, 13).Value = 1
    End If

    Cells(n, 11).Value = Cells(n, 10).Value / Cells(n, 13).Value
    
    Next n

For Z = 2 To sumrows
    If Cells(Z, 10).Value > 0 Then
        Cells(Z, 10).Interior.ColorIndex = 4
    ElseIf Cells(Z, 10).Value < 0 Then
        Cells(Z, 10).Interior.ColorIndex = 3
    End If

Next Z
 
End Sub
