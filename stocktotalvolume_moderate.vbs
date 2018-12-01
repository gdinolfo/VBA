sub stocktotalvolume()

numrows = Range("A1", Range("A1").end(xldown)).Rows.Count

    volume_total = 0
    Sum_row = 2
    Ticker  = ""

for i = 2 to numrows
    if cells(i,1).value <> cells(i+1, 1).value then
        ticker = cells(i,1).value
        volume_total = volume_total + cells(i,7).value
        Range("I" & Sum_row).value = ticker
        Range("J" & Sum_row).value = volume_total
        Sum_row = Sum_row + 1
        volume_total = 0
    else
        volume_total = volume_total + cells(i,7).value
    End if

    next i

end sub