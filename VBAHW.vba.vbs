Sub sheet2014()

Dim ttl As Double, i As Long, change As Single, j As Integer, strt As Long, rowNo As Integer
Dim perChange As Single, days As Integer, dailyChange As Single, avgChange As Single


cells(1, 9) = "Ticker"
cells(1, 10) = "Yearly Change"
cells(1, 11) = "Percent Change"
cells(1, 12) = "Total Stock Volume"


j = 0
ttl = 0
change = 0
strt = 2


' get the row number of the last row with data
RowCount = cells(Rows.Count, "A").End(xlUp).Row

'Create a script that will loop through all the stocks for one year and output the following information.


For i = 2 To RowCount
    If cells(i + 1, 1).Value <> cells(i, 1).Value Then

        ' Stores results in variables
        ttl = ttl + cells(i, 7).Value
        change = (cells(i, 6) - cells(strt, 3))
        perChange = (change / cells(strt, 3) * 100)
        

        
        
        ' start of the next stock ticker
        strt = i + 1

        ' print the results to the appropriate columns
        Range("I" & 2 + j).Value = cells(i, 1).Value
        Range("J" & 2 + j).Value = change
        Range("K" & 2 + j).Value = "%" & perChange
        Range("L" & 2 + j).Value = ttl

        ' colors positives green and negatives red
        Select Case change
            Case Is > 0
               Range("J" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                Range("J" & 2 + j).Interior.ColorIndex = 3
            Case Else
                Range("J" & 2 + j).Interior.ColorIndex = 0
        End Select
        
        ' reset variables for new stock ticker
        ttl = 0
        change = 0
        j = j + 1
        
        
        
   
    Else
        ttl = ttl + cells(i, 7).Value
        change = change + (cells(i, 6) - cells(i, 3))

        

    End If
Next i




End Sub
