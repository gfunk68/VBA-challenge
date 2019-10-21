Sub vba_challenge()

Dim summarybox As Integer
Dim volume As Double
Dim lastrow As Long
Dim yearlychange As Currency
Dim worksheetname As String
Dim maxpercent As Double
Dim maxvolume As Double
Dim minpercent As Double

summarybox = 2
volume = 0

For Each ws In Worksheets


    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    worksheetname = ws.Name
    ticker = Split(worksheetname, "_")
    yearlychange = ws.Cells(lastrow, 6).Value - ws.Cells(2, 3).Value
    Cells(summarybox, 9).Value = ws.Cells(summarybox, 1).Value
    Cells(summarybox, 10).Value = yearlychange
    Cells(summarybox, 10).NumberFormat = "$0.00"
    If yearlychange > 0 Then
    Cells(summarybox, 10).Interior.ColorIndex = "10"
    Else: Cells(summarybox, 10).Interior.ColorIndex = "3"
    End If
    Cells(summarybox, 11).Value = yearlychange / ws.Cells(2, 3).Value
    Cells(summarybox, 11).NumberFormat = "0.00%"
    For i = 2 To lastrow
        volume = volume + ws.Cells(i, 7)
    Next i
 
 
    Cells(summarybox, 12).Value = volume
    Cells(summarybox, 12).NumberFormat = "0,000"
    volume = 0
    summarybox = summarybox + 1
    
    'maxvolume = WorksheetFunction.max(Range("L2:L8"))
    maxvolume = 0
    maxpercent = 0
    minpercent = 0
    For j = 2 To 8
        If Cells(j, 12) > maxvolume Then
        maxvolume = Cells(j, 12).Value
        Range("P4").Value = Cells(j, 9).Value
        End If
    Next j
        
        For k = 2 To 8
            If Cells(k, 10) > maxpercent Then
                maxpercent = Cells(k, 11).Value
                Range("P2").Value = Cells(k, 9).Value
            End If
        Next k
        For l = 2 To 8
            If Cells(l, 10) < minpercent Then
                minpercent = Cells(l, 11).Value
                Range("P3").Value = Cells(l, 9).Value
            End If
        Next l
    
    Range("Q4").Value = maxvolume
    Range("Q2").Value = maxpercent
    Range("Q3").Value = minpercent
    

    
 
    Next ws
    
End Sub


