Sub vba_challenge()

Dim summarybox As Double
Dim volume As Double
Dim lastrow As Long
Dim yearlychange As Currency
Dim worksheetname As String
Dim maxpercent As Double
Dim maxvolume As Double
Dim minpercent As Double
Dim endofyearvalue As Double
Dim starofyearvalue As Double

summarybox = 2
volume = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For m = 2 To lastrow

    volume = volume + Cells(m, 7)
            
    If Cells(m + 1, 1).Value <> Cells(m, 1) Then
        endofyearvalue = Cells(m, 6).Value

       
        yearlychange = endofyearvalue - startofyearvalue
        Cells(summarybox, 11).Value = yearlychange / Cells(2, 3).Value
        Cells(summarybox, 11).NumberFormat = "0.00%"
        ticker = Cells(m, 1).Value
        Cells(summarybox, 9).Value = ticker
    Cells(summarybox, 10).Value = yearlychange
    Cells(summarybox, 10).NumberFormat = "$0.00"
    If yearlychange > 0 Then
    Cells(summarybox, 10).Interior.ColorIndex = "10"
    Else: Cells(summarybox, 10).Interior.ColorIndex = "3"
    End If
    Cells(summarybox, 11).Value = yearlychange / startofyearvalue
    Cells(summarybox, 11).NumberFormat = "0.00%"
        
    
     
    Cells(summarybox, 12).Value = volume
    Cells(summarybox, 12).NumberFormat = "0,000"
    volume = 0
        summarybox = summarybox + 1
    End If
    
    If Cells(m - 1, 1).Value <> Cells(m, 1) And Cells(m, 3).Value <> 0 Then
        startofyearvalue = Cells(m, 3).Value
       
        
    End If


    
 Next m
    
      
    'maxvolume = WorksheetFunction.max(Range("L2:L8"))
    maxvolume = 0
    maxpercent = 0
    minpercent = 0
    
    lastrowsummary = Cells(Rows.Count, 12).End(xlUp).Row
    
    For j = 2 To lastrowsummary
        If Cells(j, 12) > maxvolume Then
        maxvolume = Cells(j, 12).Value
        Range("P4").Value = Cells(j, 9).Value
        End If
    Next j
        
    For k = 2 To lastrowsummary
        If Cells(k, 11) > maxpercent Then
            maxpercent = Cells(k, 11).Value
            Range("P2").Value = Cells(k, 9).Value
        End If
    Next k
        
    For l = 2 To lastrowsummary
        If Cells(l, 11) < minpercent Then
            minpercent = Cells(l, 11).Value
                Range("P3").Value = Cells(l, 9).Value
        End If
    Next l
    
    Range("Q4").Value = maxvolume
    Range("Q4").NumberFormat = "0,000"
    Range("Q2").Value = maxpercent
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").Value = minpercent
    Range("Q3").NumberFormat = "0.00%"

    
 

    
End Sub



