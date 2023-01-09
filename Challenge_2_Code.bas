Attribute VB_Name = "Module11"
'Defining Worksheet sub

Sub DefineWS():
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Select
        Call Challenge2
    Next ws
End Sub
'Defining main sub
Sub Challenge2():
    Dim rowtotal As Long
    Dim opener As Double
    Dim closer As Double
    Dim yearlychange As Single
    Dim percentchange As Single
    Dim totalstockvolume As Double
    Dim newticker As String
    Dim rowcountnew As LongPtr
    Dim maximumpercent As Double
    Dim maximumticker As String
    Dim minimumpercent As Double
    Dim minimumticker As String
    Dim maximumvolume As Double
    Dim maximumvolumeticker As String
    
    [I1] = "Ticker"
    [J1] = "Yearly Change"
    [K1] = "Percent Change"
    [L1] = "Total Stock Volume"
    [O1] = "Ticker"
    [P1] = "Value"
    [N2] = "Greatest % Increase"
    [N3] = "Greatest % Decrease"
    [N4] = "Greatest Total Volume"
    
    maximumpercent = 0.001
    minimumpercent = 0
    maximumvolume = 1
    rowtotal = Cells(Rows.Count, "A").End(xlUp).Row
    Cells(2, 9).Value = Cells(2, 1).Value
    nextrow = 2
    opener = [C2]
    totalstockvolume = 0
    For i = 2 To (rowtotal)
        closer = Cells(i, 6).Value
        totalstockvolume = totalstockvolume + Cells(i, 7).Value
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Cells(nextrow, 12) = totalstockvolume
            totalstockvolume = 0
            newticker = Cells(i, 1).Value
            Range("I" & nextrow).Value = newticker
            yearlychange = closer - opener
            'Populating yearly change column
            Cells(nextrow, 10).Value = yearlychange
            If opener > 0 Then
                'MsgBox (opener & " " & closer)
                percentchange = (closer - opener) / opener
                'MsgBox (percentchange)
                Cells(nextrow, 11).Value = percentchange
            Else
                Cells(nextrow, 11).Value = "Invalid"
            End If
            nextrow = nextrow + 1
            opener = Cells(i + 1, 3).Value
            'MsgBox (opener & " " & i)
        Else
        End If
    Next i
           
    rowcountnew = Cells(Rows.Count, "I").End(xlUp).Row
    Range("K2:K" & rowcountnew).NumberFormat = "0.00%"
    For j = 2 To rowcountnew
        'Make cell colorless for first row
        Cells(1, 10).Interior.ColorIndex = 0
        If Cells(j, 10).Value < 0 Then
            Cells(j, 10).Interior.ColorIndex = 3
        Else
            Cells(j, 10).Interior.ColorIndex = 4
        End If
    Next j
    'For loop for percent change conditional formatting
    For k = 2 To rowcountnew
        If (Cells(k, 11).Value <> "Invalid") Then
            If Cells(k, 11).Value > 0 Then
                Cells(k, 11).Interior.ColorIndex = 4
            End If
            If Cells(k, 11).Value < 0 Then
                    Cells(k, 11).Interior.ColorIndex = 3
            End If
            'Define greatest percent cahnge
            If Cells(k, 11).Value > maximumpercent Then
                maximumpercent = Cells(k, 11).Value
                maximumticker = Cells(k, 9)
            End If
                       
        Else
            Cells(k, 11).Value = "Invalid"
        End If
    Next k


    [O2] = maximumticker
    [P2] = maximumpercent

    For l = 2 To rowcountnew
        If (Cells(l, 11).Value < minimumpercent) Then
            minimumpercent = Cells(l, 11).Value
            minimumticker = Cells(l, 9).Value
        End If
    Next l

    [O3] = minimumticker
    [P3] = minimumpercent

    For m = 2 To rowcountnew
        If (Cells(m, 12).Value > maximumvolume) Then
            maximumvolume = Cells(m, 12).Value
            maximumvolumeticker = Cells(m, 9).Value
        End If
    Next m

    [O4] = maximumvolumeticker
    [P4] = maximumvolume
    Range("P2:P3").NumberFormat = "0.00%"
       
        
    
End Sub


