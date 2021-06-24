Sub stockmarket()
    
    'define variables
    Dim stockname   As String
    Dim percentage   As Double
    Dim yearlychange As Double
    Dim volume      As Double
    Dim begyearprice As Double
    Dim endyearprice As Double
    Dim enddateyear As Long
    Dim begdateyear As Long
    Dim lastrow     As Long
    
    volume = 0
    endyearprice = 1
    begyearprice = 1
    enddateyear = 0
    begdateyear = 100000000
    
    ' define table
    Dim tablerow    As Integer
    tablerow = 2
    
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    'create loop for volume and name
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            stockname = Cells(i, 1).Value
            volume = volume + Cells(i, 7).Value
            
            'print info in table
            Range("J" & tablerow).Value = stockname
            Range("M" & tablerow).Value = volume
            
            If Cells(i, 2).Value > enddateyear Then
                enddateyear = Cells(i, 2).Value
                endyearprice = Cells(i, 6).Value
            End If
            
            If Cells(i, 2).Value < begdateyear Then
                begdateyear = Cells(i, 2).Value
                begyearprice = Cells(i, 3).Value
            End If
            
            'calculate yearly change and percentage
            yearlychange = endyearprice - begyearprice
            
            If begyearprice = 0 Then
                percentage = 1
            Else
                percentage = yearlychange / begyearprice
            End If
            
            'print info in table
            Range("K" & tablerow).Value = yearlychange
            Range("L" & tablerow).Value = percentage
            
            endyearprice = 1
            begyearprice = 1
            enddateyear = 0
            begdateyear = 100000000
            
            If Cells(tablerow, "K").Value > 0 Then
                Cells(tablerow, "K").Interior.ColorIndex = 4
            End If
            If Cells(tablerow, "K").Value < 0 Then
                Cells(tablerow, "K").Interior.ColorIndex = 3
            End If
            
            tablerow = tablerow + 1
            volume = 0
        Else
            volume = volume + Cells(i, 7).Value
            
            If Cells(i, 2).Value > enddateyear Then
                enddateyear = Cells(i, 2).Value
                endyearprice = Cells(i, 6).Value
            End If
            
            If Cells(i, 2).Value < begdateyear Then
                begdateyear = Cells(i, 2).Value
                begyearprice = Cells(i, 3).Value
            End If
            
        End If
        
    Next i
    
End Sub