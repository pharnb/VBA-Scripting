Attribute VB_Name = "Module1"
Sub stockdatascript()

Dim lastrow As Long
Dim i As Long
Dim ticker As String
Dim chngopn As Double
Dim chngcls As Double
Dim stkvol As Long
Dim chngclsold As Double
lastrow = Cells(Rows.Count, 1).End(xlUp).Row - 1
freshdata = True

' for each line of data
For i = 2 To lastrow
    ' assign each variable to a column
    ticker = Cells(i, 1).Value
    stkvol = Cells(i, 7).Value

   
        
    ' input/add to summary table
    For j = 2 To 100
        ' if ticker is the same
        If Cells(j, 9) = ticker Then
            ' add to stock volume
            Cells(j, 12).Value = Cells(j, 12).Value + stkvol
            'chngclsold = chngcls
            chngcls = Cells(i, 6).Value
            Exit For
        ' if ticker DNE
        ElseIf Cells(j, 9) = "" Then
        
            If freshdata = False Then
                ' calculate yearly/percentage change
                ' year closing price - year opening price
                ' for the previous ticker
                Cells(j - 1, 10).Value = chngcls - chngopn
                'Cells(j - 1, 10).Value = chngclsold - Cells(j - 1, 10).Value
                ' (year closing price - year opening price)/year opening price
                Cells(j - 1, 11).Value = Cells(j - 1, 10).Value / chngopn
                'Cells(j - 1, 11).Value = Cells(j - 1, 10).Value / Cells(j - 1, 11).Value
                chngclsold = chngcls
            End If
        
            
            Cells(j, 9).Value = ticker
            ' opening price at beginning of the year, this way it's
            ' only added once and saved until the next ticker
            chngopn = Cells(i, 3).Value
            Cells(j, 12).Value = stkvol
            
            freshdata = False
            Exit For
        ' if ticker is a different ticker, go to next j
        End If
    Next j

    'MsgBox (cctype & " " & ccnum & " " & ccamt)
    

Next i


' conditional formatting
For j = 2 To 100
    ' positive
    If Cells(j, 10).Value > 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
    ' negative
    ElseIf Cells(j, 10).Value < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
    ' else, do nothing
    End If
    Cells(j, 11) = Format(Cells(j, 11), "Percent")
    
    
Next j

        







End Sub
