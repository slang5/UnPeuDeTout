Sub BeauGraph()
    Dim ws As Worksheet
    Dim Chart As ChartObject
    Dim minYValue As Double
    Dim maxYValue As Double
    
    Set ws = ActiveSheet
    
    minYValue = ws.Range("F1").Value
    maxYValue = ws.Range("F2").Value
    
    For Each Chart In ws.ChartObjects
        With Chart.Chart.Axes(xlValue)
            .MinimumScale = minYValue
            .MaximumScale = maxYValue
        End With
    Next Chart
End Sub


Function TexteAvantChaineDeCaracteres(Inpt As String, Chaine As String) As String
    Dim position As Integer
    
    position = InStr(Inpt, Chaine)
    
    If position > 0 Then
        TexteAvantChaineDeCaracteres = Left(Inpt, position - 1)
    Else

        TexteAvantChaineDeCaracteres = ""
    End If
End Function

Function TexteApresChaineDeCaracteres(Inpt As String, Chaine As String) As String
    Dim position As Integer
    
    position = InStr(Inpt, Chaine)
    
    If position > 0 Then
        TexteApresChaineDeCaracteres = Right(Inpt, Len(Inpt) - position)
    Else

        TexteApresChaineDeCaracteres = ""
    End If
End Function


Sub MergeSameTextInColumnA()
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim currentRow As Long
    Dim textValue As String
    
    Set ws = ActiveSheet
    startRow = 1
    currentRow = 1
    textValue = ws.Cells(currentRow, 1).Value
    

    Do While ws.Cells(currentRow, 1).Value <> ""
        If ws.Cells(currentRow, 1).Value <> textValue Then
        
            If currentRow - startRow > 1 Then
                ws.Range(ws.Cells(startRow, 1), ws.Cells(currentRow - 1, 1)).Merge
            End If
            

            textValue = ws.Cells(currentRow, 1).Value
            startRow = currentRow
        End If
        currentRow = currentRow + 1
    Loop
    

    If currentRow - startRow > 1 Then
        ws.Range(ws.Cells(startRow, 1), ws.Cells(currentRow - 1, 1)).Merge
    End If
    Application.DisplayAlerts = True
End Sub

Function FRENGeasy(binaire As Integer, textFR As String, textEN As String)
    If binaire = 1 Then
        FRENGeasy = textFR
    ElseIf binaire = 2 Then
        FRENGeasy = textEN
    Else
        FRENGeasy = "#N/A"
    End If
       
End Function
