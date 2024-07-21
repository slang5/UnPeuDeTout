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


Private Sub Autocompletion(ByVal Target As Range)
    ' Define the specific cell you want to monitor
    Dim MonitoredCell As Range
    Set MonitoredCell = Me.Range("G2")
    
    ' Check if the changed cell is the one we're monitoring
    If Not Application.Intersect(MonitoredCell, Target) Is Nothing Then
        ' Disable events to prevent infinite loops when changing cells
        Application.EnableEvents = False
        
        ' Perform the XLookup and update G3
        Me.Range("G3").Value = Application.WorksheetFunction.XLookup(Me.Range("G2").Value, _
                                                                      Me.Range("A2:A23"), _
                                                                      Me.Range("B2:B23"), _
                                                                      "_RIEN_", _
                                                                      0, _
                                                                      1)
        ' Re-enable events after the operation
        Application.EnableEvents = True
    End If
End Sub



Private Sub Worksheet_Change(ByVal Target As Range)
    Dim duree As Double

    ' Check if the changed cell is within the monitored ranges
    If Not Intersect(Target, Me.Range("A7:B7")) Is Nothing Or Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        
        ' Calculate the duration
        duree = (Me.Cells(7, 2).Value - Me.Cells(7, 1).Value) * 90 / 100
        
        ' Check for negative duration
        If duree < 0 Then
            MsgBox "Periode du graphe negative"
            Exit Sub
        ElseIf duree <= ActiveSheet.Cells(3, 2).Value * 30 Then
            ' Set the cell color to red
            ActiveSheet.Cells(7, 1).Interior.Color = RGB(255, 0, 0)
            ActiveSheet.Cells(7, 2).Interior.Color = RGB(255, 0, 0)
        Else
            ' Reset the cell color if the condition is not met
            ActiveSheet.Cells(7, 1).Interior.Color = RGB(255, 255, 0)
            ActiveSheet.Cells(7, 2).Interior.Color = RGB(255, 255, 0)
        End If
    End If
End Sub
