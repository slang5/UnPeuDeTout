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
