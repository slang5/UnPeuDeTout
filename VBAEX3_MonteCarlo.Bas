Function Monte_Carlo(Spot As Double, Strike As Double, Time As Double, Vol As Double, TauxSR As Double, Instrument As String, Rep As Integer) As Double

    Dim Prix, Norm, St, Gain As Double: Prix = 0: Norm = 0: St = 0: Gain = 0
    
    For I = 1 To Rep
        Norm = WorksheetFunction.Norm_S_Inv(Rnd())
        St = Spot * Exp((TauxSR - Vol * Vol / 2) * Time + (Norm * Vol * Sqr(Time)))
        If Instrument = "Call" Then
            Gain = Gain + WorksheetFunction.Max(St - Strike, 0) * Exp(-TauxSR * Time)
        ElseIf Instrument = "Put" Then
            Gain = Gain + WorksheetFunction.Max(Strike - St, 0) * Exp(-TauxSR * Time)
        End If
    Next I
    
    Prix = Gain / Rep
    Monte_Carlo = Prix
End Function
