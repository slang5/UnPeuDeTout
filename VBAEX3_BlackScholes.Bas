Function Black_Scholes(Spot, Strike, Time, Vol, TauxSR, Instrument)
    Dim D1, D2, C As Double
    
    D1 = (WorksheetFunction.Ln(Spot / Strike) + (TauxSR + Vol * Vol / 2) * Time) / (Vol * Sqr(Time))
    D2 = D1 - Vol * Sqr(Time)
    
    If (Instrument) = "Call" Then
        C = Spot * WorksheetFunction.Norm_S_Dist(D1, True) - WorksheetFunction.Norm_S_Dist(D2, True) * Strike * Exp(-1 * TauxSR * Time)
    Else
        C = WorksheetFunction.Norm_S_Dist(-D2, True) * Strike * Exp(-TauxSR * (Time)) - Spot * WorksheetFunction.Norm_S_Dist(-D1, True)
    End If
    
    Black_Scholes = C
End Function
