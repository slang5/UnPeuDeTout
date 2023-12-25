Function BinomialTree(Spot, Strike, Time, Vol, TauxSR, Instrument, Netapes) As Double
    Dim Up As Double, Deltatime As Double, Down As Double, A As Double, P As Double, Gain As Double: Gain = 0
    Dim N As Integer
    Dim Binomiale As Double
    Deltatime = Time / Netapes
    Up = Exp(Vol * Sqr(Deltatime))
    Down = Exp(-Vol * Sqr(Deltatime))
    A = Exp(TauxSR * Deltatime)
    N = Netapes
    P = (A - Down) / (Up - Down)
    Dim St As Double

    If N > 0 Then
        Dim I As Integer
        Dim Temp As Double
        For I = 0 To N
            St = Spot * (Up ^ I) * (Down ^ (N - I))
            If Instrument = "Call" Then
                Temp = WorksheetFunction.Max(0, St - Strike)
            Else
                Temp = WorksheetFunction.Max(0, Strike - St)
            End If
            If Temp > 0 Then
                Binomiale = WorksheetFunction.BinomDist(I, N, P, False)
                Gain = Gain + Binomiale * Temp
            End If
            
        Next I
        
        End If

    Gain = Gain * Exp(-TauxSR * Time)
    
    BinomialTree = Gain
End Function
