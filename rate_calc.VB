Function PresentFuture(r As Double, n As Double) As Double
    PresentFuture = 1 / (1 + r) ^ (n)
End Function
Function FuturePresent(r As Double, n As Double) As Double
    FuturePresent = (1 + r) ^ (n)
End Function
Function PresentAnnual(r As Double, n As Double) As Double
    If (r = 0) Then
        PresentAnnual = n
    Else
        PresentAnnual = ((1 + r) ^ n - 1) / (r * (1 + r) ^ (n))
    End If
End Function
Function RateAnnual(ann As Double, n As Double) As Double
    Dim r As Double, a As Double, b As Double
    Dim fr As Double, fa As Double, fb As Double
    
    a = -0.2
    b = 0.5
    fr = 1.5
    Dim iter As Integer
    iter = 1
    While ((Abs(fr) > 0.00001) And (iter < 50))
        iter = iter + 1
        r = 0.5 * (a + b)
        fr = PresentAnnual(r, n) * ann - 1
        fa = PresentAnnual(a, n) * ann - 1
        fb = PresentAnnual(b, n) * ann - 1
        If (fa * fr < 0) Then
            b = r
        Else
            a = r
        End If
    Wend
    RateAnnual = r
End Function
Function FutureAnnual(r As Double, n As Double) As Double
    FutureAnnual = ((1 + r) ^ n - 1) / r
End Function
Function PrincipleOverTime(r As Double, n As Double, y As Double) As Double
    Dim ann As Double
    ann = 1# / PresentAnnual(r, n)
    PrincipleOverTime = (1 + r) ^ y - ann * FutureAnnual(r, y)
End Function
Function PointIRR(PV As Double, r As Double, n As Double, q As Double, y As Double) As Double
    Dim A1 As Double
    Dim A2 As Double
    Dim P1 As Double
    Dim P2 As Double
    Dim iter As Integer
    Dim a As Double, irr As Double, b As Double
    Dim fa As Double, firr As Double, fb As Double
    
    A1 = 1 / PresentAnnual(r, n)
    A2 = 1 / PresentAnnual(q, n)
    P1 = PrincipleOverTime(r, n, y)
    P2 = PrincipleOverTime(q, n, y)
    
    a = -0.5
    b = 0.5
    firr = 1.5
    
    iter = 1
    While ((Abs(firr) > 0.00001) And (iter < 50))
        iter = iter + 1
        irr = 0.5 * (a + b)
        firr = (A1 - A2) * PresentAnnual(irr, y) + (P1 - P2) / (1 + irr) ^ y - PV
        fa = (A1 - A2) * PresentAnnual(a, y) + (P1 - P2) / (1 + a) ^ y - PV
        fb = (A1 - A2) * PresentAnnual(b, y) + (P1 - P2) / (1 + b) ^ y - PV
        
        If (fa * firr < 0) Then
            b = irr
        Else
            a = irr
        End If
    Wend
    PointIRR = irr
        
    
End Function
