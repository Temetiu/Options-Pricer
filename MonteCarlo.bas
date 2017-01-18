Attribute VB_Name = "MonteCarlo"
Option Explicit

Function MonteCarlo(S0 As Double, T As Double, K As Double, r As Double, sigma As Double, q As Double, n As Double, nbSimulation As Double, CallPut As Integer) As Variant
Dim dT, e, dlns, price, Simulation(), Payoff() As Double
Dim i, j, a, p As Integer
ReDim Simulation(nbSimulation, n + 1)
ReDim Payoff(nbSimulation)
dT = T / n
a = 0
For i = 1 To nbSimulation
    Simulation(i + a, 1) = S0
    Randomize
    p = 0
    For j = 1 To n
        If (j - 1) / 5000 - Int((j - 1) / 5000) = 0 And j > 1 Then p = p + 1
        e = WorksheetFunction.NormSInv(Rnd())
        dlns = (r - q - sigma ^ 2 / 2) * dT + sigma * e * dT ^ 0.5
        
        If j - 5000 * p = 1 And p > 0 Then
            Simulation(i + p + a, 2) = Simulation(i + p - 1 + a, 5001) * Exp(dlns)
        Else
            Simulation(i + p + a, j - 5000 * p + 1) = Simulation(i + p + a, j - 5000 * p) * Exp(dlns)
        End If
    Next j
    If CallPut = 1 Then
        Payoff(i) = WorksheetFunction.Max(Simulation(i + p + a, j - 5000 * p) - K, 0) * Exp(-r * T)
    Else
        Payoff(i) = WorksheetFunction.Max(K - Simulation(i + p + a, j - 5000 * p), 0) * Exp(-r * T)
    End If
    a = a + p
Next i
price = 0
For i = 1 To nbSimulation
    price = price + Payoff(i)
Next i
price = price / nbSimulation
MonteCarlo = price
End Function

