Attribute VB_Name = "ArbresBinomiaux"
Option Explicit

Function Binomial(S0 As Double, T As Double, K As Double, r As Double, sigma As Double, Dividende As Double, n As Integer, CallPut As Integer)
    Dim u As Double     'Coeff de hausse
    Dim d As Double     'Coeff de baisse
    Dim dT As Double    'Durée d'une période
    Dim p As Double     'Probabilité risque neutre de hausse du sous-jacent
    Dim q As Double     'Proba complémentaire à p
    Dim somme As Double
    dT = T / n
    u = Math.Exp(sigma * Math.Sqr(dT))
    d = 1 / u
    p = (Math.Exp((r - Dividende) * dT) - d) / (u - d)
    q = 1 - p
    somme = 0
    
    Dim i As Long
    
    If CallPut = 1 Then
        For i = 0 To n
            somme = somme + WorksheetFunction.Combin(n, i) * p ^ i * q ^ (n - i) * WorksheetFunction.Max(S0 * u ^ i * d ^ (n - i) - K, 0)
        Next
    Else
        For i = 0 To n
            somme = somme + WorksheetFunction.Combin(n, i) * p ^ i * q ^ (n - i) * WorksheetFunction.Max(K - S0 * u ^ i * d ^ (n - i), 0)
        Next
    End If
    Binomial = Math.Exp(-r * T) * somme
End Function
