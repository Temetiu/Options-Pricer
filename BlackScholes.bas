Attribute VB_Name = "BlackScholes"
Option Explicit


Function d1(S0 As Double, T As Double, K As Double, r As Double, sigma As Double) As Double
    d1 = (Math.Log(S0 / K) + (r + sigma * sigma / 2) * T) / (sigma * Math.Sqr(T))
End Function

Function d2(d1 As Double, T As Double, sigma As Double) As Double
    d2 = d1 - sigma * Math.Sqr(T)
End Function

Function BlackScholes(S0 As Double, T As Double, K As Double, r As Double, sigma As Double, CallPut As Integer)
    'Calcul d'un Call si CallPut = 1, calcul d'un Put sinon
    If CallPut = 1 Then
        BlackScholes = S0 * WorksheetFunction.NormSDist(d1(S0, T, K, r, sigma)) - K * Math.Exp(-r * T) * WorksheetFunction.NormSDist(d2(d1(S0, T, K, r, sigma), T, sigma))
    Else
        BlackScholes = -S0 * WorksheetFunction.NormSDist(-d1(S0, T, K, r, sigma)) + K * Math.Exp(-r * T) * WorksheetFunction.NormSDist(-d2(d1(S0, T, K, r, sigma), T, sigma))
    End If
End Function


'=============================
'|   CALCUL DES GRECQUES     |
'=============================

Function Delta(S0 As Double, T As Double, K As Double, r As Double, sigma As Double, CallPut As Integer) As Double
    'Si Call
    If CallPut = 1 Then
        Delta = WorksheetFunction.NormSDist(d1(S0, T, K, r, sigma))
    'Sinon Put
    Else
        Delta = WorksheetFunction.NormSDist(d1(S0, T, K, r, sigma)) - 1
    End If
End Function


Function Gamma(S0 As Double, T As Double, K As Double, r As Double, sigma As Double) As Double
    Gamma = WorksheetFunction.NormSDist(d1(S0, T, K, r, sigma)) / (S0 * sigma * Math.Sqr(T))
End Function


Function Theta(S0 As Double, T As Double, K As Double, r As Double, sigma As Double, CallPut As Integer) As Double
    If CallPut = 1 Then
        Theta = -S0 * WorksheetFunction.NormSDist(d1(S0, T, K, r, sigma)) * sigma / 2 * Math.Sqr(T) - r * K * Math.Exp(-r * T) * WorksheetFunction.NormSDist(d2(d1(S0, T, K, r, sigma), T, sigma))
    Else
        Theta = -S0 * WorksheetFunction.NormSDist(d1(S0, T, K, r, sigma)) * sigma / 2 * Math.Sqr(T) - r * K * Math.Exp(-r * T) * WorksheetFunction.NormSDist(-d2(d1(S0, T, K, r, sigma), T, sigma))
    End If
End Function


Function Vega(S0 As Double, T As Double, K As Double, r As Double, sigma As Double) As Double
    Vega = S0 * Math.Sqr(T) * WorksheetFunction.NormSDist(d1(S0, T, K, r, sigma))
End Function


Function Rho(S0 As Double, T As Double, K As Double, r As Double, sigma As Double, CallPut As Integer)
    If CallPut = 1 Then
        Rho = T * K * Math.Exp(-r * T) * WorksheetFunction.NormSDist(d2(d1(S0, T, K, r, sigma), T, sigma))
    Else
        Rho = -T * K * Math.Exp(-r * T) * WorksheetFunction.NormSDist(-d2(d1(S0, T, K, r, sigma), T, sigma))
    End If
End Function








