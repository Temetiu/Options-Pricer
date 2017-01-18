VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pricer 
   Caption         =   "UserForm1"
   ClientHeight    =   9858.001
   ClientLeft      =   90
   ClientTop       =   408
   ClientWidth     =   9534.001
   OleObjectBlob   =   "Pricer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Pricer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BlackScholesMethod_Click()
    DividendYield.Enabled = False
    nbPeriode.Enabled = False
    nbSimulation.Enabled = False
End Sub

Private Sub MonteCarloMethod_Click()
    DividendYield.Enabled = True
    nbPeriode.Enabled = True
    nbSimulation.Enabled = True
End Sub

Private Sub ArbreBinomMethod_Click()
    DividendYield.Enabled = True
    nbPeriode.Enabled = True
    nbSimulation.Enabled = False
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim Mois As String
    
    For i = 1 To 31
        cbJour.AddItem i
    Next i
    
    ' Création d'un tableau des noms de mois
    For i = 1 To 12
        Mois = Format(DateSerial(1, i, 1), "mmmm")
        cbMois.AddItem Mois
    Next i
End Sub


Private Sub GetPrice_Click()
    Dim RF As Double
    Dim Vola As Double
    Dim Mois As Integer
    Dim T As Double
    Dim Dividende As Double
    
    If ztCours.Value <= 0 Or ztStrike.Value <= 0 Or ztRF.Value <= 0 Or ztVola.Value <= 0 Or ztRF.Value > 100 Or ztVola.Value > 100 Then
        MsgBox ("Erreur: Vérifier les valeurs")
    ElseIf nbPeriode.Value > 1000 Then
        MsgBox ("Erreur: Nombre de périodes trop grand")
    ElseIf nbSimulation.Value > 5000 Then
        MsgBox ("Erreur: Nombre de simulations trop élevé")
    Else
        'Conversion du mois en nombre
        Mois = Month(DateValue("01 " & cbMois.Value))
    
        'Calcul du nombre dannée avant la date d'échéance
        T = Abs(DateDiff("d", Now, DateSerial(ztAnnee.Value, Mois, cbJour.Value)) / 365)
    
        RF = ztRF.Value / 100
        Vola = ztVola.Value / 100
        Dividende = DividendYield.Value / 100
      
        'CALCUL PAR LA METHODE BLACK-SCHOLES
        '------------------------------------------------------------------
        If BlackScholesMethod.Value = True Then
            If TypeCall.Value = True Then
                Prix.Value = BlackScholes.BlackScholes(ztCours.Value, T, ztStrike.Value, RF, Vola, 1)
            
                'Mise à jour des grecques
                DeltaVal.Value = BlackScholes.Delta(ztCours.Value, T, ztStrike.Value, RF, Vola, 1)
                GammaVal.Value = BlackScholes.Gamma(ztCours.Value, T, ztStrike.Value, RF, Vola)
                ThetaVal.Value = BlackScholes.Theta(ztCours.Value, T, ztStrike.Value, RF, Vola, 1)
                VegaVal.Value = BlackScholes.Vega(ztCours.Value, T, ztStrike.Value, RF, Vola)
                RhoVal.Value = BlackScholes.Rho(ztCours.Value, T, ztStrike.Value, RF, Vola, 1)
            
            ElseIf TypePut.Value = True Then
                Prix.Value = BlackScholes.BlackScholes(ztCours.Value, T, ztStrike.Value, RF, Vola, 0)
            
                DeltaVal.Value = BlackScholes.Delta(ztCours.Value, T, ztStrike.Value, RF, Vola, 0)
                GammaVal.Value = BlackScholes.Gamma(ztCours.Value, T, ztStrike.Value, RF, Vola)
                ThetaVal.Value = BlackScholes.Theta(ztCours.Value, T, ztStrike.Value, RF, Vola, 0)
                VegaVal.Value = BlackScholes.Vega(ztCours.Value, T, ztStrike.Value, RF, Vola)
                RhoVal.Value = BlackScholes.Rho(ztCours.Value, T, ztStrike.Value, RF, Vola, 0)
            
            Else
                MsgBox ("Choisir Call ou Put")
            End If
        
        'CALCUL PAR LA METHODE MONTE CARLO
        '------------------------------------------------------------------
        ElseIf MonteCarloMethod.Value = True Then
            DeltaVal.Value = ""
            GammaVal.Value = ""
            ThetaVal.Value = ""
            VegaVal.Value = ""
            RhoVal.Value = ""
            If TypeCall.Value = True Then
                Prix.Value = MonteCarlo.MonteCarlo(ztCours.Value, T, ztStrike.Value, RF, Vola, DividendYield.Value, nbPeriode.Value, nbSimulation.Value, 1)
            ElseIf TypePut.Value = True Then
                Prix.Value = MonteCarlo.MonteCarlo(ztCours.Value, T, ztStrike.Value, RF, Vola, DividendYield.Value, nbPeriode.Value, nbSimulation.Value, 0)
            Else
                MsgBox ("Choisir Call ou Put")
            End If
        
    
        'CALCUL PAR LA METHODE DES ARBRES BINOMIAUX
        '------------------------------------------------------------------
        ElseIf ArbreBinomMethod.Value = True Then
            DeltaVal.Value = ""
            GammaVal.Value = ""
            ThetaVal.Value = ""
            VegaVal.Value = ""
            RhoVal.Value = ""
            If TypeCall.Value = True Then
                Prix.Value = ArbresBinomiaux.Binomial(ztCours.Value, T, ztStrike.Value, RF, Vola, Dividende, nbPeriode.Value, 1)
            ElseIf TypePut.Value = True Then
                Prix.Value = ArbresBinomiaux.Binomial(ztCours.Value, T, ztStrike.Value, RF, Vola, Dividende, nbPeriode.Value, 0)
            Else
                MsgBox ("Choisir Call ou Put")
            End If
        
        Else
            MsgBox ("Selectionnez une méthode de calcul")
        End If
    End If
End Sub


