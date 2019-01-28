VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncionesCad"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Eps_Pass() As String
    Eps_Pass = "010419762005"
End Function

Function Eps_User() As String
    Eps_User = "cav2005sialp"
End Function


Public Function NulosN(valor_nulo As Variant) As Double
    NulosN = F_NulosN(valor_nulo)
End Function

Public Function NulosC(valor_nulo As Variant) As String
    NulosC = F_NulosC(valor_nulo)
End Function

Function RellenarBlancos(Cadena As String, AnchoCadena As Integer, Optional Lado As Integer) As String
    '1 = Izquierda
    '2 = derecha
    Dim B As Integer
    Dim xCadena As String
    Dim xLongCadena As Integer
    xLongCadena = AnchoCadena - Len(Trim(Cadena))
    
    xCadena = Cadena
    If Lado = 2 Then 'rellena a la derecha
        For B = 1 To xLongCadena
            xCadena = " " + xCadena
        Next B
    Else
        For B = 1 To xLongCadena
            xCadena = xCadena + " "
        Next B
    End If
    RellenarBlancos = xCadena
End Function

