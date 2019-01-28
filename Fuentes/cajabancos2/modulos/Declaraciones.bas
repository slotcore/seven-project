Attribute VB_Name = "Declaraciones"
Option Explicit

Public xCon As New ADODB.Connection
Public xTitulo As String
Public NomEmp, NumRuc As String
Public AnoTra As String
Public xMes As Integer
Public xIdUsuario As Integer

Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Declare Sub InitCommonControls Lib "Comctl32" ()

Public T_ToolTipText() As String

Sub CargaDatos()
    Dim Rst As New ADODB.Recordset
    
    rst_busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = Rst("nomemp")
    NumRuc = Rst("numruc")
    AnoTra = Rst("anotra")
    Set Rst = Nothing
End Sub
