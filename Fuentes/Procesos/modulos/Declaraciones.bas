Attribute VB_Name = "Declaraciones"
Option Explicit

Public xCon As New ADODB.Connection
Public xTitulo As String

Public NomSIS As String
Public NomEmp As String
Public NumRUC As String
Public CONTABILIZAR As Boolean
Public xMes As Integer
Public AnoTra As String
Public AP_RUTABD As String
Public AP_RUTASY As String
Public AP_RUTABM As String

Sub CargaDatosEmpresa()
    Dim rst As New ADODB.Recordset
    
    RST_Busq rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = rst("nomemp")
    NumRUC = rst("numruc")
    CONTABILIZAR = rst("procon")
    AnoTra = rst("anotra")
    Set rst = Nothing
End Sub


