Attribute VB_Name = "declaraciones"
Option Explicit

Public xCon As New ADODB.Connection
Public xTitulo As String

Public NomEmp As String
Public NumRUC As String
Public Nomsis As String
Public AnoTra As String
Public xMes As Integer
Public xIdUsuario As Integer

Public CONTABILIZAR As Boolean

Global AP_RUTASY As String
Global AP_RUTABD As String
Global AP_RUTABM As String

Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)


Sub CargaDatosEmpresa(xCon As ADODB.Connection)
    On Error Resume Next
    Dim rst As New ADODB.Recordset
    RST_Busq rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = NulosC(rst("nomemp"))
    NumRUC = NulosC(rst("numruc"))
    CONTABILIZAR = NulosN(rst("procon"))
    AnoTra = NulosC(rst("anotra"))
    
    Set rst = Nothing
    Err.Clear
End Sub
