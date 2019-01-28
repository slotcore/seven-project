Attribute VB_Name = "Declaraciones"
Option Explicit

Public Grabo As Boolean

Public xCon As New ADODB.Connection
Public xMes As Integer
Public xTitulo As String

Public NomEmp As String
Public NumRUC As String
Public AnoTra As String
Public Nomsis As String
Public CONTABILIZAR As Boolean

Public xIdFormatos As Integer

Public AP_RUTASY As String
Public AP_RUTABD As String
Public AP_RUTABM As String
Public AP_AÑODAT As String
Public AP_MESTRA As Integer

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'---colocar el formulario sobre otro
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'----------funcion ocultar_boton frm ------------------
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'------------------------------------------------------
' API declares
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long


Sub CargaDatosEmpresa()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = NulosC(Rst("nomemp"))
    NumRUC = NulosC(Rst("numruc"))
    AnoTra = NulosC(Rst("anotra"))
    
    Set Rst = Nothing
    Nomsis = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub
