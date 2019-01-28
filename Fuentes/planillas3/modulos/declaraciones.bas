Attribute VB_Name = "declaraciones"
Option Explicit

Public xTitulo As String
Public xCon As New ADODB.Connection     'coneccion a la data planilla
'Public xConPri As New ADODB.Connection   'coneccion a la data principal

Public NomEmp, NumRuc, DirEmp As String
Public xMes As Integer
Public AnoTra As String
Public xIdUsuario As Integer

Public AP_RUTASY As String
Public AP_RUTABD As String
Public AP_RUTABM As String
Public AP_AÑODAT As String
Public AP_MESTRA As Integer
Global AP_RUTDATTRA As String   ' Especifica la ruta de datos de la empresa actual
Public NomSis As String

Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)



Sub CargaDatos()
    Dim Rst As New ADODB.Recordset
    
    'RST_Busq Rst, "SELECT * FROM mae_empresa", xConPri
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = NulosC(Rst("nomemp"))
    NumRuc = NulosC(Rst("numruc"))
    AnoTra = NulosC(Rst("anotra"))
    DirEmp = NulosC(Rst("diremp"))
    Set Rst = Nothing
    
    NomSis = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub

Function AbrirConPlanilla(Ruta As String) As ADODB.Connection
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCone As ADODB.Connection
    
    xFun.F_BASEDATOS = NulosC(AP_RUTABD) + Mid(NulosC(AP_RUTDATTRA), 1, Len(AP_RUTDATTRA) - 8) + "planillas.mdb"
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set xCone = xFun.AbrirConeccion
    Set xFun = Nothing
    Set AbrirConPlanilla = xCone
End Function

