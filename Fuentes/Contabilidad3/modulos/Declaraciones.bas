Attribute VB_Name = "Declaraciones"
Option Explicit

Public xCon As New ADODB.Connection
Public xTitulo As String

Public NomEmp As String
Public NumRUC As String
Public TIPOPERSONA As String        'especifica el tipo de persona que es la empresa que se esta llevando
Public TIPODOCUMENTOIDEN As String  'especifiva el tipo de documento de indentidad de la empresa que se esta llevando
Public CONTABILIZAR As Boolean
Public xMes As Integer
Public AnoTra As String
Public xIdEmpresa As Integer
Public MostrarValorizado  As Boolean

Public AP_RUTASY As String
Public AP_RUTABD As String
Public AP_RUTABM As String
Public AP_AÑODAT As String
Public AP_MESTRA As Integer
Public NomSis As String

Public xIdUsuario As Integer
Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)


Sub CargaDatosEmpresa()
    Dim rst As New ADODB.Recordset
    
    RST_Busq rst, "SELECT mae_empresa.*, mae_tipoempresa.codsun AS tipempcodsun, mae_dociden.codsun AS docidencodsun FROM mae_tipoempresa " _
        & " RIGHT JOIN (mae_dociden RIGHT JOIN mae_empresa ON mae_dociden.id = mae_empresa.idtipdoc) ON mae_tipoempresa.id = mae_empresa.idtipper", xCon

    NomEmp = rst("nomemp")
    NumRUC = rst("numruc")
    CONTABILIZAR = rst("procon")
    TIPODOCUMENTOIDEN = rst("docidencodsun")
    TIPOPERSONA = rst("tipempcodsun")
    AnoTra = rst("anotra")
    Set rst = Nothing
    
    
    NomSis = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub

Function AbrirConecciones(Ruta As String) As ADODB.Connection
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCone As ADODB.Connection
    
    xFun.F_BASEDATOS = Ruta
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set xCone = xFun.AbrirConeccion
    Set xFun = Nothing
    Set AbrirConecciones = xCone
End Function

