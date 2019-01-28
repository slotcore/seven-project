Attribute VB_Name = "Declaraciones"
Public xCon As New ADODB.Connection
Public xTitulo As String

Public NomEmp As String
Public NumRUC As String
Public xMes As Integer
Public AnoTra As String
Public CONTABILIZAR As Boolean
Global AP_RUTDATTRA As String  ' Especifica la ruta de datos de la empresa actual

Public AP_NOMSIS As String

Public AP_RUTASY As String
Public AP_RUTABD As String
Public AP_RUTABM As String
Public AP_AÑODAT As String
Public AP_MESTRA As Integer


Public MostrarValorizado As Boolean

Sub CargaDatosEmpresa()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = Rst("nomemp")
    NumRUC = Rst("numruc")
    CONTABILIZAR = Rst("procon")
    AnoTra = Rst("anotra")
    Set Rst = Nothing
End Sub

Sub Main()
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCad As String
    Dim xRutaData As String
    Dim xRst As New ADODB.Recordset
    
    xTitulo = "Sistema Gestion Informatica"
    xRutaData = "C:\seven\data\2016\0031\data.mdb"
    'xRutaData = "J:\seven\data\2015\0031\data.mdb"
    'xRutaData = "C:\seven\data\2015\0031\data.mdb"
    'xRutaData = "J:\seven\data\2012\0002\data.mdb"
    'xRutaData = "J:\seven\data\2012\0001\data.mdb"
    
    xFun.F_BASEDATOS = xRutaData
    xFun.F_GRUPOTRABAJO = "C:\seven\seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set xCon = xFun.AbrirConeccion
    Set xFun = Nothing
    CargaDatosEmpresa
End Sub

