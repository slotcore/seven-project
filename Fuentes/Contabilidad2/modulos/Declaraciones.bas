Attribute VB_Name = "Declaraciones"
'*****************************************************************************************************
'* Nombre Archivo   : DECLARACIONES.BAS
'* Tipo             : MODULO
'* Descripcion      : MODULO DONDE SE DECLARAN LAS VARIABLES PUBLICAS QUE SE UTILIZARAN EN LA CLASE
'*                    ASI COMO LA DEFINICION DE ALGUNAS FUNCIONES PROPIAS DE LA CLASE
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 26/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public xCon As New ADODB.Connection
Public xTitulo As String

Public NomEmp As String                ' ALMACENA EL NOMBRE DE LA EMPRESA
Public NumRUC As String                ' ALAMCENA EL NUMERO DE RUC DE LA EMPRESA
Public TIPOPERSONA As String           ' especifica el tipo de persona que es la empresa que se esta llevando
Public TIPODOCUMENTOIDEN As String     ' especifiva el tipo de documento de indentidad de la empresa que se esta llevando
Public CONTABILIZAR As Boolean         ' ESPECIFICA SI SE REALIZARAN LOS PROCESOS CONTABLES
Public xMes As Integer                 ' ESPECIFICA EL MES DE TRABAJO ACTUAL
Public AnoTra As String                ' ESPECICA EL AÑO DE TRABAJO
Public xIdEmpresa As Integer           ' ESPECIFICA EL ID DE LA EMPRESA
Public MostrarValorizado  As Boolean   '

Public AP_RUTASY As String             ' ALMACENA EL LA RUTA DEL SISTEMA
Public AP_RUTABD As String             ' ALAMACENA LA RUTA DE LA BASE DE DATOS
Public AP_RUTABM As String             ' ALAMACENA LA RUTA DE LOS ARCHIVOS BMP
Public AP_AÑODAT As String             ' ESPECIFICA EL AÑO DE TRABAJO ACTUAL
Public AP_MESTRA As Integer            ' ESPECIFICA EL MES DE TRABAJO ACTUAL
Public NomSis As String                ' ESPECIFICA EL NOMBRE DEL SISTEMA

Public xIdUsuario As Integer           'ESPECIFICA EL CODIGO DEL USUARIO QUE INGRESA AL SISTEMA
Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)


'*****************************************************************************************************
'* Nombre           : CargaDatosEmpresa
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA EMPRESA DE TRABAJO ACTUAL Y LO ALMACENA EN VARIABLES
'*                    GLOBALES
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargaDatosEmpresa()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT mae_empresa.*, mae_tipoempresa.codsun AS tipempcodsun, mae_dociden.codsun AS docidencodsun FROM mae_tipoempresa " _
        & " RIGHT JOIN (mae_dociden RIGHT JOIN mae_empresa ON mae_dociden.id = mae_empresa.idtipdoc) ON mae_tipoempresa.id = mae_empresa.idtipper", xCon

    NomEmp = Rst("nomemp")
    NumRUC = Rst("numruc")
    CONTABILIZAR = Rst("procon")
    TIPODOCUMENTOIDEN = Rst("docidencodsun")
    TIPOPERSONA = Rst("tipempcodsun")
    AnoTra = Rst("anotra")
    Set Rst = Nothing
        
    NomSis = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub

'*****************************************************************************************************
'* Nombre           : AbrirConecciones
'* Tipo             : FUNCCION
'* Descripcion      : ABRE LA CONECCION A UNA BASE DE DATOS, ESTA FUNCION DEVUELVE UNA VARIABLE DE
'*                    CONECCION ABIERTA
'* Paranetros       : NOMBRE    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Ruta      |  String    |  ESPECIFICA LA RUTA DE LA BASE DE DATOS QUE SE DESEA
'*                                              ABRIR
'* Devuelve         : ADODB.Connection
'*****************************************************************************************************
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

