Attribute VB_Name = "Declaraciones"
'*****************************************************************************************************
'* Nombre Archivo   : DECLARACIONES
'* Tipo             : MODULO
'* Descripcion      : MODULO EN EL QUE SE DEFINEN LA PRINCIPALES VARIABLES A UTILIZARCE EN LA LIBRERIA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 21/09/09
'* VERSION          : 1.0
'*****************************************************************************************************

Option Explicit
Public xAño As Integer                        ' ESPECIFICA EL AÑO DE TRABAJO ACTUAL
Public xCon As New ADODB.Connection           ' VARIABLE QUE ALMACENA LA CONECCIONA ACTUAL
Public xTitulo As String                      ' ALAMACENA EL TITULO PARA LA CLASE
Public xMes As Integer                        ' ESPECIFICA EL MES DE TRABAJO ACTUAL
Public NomEmp As String                       ' ALMACENA EL NOMBRE DE LA EMPRESA
Public NumRUC As String                       ' ALMACENA EL NUMERO DE RUC DE LA EMPRESA
Public AnoTra As String                       ' ESPECIFICA EL AÑO DE TRABAJO ACTUAL
Public CONTABILIZAR As Boolean                ' ESPECIFICA SI SE REALIZARAN PROCESOS CONTABLES CUANDO SE INGRESE UN NUEVO REGISTRO
Public NomSis As String                       ' ALMACENA EL NOMBRE DEL SISTEMA
Public xIdUsuario As Integer                  ' ALMACENA EL ID DEL USUARIO ACTUAL
Public xDeDonde As Integer                    ' ESPECIFICA DESDE DONDE SE INVOCA AL FORMULARIO
Public AP_RUTASY As String                    ' ALMACENA LA RUTA DEL SISTEMA
Public AP_RUTABD As String                    ' ALMACENA LA RUTA DE LA BD
Public AP_RUTABM As String                    ' ALMACENA LA RUTA DE LOS ARCHIVOS DE GRAFICO
Public AP_AÑODAT As String
Public AP_MESTRA As Integer                   ' ALMACENA EL MES DE TRABAJO ACTUAL
Public xValidarStckVenta As Integer           'ESPECIFICA SI SE VALIDA EL STOCK AL MOMENTO DE EMITIR LA FACTURACION 0=NO VALIDA, PERMITE EL REGISTRO; -1=VALIDA EL STOCK

Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)

'*****************************************************************************************************
'* Nombre           : RellenaNumdoc
'* Tipo             : FUNCION
'* Descripcion      : FORMATEAR EL NUMERO DE SERIE Y EL NUMERO DE DOCUMENTOS RELLENANDOLO DE CEROS,
'*                    ESTA FUNCION DEVUELVE UNA CADENA QUE ES EL NUMERO DE DOCUMENTO
'* Paranetros       : NOMBRE    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    xnumser   |  LONG      |  ESPECIFICA EL NUMERO DE SERIE DEL DOCUMENTO
'*                    xnumdoc   |  LONG      |  ESPECIFICA EL NUMERO DEL DOCUMENTO
'* Devuelve         : STRING
'*****************************************************************************************************
Function RellenaNumdoc(xnumser As Long, xnumdoc As Long) As String
    RellenaNumdoc = Format(xnumser, "0000") & "-" & Format(xnumdoc, "0000000000")
End Function

'*****************************************************************************************************
'* Nombre           : ActualizaNroDocumento
'* Tipo             : FUNCION
'* Descripcion      : ACTUALIZA EL NUMERO ACTUAL DEL DOCUMENTO QUE SE ESTA EMITIENDO
'* Paranetros       : NOMBRE    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    xnumdoc   |  LONG      |  ESPECIFICA EL NUMERO DE DOCUMENTO
'*                    xtipdoc   |  LONG      |  ESPECIFICA EL TIPO DE DOCUMENTO
'*                    xSerie    |  LONG      |  ESPECIFICA EL NUMERO DE SERIE DEL DOCUMENTO
'* Devuelve         :
'* Observaciones    : ELIMINAR ESTE PROCEDIMIENTO, YA QUE NO SE APLICA EL PROCESO DE ACTUALIZAR EL
'*                    NUMERO DE DOCUMENTO EN LA TABLA mae_series
'*****************************************************************************************************
Sub ActualizaNroDocumento(xnumdoc As Long, xtipdoc As Long, xSerie As Long)
    Dim Rst As New ADODB.Recordset
    Dim NumDoc As String

    RST_Busq Rst, "SELECT * from mae_Series where iddoc = " & xtipdoc & " and numser = " & xSerie & "" _
        & " ORDER BY numser,numdoc ", xCon

    If Rst.RecordCount = 0 Then
        Rst.AddNew
        Rst("numdoc") = xnumdoc
        Rst.Update
    Else
        Rst("numdoc") = xnumdoc
        Rst.Update
    End If

    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : CargaDatosEmpresa
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA EMPRESA: Nombre de la empresa, Numero de Ruc, Año de trabajo
'*                    TAMBIEN CARGA LOS SIGUIENTES DATOS DEL SISTEMA: Nombre del sistema, Ruta de la
'*                    base de datos, Ruta del sistema, Ruta de los archivo de grafico, ADEMAS ESPECIFICA
'*                    SI EL SISTEMA HARA PROCESOS CONTABLES
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargaDatosEmpresa()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = Rst("nomemp")
    NumRUC = Rst("numruc")
    CONTABILIZAR = Rst("procon")
    AnoTra = Rst("anotra")
    xValidarStckVenta = NulosN(Rst("stckvta"))
    xAño = AnoTra
    Set Rst = Nothing
    
    NomSis = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub

'*****************************************************************************************************
'* Nombre           : AbriConeccion2
'* Tipo             : FUNCION
'* Descripcion      : DEVUELVE LA CONECCION A LA BASE DE DATOS planillas.mdb
'* Paranetros       : NOMBRE               |  TIPO              |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    ConeccionDataActual  |  ADODB.Connection  |  ESPECIFICA LA CONECCION ACTUAL
'* Devuelve         : ADODB.Connection
'*****************************************************************************************************
Function AbriConeccion2(ConeccionDataActual As ADODB.Connection) As ADODB.Connection
    Dim xCon2 As New ADODB.Connection
    Dim RutaBD As String
    Dim Rst As New ADODB.Recordset
    
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    
    RST_Busq Rst, "SELECT * FROM mae_empresa", ConeccionDataActual
    RutaBD = AP_RUTABD & Rst("ruta") & "planillas.mdb"
    
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCad As String
    Dim xRutaData As String
    Dim xRst As New ADODB.Recordset
    
    xTitulo = "Sistema Gestion Informatica"
    
    xFun.F_BASEDATOS = RutaBD
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set xCon2 = xFun.AbrirConeccion
    Set xFun = Nothing
    Set AbriConeccion2 = xCon2
End Function
