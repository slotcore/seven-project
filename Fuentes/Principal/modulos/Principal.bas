Attribute VB_Name = "Principal"
'*****************************************************************************************************
'* Nombre Archivo   : PRINCIPAL
'* Tipo             : MODULO
'* Descripcion      : MODULO EN EL QUE SE DEFINEN LA PRINCIPALES VARIABLES A UTILIZARCE EN EL SISTEMA,
'*                    ES EL PRIMER OBJETO EN EJECUTARSE. AQUI ES DONDE SE EFECTUA EL ACCESO A LA BASE
'*                    DE DATOS Y SE EJECUTA EL PRIMER NIVEL DE SEGURIDAD DEL SISTEMA.
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 01/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
'codigo de activacion de accounting 2009
'R2KVT-7WFCX-QQ8MF-WGHQ9-9P8WD
Option Explicit
Public xDataSource As String                ' Almacena el Ruta de la Conexion actual
Public xCon As New ADODB.Connection         ' Coneccion principal para accesar datos
Public xNumRuc As String                    ' Almacena el Numero de RUC de la empresa
Public xNomEmp As String                    ' Almacena el nombre de la empresa
Public AnoTra As Integer                    ' Almacena el año de trabajo actual
Public SeEjecutoEmp As Boolean              ' Especifica si la empresa actual esta accesada
Public xIdEmpresa As Integer                ' Especifica el id de la empresa actual
Public xIdUsuario As Integer                ' Especifica el id del usuario
Public xTitulo As String                    ' Especifica el rotulo que mostraran los cuadros de texto (msgbox)

Public CONTABILIZAR As Boolean              ' Especifica si el sistema hara procesos contables a las operaciones que se realicen
Global AP_RUTASY As String                  ' Almacena la ruta de la aplicacion, lo extrae del archivo ini
Global AP_RUTABD As String                  ' Almacena la ruta de la BD, lo extrae del archivo ini
Global AP_RUTABM As String                  ' Almacena la ruta de los graficos que se usaran en el sistema, lo extrae del archivo ini
Global AP_RUTAAR As String                  ' Almacena la ruta de archivos que se usaran para el sistema, lo extrae del archivo ini
Global AP_AÑODAT As String                  ' Almacena el año de trabajo actual, lo extrae de la base de datos principal
Global AP_NOMSIS As String                  ' Almacena el nombre del sistema, lo extrae del archivo ini
Global AP_MESTRA As Integer                 ' Almacena el mes de trabajo actual, se cargara en funcion a la fecha acual, pudiendo ser actualizada por el usuario a la hora de cambiar el periodo de trabajo
Global AP_RUTDATTRA As String               ' Especifica la ruta de datos de la empresa actual
Global AP_VERSION As String               ' Especifica la ruta de datos de la empresa actual

Public PedirUsuario As Boolean              ' Variable para controlar si se ingreso el usuario para entrar al sistema

Public xConDATA As New ADODB.Connection     'esta coneccion abre la data del antiguo sgi

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal H As Long, ByVal s As String, ByVal I As Integer, d As Any) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'variables publicas para HtmlHelp
Public Const HH_DISPLAY_TOPIC = &H0         ' Muestra la Ayuda
Public Const HH_DISPLAY_TOC = &H1           ' Muestra la Ayuda, concretando TOPIC por nombre
Public Const HH_DISPLAY_INDEX = &H2         ' Muestra el indice de la ayuda
Public Const HH_DISPLAY_SEARCH = &H3        ' Despliega la busqueda de la ayuda
Public Const HH_HELP_CONTEXT = &HF          ' Muestra la Ayuda, concretando TOPIC por su ID de contexto

Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)

'*****************************************************************************************************
'* Nombre Modulo  : Main()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : MODULO PRINCIPAL, ES EL PRIMER CODIGO QUE SE EJECUTARA CUANDO UNICIE LA APLICACION,
'*                  AQUI SE CARGARAN TODAS LA CARIABLES PUBLICAS Y GLOBALES, ASI MISMO SE HARA LA
'*                  LA CONECCION A LA BASE DE DATOS
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Main()
    '----------

    If App.PrevInstance = True Then
        MsgBox "SEVEN ya esta iniciado", vbInformation
        End
    End If

    SeEjecutoEmp = False
    PedirUsuario = True                                                                 ' Le indicamos al sistema que pida usuario
    
    AP_NOMSIS = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")       ' CARGAMOS EL NOMBRE DEL SISTEMA
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")          ' LEEMOS LA RUTA DEL SISTEA
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")          ' LEEMOS LA RUTA DE LA BASE DE DATOS
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")          ' LEEMOS LA RUTA DE LOS ARCHIVOS GRAFICOS QUE USARA EL SISTEMA
    AP_RUTAAR = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTAAR", "RUTAS")          ' LEEMOS LA RUTA DE LOS ARCHIVOS ADICIONALES QUE USARA EL SISTEMA
    
    
    ' Se carga la version del sistema
    AP_VERSION = LeerLineaINI(Trim(App.Path) + "\app.ini", "VERSION", "SOFTWARE")
    
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCad As String
    Dim xRutaData As String
    Dim xRst As New ADODB.Recordset
    
    xTitulo = "Sistema Gestion Informatica"                                             ' CARGAMOS LOS TITULOS PARA LOS MSGBOX Y OTROS MENSAJES QUE EMITIERA EL SISTEMA
        
    xFun.F_BASEDATOS = AP_RUTABD + "data.mdb"                                           ' PASAMOS LA RUTA DE LA BASE DE DATOS PARA ABRIR LA CONECCION
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"                                       ' PASAMOS LA RUTA DEL ARCHIVO DE TRABJO DE LA BASE DE DATOS
    xFun.F_PASSWORD = Eps_Pass                                                          ' PASAMOS EL PASWORD DE LA BASE DE DATOS
    xFun.F_USUARIO = Eps_User                                                           ' PASAMOS EL USUARIO DE LA BASE DE DATOS
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"                                        ' PASAMOS EL NOMBRE DEL PROVEEDORE DE DATOS PARA ADO 2.5
    
    Set xCon = xFun.AbrirConeccion                                                      ' ABRIMOS LA CONECCION DE DATOS
    Set xFun = Nothing

    ' BUSCAMOS QUE LA PC EN LA QUE SE EJECUTA LA APLICACION ESTE REGISTRADA EN LA BASE DE DATOS DEL SISTEMA,
    ' ESTO CON EL FIN DE EVITAR LA REPRODUCCION DE COPIAS ILEGALES, PARA ELLO SE IDENTIFICA EL NUMERO DE SERIE
    ' DEL DISCO DURO DE LA PC
    RST_Busq xRst, "SELECT * FROM mae_pc WHERE serdis = '" & Trim(Str(LeerNumeroDisco("c:"))) & "'", xCon

    If xRst.RecordCount = 0 Then
        'NO SE ENCONTRO EL DISCO REGISTRADO, EJECUTAMOS EL FORMULARIO QUE PIDE EL REGISTRO Y AUTENTIFICACION DEL SOFTWARE
        FrmPrimeraVez.Show
    Else
    
        ' LIMPIAMOS TODAS LAS PC REGISTRADAS EN EL DIA ESPECIFICADO
        
'        If Now() >= CDate("01/04/2011") And Time > TimeValue("12:00:00 a.m.") Then
'            xCon.Execute "DELETE * FROM mae_pc"
'        End If
    
        '--Actualizar nombre de equipo
        xCon.Execute "UPDATE mae_pc SET mae_pc.pc = '" & NulosC(ComputerName()) & "' WHERE (((mae_pc.id)=" & NulosN(xRst("id")) & "));"
    
        'SE ENCONTRO EL DISCO LLAMAR AL FORMULARIO DE INGRESO DE USUARIOS
        MDIPrincipal.Show
    End If
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : ChequeEmpresa()
'* Tipo           : FUNCION
'* Descripcion    : FUNCION QUE VERIFICA QUE SE HAYA SELECCIONADO UNA EMPRESA PARA PODER INGRESAR AL
'*                  SISTEMA.
'* Paranetros     : NULL
'* Retorna        : BOOLEAN
'*****************************************************************************************************
Function ChequeEmpresa() As Boolean
    If xIdEmpresa = 0 Then
        MsgBox "No ha especificado ninguna empresa a procesar, seleccione " & Chr(13) _
            & "una empresa y vuelva a intentar esta opcion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        ChequeEmpresa = False
        Exit Function
    End If
    ChequeEmpresa = True
End Function

'*****************************************************************************************************
'* Nombre Modulo  : CargaDatosEmpresa()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : PROCEDIMIENTO QUE CARGA LOS DATOS DE LA EMPRESA Y LOS ALMACENA EN VARIABLES
'*                  PREVIAMENTE DEFINIDAD
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub CargaDatosEmpresa()
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT * FROM  mae_empresa", xCon
    xNumRuc = NulosC(Rst("numruc"))
    xNomEmp = NulosC(Rst("nomemp"))
    CONTABILIZAR = NulosN(Rst("procon"))
    AnoTra = NulosC(Rst("anotra"))
    
    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : AbrirDataEnlace()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : PROCEDIMIENTO QUE ABRE LA BASE DE DATOS  DE ACCESO, PARA SELECCIONAR LA EMPRESA DE
'*                  TRABAJO ACCTUAL
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub AbrirDataEnlace()
    SeEjecutoEmp = False
    
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCad As String
    Dim xRutaData As String
    
    xCon.Close                                          ' CERRAMOS LA CONECCION
    Set xCon = Nothing                                  ' DESTRUIMOS LA CONECCION
    
    xRutaData = Trim(AP_RUTABD) + "data.mdb"
    xFun.F_BASEDATOS = xRutaData                        ' PASAMOS LA RUTA DE LA BASE DE DATOS
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"       ' PASAMOS LA RUTA DEL ARCHIVO DE TRABAJO DE LA BASE DE DATOS
    xFun.F_PASSWORD = Eps_Pass                          ' PASOMOS EL PASSWORD DE LA BASE DE DATOS
    xFun.F_USUARIO = Eps_User                           ' PASAMOS EL USUARIO DE LA BASE DE DATOS
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"        ' PASAMOS EL NOMBRE DEL PROVEEDOR DE DATOS PARA ADO 2.5
    
    Set xCon = xFun.AbrirConeccion                      ' ABRIMOS LA CONECCION A LA BASE DE DATOS
    Set xFun = Nothing
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : ActivarOpcionesMenu10()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : PROCEDIMIENTO PARA ACTIVAR O DESACTIVAR ALGUNAS OPCIONES DEL MENU_10
'* Paranetros     : NULL
'* Retorna        : NULL
'* Observaciones  : ESTA OPCION DEBERIA DE OPTIMISARSE JUNTO CON EL PROCESO DE ACTIVACION DE MENUS DEL
'*                  SISTEMA
'*****************************************************************************************************
Sub ActivarOpcionesMenu10()
    'MDIPrincipal.menu10_1.Enabled = Not MDIPrincipal.menu10_1.Enabled
    'MDIPrincipal.menu10_13.Enabled = Not MDIPrincipal.menu10_13.Enabled
    
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    ' CARGANOS LAS OPCIONES DEL MENU PRINCIPAL GUARDADAS EN LA BASE DE DATOS
    RST_Busq Rst, "SELECT * FROM mae_menu WHERE oculto = -1", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        ' RECORREMOS LAS OPCIONES CARGADAS
        For A = 1 To Rst.RecordCount
            ' MOSTRAMOS LAS OPCIONES QUE ESTEN REGISTRADAS EN LA BASE DE DATOS
            MDIPrincipal(Rst("nomcon")).Visible = False
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : ActivarMenus()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : PROCEDIMIENTO PARA ACTIVAR O DESACTIVAR EL MENU PRINCIPAL DEL SISTEMA
'* Paranetros     : NULL
'* Retorna        : NULL
'* Observaciones  : ESTA OPCION DEBERIA DE OPTIMISARSE JUNTO CON EL PROCESO DE ACTIVACION DE MENUS DEL
'*                  SISTEMA
'*****************************************************************************************************
Sub ActivarMenus()
    MDIPrincipal.almacen.Enabled = Not MDIPrincipal.almacen.Enabled
    MDIPrincipal.Compras.Enabled = Not MDIPrincipal.Compras.Enabled
    MDIPrincipal.ventas.Enabled = Not MDIPrincipal.ventas.Enabled
    MDIPrincipal.contabilidad.Enabled = Not MDIPrincipal.contabilidad.Enabled
    MDIPrincipal.tesoreria.Enabled = Not MDIPrincipal.tesoreria.Enabled
    MDIPrincipal.produccion.Enabled = Not MDIPrincipal.produccion.Enabled
    MDIPrincipal.planillas.Enabled = Not MDIPrincipal.planillas.Enabled
    MDIPrincipal.gestion.Enabled = Not MDIPrincipal.gestion.Enabled
'''    MDIPrincipal.maestros.Enabled = Not MDIPrincipal.maestros.Enabled
    'MDIPrincipal.mantenimiento.Enabled = Not MDIPrincipal.mantenimiento.Enabled
    MDIPrincipal.setup.Enabled = Not MDIPrincipal.setup.Enabled
    
End Sub




'*****************************************************************************************************
'* Nombre Modulo  : SetearMenus()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : CARGA LAS OPCIONES DEL MENU DISPONIBLES PARA EL USUARIO ACTUAL
'* Paranetros     : NOMBRE    |  TIPO     |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Idusuario | INTEGER   | CODIGO UNICO DEL USUARIO ACTUAL
'* Retorna        : NULL
'*****************************************************************************************************
Sub SetearMenus(Idusuario As Integer)
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    On Error Resume Next
    
    ' CARGAMOS LAS OPCIONES DEL MENU DISPONIBLES PARA EL USUARIO
    RST_Busq Rst, "SELECT mae_menu.id, mae_menu.tipo, mae_menu.descripcion, mae_menu.nomcon,  mae_menu.oculto, mae_menuusuario.opcion1, " _
        & " mae_menuusuario.opcion2, mae_menuusuario.opcion3, mae_menuusuario.acceso FROM mae_menu LEFT JOIN mae_menuusuario ON " _
        & " mae_menu.id = mae_menuusuario.idmenu WHERE (((mae_menuusuario.idusuario)= " & Idusuario & "))", xCon

    ' PREGUNTAMOS SI HAY OPCIONES DISPONIBLES
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        'RECORREMOS LAS OPCIONES DISPONIBLES PARA SU RESPECTIVA ACTIVACION
        For A = 1 To Rst.RecordCount
            MDIPrincipal(Rst("nomcon")).Enabled = NulosN(Rst("acceso"))
            '--verificar si esta oculto el menu
            If NulosN(Rst("oculto")) = -1 Then
                MDIPrincipal(Rst("nomcon")).Visible = False
            Else
                MDIPrincipal(Rst("nomcon")).Visible = True
            End If
            
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
        Next A
    End If
    Err.Clear
End Sub

Function ComputerName() As String
'20/02/2012 Johan Castro
'--Obtiene el nombre del equipo
'--Fuente: http://www.forosdelweb.com/f69/como-obtener-nombre-del-equipo-vb6-850692/
  '-- Funcion auxiliar que devuelve el nombre del equipo llamando al API
  ComputerName = Space$(260)
  GetComputerName ComputerName, Len(ComputerName)
  ComputerName = Left$(ComputerName, InStr(ComputerName, vbNullChar) - 1)
End Function

