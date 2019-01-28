VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "produccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'* Nombre Archivo   : PRODUCCION.CLS
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO QUE ALMACENA LAS PROPIEDADES Y EVENTOS PUBLICOS DE LA CLASE SGI2_PRODUCCION
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 28/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public Enum e_PROGRAMA
    E_INSUMO = 0
    e_TAREA = 1
    e_EQUIPO = 2
    E_COSTO = 3
End Enum

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL

'*****************************************************************************************************
'* Nombre Metodo  : MamRecetas()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManReceta
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub MamRecetas(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa Con
    FrmManReceta2.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : OrdenProduccion()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManProduccion
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'*                  MesTrabajo |  Integer           |  ESPECIFICA EL MES DE TRABAJO
'* Retorna        :
'*****************************************************************************************************
Sub OrdenProduccion(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa Con
'    FrmManProduccion2.Show
'    FrmManProduccion2.SetFocus
    
    FrmManRegProd.Show
    FrmManRegProd.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ProgramaProduccion()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManProduccionPrograma
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'*                  MesTrabajo |  Integer           |  ESPECIFICA EL MES DE TRABAJO
'* Retorna        :
'*****************************************************************************************************
Sub ProgramaProduccion(Con As ADODB.Connection, MesTrabajo As Integer)
'    Set xCon = Con
'    xMes = MesTrabajo
'    xIdUsuario = IdUsuario
'    CargaDatosEmpresa Con
'    FrmManProduccionPrograma.Show
'    FrmManProduccionPrograma.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : Estacionalidad()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManEstacionalidad
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub Estacionalidad(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManEstacionalidad.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : RepProduccion()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsProduccion
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub RepProduccion(Con As ADODB.Connection)
    Set xCon = Con
    FrmConsProduccion.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManTareas()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManTareas
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub ManTareas(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa Con
    FrmManTareas.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManGrupos()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManGrupos
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub ManGrupos(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa Con
    FrmManGrupos.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : PersonalProduccion()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManControlPers
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub PersonalProduccion(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa Con
    FrmManControlPers2.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : IngresoTareas()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO IngresoTareas
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'*                  MesTrabajo |  Integer           |  ESPECIFICA EL MES DE TRABAJO
'* Retorna        :
'*****************************************************************************************************
Sub IngresoTareas(Con As ADODB.Connection, MesTrabajo As Integer)
    '--04/08/2008
    Set xCon = Con
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa Con
    FrmControlTarea1.Show
    FrmControlTarea1.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ProgramarDiaTmp()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmProgramaDiaTmp
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'*                  MesTrabajo |  Integer           |  ESPECIFICA EL MES DE TRABAJO
'* Retorna        :
'*****************************************************************************************************
Sub ProgramarDiaTmp(Con As ADODB.Connection, MesTrabajo As Integer)
'    Set xCon = Con
'    xMes = MesTrabajo
'    CargaDatosEmpresa Con
'    FrmProgramaDiaTmp.Show
'    FrmProgramaDiaTmp.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : DistribucionTareas()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManControlArea
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub DistribucionTareas(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManControlArea.Show
    FrmManControlArea.SetFocus
End Sub


Sub ProgramarDia(Con As ADODB.Connection, MesTrabajo As Integer)
'--26/08/08
'    Set xCon = Con
'    xMes = MesTrabajo
'    CargaDatosEmpresa Con
'    FrmManProgramaDia.Show
'    FrmManProgramaDia.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : RepTarea()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsTarea
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub RepTarea(Con As ADODB.Connection)
    Set xCon = Con
    FrmConsTarea.Show
    FrmConsTarea.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ConfigurarCosto()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManCosto
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub ConfigurarCosto(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa Con
    FrmManCosto.Show
    FrmManCosto.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : Class_Initialize()
'* Tipo           : METODO DE CLASE
'* Descripcion    : CARGA DATOS DEL ARCHIVO DE CONFIGURACION SEVEN.INI
'* Paranetros     :
'* Retorna        :
'*****************************************************************************************************

Private Sub Class_Initialize()
    Nomsis = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : CostoProduccion()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmEvaluarCosto
'* Paranetros     : NOMBRE     |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        |  ADODB.Connection  |  ESPECIFICA LA CONECCION DE LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub CostoProduccion(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa Con
    FrmEvaluarCosto1.Show
    FrmEvaluarCosto1.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : RepCosto()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsCosto
'* Paranetros     :
'* Retorna        :
'*****************************************************************************************************
Sub RepCosto(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa Con
    FrmConsCosto.Show
    FrmConsCosto.SetFocus
End Sub

Sub CronogramaProduccion(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa Con
    'FrmCronoProduccion.Show
    'FrmCronoProduccion2_1.Show
    FrmCronoProduccion2_3_1.Show
End Sub

Sub Rendimiento(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa Con
    FrmProRendimiento.Show
End Sub

Sub CronogramaTareas(Con As ADODB.Connection)
'    Set xCon = Con
'    CargaDatosEmpresa Con
'    xIdMenu = IdMenu
'    xIdUsuario = IdUsuario
'    FrmCronoTarea2.Show
End Sub

Sub LineaDeTiempo(Con As ADODB.Connection)
'    Set xCon = Con
'    CargaDatosEmpresa Con
'    xIdMenu = IdMenu
'    xIdUsuario = IdUsuario
'    FrmCronoLineaTiempo.Show
End Sub

Sub HojaDeRuta(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa Con
    FrmGeneraHoja.Show
End Sub

'Sub GenOrdenProduccion(Con As ADODB.Connection, MesTrabajo As Integer)
'    Set xCon = Con
'    xMes = MesTrabajo
'    CargaDatosEmpresa Con
'    xIdMenu = IdMenu
'    xIdUsuario = IdUsuario
'    FrmGenerarOrdProd2_1.Show
'End Sub

Sub ConfigurarPersonalxTareas(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManPersTarea.Show
    
End Sub

Sub CronogramaMantLinea(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManLinea.Show
    
End Sub

Sub RepLinea(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmRepLinea.Show
End Sub

Sub TempusImportarMarcacion(Con As ADODB.Connection)
    '--12-02-04 Johan Castro
    '--Importar Marcacion de sistema Tempus
    ' 05/03/2012 - Jose Chacon - Nueva referencia frmCargarTemp2
    Set xCon = Con
    CargaDatosEmpresa Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    frmCargarTemp2.Show
End Sub

Sub RepCompProduccion(Con As ADODB.Connection)
    '--12-03-22 Jose Chacon
    '--Comparativo entre Planeacion y Produccion
    Set xCon = Con
    FrmRepProduccion.Show
End Sub


Sub RepPlaneacion(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmPlaneProduccion.Show
End Sub

'********************************
' Creado 30/05/2012 - Jose Chacon
'********************************
Sub GenSolicitudMat(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    CargaDatosEmpresa Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManSolMat.Show
End Sub

Sub GenOrdenProd(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    CargaDatosEmpresa Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManOrdProd.Show
End Sub
