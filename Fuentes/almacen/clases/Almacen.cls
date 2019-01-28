VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'* Nombre Archivo   : ALMACEN.CLS
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO QUE ALMACENA LAS PROPIEDADES Y EVENTOS PUBLICOS DE LA CLASE SGI2_ALMACEN
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 04/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit
Public IdUsuario As Integer         'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL


'*****************************************************************************************************
'* Nombre Metodo  : MantItem()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManAlamacen
'* Parametros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |
'*                  DeDonde   | Integer           |
'* Retorna        :
'*****************************************************************************************************
Sub MantItem(Con As ADODB.Connection, DeDonde As Integer)
    'DeDonde   1 = menu compras
    'DeDonde   2 = menu de opciones
    xDeDonde = DeDonde
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    xIdMenu = IdMenu
    Set xCon = Con
    If xDeDonde = 1 Then CargaDatosEmpresa
    FrmManItems.Show
End Sub

Sub AlmacenajeAuto(Con As ADODB.Connection)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    xIdMenu = IdMenu
    Set xCon = Con
    CargaDatosEmpresa
    FrmManAlmAuto.Show
End Sub

Sub TomaInventario(Con As ADODB.Connection)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    xIdMenu = IdMenu
    Set xCon = Con
    CargaDatosEmpresa
    FrmManInventario.Show
End Sub

Sub TransferenciaAlmacen(Con As ADODB.Connection, MesTrabajo As Integer)
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    xIdMenu = IdMenu
    Set xCon = Con
    CargaDatosEmpresa
    FrmTransfAlmacen.Show
End Sub

Sub IngresoAlmacen(Con As ADODB.Connection, MesTrabajo As Integer)
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    xIdMenu = IdMenu
    Set xCon = Con
    CargaDatosEmpresa
    FrmIngresoAlmacen3.Show
End Sub

Sub IngresoAlmacen_(DataSource As String, MesTrabajo As Integer)
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    xIdMenu = IdMenu
    xDataSource = DataSource
    CargaDatosEmpresa_
    FrmIngresoAlmacen3.Show
End Sub

Sub RecepcionAlmacen(Con As ADODB.Connection, MesTrabajo As Integer)
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    Set xCon = Con
    CargaDatosEmpresa
    FrmManRecepcion.Show
End Sub

Sub ManAlmacen(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa
    FrmManAlmacen.Show
End Sub

Sub ConsultaStock(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa
    FrmRepStock.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ActualizaPrecio()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmActualizaDatos
'* Parametros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |
'* Retorna        :
'*****************************************************************************************************
Sub ActualizaPrecio(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmActualizaDatos.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ActualizaDatosContables()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmActDatosContables
'* Parametros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |
'* Retorna        :
'*****************************************************************************************************
Sub ActualizaDatosContables(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmActDatosContables.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : CargarAlmacenes()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmItemsAlmacen
'* Parametros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |
'* Retorna        :
'*****************************************************************************************************
Sub CargarAlmacenes(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmItemsAlmacen.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManClase()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManClase
'* Parametros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |
'*                  DeDonde   | Integer           |
'* Retorna        :
'*****************************************************************************************************
Sub ManClase(Con As ADODB.Connection, DeDonde As Integer)
    'DeDonde   1 = menu compras
    'DeDonde   2 = menu de opciones
    xDeDonde = DeDonde
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    Set xCon = Con
    FrmManClase.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManSubClase()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManSubClase
'* Parametros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |
'*                  DeDonde   | Integer           |
'* Retorna        :
'*****************************************************************************************************
Sub ManSubClase(Con As ADODB.Connection, DeDonde As Integer)
    'DeDonde   1 = menu compras
    'DeDonde   2 = menu de opciones
    xDeDonde = DeDonde
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    Set xCon = Con
    FrmManSubClase.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManFamilia()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManFamilia
'* Parametros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |
'*                  DeDonde   | Integer           |
'* Retorna        :
'*****************************************************************************************************
Sub ManFamilia(Con As ADODB.Connection, DeDonde As Integer)
    'DeDonde   1 = menu compras
    'DeDonde   2 = menu de opciones
    xDeDonde = DeDonde
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    Set xCon = Con
    FrmManFamilia.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManTipoProducto()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManTipoProducto
'* Parametros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |
'*                  DeDonde   | Integer           |
'* Retorna        :
'*****************************************************************************************************
Sub ManTipoProducto(Con As ADODB.Connection, DeDonde As Integer)
    'DeDonde   1 = menu compras
    'DeDonde   2 = menu de opciones
    xDeDonde = DeDonde
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    Set xCon = Con
    FrmManTipoProducto.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManUnidades()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManUnidades
'* Parametros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |
'*                  DeDonde   | Integer           |
'* Retorna        :
'*****************************************************************************************************
Sub ManUnidades(Con As ADODB.Connection, DeDonde As Integer)
    'DeDonde   1 = menu compras
    'DeDonde   2 = menu de opciones
    xDeDonde = DeDonde
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    Set xCon = Con
    FrmManUnidades.Show
End Sub

Sub ConsultaItems(Con As ADODB.Connection)
    Set xCon = Con
    FrmConsultaItems.Show
End Sub


Sub DevolucionAlmacen(Con As ADODB.Connection, MesTrabajo As Integer)
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    Set xCon = Con
    CargaDatosEmpresa
    FrmManDevolucion.Show
End Sub

Sub ManAlmacenajeAuto(Con As ADODB.Connection)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    Set xCon = Con
    CargaDatosEmpresa
    FrmManAlmAuto.Show
End Sub

Sub ManDespachoAuto(Con As ADODB.Connection)
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    Set xCon = Con
    CargaDatosEmpresa
    FrmManDespAuto.Show
End Sub
