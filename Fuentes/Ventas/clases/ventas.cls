VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 '*****************************************************************************************************
'* Nombre Archivo   : VENTAS.CLS
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO DONDE SE DECLARAN LOS METODOS Y LAS PROPIEDADES DE LA CLASE Sgi2_ventas
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 21/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL


'*****************************************************************************************************
'* Nombre Metodo  : Ventas()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmVentas2
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'*                  MesTrabajo| Integer           |  INDICA EL MES DE TRABAJO ACTUAL
'* Retorna        :
'*****************************************************************************************************
Sub Ventas(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa
    FrmVentas2.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : Clientes()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManClientes
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'*                  DeDonde   | Integer           |  INDICA DESDE DONDE ES INVOCADA LA CLASE
'* Retorna        :
'*****************************************************************************************************
Sub Clientes(Con As ADODB.Connection, DeDonde As Integer)
    xDeDonde = DeDonde
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    Set xCon = Con
    FrmManClientes.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : PuntosVenta()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManPuntosVenta
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub PuntosVenta(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManPuntosVenta.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : GuiasRemision()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManGuias
'* Paranetros     : NOMBRE     |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'*                  MesTrabajo | MesTrabajo        |  ESPECIFICA EL MES DE TRABAJO
'* Retorna        :
'*****************************************************************************************************
Sub GuiasRemision(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManGuias.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : Cotizaciones()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmCotizaciones
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub Cotizaciones(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmCotizaciones.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : MantProdCen()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManProdCEN
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub MantProdCen(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManProdCEN.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : LevantarPedidos()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmLevantar
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub LevantarPedidos(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmLevantar.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : MuestraPedidos()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmMuestraPedidos
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub MuestraPedidos(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmMuestraPedidos.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ReporteVentas()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsVenta_Administ
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub ReporteVentas(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmConsVenta1.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : Empaques()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmEmpaques
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
''*****************************************************************************************************
'Sub Empaques(Con As ADODB.Connection)
'    Set xCon = Con
'    FrmEmpaques.Show
'End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ActualizarDatosVenta()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmActualizaDatosFactura
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
'Sub ActualizarDatosVenta(Con As ADODB.Connection)
'    Set xCon = Con
'    FrmActualizaDatosFactura.Show
'End Sub

'*****************************************************************************************************
'* Nombre Metodo  : LiqGasDebito()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmGastosDebito1
'* Paranetros     : NOMBRE     |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'*                  MesTrabajo | Integer           |  ESPECIFICA EL MES DE TRABAJO ACTUAL
'* Retorna        :
'*****************************************************************************************************
Sub LiqGasDebito(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa
    FrmGastosDebito1.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManConceptoNC_ND()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManConceptoNC_ND
'* Paranetros     : NOMBRE     |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con        | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub ManConceptoNC_ND(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa
    FrmManConceptoNC_ND.Show
End Sub


'*****************************************
' Creado: 30/03/2012 - Jose Chaco - Reporte de Devoluciones
'*****************************************
Sub ConDevoluciones(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmRepDevoluciones.Show
End Sub
