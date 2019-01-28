VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mantenimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'* Nombre Archivo   : MANTENIMIENTOS.CLS
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO QUE ALMACENA LAS PROPIEDADES Y EVENTOS PUBLICOS DE LA CLASE
'*                    SGI2_CONTABILIDAD2.MANTENIMIENTO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 26/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL


'*****************************************************************************************************
'* Nombre Metodo  : ManBalance()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManBalance
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                   SE USARA
'* Retorna        :
'*****************************************************************************************************
'Sub ManBalance(Con As ADODB.Connection)
'    Set xCon = Con
'    FrmManBalance.Show
'End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManEstados()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManEstados
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                   SE USARA
'* Retorna        :
'*****************************************************************************************************
'Sub ManEstados(Con As ADODB.Connection)
'    Set xCon = Con
'    FrmManEstados.Show
'End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManDestinos()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmMaeDestinos
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                   SE USARA
'* Retorna        :
'*****************************************************************************************************
Sub ManDestinos(Tipo As Integer, Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmMaeDestinos.TipoMovimmiento = Tipo
    FrmMaeDestinos.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManOrigenes()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmMaeDestinos
'* Paranetros     : NOMBRE    |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Tipo      |  Integer           |  ESPECIFICA EL TIPO DE MOVIMIENTO
'*                  Con       |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                    SE USARA
'* Retorna        :
'*****************************************************************************************************
Sub ManOrigenes(Tipo As Integer, Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmMaeOrigen.TipoMovimmiento = Tipo
    FrmMaeOrigen.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManTC()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManTC
'* Paranetros     : NOMBRE    |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                    SE USARA
'* Retorna        :
'*****************************************************************************************************
Sub ManTC(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa
    FrmManTC.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManCentroCostoArea()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmCenCosArea
'* Paranetros     : NOMBRE    |  TIPO              |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                    SE USARA
'* Retorna        :
'*****************************************************************************************************
Sub ManCentroCostoArea(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmCenCosArea.Show
End Sub

