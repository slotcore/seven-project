VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'* Nombre Archivo   : COMPRAS.CLS
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO DONDE SE DECLARAN LOS METODOS Y LAS PROPIEDADES DE LA CLASE Sgi2Compras
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 09/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL

'*****************************************************************************************************
'* Nombre Metodo  : RegCompras2()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmCompras2, DEVUELVE EL ID DEL REGISTRO LUEGO QUE SE TERMINA
'*                  LA OPERACION, ESTE VALOR SERA UTILIZADO CUANDO Origen = 1
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'*                  MesTrabajo| Integer           |  INDICA EL MES DE TRABAJO ACTUAL
'*                  Origem    | Integer           |  ESPECIFICA DE DONDE ES INVOCADO EL METODO
'*                                                   0 = MENU PRINCIPAL; 1 = OTRA LIBRERIA O PROGRAMA
'* Retorna        : INTEGER
'*****************************************************************************************************
Function RegCompras2(Con As ADODB.Connection, MesTrabajo As Integer, Origen As Integer) As Integer
    Set xCon = Con
    xMes = MesTrabajo
    xOrigen = Origen
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmCompras2.Show
'    RegCompras2. = IdCompraReg
End Function

'*****************************************************************************************************
'* Nombre Metodo  : RegHonorarios()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmHonorarios
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'*                  MesTrabajo| Integer           |  INDICA EL MES DE TRABAJO ACTUAL
'*                  Origem    | Integer           |  ESPECIFICA DE DONDE ES INVOCADO EL METODO
'*                            |                   |  0 = MENU PRINCIPAL; 1 = OTRA LIBRERIA O PROGRAMA
'* Retorna        :
'*****************************************************************************************************
Sub RegHonorarios(Con As ADODB.Connection, MesTrabajo As Integer, Origen As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    xOrigen = Origen
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmHonorarios.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : AsignaPrecioItem()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmMantComPrecios
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub AsignaPrecioItem(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmMantComPrecios.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ManProveedor()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmManProveedores
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'*                  DeDonde   | Integer           |  ESPECIFICA SI LA TABLA SE SINCRONIZARA CON OTRAS
'*                            |                   |  BASE DE DATOS 1 = NO SINCRONIZADO, 2 = SINCRONIZA
'* Retorna        :
'*****************************************************************************************************
Sub ManProveedor(Con As ADODB.Connection, DeDonde As Integer)
    xDeDonde = DeDonde
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    If xDeDonde = 1 Then CargaDatosEmpresa
    FrmManProveedores.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : OrdenCompra()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmOrdenCompra
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub OrdenCompra(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmOrdenCompra.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : CentroCostos()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmGastosCentroCosto
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub CentroCostos(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmGastosCentroCosto.Show
End Sub

''*****************************************************************************************************
''* Nombre Metodo  : IngresoAlmacen()
''* Tipo           : METODO DE CLASE
''* Descripcion    : LLAMA AL FORMULARIO FrmIngresoAlmacen
''* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
''*                  ----------------------------------------------------------------------------------
''*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
''* Retorna        :
''*****************************************************************************************************
'Sub IngresoAlmacen(Con As ADODB.Connection, MesTrabajo As Integer)
'
'    Set xCon = Con
'    CargaDatosEmpresa
'    xMes = MesTrabajo
'    xIdMenu = IdMenu
'    xIdUsuario = IdUsuario
'    FrmIngresoAlmacen2.Show
'End Sub

'*****************************************************************************************************
'* Nombre Metodo  : RepCompras()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsCompra_Administ
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub RepCompras(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmConsCompra.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ConsultaIngSalAlm()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsIngAlmacen
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub ConsultaIngSalAlm(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmConsIngAlmacen.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : RepHonorario()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsHonorario
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub RepHonorario(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmConsHonorario.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : Reembolsables()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmReembolsables
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub Reembolsables(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    CargaDatosEmpresa
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmReembolsables.Show
End Sub

Sub ManCotizacionCompra(Con As ADODB.Connection, MesTrabajo As Integer)
'    Set xCon = Con
'    CargaDatosEmpresa
'    xMes = MesTrabajo
'    xIdMenu = IdMenu
'    xIdUsuario = IdUsuario
'    FrmManCotizacion.Show
End Sub
