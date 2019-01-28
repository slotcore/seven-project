VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cajabancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'* Nombre Archivo   : CAJABANCOS.CLS
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO QUE ALMACENA LAS PROPIEDADES Y EVENTOS PUBLICOS DE LA CLASE sgi2_cajabancos
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 12/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL

'*****************************************************************************************************
'* Nombre Archivo   : EgresoCajaBanco2
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO FrmEgresoCajaBanco
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO              |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'*                    MesTrabajo   |  Integer           |  ALMACENA EL MES DE TRABAJO
'* DEVUELVE         :
'*****************************************************************************************************
Sub EgresoCajaBanco2(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatos
    FrmEgresoCajaBanco.Show
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : IngresoCajaBanco2
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO FrmIngresoCajaBanco
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'*                    MesTrabajo   |  Integer           |  ALMACENA EL MES DE TRABAJO
'* DEVUELVE         :
'*****************************************************************************************************
Sub IngresoCajaBanco2(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatos
    FrmIngresoCajaBanco.Show
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : ConsultaCtaCte
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO ConsultaCtaCte
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* DEVUELVE         :
'*****************************************************************************************************
Sub ConsultaCtaCte(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatos
    FrmCtaCte4.Show
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : FrmCanjesFacturas
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO FrmCanjesFacturas
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* DEVUELVE         :
'*****************************************************************************************************
Sub CanjeDocumentos(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatos
    FrmCanjesFacturas.Show
End Sub


'*****************************************************************************************************
'* Nombre Archivo   : CanjeLetras
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO FrmCanjeLetras
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO              |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* DEVUELVE         :
'*****************************************************************************************************
'Sub CanjeLetras(Con As ADODB.Connection)
'    Set xCon = Con
'    CargaDatos
'    FrmCanjeLetras.Show
'End Sub

'*****************************************************************************************************
'* Nombre Archivo   : ProgramarPagos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO FrmProgramarPagos
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'*                    IdUsuario    |  Integer           |  ESPECIFICA EL ID DEL USUARIO
'*                    MesTrabajo   |  Integer           |  ESPECIFICA EL MES DE TRABAJO ACTUAL
'* DEVUELVE         :
'*****************************************************************************************************
Sub ProgramarPagos(Con As ADODB.Connection, IdUsuario As Integer, MesTrabajo As Integer)
    Set xCon = Con
    xIdUsuario = IdUsuario
    xMes = MesTrabajo
    CargaDatos
    FrmProgramarPagos.Show
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : EmiRendir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO FrmCuentasRendir
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'*                    IdUsuario    |  Integer           |  ESPECIFICA EL ID DEL USUARIO
'*                    MesTrabajo   |  Integer           |  ESPECIFICA EL MES DE TRABAJO ACTUAL
'* DEVUELVE         :
'*****************************************************************************************************
Sub EmiRendir(Con As ADODB.Connection, IdUsuario As Integer, MesTrabajo As Integer)
    Set xCon = Con
    xIdUsuario = IdUsuario
    xMes = MesTrabajo
    CargaDatos
    FrmCuentasRendir.Show
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : DevRendir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA FrmCuentasRendir_Dev
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'*                    IdUsuario    |  Integer           |  ESPECIFICA EL ID DEL USUARIO
'*                    MesTrabajo   |  Integer           |  ESPECIFICA EL MES DE TRABAJO ACTUAL
'* DEVUELVE         :
'*****************************************************************************************************
Sub DevRendir(Con As ADODB.Connection, IdUsuario As Integer, MesTrabajo As Integer)
    Set xCon = Con
    xIdUsuario = IdUsuario
    xMes = MesTrabajo
    CargaDatos
    FrmCuentasRendir_Dev.Show
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : ConsultaProgramacionPagos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO FrmConsProgramarPagos
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'*                    MesTrabajo   |  Integer           |  ESPECIFICA EL MES DE TRABAJO ACTUAL
'* DEVUELVE         :
'*****************************************************************************************************
Sub ConsultaProgramacionPagos(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    CargaDatos
    FrmConsProgramarPagos.Show
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : DiferenciaCambio
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO FrmTipoCambio
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'*                    MesTrabajo   |  Integer           |  ESPECIFICA EL MES DE TRABAJO ACTUAL
'* DEVUELVE         :
'*****************************************************************************************************
Sub DiferenciaCambio(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    CargaDatos
    FrmTipoCambio.Show
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Librobancos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO FrmLibroBancos3
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO              |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* DEVUELVE         :
'*****************************************************************************************************
Sub Librobancos(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatos
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmLibroBancos3.Show
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Anticuamiento
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL FORMULARIO FrmAnticuamiento
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO              |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con          |  ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* DEVUELVE         :
'*****************************************************************************************************
Sub Anticuamiento(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatos
    FrmAnticuamiento.Show
End Sub


'Sub ConsultaCtaCteOtros(Con As ADODB.Connection)
'    '--10/02/10
'    Set xCon = Con
'    CargaDatos
'    FrmCtaCteOtros.Show
'End Sub


