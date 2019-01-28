VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'* Nombre Archivo   : REPORTES.CLS
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO QUE ALMACENA LAS PROPIEDADES Y EVENTOS PUBLICOS DE LA CLASE
'*                    SGI2_CONTABILIDAD2.REPORTES
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 26/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

'*****************************************************************************************************
'* Nombre Metodo  : RepCaja()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsLibroDiarioEfectivo
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                   SE USARA
'* Retorna        :
'*****************************************************************************************************
Sub RepCaja(Con As ADODB.Connection)
    Set xCon = Con
    FrmConsLibroDiarioEfectivo.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : RepBanco()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsLibDiarioCtaCte
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                   SE USARA
'* Retorna        :
'*****************************************************************************************************
Sub RepBanco(Con As ADODB.Connection)
    Set xCon = Con
    FrmConsLibDiarioCtaCte.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : RepDAOT()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO CargaDatosEmpresa
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                   SE USARA
'* Retorna        :
'*****************************************************************************************************
Sub RepDAOT(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmDaot1.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : CentroCostos()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO CargaDatosEmpresa
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                   SE USARA
'* Retorna        :
'*****************************************************************************************************
Sub CentroCostos(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmCentroCostos1.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : CentroCostosUnificado()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmCentroCostosUni
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                   SE USARA
'* Retorna        :
'*****************************************************************************************************
'Sub CentroCostosUnificado(Con As ADODB.Connection)
'    Set xCon = Con
'    CargaDatosEmpresa
'    FrmCentroCostosUni.Show
'End Sub

'*****************************************************************************************************
'* Nombre Metodo  : KardexUnificado()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmKarResUnificado
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION A LA BASE DE DATOS QUE
'*                                                   SE USARA
'* Retorna        :
'*****************************************************************************************************
'Sub KardexUnificado(Con As ADODB.Connection)
'    Set xCon = Con
'    CargaDatosEmpresa
'    FrmKarResUnificado.Show
'End Sub
