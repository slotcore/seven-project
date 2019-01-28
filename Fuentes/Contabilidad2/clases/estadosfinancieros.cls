VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "estadosfinancieros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'* Nombre Archivo   : ESTADOSFINANCIEROS.CLS
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO QUE ALMACENA LAS PROPIEDADES Y EVENTOS PUBLICOS DE LA CLASE
'*                    SGI2_CONTABILIDAD2.ESTADOSFINANCIEROS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 26/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

'*****************************************************************************************************
'* Nombre Metodo  : BalanceGeneral()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmBalanceGeneral
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |
'* Retorna        :
'*****************************************************************************************************
'Sub BalanceGeneral(Con As ADODB.Connection)
'    Set xCon = Con
'    CargaDatosEmpresa
'    FrmBalanceGeneral.Show
'End Sub

'*****************************************************************************************************
'* Nombre Metodo  : AnalisisCuenta()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmAnalisisCuenta
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |
'* Retorna        :
'*****************************************************************************************************
Sub AnalisisCuenta(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmAnalisisCuenta1.Show
End Sub
