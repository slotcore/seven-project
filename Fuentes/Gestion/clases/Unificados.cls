VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Unificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'* Nombre Archivo   : UNIFICADOS
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO DE CLASE DONDE SE DEFINEN LOS METODOS UNIFICAR LA INFORMACION DE TODAS LAS
'*                    BASES DE DATOS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

'*****************************************************************************************************
'* Nombre Metodo  : UnificadoProduccion
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmProduccion
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub UnificadoProduccion(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmProduccion2.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : UnificadoAbastecimiento
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmAbastecimiento
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub UnificadoAbastecimiento(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmAbastecimiento.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : UnificadoProducido
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsultaProduccion
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub UnificadoProducido(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmConsultaProduccion2.Show
End Sub
