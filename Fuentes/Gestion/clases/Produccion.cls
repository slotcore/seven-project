VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Produccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'* Nombre Archivo   : PRODUCCION
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO DE CLASE DONDE SE DEFINEN LOS METODOS PARA EFECTUAR EL ANALISIS DE LA
'*                    PRODUCCION
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public Enum e_PROGRAMA
    E_INSUMO = 0
    e_TAREA = 1
    e_EQUIPO = 2
End Enum

'*****************************************************************************************************
'* Nombre Metodo  : AnalizisProduccion
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsProd_Gerencial
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub AnalizisProduccion(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmConsProd_Gerencial.Show
    FrmConsProd_Gerencial.SetFocus
End Sub



