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
'* Nombre Archivo   : COMPRAS
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO DE CLASE DONDE SE DEFINEN LOS METODOS PARA LA GESTION DE LA INFORMACION
'*                    DEL MODULO DE COMPRAS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

'*****************************************************************************************************
'* Nombre Metodo  : AnalizisCompras
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO AnalizisCompras
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection |  ESPECIFICA LA CONECCION ACTUAL
'* Retorna        :
'*****************************************************************************************************
Sub AnalizisCompras(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    'FrmConsComp_Gerencial.Show
    FrmGestionCompras.Show
End Sub
