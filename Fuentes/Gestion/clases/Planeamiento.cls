VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Planeamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'* Nombre Archivo   : PLANEAMIENTO
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO DE CLASE DONDE SE DEFINEN LOS METODOS PARA REALIZAR EL PLANEAMIENTO DE LA
'*                    PRODUCCION
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer         'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL


'*****************************************************************************************************
'* Nombre Metodo  : PlanVentasEstimado
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmPVEstimado
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub PlanVentasEstimado(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmPVEstimado2.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : PlanVentas
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmPlanVentas
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub PlanVentas(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmPlanVentas2.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : PlanProduccion
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmPlanProduccion
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub PlanProduccion(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmPlanProduccion3.Show
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : PlanAbastecimiento
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO PlanAbastecimiento
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub PlanAbastecimiento(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmPlanAbastecimiento3.Show
End Sub



'*****************************************************************************************************
'* Nombre Metodo  : PlanProduccion3
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmPlanProduccion
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub PlanProduccion3(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmPlanProduccion3.Show
End Sub


'*****************************************************************************************************
'* Nombre Metodo  : PlanAbastecimiento3
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO PlanAbastecimiento
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       |  ADODB.Connection |  ESPECIFICA LA CONECCION A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub PlanAbastecimiento3(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmPlanAbastecimiento3.Show
End Sub
