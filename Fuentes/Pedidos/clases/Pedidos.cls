VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'* Nombre Archivo   : PEDIDOS.CLS
'* Tipo             : MODULO DE CLASE
'* Descripcion      : MODULO QUE ALMACENA LAS PROPIEDADES Y EVENTOS PUBLICOS DE LA CLASE SGI2_PEDIDOS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 28/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL


'*****************************************************************************************************
'* Nombre Metodo  : Pedidos()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmPedido
'* Paranetros     : NOMBRE   |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con      | ADODB.Connection  |  CONECCION ACTIVA QUE SE USARA PARA ACCEDER A LA BD
'*                  MesTra   | Integer           |  ESPECIFICA EL MES DE TRABAJO ACTUAL
'* Retorna        :
'*****************************************************************************************************
Sub Pedidos(Con As ADODB.Connection, MesTra As Integer)
    Set xCon = Con
    xMes = MesTra
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa
    FrmPedido2.Show
End Sub

'*****************************************************************************************************
'* Nombre           : MostrarCronogramaEntregas
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLAMA AL METODO CargarCalendario DE LA CLASE Eps_MantTablas.formularios, ESTE
'*                    METODO MUESTRA EN UN CALENDARIO EL CRONOGRAMA DE ENTREGA DE PEDIDOS, ESTE
'*                    FORMULARIO ES DE SOLO CONSULTA
'* Paranetros       : NOMBRE |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Con    |  ADODB.Connection |  CONECCION ACTIVA QUE SE USARA PARA ACCEDER A LA BD
'* Devuelve         :
'*****************************************************************************************************
Sub MostrarCronogramaEntregas(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
'    LlenarDatos
'    Dim xFrm As New Eps_MantTablas.Formularios
'    xFrm.XTAMAÑO_TOOL = I24x24
'    xFrm.Titulo = "VENTAS - Cronograma de Entregas"
'    xFrm.P_RutaData = AP_RUTASY & "calendario.mdb"
'    xFrm.P_RutaFileTrabajo = AP_RUTASY & "seven.mdw"
'    xFrm.CargarCalendario Trim(App.Path) + "\calendario.mdb", Con
    FrmCronoPedidos2.Show
'    Set xFrm = Nothing
End Sub

'*****************************************************************************************************
'* Nombre Metodo  : ReportePedidos()
'* Tipo           : METODO DE CLASE
'* Descripcion    : LLAMA AL FORMULARIO FrmConsPedido
'* Paranetros     : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                  Con       | ADODB.Connection  |  ESPECIFICA LA CONECCION ACTIVA A LA BASE DE DATOS
'* Retorna        :
'*****************************************************************************************************
Sub ReportePedidos(Con As ADODB.Connection)
    ' Modificado: 23/03/12 Jose Chacon -Agregar referencia FrmRepPedido2
    Set xCon = Con
    CargaDatosEmpresa
    FrmRepPedido2.Show
End Sub
