VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL

Public Enum TIPO_MOVIMIENTO
'--SE USA ESTOS VALORES CON LA FINALIDAD DE SER UNICOS
    DET_Compra = 1
    DET_Venta = 2
End Enum

Sub ManPlanCuentas(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManPlanCtas.Show
End Sub

Sub ManCentroCostos(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManCentroCostos.Show
End Sub

Sub ManDetraccion(Con As ADODB.Connection, MesTrabajo As Integer, Tipo As TIPO_MOVIMIENTO)
    Set xCon = Con
    CargaDatosEmpresa
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManDetraccion.xTIPO_MOVIMIETO = Tipo
    FrmManDetraccion.Show
End Sub

Sub ManRetencion(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    CargaDatosEmpresa
    FrmRetencion.Show
End Sub

Sub ManPercepcion(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    xMes = MesTrabajo
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmRegPercepciones.Show
End Sub

Sub ManCtaDocumento(Con As ADODB.Connection)
    Set xCon = Con
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmMantDocumCta.Show
End Sub

Sub ManProviciones(Con As ADODB.Connection, MesTrabajo As Integer)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    xMes = MesTrabajo
    FrmProvisiones2.Show
End Sub

'Sub ManProviciones2(Con As ADODB.Connection, MesTrabajo As Integer)
'    Set xCon = Con
'    CargaDatosEmpresa
'    xIdUsuario = IdUsuario
'    xMes = MesTrabajo
'    FrmProvisiones2.Show
'End Sub

'--------------------------------------
' MODIFICADO 13/07/2012 - JOSE CHACON
'--------------------------------------
Sub verLibroCosto(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManLibroCosto3.Show
End Sub

