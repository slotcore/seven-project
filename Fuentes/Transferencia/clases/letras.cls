VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "letras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL

Sub ManLetras(MesTrabajo As Integer, Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    xMes = MesTrabajo
    FrmLetras1.Show
End Sub

Sub ManPlanilla(MesTrabajo As Integer, Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    xMes = MesTrabajo
    FrmPlanilla1.Show
End Sub

