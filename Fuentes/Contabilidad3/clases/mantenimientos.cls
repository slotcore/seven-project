VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mantenimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL


Sub ManConcepto(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmConcepto.Show
End Sub

Sub ManInforme(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmInforme.Show
End Sub

Sub VerInforme(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    FrmVerInforme.Show
End Sub

