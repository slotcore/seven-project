VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public IdUsuario As Integer      'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL


Sub Mantenimiento(Con As ADODB.Connection, Usuario As Integer)
    Set xCon = Con
    CargaDatosEmpresa xCon
    FrmOrdenTrabajo.Show
End Sub

Sub MantEquipos(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa xCon
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManEquipos.Show
End Sub

Private Sub Class_Initialize()
    Nomsis = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub

Sub ManPreventivo(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa xCon
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManPreven.Show
End Sub
