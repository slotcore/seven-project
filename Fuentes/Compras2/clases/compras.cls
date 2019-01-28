VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public IdUsuario As Integer         'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL
Public IdMenu As Integer            'PROPIEDAD DE CLASE : RECIBE EL CODIGO UNICO DEL USUARIO ACTUAL


Sub ManOrdenrequerimiento(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmRequerimientos.DeDonde = 1
    FrmRequerimientos.Show
End Sub

Sub ManOrdenCotizacion(Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = IdUsuario
    FrmManOrdCotiza.Show
End Sub

Sub ManOrdenCompra(Usuario As Integer, Con As ADODB.Connection)
    Set xCon = Con
    CargaDatosEmpresa
    xIdMenu = IdMenu
    xIdUsuario = Usuario
    FrmManOrdenCompra.Show
End Sub

Sub AprobarRequerimiento(Usuario As Integer, Con As ADODB.Connection)
    Set xCon = Con
    xIdUsuario = Usuario
    CargaDatosEmpresa
    FrmApruebaRequerimiento.Show
End Sub

Sub AprobarCotizacion(Usuario As Integer, Con As ADODB.Connection)
    Set xCon = Con
    xIdUsuario = Usuario
    CargaDatosEmpresa
    FrmApruebaCotiza.Show
End Sub

Sub AprobarOrdenCompra(Usuario As Integer, Con As ADODB.Connection)
    Set xCon = Con
    xIdUsuario = Usuario
    CargaDatosEmpresa
    FrmApruebaOrdenCompra.Show
End Sub
