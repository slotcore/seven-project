VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Punto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Sub PuntoVenta(Con As ADODB.Connection, xIdUsuario As Integer)
    Set xCon = Con
    mIdEmpleado = xIdUsuario
    mIdAlmacen = 1
    FrmManVenta.Show
End Sub

Sub PuntoItems(Con As ADODB.Connection)
    Set xCon = Con
    FrmManItems.Show
End Sub

Sub PuntoDocumentos(Con As ADODB.Connection)
    Set xCon = Con
    FrmControlDoc.Show
End Sub

Sub PuntoPersonal(Con As ADODB.Connection)
    Set xCon = Con
    FrmControlPersPtoVta.Show
End Sub
