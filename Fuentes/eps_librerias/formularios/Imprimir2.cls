VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Imprimir2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public xVS As VSPrinter7LibCtl.VSPrinter


Sub ImprimirReporte(xObj As Object)
    'With xVS
        
    'End With
    'FrmPrinter2.Vp = xVS
    Set xVS1 = xVS
    FrmPrinter2.Show
End Sub
