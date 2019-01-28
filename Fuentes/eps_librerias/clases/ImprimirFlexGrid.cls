VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImprimirFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Public xFg1 As Object

Sub ImprimirFlex(xFg1 As Object, xTitulo1 As String, xTitulo2 As String, xNomEmp As String, xNumRuc As String)
    FrmPrinterFlex.NomEmp = xNomEmp
    FrmPrinterFlex.NumRUC = xNumRuc
    FrmPrinterFlex.Titulo1 = xTitulo1
    FrmPrinterFlex.Titulo2 = xTitulo2
    Set xFg = xFg1
    FrmPrinterFlex.Show vbModal
End Sub
