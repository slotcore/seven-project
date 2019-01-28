VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncionesFlexGrid"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Sub FlexFormatoCelda(Grid As Object, xRow As Long, xCol As Integer, _
                   Optional xForeColor As Long = &H800000, _
                   Optional xNegrita As Boolean = False, _
                   Optional xBackColor As Long = &HFFFFFF, _
                   Optional xCaption As String = "")
                   '--::x_ForeColor=negro
                   '--::x_BackColor=blanco
    
    '--DAR LA FUENTE A LA CELDA
    Grid.Row = xRow: Grid.Col = xCol
    If xCaption <> "" Then Grid.TextMatrix(xRow, xCol) = xCaption
    Grid.CellFontBold = xNegrita
    Grid.CellForeColor = xForeColor
    Grid.CellBackColor = xBackColor
End Sub


