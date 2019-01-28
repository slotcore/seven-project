VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Codejock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TOOL_TAMAÑO_ICO
    'VALORES PARA BUSCAR EL TIPO DE CAMBIO
    I16x16 = 16
    I24x24 = 24
    I32x32 = 32
    I48x48 = 48
    I64x64 = 64
    I72x72 = 72
    I96x96 = 96
    I128x128 = 128
End Enum

Public BORRARMENU As Boolean

Sub CrearToolBar(xOpciones() As String, BarraDeComando As Object, ListaImagenes As Object, xTamaño As TOOL_TAMAÑO_ICO)
    If BORRARMENU = True Then
        BarraDeComando.ActiveMenuBar.Delete
    End If
    
    Dim A As Integer
    Dim ToolBar As CommandBar
    'Set ToolBar = CommandBars.Add("Standard", xtpBarTop)
    Set ToolBar = BarraDeComando.Add("Standar", xtpBarTop)
    
    For A = 0 To UBound(xOpciones)
        If xOpciones(A, 2) = 1 Then
            AddControl ToolBar.Controls, xtpControlButton, CLng(xOpciones(A, 0)), Trim(xOpciones(A, 1)), True, Trim(xOpciones(A, 3))
        Else
            AddControl ToolBar.Controls, xtpControlButton, CLng(xOpciones(A, 0)), Trim(xOpciones(A, 1)), False, Trim(xOpciones(A, 3))
        End If
        
    Next A
    
    Set BarraDeComando.Icons = ListaImagenes.Icons
    BarraDeComando.Options.UseDisabledIcons = True
    BarraDeComando.Options.SetIconSize False, xTamaño, xTamaño
    BarraDeComando.Options.UseFadedIcons = 0
    BarraDeComando.RecalcLayout
End Sub
