VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncionesGraficos"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum Tamaño
    'VALORES PARA BUSCAR EL TIPO DE CAMBIO
    T16x16 = 16
    T24x24 = 24
    T32x32 = 32
    T48x48 = 48
    T64x64 = 64
    T72x72 = 72
    T96x96 = 96
    T128x128 = 128
End Enum

Public Function LeerIcono(Path As String, xTamaño As Tamaño, Frm As Object, COLOR_MASCARA As Long) As Object
    Set LeerIcono = x_LeerIcono(Path, xTamaño, Frm, COLOR_MASCARA)
End Function

Sub CargarToolbar(xListaImagen As Object, xToolbar As Object, xFormulario As Object, xNumeroBotones As Integer, ArchivosImagen() As String, TextoToolTipText() As String, xTamaño As Tamaño)
    Dim I As Integer
    
    For I = 1 To UBound(ArchivosImagen)
        xListaImagen.ListImages.Add , , x_LeerIcono(ArchivosImagen(I), xTamaño, xFormulario, xFormulario.BackColor)
    Next I
    Dim xImg As Integer
    
    With xToolbar
        xToolbar.Buttons.Clear
        .ImageList = xListaImagen
        .Style = tbrFlat
        xImg = 1
        For I = 1 To xNumeroBotones
            If I = 4 Or I = 7 Or I = 12 Or I = 19 Then
                .Buttons.Add , , , 3
                xImg = xImg - 1
            Else
                .Buttons.Add , , , , xImg
                .Buttons(I).ToolTipText = TextoToolTipText(I)
            End If
            xImg = xImg + 1
        Next
    End With
End Sub


Sub ConfigurarSizerTool(xFormulario As Object, xContenedor As Object, xToolbar As Object, xTamaño As Tamaño)
    xFormulario.ScaleMode = 1
    
    If xTamaño = T16x16 Then
        xContenedor.Top = 370
    End If

    If xTamaño = T24x24 Then
        xContenedor.Top = 490
        
    End If

    If xTamaño = T32x32 Then
        xContenedor.Top = 40
        xContenedor.Left = 1
    End If

    If xTamaño = T48x48 Then
        xContenedor.Top = 860
    End If
    
    xFormulario.Height = xToolbar.Height + xContenedor.Height
    xFormulario.Width = xContenedor.Width
    xFormulario.ScaleMode = 3
End Sub



