Attribute VB_Name = "graficos"
Option Explicit

Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Sub InitCommonControls Lib "Comctl32" ()

Public Const LR_LOADFROMFILE = &H10
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2
Public Const IMAGE_ENHMETAFILE = 3
Public Const CF_BITMAP = 2

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long

' constantes DrawIconEx
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = DI_MASK Or DI_IMAGE

Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Enum T_TAMAÑO
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


Public Function x_LeerIcono(Path As String, xTamaño As T_TAMAÑO, Frm As Object, COLOR_MASCARA As Long) As Object
    Dim ANCHO_ICON, ALTO_ICON As Integer
    Dim mIcon As Long
    Dim PicTemp As PictureBox
    
    ANCHO_ICON = xTamaño
    ALTO_ICON = xTamaño
       
    Set PicTemp = Frm.Controls.Add("Vb.PictureBox", "Pic1")
    
    With PicTemp
        .Cls
        .ScaleMode = vbPixels
        .Width = ANCHO_ICON
        .Height = ALTO_ICON
        .BorderStyle = 0
        .AutoRedraw = True
        PicTemp.BackColor = COLOR_MASCARA
    End With
    
    mIcon = LoadImage(App.hInstance, Path, IMAGE_ICON, ANCHO_ICON, ALTO_ICON, LR_LOADFROMFILE)
    
    If mIcon <> 0 Then
        DrawIconEx PicTemp.hdc, 0, 0, mIcon, 0, 0, 0, 0, DI_NORMAL
        DestroyIcon mIcon
    Else
        MsgBox "Error", vbCritical
    End If
        
    PicTemp.Picture = PicTemp.Image
    Set x_LeerIcono = PicTemp.Picture
    
    Frm.Controls.Remove ("Pic1")
    Set PicTemp = Nothing
    
End Function

