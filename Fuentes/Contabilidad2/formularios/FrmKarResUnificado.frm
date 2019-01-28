VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form FrmKarResUnificado 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4710
      TabIndex        =   1
      Top             =   225
      Width           =   1545
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6990
      Left            =   0
      TabIndex        =   0
      Top             =   1110
      Width           =   11880
      _cx             =   20955
      _cy             =   12330
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmKarResUnificado.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "FrmKarResUnificado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xCon2 As New ADODB.Connection
Dim xConData As New ADODB.Connection

Sub Cargar()
    Dim Rst As New ADODB.Recordset
    Dim RstEmpresas As New ADODB.Recordset
    
    Dim A As Integer
    
    'AbriConeccion
    Set xCon2 = AbrirNuevaConeccion(AP_RUTABD + "data.mdb")
    Fg1.Rows = 2
    RST_Busq Rst, "SELECT alm_inventario.descripcion, alm_inventario.codpro, alm_inventario.id, mae_unidades.abrev, alm_inventario.tippro " _
        & " FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = 4)) " _
        & " ORDER BY alm_inventario.descripcion", xCon2

    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Rst("descripcion")
            Fg1.TextMatrix(A, 2) = Rst("id")
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
End Sub

Private Sub Command1_Click()
    Cargar
End Sub

Function AbrirNuevaConeccion(Ruta As String) As ADODB.Connection
    Dim xFun As New eps_librerias.FuncionesData
    Dim Con As New ADODB.Connection
    
    xFun.F_BASEDATOS = Ruta
    xFun.F_GRUPOTRABAJO = Trim(App.Path) + "\sigpyme.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    Set Con = xFun.AbrirConeccion
    Set xFun = Nothing
    
    Set AbrirNuevaConeccion = Con
End Function
