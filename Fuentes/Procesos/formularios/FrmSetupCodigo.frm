VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form FrmSetupCodigo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Herramientas - Definicion Codigo Autogenerado Sunat"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   2910
      Left            =   105
      TabIndex        =   1
      Top             =   705
      Width           =   5895
      _cx             =   10398
      _cy             =   5133
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmSetupCodigo.frx":0000
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Configuracion de Codigo Unicos S.U.N.A.T."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   255
      TabIndex        =   0
      Top             =   180
      Width           =   5595
   End
End
Attribute VB_Name = "FrmSetupCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim A As Integer
        
        SeEjecuto = True
        RST_Busq Rst, "SELECT var_codificacion.*, var_formatos.formato " _
            & " FROM var_codificacion LEFT JOIN var_formatos ON var_codificacion.iddato = var_formatos.id " _
            & " ORDER BY var_codificacion.orden", xCon

        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            Fg1.Rows = 1
            For A = 1 To Rst.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(A, 1) = Rst("orden")
                Fg1.TextMatrix(A, 2) = Rst("descripcion")
                Fg1.TextMatrix(A, 3) = Rst("formato")
                If Rst("activo") = -1 Then
                    Fg1.TextMatrix(A, 4) = -1
                Else
                    Fg1.TextMatrix(A, 4) = 0
                End If
                Fg1.TextMatrix(A, 5) = Rst("id")
                
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
    End If
End Sub

Private Sub Form_Load()
    Fg1.ColWidth(5) = 0
    SeEjecuto = False
End Sub

Private Sub VSFlexGrid1_Click()

End Sub
