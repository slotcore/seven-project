VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmVistaEstacionalidad 
   Caption         =   "Produccion - Estacionalidad de las frutas"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNumGrid 
      Height          =   285
      Left            =   6855
      TabIndex        =   17
      Text            =   "TxtNumGrid"
      Top             =   3255
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg3 
      Height          =   795
      Left            =   30
      TabIndex        =   12
      Top             =   4845
      Width           =   11445
      _cx             =   20188
      _cy             =   1402
      _ConvInfo       =   1
      Appearance      =   0
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
      Rows            =   2
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmVistaEstacionalidad.frx":0000
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
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   3150
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11445
      _cx             =   20188
      _cy             =   5556
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColorSel    =   -2147483643
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
      Rows            =   2
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmVistaEstacionalidad.frx":0159
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
   Begin VSFlex7Ctl.VSFlexGrid Fg2 
      Height          =   795
      Left            =   30
      TabIndex        =   2
      Top             =   3750
      Width           =   11445
      _cx             =   20188
      _cy             =   1402
      _ConvInfo       =   1
      Appearance      =   0
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
      Rows            =   2
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmVistaEstacionalidad.frx":02D6
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
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   30
      TabIndex        =   4
      Top             =   5610
      Width           =   11595
      Begin VB.CommandButton CmdMenos 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4860
         TabIndex        =   15
         Top             =   195
         Width           =   420
      End
      Begin VB.CommandButton CmdMas 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4380
         TabIndex        =   14
         Top             =   195
         Width           =   420
      End
      Begin VB.CommandButton CmdAcep 
         Caption         =   "&Aceptar"
         Height          =   420
         Left            =   7470
         TabIndex        =   6
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton CmdCan 
         Caption         =   "&Cancelar"
         Height          =   420
         Left            =   8715
         TabIndex        =   5
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "+/- Ancho"
         Height          =   195
         Left            =   4395
         TabIndex        =   16
         Top             =   555
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   5625
         X2              =   5625
         Y1              =   180
         Y2              =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   5610
         X2              =   5610
         Y1              =   180
         Y2              =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Abundancia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2790
         TabIndex        =   9
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Regular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   930
         TabIndex        =   8
         Top             =   525
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Escaces"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   930
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         Height          =   225
         Left            =   1950
         Top             =   225
         Width           =   720
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         Height          =   225
         Left            =   120
         Top             =   510
         Width           =   720
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   225
         Left            =   120
         Top             =   225
         Width           =   720
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Produccion Optima Segun Estacionalidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   30
      TabIndex        =   13
      Top             =   4620
      Width           =   3525
   End
   Begin VB.Label LblProducto 
      AutoSize        =   -1  'True
      Caption         =   "LblProducto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2505
      TabIndex        =   11
      Top             =   3525
      Width           =   1035
   End
   Begin VB.Label LblFruta 
      AutoSize        =   -1  'True
      Caption         =   "LblFruta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2505
      TabIndex        =   10
      Top             =   3255
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Produccion Aproximada de :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   3525
      Width           =   2400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Estacionalidad de :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   30
      TabIndex        =   1
      Top             =   3255
      Width           =   1650
   End
End
Attribute VB_Name = "FrmVistaEstacionalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim xTotal As Double
Dim Agregando As Boolean

Private Sub CmdAcep_Click()
    'If Val(Fg2.TextMatrix(1, 13)) <> Val(Fg3.TextMatrix(1, 13)) Then
    '    MsgBox "Los montos finales del producto son diferentes en los cuadros produccion aproximada y produccion optima", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    '    Exit Sub
    'End If
    If Trim(TxtNumGrid.Text) = "2" Then
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 4) = Fg3.TextMatrix(1, 1)
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 5) = Fg3.TextMatrix(1, 2)
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 6) = Fg3.TextMatrix(1, 3)
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 7) = Fg3.TextMatrix(1, 4)
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 8) = Fg3.TextMatrix(1, 5)
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 9) = Fg3.TextMatrix(1, 6)
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 10) = Fg3.TextMatrix(1, 7)
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 11) = Fg3.TextMatrix(1, 8)
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 12) = Fg3.TextMatrix(1, 9)
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 13) = Fg3.TextMatrix(1, 10)
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 14) = Fg3.TextMatrix(1, 11)
        FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 15) = Fg3.TextMatrix(1, 12)
    Else
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 4) = Fg3.TextMatrix(1, 1)
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 5) = Fg3.TextMatrix(1, 2)
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 6) = Fg3.TextMatrix(1, 3)
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 7) = Fg3.TextMatrix(1, 4)
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 8) = Fg3.TextMatrix(1, 5)
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 9) = Fg3.TextMatrix(1, 6)
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 10) = Fg3.TextMatrix(1, 7)
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 11) = Fg3.TextMatrix(1, 8)
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 12) = Fg3.TextMatrix(1, 9)
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 13) = Fg3.TextMatrix(1, 10)
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 14) = Fg3.TextMatrix(1, 11)
        FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 15) = Fg3.TextMatrix(1, 12)
    End If
    
    CmdCan_Click
End Sub

Private Sub CmdCan_Click()
    Set Rst = Nothing
    Unload Me
End Sub

Private Sub CmdMas_Click()
    Fg2.ColWidth(1) = Fg2.ColWidth(1) - 10
    Fg2.ColWidth(2) = Fg2.ColWidth(2) - 10
    Fg2.ColWidth(3) = Fg2.ColWidth(3) - 10
    Fg2.ColWidth(4) = Fg2.ColWidth(4) - 10
    Fg2.ColWidth(5) = Fg2.ColWidth(5) - 10
    Fg2.ColWidth(6) = Fg2.ColWidth(6) - 10
    Fg2.ColWidth(7) = Fg2.ColWidth(7) - 10
    Fg2.ColWidth(8) = Fg2.ColWidth(8) - 10
    Fg2.ColWidth(9) = Fg2.ColWidth(9) - 10
    Fg2.ColWidth(10) = Fg2.ColWidth(10) - 10
    Fg2.ColWidth(11) = Fg2.ColWidth(11) - 10
    Fg2.ColWidth(12) = Fg2.ColWidth(12) - 10

    Fg3.ColWidth(1) = Fg3.ColWidth(1) - 10
    Fg3.ColWidth(2) = Fg3.ColWidth(2) - 10
    Fg3.ColWidth(3) = Fg3.ColWidth(3) - 10
    Fg3.ColWidth(4) = Fg3.ColWidth(4) - 10
    Fg3.ColWidth(5) = Fg3.ColWidth(5) - 10
    Fg3.ColWidth(6) = Fg3.ColWidth(6) - 10
    Fg3.ColWidth(7) = Fg3.ColWidth(7) - 10
    Fg3.ColWidth(8) = Fg3.ColWidth(8) - 10
    Fg3.ColWidth(9) = Fg3.ColWidth(9) - 10
    Fg3.ColWidth(10) = Fg3.ColWidth(10) - 10
    Fg3.ColWidth(11) = Fg3.ColWidth(11) - 10
    Fg3.ColWidth(12) = Fg3.ColWidth(12) - 10

End Sub

Private Sub CmdMenos_Click()
    Fg2.ColWidth(1) = Fg2.ColWidth(1) + 10
    Fg2.ColWidth(2) = Fg2.ColWidth(2) + 10
    Fg2.ColWidth(3) = Fg2.ColWidth(3) + 10
    Fg2.ColWidth(4) = Fg2.ColWidth(4) + 10
    Fg2.ColWidth(5) = Fg2.ColWidth(5) + 10
    Fg2.ColWidth(6) = Fg2.ColWidth(6) + 10
    Fg2.ColWidth(7) = Fg2.ColWidth(7) + 10
    Fg2.ColWidth(8) = Fg2.ColWidth(8) + 10
    Fg2.ColWidth(9) = Fg2.ColWidth(9) + 10
    Fg2.ColWidth(10) = Fg2.ColWidth(10) + 10
    Fg2.ColWidth(11) = Fg2.ColWidth(11) + 10
    Fg2.ColWidth(12) = Fg2.ColWidth(12) + 10

    Fg3.ColWidth(1) = Fg3.ColWidth(1) + 10
    Fg3.ColWidth(2) = Fg3.ColWidth(2) + 10
    Fg3.ColWidth(3) = Fg3.ColWidth(3) + 10
    Fg3.ColWidth(4) = Fg3.ColWidth(4) + 10
    Fg3.ColWidth(5) = Fg3.ColWidth(5) + 10
    Fg3.ColWidth(6) = Fg3.ColWidth(6) + 10
    Fg3.ColWidth(7) = Fg3.ColWidth(7) + 10
    Fg3.ColWidth(8) = Fg3.ColWidth(8) + 10
    Fg3.ColWidth(9) = Fg3.ColWidth(9) + 10
    Fg3.ColWidth(10) = Fg3.ColWidth(10) + 10
    Fg3.ColWidth(11) = Fg3.ColWidth(11) + 10
    Fg3.ColWidth(12) = Fg3.ColWidth(12) + 10

End Sub

Private Sub Fg1_RowColChange()
    If Agregando = True Then Exit Sub
    If Fg1.Rows = 1 Then Exit Sub
    LblFruta.Caption = Fg1.TextMatrix(Fg1.Row, 1)
    MostrarValoresEnEstacionalidad Fg1.Row
End Sub

Private Sub Fg3_CellChanged(ByVal Row As Long, ByVal Col As Long)
    SumarTotales
End Sub

Sub SumarTotales()
    Dim Total As Double
    Total = Total + Val(Fg3.TextMatrix(1, 1))
    Total = Total + Val(Fg3.TextMatrix(1, 2))
    Total = Total + Val(Fg3.TextMatrix(1, 3))
    Total = Total + Val(Fg3.TextMatrix(1, 4))
    Total = Total + Val(Fg3.TextMatrix(1, 5))
    Total = Total + Val(Fg3.TextMatrix(1, 6))
    Total = Total + Val(Fg3.TextMatrix(1, 7))
    Total = Total + Val(Fg3.TextMatrix(1, 8))
    Total = Total + Val(Fg3.TextMatrix(1, 9))
    Total = Total + Val(Fg3.TextMatrix(1, 10))
    Total = Total + Val(Fg3.TextMatrix(1, 11))
    Total = Total + Val(Fg3.TextMatrix(1, 12))
    
    Fg3.TextMatrix(1, 13) = Format(Total, "0.00")
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Dim A As Integer
        RST_Busq Rst, "SELECT * FROM mae_estacionalidad ORDER BY descripcion", xCon
        
        Rst.MoveFirst
        Fg1.Rows = 1
        Agregando = True
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Trim(Rst("descripcion"))
            
            MuestraEstacionalidad A
            Rst.MoveNext
            
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        Agregando = False
        LblFruta.Caption = Fg1.TextMatrix(1, 1)
        MuestraPlan
        MostrarValoresEnEstacionalidad 1
    End If
End Sub

Sub MuestraEstacionalidad(xFila As Integer)
    Dim A As Integer
    Dim xCol As Integer
    Dim NumMeses As Integer

    With Fg1
        If Rst("ene") = 1 Then .Select xFila, 2, xFila, 2: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 2) = "1"
        If Rst("feb") = 1 Then .Select xFila, 3, xFila, 3: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 3) = "1"
        If Rst("mar") = 1 Then .Select xFila, 4, xFila, 4: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 4) = "1"
        If Rst("abr") = 1 Then .Select xFila, 5, xFila, 5: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 5) = "1"
        If Rst("may") = 1 Then .Select xFila, 6, xFila, 6: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 6) = "1"
        If Rst("jun") = 1 Then .Select xFila, 7, xFila, 7: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 7) = "1"
        If Rst("jul") = 1 Then .Select xFila, 8, xFila, 8: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 8) = "1"
        If Rst("ago") = 1 Then .Select xFila, 9, xFila, 9: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 9) = "1"
        If Rst("set") = 1 Then .Select xFila, 10, xFila, 10: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 10) = "1"
        If Rst("oct") = 1 Then .Select xFila, 11, xFila, 11: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 11) = "1"
        If Rst("nov") = 1 Then .Select xFila, 12, xFila, 12: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 12) = "1"
        If Rst("dic") = 1 Then .Select xFila, 13, xFila, 13: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 13) = "1"
        
        If Rst("ene") = 2 Then .Select xFila, 2, xFila, 2: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 2) = "2"
        If Rst("feb") = 2 Then .Select xFila, 3, xFila, 3: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 3) = "2"
        If Rst("mar") = 2 Then .Select xFila, 4, xFila, 4: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 4) = "2"
        If Rst("abr") = 2 Then .Select xFila, 5, xFila, 5: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 5) = "2"
        If Rst("may") = 2 Then .Select xFila, 6, xFila, 6: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 6) = "2"
        If Rst("jun") = 2 Then .Select xFila, 7, xFila, 7: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 7) = "2"
        If Rst("jul") = 2 Then .Select xFila, 8, xFila, 8: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 8) = "2"
        If Rst("ago") = 2 Then .Select xFila, 9, xFila, 9: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 9) = "2"
        If Rst("set") = 2 Then .Select xFila, 10, xFila, 10: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 10) = "2"
        If Rst("oct") = 2 Then .Select xFila, 11, xFila, 11: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 11) = "2"
        If Rst("nov") = 2 Then .Select xFila, 12, xFila, 12: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 12) = "2"
        If Rst("dic") = 2 Then .Select xFila, 13, xFila, 13: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 13) = "2"
        
        If Rst("ene") = 3 Then .Select xFila, 2, xFila, 2: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 2) = "3"
        If Rst("feb") = 3 Then .Select xFila, 3, xFila, 3: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 3) = "3"
        If Rst("mar") = 3 Then .Select xFila, 4, xFila, 4: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 4) = "3"
        If Rst("abr") = 3 Then .Select xFila, 5, xFila, 5: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 5) = "3"
        If Rst("may") = 3 Then .Select xFila, 6, xFila, 6: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 6) = "3"
        If Rst("jun") = 3 Then .Select xFila, 7, xFila, 7: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 7) = "3"
        If Rst("jul") = 3 Then .Select xFila, 8, xFila, 8: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 8) = "3"
        If Rst("ago") = 3 Then .Select xFila, 9, xFila, 9: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 9) = "3"
        If Rst("set") = 3 Then .Select xFila, 10, xFila, 10: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 10) = "3"
        If Rst("oct") = 3 Then .Select xFila, 11, xFila, 11: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 11) = "3"
        If Rst("nov") = 3 Then .Select xFila, 12, xFila, 12: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 12) = "3"
        If Rst("dic") = 3 Then .Select xFila, 13, xFila, 13: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 13) = "3"
    End With
    
    If Rst("ene") = 2 Then NumMeses = NumMeses + 1
    If Rst("feb") = 2 Then NumMeses = NumMeses + 1
    If Rst("mar") = 2 Then NumMeses = NumMeses + 1
    If Rst("abr") = 2 Then NumMeses = NumMeses + 1
    If Rst("may") = 2 Then NumMeses = NumMeses + 1
    If Rst("jun") = 2 Then NumMeses = NumMeses + 1
    If Rst("jul") = 2 Then NumMeses = NumMeses + 1
    If Rst("ago") = 2 Then NumMeses = NumMeses + 1
    If Rst("set") = 2 Then NumMeses = NumMeses + 1
    If Rst("oct") = 2 Then NumMeses = NumMeses + 1
    If Rst("nov") = 2 Then NumMeses = NumMeses + 1
    If Rst("dic") = 2 Then NumMeses = NumMeses + 1
    
    Fg1.TextMatrix(xFila, Fg1.Cols - 1) = NumMeses
    'Agregando = False
End Sub

Sub MostrarValoresEnEstacionalidad(xFila As Integer)
    Dim xCanProMes As Double
    Dim xTotal As Double
    Dim xNumMeses As Integer
    
    xNumMeses = NulosN(Fg1.TextMatrix(xFila, Fg1.Cols - 1))
    xTotal = NulosN(Fg2.TextMatrix(1, Fg2.Cols - 1))
    xCanProMes = xTotal / xNumMeses

    'sumamos el total de la nueva cantidad calculada por mes
    xTotal = 0
    Fg3.TextMatrix(1, 1) = ""
    Fg3.TextMatrix(1, 2) = ""
    Fg3.TextMatrix(1, 3) = ""
    Fg3.TextMatrix(1, 4) = ""
    Fg3.TextMatrix(1, 5) = ""
    Fg3.TextMatrix(1, 6) = ""
    Fg3.TextMatrix(1, 7) = ""
    Fg3.TextMatrix(1, 8) = ""
    Fg3.TextMatrix(1, 9) = ""
    Fg3.TextMatrix(1, 10) = ""
    Fg3.TextMatrix(1, 11) = ""
    Fg3.TextMatrix(1, 12) = ""

    If Trim(Fg1.TextMatrix(xFila, 2)) = "2" Then Fg3.TextMatrix(1, 1) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes
    If Trim(Fg1.TextMatrix(xFila, 3)) = "2" Then Fg3.TextMatrix(1, 2) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes
    If Trim(Fg1.TextMatrix(xFila, 4)) = "2" Then Fg3.TextMatrix(1, 3) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes
    If Trim(Fg1.TextMatrix(xFila, 5)) = "2" Then Fg3.TextMatrix(1, 4) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes
    If Trim(Fg1.TextMatrix(xFila, 6)) = "2" Then Fg3.TextMatrix(1, 5) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes
    If Trim(Fg1.TextMatrix(xFila, 7)) = "2" Then Fg3.TextMatrix(1, 6) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes
    If Trim(Fg1.TextMatrix(xFila, 8)) = "2" Then Fg3.TextMatrix(1, 7) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes
    If Trim(Fg1.TextMatrix(xFila, 9)) = "2" Then Fg3.TextMatrix(1, 8) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes
    If Trim(Fg1.TextMatrix(xFila, 10)) = "2" Then Fg3.TextMatrix(1, 9) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes
    If Trim(Fg1.TextMatrix(xFila, 11)) = "2" Then Fg3.TextMatrix(1, 10) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes
    If Trim(Fg1.TextMatrix(xFila, 12)) = "2" Then Fg3.TextMatrix(1, 11) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes
    If Trim(Fg1.TextMatrix(xFila, 13)) = "2" Then Fg3.TextMatrix(1, 12) = Format(xCanProMes, "0.00"): xTotal = xTotal + xCanProMes

    Fg3.TextMatrix(1, 13) = Format(xTotal, "0.00")
End Sub

Sub MuestraPlan()
    Fg1.ColWidth(Fg1.Cols - 1) = 0
    If Trim(TxtNumGrid.Text) = "2" Then
        LblProducto.Caption = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 1)
        xTotal = 0
        Fg2.TextMatrix(1, 1) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 4)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 1))
        Fg2.TextMatrix(1, 2) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 5)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 2))
        Fg2.TextMatrix(1, 3) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 6)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 3))
        Fg2.TextMatrix(1, 4) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 7)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 4))
        Fg2.TextMatrix(1, 5) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 8)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 5))
        Fg2.TextMatrix(1, 6) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 9)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 6))
        Fg2.TextMatrix(1, 7) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 10)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 7))
        Fg2.TextMatrix(1, 8) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 11)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 8))
        Fg2.TextMatrix(1, 9) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 12)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 9))
        Fg2.TextMatrix(1, 10) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 13)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 10))
        Fg2.TextMatrix(1, 11) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 14)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 11))
        Fg2.TextMatrix(1, 12) = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 15)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 12))
    Else
        LblProducto.Caption = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 1)
        xTotal = 0
        Fg2.TextMatrix(1, 1) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 4)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 1))
        Fg2.TextMatrix(1, 2) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 5)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 2))
        Fg2.TextMatrix(1, 3) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 6)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 3))
        Fg2.TextMatrix(1, 4) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 7)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 4))
        Fg2.TextMatrix(1, 5) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 8)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 5))
        Fg2.TextMatrix(1, 6) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 9)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 6))
        Fg2.TextMatrix(1, 7) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 10)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 7))
        Fg2.TextMatrix(1, 8) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 11)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 8))
        Fg2.TextMatrix(1, 9) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 12)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 9))
        Fg2.TextMatrix(1, 10) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 13)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 10))
        Fg2.TextMatrix(1, 11) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 14)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 11))
        Fg2.TextMatrix(1, 12) = FrmPlanProduccion.Fg1.TextMatrix(FrmPlanProduccion.Fg1.Row, 15)
        xTotal = xTotal + Val(Fg2.TextMatrix(1, 12))
    End If
    Fg2.TextMatrix(1, 13) = Format(xTotal, "0.00")
    
    'LblProducto.Caption = FrmPlanProduccion.Fg2.TextMatrix(FrmPlanProduccion.Fg2.Row, 1)
    
    Fg3.Editable = flexEDKbdMouse
End Sub


Private Sub Form_Load()
    SeEjecuto = False
End Sub
