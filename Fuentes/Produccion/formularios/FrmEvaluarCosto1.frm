VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEvaluarCosto1 
   Caption         =   "Producción - Costo de Producción de Personal"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   615
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11775
   Begin VB.Frame Frm4 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   3600
      Left            =   3360
      TabIndex        =   39
      Top             =   2910
      Visible         =   0   'False
      Width           =   6990
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   6690
         Picture         =   "FrmEvaluarCosto1.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   43
         ToolTipText     =   "Cerrar"
         Top             =   50
         Width           =   195
      End
      Begin VB.Frame Frame8 
         Height          =   495
         Left            =   70
         TabIndex        =   40
         Top             =   3000
         Width           =   6825
         Begin VB.CommandButton Cmd 
            Caption         =   "&Exportar Excel"
            Height          =   330
            Index           =   4
            Left            =   1140
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Personal"
            Top             =   135
            Width           =   1515
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Salir"
            Height          =   330
            Index           =   3
            Left            =   60
            TabIndex        =   41
            ToolTipText     =   "Agregar Personal"
            Top             =   135
            Width           =   1065
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   2640
         Index           =   1
         Left            =   90
         TabIndex        =   44
         Top             =   360
         Width           =   6795
         _cx             =   11986
         _cy             =   4657
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
         BackColorSel    =   128
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmEvaluarCosto1.frx":02EC
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
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   4
         X1              =   6960
         X2              =   6960
         Y1              =   0
         Y2              =   3570
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   6
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   7
         X1              =   0
         X2              =   6960
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tareas sin Costo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   45
         Top             =   45
         Width           =   1440
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   45
         Top             =   30
         Width           =   6870
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Seleccionar ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   4065
      TabIndex        =   18
      Top             =   360
      Width           =   6615
      Begin VB.OptionButton OptSeleccion 
         Caption         =   "x Personal"
         Height          =   225
         Index           =   1
         Left            =   900
         TabIndex        =   21
         Top             =   240
         Width           =   1065
      End
      Begin VB.OptionButton OptSeleccion 
         Caption         =   "x Area"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   0
         Left            =   2460
         Picture         =   "FrmEvaluarCosto1.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   180
         Width           =   195
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   2070
         MaxLength       =   12
         TabIndex        =   22
         Text            =   "txt_cb(0)"
         Top             =   150
         Width           =   615
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   2010
         X2              =   2010
         Y1              =   180
         Y2              =   450
      End
      Begin VB.Label lbl_cb_capt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
         Height          =   195
         Index           =   0
         Left            =   2070
         TabIndex        =   25
         Top             =   990
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lbl_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cod(0)"
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
         Height          =   300
         Index           =   0
         Left            =   4200
         TabIndex        =   24
         Top             =   150
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lbl_cb 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb(0)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   0
         Left            =   2715
         TabIndex        =   23
         Top             =   150
         Width           =   3825
      End
   End
   Begin VB.Frame FraTarea 
      BorderStyle     =   0  'None
      Height          =   3720
      Left            =   12330
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   6450
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   540
         Left            =   4890
         TabIndex        =   7
         Top             =   3060
         Width           =   1365
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   6180
         Picture         =   "FrmEvaluarCosto1.frx":0480
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   6
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg6 
         Height          =   2205
         Left            =   90
         TabIndex        =   8
         Top             =   780
         Width           =   6255
         _cx             =   11033
         _cy             =   3889
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
         BackColorSel    =   128
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   12
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmEvaluarCosto1.frx":076C
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
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblItem(1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   1
         Left            =   930
         TabIndex        =   15
         Top             =   315
         Width           =   5460
      End
      Begin VB.Line Line9 
         X1              =   120
         X2              =   6285
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "Producto: "
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
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   420
         Width           =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   6435
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   6435
         X2              =   6435
         Y1              =   -120
         Y2              =   4770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Listado de Tareas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   75
         TabIndex        =   13
         Top             =   60
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Costo Total:"
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
         Left            =   210
         TabIndex        =   12
         Top             =   3090
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Costo Linea:"
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
         Left            =   210
         TabIndex        =   11
         Top             =   3360
         Width           =   1080
      End
      Begin VB.Label LblTotal 
         Caption         =   "LblTotal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1410
         TabIndex        =   10
         Top             =   3090
         Width           =   1860
      End
      Begin VB.Label LblLinea 
         Caption         =   "LblLinea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1410
         TabIndex        =   9
         Top             =   3360
         Width           =   1860
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   270
         Index           =   2
         Left            =   30
         Top             =   15
         Width           =   6375
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6345
      Left            =   0
      TabIndex        =   31
      Top             =   960
      Width           =   11805
      _cx             =   20823
      _cy             =   11192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "   Horas   |   Destajo   |   Linea   "
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   0
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   6030
         Left            =   15
         TabIndex        =   32
         Top             =   15
         Width           =   11775
         _cx             =   20770
         _cy             =   10636
         _ConvInfo       =   1
         Appearance      =   1
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
         BackColorSel    =   128
         ForeColorSel    =   16777215
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmEvaluarCosto1.frx":0817
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
      Begin SizerOneLibCtl.TabOne TabOne2 
         Height          =   6030
         Left            =   12420
         TabIndex        =   33
         Top             =   15
         Width           =   11775
         _cx             =   20770
         _cy             =   10636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "     Detalle    |    Resumen    "
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   0
         Position        =   2
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   6030
            Left            =   330
            TabIndex        =   34
            Top             =   15
            Width           =   3570
            _cx             =   6297
            _cy             =   10636
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
            BackColorSel    =   128
            ForeColorSel    =   16777215
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
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEvaluarCosto1.frx":0A28
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
         Begin VSFlex7Ctl.VSFlexGrid Fg3 
            Height          =   6030
            Left            =   4845
            TabIndex        =   35
            Top             =   15
            Width           =   3570
            _cx             =   6297
            _cy             =   10636
            _ConvInfo       =   1
            Appearance      =   1
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
            BackColorSel    =   128
            ForeColorSel    =   16777215
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
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEvaluarCosto1.frx":0A51
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
      Begin SizerOneLibCtl.TabOne TabOne3 
         Height          =   6030
         Left            =   12720
         TabIndex        =   36
         Top             =   15
         Width           =   11775
         _cx             =   20770
         _cy             =   10636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "     Detalle    |    Resumen    "
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   0
         Position        =   2
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         Begin VSFlex7Ctl.VSFlexGrid Fg4 
            Height          =   6000
            Left            =   330
            TabIndex        =   37
            Top             =   15
            Width           =   11430
            _cx             =   20161
            _cy             =   10583
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
            BackColorSel    =   128
            ForeColorSel    =   16777215
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
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEvaluarCosto1.frx":0C62
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
         Begin VSFlex7Ctl.VSFlexGrid Fg5 
            Height          =   6000
            Left            =   12705
            TabIndex        =   38
            Top             =   15
            Width           =   11430
            _cx             =   20161
            _cy             =   10583
            _ConvInfo       =   1
            Appearance      =   1
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
            BackColorSel    =   128
            ForeColorSel    =   16777215
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
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEvaluarCosto1.frx":0C8B
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Seleccionar Fecha ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   0
      TabIndex        =   26
      Top             =   360
      Width           =   4005
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   300
         Index           =   0
         Left            =   645
         TabIndex        =   27
         Top             =   225
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "25/09/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   28
         Top             =   225
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "25/09/2007"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2040
         TabIndex        =   30
         Top             =   330
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   90
         TabIndex        =   29
         Top             =   330
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   555
      Left            =   10740
      TabIndex        =   16
      Top             =   360
      Width           =   1035
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   390
         Left            =   60
         TabIndex        =   17
         Top             =   90
         Width           =   945
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4860
         Top             =   90
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":0E9C
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":13E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":1772
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":18F6
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":1D4A
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":1E62
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":23A6
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":28EA
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":29FE
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":2B12
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":2F66
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":30D2
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":361A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEvaluarCosto1.frx":3934
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   3930
      TabIndex        =   1
      Top             =   8340
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   2
         Top             =   345
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando: Registros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   75
         Width           =   1890
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Interrumpir = ESC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   4140
         TabIndex        =   3
         Top             =   75
         Width           =   1530
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5745
         X2              =   5745
         Y1              =   -90
         Y2              =   4800
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   -60
         X2              =   6360
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   3
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmEvaluarCosto1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--HISTORIA
'--modificado:
'18/11/09 Johan Castro
'              cambiar la presentacion de la planilla por horas, antes se mostraba en formato hh:mm:ss AM/PM
'              ahora se muestra segun formato HH:MM:SS (Total Horas, Total HN, Total HE)
'              adicionalmente se cambia la presentacion a la planilla destajo(diferencia de horas)
'07/12/09 Johan Castro
'              Cálculo del turno noche, HE salia duplicado cuando se detallaba las tareas del personal, esto originaba
'              el pago exesivo al personal.
'12/12/09 Johan Castro
'              Mostrar el campo incentivos en pago por horas luego de haber modificado y grabado.
'20/05/11 Johan Castro
'              Considerar el proceso Lineas de produccion.
'23/05/11 Johan Castro
'              Agregar ElasticOne para ampliar el formulario
'05/10/11 Johan Castro
'              Revisar consultas cuando trabajado por horas ej. 08:00 hasta 01:00 del sig. dia, esto cuando son varios registros
'              Agregar filtro para mostrar como primer registro aquel que supere al ingreso 04:30 AM
'              Agregar consulta pra obtener la hora final cuando se trate de horarios otros, es decir cuando el horario abarque el siguiente dia
'20/10/11 Johan Castro
'              Eliminar lineas de codigo en evento grabar(), al eliminar registros cualdo se seleciona por area, permitiendo eliminar el registro cuando recorra persona por persona. Se aplica para hora, destajo y linea
'              Modificar filtro en sentencia SQL NulosN(Rst("idemp")), evento pCargarHoras() cuando se muestra datos de los insentivos
'21/10/11
'              Agregar campo en seleccion de personal CanReg, indica la cantidad de registros que tienen en el retalle. Esto servira para comprobar las horas de trabajo
'              Agregar variables HoraInicio, HoraFin cuando el calculo de horas trabajadas
'24/10/11
'              Modificar calculo de Horas, antes base=10 horas, ahora Base=8 horas el resto es como horas extras
'              Agregar orden a consulta de seleccion de lineas para agrupar por fecha,personal,hora inicio.

'11/11/11 Jose Chacon
'              Agregar cuadro informativo de Tareas sin costo tanto para Lineas como para Destajo
'06/01/12 Johan Castro
'              Agregar linea al grabar, eliminar registros segun seleccionen la fecha y seleccion de area y personal este vacio.
'21/01/12 Johan Castro
'              Agregar filtro de area al eliminar registro en evento grabar "& " and pro_pagos.idarea = " & NulosN(.TextMatrix(xFil, 10))"

'04/02/12 Jose Chacon
'              Agrega tipo en consulta  detallada de Destajo para insentivos

Option Explicit

Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
'--DE LA IMPRESION
Dim T_RPT_PERIODO As String '--PERIODO DEL REPORTE
Dim T_RPT_TITULO As String  '--TITULO DE REPORTE
'------------

Dim SeEjecuto  As Boolean
Dim Agregando  As Boolean
'Para mover el frame
Dim OrigFX As Long
Dim OrigFY As Long
'------------

Private Sub pConsultar()
'    On Error GoTo LaCague
    Dim rst_select As New ADODB.Recordset
    Dim nSQLSelect As String '--RECIBIR LA CONSULTA
    Dim mTipoConsulta As Integer '--valor para configurar el tipo de consulta y obtener el script sql
        
    If fValidarConsulta() = False Then Exit Sub
    BAND_INTERRUMPIR = False
    
    '***************************************************
    Frm4.Visible = False
    '***************************************************
    
    Me.MousePointer = vbHourglass
    DoEvents
    PosicionarProgBar
    
    lbl(0).Caption = "Procesando: Registros"
    lbl(1).Caption = "Interrumpir = ESC"
    Frm4.Visible = False
    FraTarea.Visible = False
    
    If TabOne1.CurrTab = 0 Then
        '--cargar las horas
        pCargarHoras
    ElseIf TabOne1.CurrTab = 1 Then
        '--cargar los destajos
        pCargarDestajo
    Else
        '--cargar los linea
        pCargarLinea
    
    End If
    '------------------------------------------------
    '*********************************************************************
    '*********************************************************************
   '
SALIR:
    FraProgreso.Visible = False
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
LaCague:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Set rst_select = Nothing
    MsgBox Err.Description & vbCr & Err.Source, vbCritical, xTitulo
    Err.Clear
End Sub

Private Sub pImprimir()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid IIf(TabOne1.CurrTab = 0, Fg1, Fg2), T_RPT_TITULO + " ", "", T_RPT_PERIODO, True, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    MsgBox Err.Description & vbCr & Err.Source & vbCr & Err.Number, vbCritical, xTitulo
    Err.Clear

End Sub

Private Sub cmd_Click(Index As Integer)
    If Index = 3 Then
        fg(1).Rows = fg(1).FixedRows
        Frm4.Visible = False
    End If
    
    If Index = 4 Then
        On Error GoTo error
        Dim X_PRINT As New SGI2_funciones.formularios
        Dim xTitulo As String
        Dim xPeriodo As String
        Dim xFg As VSFlexGrid
        
        Me.MousePointer = vbHourglass
        
        xPeriodo = "De " & TxtFecha(0).valor & " Al " & TxtFecha(1).valor
        
        If TabOne1.CurrTab = 0 Then '--horas
            xTitulo = "Detalle de Tareas sin Costo - Horas"
        ElseIf TabOne1.CurrTab = 1 Then '--destajo
            xTitulo = "Detalle de Tareas sin Costo - Destajo"
        ElseIf TabOne1.CurrTab = 2 Then '--linea
            xTitulo = "Detalle de Tareas sin Costo - Linea"
        End If
        
        Set xFg = fg(1)
        X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, xFg, xTitulo, xPeriodo, " ", xTitulo
        Set X_PRINT = Nothing
        Me.MousePointer = vbDefault
        Exit Sub
error:
        Me.MousePointer = vbDefault
        MsgBox Err.Description & vbCr & Err.Source & vbCr & Err.Number, vbCritical, xTitulo
        Err.Clear
    End If
End Sub

Private Sub CmdCancelar_Click()
    '--ocultar ventana
    FraTarea.Visible = False
    '--desbloquear objetos
    CmdGrabar.Enabled = True
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(6).Enabled = True
    
End Sub

Private Sub CmdGrabar_Click()
    '--0=Grabar Horas
    '--1=Grabar Destajo
    '--2=Grabar Linea
    pGrabar TabOne1.CurrTab
    
End Sub

Private Sub Fg4_DblClick()
    '--validar datos antes de invocar ventana emergente
    
    If Fg4.Row < Fg4.FixedRows Then Exit Sub
    If NulosN(Fg4.TextMatrix(Fg4.Row, 16)) = 0 Then
        
        Exit Sub
    End If
    
    '--bloquear controles
    
    '--limpiar objetos
    lblItem(1).Caption = Fg4.TextMatrix(Fg4.Row, 5)
    Fg6.Rows = Fg6.FixedRows
    lblTotal.Caption = "0.00"
    LblLinea.Caption = "0.00"
    LblLinea.ForeColor = &H800000
    '--mostrar ventana
    FraTarea.Visible = True
    
    FraTarea.Top = 1140
    FraTarea.Left = 2340
    
    
    '--definir variables
    Dim nSQL As String
    Dim xId As Double   '--Codigo de control de tareas
    Dim xCorr As Long   '--Correlativo
    Dim xIdRec As Double  '--Codigo de la receta
    Dim xIdUnimed As Long '--Codigio de Unidad de medidia
    Dim RstTarea As New ADODB.Recordset '--listado de todas las tareas del producto con sus costos
    Dim RstTareaLinea As New ADODB.Recordset '--listado de las tareas utilizados en la linea
    Dim xFila As Double '--indica el cambio de fila
    '--asignando datos a las variables
    
    xId = NulosN(Fg4.TextMatrix(Fg4.Row, 16))
    xCorr = NulosN(Fg4.TextMatrix(Fg4.Row, 17))
    xIdRec = NulosN(Fg4.TextMatrix(Fg4.Row, 18))
    xIdUnimed = NulosN(Fg4.TextMatrix(Fg4.Row, 19))
    
    '--limpiar datos
    '--definir sentencia SQL de todas las tareas con sus respectivos costos para un producto
    nSQL = "SELECT alm_inventario.descripcion AS producto, pro_tareas.descripcion AS tarea, pro_recetatar.orden, pro_recetatar.idtar, vwcosto.costo " _
        + vbCr + " FROM ((pro_receta INNER JOIN (pro_recetatar INNER JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id) ON pro_receta.id = pro_recetatar.idrec) INNER JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) " _
        + vbCr + " Left Join " _
        + vbCr + " ( SELECT pro_costo.idref, pro_costodet.idtar, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.costo " _
        + vbCr + " FROM pro_tareas INNER JOIN (pro_costo INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " Where (((pro_costo.idref) = " & xIdRec & ") And ((pro_costodet.idunimed) = " & xIdUnimed & ") And ((pro_costo.Tipo) = 1)) " _
        + vbCr + " ) AS vwcosto ON pro_recetatar.idtar = vwcosto.idtar " _
        + vbCr + " Where pro_receta.id = " & xIdRec & "" _
        + vbCr + " ORDER BY pro_recetatar.orden "
        
    '--ejecutar sentencia
    RST_Busq RstTarea, nSQL, xCon
    
    '--definir sentencia SQL de las tareas ingresadas en la linea para determinar solo las tareas activas
    nSQL = "SELECT pro_controltardettar.idctr, pro_controltardettar.idtar, pro_controltardettar.corr, pro_tareas.descripcion AS tarea, pro_controltardettar.activo " _
        + vbCr + " FROM pro_controltardettar LEFT JOIN pro_tareas ON pro_controltardettar.idtar = pro_tareas.id " _
        + vbCr + " WHERE (((pro_controltardettar.idctr)=" & xId & ") AND ((pro_controltardettar.corr)=" & xCorr & ")) and pro_controltardettar.activo = -1 "

    '--ejecutar sentencia
    RST_Busq RstTareaLinea, nSQL, xCon
    
    '--mostrar datos en grilla
    If RstTarea.RecordCount <> 0 Then
        RstTarea.MoveFirst
        Do While Not RstTarea.EOF
            Fg6.Rows = Fg6.Rows + 1
            Fg6.TextMatrix(Fg6.Rows - 1, 2) = NulosC(RstTarea("tarea"))
            Fg6.TextMatrix(Fg6.Rows - 1, 3) = Format(NulosN(RstTarea("costo")), "0.#00000000")
            Fg6.TextMatrix(Fg6.Rows - 1, 4) = NulosN(RstTarea("orden"))
            Fg6.TextMatrix(Fg6.Rows - 1, 5) = NulosN(RstTarea("idtar"))
            
            '--verificar si tarea esta activa para activar en grilla, caso contrario desactivar
            RstTareaLinea.Filter = ""
            RstTareaLinea.Filter = "idtar=" & NulosN(RstTarea("idtar"))
            If RstTareaLinea.RecordCount <> 0 Then
                Fg6.TextMatrix(Fg6.Rows - 1, 1) = -1 '--activo
                '--verificar que tarea activa tenga costo, caso contrario alertar con color rojo en la celda
                If NulosN(RstTarea("costo")) = 0 Then
                    GRID_COLOR_FONDO Fg6, Fg6.Rows - 1, 3, Fg6.Rows - 1, 4, vbRed
                    '--pintarde color rojo el costo total de la linea
                    LblLinea.ForeColor = vbRed
                End If
                '--acumular el costo de linea
                LblLinea.Caption = LblLinea.Caption + NulosN(RstTarea("costo"))
                
            Else
                Fg6.TextMatrix(Fg6.Rows - 1, 1) = 0 '--inactivo
            End If
            '--acumular el costo total
            lblTotal.Caption = lblTotal.Caption + NulosN(RstTarea("costo"))
            
            RstTarea.MoveNext
        Loop
    End If
    '--liberar variables del recordset
    Set RstTarea = Nothing
    Set RstTareaLinea = Nothing
    
    '--dar formato a los totales
    lblTotal.Caption = Format(lblTotal.Caption, "0.00000000")
    LblLinea.Caption = Format(LblLinea.Caption, "0.00000000")
    
    CmdCancelar.SetFocus
    
End Sub

Private Sub Fg5_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col <> 7 Then Exit Sub
    
    If IsNumeric(fg5.TextMatrix(Row, 7)) = False Then
        fg5.TextMatrix(Row, 8) = fg5.TextMatrix(Row, 6)
        fg5.TextMatrix(Row, 7) = 0
        fg5.SetFocus
        Exit Sub
    End If
    fg5.TextMatrix(Row, 7) = Format(fg5.TextMatrix(Row, 7), FORMAT_MONTO)
    
    fg5.TextMatrix(Row, 8) = NulosN(fg5.TextMatrix(Row, 7) + NulosN(fg5.TextMatrix(Row, 6)))
    
    '--totalizar
    fg5.TextMatrix(fg5.Rows - 1, 6) = Format(GRID_SUMAR_COL(fg5, 6) - NulosN(fg5.TextMatrix(fg5.Rows - 1, 6)), FORMAT_MONTO)
    fg5.TextMatrix(fg5.Rows - 1, 7) = Format(GRID_SUMAR_COL(fg5, 7) - NulosN(fg5.TextMatrix(fg5.Rows - 1, 7)), FORMAT_MONTO)
    fg5.TextMatrix(fg5.Rows - 1, 8) = Format(GRID_SUMAR_COL(fg5, 8) - NulosN(fg5.TextMatrix(fg5.Rows - 1, 8)), FORMAT_MONTO)


End Sub

Private Sub Fg5_EnterCell()
    If fg5.Col = 7 Then
        fg5.AutoSearch = flexSearchNone
        fg5.SelectionMode = flexSelectionFree
        fg5.Editable = flexEDKbdMouse
    Else
        fg5.AutoSearch = flexSearchFromTop
        fg5.SelectionMode = flexSelectionByRow
        fg5.Editable = flexEDNone
    End If
End Sub

Private Sub Form_Activate()
'    On Error GoTo error
    Dim mTipoConsulta As Integer
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = True
    
    
    TxtFecha(0).valor = Date
    TxtFecha(1).valor = Date
    txt_cb(0).Text = ""
    lbl_cb(0).Caption = ""
    lbl_cod(0).Caption = ""
    
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    TabOne3.CurrTab = 0
    '----------------------------
    pConfigurarGrilla
    TxtFecha(0).SetFocus
    Exit Sub
error:
    
    MsgBox Err.Description & vbCr & Err.Source, vbCritical, xTitulo
    Err.Clear
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    End If
End Sub

Private Sub Form_Load()

    SeEjecuto = False
    CentrarFrm Me
    Fg1.AutoSearch = flexSearchFromTop
    Fg2.AutoSearch = flexSearchFromTop
    Fg3.AutoSearch = flexSearchFromTop
    Fg4.AutoSearch = flexSearchFromTop
    fg5.AutoSearch = flexSearchFromTop
    Fg6.AutoSearch = flexSearchFromTop
    
    

    '---
    Me.WindowState = 2
    Me.Height = 7950
    Me.Width = 11910
    

    
End Sub


Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    TabOne1.Width = Me.Width - 150
    TabOne1.Top = 960
    If Me.Height > 1400 Then
        TabOne1.Height = Me.Height - 1400
    Else
        TabOne1.Height = Me.Height - 400
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    BAND_INTERRUMPIR = True
End Sub

'------
Private Function fValidarConsulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    '--DE LA FECHA ES NULL
    If TxtFecha(0).valor = "" Or TxtFecha(1).valor = "" Then
        MsgBox "Ingrese una fecha", vbExclamation, xTitulo
        If TxtFecha(0).valor = "" Then TxtFecha(0).SetFocus Else TxtFecha(1).SetFocus
        Exit Function
    End If
    If CDate(TxtFecha(0).valor) > CDate(TxtFecha(1).valor) Then
        MsgBox "La fecha inicial es superior al Final", vbExclamation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If
    fValidarConsulta = True
End Function

Private Sub PosicionarProgBar()
    '--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
    '    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    Me.PgBar.Value = 1
    FraProgreso.Visible = True
End Sub


'--------
'--------
Private Sub pExportarExcel()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Dim nTitulo As String
    Dim nPeriodo As String
    Dim xFg As VSFlexGrid
    
    Me.MousePointer = vbHourglass
    
    nPeriodo = "De " & TxtFecha(0).valor & " Al " & TxtFecha(1).valor
    
    If TabOne1.CurrTab = 0 Then '--horas
        nTitulo = "Costo de Personal - Horas"
        Set xFg = Fg1
    ElseIf TabOne1.CurrTab = 1 Then '--destajo
        If TabOne2.CurrTab = 0 Then '--detalle
            nTitulo = "Costo de Personal - Destajo Detalle"
            Set xFg = Fg2
        Else '--resumen
            nTitulo = "Costo de Personal - Destajo Resumen"
            Set xFg = Fg3
        End If
    ElseIf TabOne1.CurrTab = 2 Then '--linea
        If TabOne3.CurrTab = 0 Then '--detalle
            nTitulo = "Costo de Personal - Linea Detalle"
            Set xFg = Fg4
        Else '--resumen
            nTitulo = "Costo de Personal - Linea Resumen"
            Set xFg = fg5
        End If
    End If
    
    X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, xFg, nTitulo, nPeriodo, " ", nTitulo
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    MsgBox Err.Description & vbCr & Err.Source & vbCr & Err.Number, vbCritical, xTitulo
    Err.Clear
End Sub



Private Sub OptSeleccion_Click(Index As Integer)
    lbl_cb(0).Caption = ""
    txt_cb(0).Text = ""
    lbl_cod(0).Caption = ""
    If OptSeleccion(0).Value = True Then
        lbl_cb_capt(0).Caption = "Area"
    Else
        lbl_cb_capt(0).Caption = "Personal"
    End If
End Sub

Private Sub PbCerrar_Click(Index As Integer)
    If Index = 1 Then
        fg(1).Rows = fg(1).FixedRows
        Frm4.Visible = False
    End If
End Sub

Private Sub pic_Click(Index As Integer)
    CmdCancelar_Click
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    '*********************************************
    Frm4.Visible = False
    '*********************************************
End Sub

'************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 5 Then pExportarExcel
    If Button.Index = 6 Then pImprimir
    If Button.Index = 8 Then
        BAND_INTERRUMPIR = True
        Unload Me
    End If
End Sub

'************************************************


'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error
    Select Case Index
        Case 0 '--area
            If OptSeleccion(0).Value = True Then
                nTitulo = "Buscando Area"
                nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea " _
                    + vbCr + " FROM pro_area INNER JOIN mae_area ON pro_area.idarea = mae_area.id; "
            Else
                nTitulo = "Buscando Personal"
                nSQL = "SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.id AS cod " _
                    + vbCr + " FROM pla_empleados " _
                    + vbCr + " GROUP BY pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom], pla_empleados.id " _
                    + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"
            
            End If
            
            
    End Select
    
    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
    
    Dim RstTmp As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index).Text = NulosC(RstTmp.Fields(0))  '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
    lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
    lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1))  '--NOMBRE
      

SALIR:
    Set RstTmp = Nothing
Exit Sub
error:
    Set RstTmp = Nothing
    MsgBox Err.Description & vbCr & Err.Source & vbCr & Err.Number, vbCritical, xTitulo
    Err.Clear
End Sub


Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index <> 1 Then
            SendKeys vbTab
        Else
            If Fg1.Rows >= 2 Then
                Fg1.Row = 1: Fg1.Col = 1
            Else
                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 1
            End If
            Fg1.SetFocus
        End If
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)

    If txt_cb(Index).Text = "" Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--area
            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea " _
                + vbCr + " FROM mae_area where mae_area.id = " & NulosN(txt_cb(Index).Text)
        
        Case Else
            Exit Sub
            
    End Select

    If xCon.State = 0 Then GoTo SALIR
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))   '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
        lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
        lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1)) '--NOMBRE
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
SALIR:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub

Private Sub pCargarHoras()
    Dim Rst As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLFiltro As String
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        If OptSeleccion(0).Value = True Then
            nSQLFiltro = " and pro_controltar.idarea=" & NulosN(lbl_cod(0).Caption) & " "
        Else
            nSQLFiltro = " and pla_empleados.id=" & NulosN(lbl_cod(0).Caption) & " "
        End If
    End If
    
    Fg1.Rows = Fg1.FixedRows
    
    DoEvents
    
    '--consulta para determinar la lista de pagos por hora segun filtro seleccionado
    nSQL = "SELECT pla_empleados.id as idemp,mae_area.id as idarea, pro_controltar.fchtra, mae_area.descripcion AS area, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS personal, pla_empleados.paghornor, pla_empleados.paghorext,Count(pro_controltardet.idctr) AS canreg " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN ((pro_controltardet LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltar.tipo)=1)) and pro_controltardet.tipo = 1 " & nSQLFiltro _
        + vbCr + " GROUP BY pla_empleados.id, mae_area.id ,pro_controltar.fchtra, mae_area.descripcion, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom, pla_empleados.paghornor, pla_empleados.paghorext " _
        + vbCr + " ORDER BY pro_controltar.fchtra, mae_area.descripcion, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom; "
    
    '--falta agregar de detallegr
    
    RST_Busq Rst, nSQL, xCon
    
    Dim Difhora As String
    Dim HoraIniNoche As String
    Dim HoraFinNoche As String
    
    Dim HoraInicio As String
    Dim HoraFin As String
    
    Dim rstTmp1 As New ADODB.Recordset
    Dim rstBonif As New ADODB.Recordset '--registro de los incentivos que se le aplican
    
    Dim xHoraBase As String '--indica la hora base para aplicar calculo de HN ,HE
                            '--antes del 24/10/11  10:00:00
                            '--depues del 24/10/11 08:00:00
    
    If Rst.RecordCount = 0 Then Exit Sub
    Agregando = True
    PgBar.Min = 0
    PgBar.Value = 0
    PgBar.Max = Rst.RecordCount
    
    '---cargando listado de incentivos
    With Fg1
        Do While Not Rst.EOF
            DoEvents
            If BAND_INTERRUMPIR = True Then GoTo SALIR
            PgBar.Value = PgBar.Value + 1
            '------------------------------------------------
            Fg1.Rows = Fg1.Rows + 1
            .TextMatrix(Fg1.Rows - 1, 1) = Rst.Bookmark
            .TextMatrix(Fg1.Rows - 1, 2) = Format(Rst("fchtra"), FORMAT_DATE)
            .TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst("area"))
            .TextMatrix(Fg1.Rows - 1, 4) = NulosC(Rst("personal"))
            '--obteniendo la hora de inicio y de termino
            
            '--Establecer parametro de calculo para pago de horas
            If Rst("fchtra") <= CDate("23/10/2011") Then
                xHoraBase = "10:00:00"
            Else
                xHoraBase = "08:00:00"
            End If
            '--Si las horas de trabajo superan a la base se tomara como horas extras el resto.
            '--Ej. Horas trabajo = 11:20:00, NH=08:00:00, HE=03:20:00 cuando Base=08:00:00
            '--                              NH=10:00:00, HE=01:20:00 cuando Base=10:00:00
            
            nSQL = ""
            If NulosN(Rst("canreg")) > 1 Then
                nSQL = " And (pro_controltardet.horini)>=CDate('04:30')  "
            End If
            
            nSQL = "SELECT ini.fchtra, ini.idref, ini.hinipri, ini.hfinpri, fin.hiniult, fin.hfinult " _
                + vbCr + " From " _
                + vbCr + " (SELECT DISTINCT TOP 1 pro_controltar.fchtra, pro_controltardet.idref, pro_controltardet.horini AS hinipri, pro_controltardet.horfin AS hfinpri " _
                + vbCr + " FROM pro_controltar INNER JOIN pro_controltardet ON pro_controltar.id = pro_controltardet.idctr " _
                + vbCr + " WHERE (((pro_controltar.fchtra)=CDate('" & Rst("fchtra") & "')) AND ((pro_controltardet.idref)=" & NulosN(Rst("idemp")) & ") AND ((pro_controltar.tipo)=1)) and pro_controltardet.horini Is Not Null AND pro_controltardet.horfin Is Not Null " & nSQL _
                + vbCr + " ORDER BY pro_controltardet.horini ) AS ini "
                
            nSQL = nSQL & vbCr + " Left Join " _
                + vbCr + " (SELECT DISTINCT TOP 1 pro_controltar.fchtra, pro_controltardet.idref, pro_controltardet.horini AS hiniult, pro_controltardet.horfin AS hfinult " _
                + vbCr + " FROM pro_controltar INNER JOIN pro_controltardet ON pro_controltar.id = pro_controltardet.idctr " _
                + vbCr + " WHERE (((pro_controltar.fchtra)=CDate('" & Rst("fchtra") & "')) AND ((pro_controltardet.idref)=" & NulosN(Rst("idemp")) & ") AND ((pro_controltar.tipo)=1)) and pro_controltardet.horini Is Not Null AND pro_controltardet.horfin Is Not Null " _
                + vbCr + " ORDER BY pro_controltardet.horini desc " _
                + vbCr + " ) as fin " _
                + vbCr + " ON (ini.idref = fin.idref) AND (ini.fchtra = fin.fchtra);"
                        
            RST_Busq RstTmp, nSQL, xCon
            
            If RstTmp.RecordCount <> 0 Then
                '--evaluando las horas
                'Si la hora de inicio es mayor a la hora de termino de la tarea
                If (CDate(RstTmp("hinipri")) > CDate(RstTmp("hfinult"))) Then
                    Difhora = DiferenciaHoras(RstTmp("hinipri"), CDate("00:00"))
                    .TextMatrix(Fg1.Rows - 1, 5) = Format(RstTmp("hinipri"), FORMAT_HORA_SIN_SEGUNDO)
                    
                    '--identificar la ultima hora
                    nSQL = "SELECT TOP 1 pro_controltar.fchtra, pro_controltardet.idref, pro_controltardet.horini AS hiniult, pro_controltardet.horfin AS hfinult " _
                        + vbCr + " FROM pro_controltar INNER JOIN pro_controltardet ON pro_controltar.id = pro_controltardet.idctr " _
                        + vbCr + " WHERE (((pro_controltar.fchtra)=CDate('" & Rst("fchtra") & "')) AND ((pro_controltardet.idref)=" & NulosN(Rst("idemp")) & ") AND ((pro_controltar.tipo)=1)) and pro_controltardet.horini Is Not Null AND pro_controltardet.horfin Is Not Null And (pro_controltardet.horfin)<CDate('07:00')" _
                        + vbCr + " ORDER BY pro_controltardet.horfin DESC; "
                    
                    RST_Busq RstTmp, nSQL, xCon
                    
                    If RstTmp.RecordCount <> 0 Then
                    
                        Difhora = Format(CDate(Difhora) + CDate(DiferenciaHoras(CDate("00:00"), RstTmp("hfinult"))), "HH:mm")
                        .TextMatrix(Fg1.Rows - 1, 6) = Format(RstTmp("hfinult"), FORMAT_HORA_SIN_SEGUNDO)
                        .TextMatrix(Fg1.Rows - 1, 7) = "Otro"
                    Else
                        Difhora = ""
                    End If
                    
                Else
                    'Si los intervalos de tiempo estan en turno noche
                    If CDate(RstTmp("hfinult")) < CDate("10:00") Or CDate(RstTmp("hfinult")) > CDate("23:50") Or (CDate(RstTmp("hinipri")) > CDate("19:00")) Then
                    
                        '-----------------------------------------------------
                        If NulosN(Rst("canreg")) > 1 Then
                         '--obteniendo la ultima hora cuando una persona labora de noche
                         nSQL = "SELECT TOP 1 pro_controltardet.horfin " _
                             + vbCr + " FROM pro_controltar INNER JOIN pro_controltardet ON pro_controltar.id = pro_controltardet.idctr " _
                             + vbCr + " WHERE (((pro_controltardet.horfin) <= CDate('20:00')) AND ((pro_controltar.tipo)=1) AND ((pro_controltar.fchtra)=CDate('" & Rst("fchtra") & "')) AND ((pro_controltardet.idref)=" & Rst("idemp") & ")) " _
                             + vbCr + " ORDER BY pro_controltardet.horfin DESC;"
                             
                         RST_Busq rstTmp1, nSQL, xCon
                         If rstTmp1.RecordCount <> 0 Then
                             HoraFinNoche = NulosC(rstTmp1("horfin"))
                         End If
                         Set rstTmp1 = Nothing
                         
                        '--obteniendo la primera hora cuando una persona labora de noche
                         nSQL = "SELECT TOP 1 pro_controltardet.horini " _
                             + vbCr + " FROM pro_controltar INNER JOIN pro_controltardet ON pro_controltar.id = pro_controltardet.idctr " _
                             + vbCr + " WHERE pro_controltardet.horini >= CDate('10:00') and ((pro_controltar.tipo)=1) AND ((pro_controltar.fchtra)=CDate('" & Rst("fchtra") & "')) AND ((pro_controltardet.idref)=" & Rst("idemp") & ") " _
                             + vbCr + " ORDER BY pro_controltardet.horini asc;"
                             
                         RST_Busq rstTmp1, nSQL, xCon
                         If rstTmp1.RecordCount <> 0 Then
                             HoraIniNoche = NulosC(rstTmp1("horini"))
                         End If
                         Set rstTmp1 = Nothing
                        Else
                            HoraIniNoche = NulosC(RstTmp("hinipri"))
                            HoraFinNoche = NulosC(RstTmp("hfinpri"))
                        End If
                        
                        '-----------------------------------------------------
                        Difhora = DiferenciaHoras(HoraIniNoche, HoraFinNoche)
                        
                        .TextMatrix(Fg1.Rows - 1, 5) = Format(HoraIniNoche, FORMAT_HORA_SIN_SEGUNDO)
                        .TextMatrix(Fg1.Rows - 1, 6) = Format(HoraFinNoche, FORMAT_HORA_SIN_SEGUNDO)
                        .TextMatrix(Fg1.Rows - 1, 7) = "Noche"
                        
                    Else 'Si los intervalos de tiempo estan en turno dia
                        If NulosN(Rst("canreg")) > 1 Then
                            HoraInicio = NulosC(RstTmp("hinipri"))
                            HoraFin = NulosC(RstTmp("hfinult"))
                        Else
                            HoraInicio = NulosC(RstTmp("hinipri"))
                            HoraFin = NulosC(RstTmp("hfinpri"))
                        End If
                    
                        'Se calcula manualmente la diferencia de horas
                        Difhora = Format(CDate(HoraInicio) - CDate(HoraFin), "HH:mm")
                        
                        .TextMatrix(Fg1.Rows - 1, 5) = Format(HoraInicio, FORMAT_HORA_SIN_SEGUNDO)
                        .TextMatrix(Fg1.Rows - 1, 6) = Format(HoraFin, FORMAT_HORA_SIN_SEGUNDO)
                        .TextMatrix(Fg1.Rows - 1, 7) = "Dia"
                    End If
                End If
                
                .TextMatrix(Fg1.Rows - 1, 8) = Format(Difhora, FORMAT_HORA_LARGO) '--tot horas
                
                Dim h() As String
                Dim tiempo As Double
                '--si es de dia
                If .TextMatrix(Fg1.Rows - 1, 7) = "Dia" Then
                    If CDate(Difhora) > CDate(xHoraBase) Then
                    'If CDate(Difhora) > CDate("10:00") Then
                    
                        .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate(xHoraBase), FORMAT_HORA_LARGO) ' H.Normal
                        '.TextMatrix(Fg1.Rows - 1, 9) = Format(CDate("10:00"), FORMAT_HORA_LARGO) ' H.Normal
                        .TextMatrix(Fg1.Rows - 1, 10) = NulosN(Rst("paghornor")) 'Costo HN
                        .TextMatrix(Fg1.Rows - 1, 11) = Hour(CDate(xHoraBase)) * NulosN(Rst("paghornor")) ' Total HN
                        '.TextMatrix(Fg1.Rows - 1, 11) = 10 * NulosN(Rst("paghornor")) ' Total HN
                        
                        .TextMatrix(Fg1.Rows - 1, 12) = Format(CDate(CDate(Difhora) - CDate(xHoraBase)), FORMAT_HORA_LARGO)  ' H.Extra
                        '.TextMatrix(Fg1.Rows - 1, 12) = Format(CDate(CDate(Difhora) - CDate("10:00")), FORMAT_HORA_LARGO)  ' H.Extra
                        
                        .TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("paghorext")), FORMAT_MONTO) 'Costo HE
                        
                        'Se calcula el tiempo en horas formato decimales
                        h = Split(Format(.TextMatrix(Fg1.Rows - 1, 12), "HH:mm"), ":")
                        tiempo = Val(h(0)) + (Val(h(1)) / 60)
                        .TextMatrix(Fg1.Rows - 1, 14) = tiempo * NulosN(Rst("paghorext")) 'Total HE
                    Else
                        .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate(Difhora), FORMAT_HORA_LARGO) ' H.Normal
                        .TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(Rst("paghornor")), FORMAT_MONTO) 'Costo HN
                        .TextMatrix(Fg1.Rows - 1, 11) = Convert1HoraFaccion(Difhora) * NulosN(Rst("paghornor")) ' Total HN
                                                
                        .TextMatrix(Fg1.Rows - 1, 12) = "" 'Format(CDate("00:00"), FORMAT_HORA_SIN_SEGUNDO) ' H.Extra
                        .TextMatrix(Fg1.Rows - 1, 13) = NulosN(Rst("paghorext")) 'Costo HE
                        .TextMatrix(Fg1.Rows - 1, 14) = 0 'Total HE
                    End If
                End If
                '--si es de noche
                If .TextMatrix(Fg1.Rows - 1, 7) = "Noche" Then '--si es de dia
                    If Difhora <> "" Then
                        If CDate(Difhora) > CDate(xHoraBase) Then
                        'If CDate(Difhora) > CDate("18:00") Then
                        
                            .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate(xHoraBase), FORMAT_HORA_LARGO) ' H.Normal
                            '.TextMatrix(Fg1.Rows - 1, 9) = Format(CDate("10:00"), FORMAT_HORA_LARGO) ' H.Normal
                            .TextMatrix(Fg1.Rows - 1, 10) = NulosN(Rst("paghornor")) 'Costo HN
                            .TextMatrix(Fg1.Rows - 1, 11) = Hour(CDate(xHoraBase)) * NulosN(Rst("paghornor")) ' Total HN
                            '.TextMatrix(Fg1.Rows - 1, 11) = 10 * NulosN(Rst("paghornor")) ' Total HN
                                
                            .TextMatrix(Fg1.Rows - 1, 12) = Format(CDate(CDate(Difhora) - CDate(xHoraBase)), FORMAT_HORA_LARGO)  ' H.Extra
                            '.TextMatrix(Fg1.Rows - 1, 12) = Format(CDate(CDate(Difhora) - CDate("10:00")), FORMAT_HORA_LARGO)  ' H.Extra
                            .TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("paghorext")), FORMAT_MONTO) 'Costo HE
                            .TextMatrix(Fg1.Rows - 1, 14) = Convert1HoraFaccion(CDate(.TextMatrix(Fg1.Rows - 1, 12))) * NulosN(Rst("paghorext")) 'Total HE

                            
                        Else
                            .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate(Difhora), FORMAT_HORA_LARGO) ' H.Normal
                            .TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(Rst("paghornor")), FORMAT_MONTO) 'Costo HN
                            .TextMatrix(Fg1.Rows - 1, 11) = Convert1HoraFaccion(Difhora) * NulosN(Rst("paghornor")) ' Total HN
                            
                            .TextMatrix(Fg1.Rows - 1, 12) = "" 'Format(CDate("00:00"), FORMAT_HORA_SIN_SEGUNDO) ' H.Extra
                            .TextMatrix(Fg1.Rows - 1, 13) = NulosN(Rst("paghorext")) 'Costo HE
                            .TextMatrix(Fg1.Rows - 1, 14) = 0 'Total HE
                        End If
                    End If
                End If
                
                '--si es otro caso distinto
                If .TextMatrix(Fg1.Rows - 1, 7) = "Otro" Then
                    If CDate(Difhora) > CDate(xHoraBase) Then
                    'If CDate(Difhora) > CDate("10:00") Then
                    
                        .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate(xHoraBase), FORMAT_HORA_LARGO)  ' H.Normal
                        '.TextMatrix(Fg1.Rows - 1, 9) = Format(CDate("10:00"), FORMAT_HORA_LARGO) ' H.Normal
                        .TextMatrix(Fg1.Rows - 1, 10) = NulosN(Rst("paghornor")) 'Costo HN
                        .TextMatrix(Fg1.Rows - 1, 11) = Hour(CDate(xHoraBase)) * NulosN(Rst("paghornor")) ' Total HN
                        '.TextMatrix(Fg1.Rows - 1, 11) = 10 * NulosN(Rst("paghornor")) ' Total HN
                        
                        .TextMatrix(Fg1.Rows - 1, 12) = Format(CDate(CDate(Difhora) - CDate(xHoraBase)), FORMAT_HORA_LARGO)   ' H.Extra
                        '.TextMatrix(Fg1.Rows - 1, 12) = Format(CDate(CDate(Difhora) - CDate("10:00")), FORMAT_HORA_LARGO)  ' H.Extra
                        .TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("paghorext")), FORMAT_MONTO) 'Costo HE
                        .TextMatrix(Fg1.Rows - 1, 14) = Convert1HoraFaccion(CDate(.TextMatrix(Fg1.Rows - 1, 12))) * NulosN(Rst("paghorext")) 'Total HE
                        
                    Else
                        .TextMatrix(Fg1.Rows - 1, 9) = Format(CDate(Difhora), FORMAT_HORA_LARGO) ' H.Normal
                        .TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(Rst("paghornor")), FORMAT_MONTO) 'Costo HN
                        .TextMatrix(Fg1.Rows - 1, 11) = Convert1HoraFaccion(Difhora) * NulosN(Rst("paghornor")) ' Total HN
                                                
                        .TextMatrix(Fg1.Rows - 1, 12) = "" 'Format(CDate("00:00"), FORMAT_HORA_SIN_SEGUNDO) ' H.Extra
                        .TextMatrix(Fg1.Rows - 1, 13) = NulosN(Rst("paghorext")) 'Costo HE
                        .TextMatrix(Fg1.Rows - 1, 14) = 0 'Total HE
                    End If
                End If
                
            End If
            
            .TextMatrix(.Rows - 1, 11) = Format(.TextMatrix(Fg1.Rows - 1, 11), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, 14) = Format(.TextMatrix(Fg1.Rows - 1, 14), FORMAT_MONTO)
            
            .TextMatrix(.Rows - 1, 15) = NulosN(.TextMatrix(Fg1.Rows - 1, 11)) + NulosN(.TextMatrix(Fg1.Rows - 1, 14)) 'Tot Pagar
            .TextMatrix(.Rows - 1, 15) = Format(.TextMatrix(Fg1.Rows - 1, 15), FORMAT_MONTO)
            
            '--copiando los datos del pago total
            '-------------------------------------------'-------------------------------------------'-------------------------------------------
            '--incentivos
            nSQL = "SELECT pro_pagos.imptot, pro_pagos.impbon " _
                & " From pro_pagos " _
                & " WHERE (((pro_pagos.idemp)=" & NulosN(Rst("idemp")) & ") AND ((pro_pagos.idarea)=" & Rst("idarea") & ") AND ((pro_pagos.fchtra)=cdate('" & Rst("fchtra") & "')) AND ((pro_pagos.tipo)=1));"
            
            RST_Busq rstBonif, nSQL, xCon
            
            If rstBonif.RecordCount <> 0 Then
                .TextMatrix(.Rows - 1, 16) = NulosN(rstBonif("impbon"))
            Else
                .TextMatrix(.Rows - 1, 16) = 0
            End If
            
            Set rstBonif = Nothing
            '-------------------------------------------
            .TextMatrix(.Rows - 1, 17) = Format(NulosN(.TextMatrix(.Rows - 1, 15)) + NulosN(.TextMatrix(.Rows - 1, 16)), FORMAT_MONTO)
            
            .TextMatrix(.Rows - 1, 18) = NulosN(Rst("idarea"))
            .TextMatrix(.Rows - 1, 19) = NulosN(Rst("idemp"))
            
            Set RstTmp = Nothing
            
            Rst.MoveNext
        Loop
    End With
    Set Rst = Nothing
    Set RstTmp = Nothing
    '----------
    GRID_AGRUPAR Fg1, 3
    
    Dim mRow As Long
    Dim xFlat As Boolean
    
    For mRow = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 17)) = 0 Then
            GRID_COLOR_FONDO Fg1, mRow, 1, mRow, 17, vbRed
            xFlat = True
        End If
    Next
    If xFlat = True Then MsgBox "Se presentaron registros observados", vbInformation, xTitulo
                
    '--colocando los totales
    Fg1.Rows = Fg1.Rows + 1
    Fg1.TextMatrix(Fg1.Rows - 1, 4) = "Totales"
    Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(GRID_SUMAR_COL(Fg1, 11), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(GRID_SUMAR_COL(Fg1, 14), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(GRID_SUMAR_COL(Fg1, 15), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(GRID_SUMAR_COL(Fg1, 17), FORMAT_MONTO)
    
    GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, vbGreen
    
SALIR:
Agregando = False
    
End Sub



Private Sub pConfigurarGrilla()
    '===================================================================================================
    'Propósito: Establecer los encabezados del grid
    '
    'Entradas:  Ninguno
    '
    'Resultados: Grilla con Encabezado
    '===================================================================================================
    Dim k As Integer
    
    Agregando = True
    
    '-------------------------------------------------------------------------------------------------------------
    '-----Pago por Horas
    '-------------------------------------------------------------------------------------------------------------
    With Fg1
       
        .Cols = 20
        .Rows = 1
        
        .ColWidth(0) = 200
        .ColWidth(1) = 0
        .FrozenCols = 4
        .TextMatrix(0, 1) = "Nº":           .ColWidth(1) = 450:         .ColAlignment(1) = flexAlignRightBottom:       .Row = 0: .Col = 1: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 2) = "Fecha":        .ColWidth(2) = 800:      .ColAlignment(2) = flexAlignCenterCenter:         .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Area":         .ColWidth(3) = 900:       .ColAlignment(3) = flexAlignLeftBottom:       .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(0, 4) = "Personal":     .ColWidth(4) = 1500:       .ColAlignment(4) = flexAlignLeftBottom:         .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(0, 5) = "Hor. Inicio":  .ColWidth(5) = 900:      .ColAlignment(5) = flexAlignCenterCenter:         .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterCenter
        '--------
        .TextMatrix(0, 6) = "Hor. Fin":     .ColWidth(6) = 900:      .ColAlignment(6) = flexAlignCenterCenter:         .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(0, 7) = "Horario":       .ColWidth(7) = 700:     .ColAlignment(7) = flexAlignLeftBottom:        .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(0, 8) = "Tot. Horas":   .ColWidth(8) = 900:      .ColAlignment(8) = flexAlignCenterCenter:         .Row = 0: .Col = 8: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 9) = "H.Normal":     .ColWidth(9) = 900:       .ColAlignment(9) = flexAlignCenterCenter:       .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 10) = "Costo HN":     .ColWidth(10) = 800:     .ColAlignment(10) = flexAlignRightCenter:       .Row = 0: .Col = 10: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 11) = "Total HN":    .ColWidth(11) = 800:      .ColAlignment(11) = flexAlignRightBottom:       .Row = 0: .Col = 11: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 12) = "H.Extra":     .ColWidth(12) = 900:      .ColAlignment(12) = flexAlignCenterCenter:      .Row = 0: .Col = 12: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 13) = "Costo HE":    .ColWidth(13) = 800:      .ColAlignment(13) = flexAlignRightCenter:        .Row = 0: .Col = 13: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 14) = "Total HE":    .ColWidth(14) = 800:      .ColAlignment(14) = flexAlignRightCenter:        .Row = 0: .Col = 14: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 15) = "Tot Pagar":   .ColWidth(15) = 900:      .ColAlignment(15) = flexAlignRightCenter:        .Row = 0: .Col = 15: .CellAlignment = flexAlignRightCenter
        
        
        .TextMatrix(0, 16) = "Incentivos":   .ColWidth(16) = 900:     .ColAlignment(16) = flexAlignRightCenter:        .Row = 0: .Col = 16: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 17) = "Neto Pagar":   .ColWidth(17) = 900:     .ColAlignment(17) = flexAlignRightCenter:        .Row = 0: .Col = 17: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 18) = "IdArea":   .ColWidth(18) = 0:
        
        .TextMatrix(0, 19) = "IdEmp":   .ColWidth(19) = 0:
        
    End With
    '-------------------------------------------------------------------------------------------------------------
    '-----Pago por Destajo
    '-------------------------------------------------------------------------------------------------------------
    With Fg2 '--detalle
        .Cols = 23
        .Rows = 2
        .FixedRows = 2
        
        .ColWidth(0) = 200
        .ColWidth(1) = 0
        .FrozenCols = 6
        
        
        GRID_COMBINAR Fg2, 0, 1, 0, 11, "Información de Trabajo", flexAlignLeftCenter, True, , , &HD8E9EC, True
        GRID_COMBINAR Fg2, 0, 12, 1, 17, "Eficiencia", flexAlignLeftCenter, False, , , &HD8E9EC, True
        GRID_COMBINAR Fg2, 0, 19, 0, 22, "Pago", flexAlignLeftCenter, True, , , &HD8E9EC, True
        
        
        .TextMatrix(1, 2) = "Fecha":        .ColWidth(2) = 800:      .ColAlignment(2) = flexAlignCenterCenter:         .Row = 1: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 3) = "Area":         .ColWidth(3) = 450:       .ColAlignment(3) = flexAlignLeftBottom:       .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 4) = "Personal":     .ColWidth(4) = 1200:       .ColAlignment(4) = flexAlignLeftBottom:         .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 5) = "Tarea":  .ColWidth(5) = 1500:             .ColAlignment(5) = flexAlignLeftBottom:         .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftBottom
        '--------
        .TextMatrix(1, 6) = "Producto":    .ColWidth(6) = 1800:       .ColAlignment(6) = flexAlignLeftBottom:         .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftBottom
        
        
        .TextMatrix(1, 7) = "Observación":    .ColWidth(7) = 1200:       .ColAlignment(7) = flexAlignLeftBottom:    .Row = 1: .Col = 7: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(1, 8) = "H.Inicio":    .ColWidth(8) = 800:    .ColAlignment(8) = flexAlignLeftBottom:       .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 9) = "H.Final":     .ColWidth(9) = 800:        .ColAlignment(9) = flexAlignLeftBottom:         .Row = 1: .Col = 9: .CellAlignment = flexAlignLeftBottom
        
        
        .TextMatrix(1, 10) = "Cant.":     .ColWidth(10) = 700:            .ColAlignment(10) = flexAlignRightBottom:         .Row = 1: .Col = 10: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 11) = "U.M.":     .ColWidth(11) = 450:            .ColAlignment(11) = flexAlignCenterCenter:       .Row = 1: .Col = 11: .CellAlignment = flexAlignCenterCenter
        
        
        .TextMatrix(1, 12) = "Dif.Hora":        .ColWidth(12) = 800:       .ColAlignment(12) = flexAlignRightCenter:      .Row = 1: .Col = 12: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 13) = "Tot.Min":         .ColWidth(13) = 600:     .ColAlignment(13) = flexAlignRightCenter:        .Row = 1: .Col = 13: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 14) = "Unid x Min":      .ColWidth(14) = 0:       .ColAlignment(14) = flexAlignRightCenter:        .Row = 1: .Col = 14: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 15) = "Unid x Hor":      .ColWidth(15) = 950:    .ColAlignment(15) = flexAlignRightCenter:         .Row = 1: .Col = 15: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 16) = "Cant Teo":        .ColWidth(16) = 730:     .ColAlignment(16) = flexAlignRightCenter:        .Row = 1: .Col = 16: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 17) = "%":      .ColWidth(17) = 800:     .ColAlignment(17) = flexAlignRightCenter:        .Row = 1: .Col = 17: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 18) = " ":           .ColWidth(18) = 0:       .ColAlignment(18) = flexAlignRightCenter:       .Row = 1: .Col = 18: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 19) = "Cant":        .ColWidth(19) = 800:       .ColAlignment(19) = flexAlignRightCenter:       .Row = 1: .Col = 19: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 20) = "U.M.":        .ColWidth(20) = 450:       .ColAlignment(20) = flexAlignCenterCenter:      .Row = 1: .Col = 20: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 21) = "Pre.Uni":     .ColWidth(21) = 800:       .ColAlignment(21) = flexAlignRightCenter:       .Row = 1: .Col = 21: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 22) = "Total":       .ColWidth(22) = 750:       .ColAlignment(22) = flexAlignRightBottom:       .Row = 1: .Col = 22: .CellAlignment = flexAlignRightBottom
        
    End With
            
            
    With Fg3 '--destajo resumen
        '-----
        .Cols = 11
        .Rows = 1
        
        .ColWidth(0) = 200
        .ColWidth(1) = 0
        .FrozenCols = 4
        .TextMatrix(0, 1) = "Nº":           .ColWidth(1) = 450:      .ColAlignment(1) = flexAlignRightBottom:       .Row = 0: .Col = 1: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 2) = "Fecha":        .ColWidth(2) = 800:      .ColAlignment(2) = flexAlignCenterCenter:         .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Area":         .ColWidth(3) = 900:      .ColAlignment(3) = flexAlignLeftBottom:       .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(0, 4) = "Personal":     .ColWidth(4) = 3500:     .ColAlignment(4) = flexAlignLeftBottom:         .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(0, 5) = "Horario":      .ColWidth(5) = 700:      .ColAlignment(5) = flexAlignLeftBottom:         .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(0, 6) = "Total":        .ColWidth(6) = 900:     .ColAlignment(6) = flexAlignRightBottom:       .Row = 0: .Col = 6: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 7) = "Incentivos":   .ColWidth(7) = 900:     .ColAlignment(7) = flexAlignRightBottom:      .Row = 0: .Col = 7: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 8) = "Neto Pagar":   .ColWidth(8) = 900:     .ColAlignment(8) = flexAlignRightCenter:        .Row = 0: .Col = 8: .CellAlignment = flexAlignRightCenter
        
        
        
        .TextMatrix(0, 9) = "IdEmp":   .ColWidth(9) = 0:
        .TextMatrix(0, 10) = "IdArea":   .ColWidth(10) = 0:
                                                
    End With
    
    '-------------------------------------------------------------------------------------------------------------
    '-----Pago por Linea
    '-------------------------------------------------------------------------------------------------------------
    With Fg4 '--Linea Detalle
        .Cols = 20
        .Rows = 2
        .FixedRows = 2
        
        .ColWidth(0) = 200
        .ColWidth(1) = 0
'        .FrozenCols = 6
        
        
        GRID_COMBINAR Fg4, 0, 1, 0, 11, "Información de Trabajo", flexAlignLeftCenter, True, , , &HD8E9EC, True
        GRID_COMBINAR Fg4, 0, 12, 0, 15, "Pago", flexAlignLeftCenter, True, , , &HD8E9EC, True
        
        
        .TextMatrix(1, 2) = "Fecha":        .ColWidth(2) = 800:     .ColAlignment(2) = flexAlignCenterCenter:       .Row = 1: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 3) = "Area":         .ColWidth(3) = 450:     .ColAlignment(3) = flexAlignLeftBottom:         .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 4) = "Personal":     .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftBottom:         .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 5) = "Producto":     .ColWidth(5) = 1800:    .ColAlignment(5) = flexAlignLeftBottom:         .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 6) = "Observación":  .ColWidth(6) = 400:     .ColAlignment(6) = flexAlignLeftBottom:         .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 7) = "H.Inicio":     .ColWidth(7) = 800:     .ColAlignment(7) = flexAlignLeftBottom:         .Row = 1: .Col = 7: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 8) = "H.Final":      .ColWidth(8) = 800:     .ColAlignment(8) = flexAlignLeftBottom:         .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 9) = "Cant.":        .ColWidth(9) = 700:     .ColAlignment(9) = flexAlignRightBottom:        .Row = 1: .Col = 9: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 10) = "U.M.":        .ColWidth(10) = 450:    .ColAlignment(10) = flexAlignCenterCenter:      .Row = 1: .Col = 10: .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(1, 11) = " ":           .ColWidth(11) = 0:      .ColAlignment(11) = flexAlignRightCenter:       .Row = 1: .Col = 11: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 12) = "Cant":        .ColWidth(12) = 800:    .ColAlignment(12) = flexAlignRightCenter:       .Row = 1: .Col = 12: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 13) = "U.M.":        .ColWidth(13) = 450:    .ColAlignment(13) = flexAlignCenterCenter:      .Row = 1: .Col = 13: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 14) = "Pre.Uni":     .ColWidth(14) = 1000:    .ColAlignment(14) = flexAlignRightCenter:       .Row = 1: .Col = 14: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 15) = "Total":       .ColWidth(15) = 750:    .ColAlignment(15) = flexAlignRightBottom:       .Row = 1: .Col = 15: .CellAlignment = flexAlignRightBottom
                                                
        .TextMatrix(1, 16) = "idctr":        .ColWidth(16) = 0:
        .TextMatrix(1, 17) = "corr":         .ColWidth(17) = 0:
        .TextMatrix(1, 18) = "IdRec":        .ColWidth(18) = 0:
        .TextMatrix(1, 19) = "IdUniMed":     .ColWidth(19) = 0:
        
    End With
            
            
    With fg5 '--linea resumen
        '-----
        .Cols = 11
        .Rows = 1
        
        .ColWidth(0) = 200
        .ColWidth(1) = 0
        .FrozenCols = 4
        .TextMatrix(0, 1) = "Nº":           .ColWidth(1) = 450:         .ColAlignment(1) = flexAlignRightBottom:       .Row = 0: .Col = 1: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 2) = "Fecha":        .ColWidth(2) = 800:      .ColAlignment(2) = flexAlignCenterCenter:         .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Area":         .ColWidth(3) = 900:       .ColAlignment(3) = flexAlignLeftBottom:       .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(0, 4) = "Personal":     .ColWidth(4) = 3500:       .ColAlignment(4) = flexAlignLeftBottom:         .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(0, 5) = "Horario":      .ColWidth(5) = 700:       .ColAlignment(5) = flexAlignLeftBottom:         .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(0, 6) = "Total":        .ColWidth(6) = 900:      .ColAlignment(6) = flexAlignRightBottom:       .Row = 0: .Col = 6: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 7) = "Incentivos":   .ColWidth(7) = 900:     .ColAlignment(7) = flexAlignRightBottom:      .Row = 0: .Col = 7: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, 8) = "Neto Pagar":   .ColWidth(8) = 900:     .ColAlignment(8) = flexAlignRightCenter:        .Row = 0: .Col = 8: .CellAlignment = flexAlignRightCenter
        
        
        
        .TextMatrix(0, 9) = "IdEmp":   .ColWidth(9) = 0:
        .TextMatrix(0, 10) = "IdArea":   .ColWidth(10) = 0:
                                                
    End With
    
    '--detalle de tareas por linea
    Fg6.ColWidth(5) = 0
    
    Agregando = False
    
    DoEvents
End Sub




Private Sub Fg1_EnterCell()

    If Fg1.Col = 16 Then
        Fg1.AutoSearch = flexSearchNone
        Fg1.SelectionMode = flexSelectionFree
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.AutoSearch = flexSearchFromTop
        Fg1.SelectionMode = flexSelectionByRow
        Fg1.Editable = flexEDNone
    End If
    
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col <> 16 Then Exit Sub
    
    If IsNumeric(Fg1.TextMatrix(Row, 16)) = False Then
        Fg1.TextMatrix(Row, 17) = Fg1.TextMatrix(Row, 15)
        Fg1.TextMatrix(Row, 16) = 0
        Fg1.SetFocus
        Exit Sub
    End If
    Fg1.TextMatrix(Row, 16) = Format(Fg1.TextMatrix(Row, 16), FORMAT_MONTO)
    
    Fg1.TextMatrix(Row, 17) = NulosN(Fg1.TextMatrix(Row, 16) + NulosN(Fg1.TextMatrix(Row, 15)))
    
    '--totalizar
    Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(GRID_SUMAR_COL(Fg1, 16) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 16)), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(GRID_SUMAR_COL(Fg1, 17) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 17)), FORMAT_MONTO)
    
End Sub


Private Sub Fg3_EnterCell()

    If Fg3.Col = 7 Then
        Fg3.AutoSearch = flexSearchNone
        Fg3.SelectionMode = flexSelectionFree
        Fg3.Editable = flexEDKbdMouse
    Else
        Fg3.AutoSearch = flexSearchFromTop
        Fg3.SelectionMode = flexSelectionByRow
        Fg3.Editable = flexEDNone
    End If
    
End Sub

Private Sub Fg3_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col <> 7 Then Exit Sub
    
    If IsNumeric(Fg3.TextMatrix(Row, 7)) = False Then
        Fg3.TextMatrix(Row, 8) = Fg3.TextMatrix(Row, 6)
        Fg3.TextMatrix(Row, 7) = 0
        Fg3.SetFocus
        Exit Sub
    End If
    Fg3.TextMatrix(Row, 7) = Format(Fg3.TextMatrix(Row, 7), FORMAT_MONTO)
    
    Fg3.TextMatrix(Row, 8) = NulosN(Fg3.TextMatrix(Row, 7) + NulosN(Fg3.TextMatrix(Row, 6)))
    
    '--totalizar
    Fg3.TextMatrix(Fg3.Rows - 1, 6) = Format(GRID_SUMAR_COL(Fg3, 6) - NulosN(Fg3.TextMatrix(Fg3.Rows - 1, 6)), FORMAT_MONTO)
    Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(GRID_SUMAR_COL(Fg3, 7) - NulosN(Fg3.TextMatrix(Fg3.Rows - 1, 7)), FORMAT_MONTO)
    Fg3.TextMatrix(Fg3.Rows - 1, 8) = Format(GRID_SUMAR_COL(Fg3, 8) - NulosN(Fg3.TextMatrix(Fg3.Rows - 1, 8)), FORMAT_MONTO)
    
End Sub


Private Sub pGrabar(Tipo As Integer)
    '--tipo 0=horas, 1=destajo
    '--validando datos
    If Tipo = 0 Then '--Horas
        If Fg1.Rows = 1 Then
            MsgBox "No hay Registros Horas para grabar", vbInformation
            Exit Sub
        End If
    ElseIf Tipo = 1 Then '--Destajo
        If Fg3.Rows = 1 Then
            MsgBox "No hay Registros de Destajo para grabar", vbInformation
            Exit Sub
        End If
    Else '--Linea
        If fg5.Rows = 1 Then
            MsgBox "No hay Registros de Linea para grabar", vbInformation
            Exit Sub
        End If
    End If
    
    If MsgBox("Seguro desea continuar", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    
    habilitar TxtFecha, False
    
    Dim xFil&
    Dim xCod&
    Dim RstDet As New ADODB.Recordset
    Dim nSQLArea As String '--almacenará el filtro si elige area
    On Error GoTo error
    xCon.BeginTrans
    RST_Busq RstDet, "SELECT top 1 * FROM pro_pagos ", xCon
    
    '--registro por horas
    If Tipo = 0 Then
        
        '--eliminar los datos de pagos previamente grabados en forma grupal.
        If NulosN(txt_cb(0).Text) <> 0 Then
            
            If OptSeleccion(0).Value = True Then nSQLArea = " and pro_pagos.idarea = " & NulosN(txt_cb(0).Text) & " "
        
            xCon.Execute "DELETE pro_pagos.* FROM pro_pagos WHERE (((pro_pagos.tipo)=1) AND ((CDate([pro_pagos].[fchtra])) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "'))) " & nSQLArea
        Else
            
            xCon.Execute "DELETE pro_pagos.* FROM pro_pagos WHERE (((pro_pagos.tipo)=1) AND ((CDate([pro_pagos].[fchtra])) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "'))) "
            
        End If
        
        With Fg1
            
            For xFil = 1 To Fg1.Rows - 2
            
                DoEvents
                
                '--pro_pagos.tipo:: 1 =hora
                '--eliminar registro de pagos
                xCon.Execute "DELETE * FROM pro_pagos where pro_pagos.tipo=1 and cdate(pro_pagos.fchtra) = '" & CDate(.TextMatrix(xFil, 2)) & "' and pro_pagos.idemp = " & NulosN(.TextMatrix(xFil, 19))
                
                '--agragar nuevo registro
                RstDet.AddNew
                RstDet("tipo") = 1
                RstDet("fchtra") = CDate(.TextMatrix(xFil, 2))
                RstDet("imptot") = NulosN(.TextMatrix(xFil, 15))
                RstDet("impbon") = NulosN(.TextMatrix(xFil, 16))
                RstDet("impbrut") = NulosN(.TextMatrix(xFil, 17))
                RstDet("idemp") = NulosN(.TextMatrix(xFil, 19))
                RstDet("idarea") = NulosN(.TextMatrix(xFil, 18))
                
                If UCase(.TextMatrix(xFil, 7)) = "NOCHE" Then
                    RstDet("turno") = 2
                Else
                    RstDet("turno") = 1
                End If
                
                RstDet.Update
            Next
        End With
        
    '--registro por destajo
    ElseIf Tipo = 1 Then
        
        '--eliminar los datos de pagos previamente grabados en forma grupal.
        If NulosN(txt_cb(0).Text) <> 0 Then
            
            If OptSeleccion(0).Value = True Then nSQLArea = " and pro_pagos.idarea = " & NulosN(txt_cb(0).Text) & " "
        
            xCon.Execute "DELETE pro_pagos.* FROM pro_pagos WHERE (((pro_pagos.tipo)=2) AND ((CDate([pro_pagos].[fchtra])) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "'))) " & nSQLArea
        Else
            xCon.Execute "DELETE pro_pagos.* FROM pro_pagos WHERE (((pro_pagos.tipo)=2) AND ((CDate([pro_pagos].[fchtra])) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "'))) "
        
        End If
    
        '--para grabar se utiliza el resumen del pago
        With Fg3
            For xFil = 1 To .Rows - 2
                DoEvents
                '--pro_pagos.tipo:: 2 =destajo
                '--eliminar registro de pagos
                xCon.Execute "DELETE * FROM pro_pagos where pro_pagos.tipo=2 and cdate(pro_pagos.fchtra) = '" & CDate(.TextMatrix(xFil, 2)) & "' and pro_pagos.idemp = " & NulosN(.TextMatrix(xFil, 9)) & " and pro_pagos.idarea = " & NulosN(.TextMatrix(xFil, 10))
                '--agregar nuevo registro
                RstDet.AddNew
                RstDet("tipo") = 2
                
                RstDet("fchtra") = CDate(.TextMatrix(xFil, 2))
                RstDet("imptot") = NulosN(.TextMatrix(xFil, 6))
                RstDet("impbon") = NulosN(.TextMatrix(xFil, 7))
                RstDet("impbrut") = NulosN(.TextMatrix(xFil, 8))
                RstDet("idemp") = NulosN(.TextMatrix(xFil, 9))
                RstDet("idarea") = NulosN(.TextMatrix(xFil, 10))
                
                If UCase(.TextMatrix(xFil, 5)) = "NOCHE" Then
                    RstDet("turno") = 2
                Else
                    RstDet("turno") = 1
                End If
                
                RstDet.Update
            Next
        End With
    
    '--registro por linea
    ElseIf Tipo = 2 Then
    
        '--eliminar los datos de pagos previamente grabados en forma grupal.
        If NulosN(txt_cb(0).Text) <> 0 Then
            
            If OptSeleccion(0).Value = True Then nSQLArea = " and pro_pagos.idarea = " & NulosN(txt_cb(0).Text) & " "
            
            xCon.Execute "DELETE pro_pagos.* FROM pro_pagos WHERE (((pro_pagos.tipo)=3) AND ((CDate([pro_pagos].[fchtra])) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "'))) " & nSQLArea
            
        Else
            xCon.Execute "DELETE pro_pagos.* FROM pro_pagos WHERE (((pro_pagos.tipo)=3) AND ((CDate([pro_pagos].[fchtra])) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "'))) "
            
        End If
        
        

        '--para grabar se utiliza el resumen del pago
        With fg5
            For xFil = 1 To .Rows - 2
                DoEvents
                
                '--pro_pagos.tipo:: 3 =linea
                '--eliminar registro de pagos
                xCon.Execute "DELETE * FROM pro_pagos where pro_pagos.tipo=3 and cdate(pro_pagos.fchtra) = '" & CDate(.TextMatrix(xFil, 2)) & "' and pro_pagos.idemp = " & NulosN(.TextMatrix(xFil, 9)) & " and pro_pagos.idarea = " & NulosN(.TextMatrix(xFil, 10))
                '--agregar nuevo registro
                RstDet.AddNew
                RstDet("tipo") = 3
                RstDet("fchtra") = CDate(.TextMatrix(xFil, 2))
                RstDet("imptot") = NulosN(.TextMatrix(xFil, 6))
                RstDet("impbon") = NulosN(.TextMatrix(xFil, 7))
                RstDet("impbrut") = NulosN(.TextMatrix(xFil, 8))
                
                RstDet("idemp") = NulosN(.TextMatrix(xFil, 9))
                RstDet("idarea") = NulosN(.TextMatrix(xFil, 10))
                
                If UCase(.TextMatrix(xFil, 5)) = "NOCHE" Then
                    RstDet("turno") = 2
                Else
                    RstDet("turno") = 1
                End If
                
                RstDet.Update
            Next
        End With
    
    
    End If

    xCon.CommitTrans
    Set RstDet = Nothing
    
    habilitar TxtFecha, True
    
    MsgBox "Los registros por " & IIf(Tipo = 0, "Horas", IIf(Tipo = 1, "Destajo", "Linea")) & " se grabó con éxito", vbInformation, xTitulo
    
    Exit Sub
error:
    xCon.RollbackTrans
    habilitar TxtFecha, True
    MsgBox Err.Description & vbCr & Err.Source & vbCr & Err.Number, vbCritical, xTitulo
    Err.Clear
    Set RstDet = Nothing
End Sub

Private Sub pCargarTareasSinCosto(xRs As Recordset)
    If xRs.State = 0 Then Exit Sub
    
    xRs.Filter = adFilterNone
    If xRs.RecordCount = 0 Then Exit Sub
    
    xRs.MoveFirst
    fg(1).Rows = fg(1).FixedRows
    While Not xRs.EOF
        fg(1).Rows = fg(1).Rows + 1
        fg(1).TextMatrix(fg(1).Rows - 1, 1) = xRs("despro")
        fg(1).TextMatrix(fg(1).Rows - 1, 2) = xRs("destar")
        xRs.MoveNext
    Wend
End Sub

Sub preparaRST(ByRef xRs As Recordset)
    Dim xFun As New eps_librerias.FuncionesData
    
    Dim xCampos(2, 3) As String

    xCampos(0, 0) = "despro":      xCampos(0, 1) = "C":      xCampos(0, 2) = "100"
    xCampos(1, 0) = "destar":      xCampos(1, 1) = "C":      xCampos(1, 2) = "100"
    Set xRs = xFun.CrearRstTMP(xCampos)
    xRs.Open
End Sub

Private Sub pCargarDestajo()

    Dim Rst As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLFiltro As String
    
    Dim RstAux As New ADODB.Recordset ' Recordset que almacena tareas sin Costo
    
    Dim mRow&
    
'    On Error GoTo error
    
    Fg2.Rows = Fg2.FixedRows
    Fg3.Rows = Fg3.FixedRows
    DoEvents

    If NulosN(lbl_cod(0).Caption) <> 0 Then
        If OptSeleccion(0).Value = True Then
            nSQLFiltro = " and pro_controltar.idarea=" & NulosN(lbl_cod(0).Caption) & " "
        Else
            nSQLFiltro = " and pla_empleados.id=" & NulosN(lbl_cod(0).Caption) & " "
        End If

    End If
    
    
    '--actualizar el costo en la base de datos
    pActualizarCostoDestajo1
    
    DoEvents
    lbl(0).Caption = "Procesando: Registros"
    lbl(1).Caption = "Interrumpir = ESC"

    '--generar la consulta para presentar el informe
    nSQL = "SELECT vwtarea.*, vwcosto.canteo, iif(vwtarea.unidxhor = 0 or vwcosto.canteo = 0 or vwcosto.canteo is null ,0, (vwtarea.unidxhor/vwcosto.canteo)*100 ) as Eficiencia  " _
        + vbCr + " FROM ( "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.idctr, pro_controltardet.corr, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.cant AS CantReal, mae_unidades.abrev, pro_controltardet.observacion, " _
        + vbCr + " pro_controltardet.horini, pro_controltardet.horfin, IIf([pro_controltardet].[horini] Is Null Or [pro_controltardet].[horfin] Is Null,'',IIf([pro_controltardet].[horini]<CDate('13:20:00') And [pro_controltardet].[horfin]>CDate('14:00:00'),Format(CDate(Format([pro_controltardet].[horfin]-[pro_controltardet].[horini],'hh:mm:ss'))-CDate('01:00:00'),'hh:mm:ss'),Format([pro_controltardet].[horfin]-[pro_controltardet].[horini],'hh:mm:ss'))) AS difhora, IIf([difhora] Is Null Or [difhora]='',0,Hour([difhora])*60+Minute([difhora])) AS totmin, IIf([CantReal]=0 Or [totmin]=0,0,[CantReal]/[totmin]) AS UnidXMin, [UnidXMin]*60 AS UnidXHor, pro_controltardet.canpro, mae_unidades_1.abrev AS abrev1, pro_controltardet.preuni, pro_controltardet.imptot " _
        + vbCr + " FROM ((pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr) LEFT JOIN mae_unidades AS mae_unidades_1 ON pro_controltardet.idunid = mae_unidades_1.id " _
        + vbCr + " WHERE ((pro_controltar.tipo =2 AND pro_controltardet.tipo =1 ) AND ((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) ) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " UNION "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.idctr, pro_controltardet.corr, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardetgr.cant AS CantReal, mae_unidades.abrev, pro_controltardet.observacion,  " _
        + vbCr + "  pro_controltardetgr.horini, pro_controltardetgr.horfin, IIf([pro_controltardetgr].[horini] Is Null Or [pro_controltardetgr].[horfin] Is Null,'',IIf([pro_controltardetgr].[horini]<CDate('13:20:00') And [pro_controltardetgr].[horfin]>CDate('14:00:00'),Format(CDate(Format([pro_controltardetgr].[horfin]-[pro_controltardetgr].[horini],'hh:mm:ss'))-CDate('01:00:00'),'hh:mm:ss'),Format([pro_controltardetgr].[horfin]-[pro_controltardetgr].[horini],'hh:mm:ss'))) AS difhora, IIf([difhora] Is Null Or [difhora]='',0,Hour([difhora])*60+Minute([difhora])) AS totmin, IIf([CantReal]=0 Or [totmin]=0,0,[CantReal]/[totmin]) AS UnidXMin, [UnidXMin]*60 AS UnidXHor, pro_controltardetgr.canpro, mae_unidades_1.abrev AS abrev1, pro_controltardetgr.preuni, pro_controltardetgr.imptot " _
        + vbCr + " FROM ((pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN ((alm_inventario RIGHT JOIN (((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pro_controltar.id = pro_controltardet.idctr) LEFT JOIN mae_unidades AS mae_unidades_1 ON pro_controltardetgr.idunid = mae_unidades_1.id " _
        + vbCr + " WHERE pro_controltar.tipo =2 AND (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.idtar)<>0) AND ((pro_controltardet.tipo)=2) AND ((pro_controltardetgr.activo)=-1)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden " _
        + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " ) AS vwcosto"
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk "
            
    nSQL = nSQL _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.area,vwtarea.personal,vwtarea.horini; "
    
    
    
    RST_Busq Rst, nSQL, xCon
    
    PgBar.Min = 0
    PgBar.Value = 0
    If Rst.RecordCount <> 0 Then PgBar.Max = Rst.RecordCount
    With Fg2
        
        Do While Not Rst.EOF
            DoEvents
            If BAND_INTERRUMPIR = True Then GoTo SALIR
            PgBar.Value = PgBar.Value + 1
            '------------------------------------------------
            .Rows = .Rows + 1
            .TextMatrix(Fg2.Rows - 1, 1) = NulosN(Rst("idemp"))
            .TextMatrix(Fg2.Rows - 1, 2) = Format(Rst("fchtra"), FORMAT_DATE)
            .TextMatrix(Fg2.Rows - 1, 3) = NulosC(Rst("area"))
            .TextMatrix(Fg2.Rows - 1, 4) = NulosC(Rst("personal"))

            .TextMatrix(Fg2.Rows - 1, 5) = NulosC(Rst("tarea")) 'Tarea
            '--------
            .TextMatrix(Fg2.Rows - 1, 6) = NulosC(Rst("Producto")) 'Producto
            
            .TextMatrix(Fg2.Rows - 1, 7) = NulosC(Rst("observacion")) 'observacion
            
            
            .TextMatrix(Fg2.Rows - 1, 8) = Format(Rst("horini"), FORMAT_HORA_SIN_SEGUNDO)   'hini
            .TextMatrix(Fg2.Rows - 1, 9) = Format(Rst("horfin"), FORMAT_HORA_SIN_SEGUNDO)   'hfin
            
            
            .TextMatrix(Fg2.Rows - 1, 10) = Format(NulosN(Rst("cantreal")), FORMAT_MONTO) 'Cantidad
            .TextMatrix(Fg2.Rows - 1, 11) = Rst("abrev") 'U.M.
            
            
            .TextMatrix(Fg2.Rows - 1, 12) = Format(Rst("difhora"), FORMAT_HORA_LARGO)    'difhora
            .TextMatrix(Fg2.Rows - 1, 13) = NulosN(Rst("totmin")) 'Tot.Min"
            .TextMatrix(Fg2.Rows - 1, 14) = NulosN(Rst("unidxmin")) 'Unid x Min
            .TextMatrix(Fg2.Rows - 1, 15) = Format(NulosN(Rst("unidxhor")), "#,##0.00000") 'Unid x Hor
            .TextMatrix(Fg2.Rows - 1, 16) = Format(Rst("canteo"), FORMAT_MONTO) 'Cant Teo
            
            If NulosN(Rst.Fields("eficiencia")) = 100 Then  '--negro
                .TextMatrix(Fg2.Rows - 1, 17) = Format(Rst.Fields("eficiencia"), FORMAT_PORCENTAJE) & "%"
            ElseIf NulosN(Rst.Fields("eficiencia")) = 0 Then  '--no mostrar datos
                
            ElseIf NulosN(Rst.Fields("eficiencia")) > 100 Then '--azul (supera la eficiencia)
                FORMATO_CELDA Fg2, .Rows - 1, 17, &HFF0000, False, &HFFFFFF, Format(NulosN(Rst.Fields("eficiencia")), FORMAT_PORCENTAJE) + "%"
            ElseIf NulosN(Rst.Fields("eficiencia")) < 100 Then '--rojo (menos eficiente)
                FORMATO_CELDA Fg2, .Rows - 1, 17, &HFF, False, &HFFFFFF, Format(NulosN(Rst.Fields("eficiencia")), FORMAT_PORCENTAJE) + "%"
            End If
                
            
            .TextMatrix(Fg2.Rows - 1, 19) = Format(NulosN(Rst("canpro")), FORMAT_MONTO)
            .TextMatrix(Fg2.Rows - 1, 20) = NulosC(Rst("abrev1"))
            .TextMatrix(Fg2.Rows - 1, 21) = Format(NulosN(Rst("preuni")), "0.000000") 'Pre.Uni
            .TextMatrix(Fg2.Rows - 1, 22) = Format(NulosN(Rst("imptot")), FORMAT_MONTO) 'Total
                

            Rst.MoveNext
        Loop
    End With
    Set Rst = Nothing
    '----------
    GRID_AGRUPAR Fg2, 4
    
    '--pintar los montos =0
    Dim xFlat As Boolean

    For mRow = Fg2.FixedRows To Fg2.Rows - 1
        If NulosN(Fg2.TextMatrix(mRow, 22)) = 0 Then
        
            '*************************************************************************************************
            ' Se verifica que la el registro no este guardado en Tareas sin costo
            If RstAux.State = 0 Then preparaRST RstAux
            RstAux.Filter = "despro = '" & Fg2.TextMatrix(mRow, 6) & "' And destar = '" & Fg2.TextMatrix(mRow, 5) & "'"
            If RstAux.RecordCount = 0 Then
                ' Se guarda en registro
                RstAux.AddNew
                RstAux("despro") = NulosC(Fg2.TextMatrix(mRow, 6))
                RstAux("destar") = NulosC(Fg2.TextMatrix(mRow, 5))
                RstAux.Update
            End If
            '*************************************************************************************************
            
            GRID_COLOR_FONDO Fg2, mRow, 1, mRow, 22, vbRed
            xFlat = True
        End If
    Next
    If xFlat = True Then MsgBox "Se presentaron registros observados", vbInformation, xTitulo
    
    ' ************************************
    ' Se muestran las Tareas sin costo
    If xFlat = True Then
        CentrarFrm Frm4
        pCargarTareasSinCosto RstAux
        Frm4.Visible = True
    End If
    ' ************************************
    '--colocando los totales
    Fg2.Rows = Fg2.Rows + 1
    Fg2.TextMatrix(Fg2.Rows - 1, 4) = "Totales"
    Fg2.TextMatrix(Fg2.Rows - 1, 22) = Format(GRID_SUMAR_COL(Fg2, 22), FORMAT_MONTO)
    
    GRID_COLOR_FONDO Fg2, Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1, vbGreen
    
    '----------------------------------------------------------------------------------------
    '-- colocando los datos en el resumen
    
    nSQL = "SELECT vwtarea.idemp, vwtarea.idarea, vwtarea.fchtra, vwtarea.area, vwtarea.personal, Sum(vwtarea.total) AS toting, IIf([vwbono].[impbon] Is Null,0,[vwbono].[impbon]) AS totbono, [toting]+[totbono] AS totneto, First(vwtarea.hinipri) AS hinipri1, First(vwtarea.hfinpri) AS hfinpri1, Last(vwtarea.hiniult) AS hiniult1, Last(vwtarea.hfinult) AS hfinult1 " _
        + vbCr + " FROM ( "
    nSQL = nSQL _
        + vbCr + " SELECT Format([pro_controltar].[fchtra],'dd/mm/yy') & '-' & [pla_empleados].[id] AS codigopk, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_controltardet.imptot AS total, pro_controltardet.horini AS hinipri, pro_controltardet.horfin AS hfinpri, pro_controltardet.horini AS hiniult, pro_controltardet.horfin AS hfinult " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (pro_controltardet LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltar.tipo)=2) AND ((pro_controltardet.tipo)=1)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " UNION "
    nSQL = nSQL _
        + vbCr + " SELECT Format([pro_controltar].[fchtra],'dd/mm/yy') & '-' & [pla_empleados].[id] AS codigopk, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_controltardetgr.imptot AS total, pro_controltardetgr.horini AS hinipri, pro_controltardetgr.horfin AS hfinpri, pro_controltardetgr.horini AS hiniult, pro_controltardetgr.horfin AS hfinult " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (pro_controltardet INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltar.tipo)=2) AND ((pro_controltardet.tipo)=2) AND ((pro_controltardetgr.activo)=-1)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT Format([pro_pagos].[fchtra],'dd/mm/yy') & '-' & [pro_pagos].[idemp] AS codigopk, pro_pagos.idemp, pro_pagos.fchtra, pro_pagos.impbon " _
        + vbCr + " FROM pro_pagos WHERE (((pro_pagos.tipo) = 2) And ((pro_pagos.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "'))) " _
        + vbCr + " GROUP BY Format([pro_pagos].[fchtra],'dd/mm/yy') & '-' & [pro_pagos].[idemp], pro_pagos.idemp, pro_pagos.fchtra, pro_pagos.impbon  " _
        + vbCr + " ) AS vwbono"
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwbono.codigopk "
            
    nSQL = nSQL _
     + vbCr + " GROUP BY vwtarea.idemp, vwtarea.idarea, vwtarea.fchtra, vwtarea.area, vwtarea.personal, IIf([vwbono].[impbon] Is Null,0,[vwbono].[impbon]) " _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.area, vwtarea.personal, First(vwtarea.hinipri); "
    
    
    
    RST_Busq Rst, nSQL, xCon
    
    If Rst.RecordCount = 0 Then Exit Sub
    Agregando = True
    PgBar.Min = 0
    PgBar.Value = 0
    PgBar.Max = Rst.RecordCount
    
    
    With Fg3
        Do While Not Rst.EOF
            DoEvents
            If BAND_INTERRUMPIR = True Then GoTo SALIR
            PgBar.Value = PgBar.Value + 1
            '------------------------------------------------
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = Rst.Bookmark
            .TextMatrix(.Rows - 1, 2) = Format(Rst("fchtra"), FORMAT_DATE)
            .TextMatrix(.Rows - 1, 3) = NulosC(Rst("area"))
            .TextMatrix(.Rows - 1, 4) = NulosC(Rst("personal"))

            '--determinar el horario
            
            If IsDate(Rst("hfinult1")) = True Then
                '--evaluando los las horas
                If CDate(Rst("hfinult1")) < CDate("10:00") Then
                    .TextMatrix(.Rows - 1, 5) = "Noche"
                Else
                    .TextMatrix(.Rows - 1, 5) = "Dia"
                End If
            End If
            
            '--------
            .TextMatrix(.Rows - 1, 6) = Format(NulosN(Rst("toting")), FORMAT_MONTO) 'totingreso
            .TextMatrix(.Rows - 1, 7) = Format(NulosN(Rst("totbono")), FORMAT_MONTO) 'totbono
            .TextMatrix(.Rows - 1, 8) = Format(NulosN(Rst("totneto")), FORMAT_MONTO) 'totneto
            
'            '--si es de noche el costo de la tarea incrementar en un porentaje ejm 30%
'            If .TextMatrix(.Rows - 1, 5) = "Noche" Then
'                .TextMatrix(.Rows - 1, 6) = Format(NulosN(.TextMatrix(.Rows - 1, 6)) * 1.3, FORMAT_MONTO)
'                .TextMatrix(.Rows - 1, 8) = Format(NulosN(.TextMatrix(.Rows - 1, 6)) + NulosN(.TextMatrix(.Rows - 1, 7)), FORMAT_MONTO)
'            End If
'
            
            .TextMatrix(.Rows - 1, 9) = NulosN(Rst("idemp"))
            .TextMatrix(.Rows - 1, 10) = NulosN(Rst("idarea"))
            
            Rst.MoveNext
        Loop
    End With
    Set Rst = Nothing
    '----------
    GRID_AGRUPAR Fg3, 3
    
    '--colocando los totales
    Fg3.Rows = Fg3.Rows + 1
    Fg3.TextMatrix(Fg3.Rows - 1, 4) = "Totales"
    Fg3.TextMatrix(Fg3.Rows - 1, 6) = Format(GRID_SUMAR_COL(Fg3, 6), FORMAT_MONTO)
    Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(GRID_SUMAR_COL(Fg3, 7), FORMAT_MONTO)
    Fg3.TextMatrix(Fg3.Rows - 1, 8) = Format(GRID_SUMAR_COL(Fg3, 8), FORMAT_MONTO)
    
    GRID_COLOR_FONDO Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, Fg3.Cols - 1, vbGreen
    
    
    '----------------------------------------------------------------------------------------
    
    
SALIR:
Agregando = False
Exit Sub
error:
    MsgBox Err.Description & vbCr & Err.Source & vbCr & Err.Number, vbCritical, xTitulo
    Err.Clear
End Sub

Private Sub pActualizarCostoDestajo()

    Dim Rst As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLFiltro  As String
    
    '--actualizando el destajo individual
    lbl(0).Caption = "Actualizando Costos 1/2"
    lbl(1).Caption = "No Interrumpir"
    DoEvents
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        If OptSeleccion(0).Value = True Then
            nSQLFiltro = " and pro_controltar.idarea=" & NulosN(lbl_cod(0).Caption) & " "
        Else
            nSQLFiltro = " and pla_empleados.id=" & NulosN(lbl_cod(0).Caption) & " "
        End If
        
    End If
    
    
    nSQL = "SELECT vwtarea.idctr,vwtarea.corr, vwtarea.idemp, vwtarea.idarea, vwtarea.idtar, vwtarea.idrec, vwtarea.idunimed, vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto, vwtarea.tarea, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.CantReal,vwtarea.tothor) AS canpro, vwtarea.abrev,  iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwcosto.costo,0)  AS preuni, [preuni]*[canpro] AS tot ,  vwcosto.paghor,vwtarea.tothor, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.idunimed,7)  as idunid " _
        + vbCr + " FROM ( "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.idctr, pro_controltardet.corr,pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.cant AS CantReal, mae_unidades.abrev,pro_controltardet.tothor " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE ((pro_controltar.tipo =2 AND pro_controltardet.tipo =1 ) AND ((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.cant)<>0)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden,pro_costodet.paghor " _
        + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " ) AS vwcosto"
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk "
            
    nSQL = nSQL _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto; "
    
    
    RST_Busq RstTmp, nSQL, xCon
    
    PgBar.Min = 0
    PgBar.Value = 0
    
    Dim xHoras As Double
    If RstTmp.RecordCount <> 0 Then
        PgBar.Max = RstTmp.RecordCount
        RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            PgBar.Value = RstTmp.Bookmark
            xCon.Execute "update pro_controltardet set canpro=" & NulosN(RstTmp("canpro")) & " ,idunid=" & NulosN(RstTmp("idunid")) & ", preuni = " & NulosN(RstTmp("preuni")) & ", imptot =" & NulosN(RstTmp("tot")) & " where idctr = " & NulosN(RstTmp("idctr")) & " and corr = " & NulosN(RstTmp("corr")) & " and tipo=1 and idref = " & NulosN(RstTmp("idemp")) & " and idtar = " & NulosN(RstTmp("idtar")) & " and idrec = " & NulosN(RstTmp("idrec")) & " and idunimed =" & NulosN(RstTmp("idunimed"))
            RstTmp.MoveNext
        Loop
    End If
    Set RstTmp = Nothing
    
    
    '-----------
    '--actualizando el destajo grupal
    lbl(0).Caption = "Actualizando Costos 2/2"
    DoEvents
    
    nSQL = "SELECT vwtarea.idctr,vwtarea.corr, vwtarea.idemp, vwtarea.idarea, vwtarea.idtar, vwtarea.idrec, vwtarea.idunimed, vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto, vwtarea.tarea, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.CantReal,vwtarea.tothor) AS canpro, vwtarea.abrev,  iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwcosto.costo,0)  AS preuni, [preuni]*[canpro] AS tot ,  vwcosto.paghor,vwtarea.tothor, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.idunimed,7)  as idunid " _
        + vbCr + " FROM ( "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk,pro_controltardet.idctr, pro_controltardet.corr, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardetgr.cant AS CantReal, mae_unidades.abrev,pro_controltardetgr.tothor " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE pro_controltar.tipo =2 AND (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.idtar)<>0) AND ((pro_controltardet.tipo)=2) AND ((pro_controltardetgr.cant)<>0) AND ((pro_controltardetgr.activo)=-1)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden,pro_costodet.paghor " _
        + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " ) AS vwcosto"
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk "
            
    nSQL = nSQL _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto; "
    
    
    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.RecordCount <> 0 Then
        PgBar.Min = 0
        PgBar.Value = 0
        PgBar.Max = RstTmp.RecordCount
        
        RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            
            PgBar.Value = RstTmp.Bookmark
            
            xCon.Execute "UPDATE pro_controltardet INNER JOIN pro_controltardetgr ON (pro_controltardet.idctr = pro_controltardetgr.idctr) AND (pro_controltardet.corr = pro_controltardetgr.corr) SET pro_controltardetgr.canpro=" & NulosN(RstTmp("canpro")) & ",pro_controltardetgr.idunid=" & NulosN(RstTmp("idunid")) & ", pro_controltardetgr.preuni = " & NulosN(RstTmp("preuni")) & ", pro_controltardetgr.imptot = " & NulosN(RstTmp("tot")) _
                & " WHERE (((pro_controltardet.idctr)=" & NulosN(RstTmp("idctr")) & ") AND ((pro_controltardet.corr)=" & NulosN(RstTmp("corr")) & ") AND " _
                & " ((pro_controltardet.idtar)=" & NulosN(RstTmp("idtar")) & ") AND ((pro_controltardet.idrec)=" & NulosN(RstTmp("idrec")) & " ) AND " _
                & " ((pro_controltardetgr.idper)=" & NulosN(RstTmp("idemp")) & ") AND ((pro_controltardet.idunimed)=" & NulosN(RstTmp("idunimed")) & " ) AND " _
                & " ((pro_controltardet.tipo)=2)); "

            
            RstTmp.MoveNext
        Loop
    End If
    Set RstTmp = Nothing
        
SALIR:
Agregando = False

End Sub



Private Sub pActualizarCostoDestajo1()

    Dim Rst As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLFiltro  As String
    
    '--actualizando el destajo individual
    lbl(0).Caption = "Actualizando Costos 1/2"
    lbl(1).Caption = "No Interrumpir"
    DoEvents
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        If OptSeleccion(0).Value = True Then
            nSQLFiltro = " and pro_controltar.idarea=" & NulosN(lbl_cod(0).Caption) & " "
        Else
            nSQLFiltro = " and pla_empleados.id=" & NulosN(lbl_cod(0).Caption) & " "
        End If
        
    End If
    
    
    nSQL = "SELECT vwtarea.idctr,vwtarea.corr, vwtarea.idemp, vwtarea.idarea, vwtarea.idtar, vwtarea.idrec, vwtarea.idunimed, vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto, vwtarea.tarea, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.CantReal,vwtarea.tothor) AS canpro, vwtarea.abrev,  IIf([vwcosto].[paghor]=0,[vwcosto].[costo],[vwcostoh].[costo]) AS preuni, [preuni]*[canpro] AS tot ,  vwcosto.paghor,vwtarea.tothor, iif(vwcosto.paghor=0,vwtarea.idunimed,7)  as idunid " _
        + vbCr + " FROM (( "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk,IIf(pro_controltardet.idrec=0 Or pro_controltardet.idrec Is Null,'-',pro_controltardet.idrec) & '*' & pro_controltardet.idtar AS codigopk1, pro_controltardet.idctr, pro_controltardet.corr,pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.cant AS CantReal, mae_unidades.abrev,pro_controltardet.tothor " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE ((pro_controltar.tipo =2 AND pro_controltardet.tipo =1 ) AND ((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.cant)<>0)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden,pro_costodet.paghor " _
        + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " ) AS vwcosto "
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk ) "
            
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] AS codigopk1, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden, pro_costodet.paghor " _
        + vbCr + " FROM (alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (pro_tareas INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_tareas.id = pro_costodet.idtar) ON pro_costo.id = pro_costodet.idcos " _
        + vbCr + " WHERE (((pro_costodet.idunimed)=7)) " _
        + vbCr + " ) AS vwcostoh "
    nSQL = nSQL _
        + vbCr + "  ON vwtarea.codigopk1 = vwcostoh.codigopk1 "
            
            
    nSQL = nSQL _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto; "
    
    
    RST_Busq RstTmp, nSQL, xCon
    
    PgBar.Min = 0
    PgBar.Value = 0
    
    Dim xHoras As Double
    If RstTmp.RecordCount <> 0 Then
        PgBar.Max = RstTmp.RecordCount
        RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            PgBar.Value = RstTmp.Bookmark
            xCon.Execute "update pro_controltardet set canpro=" & NulosN(RstTmp("canpro")) & " ,idunid=" & NulosN(RstTmp("idunid")) & ", preuni = " & NulosN(RstTmp("preuni")) & ", imptot =" & NulosN(RstTmp("tot")) & " where idctr = " & NulosN(RstTmp("idctr")) & " and corr = " & NulosN(RstTmp("corr")) & " and tipo=1 and idref = " & NulosN(RstTmp("idemp")) & " and idtar = " & NulosN(RstTmp("idtar")) & " and idrec = " & NulosN(RstTmp("idrec")) & " and idunimed =" & NulosN(RstTmp("idunimed"))
            RstTmp.MoveNext
        Loop
    End If
    Set RstTmp = Nothing
    
    
    '-----------
    '--actualizando el destajo grupal
    lbl(0).Caption = "Actualizando Costos 2/2"
    DoEvents
    
    nSQL = "SELECT vwtarea.idctr,vwtarea.corr, vwtarea.idemp, vwtarea.idarea, vwtarea.idtar, vwtarea.idrec, vwtarea.idunimed, vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto, vwtarea.tarea, iif(vwcosto.paghor=0 or vwcosto.paghor is null,vwtarea.CantReal,vwtarea.tothor) AS canpro, vwtarea.abrev, IIf([vwcosto].[paghor]=0,[vwcosto].[costo],[vwcostoh].[costo]) AS preuni, [preuni]*[canpro] AS tot ,  vwcosto.paghor,vwtarea.tothor, iif(vwcosto.paghor=0,vwtarea.idunimed,7)  as idunid " _
        + vbCr + " FROM (( "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk,IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] AS codigopk1, pro_controltardet.idctr, pro_controltardet.corr, pla_empleados.id AS idemp, pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardetgr.cant AS CantReal, mae_unidades.abrev,pro_controltardetgr.tothor " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE pro_controltar.tipo =2 AND (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.idtar)<>0) AND ((pro_controltardet.tipo)=2) AND ((pro_controltardetgr.cant)<>0) AND ((pro_controltardetgr.activo)=-1)) " & nSQLFiltro
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden,pro_costodet.paghor " _
        + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " ) AS vwcosto"
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk ) "
            
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] AS codigopk1, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden, pro_costodet.paghor " _
        + vbCr + " FROM (alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (pro_tareas INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_tareas.id = pro_costodet.idtar) ON pro_costo.id = pro_costodet.idcos " _
        + vbCr + " WHERE (((pro_costodet.idunimed)=7)) " _
        + vbCr + " ) AS vwcostoh "
    nSQL = nSQL _
        + vbCr + "  ON vwtarea.codigopk1 = vwcostoh.codigopk1 "
            
    nSQL = nSQL _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.personal, vwtarea.area, vwtarea.producto; "
    
    
    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.RecordCount <> 0 Then
        PgBar.Min = 0
        PgBar.Value = 0
        PgBar.Max = RstTmp.RecordCount
        
        RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            
            PgBar.Value = RstTmp.Bookmark
            
            xCon.Execute "UPDATE pro_controltardet INNER JOIN pro_controltardetgr ON (pro_controltardet.idctr = pro_controltardetgr.idctr) AND (pro_controltardet.corr = pro_controltardetgr.corr) SET pro_controltardetgr.canpro=" & NulosN(RstTmp("canpro")) & ",pro_controltardetgr.idunid=" & NulosN(RstTmp("idunid")) & ", pro_controltardetgr.preuni = " & NulosN(RstTmp("preuni")) & ", pro_controltardetgr.imptot = " & NulosN(RstTmp("tot")) _
                & " WHERE (((pro_controltardet.idctr)=" & NulosN(RstTmp("idctr")) & ") AND ((pro_controltardet.corr)=" & NulosN(RstTmp("corr")) & ") AND " _
                & " ((pro_controltardet.idtar)=" & NulosN(RstTmp("idtar")) & ") AND ((pro_controltardet.idrec)=" & NulosN(RstTmp("idrec")) & " ) AND " _
                & " ((pro_controltardetgr.idper)=" & NulosN(RstTmp("idemp")) & ") AND ((pro_controltardet.idunimed)=" & NulosN(RstTmp("idunimed")) & " ) AND " _
                & " ((pro_controltardet.tipo)=2)); "

            
            RstTmp.MoveNext
        Loop
    End If
    Set RstTmp = Nothing
        
SALIR:
Agregando = False

End Sub

Private Function pBuscarTareasSinCosto(xId As Double, xCorr As Double, xIdRec As Double, xIdUnimed As Double) As Recordset
    Dim xRs As New ADODB.Recordset
    Dim RstTarea As New ADODB.Recordset
    Dim RstTareaLinea As New ADODB.Recordset
    Dim nSQL As String
    
    ' Se hallan las tareas con costo cero de la receta
    nSQL = "SELECT alm_inventario.descripcion AS producto, pro_tareas.descripcion AS tarea, pro_recetatar.orden, pro_recetatar.idtar, vwcosto.costo " _
        + vbCr + " FROM ((pro_receta INNER JOIN (pro_recetatar INNER JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id) ON pro_receta.id = pro_recetatar.idrec) INNER JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) " _
        + vbCr + " Left Join " _
        + vbCr + " ( SELECT pro_costo.idref, pro_costodet.idtar, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.costo " _
        + vbCr + " FROM pro_tareas INNER JOIN (pro_costo INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " Where (((pro_costo.idref) = " & xIdRec & ") And ((pro_costodet.idunimed) = " & xIdUnimed & ") And ((pro_costo.Tipo) = 1)) " _
        + vbCr + " ) AS vwcosto ON pro_recetatar.idtar = vwcosto.idtar " _
        + vbCr + " Where (((pro_receta.id) = " & xIdRec & ") AND ((vwcosto.costo)=0 Or (vwcosto.costo) Is Null));"
        
'WHERE (((vwcosto.costo)=0 Or (vwcosto.costo) Is Null) AND ((pro_receta.id)=195));


'    nSQL = "SELECT pro_receta.id AS idrec, pro_recetatar.idtar " _
'        + vbCr + "FROM (pro_receta INNER JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec) LEFT JOIN " _
'        + vbCr + "( " _
'        + vbCr + "SELECT pro_costo.idref, pro_costodet.idtar, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.costo " _
'        + vbCr + "FROM pro_tareas INNER JOIN (pro_costo INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
'        + vbCr + "Where (((pro_costo.idref) = " & xIdRec & ") And ((pro_costodet.idunimed) = " & xIdUnimed & ") And ((pro_costo.Tipo) = 1)) " _
'        + vbCr + ") " _
'        + vbCr + "AS vwcosto ON pro_recetatar.idtar = vwcosto.idtar " _
'        + vbCr + "Where (((pro_receta.id) = " & xIdRec & ") AND ((vwcosto.costo)=0));"

    '--ejecutar sentencia
    RST_Busq RstTarea, nSQL, xCon
    
    ' Se hallan las Tareas activas registradas en la linea
    nSQL = "SELECT pro_controltardettar.idtar " _
        + vbCr + "From pro_controltardettar " _
        + vbCr + "WHERE (((pro_controltardettar.idctr)=" & xId & ") AND ((pro_controltardettar.corr)=" & xCorr & ")) and pro_controltardettar.activo = -1 "

    '--ejecutar sentencia
    RST_Busq RstTareaLinea, nSQL, xCon
    
    ' Si no se ha cargado correctamente los recordset
    If RstTarea.State = 0 Then Set xRs = Nothing: GoTo SALIR
    If RstTareaLinea.State = 0 Then Set xRs = Nothing: GoTo SALIR
    ' Si no hay tareas con costo cero
    If RstTarea.RecordCount = 0 Then Set xRs = Nothing: GoTo SALIR
    
    RstTarea.MoveFirst
    While Not RstTarea.EOF
        ' Se busca la tarea si esta registrada o no
        RstTareaLinea.Filter = ""
        RstTareaLinea.Filter = "idtar=" & NulosN(RstTarea("idtar"))
        If RstTareaLinea.RecordCount = 0 Then GoTo SIGUIENTE
        
        ' Se agrega al recordset temporal
        If xRs.State = 0 Then preparaRST xRs
        xRs.AddNew
        xRs("despro") = RstTarea("producto")
        xRs("destar") = RstTarea("tarea")
SIGUIENTE:
        RstTarea.MoveNext
    Wend
    Set RstTarea = Nothing
    Set RstTareaLinea = Nothing
    
SALIR:
    If xRs.State <> 0 Then xRs.Filter = adFilterNone
    Set pBuscarTareasSinCosto = xRs
End Function

Private Sub pCargarLinea()
    '===================================================================================================
    'Creado : 20/05/11 Johan Castro
    'Propósito: Mostrar en pantalla el tipo Lineas de Produccion
    '
    'Entradas:  Ninguno
    '
    'Resultados: Datos en pantalla, detalle y resumen
    '
    '===================================================================================================

    Dim Rst As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLFiltro As String
    
    Dim RstAux As New ADODB.Recordset
    
    Dim mRow&
    
    On Error GoTo error
    
    Fg4.Rows = Fg4.FixedRows
    fg5.Rows = fg5.FixedRows
    DoEvents

    If NulosN(lbl_cod(0).Caption) <> 0 Then
        If OptSeleccion(0).Value = True Then
            nSQLFiltro = " and pro_controltar.idarea=" & NulosN(lbl_cod(0).Caption) & " "
        Else
            nSQLFiltro = " and pla_empleados.id=" & NulosN(lbl_cod(0).Caption) & " "
        End If

    End If
    
    '--actualizar el costo en la base de datos
    pActualizarCostoLinea
    
    DoEvents
    lbl(0).Caption = "Procesando: Registros"
    
    lbl(1).Caption = "Interrumpir = ESC"

    '--generar la consulta para presentar el informe
    
    nSQL = "SELECT pro_controltardet.idctr, pro_controltardet.corr, pro_controltardet.idrec, pro_controltar.idarea, pla_empleados.id AS idemp, pro_controltardet.idunimed, pro_controltar.fchtra, mae_area.descripcion AS area, pla_empleados.nombre AS personal, " _
        + vbCr + " alm_inventario.descripcion AS producto, pro_controltardetgr.cant AS CantReal, mae_unidades.abrev, pro_controltardet.observacion, pro_controltardetgr.horini, pro_controltardetgr.horfin, pro_controltardetgr.canpro, mae_unidades_1.abrev AS abrev1, pro_controltardetgr.preuni, pro_controltardetgr.imptot " _
        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (((alm_inventario RIGHT JOIN ((pro_controltardet LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) " _
        + vbCr + " INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) LEFT JOIN mae_unidades AS mae_unidades_1 ON pro_controltardetgr.idunid = mae_unidades_1.id) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.tipo)=3) AND ((pro_controltardetgr.activo)=-1)) " & nSQLFiltro & " ORDER BY pro_controltar.fchtra, pla_empleados.nombre, pro_controltardetgr.horini "

    
    
    RST_Busq Rst, nSQL, xCon
    
    PgBar.Min = 0
    PgBar.Value = 0
    If Rst.RecordCount <> 0 Then PgBar.Max = Rst.RecordCount
    With Fg4
        
        Do While Not Rst.EOF
            DoEvents
            If BAND_INTERRUMPIR = True Then GoTo SALIR
            PgBar.Value = PgBar.Value + 1
            '------------------------------------------------
            .Rows = .Rows + 1
            .TextMatrix(Fg4.Rows - 1, 1) = NulosN(Rst("idemp"))
            .TextMatrix(Fg4.Rows - 1, 2) = Format(Rst("fchtra"), FORMAT_DATE)
            .TextMatrix(Fg4.Rows - 1, 3) = NulosC(Rst("area"))
            .TextMatrix(Fg4.Rows - 1, 4) = NulosC(Rst("personal"))
            .TextMatrix(Fg4.Rows - 1, 5) = NulosC(Rst("Producto")) 'Producto
            .TextMatrix(Fg4.Rows - 1, 6) = NulosC(Rst("observacion")) 'observacion
            .TextMatrix(Fg4.Rows - 1, 7) = Format(Rst("horini"), FORMAT_HORA_SIN_SEGUNDO)   'hini
            .TextMatrix(Fg4.Rows - 1, 8) = Format(Rst("horfin"), FORMAT_HORA_SIN_SEGUNDO)   'hfin
            .TextMatrix(Fg4.Rows - 1, 9) = Format(NulosN(Rst("cantreal")), FORMAT_MONTO) 'Cantidad
            .TextMatrix(Fg4.Rows - 1, 10) = Rst("abrev") 'U.M.
                            
            .TextMatrix(Fg4.Rows - 1, 12) = Format(NulosN(Rst("canpro")), FORMAT_MONTO) '--cantidad para pago
            .TextMatrix(Fg4.Rows - 1, 13) = NulosC(Rst("abrev1"))
            .TextMatrix(Fg4.Rows - 1, 14) = Format(NulosN(Rst("preuni")), "0.00000000") 'Pre.Uni
            .TextMatrix(Fg4.Rows - 1, 15) = Format(NulosN(Rst("imptot")), FORMAT_MONTO) 'Total
            
            .TextMatrix(Fg4.Rows - 1, 16) = NulosN(Rst("idctr"))
            .TextMatrix(Fg4.Rows - 1, 17) = NulosN(Rst("corr"))
            .TextMatrix(Fg4.Rows - 1, 18) = NulosN(Rst("idrec"))
            .TextMatrix(Fg4.Rows - 1, 19) = NulosN(Rst("idunimed"))
            
                
            Rst.MoveNext
        Loop
    End With
    Set Rst = Nothing
    '----------
    GRID_AGRUPAR Fg4, 4
    
    '--pintar los montos =0
    Dim xFlat As Boolean
    
    ' **************************************************************
    Dim xId As Double
    Dim xCorr As Double
    Dim xIdRec As Double
    Dim xIdUnimed As Double
    Dim xRs As New ADODB.Recordset
    ' **************************************************************
    
    For mRow = Fg4.FixedRows To Fg4.Rows - 1
        If NulosN(Fg4.TextMatrix(mRow, 15)) = 0 Then
            '*************************************************************************************************
            ' Se encuentra las Tareas sin Costo relacionadas a la linea
            xId = NulosN(Fg4.TextMatrix(mRow, 16))
            xCorr = NulosN(Fg4.TextMatrix(mRow, 17))
            xIdRec = NulosN(Fg4.TextMatrix(mRow, 18))
            xIdUnimed = NulosN(Fg4.TextMatrix(mRow, 19))
            
            Set xRs = pBuscarTareasSinCosto(xId, xCorr, xIdRec, xIdUnimed)
            
            If xRs.State = 0 Then GoTo SIGUIENTE
            If xRs.RecordCount = 0 Then GoTo SIGUIENTE
            
            xRs.MoveFirst
            While Not xRs.EOF
                ' Se verifica que el registro no este guardado en Tareas sin costo
                If RstAux.State = 0 Then preparaRST RstAux
                ' Se filtra la tarea involucrada si es que el recordset tuviera tareas registradas
                RstAux.Filter = adFilterNone
                If RstAux.RecordCount > 0 Then
                    RstAux.Filter = "despro = '" & Fg4.TextMatrix(mRow, 5) & "' And destar = '" & xRs("destar") & "'"
                End If
                If RstAux.RecordCount = 0 Then
                    ' Se guarda en registro
                    RstAux.AddNew
                    RstAux("despro") = NulosC(Fg4.TextMatrix(mRow, 5))
                    RstAux("destar") = NulosC(xRs("destar"))
                    RstAux.Update
                End If
                xRs.MoveNext
            Wend
SIGUIENTE:
            '*************************************************************************************************
            
            GRID_COLOR_FONDO Fg4, mRow, 1, mRow, 15, vbRed
            xFlat = True
        End If
    Next
    If xFlat = True Then MsgBox "Se presentaron registros observados", vbInformation, xTitulo
    
    ' ************************************
    ' Se muestran las Tareas sin costo
    If xFlat = True Then
        CentrarFrm Frm4
        pCargarTareasSinCosto RstAux
        Frm4.Visible = True
    End If
    ' ************************************
    
    '--colocando los totales
    Fg4.Rows = Fg4.Rows + 1
    Fg4.TextMatrix(Fg4.Rows - 1, 4) = "Totales"
    Fg4.TextMatrix(Fg4.Rows - 1, 15) = Format(GRID_SUMAR_COL(Fg4, 15), FORMAT_MONTO)
    
    GRID_COLOR_FONDO Fg4, Fg4.Rows - 1, 1, Fg4.Rows - 1, Fg4.Cols - 1, vbGreen
    
    '----------------------------------------------------------------------------------------
    '-- colocando los datos en el resumen
    nSQL = "SELECT vwtarea.idemp, vwtarea.idarea, vwtarea.fchtra, vwtarea.area, vwtarea.personal, Sum(vwtarea.imptot) AS toting, IIf([vwbono].[impbon] Is Null,0,[vwbono].[impbon]) AS totbono, [toting]+[totbono] AS totneto, First(vwtarea.hinipri) AS hinipri1, First(vwtarea.hfinpri) AS hfinpri1, Last(vwtarea.hiniult) AS hiniult1, Last(vwtarea.hfinult) AS hfinult1 " _
     + vbCr + " FROM ( " _
     + vbCr + " SELECT Format([pro_controltar].[fchtra],'dd/mm/yy') & '-' & [pla_empleados].[id] AS codigopk, pro_controltardet.idctr, pro_controltardet.idrec, pro_controltar.idarea, pla_empleados.id AS idemp, pro_controltar.fchtra, mae_area.descripcion AS area, pla_empleados.nombre AS personal, alm_inventario.descripcion AS producto, pro_controltardetgr.cant AS CantReal, " _
     + vbCr + " pro_controltardetgr.horini, pro_controltardetgr.horfin, pro_controltardetgr.canpro, pro_controltardetgr.preuni, pro_controltardetgr.imptot,pro_controltardetgr.horini AS hinipri, pro_controltardetgr.horfin AS hfinpri, pro_controltardetgr.horini AS hiniult, pro_controltardetgr.horfin AS hfinult " _
     + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN ((alm_inventario RIGHT JOIN (pro_controltardet LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados " _
     + vbCr + " ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.idctr = pro_controltardetgr.idctr) AND (pro_controltardet.corr = pro_controltardetgr.corr)) ON pro_controltar.id = pro_controltardet.idctr " _
     + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.tipo)=3) AND ((pro_controltardetgr.activo)=-1)) " & nSQLFiltro _
     + vbCr + " ) AS vwtarea "
    '--consulta de los pagos adicionales que se le haya hecho en el dia
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT Format([pro_pagos].[fchtra],'dd/mm/yy') & '-' & [pro_pagos].[idemp] AS codigopk, pro_pagos.idemp, pro_pagos.fchtra, pro_pagos.impbon " _
        + vbCr + " FROM pro_pagos " _
        + vbCr + " WHERE (((pro_pagos.tipo) = 3) And ((pro_pagos.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')))" _
        + vbCr + " GROUP BY Format([pro_pagos].[fchtra],'dd/mm/yy') & '-' & [pro_pagos].[idemp], pro_pagos.idemp, pro_pagos.fchtra, pro_pagos.impbon  " _
        + vbCr + " ) AS vwbono"
     
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwbono.codigopk "
            
    nSQL = nSQL _
     + vbCr + " GROUP BY vwtarea.idemp, vwtarea.idarea, vwtarea.fchtra, vwtarea.area, vwtarea.personal, IIf([vwbono].[impbon] Is Null,0,[vwbono].[impbon]) " _
     + vbCr + " ORDER BY vwtarea.fchtra, vwtarea.area, vwtarea.personal, First(vwtarea.hinipri); "
    
    
    RST_Busq Rst, nSQL, xCon
    
    If Rst.RecordCount = 0 Then Exit Sub
    Agregando = True
    PgBar.Min = 0
    PgBar.Value = 0
    PgBar.Max = Rst.RecordCount
    
    
    With fg5
        Do While Not Rst.EOF
            DoEvents
            If BAND_INTERRUMPIR = True Then GoTo SALIR
            PgBar.Value = PgBar.Value + 1
            '------------------------------------------------
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = Rst.Bookmark
            .TextMatrix(.Rows - 1, 2) = Format(Rst("fchtra"), FORMAT_DATE)
            .TextMatrix(.Rows - 1, 3) = NulosC(Rst("area"))
            .TextMatrix(.Rows - 1, 4) = NulosC(Rst("personal"))

            '--determinar el horario
            
            If IsDate(Rst("hfinult1")) = True Then
                '--evaluando los las horas
                If CDate(Rst("hfinult1")) < CDate("10:00") Then
                    .TextMatrix(.Rows - 1, 5) = "Noche"
                Else
                    .TextMatrix(.Rows - 1, 5) = "Dia"
                End If
            End If
            
            '--------
            .TextMatrix(.Rows - 1, 6) = Format(NulosN(Rst("toting")), FORMAT_MONTO) 'totingreso
            .TextMatrix(.Rows - 1, 7) = Format(NulosN(Rst("totbono")), FORMAT_MONTO) 'totbono
            .TextMatrix(.Rows - 1, 8) = Format(NulosN(Rst("totneto")), FORMAT_MONTO) 'totneto
            
'            '--si es de noche el costo de la tarea incrementar en un porentaje ejm 30%
'            If .TextMatrix(.Rows - 1, 5) = "Noche" Then
'                .TextMatrix(.Rows - 1, 6) = Format(NulosN(.TextMatrix(.Rows - 1, 6)) * 1.3, FORMAT_MONTO)
'                .TextMatrix(.Rows - 1, 8) = Format(NulosN(.TextMatrix(.Rows - 1, 6)) + NulosN(.TextMatrix(.Rows - 1, 7)), FORMAT_MONTO)
'            End If
'
            
            .TextMatrix(.Rows - 1, 9) = NulosN(Rst("idemp"))
            .TextMatrix(.Rows - 1, 10) = NulosN(Rst("idarea"))
            
            Rst.MoveNext
        Loop
    End With
    Set Rst = Nothing
    '----------
    GRID_AGRUPAR fg5, 3
    
    '--colocando los totales
    fg5.Rows = fg5.Rows + 1
    fg5.TextMatrix(fg5.Rows - 1, 4) = "Totales"
    fg5.TextMatrix(fg5.Rows - 1, 6) = Format(GRID_SUMAR_COL(fg5, 6), FORMAT_MONTO)
    fg5.TextMatrix(fg5.Rows - 1, 7) = Format(GRID_SUMAR_COL(fg5, 7), FORMAT_MONTO)
    fg5.TextMatrix(fg5.Rows - 1, 8) = Format(GRID_SUMAR_COL(fg5, 8), FORMAT_MONTO)
    
    GRID_COLOR_FONDO fg5, fg5.Rows - 1, 1, fg5.Rows - 1, fg5.Cols - 1, vbGreen
    
    
    '----------------------------------------------------------------------------------------
SALIR:
Agregando = False
Exit Sub
error:
    MsgBox Err.Description & vbCr & Err.Source & vbCr & Err.Number, vbCritical, xTitulo
    Err.Clear
End Sub





Private Sub pActualizarCostoLinea()
    '===================================================================================================
    'Creado : 20/05/11 Johan Castro
    'Propósito: Actualizar los registros en la data correspondiente al tipo Lineas de Produccion
    '
    'Entradas:  Ninguno
    '
    'Resultados: Datos en pantalla, detalle y resumen
    '
    '===================================================================================================

    Dim nSQL As String
    Dim nSQLWhere As String '--Cadena SQL para filtrar solo las lineas que tienen costo en todas sus tareas
    Dim RstCosto As New ADODB.Recordset '--Utilizado para indicar 1.- Lineas sin costo. 2.- Lineas solo con costo
    Dim RstTarea As New ADODB.Recordset '--Utilizado para indicar el detalle de los registros
    Dim nSQLFiltro  As String
    
    '--actualizando el destajo individual
    lbl(0).Caption = "Actualizando Costos Linea"
    lbl(1).Caption = "No Interrumpir"
    DoEvents
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        If OptSeleccion(0).Value = True Then
            nSQLFiltro = " and pro_controltar.idarea=" & NulosN(lbl_cod(0).Caption) & " "
        Else
            nSQLFiltro = " and pro_controltardetgr.idper=" & NulosN(lbl_cod(0).Caption) & " "
        End If
        
    End If
    
    '-------------------------------------------------------------------------------------------------------
    '--limpiando los costos
    '--se procede a poner valores a cero los campos que se utilizan para el pago. IdUnid=0, PreUni=0, ImpTot=0
    nSQL = "UPDATE pro_controltar INNER JOIN (pro_controltardet INNER JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) " _
            + vbCr + " AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pro_controltar.id = pro_controltardet.idctr " _
            + vbCr + " SET pro_controltardetgr.idunid = 0, pro_controltardetgr.preuni = 0, pro_controltardetgr.imptot = 0 " _
            + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.tipo)=3)) " & nSQLFiltro
    '--ejecutando consutlta de actualizacion
    xCon.Execute nSQL
    
    '-------------------------------------------------------------------------------------------------------
    '--Recalculando los costos
    
    '--Consulta para determinar las lineas que falta configurar el costo
    nSQL = ""
    nSQL = "SELECT  Linea.idctr & '*' &  Linea.corr & '*' &  Linea.idrec as codigopk,  Linea.idctr, Linea.corr,Linea.idrec, Linea.producto, Costo.costo " _
     + vbCr + " From " _
     + vbCr + " ( SELECT [pro_controltardet].[idrec] & '*' & [pro_controltardettar].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltar.fchtra, pro_controltardet.idctr, pro_controltardet.corr, pro_controltardet.idrec, pro_controltardettar.idtar, pro_controltardet.idunimed, pro_receta.descripcion AS producto, pro_tareas.descripcion AS tarea, pro_controltardet.tipo " _
     + vbCr + " FROM pro_controltar INNER JOIN (((pro_controltardet INNER JOIN pro_controltardettar ON (pro_controltardet.idctr = pro_controltardettar.idctr) AND (pro_controltardet.corr = pro_controltardettar.corr)) INNER JOIN pro_tareas ON pro_controltardettar.idtar = pro_tareas.id) INNER JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) ON pro_controltar.id = pro_controltardet.idctr " _
     + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.tipo)=3) AND ((pro_controltardettar.activo)=-1)) " _
     + vbCr + " ) AS Linea LEFT JOIN " _
     + vbCr + " ( SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden " _
     + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
     + vbCr + " ) AS Costo " _
     + vbCr + " ON Linea.codigopk = Costo.codigopk " _
     + vbCr + " GROUP BY Linea.idctr, Linea.corr, Linea.idrec, Linea.producto, Costo.costo " _
     + vbCr + " HAVING (((Costo.costo)=0 Or (Costo.costo) Is Null));"
     
     RST_Busq RstCosto, nSQL, xCon
     
    '--Generando el filtro de lineas que faltan asignar costo
    nSQLWhere = RstRegistroGenerarId(RstCosto, "codigopk", "", "NOT IN", False)
    If nSQLWhere <> "" Then nSQLWhere = Replace(nSQLWhere, "codigopk", " WHERE Linea.idctr & '*' & Linea.corr & '*' & Linea.idrec")
    '--liberando el recordset
    Set RstCosto = Nothing
    
    
    '--consulta de costo x producto segun linea
    nSQL = "SELECT Linea.idctr, Linea.corr,Linea.idrec, Linea.producto, sum(Costo.costo) as totcosto " _
     + vbCr + " From " _
     + vbCr + " ( SELECT [pro_controltardet].[idrec] & '*' & [pro_controltardettar].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltar.fchtra, pro_controltardet.idctr, pro_controltardet.corr, pro_controltardet.idrec, pro_controltardettar.idtar, pro_controltardet.idunimed, pro_receta.descripcion AS producto, pro_tareas.descripcion AS tarea, pro_controltardet.tipo " _
     + vbCr + " FROM pro_controltar INNER JOIN (((pro_controltardet INNER JOIN pro_controltardettar ON (pro_controltardet.idctr = pro_controltardettar.idctr) AND (pro_controltardet.corr = pro_controltardettar.corr)) INNER JOIN pro_tareas ON pro_controltardettar.idtar = pro_tareas.id) INNER JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) ON pro_controltar.id = pro_controltardet.idctr " _
     + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.tipo)=3) AND ((pro_controltardettar.activo)=-1)) " _
     + vbCr + " ) AS Linea LEFT JOIN " _
     + vbCr + " ( SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden " _
     + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
     + vbCr + " ) AS Costo " _
     + vbCr + " ON Linea.codigopk = Costo.codigopk " _
     + vbCr + " " & nSQLWhere _
     + vbCr + " GROUP BY Linea.idctr, Linea.corr, Linea.idrec, Linea.producto;"
    
    RST_Busq RstCosto, nSQL, xCon
    
    If RstCosto.RecordCount <> 0 Then
    
        '--cargando listado de personal
        
        '--consulta de todos los registros de personal vs producto
        nSQL = "SELECT pro_controltardet.idctr, pro_controltardet.corr, pro_controltardet.idrec, pro_controltardetgr.idper as idemp, pro_controltardet.idunimed, pro_controltar.fchtra, pro_controltardetgr.canpro, pro_controltardetgr.idunid, pro_controltardetgr.preuni, pro_controltardetgr.imptot " _
            + vbCr + " FROM pro_controltar INNER JOIN (pro_controltardet INNER JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pro_controltar.id = pro_controltardet.idctr " _
            + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltardet.tipo)=3) AND ((pro_controltardetgr.activo)=-1)) " & nSQLFiltro
        
        
        RST_Busq RstTarea, nSQL, xCon
        
        PgBar.Min = 0
        PgBar.Value = 0
        
        If RstTarea.RecordCount <> 0 Then
            PgBar.Max = RstTarea.RecordCount
            RstTarea.MoveFirst
            Do While Not RstTarea.EOF
                PgBar.Value = RstTarea.Bookmark
                RstCosto.Filter = ""
                RstCosto.Filter = " idctr = " & RstTarea("idctr") & " and corr = " & RstTarea("corr") & " and idrec = " & NulosN(RstTarea("idrec"))
                If RstCosto.RecordCount <> 0 Then
                    xCon.Execute "Update pro_controltardetgr set idunid=" & NulosN(RstTarea("idunimed")) & ", preuni = " & NulosN(RstCosto("totcosto")) & ", imptot =" & NulosN(RstTarea("canpro")) * NulosN(RstCosto("totcosto")) & "" _
                                + vbCr + " where idctr = " & NulosN(RstTarea("idctr")) & " and corr = " & NulosN(RstTarea("corr")) & " and idper = " & NulosN(RstTarea("idemp")) & " "
                End If
                RstTarea.MoveNext
            Loop
        End If
    
    End If
    
    Set RstCosto = Nothing
    
    Set RstTarea = Nothing
    
  
        
SALIR:
Agregando = False

End Sub

Private Sub FraTarea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    FraTarea.ZOrder 0
End Sub

Private Sub FraTarea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With FraTarea
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub


Private Sub Frm4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frm4.ZOrder 0
End Sub

Private Sub Frm4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frm4
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub
