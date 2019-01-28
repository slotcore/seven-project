VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmCentroCostos1 
   Caption         =   "Contabilidad - Centro de Costos"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "[ Buscar por ]"
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
      Height          =   1050
      Left            =   30
      TabIndex        =   24
      Top             =   360
      Width           =   1365
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Fch. Reg."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   26
         Top             =   465
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Fch. Doc."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   25
         Top             =   225
         Width           =   1065
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6150
      Left            =   30
      TabIndex        =   21
      Top             =   1470
      Width           =   11940
      _cx             =   21061
      _cy             =   10848
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "   Detalle   |  Resumen  "
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
      BoldCurrent     =   0   'False
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
         Height          =   5730
         Left            =   45
         TabIndex        =   22
         Top             =   45
         Width           =   11850
         _cx             =   20902
         _cy             =   10107
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
         BackColor       =   14745342
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   128
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14745342
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
         Rows            =   1
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCentroCostos1.frx":0000
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
         Height          =   5730
         Left            =   12585
         TabIndex        =   23
         Top             =   45
         Width           =   11850
         _cx             =   20902
         _cy             =   10107
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
         BackColor       =   14745342
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   128
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14745342
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
         Rows            =   3
         Cols            =   7
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCentroCostos1.frx":01E3
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
   Begin VB.Frame Frame6 
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
      Height          =   1050
      Left            =   10425
      TabIndex        =   17
      Top             =   360
      Width           =   1485
      Begin VB.OptionButton OptLib 
         Caption         =   "Honorarios"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   20
         Top             =   705
         Width           =   1125
      End
      Begin VB.OptionButton OptLib 
         Caption         =   "Todos"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   19
         Top             =   225
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.OptionButton OptLib 
         Caption         =   "Compras"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   18
         Top             =   465
         Width           =   1050
      End
   End
   Begin VB.Frame FraBarra 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   690
      Left            =   1920
      TabIndex        =   14
      Top             =   7770
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Top             =   300
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Consultando"
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
         Left            =   90
         TabIndex        =   16
         Top             =   60
         Width           =   1050
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   5940
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   5925
         X2              =   5925
         Y1              =   15
         Y2              =   945
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   30
         X2              =   5940
         Y1              =   675
         Y2              =   660
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Selecc.Fecha]"
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
      Height          =   1050
      Left            =   1410
      TabIndex        =   9
      Top             =   360
      Width           =   1785
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   405
         TabIndex        =   10
         Top             =   270
         Width           =   1290
         _ExtentX        =   2275
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
         Valor           =   "11/01/2009"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   405
         TabIndex        =   11
         Top             =   645
         Width           =   1290
         _ExtentX        =   2275
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
         Valor           =   "11/01/2009"
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Top             =   750
         Width           =   135
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   330
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Ordenado por ]"
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
      Height          =   1050
      Left            =   8790
      TabIndex        =   5
      Top             =   360
      Width           =   1635
      Begin VB.OptionButton opt 
         Caption         =   "Nº Documento"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   465
         Width           =   1350
      End
      Begin VB.OptionButton opt 
         Caption         =   "Fch Emisión"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   225
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton opt 
         Caption         =   "Nº Registro"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   705
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Centro de Costo]"
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
      Height          =   1050
      Left            =   3210
      TabIndex        =   1
      Top             =   360
      Width           =   5565
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Agregar "
         Height          =   375
         Left            =   4710
         TabIndex        =   3
         Top             =   210
         Width           =   795
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   4710
         TabIndex        =   2
         Top             =   615
         Width           =   795
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg3 
         Height          =   780
         Left            =   60
         TabIndex        =   4
         Top             =   210
         Width           =   4620
         _cx             =   8149
         _cy             =   1376
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCentroCostos1.frx":02E6
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12000
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":0367
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":08AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":0C3D
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":0DC1
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":1215
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":132D
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":1871
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":1DB5
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":1EC9
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":1FDD
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":2431
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":259D
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":2AE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":2DFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":3191
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostos1.frx":3523
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12405
      _ExtentX        =   21881
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
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Configurar Formatos"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar a PDT"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmCentroCostos1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCENTROCOSTO.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : MUESTRA EN FORMA DETALLADA Y RESUMIDA LOS SALDOS POR CADA CENTRO DE COSTOS, EN
'*                    FUNCION A CRITERIOS ESPECIFICADOS POR EL USUARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 26/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim SeEjecuto As Boolean       ' ESPECIFICA SI SE EJECUTO EL EVENTO ACTIVATE DEL FORMULARIO
Dim Agregando As Boolean

Private Sub Exportar()
    ' EXPORTAR A EXCEL LOS DATOS DEL CONTROL Fg1
    If Fg2.Rows = 2 Then
        MsgBox "No hay registro que exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    Dim xFun As New SGI2_funciones.formularios
    Dim nPeriodo As String
    nPeriodo = "Del " & TxtFchIni.Valor & " Al " & TxtFchFin.Valor
    
    If TabOne1.CurrTab = 0 Then
        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "CENTRO DE COSTOS - DETALLADO", nPeriodo, , "CENTRO DE COSTOS DETALLADO"
    Else
        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg2, "CENTRO DE COSTOS - RESUMIDO", nPeriodo, , "CENTRO DE COSTO RESUMIDO"
    End If
        
    Set xFun = Nothing
End Sub




'*****************************************************************************************************
'* Nombre           : Detalle
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA DE FORMA DETALLADA LOS SALDOS DE DE CADA CENTRO DE COSTO, EN FUNCION A
'*                    CRITERIOS ESPECIFICADOS POR EL USUARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Detalle()
    Dim A, B, xFila As Double
    Dim xRst As New ADODB.Recordset
    Dim xTotal, xTotalTotal As Double
    
    Dim nSQL As String '--Sentencia SQL para la consulta
    Dim FilaInicioGrupo As Long '--variable para almacenar la fila donde se inicia un grupo
                            '--util para hacer las sumas de los importes
                            
    Dim TotImpCosMn As Double '--Acumular los subtotales en moneda nacional
    Dim TotImpCosMe As Double '--Acumular los subtotales en moneda extranjera
    Dim TotImpCosExpMn As Double '--Acumular los subtotales expresado en moneda nacional
    Dim TotImpCosExpMe As Double '--Acumular los subtotales expresado en moneda extranjera
    Dim nSQLCencos As String    '--Sentencia SQL para filtrar centros de costo
    Dim SQLCampoOrden As String '--Indica el campo para aplicar el orden
    Dim nSQLVerCompras As String '--Condicion para mostrar registro de compras
    Dim nSQLVerHonorarios As String '--Condicion para mostrar registro de Honorarios
    Dim SQLCampoFecha As String  '--Indica el tipo de filtro por fecha(x fch documento,x fch registro)
    
    '--mostrar la barra de progreso
    FraBarra.Visible = True
    FraBarra.Left = 2850
    FraBarra.Top = 3660
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Value = 0
    Label3.Caption = "Procesando Detalle"
    
    '--Verificar el tipo de filtro de fechas
    If opt_fecha(0).Value = True Then
        SQLCampoFecha = "com_compras.fchdoc"
    Else
        SQLCampoFecha = "com_compras.fchreg"
    End If
    
    '--Verificar los registros a mostrar
    If OptLib(0).Value = True Then '--Muestra todos
        nSQLVerCompras = ""
        nSQLVerHonorarios = ""
    ElseIf OptLib(1).Value = True Then '--Muestra Registro de Compras
        nSQLVerCompras = ""
        nSQLVerHonorarios = " and 1=0 "
    ElseIf OptLib(2).Value = True Then '--Muestra Registro de Honorarios
        nSQLVerCompras = " and 1=0 "
        nSQLVerHonorarios = ""
    End If
    
    '--Generar el filtro por centro de costo, saegun criterio del usuario
    For A = 1 To Fg3.Rows - 1
        If Trim(Fg3.TextMatrix(A, 1)) <> "" Then
            nSQLCencos = nSQLCencos + " con_centrocosto.codigo Like '" & Trim(Fg3.TextMatrix(A, 1)) & "%' OR "
        End If
    Next A
    If nSQLCencos <> "" Then nSQLCencos = " AND (" + Left(nSQLCencos, Len(nSQLCencos) - 3) + ") "
        
    '--aplicamos el orden de persentacion de los datos
    If opt(0).Value = True Then '--Fch emision
        SQLCampoOrden = ", CenCos.fchdoc "
    ElseIf opt(1).Value = True Then '--Nro Documento
        SQLCampoOrden = ", CenCos.numerodoc "
    ElseIf opt(2).Value = True Then '--Nro Registro
        SQLCampoOrden = ", CenCos.registro "
    End If

    ' CREAMOS LA SENTENCIA SQL PARA CARGAR LOS DATOS
    '--compras
    nSQL = "SELECT con_centrocosto.id, con_centrocosto.codigo, con_centrocosto.descripcion, Mid([com_compras]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_compras]![numreg],3,4) AS registro, mae_documento.abrev, [com_compras]![numser] & '-'  & [com_compras]![numdoc] AS numerodoc, com_compras.fchdoc, " _
        + vbCr + " com_compras.fchven, mae_moneda.simbolo,mae_prov.numruc, mae_prov.nombre, com_comprascosto.imppor, com_compras.fchreg, com_compras.glosa, " _
        + vbCr + " iif(com_compras.tc=0 or com_compras.tc is null, con_tc.impven,com_compras.tc) as tipcam, " _
        + vbCr + " iif(com_compras.tipdoc=7, (-1) * [com_comprascosto]![impcos],[com_comprascosto]![impcos]) as impcosreal, " _
        + vbCr + " iif(com_compras.idmon=1, impcosreal,0 ) as impcosmn, " _
        + vbCr + " iif(com_compras.idmon=2, impcosreal,0 ) as impcosme, " _
        + vbCr + " iif(com_compras.idmon=1, impcosmn,impcosme * tipcam ) as impcosexpmn, " _
        + vbCr + " iif(com_compras.idmon = 2, impcosme, iif(tipcam = 0, 0, impcosmn / tipcam)) As impcosexpme " _
        + vbCr + " FROM (con_centrocosto LEFT JOIN (mae_documento RIGHT JOIN (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN ((com_comprascosto LEFT JOIN com_compras ON com_comprascosto.idcom = com_compras.id) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = com_compras.idmon) ON mae_documento.id = com_compras.tipdoc) ON con_centrocosto.id = com_comprascosto.idcencos) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE (((" & SQLCampoFecha & ") Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLCencos & nSQLVerCompras
    
    '--honorarios
    '--reemplazando el campo para filtrar la fecha
    SQLCampoFecha = Replace(SQLCampoFecha, "com_compras", "com_honorarios")
    
    nSQL = nSQL _
        + vbCr + " UNION " _
        + vbCr + " SELECT con_centrocosto.id, con_centrocosto.codigo, con_centrocosto.descripcion, Mid([com_honorarios]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_honorarios]![numreg],3,4) AS registro, mae_documento.abrev, [com_honorarios]![numser] & '-'  & [com_honorarios]![numdoc] AS numerodoc, com_honorarios.fchdoc, " _
        + vbCr + " com_honorarios.fchven, mae_moneda.simbolo,mae_prov.numruc, mae_prov.nombre, com_honorarioscosto.imppor, com_honorarios.fchreg, com_honorarios.glosa, " _
        + vbCr + " iif(com_honorarios.tc=0 or com_honorarios.tc is null, con_tc.impven,com_honorarios.tc) as tipcam, " _
        + vbCr + " iif(com_honorarios.tipdoc=7, (-1) * [com_honorarioscosto]![impcos],[com_honorarioscosto]![impcos]) as impcosreal, " _
        + vbCr + " iif(com_honorarios.idmon=1, impcosreal,0 ) as impcosmn, " _
        + vbCr + " iif(com_honorarios.idmon=2, impcosreal,0 ) as impcosme, " _
        + vbCr + " iif(com_honorarios.idmon=1, impcosmn,impcosme * tipcam ) as impcosexpmn, " _
        + vbCr + " iif(com_honorarios.idmon = 2, impcosme, iif(tipcam = 0, 0, impcosmn / tipcam)) As impcosexpme " _
        + vbCr + " FROM (con_centrocosto LEFT JOIN (mae_documento RIGHT JOIN (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN ((com_honorarioscosto LEFT JOIN com_honorarios ON com_honorarioscosto.idcom = com_honorarios.id) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) ON mae_prov.id = com_honorarios.idpro) ON mae_moneda.id = com_honorarios.idmon) ON mae_documento.id = com_honorarios.tipdoc) ON con_centrocosto.id = com_honorarioscosto.idcencos) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE (((" & SQLCampoFecha & ") Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLCencos & nSQLVerHonorarios

    '--generamos una select de la subconsulta para aplicar el orden
    nSQL = "SELECT CenCos.* " _
        + vbCr + " FROM (" & nSQL & ") as CenCos " _
        + vbCr + " ORDER BY CenCos.codigo " & SQLCampoOrden
                     
    '--ejecutar la consulta
    RST_Busq xRst, nSQL, xCon

    Dim xCodCenCos As String
    Dim xCad As String

    '--Colocar la posicion de inicio de grupo
    FilaInicioGrupo = Fg1.FixedRows
    
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        '--asignar total de valores a la barra de progreso
        ProgressBar1.Max = xRst.RecordCount
        
        xCodCenCos = xRst("codigo")
        
        Fg1.Rows = Fg1.Rows + 1
        xCad = "CENTRO COSTO Nº ==> " & NulosC(xRst("codigo")) & "  " & UCase(NulosC(xRst("descripcion")))
        GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 10, xCad, flexAlignLeftCenter, True, , &H800000, &HE2FEFE, True
        
        ' ESCRIBIMOS LOS DATOS DEL RECORDSET EN LAS FILAS DEL CONTROL Fg1
        For A = 1 To xRst.RecordCount
            DoEvents
            ProgressBar1.Value = A
            
            Fg1.Rows = Fg1.Rows + 1

            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRst("registro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRst("numruc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRst("nombre"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRst("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(xRst("numerodoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(xRst("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(xRst("fchven"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(xRst("simbolo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(xRst("tipcam"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosC(xRst("glosa"))
                        
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(NulosN(xRst("impcosme")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(NulosN(xRst("impcosmn")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(xRst("impcosexpme")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(NulosN(xRst("impcosexpmn")), FORMAT_MONTO)
            '--datos del centro de costo
            Fg1.TextMatrix(Fg1.Rows - 1, 15) = NulosC(xRst("codigo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = NulosC(xRst("descripcion"))
            Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(NulosN(xRst("imppor")), FORMAT_MONTO)
            '--datos del plan de cuenta >>> PENDIENTE
'            Fg1.TextMatrix(Fg1.Rows - 1, 18) = NulosC(xRst("numcta"))
'            Fg1.TextMatrix(Fg1.Rows - 1, 19) = NulosC(xRst("nomcta"))
            
            xRst.MoveNext
            If xRst.EOF = True Then Exit For
            
            If xCodCenCos <> xRst("codigo") Then
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = "TOTAL ==>"
                
                Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(GRID_SUMAR_COL(Fg1, 11, FilaInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(GRID_SUMAR_COL(Fg1, 12, FilaInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(GRID_SUMAR_COL(Fg1, 13, FilaInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(GRID_SUMAR_COL(Fg1, 14, FilaInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
                
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, &H800000, True, &HE2FEFE
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H800000, True, &HE2FEFE
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H800000, True, &HE2FEFE
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H800000, True, &HE2FEFE
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, &H800000, True, &HE2FEFE
                
                '--acumulando los subtotales para mostrar en el total general
                TotImpCosMe = TotImpCosMe + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 11)) '--moneda extranjera
                TotImpCosMn = TotImpCosMn + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 12)) '--moneda nacional
                TotImpCosExpMe = TotImpCosExpMe + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 13)) '--expresado en moneda extranjera
                TotImpCosExpMn = TotImpCosExpMn + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 14)) '--expresado en moneda nacional
                                
                xCodCenCos = NulosC(xRst("codigo"))
                xTotal = 0
                Fg1.Rows = Fg1.Rows + 2
                xCad = "CENTRO COSTO Nº ==> " & NulosC(xRst("codigo")) & "  " & UCase(NulosC(xRst("descripcion")))
                GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 17, xCad, flexAlignLeftCenter, True, , &H800000, &HE2FEFE, True
                
                '--posicionar en nueva fila para inicio de grupo
                FilaInicioGrupo = Fg1.Rows - 1
                
            End If
        Next A
    
        '--muestra el subtotal del ultimo grupo
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = "TOTAL ==>"
        
        Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(GRID_SUMAR_COL(Fg1, 11, FilaInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(GRID_SUMAR_COL(Fg1, 12, FilaInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(GRID_SUMAR_COL(Fg1, 13, FilaInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(GRID_SUMAR_COL(Fg1, 14, FilaInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, &H800000, True, &HE2FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H800000, True, &HE2FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H800000, True, &HE2FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H800000, True, &HE2FEFE
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, &H800000, True, &HE2FEFE
        
        '--acumulando los subtotales para mostrar en el total general
        TotImpCosMe = TotImpCosMe + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 11)) '--moneda extranjera
        TotImpCosMn = TotImpCosMn + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 12)) '--moneda nacional
        TotImpCosExpMe = TotImpCosExpMe + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 13)) '--expresado en moneda extranjera
        TotImpCosExpMn = TotImpCosExpMn + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 14)) '--expresado en moneda nacional
        
        '--muestra el total general del ultimo grupo
        Fg1.Rows = Fg1.Rows + 2
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, &H800000, True, &HE2FEFE, "TOTAL GRAL==>"
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H800000, True, &HE2FEFE, Format(TotImpCosMe, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H800000, True, &HE2FEFE, Format(TotImpCosMn, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H800000, True, &HE2FEFE, Format(TotImpCosExpMe, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, &H800000, True, &HE2FEFE, Format(TotImpCosExpMn, FORMAT_MONTO)
        
        '--ajustando las columnas de acuerdo a los importes
        Fg1.AutoSizeMode = flexAutoSizeColWidth
        Fg1.AutoSize 11
        Fg1.AutoSize 12
        Fg1.AutoSize 13
        Fg1.AutoSize 14
        
    End If
    
    Set xRst = Nothing
    
    '--ocultando la barra de progreso
    FraBarra.Visible = False
End Sub

'*****************************************************************************************************
'* Nombre           : Resumen
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EN FORMA RESUMIDA EL SALDO DE CADA CENTRO DE COSTO, EN FUNCION A
'*                    CRITERIOS ESPECIFICADOS POR EL USUARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Resumen()
    Dim A As Long
    Dim xRst As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLCencos As String    '--Sentencia SQL para filtrar centros de costo
    Dim nSQLVerCompras As String '--Condicion para mostrar registro de compras
    Dim nSQLVerHonorarios As String '--Condicion para mostrar registro de Honorarios
    Dim SQLCampoFecha As String  '--Indica el tipo de filtro por fecha(x fch documento,x fch registro)

    
    '--mostrar la barra de progreso
    FraBarra.Visible = True
    FraBarra.Left = 2850
    FraBarra.Top = 3660
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Value = 0
    Label3.Caption = "Procesando Resumen"
    
    '--Verificar el tipo de filtro de fechas
    If opt_fecha(0).Value = True Then
        SQLCampoFecha = "com_compras.fchdoc"
    Else
        SQLCampoFecha = "com_compras.fchreg"
    End If
    
    '--Verificar los registros a mostrar
    If OptLib(0).Value = True Then '--Muestra todos
        nSQLVerCompras = ""
        nSQLVerHonorarios = ""
    ElseIf OptLib(1).Value = True Then '--Muestra Registro de Compras
        nSQLVerCompras = ""
        nSQLVerHonorarios = " and 1=0 "
    ElseIf OptLib(2).Value = True Then '--Muestra Registro de Honorarios
        nSQLVerCompras = " and 1=0 "
        nSQLVerHonorarios = ""
    End If
    
    '--Generar el filtro por centro de costo, saegun criterio del usuario
    For A = 1 To Fg3.Rows - 1
        If Trim(Fg3.TextMatrix(A, 1)) <> "" Then
            nSQLCencos = nSQLCencos + " con_centrocosto.codigo Like '" & Trim(Fg3.TextMatrix(A, 1)) & "%' OR "
        End If
    Next A
    If nSQLCencos <> "" Then nSQLCencos = " AND (" + Left(nSQLCencos, Len(nSQLCencos) - 3) + ") "
    '-------------------
    
    ' CREAMOS LA SENTENCIA SQL PARA CARGAR LOS DATOS

    '--compras
    nSQL = "SELECT con_centrocosto.id, con_centrocosto.codigo, con_centrocosto.descripcion, Mid([com_compras]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_compras]![numreg],3,4) AS registro, mae_documento.abrev, [com_compras]![numser] & '-'  & [com_compras]![numdoc] AS numerodoc, com_compras.fchdoc, " _
        + vbCr + " com_compras.fchven, mae_moneda.simbolo,mae_prov.numruc, mae_prov.nombre, com_comprascosto.imppor, com_compras.fchreg, com_compras.glosa, " _
        + vbCr + " iif(com_compras.tc=0 or com_compras.tc is null, con_tc.impven,com_compras.tc) as tipcam, " _
        + vbCr + " iif(com_compras.tipdoc=7, (-1) * [com_comprascosto]![impcos],[com_comprascosto]![impcos]) as impcosreal, " _
        + vbCr + " iif(com_compras.idmon=1, impcosreal,0 ) as impcosmn, " _
        + vbCr + " iif(com_compras.idmon=2, impcosreal,0 ) as impcosme, " _
        + vbCr + " iif(com_compras.idmon=1, impcosmn,impcosme * tipcam ) as impcosexpmn, " _
        + vbCr + " iif(com_compras.idmon = 2, impcosme, iif(tipcam = 0, 0, impcosmn / tipcam)) As impcosexpme " _
        + vbCr + " FROM (con_centrocosto LEFT JOIN (mae_documento RIGHT JOIN (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN ((com_comprascosto LEFT JOIN com_compras ON com_comprascosto.idcom = com_compras.id) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = com_compras.idmon) ON mae_documento.id = com_compras.tipdoc) ON con_centrocosto.id = com_comprascosto.idcencos) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE (((" & SQLCampoFecha & ") Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLCencos & nSQLVerCompras
    
    '--honorarios
    '--reemplazando el campo para filtrar la fecha
    SQLCampoFecha = Replace(SQLCampoFecha, "com_compras", "com_honorarios")
    
    nSQL = nSQL _
        + vbCr + " UNION " _
        + vbCr + " SELECT con_centrocosto.id, con_centrocosto.codigo, con_centrocosto.descripcion, Mid([com_honorarios]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_honorarios]![numreg],3,4) AS registro, mae_documento.abrev, [com_honorarios]![numser] & '-'  & [com_honorarios]![numdoc] AS numerodoc, com_honorarios.fchdoc, " _
        + vbCr + " com_honorarios.fchven, mae_moneda.simbolo,mae_prov.numruc, mae_prov.nombre, com_honorarioscosto.imppor, com_honorarios.fchreg, com_honorarios.glosa, " _
        + vbCr + " iif(com_honorarios.tc=0 or com_honorarios.tc is null, con_tc.impven,com_honorarios.tc) as tipcam, " _
        + vbCr + " iif(com_honorarios.tipdoc=7, (-1) * [com_honorarioscosto]![impcos],[com_honorarioscosto]![impcos]) as impcosreal, " _
        + vbCr + " iif(com_honorarios.idmon=1, impcosreal,0 ) as impcosmn, " _
        + vbCr + " iif(com_honorarios.idmon=2, impcosreal,0 ) as impcosme, " _
        + vbCr + " iif(com_honorarios.idmon=1, impcosmn,impcosme * tipcam ) as impcosexpmn, " _
        + vbCr + " iif(com_honorarios.idmon = 2, impcosme, iif(tipcam = 0, 0, impcosmn / tipcam)) As impcosexpme " _
        + vbCr + " FROM (con_centrocosto LEFT JOIN (mae_documento RIGHT JOIN (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN ((com_honorarioscosto LEFT JOIN com_honorarios ON com_honorarioscosto.idcom = com_honorarios.id) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) ON mae_prov.id = com_honorarios.idpro) ON mae_moneda.id = com_honorarios.idmon) ON mae_documento.id = com_honorarios.tipdoc) ON con_centrocosto.id = com_honorarioscosto.idcencos) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE (((" & SQLCampoFecha & ") Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLCencos & nSQLVerHonorarios

    '--generamos una select de la subconsulta para aplicar las sumas
    nSQL = "SELECT CenCos.codigo, CenCos.descripcion, Sum(CenCos.impcosmn) AS totimpcosmn, Sum(CenCos.impcosme) AS totimpcosme, Sum(CenCos.impcosexpmn) AS totimpcosexpmn, Sum(CenCos.impcosexpme) AS totimpcosexpme " _
        + vbCr + " FROM (" & nSQL & ") as CenCos " _
        + vbCr + " GROUP BY CenCos.codigo, CenCos.descripcion " _
        + vbCr + " ORDER BY CenCos.codigo "
                     
    '--ejecutar la consulta
    RST_Busq xRst, nSQL, xCon
    
    '--verificar que este activo el recordset
    If xRst.State = 0 Then GoTo LaCague
    
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        
        '--asignar total de valores a la barra de progreso
        ProgressBar1.Max = xRst.RecordCount
        
        For A = 1 To xRst.RecordCount
            DoEvents
            ProgressBar1.Value = A
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(xRst("codigo"))
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(xRst("descripcion"))
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(NulosN(xRst("totimpcosme")), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(xRst("totimpcosmn")), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(NulosN(xRst("totimpcosexpme")), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(NulosN(xRst("totimpcosexpmn")), FORMAT_MONTO)
            
            xRst.MoveNext
            If xRst.EOF = True Then Exit For
        Next A
        
        '--colocando el total
        Fg2.Rows = Fg2.Rows + 1
        FORMATO_CELDA Fg2, Fg2.Rows - 1, 2, &H800000, True, &HE2FEFE, "TOTAL ==> "
        FORMATO_CELDA Fg2, Fg2.Rows - 1, 3, &H800000, True, &HE2FEFE, Format(GRID_SUMAR_COL(Fg2, 3), FORMAT_MONTO)
        FORMATO_CELDA Fg2, Fg2.Rows - 1, 4, &H800000, True, &HE2FEFE, Format(GRID_SUMAR_COL(Fg2, 4), FORMAT_MONTO)
        FORMATO_CELDA Fg2, Fg2.Rows - 1, 5, &H800000, True, &HE2FEFE, Format(GRID_SUMAR_COL(Fg2, 5), FORMAT_MONTO)
        FORMATO_CELDA Fg2, Fg2.Rows - 1, 6, &H800000, True, &HE2FEFE, Format(GRID_SUMAR_COL(Fg2, 6), FORMAT_MONTO)
        
        
        '--ajustando las columnas de acuerdo a los importes
        Fg2.AutoSizeMode = flexAutoSizeColWidth
        Fg2.AutoSize 3
        Fg2.AutoSize 4
        Fg2.AutoSize 5
        Fg2.AutoSize 6
        
    End If
    
LaCague:
    Set xRst = Nothing
    
    '--ocultando la barra de progreso
    FraBarra.Visible = False
End Sub

Private Sub Consultar()
    ' CARGA LOS DATOS DE LOS CENTROS DE COSTO
    
    ' VERIFICAMOS QUE LOS DATOS NECESARIOS SEAN LOS CORRECTOS
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio para la consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "No ha especificado la fecha final para la consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    '--incializar la grilla
    Fg1.Rows = Fg1.FixedRows
    Fg2.Rows = Fg2.FixedRows
    
    If Fg1.Cols = 2 Then
        MsgBox "Falta configurar la presentación del reporte" & vbCr & "Consulte con el administrador del sistema", vbInformation, xTitulo
        Exit Sub
    End If
    
    
    DoEvents
    TabOne1.CurrTab = 1
    Detalle    ' CARGAMOS EL DETALLE DE LOS CENTROS DE COSTO
    TabOne1.CurrTab = 0
    Resumen    ' CARGAMOS EL RESUMEN DE LOS CENTROS DE COSTO
    
    
End Sub


Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        TxtFchIni.SetFocus
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    
    
    SeEjecuto = False
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    Frame2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F

    ' CONFIGURA EL RESUMEN DEL CENTRO DE COSTO
    SetearCuadricula Fg2, 6, xCon, 1, 1
    ' CONFIGURA EL DETALLE DEL CENTRO DE COSTO
    SetearCuadricula Fg1, 6, xCon, 1, 2


    Fg3.Rows = 1
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    Fg3.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDNone
    Fg2.Editable = flexEDNone
    Fg3.Editable = flexEDNone
    
    Fg3.ColWidth(3) = 0
    GRID_COMBOLIST Fg3, 1
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height > 3000 Then
        TabOne1.Top = 1470
        TabOne1.Width = Me.Width - 150
        TabOne1.Height = Me.Height - 1890
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        If IsDate(TxtFchIni.Valor) = False Then
            MsgBox "No ha especificado la fecha inicial", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            TxtFchIni.SetFocus
            Exit Sub
        End If

        If IsDate(TxtFchFin.Valor) = False Then
            MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            TxtFchFin.SetFocus
            Exit Sub
        End If

        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            TxtFchIni.SetFocus
            Exit Sub
        End If
        
        Consultar
        
    End If
    
    If Button.Index = 3 Then
        If Fg1.Rows = Fg1.FixedRows Then
            MsgBox "No hay registro que exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Exportar
    End If
    
    If Button.Index = 4 Then
'        Dim nPeriodo As String
'        Dim xPrint As New SGI2_funciones.formularios
'        nPeriodo = "Del " & TxtFchIni.Valor & " Al " & TxtFchFin.Valor
'        Me.MousePointer = vbHourglass
'        If TabOne1.CurrTab = 0 Then
'            xPrint.Imprimir_x_VSFlexGrid Fg1, "CENTRO DE COSTOS DETALLADO", " ", nPeriodo, False, True
'        Else
'            xPrint.Imprimir_x_VSFlexGrid Fg2, "RESUMEN DE CENTRO DE COSTOS ", " ", nPeriodo, False, True
'        End If
'
'        Set xPrint = Nothing
'        Me.MousePointer = vbDefault
        pImprimir
    End If
    
    'If Button.Index = 5 Then Configurar
    
    'If Button.Index = 6 Then ExportarPDT
        
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub


Private Sub CmdAdd_Click()
    If Fg3.Rows = 1 Then
        Fg3.Rows = Fg3.Rows + 1
        Exit Sub
    End If
    If NulosN(Fg3.TextMatrix(Fg3.Rows - 1, 3)) = 0 Then Exit Sub
    Fg3.Rows = Fg3.Rows + 1
    Fg3.Row = Fg3.Rows - 1
    Fg3.Col = 1
    Fg3.SetFocus
End Sub


Private Sub CmdDel_Click()
    If Fg3.Row <= 0 Then Exit Sub
    If Fg3.Rows <= 1 Then
        MsgBox "No hay cuentas seleccionadas para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    Else
        Fg3.RemoveItem Fg3.Row
        Fg3.Refresh
    End If
End Sub



Private Sub fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error

    If Col = 1 Then
        Dim Rst As New ADODB.Recordset
        Dim nSQL As String
        Dim nSQLLike As String
        Dim nSQLIdCta As String
          
        Dim xCampos(2, 4) As String
        
        xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "codigo":             xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":        xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
        
        '--generando el filtro por centro de costo
        nSQLIdCta = GRID_GENERAR_SQL_ID(Fg3, 3, " WHERE con_centrocosto.id", " NOT IN ", True, 1, Fg3.TextMatrix(Row, 3))
        '--Verificar si filtro por centro de costo para buscar
        If NulosC(Fg3.TextMatrix(Fg3.Row, 1)) <> "" Then
            nSQLLike = " and con_centrocosto.codigo like '" + Trim(Fg3.TextMatrix(Fg3.Row, 1)) + "%' "
        End If
        If nSQLIdCta = "" Then nSQLLike = Replace(nSQLLike, " and ", " WHERE ")
        '--Armar la sentencia SQL
        nSQL = "SELECT con_centrocosto.codigo, con_centrocosto.descripcion, con_centrocosto.id " _
            + vbCr + " From con_centrocosto " + nSQLIdCta + nSQLLike + vbCr + "  ORDER BY con_centrocosto.codigo"
        '--Cargar la ventana emergente para consultar
        CARGAR_DLL_EPSBUSCAR xCon, Rst, nSQL, xCampos(), "Buscando Centro de Costos", "codigo", "codigo", Principio
        
        If Rst.State = 0 Then GoTo SALIR
        If Rst.RecordCount = 0 Then GoTo SALIR
        
        If fValidarSeleccionCta(NulosC(Rst("codigo"))) = False Then GoTo SALIR

        Agregando = True
    
        Fg3.TextMatrix(Fg3.Row, 1) = NulosC(Rst("codigo"))
        Fg3.TextMatrix(Fg3.Row, 2) = NulosC(Rst("descripcion"))
        Fg3.TextMatrix(Fg3.Row, 3) = NulosN(Rst("id"))
        
        Set Rst = Nothing
    End If
    
SALIR:
    
    Agregando = False
    Exit Sub
error:
    'Resume
    Set Rst = Nothing
    Agregando = False
    MsgBox Err.Description & vbCr & Err.Source & vbCr & Err.Number, vbCritical, xTitulo
    Err.Clear
End Sub

Private Sub fg3_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    
    If Fg3.TextMatrix(Row, Col) = "" Then
        Fg3.TextMatrix(Row, 1) = ""
        Fg3.TextMatrix(Row, 2) = ""
        Fg3.TextMatrix(Row, 3) = ""
        Exit Sub
    End If
    
    If Col = 1 Then
        If fValidarSeleccionCta(NulosN(Fg3.TextMatrix(Row, Col))) = False Then
            Fg3.TextMatrix(Row, 1) = ""
            Fg3.TextMatrix(Row, 2) = ""
            Fg3.TextMatrix(Row, 3) = ""
            Exit Sub
        End If
        
        Dim Rst As New ADODB.Recordset
        RST_Busq Rst, "SELECT * FROM con_centrocosto WHERE codigo = '" & NulosC(Fg3.TextMatrix(Row, 1)) & "'", xCon
        If Rst.RecordCount = 1 Then
            Fg3.TextMatrix(Row, 2) = NulosC(Rst("descripcion"))
            Fg3.TextMatrix(Row, 3) = NulosN(Rst("id"))
        Else
            Fg3.TextMatrix(Row, 1) = ""
            Fg3.TextMatrix(Row, 2) = ""
            Fg3.TextMatrix(Row, 3) = ""
        End If
        Set Rst = Nothing
    End If
   
End Sub

Private Sub fg3_EnterCell()
    If Fg3.Col = 1 Then
        Fg3.Editable = flexEDKbdMouse
    Else
        Fg3.Editable = flexEDNone
    End If
End Sub

Private Sub fg3_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 45 Then
        CmdAdd_Click
    End If
    
    If KeyCode = 46 Then
        CmdDel_Click
    End If
End Sub

Private Sub fg3_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    
    Select Case Col
        Case 1
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub



Private Function fValidarSeleccionCta(NumCuenta As String) As Boolean
    On Error GoTo error
    
    Dim k As Integer
    Dim MSG_CUENTA As String    '--MUSTRA EL MENSAJE SI DESEA AGREGAR UN CENTRO DE COSTO, CUANDO YA EXISTE UN COSTO DE NIVEL SUPERIOR O NIVEL INFERIOR
                                '--NO MOSTRAR MENSAJE SOLO CUANDO LOS COSTOS SEA DEL MISMO NIVEL
    If GRID_BUSCAR_VALOR(Fg3, 1, NumCuenta, False, , Fg3.Row) <> "-1" Then
        MsgBox "El centro de costo Nº " & NumCuenta & " ya fue seleccionada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If
    
    For k = 1 To Fg3.Rows - 1
        If k <> Fg3.Row Then
            If Len(NumCuenta) < Len(Trim(Fg3.TextMatrix(k, 1))) Then
                
                If NumCuenta = Mid(Trim(Fg3.TextMatrix(k, 1)), 1, Len(NumCuenta)) Then
                    MSG_CUENTA = "Ya agregó el costo Nº: " + Trim(Fg3.TextMatrix(k, 1)) + " cuyo nivel es Inferior al costo Nº: " & NumCuenta & " que desea agregar" _
                                + vbCr + "Sólo puede agregar Centro de Costos del mismo nivel " _
                                + vbCr + "Si desea continuar elimine la fila que contenga el Costo Nº: " + Trim(Fg3.TextMatrix(k, 1))
                    Exit For
                End If
                
            Else
                If Trim(Fg3.TextMatrix(k, 1)) = Mid(NumCuenta, 1, Len(Trim(Fg3.TextMatrix(k, 1)))) Then
                    MSG_CUENTA = "Ya agregó el costo Nº: " + Trim(Fg3.TextMatrix(k, 1)) + " cuyo nivel es Superior al costo Nº: " & NumCuenta & " que desea agregar" _
                                + vbCr + "Sólo puede agregar Centro de Costos del mismo nivel " _
                                + vbCr + "Si desea continuar elimine la fila que contenga el Costo Nº: " + Trim(Fg3.TextMatrix(k, 1))
                    Exit For
                End If
                
            End If
        End If
    Next k
    If MSG_CUENTA <> "" Then
        MsgBox MSG_CUENTA, vbExclamation, xTitulo
        GoTo SALIR
    End If
    
    fValidarSeleccionCta = True
SALIR:
    Exit Function
error:
    MsgBox Err.Description & vbCr & Err.Source & vbCr & Err.Number, vbCritical, xTitulo
    Err.Clear
End Function



'*****************************************************************************************************
'* Nombre           : pImprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MANDA A LA IMPRESORA EL REGISTRO DE VENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub pImprimir()
    Dim xMoneda As String
    Dim nPeriodo As String

    If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
        nPeriodo = " Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
    Else
        nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
    End If

'    xMoneda = LblMoneda.Caption

    Dim RstTmp As New ADODB.Recordset
    Dim A As Long
    Dim Rst As New ADODB.Recordset
    ' SELECCIONA EL FORMATO DE IMPRESION ACTUAL PARA EL REGISTRO DE VENTAS
    'xFormatoActual = xRs("id")
    Dim xCampos() As String
    Dim xFil, xCol As Double
    Dim xFila As Double
    
    If TabOne1.CurrTab = 0 Then
        '--verificar si hay registros para imprimir
        If Fg1.Rows <= Fg1.FixedRows Then
            MsgBox "No hay registros para imprimir", vbInformation, xTitulo
            Exit Sub
        End If
        
        '--imprimir el detalle
        RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet " _
            & " Where (((con_formatostipodet.idformato) = 6) And ((con_formatostipodet.idformatotipo) = 2) " _
            & " And ((con_formatostipodet.mostrar) = -1)) ORDER BY con_formatostipodet.orden", xCon
        
        ReDim xCampos(Fg1.Rows - 2, Fg1.Cols - 1)
        
        xFila = 0
        ' PASAMOS LOS DATOS DEL CONTROL Fg1 AL ARRAY DE DATOS
        For xFil = 1 To Fg1.Rows - 1
            For xCol = 1 To Fg1.Cols - 1
                xCampos(xFila, xCol) = Fg1.TextMatrix(xFil, xCol)
            Next xCol
            xFila = xFila + 1
        Next xFil
    
    Else
        '--verificar si hay registros para imprimir
        If Fg2.Rows <= Fg2.FixedRows Then
            MsgBox "No hay registros para imprimir", vbInformation, xTitulo
            Exit Sub
        End If
    
        '--imprimir el resumen
        RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 6) And ((con_formatostipodet.idformatotipo) = 1)) ORDER BY con_formatostipodet.orden", xCon
    
        ReDim xCampos(Fg2.Rows - 2, Fg2.Cols - 1)
        
        xFila = 0
        ' PASAMOS LOS DATOS DEL CONTROL Fg1 AL ARRAY DE DATOS
        For xFil = 1 To Fg2.Rows - 1
            For xCol = 1 To Fg2.Cols - 1
                xCampos(xFila, xCol) = Fg2.TextMatrix(xFil, xCol)
            Next xCol
            xFila = xFila + 1
        Next xFil
        
    End If
        
    ' ESTABLECEMOS EL TITULO DE CADA COLUMNA PARA EL REPORTE
    Rst.MoveFirst
    For A = 1 To Rst.RecordCount
        If xCampos(0, A) = NulosC(Rst("abrev")) Then
            If Rst("imprimir") = False Then
                xCampos(0, A) = ""
            End If
        End If
        Rst.MoveNext
        If Rst.EOF = True Then Exit For
    Next A
    
    
    Dim xFrm As New eps_librerias.Imprimir
        
    xFrm.Cabecera1 = NomEmp                         ' ESPECIFICA EL NOMBRE DE LA EMPRESA
    xFrm.Cabecera2 = "RUC Nº: " & NumRUC            ' ESPECIFICA EL NUMERO DE RUC DE LA EMPRESA
    xFrm.Fecha = Format(Date, "dd/mm/yyyy")         ' ESPECIFICA LA FECHA DE EMISION DEL REPORTE
    xFrm.Titulo1 = "CENTRO DE COSTO"  ' TITULO DEL REPORTE
    xFrm.Titulo2 = nPeriodo                         ' SEGUNDO TITULO DEL REPORTE
    xFrm.TamañoFuente = 6                           ' ESPECIFICA EL TAMAÑO DE LA FUENTE
    xFrm.TamañoCabecera = 8                         ' ESPECIFICA EL TAMAÑO DE LA FUENTE DE LA CABECERA
    xFrm.FuenteCabecera = "Courier New"             ' ESPECIFICA EL NOMBRE DE LA FUENTE DE LA CABECERA
    If TabOne1.CurrTab = 0 Then                     ' ESPECIFICA LA ORIENTACION DE LA JOHA
        xFrm.Posicion_Hoja = Horizontal '--detalle
    Else
        xFrm.Posicion_Hoja = Vertical   '--resumen
    End If
    xFrm.Tamaño_Hoja = A_4                          ' ESPECIFICA EL TAMAÑO DE LA HOJA
    xFrm.TextoConsiderar = "CENTRO COSTO"
    xFrm.TextoConsiderarAncho = 12
    xFrm.ImprimirArray xCampos(), Rst
        
    Set Rst = Nothing
    Set xFrm = Nothing
    
End Sub
