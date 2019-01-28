VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRegDetraccion 
   Caption         =   "Contabilidad - Reporte de Detracción"
   ClientHeight    =   8790
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   12900
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6000
      Left            =   30
      TabIndex        =   43
      Top             =   1650
      Width           =   11850
      _cx             =   20902
      _cy             =   10583
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
      Caption         =   "      Detalle    |    Resumen   "
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
         Height          =   5580
         Left            =   45
         TabIndex        =   44
         Top             =   45
         Width           =   11760
         _cx             =   20743
         _cy             =   9842
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
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   19
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRegDetraccion.frx":0000
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
         Height          =   5580
         Left            =   12495
         TabIndex        =   45
         Top             =   45
         Width           =   11760
         _cx             =   20743
         _cy             =   9842
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
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRegDetraccion.frx":0238
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
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1245
      Left            =   3330
      TabIndex        =   36
      Top             =   3750
      Visible         =   0   'False
      Width           =   5010
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   120
         TabIndex        =   37
         Top             =   630
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5010
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   4995
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   4995
         X2              =   4995
         Y1              =   30
         Y2              =   1230
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exportando a Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   38
         Top             =   105
         Width           =   1665
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   315
         Left            =   45
         Top             =   45
         Width           =   4935
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8790
      Top             =   -150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":03D8
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":091C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":0CAE
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":0E32
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":1286
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":139E
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":18E2
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":1E26
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":1F3A
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":204E
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":24A2
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":260E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":2B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":2E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegDetraccion.frx":3202
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
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
            Object.ToolTipText     =   "Configurar Formatos"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne2 
      Height          =   1275
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   11880
      _cx             =   20955
      _cy             =   2249
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
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "Inicio|Mas"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
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
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   12825
         TabIndex        =   8
         Top             =   45
         Width           =   11490
         Begin VB.Frame Frame14 
            Caption         =   "[  Filtro por Proveedor ]"
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
            Left            =   120
            TabIndex        =   14
            Top             =   0
            Width           =   9390
            Begin VB.OptionButton OptSel1 
               Caption         =   "Todos"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   17
               Top             =   270
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.OptionButton OptSel2 
               Caption         =   "Seleccionar"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1170
               TabIndex        =   16
               Top             =   270
               Width           =   1140
            End
            Begin VB.CommandButton CmdBusCliPro 
               Enabled         =   0   'False
               Height          =   240
               Left            =   8640
               Picture         =   "FrmRegDetraccion.frx":3594
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   210
               Width           =   210
            End
            Begin VB.TextBox TxtCliPro 
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   3405
               Locked          =   -1  'True
               TabIndex        =   18
               Text            =   "TxtCliPro"
               Top             =   180
               Width           =   5475
            End
            Begin VB.Label LblIdCliPro 
               Caption         =   "LblIdCliPro"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   2280
               TabIndex        =   20
               Top             =   150
               Visible         =   0   'False
               Width           =   750
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Proveedor"
               Height          =   195
               Index           =   2
               Left            =   2610
               TabIndex        =   19
               Top             =   270
               Width           =   735
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "[ Tipo de Documento ]"
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
            Height          =   600
            Left            =   120
            TabIndex        =   9
            Top             =   570
            Width           =   5085
            Begin VB.CommandButton CmdBusTipDoc 
               Height          =   240
               Left            =   735
               Picture         =   "FrmRegDetraccion.frx":36C6
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   270
               Width           =   240
            End
            Begin VB.TextBox TxtTipDoc 
               Height          =   300
               Left            =   90
               MaxLength       =   3
               TabIndex        =   11
               Text            =   "TxtTipDoc"
               Top             =   240
               Width           =   915
            End
            Begin VB.Label LblNomDoc 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblNomDoc"
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
               Left            =   1035
               TabIndex        =   13
               Top             =   240
               Width           =   3975
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "T.D."
               Height          =   195
               Index           =   1
               Left            =   2340
               TabIndex        =   12
               Top             =   330
               Visible         =   0   'False
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1185
         Left            =   345
         TabIndex        =   2
         Top             =   45
         Width           =   11490
         Begin VB.Frame Frame1 
            Caption         =   "[ Búsqueda Por ]"
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
            Height          =   570
            Left            =   4170
            TabIndex        =   40
            Top             =   0
            Width           =   2520
            Begin VB.OptionButton OptFch1 
               Caption         =   "Fch. Detr"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Value           =   -1  'True
               Width           =   1110
            End
            Begin VB.OptionButton OptFch2 
               Caption         =   "Fch. Doc"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   1290
               TabIndex        =   41
               Top             =   240
               Width           =   1125
            End
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Solo Registrados"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   9840
            TabIndex        =   39
            Top             =   930
            Value           =   1  'Checked
            Width           =   1515
         End
         Begin VB.Frame Frame2 
            Caption         =   "[ Datos ]"
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
            Height          =   930
            Left            =   9780
            TabIndex        =   33
            Top             =   0
            Width           =   1695
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nº Registros :"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   975
            End
            Begin VB.Label LblNumreg 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   360
               Left            =   120
               TabIndex        =   34
               Top             =   480
               Width           =   1440
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "[ Ordenado Por ]"
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
            Height          =   1170
            Left            =   6765
            TabIndex        =   28
            Top             =   0
            Width           =   2940
            Begin VB.OptionButton OptSort2 
               Caption         =   "Nº de Documento"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   32
               Top             =   470
               Width           =   1800
            End
            Begin VB.OptionButton OptSort1 
               Caption         =   "Fecha  de Emisión"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   31
               Top             =   240
               Width           =   1800
            End
            Begin VB.OptionButton OptSort3 
               Caption         =   "Nº Registro"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   30
               Top             =   700
               Width           =   1650
            End
            Begin VB.OptionButton OptSort4 
               Caption         =   "Fch. Emisión y Nº de Documento"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   29
               Top             =   930
               Width           =   2670
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "[ Opciones de Vista ]"
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
            Height          =   570
            Left            =   4170
            TabIndex        =   25
            Top             =   570
            Width           =   2520
            Begin VB.OptionButton OptOpc33 
               Caption         =   "Ventas"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   1290
               TabIndex        =   27
               Top             =   240
               Width           =   1125
            End
            Begin VB.OptionButton OptOpc11 
               Caption         =   "Compras"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Value           =   -1  'True
               Width           =   1080
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "[ Expresado en ]"
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
            Height          =   600
            Left            =   30
            TabIndex        =   4
            Top             =   570
            Width           =   4095
            Begin VB.CommandButton CmdBusMon 
               Height          =   230
               Left            =   495
               Picture         =   "FrmRegDetraccion.frx":37F8
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   180
               MaxLength       =   1
               TabIndex        =   6
               Text            =   "TxtIdMon"
               Top             =   240
               Width           =   555
            End
            Begin VB.Label LblMoneda 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblMoneda"
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
               Left            =   735
               TabIndex        =   7
               Top             =   240
               Width           =   3135
            End
         End
         Begin VB.Frame Frame5 
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
            Height          =   585
            Left            =   30
            TabIndex        =   3
            Top             =   0
            Width           =   4095
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   735
               TabIndex        =   21
               Top             =   210
               Width           =   1305
               _ExtentX        =   2302
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
               Valor           =   "11/09/2008"
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
               Height          =   300
               Left            =   2700
               TabIndex        =   22
               Top             =   210
               Width           =   1305
               _ExtentX        =   2302
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
               Valor           =   "11/09/2008"
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   24
               Top             =   270
               Width           =   465
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
               Height          =   195
               Index           =   2
               Left            =   2145
               TabIndex        =   23
               Top             =   255
               Width           =   420
            End
         End
      End
   End
End
Attribute VB_Name = "FrmRegDetraccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FrmRegComVen.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : EMITE EL LIBRO REGISTRO DE COMPRAS, EN FUNCION A CRITERIOS ESPECIFICADOS POR EL
'                     USUARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 21/10/09
'* VERSION          : 1.0

'Modificado         : 08/02/10 - Johan Castro
'                     Agregar filtro por proveedor y Tipo de Documento
'                     26/04/10 - Johan Castro
'                     Agregar pestaña que muestre resumen por documento
'                     Agregar para imprimir y exportar a MSExcel el resumen por documento
'                     21/07/10 - Johan Castro
'                     Hacer que la consulta este sincronizado con la tabla con_formatostipodet
'                     en relacion al orden de los campos
'*****************************************************************************************************
Option Explicit

Dim SeEjecuto As Boolean               ' ESPECIFICA SI SE EJECUTO EL EVENTO ACTIVATE
Dim xNumPag As Integer                 ' ALMACENA EL NUMERO DE PAGINA
Dim xTotal1, xTotal2, xTotal3, xTotal4, xTotal5, xTotal6 As Double
Dim xCadOrd As String                  ' Cadenas de ordenacion para las consultas
Dim xFormatoActual As Integer          ' indica el id del formato actual que se mostrara en la cuadricula

'*****************************************************************************************************
'* Nombre           : MostrarDetalle
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GENERA EL REGISTRO DE COMPRAS EN FUNCION A LAS CONDICIONES APLICADAS POR EL
'*                    USUARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MostrarDetalle()

    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLCampos As String '--Relacion de campos a mostrar, obtenido de tabla: con_formatostipodet
    Dim nSQLSub As String '--Sentencia SQL para identificar una subconsulta; está a nivel de detalle

    
    '--obtener el orden de presentacion de los campos
    nSQLCampos = fSetearCuadriculaColumna(xCon, 9)
    '--verificar si hay campos seleccionados para mostrar el reporte
    If nSQLCampos = "" Then Exit Sub
        
    Me.MousePointer = vbHourglass
    DoEvents
    '--
        
    nSQLSub = GenerarConsulta()
    
    '--armar la sentencia SQL
    nSQL = "Select " & nSQLCampos & _
            vbCr + " from ( " _
            + vbCr + nSQLSub _
            + vbCr + ") as consulta "
    
    '--ejecutar la consulta
    RST_Busq Rst, nSQL, xCon
    
    '--Salir si hay error en la consulta
    If Rst.State = 0 Then GoTo LaCague
    
   '**************************************************************************************************
    '--obtener las posiciones de las columnas
    Dim mColCampo As Integer
    Dim mCol As Integer '--indica la posicion del campo
    '--definir el array por defecto a 15 campos
    Dim ArrCampos(15) As Integer
    '--posicionar la variable a la primera columna
    mCol = 0
    '--obtener la posicion de los campos de la consulta en el arreglo
    For mColCampo = 0 To Rst.Fields.Count - 1
        Select Case LCase(Rst.Fields(mColCampo).Name)
            Case "impdetrac":     ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impbaseexmn":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "imptotexmn":    ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            
        End Select
        
    Next mColCampo
    '**************************************************************************************************
    
    '--Aplicar orden
    If OptSort1.Value = True Then Rst.Sort = "fchdoc"
    If OptSort2.Value = True Then Rst.Sort = "numerodoc"
    If OptSort3.Value = True Then Rst.Sort = "registro"
    If OptSort4.Value = True Then Rst.Sort = "fchdoc,numerodoc"
        
    LblNumreg.Caption = Rst.RecordCount

    Do While Not Rst.EOF
        DoEvents
''        ProgressBar1.Value = Rst.Bookmark
        
        '-----------------------------------------------
        Fg1.Rows = Fg1.Rows + 1
        
        For mCol = 0 To Rst.Fields.Count - 1
        
            Select Case LCase(Rst.Fields(mCol).Name)
                Case "xitem"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Fg1.Rows - Fg1.FixedRows
                Case "fchdoc", "fchpag"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_DATE)
                
                Case "impdetrac", "impbaseexmn", "imptotexmn", "impbaseexmn", "imptotexmn"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_MONTO)
                
                Case "tipcam"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), "0.000")
                Case Else
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
            End Select
            
        Next mCol
                
        '--verificar si monto=cero y no sea anulado =>> pintar la fila para que muestre una alerta al usuario
        If InStr(LCase(Rst("numdet")), "numero") <> 0 Then
            GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HC0C0FF
        End If
            
        Rst.MoveNext
    Loop
    
    '**************************************************************************************************
    '--verificamos si se suman las columnas
    If ArrCampos(0) <> 0 Then
            
        '--sumando las columnas
        Fg1.Rows = Fg1.Rows + 1
        FORMATO_CELDA Fg1, Fg1.Rows - 1, IIf(ArrCampos(0) - 2 < 0, 1, ArrCampos(0) - 2), &H800000, False, , "TOTAL ==>"
        
        For mCol = 0 To UBound(ArrCampos())
            If ArrCampos(mCol) <> 0 Then
                FORMATO_CELDA Fg1, Fg1.Rows - 1, ArrCampos(mCol) + 1, &H800000, False, , Format(GRID_SUMAR_COL(Fg1, ArrCampos(mCol) + 1), FORMAT_MONTO)
            End If
        Next mCol
        
    End If
    '**************************************************************************************************
    
LaCague:

    Set Rst = Nothing
    
    '--restablecer cursor
    Me.MousePointer = vbDefault
    
End Sub




Private Sub fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 Then
        Fg1.SelectionMode = flexSelectionFree
        Fg1.Editable = flexEDKbdMouse
    End If

    If KeyCode = 122 Then
        Fg1.SelectionMode = flexSelectionByRow
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
    
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
        
        TxtCliPro.Text = ""
        
        TabOne2.CurrTab = 0
        
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
    
   
        OptSort3.Value = True
        SeEjecuto = True
        TxtFchIni.SetFocus
        
        '--enfocar en la pestaña del detalle
        TabOne1.CurrTab = 0
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    LblNumreg.Caption = 0
    Dim xRs As New ADODB.Recordset
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
    ' SELECCIONAMOS EL FORMATO ACTUAL PARA LA IMPRESION DEL LIBRO "REGISTRO DE COMPRAS"
    RST_Busq xRs, "SELECT con_formatostipo.* From con_formatostipo WHERE (((con_formatostipo.defecto)=-1) AND ((con_formatostipo.idformato)=9))", xCon
    If xRs.RecordCount = 0 Then
        MsgBox "No hay formato en la base de datos para mostrar el reporte" & vbCr & "Consulte con el administrador del sistema", vbInformation, xTitulo
        Exit Sub
    End If
    xFormatoActual = xRs("id")
    
    Set xRs = Nothing
    
    '--dar formato al detalle
    SetearCuadricula Fg1, 9, xCon, 1, xFormatoActual, False
        
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    OptOpc11.Value = True
    
    OptSort3.Value = True
    
    '--cargar el formato del resumen
    SetearCuadricula fg2, 9, xCon, 1, 3, False
    
    '--buscar los registros
    Fg1.AutoSearch = flexSearchFromTop
    
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height > 3000 Then
        TabOne1.Top = 1650
        TabOne1.Width = Me.Width - 150
        TabOne1.Height = Me.Height - 2050
    End If
End Sub

Private Sub OptOpc11_Click()
    Frame14.Caption = "[  Filtro por Proveedor ]"
    Label1(2).Caption = "Proveedor"
End Sub

Private Sub OptOpc33_Click()
    Frame14.Caption = "[  Filtro por Cliente ]"
    Label1(2).Caption = "Cliente"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        
        '--limpiar datos
        Fg1.Rows = Fg1.FixedRows
        fg2.Rows = fg2.FixedRows
        LblNumreg.Caption = 0
        DoEvents
    
        '--posicionar en la primera pestaña
        TabOne2.CurrTab = 0
        DoEvents
        '--
        ' VERIFICAMOS QUE LOS DATOS NECESARIOS SEAN LOS CORRECTOS
        If NulosC(TxtFchIni.Valor) = "" Then
            MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If
        
        If NulosC(TxtFchFin.Valor) = "" Then
            MsgBox "No ha especificado la fecha de final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchFin.SetFocus
            Exit Sub
        End If
    
        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            TxtFchIni.SetFocus
            Exit Sub
        End If
        
        '--VERIFICAMOS LA MONEDA
        If NulosN(TxtIdMon.Text) = 0 Then
            MsgBox "Falta especificar la moneda", vbInformation, xTitulo
            TxtIdMon.SetFocus
            Exit Sub
        End If
        
       
        MostrarDetalle
        
        MostrarResumen
    End If
    
    If Button.Index = 3 Then
        If Fg1.Rows = 2 Then
            MsgBox "No hay registro que exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Dim xFun As New SGI2_funciones.formularios
        
        If TabOne1.CurrTab = 0 Then     '--imprimir el detalle
            xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "REGISTRO DETRACCIONES - " & IIf(OptOpc11.Value = True, "COMPRAS", "VENTAS"), "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Por Fecha", "Registro de Detracciones"   ', Rst, ""
            
        Else                            '--imprimir el resumen
            xFun.VSFlexGrid_Exportar_MSExcel xCon, fg2, "RESUMEN - REGISTRO DETRACCIONES DE " & IIf(OptOpc11.Value = True, "COMPRAS", "VENTAS"), "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Por Fecha", "Registro de Detracciones"   ', Rst, ""
        End If
        
        
        Set xFun = Nothing
    End If
    
    If Button.Index = 4 Then IMPRIMIR
    
    If Button.Index = 5 Then
        Configurar
    End If
    
    If Button.Index = 7 Then
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : IMPRIMIR
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME EL REGISTRO DE COMPRAS PARA ELLO INVOCA AL EVENTO ImprimirArray DE LA
'*                    CLASE eps_librerias.IMPRIMIR
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub IMPRIMIR()
    If Fg1.Rows = 1 Then
        MsgBox "No hay registros para imprimir", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim nPeriodo   As String
    Dim xMoneda As String
        
    If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
        nPeriodo = " Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
    Else
        nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
    End If
    
    xMoneda = LblMoneda.Caption
    
    Dim RstTmp As New ADODB.Recordset
    Dim A As Long
    Dim Rst As New ADODB.Recordset
    
    Dim xCampos() As String
    Dim xFil, xCol As Double
    Dim xFila As Double
    
    ' SELECCIONAMOS EL FORMATO ACTUAL PARA LA IMPRESION DEL LIBRO "REGISTRO DE COMPRAS"
    
    If TabOne1.CurrTab = 0 Then
        '--verificar si hay registros para imprimir
        If Fg1.Rows <= Fg1.FixedRows Then
            MsgBox "No hay registros para imprimir", vbInformation, xTitulo
            Exit Sub
        End If
        
        RST_Busq Rst, "SELECT con_formatostipo.rpttitulo, con_formatostipo.rpttamcab, con_formatostipo.rpttamdet, con_formatostipodet.* " _
            & " FROM con_formatostipodet INNER JOIN con_formatostipo ON (con_formatostipodet.idformatotipo = con_formatostipo.id) AND (con_formatostipodet.idformato = con_formatostipo.idformato) " _
            & " Where (((con_formatostipodet.idformato) = 9) And ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ") And ((con_formatostipodet.mostrar) = -1)) " _
            & " ORDER BY con_formatostipodet.orden", xCon
                
        ReDim xCampos(Fg1.Rows - 2, Fg1.Cols - 1)
        
        xFila = 0
        '--asignando el nombre de los campos
        For xFil = 1 To Fg1.Rows - 1
            For xCol = 1 To Fg1.Cols - 1
                xCampos(xFila, xCol) = Fg1.TextMatrix(xFil, xCol)
            Next xCol
            xFila = xFila + 1
        Next xFil
    Else
        '--verificar si hay registros para imprimir
        If fg2.Rows <= fg2.FixedRows Then
            MsgBox "No hay registros para imprimir", vbInformation, xTitulo
            Exit Sub
        End If
        
        RST_Busq Rst, "SELECT con_formatostipo.rpttitulo, con_formatostipo.rpttamcab, con_formatostipo.rpttamdet, con_formatostipodet.* " _
            & " FROM con_formatostipodet INNER JOIN con_formatostipo ON (con_formatostipodet.idformatotipo = con_formatostipo.id) AND (con_formatostipodet.idformato = con_formatostipo.idformato) " _
            & " Where (((con_formatostipodet.idformato) = 9) And ((con_formatostipodet.idformatotipo) = 3 ) And ((con_formatostipodet.mostrar) = -1)) " _
            & " ORDER BY con_formatostipodet.orden", xCon
    
        ReDim xCampos(fg2.Rows - 2, fg2.Cols - 1)
        
        xFila = 0
        '--asignando el nombre de los campos
        For xFil = 1 To fg2.Rows - 1
            For xCol = 1 To fg2.Cols - 1
                xCampos(xFila, xCol) = fg2.TextMatrix(xFil, xCol)
            Next xCol
            xFila = xFila + 1
        Next xFil
        
    End If
    
    ' BLANQUEAMOS LAS TITULOS DE LAS COLUMNAS QUE NO SE VAN A IMPRIMIR
    Rst.MoveFirst
    For A = 1 To Rst.RecordCount
        If xCampos(0, A) = Rst("abrev") Then
            If Rst("imprimir") = False Then
                xCampos(0, A) = ""
            End If
        End If
        Rst.MoveNext
        If Rst.EOF = True Then Exit For
    Next A
    
    Rst.MoveFirst
    
    Dim xfrm As New eps_librerias.IMPRIMIR
    ' CABECERA DEL REPORTE
    xfrm.Cabecera1 = NomEmp                                                   ' NOMBRE DE LA EMPRESA
    xfrm.Cabecera2 = "RUC Nº: " & NumRUC                                      ' NUMERO DE RUC DE LA EMPRESA
    xfrm.Fecha = Format(Date, "dd/mm/yyyy")                                   ' FECHA DE EMISION DEL REPORTE
    xfrm.Titulo1 = UCase(NulosC(Rst("rpttitulo"))) & IIf(OptOpc11.Value = True, " - COMPRAS ", " - VENTAS ") & "(Expresado en " & xMoneda & ")"  ' TITULO DEL REPORTE
    xfrm.Titulo2 = nPeriodo                                                   ' SEGUNDO TITULO DEL REPORTE
    xfrm.TamañoFuente = NulosN(Rst("rpttamdet"))  '6                                                     ' TAMAÑO DE LA FUENTE DEL REPORTE
    xfrm.TamañoCabecera = NulosN(Rst("rpttamcab")) '8                                                    ' TAMAÑO DE LA FUENTE DE LA CABECERA DEL REPORTE
    xfrm.FuenteCabecera = "Courier New"                                       ' ESTABLECE LA FUENTE DE LA CABECERA
    
    If TabOne1.CurrTab = 0 Then                     ' ESPECIFICA LA ORIENTACION DE LA JOHA
        xfrm.Posicion_Hoja = Horizontal '--detalle
    Else
        xfrm.Posicion_Hoja = Vertical   '--resumen
    End If
    
    xfrm.Tamaño_Hoja = A_4                                                    ' ESTABLECE EL TAMAÑO DE LA HOJA
    xfrm.ImprimirArray xCampos, Rst
    Set xfrm = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Configurar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL REGISTRO DE COMPRAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Configurar()
    Dim xform As New SGI2_funciones.Varias
    If xform.CambioOpcionLiro(9, xCon, 1) = True Then
    
        Dim xRs As New ADODB.Recordset
        ' SELECCIONAMOS EL FORMATO ACTUAL PARA LA IMPRESION DEL LIBRO "REGISTRO DE COMPRAS"
        RST_Busq xRs, "SELECT con_formatostipo.* From con_formatostipo WHERE (((con_formatostipo.defecto)=-1) AND ((con_formatostipo.idformato)=9))", xCon
    
        xFormatoActual = xRs("id")
        
        Set xRs = Nothing
        
        SetearCuadricula Fg1, 9, xCon, 1, xFormatoActual, False
            
        If TxtFchIni.Valor = "" And TxtFchFin.Valor = "" Then
            MsgBox "No ha especificado el periodo de la consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If

        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If
        MostrarDetalle
    End If
    Set xform = Nothing
End Sub


'***********************************************************************************************
Private Sub CmdBusTipDoc_Click()
    ' EJECUTA LA BUSQUEDA DE UN TIPO DE DOCUMENTO
    If IsDate(TxtFchIni.Valor) = False Then
        MsgBox "Falta especificar la Fecha de Inicio", vbExclamation, xTitulo
        TabOne2.CurrTab = 0
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If IsDate(TxtFchFin.Valor) = False Then
        MsgBox "Falta especificar la Fecha Final", vbExclamation, xTitulo
        TabOne2.CurrTab = 0
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Abrev":    xCampos(1, 1) = "abrev":      xCampos(1, 2) = "450":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":         xCampos(2, 1) = "id":               xCampos(2, 2) = "600":         xCampos(2, 3) = "N"
    
    xform.SqlCad = "SELECT mae_documento.id, mae_documento.descripcion, mae_documento.abrev " _
    & " FROM com_compras INNER JOIN mae_documento ON com_compras.tipdoc = mae_documento.id " _
    & " WHERE (((com_compras.fchreg) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " _
    & " GROUP BY mae_documento.id, mae_documento.descripcion, mae_documento.abrev;"
    
    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipDoc.Text = xRs("id")
            LblNomDoc.Caption = NulosC(xRs("descripcion"))
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub




Private Sub TxtCliPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCliPro_Click
    End If
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc_Click
    End If
End Sub

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    Dim nSQL As String
    Dim Rst As New ADODB.Recordset
    
    nSQL = "SELECT mae_documento.id, mae_documento.descripcion, mae_documento.abrev " _
        & " FROM com_compras INNER JOIN mae_documento ON com_compras.tipdoc = mae_documento.id " _
        & " WHERE (((com_compras.fchreg) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) and com_compras.tipdoc = " & NulosN(TxtTipDoc.Text) & " " _
        & " GROUP BY mae_documento.id, mae_documento.descripcion, mae_documento.abrev;"

    RST_Busq Rst, nSQL, xCon
    
    If Rst.State = 1 Then
        If Rst.RecordCount <> 0 Then
            TxtTipDoc.Text = Rst("id")
            LblNomDoc.Caption = NulosC(Rst("descripcion"))
        Else
            TxtTipDoc.Text = ""
            LblNomDoc.Caption = ""
        End If
    End If
    Set Rst = Nothing
End Sub

'***********************************************************************************************





'***************************************************************************************************************************************

Private Sub CmdBusCliPro_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    If OptOpc11.Value = True Then
        xform.Titulo = "Buscando Proveedor"
        xform.SqlCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov where mae_prov.id <>0 ORDER BY mae_prov.nombre"
    Else
        xform.Titulo = "Buscando Cliente"
        xform.SqlCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente where mae_cliente.id<>0 ORDER BY mae_cliente.nombre"
        
    End If
    xCampos(0, 0) = "Razón Social":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":        xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"

    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCliPro.Text = NulosC(xRs("nombre"))
        LblIdCliPro.Caption = xRs("id")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub OptSel1_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    TxtCliPro.Enabled = False
    CmdBusCliPro.Enabled = False
End Sub
Private Sub OptSel2_Click()
    TxtCliPro.Enabled = True
    CmdBusCliPro.Enabled = True
    TxtCliPro.SetFocus
End Sub


Private Sub CmdBusMon_Click()
    
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":   xCampos(1, 1) = "id":              xCampos(1, 2) = "500":      xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_moneda ORDER BY descripcion ;", xCampos(), "Buscando Moneda", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.RecordCount = 0 Then GoTo SALIR
    TxtIdMon.Text = xRs("id") & ""
    LblMoneda.Caption = xRs("descripcion") & ""
    
SALIR:
    Set xRs = Nothing
End Sub

Private Sub TxtIdMon_Change()
    If Trim(TxtIdMon.Text) = "" Then LblMoneda.Caption = ""
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosC(TxtIdMon.Text) <> "" Then
        LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
        If NulosC(LblMoneda.Caption) = "" Then
            TxtIdMon.Text = ""
        End If
    End If
End Sub


'***************************************************************************************************************************************



'*****************************************************************************************************
'* Nombre           : MostrarResumen
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GENERA EL RESUMEN DEL REGISTRO DE COMPRAS EN FUNCION A LAS CONDICIONES APLICADAS POR EL
'*                    USUARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MostrarResumen()

    Dim Rst As New ADODB.Recordset

    Dim A As Long
    
    Dim nSQLSub As String '--Sentencia SQL para identificar una subconsulta; está a nivel de detalle
    Dim nSQLCampos As String '--Relacion de campos a mostrar, obtenido de tabla: con_formatostipodet
    Dim nSQL As String

    '--verificar si se puede mostrar los datos, esto dependera que esta la configuracion del grid en la base de datos
    If fg2.Cols = 1 Then
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    DoEvents
    '--
    '--obtener el orden de presentacion de los campos
    nSQLCampos = fSetearCuadriculaColumna(xCon, 9, 1, 3)
    
    '--Generar la sub consulta
    nSQLSub = GenerarConsulta()
    
    nSQL = "SELECT det.numruc, det.nombre, det.tipodetraccion, Sum(det.impdetrac) AS sumdetracc, Sum(det.impbaseexmn) AS sumbaseex, Sum(det.imptotexmn) AS sumtotex " _
            + vbCr + " FROM ( " _
            + vbCr + nSQLSub _
            + vbCr + " ) AS det GROUP BY det.numruc, det.nombre, det.tipodetraccion " _
            + vbCr + " ORDER BY det.nombre "
                
    '--armar la sentencia SQL
    nSQL = "Select " & nSQLCampos & _
            vbCr + " from ( " _
            + vbCr + nSQL _
            + vbCr + ") as consulta "
    
    '--ejecutar la consulta
    RST_Busq Rst, nSQL, xCon
    
    '--Salir si hay error en la consulta
    If Rst.State = 0 Then GoTo LaCague
    
   '**************************************************************************************************
    '--obtener las posiciones de las columnas
    Dim mColCampo As Integer
    Dim mCol As Integer '--indica la posicion del campo
    '--definir el array por defecto a 15 campos
    Dim ArrCampos(15) As Integer
    '--posicionar la variable a la primera columna
    mCol = 0
    '--obtener la posicion de los campos de la consulta en el arreglo
    For mColCampo = 0 To Rst.Fields.Count - 1
        Select Case LCase(Rst.Fields(mColCampo).Name)
            Case "sumdetracc":  ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "sumbaseex":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "sumtotex":    ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            
        End Select
        
    Next mColCampo
    '**************************************************************************************************

    Do While Not Rst.EOF
        DoEvents
''        ProgressBar1.Value = Rst.Bookmark
        
        '-----------------------------------------------
        fg2.Rows = fg2.Rows + 1
        
        For mCol = 0 To Rst.Fields.Count - 1
        
            Select Case LCase(Rst.Fields(mCol).Name)
                
                Case "sumdetracc", "sumbaseex", "sumtotex"
                    fg2.TextMatrix(fg2.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_MONTO)
                
                Case Else
                    fg2.TextMatrix(fg2.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
            End Select
            
        Next mCol
            
        Rst.MoveNext
    Loop
    
    '**************************************************************************************************
    '--verificamos si se suman las columnas
    If ArrCampos(0) <> 0 Then
            
        '--sumando las columnas
        fg2.Rows = fg2.Rows + 1
        FORMATO_CELDA fg2, fg2.Rows - 1, IIf(ArrCampos(1) - 2 < 0, 1, ArrCampos(1) - 2), &H800000, False, , "TOTAL ==>"
        
        For mCol = 0 To UBound(ArrCampos())
            If ArrCampos(mCol) <> 0 Then
                FORMATO_CELDA fg2, fg2.Rows - 1, ArrCampos(mCol) + 1, &H800000, False, , Format(GRID_SUMAR_COL(fg2, ArrCampos(mCol) + 1), FORMAT_MONTO)
            End If
        Next mCol
        
    End If
    '**************************************************************************************************


LaCague:
    Set Rst = Nothing
        
    '--restablecer cursor
    Me.MousePointer = vbDefault
    
End Sub


Function GenerarConsulta() As String
    '===================================================================================================
    'creado: 27/04/11 Por Johan Castro
    'Propósito: Generar la consulta a nivel de detalle
    '
    'Entradas:  Ninguno
    '
    'Resultados: Consulta segun parametros indicados
    '
    '===================================================================================================
    
    Dim nSQL As String
    Dim nSQLFiltro As String
    
    '--verificar si se mostraran solo los registros distintos a "sin numero" en comprobante de detraccion
    If chk.Value = 1 Then
        nSQLFiltro = " and con_detraccion.numdet<>'SIN NUMERO' "
    End If
    '--verificar si la consulta se filtra por fecha de detraccion
    If OptFch1.Value = True Then nSQLFiltro = nSQLFiltro & " and con_detraccion.fchpag between cdate('" & TxtFchIni.Valor & "') and cdate('" & TxtFchFin.Valor & "') "
    

    If OptOpc11.Value = True Then
        
        '--verificar si hay filtro por proveedor
        If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLFiltro = nSQLFiltro & " and com_compras.idpro = " & NulosN(LblIdCliPro.Caption) & " "
        
        '--verificar si hay filtro por documento
        If NulosN(TxtTipDoc.Text) <> 0 Then nSQLFiltro = nSQLFiltro & " and com_compras.tipdoc = " & NulosN(TxtTipDoc.Text) & " "
    
        '--verificar si la consulta se filtra por fecha de detraccion
        If OptFch2.Value = True Then nSQLFiltro = nSQLFiltro & " and com_compras.fchdoc between cdate('" & TxtFchIni.Valor & "') and cdate('" & TxtFchFin.Valor & "') "
    
        nSQL = "SELECT 0 as xitem,con_detraccion.*, Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS registro, mae_prov.numruc, mae_prov.nombre, mae_documento.abrev, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numerodoc, com_compras.fchdoc, " _
                + vbCr + " mae_detraccion.descripcion AS tipodetraccion, mae_detraccion.tasa,mae_moneda.simbolo, IIf([com_compras].[tc]=0,[con_tc].[impven],[com_compras].[tc]) AS tipcam, " _
                + vbCr + " (com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina) AS impbase, com_compras.imptot, "
                
        If NulosN(TxtIdMon.Text) = 1 Then
            nSQL = nSQL _
                + vbCr + " con_detraccion.imp AS impdetrac, " _
                + vbCr + " CDbl(Format(IIf([com_compras].[idmon]=1,(com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina),(com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina) * [tipcam] ),'0.00')) AS impbaseexmn, " _
                + vbCr + " CDbl(Format(IIf([com_compras].[idmon]=1,[com_compras].[imptot],[com_compras].[imptot] * [tipcam]),'0.00')) AS imptotexmn "
        ElseIf NulosN(TxtIdMon.Text) = 2 Then
            nSQL = nSQL _
                + vbCr + " CDbl(Format(IIf([tipcam]=0,0,con_detraccion.imp/[tipcam]),'0.00')) AS impdetrac, " _
                + vbCr + " CDbl(Format(IIf([com_compras].[idmon]=2,(com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina),IIf([tipcam]=0,0,(com_compras.impbru+com_compras.impbru2+com_compras.impbru3+com_compras.impina) / [tipcam] )),'0.00')) AS impbaseexmn, " _
                + vbCr + " CDbl(Format(IIf([com_compras].[idmon]=2,[com_compras].[imptot],IIf([tipcam]=0,0,[com_compras].[imptot] / [tipcam])),'0.00')) AS imptotexmn "
        End If
        
        nSQL = nSQL _
                + vbCr + " FROM mae_prov RIGHT JOIN (mae_detraccion RIGHT JOIN (mae_documento RIGHT JOIN (((com_compras INNER JOIN (con_detraccion LEFT JOIN mae_moneda ON con_detraccion.idmon = mae_moneda.id) ON com_compras.id = con_detraccion.iddoc) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) ON mae_detraccion.id = con_detraccion.iddet) ON mae_prov.id = com_compras.idpro " _
                + vbCr + " Where (((con_detraccion.Tipo) = 1)) " & nSQLFiltro _
                + vbCr + " ORDER BY com_compras.fchdoc DESC "

    Else
    
        '--verificar si hay filtro por proveedor
        If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLFiltro = nSQLFiltro & " and vta_ventas.idcli = " & NulosN(LblIdCliPro.Caption) & " "
        
        '--verificar si hay filtro por documento
        If NulosN(TxtTipDoc.Text) <> 0 Then nSQLFiltro = nSQLFiltro & " and vta_ventas.tipdoc = " & NulosN(TxtTipDoc.Text) & " "
    
        '--verificar si la consulta se filtra por fecha de detraccion
        If OptFch2.Value = True Then nSQLFiltro = nSQLFiltro & " and vta_ventas.fchdoc between cdate('" & TxtFchIni.Valor & "') and cdate('" & TxtFchFin.Valor & "') "
       
        nSQL = "SELECT 0 as xitem,con_detraccion.*, Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) AS registro, mae_prov.numruc, mae_prov.nombre, mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numerodoc, vta_ventas.fchdoc, " _
                + vbCr + " mae_detraccion.descripcion AS tipodetraccion, mae_detraccion.tasa,mae_moneda.simbolo, IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc]) AS tipcam, " _
                + vbCr + " (vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf) AS impbase, vta_ventas.imptotdoc as imptot, "
                
        If NulosN(TxtIdMon.Text) = 1 Then
            nSQL = nSQL _
                + vbCr + " con_detraccion.imp AS impdetrac, " _
                + vbCr + " CDbl(Format(IIf([vta_ventas].[idmon]=1,(vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf),(vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf) * [tipcam] ),'0.00')) AS impbaseexmn, " _
                + vbCr + " CDbl(Format(IIf([vta_ventas].[idmon]=1,[vta_ventas].[imptotdoc],[vta_ventas].[imptotdoc] * [tipcam]),'0.00')) AS imptotexmn "
        ElseIf NulosN(TxtIdMon.Text) = 2 Then
            nSQL = nSQL _
                + vbCr + " CDbl(Format(IIf([tipcam]=0,0,con_detraccion.imp/[tipcam]),'0.00')) AS impdetrac, " _
                + vbCr + " CDbl(Format(IIf([vta_ventas].[idmon]=2,(vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf),IIf([tipcam]=0,0,(vta_ventas.impbru+vta_ventas.impbru2+vta_ventas.impbru3+vta_ventas.impinaf) / [tipcam] )),'0.00')) AS impbaseexmn, " _
                + vbCr + " CDbl(Format(IIf([vta_ventas].[idmon]=2,[vta_ventas].[imptotdoc],IIf([tipcam]=0,0,[vta_ventas].[imptotdoc] / [tipcam])),'0.00')) AS imptotexmn "
        End If
        
        nSQL = nSQL _
                + vbCr + " FROM mae_prov RIGHT JOIN (mae_detraccion RIGHT JOIN (mae_documento RIGHT JOIN (((vta_ventas INNER JOIN (con_detraccion LEFT JOIN mae_moneda ON con_detraccion.idmon = mae_moneda.id) ON vta_ventas.id = con_detraccion.iddoc) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_documento.id = vta_ventas.tipdoc) ON mae_detraccion.id = con_detraccion.iddet) ON mae_prov.id = vta_ventas.idcli " _
                + vbCr + " Where (((con_detraccion.Tipo) = 2)) " & nSQLFiltro _
                + vbCr + " ORDER BY vta_ventas.fchdoc DESC "

    
        
    End If
    
    GenerarConsulta = nSQL
    
End Function
