VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRegVentasOtros1 
   Caption         =   "Contabilidad - Registro de Ventas"
   ClientHeight    =   9045
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6000
      Left            =   30
      TabIndex        =   42
      Top             =   1620
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
         TabIndex        =   43
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
         FormatString    =   $"FrmRegVentasOtros1.frx":0000
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRegVentasOtros1.frx":0238
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
   Begin SizerOneLibCtl.TabOne TabOne2 
      Height          =   1275
      Left            =   30
      TabIndex        =   4
      Top             =   360
      Width           =   11820
      _cx             =   20849
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
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1185
         Left            =   345
         TabIndex        =   6
         Top             =   45
         Width           =   11430
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
            Height          =   1170
            Left            =   9720
            TabIndex        =   39
            Top             =   0
            Width           =   1695
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
               TabIndex        =   41
               Top             =   630
               Width           =   1440
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nº Registros :"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   40
               Top             =   390
               Width           =   975
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
            Height          =   1170
            Left            =   4170
            TabIndex        =   33
            Top             =   0
            Width           =   2520
            Begin VB.OptionButton OptOpc11 
               Caption         =   "Todas las compras"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   240
               Width           =   1770
            End
            Begin VB.OptionButton OptOpc22 
               Caption         =   "Bancarización"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   37
               Top             =   475
               Width           =   1350
            End
            Begin VB.OptionButton OptOpc33 
               Caption         =   "Detracción"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   36
               Top             =   710
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.OptionButton OptOpc44 
               Caption         =   "Percepción"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   35
               Top             =   945
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txtBancarizacion 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1560
               TabIndex        =   34
               Text            =   "txtBancarizacion"
               Top             =   420
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "[Seleccionar Fecha]"
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
            TabIndex        =   16
            Top             =   0
            Width           =   4095
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   705
               TabIndex        =   17
               Top             =   225
               Width           =   1245
               _ExtentX        =   2196
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
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
               Height          =   300
               Left            =   2670
               TabIndex        =   18
               Top             =   240
               Width           =   1245
               _ExtentX        =   2196
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
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
               Height          =   195
               Index           =   2
               Left            =   2145
               TabIndex        =   20
               Top             =   330
               Width           =   420
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
               Height          =   195
               Index           =   1
               Left            =   105
               TabIndex        =   19
               Top             =   330
               Width           =   465
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
            Left            =   6750
            TabIndex        =   11
            Top             =   0
            Width           =   2910
            Begin VB.OptionButton OptSort2 
               Caption         =   "Nº de Documento"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   15
               Top             =   450
               Width           =   2010
            End
            Begin VB.OptionButton OptSort1 
               Caption         =   "Fecha  de Emisión"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   14
               Top             =   240
               Width           =   2010
            End
            Begin VB.OptionButton OptSort3 
               Caption         =   "Nº Registro"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   13
               Top             =   675
               Value           =   -1  'True
               Width           =   2010
            End
            Begin VB.OptionButton OptSort4 
               Caption         =   "Fch. Emisión y Nº de Documento"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   12
               Top             =   900
               Width           =   2730
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
            TabIndex        =   7
            Top             =   570
            Width           =   4095
            Begin VB.CommandButton CmdBusMon 
               Height          =   230
               Left            =   495
               Picture         =   "FrmRegVentasOtros1.frx":03D8
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   180
               MaxLength       =   1
               TabIndex        =   9
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
               TabIndex        =   10
               Top             =   240
               Width           =   3135
            End
         End
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   12765
         TabIndex        =   5
         Top             =   45
         Width           =   11430
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
            TabIndex        =   28
            Top             =   570
            Width           =   5085
            Begin VB.CommandButton CmdBusTipDoc 
               Height          =   240
               Left            =   735
               Picture         =   "FrmRegVentasOtros1.frx":050A
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   270
               Width           =   240
            End
            Begin VB.TextBox TxtTipDoc 
               Height          =   300
               Left            =   90
               MaxLength       =   3
               TabIndex        =   30
               Text            =   "TxtTipDoc"
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "T.D."
               Height          =   195
               Index           =   1
               Left            =   2340
               TabIndex        =   32
               Top             =   330
               Visible         =   0   'False
               Width           =   315
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
               TabIndex        =   31
               Top             =   240
               Width           =   3975
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[  Filtro por Cliente]"
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
            TabIndex        =   21
            Top             =   0
            Width           =   9390
            Begin VB.CommandButton CmdBusCliPro 
               Enabled         =   0   'False
               Height          =   240
               Left            =   8640
               Picture         =   "FrmRegVentasOtros1.frx":063C
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   210
               Width           =   210
            End
            Begin VB.OptionButton OptSel2 
               Caption         =   "Seleccionar"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1170
               TabIndex        =   23
               Top             =   270
               Width           =   1140
            End
            Begin VB.OptionButton OptSel1 
               Caption         =   "Todos"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   22
               Top             =   270
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.TextBox TxtCliPro 
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   3405
               Locked          =   -1  'True
               TabIndex        =   25
               Text            =   "TxtCliPro"
               Top             =   180
               Width           =   5475
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cliente"
               Height          =   195
               Index           =   0
               Left            =   2820
               TabIndex        =   27
               Top             =   270
               Width           =   480
            End
            Begin VB.Label LblIdCliPro 
               Caption         =   "LblIdCliPro"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   2280
               TabIndex        =   26
               Top             =   150
               Visible         =   0   'False
               Width           =   750
            End
         End
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1245
      Left            =   120
      TabIndex        =   1
      Top             =   7830
      Visible         =   0   'False
      Width           =   5010
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   615
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   609
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
         TabIndex        =   3
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
      Left            =   5460
      Top             =   45
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
            Picture         =   "FrmRegVentasOtros1.frx":076E
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":0CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":1044
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":11C8
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":161C
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":1734
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":1C78
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":21BC
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":22D0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":23E4
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":2838
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":29A4
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":2EEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":3206
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegVentasOtros1.frx":3598
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
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   1005
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
End
Attribute VB_Name = "FrmRegVentasOtros1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FrmRegVentasOtros.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : IMPRIME EL REGISTRO DE VENTAS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 22/10/09
'* VERSION          : 1.0

'Modificado         : 08/02/10 - Johan Castro
'                     Agregar filtro por cliente y Tipo de Documento
'                     26/04/10 - Johan Castro
'                     Agregar pestaña que muestre resumen por documento
'                     Agregar para imprimir y exportar a MSExcel

'*****************************************************************************************************
Option Explicit

Dim SeEjecuto As Boolean                 ' ESPECIFICA SI SE EJECUTO EL EVENTO ACTIVATE
Dim xNumPag As Integer                   ' ALMACENA EL NUMERO DE PAGINA
Dim xTotal1, xTotal2, xTotal3, xTotal4, xTotal5, xTotal6 As Double
Dim xIdPer As Integer                    ' ESPECIFICA EL ID DEL PERIODO
Dim xCadOrd As String                    ' Cadenas de ordenacion para las consultas
Dim xnomtabla As String                  ' ESPECIFICA EL NOMBRE DE LA TABLA
Dim xMonBancSol, xMonBancDol As Double   ' VARIABLE PARA ALMACENAR LOS IMPORTE MAXIMOS MINIMOS PARA LA BANCARIZACION, EN SOLES Y DOLARES
Dim xFormatoActual As Integer            ' ID DEL FORMATO ACTUAL O POR DEFECTO

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

    xMoneda = LblMoneda.Caption

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
            & " Where (((con_formatostipodet.idformato) = 3) And ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ") " _
            & " And ((con_formatostipodet.mostrar) = -1)) ORDER BY con_formatostipodet.orden", xCon
        
        'RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 3) And ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ")) " _
            & " ORDER BY con_formatostipodet.orden", xCon
    
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
        If fg2.Rows <= fg2.FixedRows Then
            MsgBox "No hay registros para imprimir", vbInformation, xTitulo
            Exit Sub
        End If
    
        '--imprimir el resumen
        RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 3) And ((con_formatostipodet.idformatotipo) = 3)) ORDER BY con_formatostipodet.orden", xCon
    
        ReDim xCampos(fg2.Rows - 2, fg2.Cols - 1)
        
        xFila = 0
        ' PASAMOS LOS DATOS DEL CONTROL Fg1 AL ARRAY DE DATOS
        For xFil = 1 To fg2.Rows - 1
            For xCol = 1 To fg2.Cols - 1
                xCampos(xFila, xCol) = fg2.TextMatrix(xFil, xCol)
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
    
    
    Dim xfrm As New eps_librerias.IMPRIMIR
    
    xfrm.Cabecera1 = NomEmp                         ' ESPECIFICA EL NOMBRE DE LA EMPRESA
    xfrm.Cabecera2 = "RUC Nº: " & NumRUC            ' ESPECIFICA EL NUMERO DE RUC DE LA EMPRESA
    xfrm.Fecha = Format(Date, "dd/mm/yyyy")         ' ESPECIFICA LA FECHA DE EMISION DEL REPORTE
    xfrm.Titulo1 = "REGISTRO DE VENTAS " & "(Expresado en " & xMoneda & ")"  ' TITULO DEL REPORTE
    xfrm.Titulo2 = nPeriodo                         ' SEGUNDO TITULO DEL REPORTE
    xfrm.TamañoFuente = 6                           ' ESPECIFICA EL TAMAÑO DE LA FUENTE
    xfrm.TamañoCabecera = 8                         ' ESPECIFICA EL TAMAÑO DE LA FUENTE DE LA CABECERA
    xfrm.FuenteCabecera = "Courier New"             ' ESPECIFICA EL NOMBRE DE LA FUENTE DE LA CABECERA
    If TabOne1.CurrTab = 0 Then                     ' ESPECIFICA LA ORIENTACION DE LA JOHA
        xfrm.Posicion_Hoja = Horizontal '--detalle
    Else
        xfrm.Posicion_Hoja = Vertical   '--resumen
    End If
    xfrm.Tamaño_Hoja = A_4                          ' ESPECIFICA EL TAMAÑO DE LA HOJA
    
    xfrm.ImprimirArray xCampos, Rst
    
    Set Rst = Nothing
    Set xfrm = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : MostrarPercepciones
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLENA EL CONTROL Fg1 CON LAS PERCEPCIONES EXISTENTE EN EL PERIODO ESPECIFICADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MostrarPercepciones()
    Dim SqlCad As String
    Dim Rst As New ADODB.Recordset
    
    Dim Totgendoc As Double
    Dim Totgenper  As Double
    Dim Totgencob   As Double
    
    SqlCad = " SELECT con_percepcion.numreg, con_percepcion.fchdoc, con_percepcion.tipdoc, con_percepcion.numser, con_percepcion.numdoc , mae_cliente.numruc, mae_cliente.nombre, con_tc.impven, con_percepcion.imptotdoc, con_percepcion.imptotper, con_percepcion.imptotcob, con_percepcion.idmon, mae_moneda.descripcion " & _
        " FROM mae_cliente RIGHT JOIN (mae_documento RIGHT JOIN (mae_moneda RIGHT JOIN (con_percepcion LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha) ON mae_moneda.id = con_percepcion.idmon) ON mae_documento.id = con_percepcion.tipdoc) ON mae_cliente.id = con_percepcion.idcli " & _
        " WHERE (((con_percepcion.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (con_percepcion.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " & xCadOrd
    
    RST_Busq Rst, SqlCad, xCon
    MsgBox SqlCad
    
    If Rst.RecordCount = 0 Then
        MsgBox "No se han encontrado movimientos de venta en el perido especificado", vbInformation + vbOKCancel + vbDefaultButton1, Me.Caption
        Set Rst = Nothing
        Exit Sub
    End If
    
    If OptSort1.Value = True Then Rst.Sort = "FchDoc"
    If OptSort2.Value = True Then Rst.Sort = "NumDoc"
    If OptSort3.Value = True Then Rst.Sort = "NumReg"
    
    Fg1.Cols = 1
    Fg1.Rows = 1
    
    Fg1.Cols = 15
    Fg1.ColWidth(1) = 800
    Fg1.ColWidth(2) = 950
    Fg1.ColWidth(3) = 350
    Fg1.ColWidth(4) = 1400
    Fg1.ColWidth(5) = 1200
    Fg1.ColWidth(6) = 2000
    Fg1.ColWidth(7) = 600
    Fg1.ColWidth(8) = 1200
    Fg1.ColWidth(9) = 1500
    Fg1.ColWidth(10) = 1500
    
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    Fg1.ColWidth(13) = 0
    Fg1.ColWidth(14) = 0
    
    Fg1.TextMatrix(0, 1) = "Nº Reg."
    Fg1.TextMatrix(0, 2) = "Fch. Doc."
    Fg1.TextMatrix(0, 3) = "T.D."
    Fg1.TextMatrix(0, 4) = "Nº Documento"
    Fg1.TextMatrix(0, 5) = "Nº R.U.C."
    Fg1.TextMatrix(0, 6) = "Cliente"
    Fg1.TextMatrix(0, 7) = "T.C."
    Fg1.TextMatrix(0, 8) = "Importe Doc(s)"
    Fg1.TextMatrix(0, 9) = "Importe Perc."
    Fg1.TextMatrix(0, 10) = "Total "
    
    Fg1.ColAlignment(1) = flexAlignLeftTop
    Fg1.ColAlignment(2) = flexAlignCenterTop
    Fg1.ColAlignment(3) = flexAlignLeftTop
    Fg1.ColAlignment(4) = flexAlignLeftTop
    Fg1.ColAlignment(5) = flexAlignLeftTop
    Fg1.ColAlignment(6) = flexAlignLeftTop
    Fg1.ColAlignment(7) = flexAlignRightTop
    Fg1.ColAlignment(8) = flexAlignRightTop
    Fg1.ColAlignment(9) = flexAlignRightTop
    Fg1.ColAlignment(10) = flexAlignRightTop
    
    Rst.MoveFirst
    
    While Rst.EOF = False
        Fg1.Rows = Fg1.Rows + 1
        
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("numreg")
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = Format(Rst("fchdoc"), "dd/mm/yy")
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("tipdoc")
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = Rst("numser") & "-" & Rst("numdoc")
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = Rst("numruc")
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = Rst("nombre")
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(Rst("impven"), "0.000")
        
        If TxtIdMon.Text = 1 Then
            If Rst("idmon") = 1 Then 'Soles
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(Rst("imptotdoc"), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Rst("imptotper"), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(Rst("imptotcob"), "0.00")
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format((Rst("imptotdoc") * Rst("impven")), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format((Rst("imptotper") * Rst("impven")), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format((Rst("imptotcob") * Rst("impven")), "0.00")
            End If
        Else
            If Rst("idmon") = 2 Then 'Dolares
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(Rst("imptotdoc"), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Rst("imptotper"), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(Rst("imptotcob"), "0.00")
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format((Rst("imptotdoc") / Rst("impven")), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format((Rst("imptotper") / Rst("impven")), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format((Rst("imptotcob") / Rst("impven")), "0.00")
            End If
        End If
    
        Totgendoc = Totgendoc + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 8))
        Totgenper = Totgenper + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 9))
        Totgencob = Totgencob + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 10))
        Rst.MoveNext
    Wend
    
    Fg1.Rows = Fg1.Rows + 1
    Fg1.TextMatrix(Fg1.Rows - 1, 6) = "TOTAL ==>"
    Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(Totgendoc, "0.00")
    Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Totgenper, "0.00")
    Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(Totgencob, "0.00")
    
    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : MostrarVentas
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LAS VENTAS REGISTRADAS EN EL PERIODO ESPECIFICADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************

Sub MostrarVentas()
'--06-08-2010
'--01/06/11 Agregar campo baseigv en las sentencias SQL para mostrar en reporte
'--         Mostrar en pantalla el T.C. en formato a 3 decimales


    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLFiltro As String '--Sentencia que indica el filtro a la consulta
    Dim nSQLCampos As String '--Relacion de campos a mostrar, obtenido de tabla: con_formatostipodet
    
    '--obtener el orden de presentacion de los campos
    nSQLCampos = fSetearCuadriculaColumna(xCon, 3)
    '--verificar si hay campos seleccionados para mostrar el reporte
    If nSQLCampos = "" Then Exit Sub
        
    '--verificar si hay filtro por proveedor
    If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLFiltro = " and vta_ventas.idcli = " & NulosN(LblIdCliPro.Caption) & " "
    
    '--verificar si hay filtro por documento
    If NulosN(TxtTipDoc.Text) <> 0 Then nSQLFiltro = nSQLFiltro & " and vta_ventas.tipdoc = " & NulosN(TxtTipDoc.Text) & " "

    Me.MousePointer = vbHourglass
    DoEvents
    '--

    If TxtIdMon.Text = 1 Then
 
        nSQL = "SELECT vta_ventas.id, Left([vta_ventas].[numreg],2)& [mae_libros].[codsun]& Right([vta_ventas].[numreg],4) AS registro, IIf(vta_ventas!anulado=-1,'',mae_dociden.codsun) AS tdpersun, IIf(vta_ventas!anulado=-1,'',[numruc]) AS numruc1, IIf(vta_ventas!anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre1, mae_documento.codsun AS tdsun, mae_documento.abrev, " _
            + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc2, vta_ventas.numser, vta_ventas.numdoc, vta_ventas.fchdoc, vta_ventas.fchrecep,mae_condpago.abrev AS condpag, vta_ventas.fchven, vta_ventas.glosa, mae_moneda.simbolo ,vta_ventas.tasaigv, " _
            + vbCr + " IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, IIf(vta_ventas.tipdoc=7,-1,1) AS xreal, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=1,[vta_ventas].[impbru],[vta_ventas].[impbru]*tipcam)  AS impbru1_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=1,[vta_ventas].[impbru2],[vta_ventas].[impbru2]*tipcam) AS impbru2_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=1,[vta_ventas].[impbru3],[vta_ventas].[impbru3]*tipcam) AS impbru3_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=1,[vta_ventas].[impinaf],[vta_ventas].[impinaf]*tipcam) AS impina_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=1,[vta_ventas].[impisc],[vta_ventas].[impisc]*tipcam) AS impisc_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=1,[vta_ventas].[impigv],[vta_ventas].[impigv]*tipcam) AS impigv1_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=1,[vta_ventas].[impotr],[vta_ventas].[impotr]*tipcam) AS impotros_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=1,[vta_ventas].[imptotdoc],[vta_ventas].[imptotdoc]*tipcam) AS imptot_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=1,[vta_ventas].[impdesc],[vta_ventas].[impdesc]*tipcam) AS impdesc_c, " _
            + vbCr + " ref1.* " _
            + vbCr + " FROM ((mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) RIGHT JOIN (mae_moneda RIGHT JOIN (mae_condpago RIGHT JOIN (((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) " _
            + vbCr + " LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_condpago.id = vta_ventas.idconpag) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli ) " _
            + vbCr + " LEFT JOIN " _
            + vbCr + " (SELECT  vta_ventas_1.id AS iddoc, vta_ventas.id AS refiddoc, Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) AS refregistro, mae_documento.abrev AS refabrev, mae_documento.codsun AS reftdsun, vta_ventas.fchdoc AS reffchdoc, vta_ventas.numser AS refnumser, vta_ventas.numdoc AS refnumdoc, mae_moneda.simbolo AS refsimbolo, vta_ventas.imptotdoc AS refimptot,vta_ventas.glosa as refglosa " _
            + vbCr + " FROM vta_ventas AS vta_ventas_1 INNER JOIN (((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) ON vta_ventas_1.iddocref = vta_ventas.id " _
            + vbCr + " WHERE vta_ventas_1.anulado=0 ) AS ref1  ON vta_ventas.id=ref1.iddoc " _
            + vbCr + " WHERE (((vta_ventas.fchreg) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')) AND ((Mid([vta_ventas].[numreg],1,2))<>'00')) " & nSQLFiltro

      
    Else
    
        nSQL = "SELECT vta_ventas.id, Left([vta_ventas].[numreg],2)& [mae_libros].[codsun]& Right([vta_ventas].[numreg],4) AS registro, IIf(vta_ventas!anulado=-1,'',mae_dociden.codsun) AS tdpersun, IIf(vta_ventas!anulado=-1,'',[numruc]) AS numruc1, IIf(vta_ventas!anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre1, mae_documento.codsun AS tdsun, mae_documento.abrev, " _
            + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc2, vta_ventas.numser, vta_ventas.numdoc, vta_ventas.fchdoc, vta_ventas.fchrecep, mae_condpago.abrev AS condpag, vta_ventas.fchven, vta_ventas.glosa, mae_moneda.simbolo ,vta_ventas.tasaigv, " _
            + vbCr + " IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, IIf(vta_ventas.tipdoc=7,-1,1) AS xreal, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=2,[vta_ventas].[impbru],IIF(tipcam=0,0,[vta_ventas].[impbru]/tipcam))  AS impbru1_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=2,[vta_ventas].[impbru2],IIF(tipcam=0,0,[vta_ventas].[impbru2]/tipcam)) AS impbru2_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=2,[vta_ventas].[impbru3],IIF(tipcam=0,0,[vta_ventas].[impbru3]/tipcam)) AS impbru3_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=2,[vta_ventas].[impinaf],IIF(tipcam=0,0,[vta_ventas].[impinaf]/tipcam)) AS impina_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=2,[vta_ventas].[impisc],IIF(tipcam=0,0,[vta_ventas].[impisc]/tipcam)) AS impisc_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=2,[vta_ventas].[impigv],IIF(tipcam=0,0,[vta_ventas].[impigv]/tipcam)) AS impigv1_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=2,[vta_ventas].[impotr],IIF(tipcam=0,0,[vta_ventas].[impotr]/tipcam)) AS impotros_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=2,[vta_ventas].[imptotdoc],IIF(tipcam=0,0,[vta_ventas].[imptotdoc]/tipcam)) AS imptot_c, " _
            + vbCr + " xreal * IIf([vta_ventas].[idmon]=2,[vta_ventas].[impdesc],IIF(tipcam=0,0,[vta_ventas].[impdesc]/tipcam)) AS impdesc_c, " _
            + vbCr + " ref1.* " _
            + vbCr + " FROM ((mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) RIGHT JOIN (mae_moneda RIGHT JOIN (mae_condpago RIGHT JOIN (((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) " _
            + vbCr + " LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_condpago.id = vta_ventas.idconpag) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli ) " _
            + vbCr + " LEFT JOIN " _
            + vbCr + " (SELECT  vta_ventas_1.id AS iddoc, vta_ventas.id AS refiddoc, Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) AS refregistro, mae_documento.abrev AS refabrev, mae_documento.codsun AS reftdsun, vta_ventas.fchdoc AS reffchdoc, vta_ventas.numser AS refnumser, vta_ventas.numdoc AS refnumdoc, mae_moneda.simbolo AS refsimbolo, vta_ventas.imptotdoc AS refimptot,vta_ventas.glosa as refglosa " _
            + vbCr + " FROM vta_ventas AS vta_ventas_1 INNER JOIN (((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) ON vta_ventas_1.iddocref = vta_ventas.id " _
            + vbCr + " WHERE vta_ventas_1.anulado=0 ) AS ref1  ON vta_ventas.id=ref1.iddoc " _
            + vbCr + " WHERE (((vta_ventas.fchreg) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')) AND ((Mid([vta_ventas].[numreg],1,2))<>'00')) " & nSQLFiltro
       
    End If
    
    '--armar la sentencia SQL
    nSQL = "Select " & nSQLCampos & _
            vbCr + " from ( " _
            + vbCr + nSQL _
            + vbCr + ") as consulta "
    
    '--ejecutar la consulta
    RST_Busq Rst, nSQL, xCon
    
    '--Salir si hay error en la consulta
    If Rst.State = 0 Then GoTo LaCague
    
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
            Case "impbru1_c":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impbru2_c":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impbru3_c":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impina_c":    ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impisc_c":    ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impigv1_c":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impotros_c":  ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "imptot_c":    ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impdesc_c":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "refimptot":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            
        End Select
    Next mColCampo
        
    
    If OptOpc11.Value = True Then Rst.Filter = adFilterNone                                ' mostramos todos los registros
    If OptOpc22.Value = True Then
        If TxtIdMon.Text = 1 Then Rst.Filter = "imptot_c > " & NulosN(txtBancarizacion.Text)   ' mostramos solo los de bancarizacion en Soles
        If TxtIdMon.Text = 2 Then Rst.Filter = "imptot_c > " & NulosN(txtBancarizacion.Text)   ' mostramos solo los de bancarizacion en Dolares
    End If
'    If OptOpc33.Value = True Then Rst.Filter = "spotnum<>null"                ' mostramos solo los detraccion
    
    '--Aplicar orden
    If OptSort1.Value = True Then Rst.Sort = "fchdoc"
    If OptSort2.Value = True Then Rst.Sort = "numdoc"
    If OptSort3.Value = True Then Rst.Sort = "registro"
    If OptSort4.Value = True Then Rst.Sort = "fchdoc,numdoc"
        
    LblNumreg.Caption = Rst.RecordCount

    Do While Not Rst.EOF
        DoEvents
''        ProgressBar1.Value = Rst.Bookmark
        '-----------------------------------------------
        Fg1.Rows = Fg1.Rows + 1
        
        For mCol = 0 To Rst.Fields.Count - 1
        
            Select Case LCase(Rst.Fields(mCol).Name)
                Case "fchdoc", "fchven", "fchrecep", "spotfchpag", "reffchdoc"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_DATE)
                
                Case "impbru1_c", "impbru2_c", "impbru3_c", "impina_c", "impisc_c", "impigv1_c", "imptot_c", "impotros_c", "impdesc_c", "refimptot"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_MONTO)
                    
                Case "tipcam"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), "0.000")
                    
                Case "tdpersun"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
                    
                    
                Case Else
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
            End Select
            
        Next mCol
                
        '--verificar si monto=cero y no sea anulado =>> pintar la fila para que muestre una alerta al usuario
        If NulosN(Rst("imptot_c")) = 0 And InStr(LCase(Rst("nombre1")), "anulado") = 0 Then
            GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HC0C0FF
        End If
            
        Rst.MoveNext
    Loop
        
    '--verificamos si se suman las columnas
    If ArrCampos(0) <> 0 Then
            
        '--sumando las columnas
        Fg1.Rows = Fg1.Rows + 1
        FORMATO_CELDA Fg1, Fg1.Rows - 1, IIf(ArrCampos(1) - 2 < 0, 1, ArrCampos(1) - 2), &H800000, False, , "TOTAL ==>"
        
        For mCol = 0 To UBound(ArrCampos())
            If ArrCampos(mCol) <> 0 Then
                FORMATO_CELDA Fg1, Fg1.Rows - 1, ArrCampos(mCol) + 1, &H800000, False, , Format(GRID_SUMAR_COL(Fg1, ArrCampos(mCol) + 1), FORMAT_MONTO)
            End If
        Next mCol
    
    End If
    
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

'*****************************************************************************************************
'* Nombre           : SumarColumna
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : SUMA LAS COLUMNAS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub SumarColumna()
    Dim xTotal1, xTotal2, xTotal3, xTotal4 As Double
    Dim xIGV1, xIGV2, xIGV3 As Double
    Dim xISC, xOTros, xTotalTot As Double
    Dim A As Integer
    Dim xFila As Integer
        
    Fg1.Rows = Fg1.Rows + 1
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H800000, False, , "TOTAL ==>"
        
    xFila = 2
    For A = 1 To Fg1.Rows - 2
        If Fg1.TextMatrix(xFila, 4) = "07" Then
            xTotal1 = xTotal1 - Abs(NulosN(Fg1.TextMatrix(xFila, 12)))
            xTotal2 = xTotal2 - Abs(NulosN(Fg1.TextMatrix(xFila, 13)))
            xTotal3 = xTotal3 - Abs(NulosN(Fg1.TextMatrix(xFila, 14)))
            xTotal4 = xTotal4 - Abs(NulosN(Fg1.TextMatrix(xFila, 15)))
            
            xISC = xISC - Abs(NulosN(Fg1.TextMatrix(xFila, 16)))
            xIGV1 = xIGV1 - Abs(NulosN(Fg1.TextMatrix(xFila, 17)))
            xOTros = xOTros - Abs(NulosN(Fg1.TextMatrix(xFila, 18)))
            xTotalTot = xTotalTot - Abs(NulosN(Fg1.TextMatrix(xFila, 19)))
        Else
            xTotal1 = xTotal1 + NulosN(Fg1.TextMatrix(xFila, 12))
            xTotal2 = xTotal2 + NulosN(Fg1.TextMatrix(xFila, 13))
            xTotal3 = xTotal3 + NulosN(Fg1.TextMatrix(xFila, 14))
            xTotal4 = xTotal4 + NulosN(Fg1.TextMatrix(xFila, 15))
            
            xISC = xISC + NulosN(Fg1.TextMatrix(xFila, 16))
            xIGV1 = xIGV1 + NulosN(Fg1.TextMatrix(xFila, 17))
            xOTros = xOTros + NulosN(Fg1.TextMatrix(xFila, 18))
            xTotalTot = xTotalTot + NulosN(Fg1.TextMatrix(xFila, 19))
        End If
        xFila = xFila + 1
    Next A
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H800000, False, , Format(xTotal1, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H800000, False, , Format(xTotal2, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, &H800000, False, , Format(xTotal3, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 15, &H800000, False, , Format(xTotal4, FORMAT_MONTO)
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &H800000, False, , Format(xISC, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 17, &H800000, False, , Format(xIGV1, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &H800000, False, , Format(xOTros, FORMAT_MONTO)
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 19, &H800000, False, , Format(xTotalTot, FORMAT_MONTO)
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
        
        TxtCliPro.Text = ""
        
        TabOne2.CurrTab = 0
        
        txtBancarizacion.Text = "0.00"
        
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
        
        '--enfocar en la pestaña del detalle
        TabOne1.CurrTab = 0
        
        SeEjecuto = True
        TxtFchIni.SetFocus
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE LE FORMULARIO
    SeEjecuto = False
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    
    OptSort3.Value = True
    OptOpc11.Value = True
    
    LblNumreg.Caption = ""
    
    Dim xRs As New ADODB.Recordset
    
    ' SELECCIONAMOS EL FORMATO ACTUAL PARA LA IMPRESION DEL LIBRO "REGISTRO DE COMPRAS"
    RST_Busq xRs, "SELECT con_formatostipo.* From con_formatostipo WHERE (((con_formatostipo.defecto)=-1) AND ((con_formatostipo.idformato)=3))", xCon

    xFormatoActual = xRs("id")
    
    Set xRs = Nothing
    
    '--dar formato al detalle
    SetearCuadricula Fg1, 3, xCon, 1, xFormatoActual, False
    
    '--dar formato al resumen
    SetearCuadricula fg2, 3, xCon, 1, 3
    
    '--buscar los registros
    Fg1.AutoSearch = flexSearchFromTop
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
End Sub

'*****************************************************************************************************
'* Nombre           : Configurar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL REGISTRO DE VENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Configurar()
    Dim xform As New SGI2_funciones.Varias
    
    If xform.CambioOpcionLiro(3, xCon, 1) = True Then
    
        Dim xRs As New ADODB.Recordset
        
        ' SELECCIONAMOS EL FORMATO ACTUAL PARA LA IMPRESION DEL LIBRO "REGISTRO DE COMPRAS"
        RST_Busq xRs, "SELECT con_formatostipo.* From con_formatostipo WHERE (((con_formatostipo.defecto)=-1) AND ((con_formatostipo.idformato)=3))", xCon
    
        xFormatoActual = xRs("id")
        
        Set xRs = Nothing
        
        SetearCuadricula Fg1, 3, xCon, 1, xFormatoActual, False
            
        ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
        If TxtFchIni.Valor = "" Or TxtFchFin.Valor = "" Then
            MsgBox "No ha especificado el periodo de la consulta", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            TxtFchIni.SetFocus
            Exit Sub
        End If

        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            TxtFchIni.SetFocus
            Exit Sub
        End If
        MostrarVentas
    End If
    Set xform = Nothing
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height > 3000 Then
        TabOne1.Top = 1650
        TabOne1.Width = Me.Width - 150
        TabOne1.Height = Me.Height - 2100
    End If
End Sub

Private Sub OptOpc11_Click()
    txtBancarizacion.Visible = False
End Sub

Private Sub OptOpc22_Click()
    txtBancarizacion.Visible = True
End Sub

Private Sub OptOpc33_Click()
    txtBancarizacion.Visible = False
End Sub

Private Sub OptOpc44_Click()
    txtBancarizacion.Visible = False
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
        '--
        '--verificar que este ingresado la base para mostrar los registros que cumplen con la bancarizacion
        If OptOpc22.Value = True Then
            If NulosN(txtBancarizacion.Text) = 0 Then
                MsgBox "Falta especificar la base de la bancarizacion expresado en " & LblMoneda.Caption, vbInformation, xTitulo
                txtBancarizacion.SetFocus
                Exit Sub
            End If
        End If
        
        MostrarVentas
        MostrarVentasResumen
    End If
    
    If Button.Index = 3 Then
        Dim xFun As New SGI2_funciones.formularios
        If Fg1.Rows = 2 Then
            MsgBox "No hay registro que exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        If TabOne1.CurrTab = 0 Then
            xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "LIBRO VENTAS", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Por Fecha", "Registro de Ventas"     ', Rst, ""
        Else
            xFun.VSFlexGrid_Exportar_MSExcel xCon, fg2, "RESUMEN - LIBRO VENTAS", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Por Fecha", "Registro de Ventas"   ', Rst, ""
        End If
        Set xFun = Nothing
    End If
    
    If Button.Index = 4 Then
        pImprimir
    End If
    
    If Button.Index = 5 Then
        Configurar
    End If
    
    If Button.Index = 7 Then
        Unload Me
    End If
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
    & " FROM vta_ventas INNER JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id " _
    & " WHERE (((vta_ventas.fchreg) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " _
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

Private Sub txtBancarizacion_Validate(Cancel As Boolean)
    If NulosN(txtBancarizacion.Text) <> 0 Then
        txtBancarizacion.Text = Format(txtBancarizacion.Text, FORMAT_MONTO)
    Else
        txtBancarizacion.Text = "0.00"
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
        & " FROM vta_ventas INNER JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id " _
        & " WHERE (((vta_ventas.fchreg) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) and vta_ventas.tipdoc = " & NulosN(TxtTipDoc.Text) & " " _
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

'Private Sub MostrarVentasResumen()
'    Dim Rst As New ADODB.Recordset
'    Dim nSQL As String
'    Fg2.Rows = Fg2.FixedRows
'    If txtidmon.text=1 Then
'    nSQL = "SELECT mae_documento.codsun AS tipdoc, mae_documento.descripcion, Sum(IIf([vta_ventas]![idmon]=1,[impbru2],[impbru2]*tipcam)) AS impbru2_c, Sum(IIf([vta_ventas]![idmon]=1,[impbru],[impbru]*tipcam)) AS impbru_c, Sum(IIf([vta_ventas]![idmon]=1,[impbru3],[impbru3]*tipcam)) AS impbru3_c, Sum(IIf([vta_ventas]![idmon]=1,[impinaf],[impinaf]*tipcam)) AS impinaf_c, Sum(IIf([vta_ventas]![idmon]=1,[impisc],[impisc]*tipcam)) AS impisc_c, Sum(IIf([vta_ventas]![idmon]=1,[impigv],[impigv]*tipcam)) AS impigv_c, 0 AS otrostrib, Sum(IIf([vta_ventas]![idmon]=1,[imptotdoc],[imptotdoc]*tipcam)) AS imptotdoc_c " _
'        + vbCr + " FROM ((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id " _
'        + vbCr + " WHERE (((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((Mid([numreg],1,2))<>0)) " _
'        + vbCr + " GROUP BY mae_documento.codsun, mae_documento.descripcion;"
'    Else
'
'    End If
'
'    DoEvents
'    RST_Busq Rst, nSQL, xCon
'
'
'    Do While Not Rst.EOF
'        DoEvents
'        Fg2.Rows = Fg2.Rows + 1
'        Fg2.TextMatrix(Fg2.Rows - 1, 1) = Rst("tipdoc")
'        Fg2.TextMatrix(Fg2.Rows - 1, 1) = Rst("descripcion")
'        If NulosN(Rst("tipdoc")) = 7 Then
'            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(NulosN(Rst("impbru2_c")), "-0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(Rst("impbru_c")), "-0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(NulosN(Rst("impbru3_c")), "-0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(NulosN(Rst("impinaf_c")), "-0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 7) = Format(NulosN(Rst("impisc_c")), "-0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(NulosN(Rst("impigv_c")), "-0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 9) = Format(NulosN(Rst("otrostrib")), "-0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 10) = Format(NulosN(Rst("imptotdoc_c")), "-0.00")
'        Else
'            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(NulosN(Rst("impbru2_c")), "0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(Rst("impbru_c")), "0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(NulosN(Rst("impbru3_c")), "0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(NulosN(Rst("impinaf_c")), "0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 7) = Format(NulosN(Rst("impisc_c")), "0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(NulosN(Rst("impigv_c")), "0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 9) = Format(NulosN(Rst("otrostrib")), "0.00")
'            Fg2.TextMatrix(Fg2.Rows - 1, 10) = Format(NulosN(Rst("imptotdoc_c")), "0.00")
'        End If
'        Rst.MoveNext
'    Loop
'
'    Set Rst = Nothing
'
'
'End Sub









'***************************************************************************************************************************************

Private Sub CmdBusCliPro_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xform.Titulo = "Buscando Clientes"
    xform.SqlCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente ORDER BY mae_cliente.nombre"
    xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"

    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"

    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCliPro.Text = xRs("nombre")
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
'* Nombre           : MostrarVentasResumen
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL RESUMEN POR DOCUMENTO DE LAS VENTAS REGISTRADAS EN EL PERIODO ESPECIFICADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MostrarVentasResumen()
    Dim SqlCad As String
    Dim A As Long
    Dim Rst As New ADODB.Recordset
    Dim nSQLFiltro As String '--Sentencia que indica el filtro a la consulta

    
    '--verificar si se puede mostrar los datos, esto dependera que esta la configuracion del grid en la base de datos
    If fg2.Cols = 1 Then
        Exit Sub
    End If
    
    
    '--verificar si hay filtro por cliente
    If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLFiltro = " and vta_ventas.idcli = " & NulosN(LblIdCliPro.Caption) & " and vta_ventas.anulado=0 "
    
    '--verificar si hay filtro por documento
    If NulosN(TxtTipDoc.Text) <> 0 Then nSQLFiltro = nSQLFiltro & " and vta_ventas.tipdoc = " & NulosN(TxtTipDoc.Text) & " "
    
    '--si selecciona en MN
    If TxtIdMon.Text = 1 Then
    
        SqlCad = "SELECT CONSULTA1.tdocnombre, CONSULTA1.abrev, CONSULTA1.tipdoc, Sum(CONSULTA1.impbru2_c) AS impbru2_c1, Sum(CONSULTA1.impbru_c) AS impbru_c1, Sum(CONSULTA1.impbru3_c) AS impbru3_c1, Sum(CONSULTA1.impinaf_c) AS impinaf_c1, Sum(CONSULTA1.impisc_c) AS impisc_c1, Sum(CONSULTA1.impigv_c) AS impigv_c1, Sum(CONSULTA1.otrostrib) AS otrostrib1, Sum(CONSULTA1.imptotdoc_c) AS imptotdoc_c1 " _
            & " FROM " _
            & vbCr & " (SELECT vta_ventas.id, Mid(vta_ventas!numreg,1,2)+mae_libros!codsun+Mid(vta_ventas!numreg,3,4) AS numreg, vta_ventas.fchdoc, vta_ventas.fchven, " _
            & " mae_documento.codsun AS tipdoc, vta_ventas.numser, vta_ventas.numdoc, IIf(vta_ventas!anulado=-1,'',mae_dociden.codsun) AS tdiden, " _
            & " IIf(vta_ventas!anulado=-1,'',[numruc]) AS numruc2, IIf(vta_ventas!anulado=-1,'ANULADO',[nombre]) AS nombre2, mae_moneda.simbolo, " _
            & vbCr & " con_tc.impven, IIf([vta_ventas].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
            & " vta_ventas.fchreg, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=1,[impbru2],[impbru2]*tipcam),'0.00') AS impbru2_c, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=1,[impbru],[impbru]*tipcam),'0.00') AS impbru_c, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=1,[impbru3],[impbru3]*tipcam),'0.00') AS impbru3_c, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=1,[impinaf],[impinaf]*tipcam),'0.00') AS impinaf_c, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=1,[impisc],[impisc]*tipcam),'0.00') AS impisc_c, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=1,[impigv],[impigv]*tipcam),'0.00') AS impigv_c, " _
            & vbCr & " 0 AS otrostrib, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=1,[imptotdoc],[imptotdoc]*tipcam),'0.00') AS imptotdoc_c, " _
            & vbCr & " mae_documento.descripcion AS tdocnombre, mae_documento.abrev " _
            & vbCr & " FROM (mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) RIGHT JOIN ((((vta_ventas LEFT JOIN mae_documento " _
            & " ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
            & " LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) ON mae_cliente.id = vta_ventas.idcli " _
            & vbCr & " WHERE (((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((Mid([numreg],1,2))<>0)) " & nSQLFiltro & " ) " _
            & vbCr & " AS CONSULTA1 " _
            & vbCr & " GROUP BY CONSULTA1.tdocnombre, CONSULTA1.abrev, CONSULTA1.tipdoc " _
            & vbCr & " ORDER BY CONSULTA1.tipdoc; "
            
    '--si selecciona en ME
    ElseIf TxtIdMon = 2 Then
    
        SqlCad = "SELECT CONSULTA1.tdocnombre, CONSULTA1.abrev, CONSULTA1.tipdoc, Sum(CONSULTA1.impbru2_c) AS impbru2_c1, Sum(CONSULTA1.impbru_c) AS impbru_c1, Sum(CONSULTA1.impbru3_c) AS impbru3_c1, Sum(CONSULTA1.impinaf_c) AS impinaf_c1, Sum(CONSULTA1.impisc_c) AS impisc_c1, Sum(CONSULTA1.impigv_c) AS impigv_c1, Sum(CONSULTA1.otrostrib) AS otrostrib1, Sum(CONSULTA1.imptotdoc_c) AS imptotdoc_c1 " _
            & " FROM " _
            & vbCr & " (SELECT vta_ventas.id, Mid(vta_ventas!numreg,1,2)+mae_libros!codsun+Mid(vta_ventas!numreg,3,4) AS numreg, vta_ventas.fchdoc, vta_ventas.fchven, " _
            & " mae_documento.codsun AS tipdoc, vta_ventas.numser, vta_ventas.numdoc, IIf(vta_ventas!anulado=-1,'',mae_dociden.codsun) AS tdiden, " _
            & " IIf(vta_ventas!anulado=-1,'',[numruc]) AS numruc2, IIf(vta_ventas!anulado=-1,'ANULADO',[nombre]) AS nombre2, mae_moneda.simbolo, " _
            & vbCr & " con_tc.impven, IIf([vta_ventas].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
            & " vta_ventas.fchreg, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=2,[impbru2],IIF(tipcam=0,0,[impbru2]/tipcam)),'0.00') AS impbru2_c, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=2,[impbru],IIF(tipcam=0,0,[impbru]/tipcam)),'0.00') AS impbru_c, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=2,[impbru3],IIF(tipcam=0,0,[impbru3]/tipcam)),'0.00') AS impbru3_c, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=2,[impinaf],IIF(tipcam=0,0,[impinaf]*tipcam)),'0.00') AS impinaf_c, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=2,[impisc],IIF(tipcam=0,0,[impisc]/tipcam)),'0.00') AS impisc_c, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=2,[impigv],IIF(tipcam=0,0,[impigv]/tipcam)),'0.00') AS impigv_c, " _
            & vbCr & " 0 AS otrostrib, " _
            & vbCr & " FORMAT(IIf([vta_ventas]![idmon]=2,[imptotdoc],IIF(tipcam=0,0,[imptotdoc]/tipcam)),'0.00') AS imptotdoc_c, " _
            & vbCr & " mae_documento.descripcion AS tdocnombre, mae_documento.abrev " _
            & vbCr & " FROM (mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) RIGHT JOIN ((((vta_ventas LEFT JOIN mae_documento " _
            & " ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
            & " LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) ON mae_cliente.id = vta_ventas.idcli  " _
            & vbCr & " WHERE (((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((Mid([numreg],1,2))<>'00')) " & nSQLFiltro & " ) AS CONSULTA1 " _
            & vbCr & " GROUP BY CONSULTA1.tdocnombre, CONSULTA1.abrev, CONSULTA1.tipdoc " _
            & vbCr & " ORDER BY CONSULTA1.tipdoc; "
            
    End If
    
    
    '--ejecutar consulta
    RST_Busq Rst, SqlCad, xCon
    
    If OptOpc11.Value = True Then Rst.Filter = adFilterNone                      ' mostramos todos los registros
''pendiente revisar luego
''    If OptOpc22.Value = True Then
''        If TxtIdMon.Text = 1 Then Rst.Filter = "imptotdoc_c > 3500"            ' mostramos solo los de bancarizacion en Soles
''        If TxtIdMon.Text = 2 Then Rst.Filter = "imptotdoc_c > 1000"            ' mostramos solo los de bancarizacion en Dolares
''    End If
    
    
    fg2.Rows = 2
    
    DoEvents
    Me.MousePointer = vbHourglass
    If Rst.RecordCount <> 0 Then
        
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            fg2.Rows = fg2.Rows + 1
            fg2.TextMatrix(fg2.Rows - 1, 1) = NulosC(Rst("tdocnombre"))
            fg2.TextMatrix(fg2.Rows - 1, 2) = NulosC(Rst("abrev"))
            
            If NulosC(Rst("tipdoc")) = "07" Then
                fg2.TextMatrix(fg2.Rows - 1, 3) = Format(NulosN(Rst("impbru2_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 4) = Format(NulosN(Rst("impbru_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 5) = Format(NulosN(Rst("impbru3_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 6) = Format(NulosN(Rst("impinaf_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 7) = Format(NulosN(Rst("impisc_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 8) = Format(NulosN(Rst("impigv_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 9) = Format(NulosN(Rst("otrostrib1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 10) = Format(NulosN(Rst("imptotdoc_c1")), "-" & FORMAT_MONTO)
            Else
                fg2.TextMatrix(fg2.Rows - 1, 3) = Format(NulosN(Rst("impbru2_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 4) = Format(NulosN(Rst("impbru_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 5) = Format(NulosN(Rst("impbru3_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 6) = Format(NulosN(Rst("impinaf_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 7) = Format(NulosN(Rst("impisc_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 8) = Format(NulosN(Rst("impigv_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 9) = Format(NulosN(Rst("otrostrib1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 10) = Format(NulosN(Rst("imptotdoc_c1")), FORMAT_MONTO)
            End If
            
            '---------------------------------------------------------------------------------------
                        
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    
    '--mostrar totales
    fg2.Rows = fg2.Rows + 1
    Dim xCol As Integer
    FORMATO_CELDA fg2, fg2.Rows - 1, 1, &H800000, False, , "TOTAL ==>"
    For xCol = 3 To 10
        FORMATO_CELDA fg2, fg2.Rows - 1, xCol, &H800000, False, , Format(NulosN(GRID_SUMAR_COL(fg2, xCol, fg2.FixedCols, fg2.Rows - 2)), FORMAT_MONTO)
    Next
    '------
    
    
    Me.MousePointer = vbDefault
End Sub


Sub MostrarVentas_xxxx()
    '--dejado de usar el 15/10/10 porque esta consulta muestra datos fijos segun una columna establecida, si el usuario cambia el orden de la presentacion
    'de la consulta los datos que se presenten no coincidiran con la cabecera; El cambio consiste en hacer que la consulta se sincronice con la configuracion del reporte.
    
    'se modifica las sgtes linea de codigo
    'Form_Load , Configurar
    'SetearCuadricula Fg1, 3, xCon, 1, xFormatoActual, True
    'pImprimir
    'RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 3) And ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ")) " _
                & " ORDER BY con_formatostipodet.orden", xCon
    
    'por lo sgte
    'Form_Load , Configurar
    'SetearCuadricula Fg1, 3, xCon, 1, xFormatoActual, False
    'pImprimir
    'RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet " _
            & " Where (((con_formatostipodet.idformato) = 3) And ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ") " _
            & " And ((con_formatostipodet.mostrar) = -1)) ORDER BY con_formatostipodet.orden", xCon
       
    
    
    Dim SqlCad As String
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    Dim nSQLTipDoc As String
    Dim nSQLCli As String
    
    '--verificar si hay filtro por cliente
    If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLCli = " and vta_ventas.idcli = " & NulosN(LblIdCliPro.Caption) & " and vta_ventas.anulado=0 "
    
    '--verificar si hay filtro por documento
    If NulosN(TxtTipDoc.Text) <> 0 Then nSQLTipDoc = " and vta_ventas.tipdoc = " & NulosN(TxtTipDoc.Text) & " "
    
    '--si selecciona en MN
    If TxtIdMon.Text = 1 Then
    
        SqlCad = "SELECT CONSULTA1.*, CONSULTA2.fac_tcp, CONSULTA2.fac_fchdoc, CONSULTA2.fac_numser, CONSULTA2.fac_numdoc" _
            & " From " _
            & " (SELECT vta_ventas.id, Mid(vta_ventas!numreg,1,2)+mae_libros!codsun+Mid(vta_ventas!numreg,3,4) AS numreg, vta_ventas.fchdoc, vta_ventas.fchven, " _
            & " mae_documento.codsun AS tipdoc, vta_ventas.numser, vta_ventas.numdoc, IIf(vta_ventas!anulado=-1,'',mae_dociden.codsun) AS tdiden, " _
            & " IIf(vta_ventas!anulado=-1,'',[numruc]) AS numruc2, IIf(vta_ventas!anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre2, mae_moneda.simbolo, " _
            & " con_tc.impven, IIf([vta_ventas].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
            & " vta_ventas.fchreg, " _
            & " IIf([vta_ventas]![idmon]=1,[impbru],[impbru]*tipcam) AS impbru_c, " _
            & " IIf([vta_ventas]![idmon]=1,[impbru2],[impbru2]*tipcam) AS impbru2_c, " _
            & " IIf([vta_ventas]![idmon]=1,[impbru3],[impbru3]*tipcam) AS impbru3_c, " _
            & " IIf([vta_ventas]![idmon]=1,[impinaf],[impinaf]*tipcam) AS impinaf_c, " _
            & " IIf([vta_ventas]![idmon]=1,[impisc],[impisc]*tipcam) AS impisc_c, " _
            & " IIf([vta_ventas]![idmon]=1,[impigv],[impigv]*tipcam) AS impigv_c, " _
            & " 0 AS otrostrib, " _
            & " IIf([vta_ventas]![idmon]=1,[imptotdoc],[imptotdoc]*tipcam) AS imptotdoc_c " _
            & " FROM (mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) RIGHT JOIN ((((vta_ventas LEFT JOIN mae_documento " _
            & " ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
            & " LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) ON mae_cliente.id = vta_ventas.idcli " _
            & " WHERE (((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((Mid([numreg],1,2))<>0)) " & nSQLTipDoc & nSQLCli & " ) AS CONSULTA1" _
            & " LEFT JOIN " _
            & " (SELECT vta_ventas.id, vta_ventas.iddocref, mae_documento.abrev AS fac_tcp, vta_ventas_1.fchdoc AS fac_fchdoc, vta_ventas_1.numser AS fac_numser,  " _
            & " vta_ventas_1.numdoc AS fac_numdoc FROM (vta_ventas LEFT JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref=vta_ventas_1.id) " _
            & " LEFT JOIN mae_documento ON vta_ventas_1.tipdoc=mae_documento.id WHERE (((vta_ventas.iddocref)<>0)) " & nSQLTipDoc & " ) AS CONSULTA2 ON CONSULTA1.id = CONSULTA2.id;"
            
    '--si selecciona en ME
    ElseIf TxtIdMon = 2 Then
    
        SqlCad = "SELECT CONSULTA1.*, CONSULTA2.fac_tcp, CONSULTA2.fac_fchdoc, CONSULTA2.fac_numser, CONSULTA2.fac_numdoc" _
            & " From " _
            & " (SELECT vta_ventas.id, Mid(vta_ventas!numreg,1,2)+mae_libros!codsun+Mid(vta_ventas!numreg,3,4) AS numreg, vta_ventas.fchdoc, vta_ventas.fchven, " _
            & " mae_documento.codsun AS tipdoc, vta_ventas.numser, vta_ventas.numdoc, IIf(vta_ventas!anulado=-1,'',mae_dociden.codsun) AS tdiden, " _
            & " IIf(vta_ventas!anulado=-1,'',[numruc]) AS numruc2, IIf(vta_ventas!anulado=-1,'ANULADO',[nombre]) AS nombre2, mae_moneda.simbolo, " _
            & " con_tc.impven, IIf([vta_ventas].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
            & " vta_ventas.fchreg, " _
            & " IIf([vta_ventas]![idmon]=2,[impbru2],IIF(tipcam=0,0,[impbru2]/tipcam)) AS impbru2_c, " _
            & " IIf([vta_ventas]![idmon]=2,[impbru],IIF(tipcam=0,0,[impbru]/tipcam)) AS impbru_c, " _
            & " IIf([vta_ventas]![idmon]=2,[impbru3],IIF(tipcam=0,0,[impbru3]/tipcam)) AS impbru3_c, " _
            & " IIf([vta_ventas]![idmon]=2,[impinaf],IIF(tipcam=0,0,[impinaf]*tipcam)) AS impinaf_c, " _
            & " IIf([vta_ventas]![idmon]=2,[impisc],IIF(tipcam=0,0,[impisc]/tipcam)) AS impisc_c, " _
            & " IIf([vta_ventas]![idmon]=2,[impigv],IIF(tipcam=0,0,[impigv]/tipcam)) AS impigv_c, " _
            & " 0 AS otrostrib, " _
            & " IIf([vta_ventas]![idmon]=2,[imptotdoc],IIF(tipcam=0,0,[imptotdoc]/tipcam)) AS imptotdoc_c " _
            & " FROM (mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) RIGHT JOIN ((((vta_ventas LEFT JOIN mae_documento " _
            & " ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
            & " LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) ON mae_cliente.id = vta_ventas.idcli WHERE (((vta_ventas.fchreg)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (vta_ventas.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((Mid([numreg],1,2))<>'00')) " & nSQLTipDoc & " ) AS CONSULTA1" _
            & " LEFT JOIN " _
            & " (SELECT vta_ventas.id, vta_ventas.iddocref, mae_documento.abrev AS fac_tcp, vta_ventas_1.fchdoc AS fac_fchdoc, vta_ventas_1.numser AS fac_numser,  " _
            & " vta_ventas_1.numdoc AS fac_numdoc FROM (vta_ventas LEFT JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref=vta_ventas_1.id) " _
            & " LEFT JOIN mae_documento ON vta_ventas_1.tipdoc=mae_documento.id WHERE (((vta_ventas.iddocref)<>0)) " & nSQLTipDoc & " ) AS CONSULTA2 ON CONSULTA1.id = CONSULTA2.id;"
    End If
    
    
    '--ejecutar consulta
    RST_Busq Rst, SqlCad, xCon
    
    If OptSort1.Value = True Then Rst.Sort = "fchdoc"                            ' ORDENAMOS POR FECHA DEL DOCUMENTO
    If OptSort2.Value = True Then Rst.Sort = "numser,numdoc"                     ' ORDENAMOS POR NUMERO DE DOCUMENTO
    If OptSort3.Value = True Then Rst.Sort = "numreg"                            ' ORDENAMOS POR NUMERO DE REGISTRO
    If OptSort4.Value = True Then Rst.Sort = "fchdoc,numser,numdoc"              ' ORDENAMOS POR FECHA DE DOCUMENTO Y NUMERO DE DOCUMENTO
    
    If OptOpc11.Value = True Then Rst.Filter = adFilterNone                      ' mostramos todos los registros
    If OptOpc22.Value = True Then
        If TxtIdMon.Text = 1 Then Rst.Filter = "imptotdoc_c > 3500"            ' mostramos solo los de bancarizacion en Soles
        If TxtIdMon.Text = 2 Then Rst.Filter = "imptotdoc_c > 1000"            ' mostramos solo los de bancarizacion en Dolares
    End If
    
    LblNumreg.Caption = Rst.RecordCount
    
    'TabOne1.CurrTab = 0
    Fg1.Rows = 2
    
    DoEvents
    Me.MousePointer = vbHourglass
    If Rst.RecordCount <> 0 Then
        
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("numreg"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Format(NulosC(Rst("fchdoc")), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(NulosC(Rst("fchven")), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(Rst("tipdoc"))
            If NulosC(Rst("numser")) <> "" Then
                Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(Rst("numser"))
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(Rst("numdoc"))
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 5) = Mid(NulosC(Rst("numdoc")), 1, 3)
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = Mid(NulosC(Rst("numdoc")), 5, 20)
            End If
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(Rst("tdiden"))
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(Rst("numruc2"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(Rst("nombre2"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(Rst("tipcam")), "0.000")
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosC(Rst("simbolo"))
            
            If Fg1.TextMatrix(Fg1.Rows - 1, 4) = "07" Then
                Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(NulosN(Rst("impbru2_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("impbru_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(NulosN(Rst("impbru3_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(NulosN(Rst("impinaf_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(NulosN(Rst("impisc_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(NulosN(Rst("impigv_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(NulosN(Rst("otrostrib")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(NulosN(Rst("imptotdoc_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(Rst("fac_fchdoc"), "dd/mm/yy")
                Fg1.TextMatrix(Fg1.Rows - 1, 21) = NulosC(Rst("fac_tcp"))
                Fg1.TextMatrix(Fg1.Rows - 1, 22) = NulosC(Rst("fac_numser"))
                Fg1.TextMatrix(Fg1.Rows - 1, 23) = NulosC(Rst("fac_numdoc"))
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(NulosN(Rst("impbru2_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("impbru_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(NulosN(Rst("impbru3_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(NulosN(Rst("impinaf_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(NulosN(Rst("impisc_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(NulosN(Rst("impigv_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(NulosN(Rst("otrostrib")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(NulosN(Rst("imptotdoc_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(Rst("fac_fchdoc"), "dd/mm/yy")
                Fg1.TextMatrix(Fg1.Rows - 1, 21) = NulosC(Rst("fac_tcp"))
                Fg1.TextMatrix(Fg1.Rows - 1, 22) = NulosC(Rst("fac_numser"))
                Fg1.TextMatrix(Fg1.Rows - 1, 23) = NulosC(Rst("fac_numdoc"))
            End If
            
            '--verificar si monto=cero y no sea anulado =>> pintar la fila para que muestre una alerta al usuario
            If NulosN(Rst("imptotdoc_c")) = 0 And InStr(LCase(Rst("nombre2")), "anulado") = 0 Then
                GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HC0C0FF
            End If
            '---------------------------------------------------------------------------------------
                        
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    
    SumarColumna
    Me.MousePointer = vbDefault
End Sub




