VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "aspatextboxfecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultaMayor2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Mayor Auxiliar"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
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
      Height          =   660
      Left            =   15
      TabIndex        =   31
      Top             =   990
      Width           =   2745
      Begin VB.CommandButton CmdBusMon 
         Height          =   230
         Left            =   345
         Picture         =   "FrmConsultaMayor2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   270
         Width           =   210
      End
      Begin VB.TextBox TxtIdMon 
         Height          =   300
         Left            =   30
         MaxLength       =   1
         TabIndex        =   33
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
         Left            =   585
         TabIndex        =   34
         Top             =   240
         Width           =   2085
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   2790
      TabIndex        =   5
      Top             =   360
      Width           =   7485
      Begin VB.CheckBox chk 
         Caption         =   "&Procesar Todas las Cuentas"
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   60
         TabIndex        =   35
         Top             =   870
         Width           =   1695
      End
      Begin VB.CommandButton cmd_periodo 
         Height          =   255
         Index           =   1
         Left            =   1410
         Picture         =   "FrmConsultaMayor2.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1650
         Width           =   285
      End
      Begin VB.CommandButton cmd_periodo 
         Height          =   255
         Index           =   0
         Left            =   1410
         Picture         =   "FrmConsultaMayor2.frx":04B4
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1320
         Width           =   285
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Eliminar"
         Height          =   405
         Left            =   6600
         TabIndex        =   4
         Top             =   750
         Width           =   765
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Agregar "
         Height          =   405
         Left            =   6600
         TabIndex        =   3
         Top             =   240
         Width           =   765
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   990
         Left            =   1770
         TabIndex        =   2
         Top             =   165
         Width           =   4740
         _cx             =   8361
         _cy             =   1746
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
         AllowUserResizing=   0
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
         FormatString    =   $"FrmConsultaMayor2.frx":0836
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
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   435
         TabIndex        =   0
         Top             =   195
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
         Left            =   435
         TabIndex        =   1
         Top             =   510
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
      Begin VB.Label lbl_periodo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_periodo(1)"
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
         Left            =   435
         TabIndex        =   29
         Top             =   1620
         Width           =   1290
      End
      Begin VB.Label lbl_periodo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_periodo(0)"
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
         Left            =   435
         TabIndex        =   27
         Top             =   1290
         Width           =   1290
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   7
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   6
         Top             =   555
         Width           =   135
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "[ Tipo de Consulta ]"
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
      Height          =   630
      Left            =   15
      TabIndex        =   23
      Top             =   360
      Width           =   2745
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Fecha"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Top             =   300
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Periodo"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   24
         Top             =   300
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   1290
      Left            =   10290
      TabIndex        =   19
      Top             =   360
      Width           =   1605
      Begin VB.CheckBox ChkResumen 
         Caption         =   "Ver Resumen"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   90
         TabIndex        =   36
         Top             =   960
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.OptionButton opt 
         Caption         =   "Nº Documento"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   22
         Top             =   465
         Width           =   1350
      End
      Begin VB.OptionButton opt 
         Caption         =   "Fch Emisión"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   21
         Top             =   225
         Width           =   1200
      End
      Begin VB.OptionButton opt 
         Caption         =   "Nº Registro"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   20
         Top             =   705
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame fra_msg 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   270
      Left            =   6015
      TabIndex        =   17
      Top             =   7350
      Visible         =   0   'False
      Width           =   5730
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Obs: Se recomienda minimizar la ventana para agilizar el proceso"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   30
         TabIndex        =   18
         Top             =   15
         Width           =   5535
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   795
      Left            =   3000
      TabIndex        =   13
      Top             =   3780
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   105
         TabIndex        =   14
         Top             =   330
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5925
         Y1              =   780
         Y2              =   765
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
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   5940
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Label lbl 
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
         Index           =   2
         Left            =   4365
         TabIndex        =   16
         Top             =   90
         Width           =   1530
      End
      Begin VB.Label Label3 
         Caption         =   "Procesando Asientos"
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
         Left            =   165
         TabIndex        =   15
         Top             =   90
         Width           =   4020
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
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5985
      Left            =   0
      TabIndex        =   8
      Top             =   1665
      Width           =   11910
      _cx             =   21008
      _cy             =   10557
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
      Caption         =   "   Detalle   |   Resumen   "
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
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5565
         Left            =   12555
         TabIndex        =   10
         Top             =   45
         Width           =   11820
         Begin VSFlex7Ctl.VSFlexGrid Fg3 
            Height          =   5550
            Left            =   15
            TabIndex        =   11
            Top             =   0
            Width           =   11790
            _cx             =   20796
            _cy             =   9790
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConsultaMayor2.frx":08BB
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5565
         Left            =   45
         TabIndex        =   9
         Top             =   45
         Width           =   11820
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5550
            Left            =   15
            TabIndex        =   12
            Top             =   0
            Width           =   11790
            _cx             =   20796
            _cy             =   9790
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
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConsultaMayor2.frx":09C1
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
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":0AFA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":103E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":13D0
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":1554
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":19A8
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":1AC0
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":2004
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":2548
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":265C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":2770
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":2BC4
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":2D30
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":3278
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":3592
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":3924
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaMayor2.frx":3CB6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Asiento"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.ToolTipText     =   "Configurar Formatos"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmConsultaMayor2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstTmp As New ADODB.Recordset
Dim RstTmp2 As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim Agregando As Boolean

Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
Dim mMesIni As Integer
Dim mMesFin As Integer

Dim mColDebe As Integer '--posicion de la columna debe
Dim mColHaber As Integer '--posicion de la columna haber
Dim mColSaldo As Integer '--posicion de la columna  saldo

Dim mPosRegistro As Integer '--indica la posicion del numero de registro


Private Sub chk_Click()
    If Fg2.Rows = 1 Then Exit Sub
    Fg2.Rows = 1
End Sub

Private Function fValidarSeleccionCta(NumCuenta As String) As Boolean
    On Error GoTo error
    
    Dim k As Integer
    Dim MSG_CUENTA As String    '--MUSTRA EL MENSAJE SI DESEA AGREGAR UNA CUENTA, CUANDO YA EXISTE UNA CUENTA DE NIVEL SUPERIOR O NIVEL INFERIOR
                                '--NO MOSTRAR MENSAJE SOLO CUANDO LAS CUENTAS SEA DEL MISMO NIVEL
    If GRID_BUSCAR_VALOR(Fg2, 1, NumCuenta, False, , Fg2.Row) <> "-1" Then
        MsgBox "La cuenta contable Nº " & NumCuenta & " ya fue seleccionada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If
    
    For k = 1 To Fg2.Rows - 1
        If k <> Fg2.Row Then
            If Len(NumCuenta) < Len(Trim(Fg2.TextMatrix(k, 1))) Then
                
                If NumCuenta = Mid(Trim(Fg2.TextMatrix(k, 1)), 1, Len(NumCuenta)) Then
                    MSG_CUENTA = "Ya agregó la cuenta Nº: " + Trim(Fg2.TextMatrix(k, 1)) + " cuyo nivel es Inferior a la cuenta Nº: " & NumCuenta & " que desea agregar" _
                                + vbCr + "Sólo puede agregar Cuentas del mismo nivel " _
                                + vbCr + "Si desea continuar elimine la fila que contenga la Cuenta Nº: " + Trim(Fg2.TextMatrix(k, 1))
                    Exit For
                End If
                
            Else
                If Trim(Fg2.TextMatrix(k, 1)) = Mid(NumCuenta, 1, Len(Trim(Fg2.TextMatrix(k, 1)))) Then
                    MSG_CUENTA = "Ya agregó la cuenta Nº: " + Trim(Fg2.TextMatrix(k, 1)) + " cuyo nivel es Superior a la cuenta Nº: " & NumCuenta & " que desea agregar" _
                                + vbCr + "Sólo puede agregar Cuentas del mismo nivel " _
                                + vbCr + "Si desea continuar elimine la fila que contenga la Cuenta Nº: " + Trim(Fg2.TextMatrix(k, 1))
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
    SHOW_ERROR Me.Name, "fValidarSeleccionCta"
End Function

Private Sub CmdAdd_Click()
    If Fg2.Rows = 1 Then
        Fg2.Rows = Fg2.Rows + 1
        Exit Sub
    End If
    If NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 3)) = 0 Then Exit Sub
    Fg2.Rows = Fg2.Rows + 1
    Fg2.Row = Fg2.Rows - 1
    Fg2.Col = 1
    Fg2.SetFocus
End Sub

Private Sub CmdDel_Click()
    If Fg2.Row <= 0 Then Exit Sub
    If Fg2.Rows <= 1 Then
        MsgBox "No hay cuentas seleccionadas para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    Else
        Fg2.RemoveItem Fg2.Row
        Fg2.Refresh
    End If
End Sub

'Private Sub pImprimir()
'
'    On Error GoTo error
'
'    If fValidarConsulta() = False Then Exit Sub
'
'
'    Me.MousePointer = vbHourglass
'    If Me.TabOne1.CurrTab = 0 Then
'        FrmPrintMayor.Show
'    Else
'        Dim X_PRINT As New SGI2_funciones.formularios
'        Dim xMoneda As String
'        Dim nPeriodo  As String
'        If opt_fecha(0).Value = True Then  '--por fecha
'            If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
'                nPeriodo = "Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
'            Else
'                nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
'            End If
'        Else '--por periodo
'            If mMesIni = mMesFin Then
'                nPeriodo = "Periodo : " + lbl_periodo(0).Caption
'            Else
'                nPeriodo = "Periodo : De " + lbl_periodo(0).Caption & " A " & lbl_periodo(1).Caption
'            End If
'        End If
'        If NulosN(TxtIdMon.Text) = 1 Then
'            xMoneda = "Nuevos Soles"
'        Else
'            xMoneda = "Dolares Americanos"
'        End If
'        X_PRINT.Imprimir_x_VSFlexGrid Fg3, "LIBRO MAYOR ", "(Expresado en " + xMoneda + ")", nPeriodo, False, True
'        Set X_PRINT = Nothing
'
'    End If
'    Me.MousePointer = vbDefault
'    Exit Sub
'error:
'    Me.MousePointer = vbDefault
'    SHOW_ERROR Me.Name, "CmdImprimir_Click"
'End Sub

Private Sub pImprimir()
    Dim xMoneda As String
    Dim nPeriodo As String
    
    If opt_fecha(0).Value = True Then  '--por fecha
        If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
            nPeriodo = "Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
        Else
            nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
        End If
    Else '--por periodo
        If mMesIni = mMesFin Then
            nPeriodo = "Periodo : " + lbl_periodo(0).Caption
        Else
            nPeriodo = "Periodo : De " + lbl_periodo(0).Caption & " A " & lbl_periodo(1).Caption
        End If
    End If
    
    If NulosN(TxtIdMon.Text) = 1 Then
        xMoneda = "Nuevos Soles"
    Else
        xMoneda = "Dolares Americanos"
    End If

    Dim RstTmp As New ADODB.Recordset
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 5) And ((con_formatostipodet.idformatotipo) = 2) " _
        & " And ((con_formatostipodet.mostrar) = -1)) ORDER BY con_formatostipodet.orden", xCon
    
    Dim xCampos() As String
    Dim xFil, xCol As Double
    
    ReDim xCampos(Fg1.Rows - 2, Fg1.Cols - 1)
    
    Dim xFila As Double
    xFila = 0
    For xFil = 1 To Fg1.Rows - 1
        For xCol = 1 To Fg1.Cols - 1
            xCampos(xFila, xCol) = Fg1.TextMatrix(xFil, xCol)
        Next xCol
        xFila = xFila + 1
    Next xFil
    
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
    
    Dim xfrm As New eps_librerias.IMPRIMIR
    
    xfrm.Cabecera1 = NomEmp
    xfrm.Cabecera2 = "RUC Nº: " & NumRUC
    xfrm.Fecha = Format(Date, "dd/mm/yyyy")
    xfrm.Titulo1 = "LIBRO MAYOR " & "(Expresado en " & xMoneda & ")"
    xfrm.Titulo2 = nPeriodo
    xfrm.TamañoFuente = 6
    xfrm.TamañoCabecera = 8
    xfrm.FuenteCabecera = "Courier New"
    xfrm.Posicion_Hoja = Vertical
    xfrm.Tamaño_Hoja = A_4
    xfrm.TextoConsiderar = "CTA"
    xfrm.TextoConsiderarAncho = 3
    xfrm.ImprimirArray xCampos, Rst
    Set xfrm = Nothing
End Sub

Private Sub pConsultar()
    If fValidarConsulta() = False Then Exit Sub

    BAND_INTERRUMPIR = False
    Me.ProgressBar1.Value = 1
    Me.TabOne1.CurrTab = 0
    pConfigurarGrilla True
    DoEvents
    
    If MuestraMayor1() = False Then Exit Sub
    
    If BAND_INTERRUMPIR = True Then Exit Sub
    
    DoEvents
    If ChkResumen.Value = 1 Then
        Me.TabOne1.CurrTab = 1
        CargarResumen1
    Else
        Me.TabOne1.CurrTab = 0
    End If
    
End Sub

Sub CargarResumen()
    On Error GoTo error
    Dim RstRes As New ADODB.Recordset
    Dim A&
    Dim xTotal1, xTotal2, xTotal3, xTotal4 As Double
    Dim xAcumulado(7) As Double
   
    Frame5.Left = 3413
    Frame5.Top = 2685
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Value = 0
    
    Frame5.Visible = True
    Label3.Caption = "Procesando Resumen"
    Fg3.Rows = Fg3.FixedRows
    
    DoEvents
    
    Dim nSQLCuenta As String
    Dim nSQL As String
    '--------------------------
    nSQLCuenta = ""
    '--SI AGREGA CUENTAS AS GRID, GENERAR EL FILTRO A CONCATENAR A LA CONSULTA
    For A = 1 To Fg2.Rows - 1
        If Trim(Fg2.TextMatrix(A, 1)) <> "" Then
            nSQLCuenta = nSQLCuenta + " con_planctas.cuenta Like '" & Trim(Fg2.TextMatrix(A, 1)) & "%' OR "
        End If
    Next A
    If nSQLCuenta <> "" Then nSQLCuenta = " WHERE (" + Left(nSQLCuenta, Len(nSQLCuenta) - 3) + ") "
    '--------------------------

    nSQL = "SELECT con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion AS descri , con_planctas.tipsal , "
    
    If NulosN(TxtIdMon.Text) = 1 Then
        nSQL = nSQL _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol=0,con_diario1.impdebsol,IIf(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,iif(con_tc1.impven is null,0,con_tc1.impven))=0,0,((iif(con_diario1.idlib in (3,6,44), con_diario1.tc,con_tc1.impven))*con_diario1.impdebdol)))) AS saldebesol  FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue  WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING con_diario1.idcue = con_diario.idcue ) AS saldebesol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol=0,con_diario1.imphabsol,IIf(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,iif(con_tc1.impven is null,0,con_tc1.impven))=0,0,((iif(con_diario1.idlib in (3,6,44), con_diario1.tc,con_tc1.impven))*con_diario1.imphabdol)))) AS salhabersol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING con_diario1.idcue = con_diario.idcue ) AS salhabersol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol=0,con_diario1.impdebsol,IIf(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,iif(con_tc1.impven is null,0,con_tc1.impven))=0,0,((iif(con_diario1.idlib in (3,6,44), con_diario1.tc,con_tc1.impven))*con_diario1.impdebdol))))  AS debesol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07') GROUP BY con_diario1.idcue  HAVING con_diario1.idcue=con_diario.idcue  ) AS  debesol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol=0,con_diario1.imphabsol,IIf(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,iif(con_tc1.impven is null,0,con_tc1.impven))=0,0,((iif(con_diario1.idlib in (3,6,44), con_diario1.tc,con_tc1.impven))*con_diario1.imphabdol))))  AS habersol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')  GROUP BY con_diario1.idcue  HAVING con_diario1.idcue=con_diario.idcue  ) AS  habersol, " _
            + vbCr + " IIf(debesol Is Null,0+IIf(saldebesol Is Null,0,saldebesol),debesol+IIf(saldebesol Is Null,0,saldebesol)) AS maydebesol, " _
            + vbCr + " IIf(habersol Is Null,0+IIf(salhabersol Is Null,0,salhabersol),habersol+IIf(salhabersol Is Null,0,salhabersol)) AS mayhabersol, " _
            + vbCr + " (IIF (con_planctas.tipsal='D' OR con_planctas.tipsal IS NULL OR con_planctas.tipsal ='', (maydebesol -  mayhabersol), (mayhabersol - maydebesol))) as saldosol, " _
            + vbCr + " IIf(maydebesol>mayhabersol,(maydebesol-mayhabersol),0) AS deudorsol, " _
            + vbCr + " IIf(mayhabersol>maydebesol,(mayhabersol-maydebesol),0) AS acreedorsol "
    Else
        nSQL = nSQL _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol<>0,con_diario1.impdebdol,IIf(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,iif(con_tc1.impven is null,0,con_tc1.impven))=0 Or con_diario1.impdebsol=0,0,(con_diario1.impdebsol/(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,con_tc1.impven)))))) AS saldebedol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING (((con_diario1.idcue)=con_diario.idcue))) AS saldebedol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol<>0,con_diario1.imphabdol,IIf(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,iif(con_tc1.impven is null,0,con_tc1.impven))=0 Or con_diario1.imphabsol=0,0,(con_diario1.imphabsol/(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,con_tc1.impven)))))) AS salhaberdol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING (((con_diario1.idcue)=con_diario.idcue))) AS salhaberdol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol<>0,con_diario1.impdebdol,IIf(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,iif(con_tc1.impven is null,0,con_tc1.impven))=0 Or con_diario1.impdebsol=0,0,(con_diario1.impdebsol/(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,con_tc1.impven)))))) AS debedol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')  GROUP BY con_diario1.idcue HAVING con_diario1.idcue=con_diario.idcue ) AS  debedol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol<>0,con_diario1.imphabdol,IIf(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,iif(con_tc1.impven is null,0,con_tc1.impven))=0 Or con_diario1.imphabsol=0,0,(con_diario1.imphabsol/(iif(con_diario1.idlib in (3,6,44), con_diario1.tc,con_tc1.impven)))))) AS haberdol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')  GROUP BY con_diario1.idcue  HAVING con_diario1.idcue=con_diario.idcue ) AS  haberdol, " _
            + vbCr + " IIf(debedol Is Null,0+IIf(saldebedol Is Null,0,saldebedol),debedol+IIf(saldebedol Is Null,0,saldebedol)) AS maydebedol, " _
            + vbCr + " IIf(haberdol Is Null,0+IIf(salhaberdol Is Null,0,salhaberdol),haberdol+IIf(salhaberdol Is Null,0,salhaberdol)) AS mayhaberdol, " _
            + vbCr + " (IIF (con_planctas.tipsal='D' OR con_planctas.tipsal IS NULL OR con_planctas.tipsal ='', (maydebedol -  mayhaberdol), (mayhaberdol - maydebedol))) as saldodol, " _
            + vbCr + " IIf(maydebedol>mayhaberdol,(maydebedol-mayhaberdol),0) AS deudordol, " _
            + vbCr + " IIf(mayhaberdol > maydebedol, (mayhaberdol - maydebedol), 0) As acreedordol "
    End If
        
    nSQL = nSQL _
        + vbCr + " FROM con_planctas INNER JOIN con_diario ON con_planctas.id = con_diario.idcue " _
        + nSQLCuenta _
        + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion,con_planctas.tipsal " _
        + vbCr + " ORDER BY con_planctas.cuenta, con_planctas.descripcion;"

    '--UNIFICANDO LOS NOMBRES DE LOS CAMPOS TANTO PARA DOLARES Y SOLES
    nSQL = Replace(nSQL, "saldebesol", "saldeb")
    nSQL = Replace(nSQL, "salhabersol", "salhab")
    nSQL = Replace(nSQL, "maydebesol", "maydeb")
    nSQL = Replace(nSQL, "mayhabersol", "mayhab")
    nSQL = Replace(nSQL, "debesol", "movdeb")
    nSQL = Replace(nSQL, "habersol", "movhab")
    nSQL = Replace(nSQL, "deudorsol", "deudor")
    nSQL = Replace(nSQL, "acreedorsol", "acreedor")
    

    nSQL = Replace(nSQL, "saldebedol", "saldeb")
    nSQL = Replace(nSQL, "salhaberdol", "salhab")
    nSQL = Replace(nSQL, "maydebedol", "maydeb")
    nSQL = Replace(nSQL, "mayhaberdol", "mayhab")
    nSQL = Replace(nSQL, "debedol", "movdeb")
    nSQL = Replace(nSQL, "haberdol", "movhab")
    nSQL = Replace(nSQL, "deudordol", "deudor")
    nSQL = Replace(nSQL, "acreedordol", "acreedor")
    
    If opt_fecha(0).Value = True Then '--por fecha
        '--REEMPLAZANDO EL INTERVALO DE FECHA
        nSQL = Replace(nSQL, "con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')", " ( con_diario1.fchasi >=CDate('" + Me.TxtFchIni.Valor + "') And con_diario1.fchasi <= CDate('" + Me.TxtFchFin.Valor + "') ) ")
        '--REEMPLAZANDO LA FECHA DE INICIO PARA OBTENER LOS SALDOS
        nSQL = Replace(nSQL, "con_diario1.fchasi < CDate('01/01/07')", " ( con_diario1.fchasi < CDate('" + Me.TxtFchIni.Valor + "') ) ")
    Else '--por periodo
        '--REEMPLAZANDO EL INTERVALO DE FECHA
        nSQL = Replace(nSQL, "con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')", " ( con_diario1.año = " & AnoTra & " and con_diario1.idmes >= " & mMesIni & " and con_diario1.idmes <= " & mMesFin & " ) ")
        '--REEMPLAZANDO LA FECHA DE INICIO PARA OBTENER LOS SALDOS
        nSQL = Replace(nSQL, "con_diario1.fchasi < CDate('01/01/07')", " ( con_diario1.año = " & AnoTra & " and con_diario1.idmes < " & mMesIni & " ) ")
    End If
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo SALIR:
    RstTmp.Filter = adFilterNone
    RstTmp.Sort = "cuenta"
    fra_msg.Visible = True
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    If RstTmp.RecordCount > 0 Then ProgressBar1.Max = RstTmp.RecordCount
    
    Label3.Caption = "Procesando Resumen"
    For A = 1 To RstTmp.RecordCount
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        '-----------------------------------------------
        ProgressBar1.Value = A
        Fg3.Rows = Fg3.Rows + 1
        Fg3.TextMatrix(A + 1, 1) = RstTmp("cuenta") & ""
        
        Fg3.TextMatrix(A + 1, 2) = RstTmp("descri") & ""
        
        Fg3.TextMatrix(A + 1, 3) = Format(NulosN(RstTmp("saldeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 4) = Format(NulosN(RstTmp("salhab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 5) = Format(NulosN(RstTmp("movdeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 6) = Format(NulosN(RstTmp("movhab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 7) = Format(NulosN(RstTmp("maydeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 8) = Format(NulosN(RstTmp("mayhab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 9) = Format(NulosN(RstTmp("deudor")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 10) = Format(NulosN(RstTmp("acreedor")), FORMAT_MONTO)
        
        
        xAcumulado(0) = xAcumulado(0) + NulosN(Fg3.TextMatrix(A + 1, 3)) '--saldeb
        xAcumulado(1) = xAcumulado(1) + NulosN(Fg3.TextMatrix(A + 1, 4)) '--salhab
        xAcumulado(2) = xAcumulado(2) + NulosN(Fg3.TextMatrix(A + 1, 5)) '--movdeb
        xAcumulado(3) = xAcumulado(3) + NulosN(Fg3.TextMatrix(A + 1, 6)) '--movhab
        xAcumulado(4) = xAcumulado(4) + NulosN(Fg3.TextMatrix(A + 1, 7)) '--maydeb
        xAcumulado(5) = xAcumulado(5) + NulosN(Fg3.TextMatrix(A + 1, 8)) '--mayhab
        xAcumulado(6) = xAcumulado(6) + NulosN(Fg3.TextMatrix(A + 1, 9)) '--deudor
        xAcumulado(7) = xAcumulado(7) + NulosN(Fg3.TextMatrix(A + 1, 10)) '--acreedor
        
        RstTmp.MoveNext
        If RstTmp.EOF = True Then Exit For
    Next A
    
    Fg3.Rows = Fg3.Rows + 2
    
    Fg3.TextMatrix(Fg3.Rows - 1, 2) = "TOTAL =>"
    '-------------------------------
    Dim Col&
    For A = 0 To UBound(xAcumulado())
        Fg3.TextMatrix(Fg3.Rows - 1, 3 + A) = Format(xAcumulado(A), FORMAT_MONTO)
        FORMATO_CELDA Fg3, Fg3.Rows - 1, 3 + A, , True
    Next A
    Erase xAcumulado()
    '-------------------------------
    If RstTmp.RecordCount > 0 Then
        GRID_COLOR_FONDO Fg3, 2, 3, Fg3.Rows - 3, 4, RGB(255, 255, 236)
        GRID_COLOR_FONDO Fg3, 2, 7, Fg3.Rows - 3, 8, RGB(255, 255, 236)
        
    End If
        GRID_COLOR_FONDO Fg3, Fg3.Rows - 2, 1, Fg3.Rows - 1, Fg3.Cols - 1, RGB(231, 254, 224)
    
SALIR:
    Set RstRes = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    
    MsgBox "El Mayor se terminó de procesar con éxito", vbInformation, xTitulo
    
    Exit Sub
error:
    Frame5.Visible = False
    Set RstRes = Nothing
    fra_msg.Visible = False
    SHOW_ERROR Me.Name, "CargarResumen"
End Sub

Private Sub pExportar()
    If TabOne1.CurrTab = 0 Then
        If Fg1.Rows = 1 Then
            MsgBox "No hay registros para exportar", vbExclamation, xTitulo
            Exit Sub
        End If
    ElseIf TabOne1.CurrTab = 1 Then
        If Fg3.Rows = 1 Then
            MsgBox "No hay registros para exportar", vbExclamation, xTitulo
            Exit Sub
        End If
    End If
    
    If fValidarConsulta() = False Then Exit Sub
    
    Dim nTitulo As String
    Dim nTitulo1 As String
    Dim nPeriodo As String
    
    If opt_fecha(0).Value = True Then
        nPeriodo = "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor
    Else
        nPeriodo = "DE " + lbl_periodo(0).Caption + " A " + lbl_periodo(1).Caption
    End If
    If TabOne1.CurrTab = 0 Then
        
'''        Dim xFun As New SGI2_funciones.formularios
'''        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "DETALLE DEL MAYOR", nPeriodo, "Expresado en " & LblMoneda.Caption, "Mayor - Detalle"         ', Rst, ""
'''        Set xFun = Nothing
        
        GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "DETALLE DEL MAYOR", nPeriodo, "Expresado en " & LblMoneda.Caption
        
    Else
        ExportarExcelResumen
    End If
    
    
End Sub


Private Sub Fg1_DblClick()
    '--mostrar el asiento
    If Fg1.Rows <= Fg1.FixedRows Then Exit Sub
    If mPosRegistro = 0 Then Exit Sub
    If mPosRegistro = 0 Then Exit Sub
    Dim xfrm As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    xfrm.AsientoVer xCon, Fg1.TextMatrix(Fg1.Row, mPosRegistro)
    Set xfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
    
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
        
        lbl_periodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
        lbl_periodo(1).Caption = lbl_periodo(0).Caption
        mMesIni = xMes
        mMesFin = xMes

        TabOne1.CurrTab = 0

        TxtFchIni.SetFocus
        
        GRID_COMBOLIST Fg2, 1
        SeEjecuto = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    ElseIf KeyCode = vbKeyF3 And Shift = 0 Then
        BuscarVSFlexGrid
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Fg2.Rows = 1
    Fg1.Rows = 1
    Fg3.Rows = 1
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    Fg3.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDNone
    Fg2.Editable = flexEDNone
    Fg3.Editable = flexEDNone
    
    Fg2.Tag = Fg2.FormatString
    
    Fg2.ColWidth(3) = 0
    Frame2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    pConfigurarGrilla True
    
    TabOne1.CurrTab = 0
    
    SetearCuadricula Fg1, 5, xCon, 1, 0, False

  
End Sub


Private Function MuestraMayor() As Boolean
    Dim RstMay As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstSal As New ADODB.Recordset
    
    Dim A&, B&, C&
    Dim nSQL As String
'    On Error GoTo error

    Frame5.Left = 3413
    Frame5.Top = 2685
    Frame5.Visible = True
    ProgressBar1.Min = 1
    DoEvents
    
    Dim nSQLCuenta As String
    nSQLCuenta = ""
    '--SI AGREGA CUENTAS AS GRID, GENERAR EL FILTRO A CONCATENAR A LA CONSULTA
    For A = 1 To Fg2.Rows - 1
        If Trim(Fg2.TextMatrix(A, 1)) <> "" Then
            nSQLCuenta = nSQLCuenta + " con_planctas.cuenta Like '" & Trim(Fg2.TextMatrix(A, 1)) & "%' OR "
        End If
    Next A
    If nSQLCuenta <> "" Then nSQLCuenta = " AND (" + Left(nSQLCuenta, Len(nSQLCuenta) - 3) + ") "
    '---------
    '--ESTABLECER EL CAMPO A TOTALIZAR EN FUNCION DEL RECORDSET TMP (RstTmp2) , TANTO A SOLES Y DOLARES
    Dim CAMPO_DEBE, CAMPO_HABER  As String
    If NulosN(TxtIdMon.Text) = 1 Then
        CAMPO_DEBE = "impdebsol":  CAMPO_HABER = "imphabsol"
    Else
        CAMPO_DEBE = "impdebdol":  CAMPO_HABER = "imphabdol"
    End If
    '-----------------------------------------------------------------------
    '--SI SE NTERRUMPE EL PROCESO => SALIR
     If BAND_INTERRUMPIR = True Then GoTo SALIR:
     '-----------------------------------------------
    Set RstTmp2 = Nothing
    'nSQL = "SELECT con_diario.idcue AS id,  con_diario.idcue as idcuenta, con_planctas.cuenta, con_planctas.descripcion AS descri, con_planctas.tipsal, con_diario.idlib, con_diario.idmov, Format([con_diario]![idmes],'00') & IIf(mae_libros.codsun Is Null OR mae_libros.codsun ='','FF',Format([mae_libros].[codsun],'00')) & Trim([con_diario]![numasi]) AS numreg, mae_libros.descripcion AS nomlib, con_tc.impven, " _
        + vbCr + " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1, mae_documento.abrev,IIf(con_diario.idlib=2, mae_documento_1.abrev,IIf(con_diario.idlib=3, mae_documento_2.abrev,IIf(con_diario.idlib=4, mae_documento_3.abrev,IIf(con_diario.idlib=5, mae_documento_4.abrev,IIf(con_diario.idlib=6, mae_doccajaban.abrev,IIf(con_diario.idlib=8,'CAN',IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,'CAN',IIf(con_diario.idlib=38,mae_doccajaban_1.abrev,IIf(con_diario.idlib=39,mae_documento_5.abrev,'OTROS LIBROS')))))))))))) AS tipdoc, " _
        + vbCr + " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.fchdoc,IIf(con_diario.idlib=2,vta_ventas.fchdoc,IIf(con_diario.idlib=3,con_proviciones.fchdoc,IIf(con_diario.idlib=4,con_percepcion.fchdoc,IIf(con_diario.idlib=5,con_retencion.fchemi,IIf(con_diario.idlib=6,con_cajabanco.fchope,IIf(con_diario.idlib=8,con_canjes.fchemi,IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,con_letra.fchemi,IIf(con_diario.idlib=38,con_ctasrendir.fchemi,IIf(con_diario.idlib=39,con_devoluciones.fchemi,'OTROS LIBROS')))))))))))) AS fchemi, " _
        + vbCr + " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.numser & '-' & com_compras.numdoc,IIf(con_diario.idlib=2,vta_ventas.numser & '-' & vta_ventas.numdoc,IIf(con_diario.idlib=3,con_proviciones.numser & '-' & con_proviciones.numdoc,IIf(con_diario.idlib=4,con_percepcion.numser & '-' & con_percepcion.numdoc,IIf(con_diario.idlib=5,con_retencion.numser & '-' & con_retencion.numdoc,IIf(con_diario.idlib=6,con_cajabanco.numdoc,IIf(con_diario.idlib=8,[con_canjes].[numser] & '-' & [con_canjes].[numdoc],IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,' ',IIf(con_diario.idlib=38,con_ctasrendir.numdoc,IIf(con_diario.idlib=39,con_devoluciones.numdoc,'OTROS LIBROS')))))))))))) AS numdoc, " _
        + vbCr + " IIf(con_diario.impdebdol<>0,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebsol, " _
        + vbCr + " IIf(con_diario.imphabdol<>0,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabsol, " _
        + vbCr + " IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_diario].[impdebsol]=0,0,([con_diario].[impdebsol]/[con_tc].[impven]))) AS impdebdol, " _
        + vbCr + " IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol],IIf([con_tc].[impven] Is Null Or [con_diario].[imphabsol]=0,0,([con_diario].[imphabsol]/[con_tc].[impven]))) AS imphabdol, " _
        + vbCr + " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebsol-imphabsol),(imphabsol-impdebsol))) AS saldosol, " _
        + vbCr + " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebdol-imphabdol),(imphabdol-impdebdol))) AS saldodol " _
        + vbCr + " FROM (((((mae_libros RIGHT JOIN (con_planctas RIGHT JOIN ((((((((((((((con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN com_compras ON con_diario.idmov = com_compras.id) LEFT JOIN vta_ventas ON con_diario.idmov = vta_ventas.id) LEFT JOIN con_retencion ON con_diario.idmov = con_retencion.id) LEFT JOIN con_percepcion ON con_diario.idmov = con_percepcion.id) LEFT JOIN con_proviciones ON con_diario.idmov = con_proviciones.id) LEFT JOIN con_cajabanco ON con_diario.idmov = con_cajabanco.id) LEFT JOIN con_canjes ON con_diario.idmov = con_canjes.id) LEFT JOIN mae_doccajaban ON con_cajabanco.iddoc = mae_doccajaban.id) LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_documento AS mae_documento_1 ON vta_ventas.tipdoc = mae_documento_1.id) LEFT JOIN mae_documento AS mae_documento_2 ON con_proviciones.tipdoc = mae_documento_2.id)  " _
        + vbCr + "      LEFT JOIN mae_documento AS mae_documento_3 ON con_percepcion.tipdoc = mae_documento_3.id) LEFT JOIN mae_documento AS mae_documento_4 ON con_retencion.iddoc = mae_documento_4.id) ON con_planctas.id = con_diario.idcue) ON mae_libros.id = con_diario.idlib) LEFT JOIN con_letra ON con_diario.idmov = con_letra.id) LEFT JOIN con_ctasrendir ON con_diario.idmov = con_ctasrendir.id) LEFT JOIN con_devoluciones ON con_diario.idmov = con_devoluciones.id) LEFT JOIN mae_doccajaban AS mae_doccajaban_1 ON con_ctasrendir.tipdoc = mae_doccajaban_1.id) LEFT JOIN mae_documento AS mae_documento_5 ON con_devoluciones.iddoc = mae_documento_5.id "
        
    nSQL = "SELECT con_diario.idcue AS id, con_diario.idcue AS idcuenta, con_planctas.cuenta, con_planctas.descripcion AS descri, con_planctas.tipsal, con_diario.idlib, con_diario.idmov, Format([con_diario]![idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',Format([mae_libros].[codsun],'00')) & Trim([con_diario]![numasi]) AS numreg, IIf([con_diario].[idlib]<>3,[mae_libros].[descripcion],[mae_librossub].[descripcion]) AS nomlib, con_tc.impven, " _
        & vbCr & " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,mae_documento.abrev,IIf(con_diario.idlib=2,mae_documento_1.abrev,IIf(con_diario.idlib=3,mae_documento_2.abrev,IIf(con_diario.idlib=4,mae_documento_3.abrev,IIf(con_diario.idlib=5,mae_documento_4.abrev,IIf(con_diario.idlib=6,mae_doccajaban.abrev,IIf(con_diario.idlib=8,'CAN',IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,'CAN',IIf(con_diario.idlib=38,mae_doccajaban_1.abrev,IIf(con_diario.idlib=39,mae_documento_5.abrev,'OTROS LIBROS')))))))))))) AS tipdoc, " _
        & vbCr & " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.fchdoc,IIf(con_diario.idlib=2,vta_ventas.fchdoc,IIf(con_diario.idlib=3,con_proviciones.fchdoc,IIf(con_diario.idlib=4,con_percepcion.fchdoc,IIf(con_diario.idlib=5,con_retencion.fchemi,IIf(con_diario.idlib=6,con_cajabanco.fchope,IIf(con_diario.idlib=8,con_canjes.fchemi,IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,con_letra.fchemi,IIf(con_diario.idlib=38,con_ctasrendir.fchemi,IIf(con_diario.idlib=39,con_devoluciones.fchemi,'OTROS LIBROS')))))))))))) AS fchemi, " _
        & vbCr & " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.numser & '-' & com_compras.numdoc,IIf(con_diario.idlib=2,vta_ventas.numser & '-' & vta_ventas.numdoc,IIf(con_diario.idlib=3,con_proviciones.numser & '-' & con_proviciones.numdoc,IIf(con_diario.idlib=4,con_percepcion.numser & '-' & con_percepcion.numdoc,IIf(con_diario.idlib=5,con_retencion.numser & '-' & con_retencion.numdoc,IIf(con_diario.idlib=6,con_cajabanco.numdoc,IIf(con_diario.idlib=8,[con_canjes].[numser] & '-' & [con_canjes].[numdoc],IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,' ',IIf(con_diario.idlib=38,con_ctasrendir.numdoc,IIf(con_diario.idlib=39,con_devoluciones.numdoc,'OTROS LIBROS')))))))))))) AS numdoc, " _
        & vbCr & " IIf(con_diario.impdebdol<>0,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebsol, " _
        & vbCr & " IIf(con_diario.imphabdol<>0,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabsol, " _
        & vbCr & " IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_diario].[impdebsol]=0,0,([con_diario].[impdebsol]/[con_tc].[impven]))) AS impdebdol, " _
        & vbCr & " IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol],IIf([con_tc].[impven] Is Null Or [con_diario].[imphabsol]=0,0,([con_diario].[imphabsol]/[con_tc].[impven]))) AS imphabdol, " _
        & vbCr & " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebsol-imphabsol),(imphabsol-impdebsol))) AS saldosol, " _
        & vbCr & " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebdol-imphabdol),(imphabdol-impdebdol))) AS saldodol, " _
        & vbCr & " IIf([con_diario].[idlib]=2,[mae_cliente]![nombre],IIf([con_diario].[idlib]=1,[mae_prov]![nombre],'')) AS nomclipro " _
        & vbCr & " FROM (mae_prov RIGHT JOIN (mae_libros RIGHT JOIN (mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN ((((((((((((((con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN com_compras ON con_diario.idmov = com_compras.id) LEFT JOIN vta_ventas ON con_diario.idmov = vta_ventas.id) LEFT JOIN con_retencion ON con_diario.idmov = con_retencion.id) LEFT JOIN con_percepcion ON con_diario.idmov = con_percepcion.id) LEFT JOIN con_proviciones ON con_diario.idmov = con_proviciones.id) LEFT JOIN con_cajabanco ON con_diario.idmov = con_cajabanco.id) LEFT JOIN con_canjes ON con_diario.idmov = con_canjes.id) LEFT JOIN mae_doccajaban ON con_cajabanco.iddoc = mae_doccajaban.id) LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_documento AS mae_documento_1 ON vta_ventas.tipdoc = mae_documento_1.id) LEFT JOIN mae_documento AS mae_documento_2 ON con_proviciones.tipdoc = mae_documento_2.id)  " _
        & vbCr & " LEFT JOIN mae_documento AS mae_documento_3 ON con_percepcion.tipdoc = mae_documento_3.id) LEFT JOIN mae_documento AS mae_documento_4 ON con_retencion.iddoc = mae_documento_4.id) ON con_planctas.id = con_diario.idcue) LEFT JOIN con_letra ON con_diario.idmov = con_letra.id) LEFT JOIN con_ctasrendir ON con_diario.idmov = con_ctasrendir.id) LEFT JOIN con_devoluciones ON con_diario.idmov = con_devoluciones.id) LEFT JOIN mae_doccajaban AS mae_doccajaban_1 ON con_ctasrendir.tipdoc = mae_doccajaban_1.id) LEFT JOIN mae_documento AS mae_documento_5 ON con_devoluciones.iddoc = mae_documento_5.id) ON mae_cliente.id = vta_ventas.idcli) ON mae_libros.id = con_diario.idlib) ON mae_prov.id = com_compras.idpro) LEFT JOIN mae_librossub ON (con_proviciones.idlib = mae_librossub.idlib) AND (con_proviciones.idsublib = mae_librossub.id) "

'WHERE (((con_planctas.cuenta) Like '42*') AND ((con_diario.fchasi)>=CDate('01/01/2007') And (con_diario.fchasi)<=CDate('31/12/2007') And (con_diario.fchasi)>=CDate('01/01/2007') And (con_diario.fchasi)<=CDate('31/12/2007')))
'ORDER BY con_planctas.cuenta;

   
    If opt_fecha(0).Value = True Then
        nSQL = nSQL + vbCr + " WHERE (con_diario.fchasi >=CDate('" + TxtFchIni.Valor + "') And con_diario.fchasi<=CDate('" + TxtFchFin.Valor + "')) " _
            + vbCr + " AND ( con_diario.fchasi >=CDate('01/01/" + AnoTra + "') And con_diario.fchasi <= CDate('31/12/" + AnoTra + "') ) "
    Else
        nSQL = nSQL + vbCr + " WHERE ( con_diario.idmes >= " & mMesIni & " and con_diario.idmes <= " & mMesFin & " ) and con_diario.año = " & AnoTra & " "
    End If
        
    nSQL = nSQL + nSQLCuenta + vbCr + " ORDER BY con_planctas.cuenta ASC "


    RST_Busq RstTmp2, nSQL, xCon

    'HACEMOS UNA CONSULTA DE LOS REGISTROS UNICOS DE LA CONSULTA ANTERIOR, PARA PODER TOTALIZARLA

    Fg1.Rows = 1
    Dim xFila&
    Dim xSaldo As Double
    Dim xTotal1, xTotal2 As Double
    Dim xTotal1_1, xTotal2_1 As Double
    
    xFila = 1

    DoEvents
    '--SI SE NTERRUMPE EL PROCESO => SALIR
     If BAND_INTERRUMPIR = True Then GoTo SALIR:
    '-----------------------------------------------
    nSQL = "SELECT con_diario.idcue as idcuenta, con_planctas.cuenta, con_planctas.descripcion, Sum(con_diario.impdebsol) AS SumaDeimpdebsol, " _
         + vbCr + " Sum(con_diario.imphabsol) AS SumaDeimphabsol, Sum(con_diario.impdebdol) AS SumaDeimpdebdol, Sum(con_diario.imphabdol) AS SumaDeimphabdol " _
         + vbCr + " FROM con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue "
                  
    If opt_fecha(0).Value = True Then
         nSQL = nSQL + vbCr + " WHERE (con_diario.fchasi >=CDate('" & TxtFchIni.Valor & "') And con_diario.fchasi <=CDate('" & TxtFchFin.Valor & "')) " _
         + vbCr + " AND ( con_diario.fchasi >=CDate('01/01/" + AnoTra + "') And con_diario.fchasi <= CDate('31/12/" + AnoTra + "') ) "
    Else
        nSQL = nSQL + vbCr + " WHERE ( con_diario.idmes >= " & mMesIni & " and con_diario.idmes <= " & mMesFin & " )and con_diario.año = " & AnoTra & " "
    End If
         
    nSQL = nSQL + vbCr + nSQLCuenta _
         + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion ORDER BY con_planctas.cuenta ASC "
    
    RST_Busq RstMay, nSQL, xCon

    xSaldo = 0
    '---
    If RstMay.RecordCount <> 0 Then
    
        fra_msg.Visible = True
        
        RstMay.MoveFirst
        
        If RstMay.RecordCount > 1 Then ProgressBar1.Max = RstMay.RecordCount
        
        DoEvents
        Do While Not RstMay.EOF
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo SALIR:
            '-----------------------------------------------
            ProgressBar1.Value = RstMay.Bookmark
    
            xSaldo = 0
            Fg1.Rows = Fg1.Rows + 1
            
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 7, "Cta Nº  :  " + NulosC(RstMay("cuenta")) & "   - " + NulosC(RstMay.Fields("descripcion")), flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 1, , True
                        
            xTotal1 = 0
            xTotal2 = 0
            
           'hallamos el saldo anterior de la cuenta
            Set RstSal = Nothing
                            
            nSQL = "SELECT con_diario.idcue as idcuenta, con_planctas.cuenta,con_planctas.tipsal, " _
                + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[impdebdol]*[con_tc].[impven]),[con_diario].[impdebsol])) AS impdebsol, " _
                + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[imphabdol]*[con_tc].[impven]),[con_diario].[imphabsol])) AS imphabsol, " _
                + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol],IIf([con_diario].[impdebsol]=0 Or [con_tc].[impven] Is Null,0,[con_diario].[impdebsol]/[con_tc].[impven]))) AS impdebdol, " _
                + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol],IIf([con_diario].[imphabsol]=0 Or [con_diario]![imphabsol] Is Null Or [con_tc].[impven] Is Null, 0, [con_diario].[imphabsol] / [con_tc].[impven]))) As imphabdol " _
                + vbCr + " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue "
            If opt_fecha(0).Value = True Then
                nSQL = nSQL + vbCr + " WHERE (con_diario.fchasi Is Null and con_diario.año = " & AnoTra & " ) Or (con_diario.fchasi < CDate('" & TxtFchIni.Valor & "')) "
            Else
                nSQL = nSQL + vbCr + " WHERE (con_diario.idmes < " & mMesIni & " AND con_diario.año = " & AnoTra & " )  or (con_diario.año < " & AnoTra & " ) "
            End If
            nSQL = nSQL + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.tipsal  " _
                + vbCr + " HAVING con_planctas.cuenta ='" & RstMay("cuenta") & "' "
                            
            RST_Busq RstSal, nSQL, xCon
            
            If RstSal.RecordCount <> 0 Then
                    If UCase(RstSal.Fields("tipsal") & "") = "D" Or NulosC(RstSal.Fields("tipsal")) = "" Then
                        xSaldo = (NulosN(RstSal(CAMPO_DEBE)) - NulosN(RstSal(CAMPO_HABER)))
                    Else
                        xSaldo = (NulosN(RstSal(CAMPO_HABER)) - NulosN(RstSal(CAMPO_DEBE)))
                    End If
                    Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xSaldo, FORMAT_MONTO)
            Else
                xSaldo = 0
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = "0.00"
            End If
                        
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, , True
            
            RstTmp2.Filter = adFilterNone
            RstTmp2.Filter = "idcuenta = '" & NulosC(RstMay("idcuenta")) & "'"
            Label3.Caption = "Procesando Cta Nº  :  " + NulosC(RstMay("cuenta"))
            DoEvents
            xFila = xFila + 1
            If RstTmp2.RecordCount <> 0 Then
            
                RstTmp2.MoveFirst
                If opt(0).Value = True Then
                    RstTmp2.Sort = "fchemi"
                ElseIf opt(1).Value = True Then
                    RstTmp2.Sort = "numdoc"
                Else
                    RstTmp2.Sort = "numreg"
                End If
                Do While Not RstTmp2.EOF
                    DoEvents
                    '--SI SE NTERRUMPE EL PROCESO => SALIR
                    If BAND_INTERRUMPIR = True Then GoTo SALIR
                    '-----------------------------------------------
                    Fg1.Rows = Fg1.Rows + 1
                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstMay("cuenta"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstTmp2("numreg"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstTmp2("nomlib"))
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstTmp2("tipdoc"))
                    
                    If IsDate(RstTmp2("fchemi")) = True Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(RstTmp2("fchemi"), FORMAT_DATE)
                    End If
                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(RstTmp2("numdoc"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(RstTmp2("impven"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(RstTmp2(CAMPO_DEBE)), FORMAT_MONTO)
                    Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosN(RstTmp2(CAMPO_HABER)), FORMAT_MONTO)
                    
                    If UCase(RstTmp2.Fields("tipsal") & "") = "D" Or NulosC(RstTmp2.Fields("tipsal")) = "" Then
                        xSaldo = xSaldo + (NulosN(RstTmp2(CAMPO_DEBE)) - NulosN(RstTmp2(CAMPO_HABER)))
                    Else
                        xSaldo = xSaldo + (NulosN(RstTmp2(CAMPO_HABER)) - NulosN(RstTmp2(CAMPO_DEBE)))
                    End If
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xSaldo, FORMAT_MONTO)
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosC(RstTmp2("nomclipro"))
                    
                    xTotal1 = xTotal1 + NulosN(RstTmp2(CAMPO_DEBE))
                    xTotal2 = xTotal2 + NulosN(RstTmp2(CAMPO_HABER))
                    RstTmp2.MoveNext
                    If RstTmp2.EOF = True Then
                        Fg1.Rows = Fg1.Rows + 1
                        Fg1.TextMatrix(Fg1.Rows - 1, 7) = "Total =>"
                        Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xTotal1, FORMAT_MONTO)
                        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(xTotal2, FORMAT_MONTO)
                        
                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, , True
                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, , True
                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, , True
                        
                        Fg1.Rows = Fg1.Rows + 1
                        Exit Do
                    End If
                    xFila = xFila + 1
                Loop
            End If
            
            xTotal1_1 = xTotal1_1 + xTotal1
            xTotal2_1 = xTotal2_1 + xTotal2

            RstMay.MoveNext
            
        Loop
        
    Else
        xSaldo = 0

        'hallamos el saldo anterior de la cuenta

        nSQL = "SELECT con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion ,con_planctas.tipsal, " _
            + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[impdebdol]*[con_tc].[impven]),[con_diario].[impdebsol])) AS impdebsol, " _
            + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[imphabdol]*[con_tc].[impven]),[con_diario].[imphabsol])) AS imphabsol, " _
            + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol],IIf([con_diario].[impdebsol]=0 Or [con_tc].[impven] Is Null,0,[con_diario].[impdebsol]/[con_tc].[impven]))) AS impdebdol, " _
            + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol], IIf([con_diario].[imphabsol] = 0 Or [con_diario].[imphabsol] Is Null Or [con_tc].[impven] Is Null, 0, [con_diario].[imphabsol] / [con_tc].[impven]))) As imphabdol " _
            + vbCr + " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue "

        If opt_fecha(0).Value = True Then
            nSQL = nSQL + vbCr + " WHERE (con_diario.fchasi Is Null and con_diario.año = " & AnoTra & " ) Or (con_diario.fchasi < CDate('" & TxtFchIni.Valor & "')) "
        Else
            nSQL = nSQL + vbCr + " WHERE (con_diario.idmes < " & mMesIni & " AND con_diario.año = " & AnoTra & " )  or (con_diario.año < " & AnoTra & " ) "
        End If
        
        nSQL = nSQL + nSQLCuenta _
            + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion,con_planctas.tipsal ; " _
            + vbCr + " "
            
        RST_Busq RstSal, nSQL, xCon
        
        If RstSal.EOF = False Or RstSal.BOF = False Or RstSal.RecordCount <> 0 Then
            Fg1.Rows = Fg1.Rows + 1
            RstSal.MoveFirst
        End If
        
        Do While Not RstSal.EOF
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 5, "Cta Nº  :  " + NulosC(RstSal("cuenta")) & "   - " + NulosC(RstSal.Fields("descripcion")), flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 1, , True
            
            If UCase(RstSal.Fields("tipsal") & "") = "D" Or RstSal.Fields("tipsal") = "" Then
                xSaldo = (NulosN(RstSal(CAMPO_DEBE)) - NulosN(RstSal(CAMPO_HABER)))
            Else
                xSaldo = (NulosN(RstSal(CAMPO_HABER)) - NulosN(RstSal(CAMPO_DEBE)))
            End If
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xSaldo, FORMAT_MONTO)
            Fg1.Rows = Fg1.Rows + 1
            RstSal.MoveNext
        Loop
        Set RstSal = Nothing
        '----------------------------
    End If

    If xTotal1_1 <> 0 Or xTotal2_1 <> 0 Then
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = "Total Gral =>"
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xTotal1_1, FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(xTotal2_1, FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, , True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, , True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, , True
    End If

SALIR:
    Set RstMay = Nothing:     Set RstDet = Nothing:     Set RstSal = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    MuestraMayor = True
    Exit Function
error:
    Set RstMay = Nothing:     Set RstDet = Nothing:     Set RstSal = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    SHOW_ERROR Me.Name, "MuestraMayor"
End Function




Sub ExportarExcelDetalle()
    Dim A&, B&, xFilas&
    
    On Error GoTo error

    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.Workbooks.Add  'Trim(App.Path) + "\RegCompras.xls"
    
    objExcel.WindowState = 1
    
    objExcel.Visible = True
    With objExcel.ActiveSheet
        
        .Cells(1, 2) = NomEmp
        .Cells(1, 11) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        
        '**********************************
        Dim nPeriodo As String
        Dim nTitulo1 As String
        If opt_fecha(0).Value = True Then
            If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
                nPeriodo = " Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
            Else
                nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
            End If
        Else
            If mMesIni = mMesFin Then
                nPeriodo = "Periodo: " + lbl_periodo(0).Caption
            Else
                nPeriodo = "Periodo: De " + lbl_periodo(0).Caption & " A " & lbl_periodo(1).Caption
            End If
            
        End If

        nTitulo1 = "(Expresado en " & LblMoneda.Caption & ")"
        .Cells(4, 2) = "Libro Mayor"
        .Cells(5, 2) = nPeriodo
        .Cells(6, 2) = nTitulo1
        '**********************************
        '--ancho de columna
        For B = 1 To Fg1.Cols - 1
            .Columns(B + 1).ColumnWidth = Fg1.ColWidth(B) / 100
        Next B
        '--encabezado
        xFilas = 8
        For B = 1 To Fg1.Cols - 1
            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(0, B)
        Next B
        
        .Range("B6:K7").Font.Bold = True
    
        xFilas = xFilas + 1
        For A = 1 To Fg1.Rows - 1
            DoEvents
            For B = 1 To Fg1.Cols - 1
                
                If B <= 7 Then
                    If B = 1 Then
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                        If InStr(Fg1.TextMatrix(A, B), "Cta Nº  :") <> 0 Then
                            .Cells(xFilas, 2) = "'" + Fg1.TextMatrix(A, B)
                            GoTo SIG_FIL
                        End If
                    Else
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                    End If
                    
                    
                Else
                    If IsNumeric(Fg1.TextMatrix(A, B)) = True Then
                        .Cells(xFilas, B + 1) = NulosN(Fg1.TextMatrix(A, B))
                    Else
                        .Cells(xFilas, B + 1) = NulosC(Fg1.TextMatrix(A, B))
                    End If
                End If

            Next B
SIG_FIL:
            xFilas = xFilas + 1
        Next A
    End With
    
    MsgBox "El proceso de exportación terminó con éxito", vbInformation, xTitulo
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "ExportarExcelDetalle", , IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
    
End Sub


Private Sub pConfigurarGrilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
                                   
    Dim k&, j&
''    With Fg1
''        '-----
''        If F_CONSERVAR_FORMATO = True Then LimpiarGrid Fg1, , 1
''        .FrozenCols = 0
''        Fg1.Cols = 12
''
''        .ColWidth(0) = 200
''        '--DATOS DE FILA
''        .TextMatrix(0, 1) = "Nº.Cuenta":  .ColWidth(1) = 1000:       .ColAlignment(1) = flexAlignLeftBottom:     .FixedAlignment(1) = flexAlignCenterTop
''        .TextMatrix(0, 2) = "Num.Reg.":     .ColWidth(2) = 850:    .ColAlignment(2) = flexAlignLeftCenter:     .FixedAlignment(2) = flexAlignCenterTop
''        .TextMatrix(0, 3) = "Libro":        .ColWidth(3) = 1500:    .ColAlignment(3) = flexAlignLeftBottom:     .FixedAlignment(3) = flexAlignCenterTop
''        .TextMatrix(0, 4) = "T.D.":         .ColWidth(4) = 450:     .ColAlignment(4) = flexAlignLeftCenter:     .FixedAlignment(4) = flexAlignCenterTop
''        .TextMatrix(0, 5) = "Fch. Doc":     .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignCenterBottom:   .FixedAlignment(5) = flexAlignCenterTop
''        .TextMatrix(0, 6) = "Nº.Documento": .ColWidth(6) = 1500:    .ColAlignment(6) = flexAlignLeftBottom:     .FixedAlignment(6) = flexAlignCenterTop
''        .TextMatrix(0, 7) = "T.C.":         .ColWidth(7) = 800:     .ColAlignment(7) = flexAlignRightBottom:   .FixedAlignment(7) = flexAlignCenterTop
''        .TextMatrix(0, 8) = "Debe":         .ColWidth(8) = 1300:    .ColAlignment(8) = flexAlignRightBottom:    .FixedAlignment(8) = flexAlignCenterTop
''        .TextMatrix(0, 9) = "Haber":        .ColWidth(9) = 1300:    .ColAlignment(9) = flexAlignRightBottom:    .FixedAlignment(9) = flexAlignCenterTop
''        .TextMatrix(0, 10) = "Saldo":       .ColWidth(10) = 1300:   .ColAlignment(10) = flexAlignRightBottom:   .FixedAlignment(10) = flexAlignCenterTop
''        .TextMatrix(0, 11) = "Cliente / Proveedor":       .ColWidth(11) = 3000:   .ColAlignment(11) = flexAlignLeftTop:  .FixedAlignment(11) = flexAlignCenterTop
''    End With
    
    With Fg3
        '-----
        If F_CONSERVAR_FORMATO = True Then LimpiarGrid Fg3, , 2
        
        .Cols = 11
        .FixedRows = 2
        .FrozenCols = 2
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        UNIR_CELDAS Fg3, 0, 1, 0, 2, "Datos de la Cuenta", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 3, 0, 4, "Saldos Iniciales", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 5, 0, 6, "Movimiento del Periodo", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 7, 0, 8, "Sumas del Mayor", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 9, 0, 10, "Saldos Finales", flexAlignCenterCenter
        
'        '--DATOS DE FILA
        .TextMatrix(1, 1) = "Nº. Cuenta":       .ColWidth(1) = 1100:       .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(1, 2) = "Descripción":      .ColWidth(2) = 3000:       .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(1, 3) = "Debe":       .ColWidth(3) = 1300:       .ColAlignment(3) = flexAlignRightCenter
        .TextMatrix(1, 4) = "Haber":      .ColWidth(4) = 1300:       .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(1, 5) = "Debe":       .ColWidth(5) = 1320:       .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(1, 6) = "Haber":      .ColWidth(6) = 1320:       .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(1, 7) = "Debe":       .ColWidth(7) = 1320:       .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(1, 8) = "Haber":      .ColWidth(8) = 1320:       .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(1, 9) = "Debe":       .ColWidth(9) = 1200:       .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(1, 10) = "Haber":     .ColWidth(10) = 1200:      .ColAlignment(10) = flexAlignRightCenter
        
        '--AGREGANDO LAS FECHAS EN LA CABECERA
        If opt_fecha(0).Value = True Then
            If IsDate(TxtFchIni.Valor) = True Then UNIR_CELDAS Fg3, 0, 3, 0, 4, "Saldos Iniciales" + vbCr + " Al " + CStr(CDate(TxtFchIni.Valor) - 1), flexAlignCenterCenter
            If IsDate(TxtFchFin.Valor) = True Then UNIR_CELDAS Fg3, 0, 9, 0, 10, "Saldos Finales" + vbCr + " Al " + CStr(CDate(TxtFchFin.Valor) - 1), flexAlignCenterCenter
        Else
            If IsDate(TxtFchIni.Valor) = True Then UNIR_CELDAS Fg3, 0, 3, 0, 4, "Saldos Iniciales" + vbCr + " A " + lbl_periodo(0).Caption, flexAlignCenterCenter
            If IsDate(TxtFchFin.Valor) = True Then UNIR_CELDAS Fg3, 0, 9, 0, 10, "Saldos Finales" + vbCr + " A " + lbl_periodo(1).Caption, flexAlignCenterCenter
        End If
    End With
    
    DoEvents
End Sub


Private Function fValidarConsulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Function
    End If
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Seleccione la Moneda", vbExclamation, xTitulo
        Exit Function
    End If
    
    If opt_fecha(0).Value = True Then '--por fecha
        If NulosC(TxtFchIni.Valor) = "" Then
            MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
        
        If NulosC(TxtFchFin.Valor) = "" Then
            MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchFin.SetFocus
            Exit Function
        End If
        
        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
        
        If (Year(TxtFchIni.Valor) <> Year(TxtFchFin.Valor)) Then
            MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        ElseIf Year(TxtFchIni.Valor) <> CStr(AnoTra) Then
            MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
    Else '--por periodo
        If mMesIni > mMesFin Then
            MsgBox "El periodo de inicio debe ser inferior o igual al periodo final", vbExclamation, xTitulo
            cmd_periodo(0).SetFocus
            Exit Function
        End If
    End If
    If chk.Value = 0 Then
        If Fg2.Rows = 1 Then
            MsgBox "No ha especificado una cuenta contable a mayorizar" + vbCr + "Si desea ver todas las cuentas, Active la opción: Procesar Todas las Cuentas...", vbExclamation, xTitulo
            CmdAdd.SetFocus
            Exit Function
        End If
        If Fg2.TextMatrix(Fg2.Rows - 1, 1) = "" Then
            MsgBox "Seleccione la Cuenta Contable", vbExclamation, xTitulo
            Fg2.Row = Fg2.Rows - 1
            Fg2.Col = 1
            Fg2.SetFocus
            Exit Function
        End If
    End If
    fValidarConsulta = True
End Function



Private Sub ExportarExcelResumen()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim nPeriodo As String
    Dim nTitulo1 As String
    If opt_fecha(0).Value = True Then  '--por fecha
        If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
            nPeriodo = " Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
        Else
            nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
        End If
    Else '--por periodo
        If mMesIni = mMesFin Then
            nPeriodo = "Periodo:  " + lbl_periodo(0).Caption
        Else
            nPeriodo = "Periodo:  De " + lbl_periodo(0).Caption & " A " & lbl_periodo(1).Caption
        End If
        
    End If
    nTitulo1 = "(Expresado en " & LblMoneda.Caption & ")"
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg3, "RESUMEN DEL MAYOR", nPeriodo, nTitulo1, "Resumen del Mayor"
    
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub


Private Sub BuscarVSFlexGrid()
    On Error GoTo error
    
    If Me.TabOne1.CurrTab <> 0 Then Exit Sub
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim xCampos(3, 3) As String
    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
    xCampos(0, 0) = "Num.Reg.":     xCampos(0, 1) = "2":    xCampos(0, 2) = "C":    xCampos(0, 3) = "-1"
    xCampos(1, 0) = "Libro":        xCampos(1, 1) = "3":    xCampos(1, 2) = "C":    xCampos(1, 3) = "0"
    xCampos(2, 0) = "Fch. Doc   ":  xCampos(2, 1) = "4":    xCampos(2, 2) = "F":    xCampos(2, 3) = "0"
    xCampos(3, 0) = "Nº Documento": xCampos(3, 1) = "5":    xCampos(3, 2) = "C":    xCampos(3, 3) = "0"
    
    X_EXPORT.VSFlexGrid_Buscar Me.hWnd, Fg1, xCampos(), Fg1.Row
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "BuscarVSFlexGrid"
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then pBuscarAsiento
    If Button.Index = 5 Then pExportar
    If Button.Index = 6 Then pImprimir
    If Button.Index = 7 Then Configurar
    If Button.Index = 9 Then
        Unload Me
    End If
End Sub

Sub Configurar()
    Dim xform As New SGI2_funciones.Varias
    If xform.CambioOpcionLiro(5, xCon, 1) = True Then
        SetearCuadricula Fg1, 5, xCon, 1, 0, False
        pConsultar
    End If
    Set xform = Nothing
End Sub

Private Sub opt_fecha_Click(Index As Integer)
    If Index = 0 Then '--por fecha
        TxtFchFin.Visible = True
        TxtFchIni.Visible = True
        lbl(0).Caption = "Del"
        lbl(1).Caption = "Al"
        
        Ocultar cmd_periodo, False
        Ocultar lbl_periodo, False
    Else '--por periodo
        TxtFchFin.Visible = False
        TxtFchIni.Visible = False
        lbl(0).Caption = "De"
        lbl(1).Caption = "A"
        cmd_periodo(0).Top = 240
        lbl_periodo(0).Top = 210
        cmd_periodo(1).Top = 540
        lbl_periodo(1).Top = 510
        Ocultar cmd_periodo, True
        Ocultar lbl_periodo, True
    End If
End Sub

Private Sub cmd_periodo_Click(Index As Integer)
    If Index = 0 Then
        mMesIni = SeleccionaMes(xCon)
        lbl_periodo(0).Caption = Busca_Codigo(mMesIni, "id", "descripcion", "con_meses", "N", xCon)
    Else
        mMesFin = SeleccionaMes(xCon)
        lbl_periodo(1).Caption = Busca_Codigo(mMesFin, "id", "descripcion", "con_meses", "N", xCon)
    End If
End Sub

Private Sub TxtFchFin_Validate(Cancel As Boolean)
    If IsDate(TxtFchFin.Valor) = True Then
        CmdAdd.SetFocus
    End If
End Sub

Private Sub TxtIdMon_Change()
    If Trim(TxtIdMon.Text) = "" Then TxtIdMon.Text = ""
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
        If NulosC(TxtIdMon.Text) = "" Then
            TxtIdMon.Text = ""
        End If
    End If
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

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error

    If Col = 1 Then
        Dim Rst As New ADODB.Recordset
        Dim xRs As New ADODB.Recordset
        Dim nSQL As String
        Dim nSQLLike As String
        Dim nSQLIdCta As String
          
        Dim xCampos(2, 4) As String
        
        xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "cuenta":             xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":        xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
        
        '--
        nSQLIdCta = GRID_GENERAR_SQL_ID(Fg2, 3, " WHERE con_planctas.id", " NOT IN ", True, 1, Fg2.TextMatrix(Row, 3))
        
        If NulosC(Fg2.TextMatrix(Fg2.Row, 1)) <> "" Then
            nSQLLike = " and con_planctas.cuenta like '" + Trim(Fg2.TextMatrix(Fg2.Row, 1)) + "%' "
        End If
        If nSQLIdCta = "" Then nSQLLike = Replace(nSQLLike, " and ", " WHERE ")
        
        
        nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id " _
            + vbCr + " From con_planctas " + nSQLIdCta + nSQLLike + vbCr + "  ORDER BY con_planctas.cuenta"
                
        CARGAR_DLL_EPSBUSCAR xCon, Rst, nSQL, xCampos(), "Buscando Cuentas Contables", "cuenta", "cuenta", Principio
        
        If Rst.State = 0 Then GoTo SALIR
        If Rst.RecordCount = 0 Then GoTo SALIR
        
        If fValidarSeleccionCta(NulosC(Rst("cuenta"))) = False Then GoTo SALIR

        Agregando = True
    
        Fg2.TextMatrix(Fg2.Row, 1) = NulosC(Rst("cuenta"))
        Fg2.TextMatrix(Fg2.Row, 2) = NulosC(Rst("descripcion"))
        Fg2.TextMatrix(Fg2.Row, 3) = NulosN(Rst("id"))
        
        Set Rst = Nothing
        Set xRs = Nothing
    End If
    
SALIR:
    
    Agregando = False
    Exit Sub
error:
    'Resume
    Set Rst = Nothing
    Set xRs = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "fg2_CellButtonClick"
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    
    If Fg2.TextMatrix(Row, Col) = "" Then
        Fg2.TextMatrix(Row, 1) = ""
        Fg2.TextMatrix(Row, 2) = ""
        Fg2.TextMatrix(Row, 3) = ""
        Exit Sub
    End If
    
    If Col = 1 Then
        If fValidarSeleccionCta(NulosN(Fg2.TextMatrix(Row, Col))) = False Then
            Fg2.TextMatrix(Row, 1) = ""
            Fg2.TextMatrix(Row, 2) = ""
            Fg2.TextMatrix(Row, 3) = ""
            Exit Sub
        End If
        
        Dim Rst As New ADODB.Recordset
        RST_Busq Rst, "SELECT * FROM con_planctas WHERE cuenta = '" & NulosC(Fg2.TextMatrix(Row, 1)) & "'", xCon
        If Rst.RecordCount = 1 Then
            Fg2.TextMatrix(Row, 2) = NulosC(Rst("descripcion"))
            Fg2.TextMatrix(Row, 3) = NulosN(Rst("id"))
        Else
            Fg2.TextMatrix(Row, 1) = ""
            Fg2.TextMatrix(Row, 2) = ""
            Fg2.TextMatrix(Row, 3) = ""
        End If
        Set Rst = Nothing
    End If
   
End Sub

Private Sub Fg2_EnterCell()
    If Fg2.Col = 1 Then
        Fg2.Editable = flexEDKbdMouse
    Else
        Fg2.Editable = flexEDNone
    End If
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 45 Then
        CmdAdd_Click
    End If
    
    If KeyCode = 46 Then
        CmdDel_Click
    End If
End Sub

Private Sub fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    
    Select Case Col
        Case 1
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Function MuestraMayor1() As Boolean
    Dim RstMay As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstSal As New ADODB.Recordset
    Dim nSQLCampos As String
    Dim mCol As Long
    Dim mColCampo As Integer
    Dim nSQLAjuste As String
        
    Dim A&, B&, C&
    Dim nSQL As String
    On Error GoTo error

    Frame5.Left = 3413
    Frame5.Top = 3780
    Frame5.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0
    DoEvents
    
    Dim nSQLCuenta As String
    nSQLCuenta = ""
    '--SI AGREGA CUENTAS AS GRID, GENERAR EL FILTRO A CONCATENAR A LA CONSULTA
    For A = 1 To Fg2.Rows - 1
        If Trim(Fg2.TextMatrix(A, 1)) <> "" Then
            nSQLCuenta = nSQLCuenta + " con_planctas.cuenta Like '" & Trim(Fg2.TextMatrix(A, 1)) & "%' OR "
        End If
    Next A
    If nSQLCuenta <> "" Then nSQLCuenta = " AND (" + Left(nSQLCuenta, Len(nSQLCuenta) - 3) + ") "
    '---------
    '--ESTABLECER EL CAMPO A TOTALIZAR EN FUNCION DEL RECORDSET TMP (RstTmp2) , TANTO A SOLES Y DOLARES
    Dim CAMPO_DEBE, CAMPO_HABER, CAMPO_SALDO As String
    If NulosN(TxtIdMon.Text) = 1 Then
        CAMPO_DEBE = "impdebsol":  CAMPO_HABER = "imphabsol": CAMPO_SALDO = "impsalsol"
    Else
        CAMPO_DEBE = "impdebdol":  CAMPO_HABER = "imphabdol": CAMPO_SALDO = "impsaldol"
    End If
    '-----------------------------------------------------------------------
    '--SI SE NTERRUMPE EL PROCESO => SALIR
    If BAND_INTERRUMPIR = True Then GoTo SALIR:
    '-----------------------------------------------
     
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) AND "
    '-----------------------------------------------
     
     
    Set RstTmp2 = Nothing
    '**********************************************************************************************
    nSQLCampos = fSetearCuadriculaColumna(xCon, 5)
    If nSQLCampos = "" Then Exit Function
    nSQLCampos = "idcuenta,tipsal," & nSQLCampos
     '**********************************************************************************************
'    'antes de 07/02/09
'    nSQL = "SELECT con_diario.idcue AS idcuenta,con_planctas.tipsal,Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, Format([con_diario].[idmes],'00') AS mes, mae_libros.codsun AS libsun, CDbl(con_diario.numasi) AS corr, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, con_diario.rglosa AS glosa, con_diario.rregistro AS registroref, iif(con_diario.ridtipper=0,tes_documentos.abrev,mae_documento.abrev) AS tdocdesc, con_diario.rfchope AS fchdoc, con_diario.rnumerodoc AS numdoc, " _
'            + vbCr + " IIf([con_diario].[ridtipper]=1,[mae_prov].[numruc],IIf([con_diario].[ridtipper]=2,[mae_cliente].[numruc],IIf([con_diario].[ridtipper]=3,[pla_empleados].[numdoc],''))) AS numruc, " _
'            + vbCr + " IIf([con_diario].[ridtipper]=1,[mae_prov].[nombre],IIf([con_diario].[ridtipper]=2,[mae_cliente].[nombre],IIf([con_diario].[ridtipper]=3,[pla_empleados].[nombre],''))) AS apenom, mae_documento.codsun AS tdocsun, con_tc.impven AS tc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo as moneda, " _
'            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
'            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
'            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
'            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
'            + vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',impdebsol-imphabsol,imphabsol-impdebsol) as impsalsol, " _
'            + vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',impdebdol-imphabdol,imphabdol-impdebdol) as impsaldol, " _
'            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc " _
'            + vbCr + " FROM (pla_empleados RIGHT JOIN (mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper) ON pla_empleados.id = con_diario.ridper) LEFT JOIN tes_documentos ON con_diario.rtipdoc = tes_documentos.id "
    
    '--obs: esta consulta es iddentica a la consulta del diario excepto
    '-- agregar con_diario.idcue AS idcuenta,con_planctas.tipsal, y
    '--+ vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',impdebsol-imphabsol,imphabsol-impdebsol) as impsalsol, " _
    '--+ vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',impdebdol-imphabdol,imphabdol-impdebdol) as impsaldol, " _

    '--con tipo de cambio todo con con_tc
'   nSQL = "SELECT con_diario.idcue AS idcuenta,con_planctas.tipsal,Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, Format([con_diario].[idmes],'00') AS mes, mae_libros.codsun AS libsun, CDbl(con_diario.numasi) AS corr, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, con_diario.rglosaope as glosaope, con_diario.rglosa AS glosaref, con_diario.rregistro AS registroref, iif(con_diario.ridtipper=0,tes_documentos.abrev,mae_documento.abrev) AS tdocdesc, con_diario.rfchope AS fchdoc, con_diario.rnumerodoc AS numdoc, " _
'            + vbCr + " IIf([con_diario].[ridtipper]=1,[mae_prov].[numruc],IIf([con_diario].[ridtipper]=2,[mae_cliente].[numruc],IIf([con_diario].[ridtipper]=3,[pla_empleados].[numdoc],IIf([con_diario].[ridtipper]=5,[mae_bancos].[numruc],'')))) AS numruc, " _
'            + vbCr + "  IIf([con_diario].[ridtipper]=1,[mae_prov].[nombre],IIf([con_diario].[ridtipper]=2,[mae_cliente].[nombre],IIf([con_diario].[ridtipper]=3,[pla_empleados].[apepat]&' '&[pla_empleados].[apemat]&', '&[pla_empleados].[nom],IIf([con_diario].[ridtipper]=5,[mae_bancos].[descripcion],'')))) AS apenom , mae_documento.codsun AS tdocsun, con_tc.impven AS tc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo AS monope, mae_moneda_1.simbolo AS monref, " _
'            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
'            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
'            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
'            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
'            + vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',impdebsol-imphabsol,imphabsol-impdebsol) as impsalsol, " _
'            + vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',impdebdol-imphabdol,imphabdol-impdebdol) as impsaldol, " _
'            + vbCr + " iif(con_diario.rnumerodoc1 is null,'',mae_documento_1.abrev) AS tdocdesc1, con_diario.rnumerodoc1 AS numdoc1, " _
'            + vbCr + " tes_documentos_1.abrev AS tdocdesc2, con_diario.rfchope2 AS fchdoc2, con_diario.rnumerodoc2 AS numdoc2,con_diario.ridtipper2, iif(con_diario.ridtipper2<>5,'', mae_bancos_1.numruc ) AS numruc2,iif(con_diario.ridtipper2<>5,'',mae_bancos_1.descripcion ) AS apenom2 " _
'            + vbCr + " FROM ((((((pla_empleados RIGHT JOIN (mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper) ON pla_empleados.id = con_diario.ridper) LEFT JOIN tes_documentos ON con_diario.rtipdoc = tes_documentos.id) LEFT JOIN mae_bancos ON con_diario.ridper = mae_bancos.id) LEFT JOIN mae_bancos AS mae_bancos_1 ON con_diario.ridper2 = mae_bancos_1.id) LEFT JOIN mae_documento AS mae_documento_1 ON con_diario.rtipdoc1 = mae_documento_1.id) LEFT JOIN tes_documentos AS tes_documentos_1 ON con_diario.rtipdoc2 = tes_documentos_1.id) LEFT JOIN mae_moneda AS mae_moneda_1 ON con_diario.ridmon = mae_moneda_1.id "



   '--tomar tipo de cambio del diario cuando idlib = bancos y diversos
   nSQL = "SELECT con_diario.idcue AS idcuenta,con_planctas.tipsal,Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, Format([con_diario].[idmes],'00') AS mes, mae_libros.codsun AS libsun, CDbl(con_diario.numasi) AS corr, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, con_diario.rglosaope as glosaope, con_diario.rglosa AS glosaref, con_diario.rregistro AS registroref, iif(con_diario.ridtipper=0,tes_documentos.abrev,mae_documento.abrev) AS tdocdesc, con_diario.rfchope AS fchdoc, con_diario.rnumerodoc AS numdoc, " _
            + vbCr + " IIf([con_diario].[ridtipper]=1,[mae_prov].[numruc],IIf([con_diario].[ridtipper]=2,[mae_cliente].[numruc],IIf([con_diario].[ridtipper]=3,[pla_empleados].[numdoc],IIf([con_diario].[ridtipper]=5,[mae_bancos].[numruc],'')))) AS numruc, " _
            + vbCr + " IIf([con_diario].[ridtipper]=1,[mae_prov].[nombre],IIf([con_diario].[ridtipper]=2,[mae_cliente].[nombre],IIf([con_diario].[ridtipper]=3,[pla_empleados].[apepat]&' '&[pla_empleados].[apemat] & ', ' & [pla_empleados].[nom],IIf([con_diario].[ridtipper]=5,[mae_bancos].[descripcion],'')))) AS apenom , mae_documento.codsun AS tdocsun, iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo AS monope, mae_moneda_1.simbolo AS monref, "
    
    If NulosN(TxtIdMon.Text) = 1 Then
        nSQL = nSQL _
            + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',impdebesol-imphabersol,imphabersol-impdebesol) as impsalsol, "
    Else
        nSQL = nSQL _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) AS imphaberdol, " _
            + vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',impdebedol-imphaberdol,imphaberdol-impdebedol) as impsaldol, "
    End If
    
    nSQL = nSQL _
        + vbCr + " iif(con_diario.rnumerodoc1 is null,'',mae_documento_1.abrev) AS tdocdesc1, con_diario.rnumerodoc1 AS numdoc1, " _
        + vbCr + " tes_documentos_1.abrev AS tdocdesc2, con_diario.rfchope2 AS fchdoc2, con_diario.rnumerodoc2 AS numdoc2,con_diario.ridtipper2, iif(con_diario.ridtipper2<>5,'', mae_bancos_1.numruc ) AS numruc2,iif(con_diario.ridtipper2<>5,'',mae_bancos_1.descripcion ) AS apenom2 " _
        + vbCr + " FROM ((((((pla_empleados RIGHT JOIN (mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper) ON pla_empleados.id = con_diario.ridper) LEFT JOIN tes_documentos ON con_diario.rtipdoc = tes_documentos.id) LEFT JOIN mae_bancos ON con_diario.ridper = mae_bancos.id) LEFT JOIN mae_bancos AS mae_bancos_1 ON con_diario.ridper2 = mae_bancos_1.id) LEFT JOIN mae_documento AS mae_documento_1 ON con_diario.rtipdoc1 = mae_documento_1.id) LEFT JOIN tes_documentos AS tes_documentos_1 ON con_diario.rtipdoc2 = tes_documentos_1.id) LEFT JOIN mae_moneda AS mae_moneda_1 ON con_diario.ridmon = mae_moneda_1.id "
        
    If opt_fecha(0).Value = True Then
        nSQL = nSQL + vbCr + " WHERE " & nSQLAjuste & " ( con_diario.fchasi >=CDate('" + TxtFchIni.Valor + "') And con_diario.fchasi<=CDate('" + TxtFchFin.Valor + "') ) " _
            + vbCr + " AND ( con_diario.fchasi >=CDate('01/01/" + AnoTra + "') And con_diario.fchasi <= CDate('31/12/" + AnoTra + "') ) "
    Else
        nSQL = nSQL + vbCr + " WHERE " & nSQLAjuste & " ( con_diario.idmes >= " & mMesIni & " and con_diario.idmes <= " & mMesFin & " ) and con_diario.año = " & AnoTra & " "
    End If
        
    nSQL = nSQL + nSQLCuenta + vbCr + " ORDER BY con_planctas.cuenta ASC "

     '**********************************************************************************************
    '--remplazando segun la moneda seleccionada
    
    If NulosN(TxtIdMon.Text) = 1 Then
        nSQL = Replace(nSQL, "impdebesol", "debe")
        nSQL = Replace(nSQL, "imphabersol", "haber")
        nSQL = Replace(nSQL, "impsalsol", "saldo")
    Else
        nSQL = Replace(nSQL, "impdebedol", "debe")
        nSQL = Replace(nSQL, "imphaberdol", "haber")
        nSQL = Replace(nSQL, "impsaldol", "saldo")
    End If
    
    '--manipulando el encabezado
'''    If NulosN(TxtIdMon.Text) = 1 Then
'''        nSQLCampos = Replace(nSQLCampos, "debe", "impdebesol")
'''        nSQLCampos = Replace(nSQLCampos, "haber", "imphabersol")
'''        nSQLCampos = Replace(nSQLCampos, "saldo", "impsalsol")
'''    Else
'''        nSQLCampos = Replace(nSQLCampos, "debe", "impdebedol")
'''        nSQLCampos = Replace(nSQLCampos, "haber", "imphaberdol")
'''        nSQLCampos = Replace(nSQLCampos, "saldo", "impsaldol")
'''    End If
    
    
    nSQL = "Select " & nSQLCampos & _
            vbCr + " from ( " _
            + vbCr + nSQL _
            + vbCr + ") as diario ORDER BY registro, ctanum ,numdoc"
     '**********************************************************************************************
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    
    RST_Busq RstTmp2, nSQL, xCon
    
    '--obtener la posicione de las columnas debe,haber,saldo
    mCol = 0
    For mColCampo = 2 To RstTmp2.Fields.Count - 1
        mCol = mCol + 1
        Select Case LCase(RstTmp2.Fields(mColCampo).Name)
            Case "debe", "impdebesol", "impdebedol": mColDebe = mCol
            Case "haber", "imphabersol", "imphaberdol": mColHaber = mCol
            Case "saldo", "impsalsol", "impsaldol": mColSaldo = mCol
            Case "registro": mPosRegistro = mCol
        End Select
    Next mColCampo


    'HACEMOS UNA CONSULTA DE LOS REGISTROS UNICOS DE LA CONSULTA ANTERIOR, PARA PODER TOTALIZARLA

    Dim xFila&
    Dim xSaldo As Double
    Dim xTotal1, xTotal2 As Double
    Dim xTotal1_1, xTotal2_1 As Double
    
    xFila = 1

    DoEvents
    '--SI SE NTERRUMPE EL PROCESO => SALIR
     If BAND_INTERRUMPIR = True Then GoTo SALIR:
    '-----------------------------------------------
    nSQL = "SELECT con_diario.idcue as idcuenta, con_planctas.cuenta, con_planctas.descripcion, Sum(con_diario.impdebsol) AS SumaDeimpdebsol, " _
         + vbCr + " Sum(con_diario.imphabsol) AS SumaDeimphabsol, Sum(con_diario.impdebdol) AS SumaDeimpdebdol, Sum(con_diario.imphabdol) AS SumaDeimphabdol " _
         + vbCr + " FROM con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue "
                  
    If opt_fecha(0).Value = True Then
        nSQL = nSQL + vbCr + " WHERE " & nSQLAjuste & " ( ( con_diario.fchasi >=CDate('" & TxtFchIni.Valor & "') And con_diario.fchasi <=CDate('" & TxtFchFin.Valor & "') ) and ( year(con_diario.fchasi)= " & AnoTra & " ) OR  con_diario.fchasi IS NULL ) "
    Else
        nSQL = nSQL + vbCr + " WHERE " & nSQLAjuste & " ( ( con_diario.idmes >= " & mMesIni & " and con_diario.idmes <= " & mMesFin & " ) and con_diario.año = " & AnoTra & " ) OR  con_diario.fchasi IS NULL ) "
    End If
         
    nSQL = nSQL + vbCr + nSQLCuenta _
         + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion ORDER BY con_planctas.cuenta ASC "
    
    RST_Busq RstMay, nSQL, xCon

    xSaldo = 0
    '---
    If RstMay.RecordCount <> 0 Then
    
        fra_msg.Visible = True
        
        RstMay.MoveFirst
        
        If RstMay.RecordCount > 1 Then ProgressBar1.Max = RstMay.RecordCount
        
        DoEvents
        Do While Not RstMay.EOF
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo SALIR:
            '-----------------------------------------------
            ProgressBar1.Value = RstMay.Bookmark
    
            xSaldo = 0
            Fg1.Rows = Fg1.Rows + 1
            
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 7, "Cta Nº  :  " + NulosC(RstMay("cuenta")) & "   - " + NulosC(RstMay.Fields("descripcion")), flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 1, , True
                        
            xTotal1 = 0
            xTotal2 = 0
            
           'hallamos el saldo anterior de la cuenta
            Set RstSal = Nothing
                                                        
'''                + vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',[impdebsol]-[imphabsol],[imphabsol]-[impdebsol]) as impsalsol, " _
'''                + vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',[impdebdol]-[imphabdol],[imphabdol]-[impdebdol]) as impsaldol " _

            nSQL = "SELECT con_diario.idcue as idcuenta, con_planctas.cuenta,con_planctas.tipsal, " _
                + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS impdebsol, " _
                + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS imphabsol, " _
                + vbCr + " Sum(IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven)))))) AS impdebdol, " _
                + vbCr + " Sum(IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven)))))) As imphabdol " _
                + vbCr + " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue "
            If opt_fecha(0).Value = True Then
                nSQL = nSQL + vbCr + " WHERE ( " & nSQLAjuste & " con_diario.fchasi Is Null and con_diario.año = " & AnoTra & " ) Or ( " & nSQLAjuste & " con_diario.fchasi < CDate('" & TxtFchIni.Valor & "')) "
            Else
                nSQL = nSQL + vbCr + " WHERE ( " & nSQLAjuste & " con_diario.idmes < " & mMesIni & " AND con_diario.año = " & AnoTra & " )  or ( " & nSQLAjuste & " con_diario.año < " & AnoTra & " ) "
            End If
            nSQL = nSQL + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.tipsal  " _
                + vbCr + " HAVING con_planctas.cuenta ='" & RstMay("cuenta") & "' "
                            
            RST_Busq RstSal, nSQL, xCon
            
            If RstSal.RecordCount <> 0 Then
                    If UCase(RstSal.Fields("tipsal") & "") = "D" Or NulosC(RstSal.Fields("tipsal")) = "" Then
                        xSaldo = (NulosN(RstSal(CAMPO_DEBE)) - NulosN(RstSal(CAMPO_HABER)))
'                        Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(Abs(xSaldo), FORMAT_MONTO)
'                        xTotal1 = Abs(xSaldo)
                    Else
                        xSaldo = (NulosN(RstSal(CAMPO_HABER)) - NulosN(RstSal(CAMPO_DEBE)))
'                        Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(Abs(xSaldo), FORMAT_MONTO)
'                        xTotal2 = Abs(xSaldo)
                    End If
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(NulosN(RstSal(CAMPO_DEBE)), FORMAT_MONTO)
                    Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(NulosN(RstSal(CAMPO_HABER)), FORMAT_MONTO)
                    xTotal1 = NulosN(RstSal(CAMPO_DEBE))
                    xTotal2 = NulosN(RstSal(CAMPO_HABER))
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, mColSaldo) = Format(xSaldo, FORMAT_MONTO)
                    
                    
                    
            Else
                xSaldo = 0
                Fg1.TextMatrix(Fg1.Rows - 1, mColSaldo) = "0.00"
            End If
                        
            FORMATO_CELDA Fg1, Fg1.Rows - 1, mColSaldo, , True
            
            RstTmp2.Filter = adFilterNone
            RstTmp2.Filter = "idcuenta = '" & NulosC(RstMay("idcuenta")) & "'"
            Label3.Caption = "Procesando Cta Nº  :  " + NulosC(RstMay("cuenta"))
            DoEvents
            xFila = xFila + 1
            If RstTmp2.RecordCount <> 0 Then
            
                RstTmp2.MoveFirst
                If opt(0).Value = True Then
                    RstTmp2.Sort = "fchdoc"
                ElseIf opt(1).Value = True Then
                    RstTmp2.Sort = "numdoc"
                Else
                    RstTmp2.Sort = "registro"
                End If
                
                Do While Not RstTmp2.EOF
                    DoEvents
                    
                    '--SI SE NTERRUMPE EL PROCESO => SALIR
                    If BAND_INTERRUMPIR = True Then GoTo SALIR
                    '-----------------------------------------------
                    Fg1.Rows = Fg1.Rows + 1
                    mCol = 0
                    For mColCampo = 2 To RstTmp2.Fields.Count - 1
                        mCol = mCol + 1
                        Select Case LCase(RstTmp2.Fields(mColCampo).Name)
                            Case "libdesc", "registro", "registroref", "glosa", "numruc", "apenom", "tdocdesc", "docsustenta", "ctanum", "ctadesc", "simbolo"
                                Fg1.TextMatrix(Fg1.Rows - 1, mCol) = NulosC(RstTmp2.Fields(mColCampo))
                            Case "fchdoc", "fchope"
                                Fg1.TextMatrix(Fg1.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), FORMAT_DATE)
                            Case "tc"
                                Fg1.TextMatrix(Fg1.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), "0.000")
                            'Case "debe"
                            Case "debe", "impdebesol", "impdebedol":
                                Fg1.TextMatrix(Fg1.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), FORMAT_MONTO)
                                'xTotal1 = xTotal1 + NulosN(RstTmp2(mColCampo))
                                xTotal1 = xTotal1 + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, mCol))
                                
                                
                            'Case "haber"
                            Case "haber", "imphabersol", "imphaberdol":
                                Fg1.TextMatrix(Fg1.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), FORMAT_MONTO)
                                'xTotal2 = xTotal2 + NulosN(RstTmp2(mColCampo))
                                xTotal2 = xTotal2 + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, mCol))
                            'Case "saldo"
                            Case "saldo", "impsalsol", "impsaldol"

'''                                Fg1.TextMatrix(Fg1.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), FORMAT_MONTO)
'''                                'xSaldo = xSaldo + NulosN(RstTmp2(mColCampo))
'''                                xSaldo = xSaldo + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, mCol))

                                If UCase(RstTmp2.Fields("tipsal") & "") = "D" Or NulosC(RstTmp2.Fields("tipsal")) = "" Then
                                    xSaldo = xSaldo + Format((NulosN(RstTmp2(mColDebe + 1)) - NulosN(RstTmp2(mColHaber + 1))), FORMAT_MONTO)
                                Else
                                    xSaldo = xSaldo + Format((NulosN(RstTmp2(mColHaber + 1)) - NulosN(RstTmp2(mColDebe + 1))), FORMAT_MONTO)
                                End If
                                
                                Fg1.TextMatrix(Fg1.Rows - 1, mCol) = Format(xSaldo, FORMAT_MONTO)
                                

                                
                            Case Else
                                Fg1.TextMatrix(Fg1.Rows - 1, mCol) = NulosC(RstTmp2.Fields(mColCampo))
                        End Select
                        
                    Next mColCampo
                    
                    RstTmp2.MoveNext
                    If RstTmp2.EOF = True Then
                        Fg1.Rows = Fg1.Rows + 1
                        Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total =>"
                        Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xTotal1, FORMAT_MONTO)
                        Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xTotal2, FORMAT_MONTO)
                        
                        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1, , True
                        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe, , True
                        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber, , True
                        
                        Fg1.Rows = Fg1.Rows + 1
                        Exit Do
                    End If
                    xFila = xFila + 1
                Loop
            Else
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total =>"
                Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xTotal1, FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xTotal2, FORMAT_MONTO)
                
                FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1, , True
                FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe, , True
                FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber, , True
                '**************************************
                
                Fg1.Rows = Fg1.Rows + 1
            End If
            
            xTotal1_1 = xTotal1_1 + xTotal1
            xTotal2_1 = xTotal2_1 + xTotal2

            RstMay.MoveNext
            
        Loop
        
    Else
        xSaldo = 0

        'hallamos el saldo anterior de la cuenta

        nSQL = "SELECT con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion ,con_planctas.tipsal, " _
            + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[impdebdol]*[con_tc].[impven]),[con_diario].[impdebsol])) AS impdebsol, " _
            + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[imphabdol]*[con_tc].[impven]),[con_diario].[imphabsol])) AS imphabsol, " _
            + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol],IIf([con_diario].[impdebsol]=0 Or [con_tc].[impven] Is Null,0,[con_diario].[impdebsol]/[con_tc].[impven]))) AS impdebdol, " _
            + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol], IIf([con_diario].[imphabsol] = 0 Or [con_diario].[imphabsol] Is Null Or [con_tc].[impven] Is Null, 0, [con_diario].[imphabsol] / [con_tc].[impven]))) As imphabdol " _
            + vbCr + " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue "

        If opt_fecha(0).Value = True Then
            nSQL = nSQL + vbCr + " WHERE (  ( " & nSQLAjuste & " con_diario.fchasi Is Null and con_diario.año = " & AnoTra & " ) Or ( " & nSQLAjuste & " con_diario.fchasi < CDate('" & TxtFchIni.Valor & "')) ) "
        Else
            nSQL = nSQL + vbCr + " WHERE (  ( " & nSQLAjuste & " con_diario.idmes < " & mMesIni & " AND con_diario.año = " & AnoTra & " )  or ( " & nSQLAjuste & " con_diario.año < " & AnoTra & " )  ) "
        End If
        
        nSQL = nSQL + nSQLCuenta _
            + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion,con_planctas.tipsal ; " _
            + vbCr + " "
            
        RST_Busq RstSal, nSQL, xCon
        
        If RstSal.EOF = False Or RstSal.BOF = False Or RstSal.RecordCount <> 0 Then
            Fg1.Rows = Fg1.Rows + 1
            RstSal.MoveFirst
        
            Do While Not RstSal.EOF
                UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 5, "Cta Nº  :  " + NulosC(RstSal("cuenta")) & "   - " + NulosC(RstSal.Fields("descripcion")), flexAlignLeftCenter
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 1, , True
                
                If UCase(RstSal.Fields("tipsal") & "") = "D" Or RstSal.Fields("tipsal") = "" Then
                    xSaldo = (NulosN(RstSal(CAMPO_DEBE)) - NulosN(RstSal(CAMPO_HABER)))
                Else
                    xSaldo = (NulosN(RstSal(CAMPO_HABER)) - NulosN(RstSal(CAMPO_DEBE)))
                End If
                
                Fg1.TextMatrix(Fg1.Rows - 1, mColSaldo) = Format(xSaldo, FORMAT_MONTO)
                
                '**************************************
                Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(NulosN(RstSal(CAMPO_DEBE)), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(NulosN(RstSal(CAMPO_HABER)), FORMAT_MONTO)
                
                xTotal1 = NulosN(RstSal(CAMPO_DEBE))
                xTotal2 = NulosN(RstSal(CAMPO_HABER))

                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total =>"
                Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xTotal1, FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xTotal2, FORMAT_MONTO)
                
                FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1, , True
                FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe, , True
                FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber, , True
                
                xTotal1_1 = xTotal1_1 + xTotal1
                xTotal2_1 = xTotal2_1 + xTotal2
                
                '**************************************
                
                Fg1.Rows = Fg1.Rows + 1
                
                
''
''
''                Fg1.Rows = Fg1.Rows + 1
                
               
                RstSal.MoveNext
            Loop
        End If
        Set RstSal = Nothing
        '----------------------------
    End If

    If xTotal1_1 <> 0 Or xTotal2_1 <> 0 Then
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total Gral =>"
        Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xTotal1_1, FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xTotal2_1, FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1, , True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe, , True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber, , True
    End If
    
    '--ajustando las columnas de acuerdo a los importes
    Fg1.AutoSizeMode = flexAutoSizeColWidth
    Fg1.AutoSize mColDebe
    Fg1.AutoSize mColHaber
    Fg1.AutoSize mColSaldo
    

SALIR:
    Set RstMay = Nothing:     Set RstDet = Nothing:     Set RstSal = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    MuestraMayor1 = True
    Exit Function
error:
    Set RstMay = Nothing:     Set RstDet = Nothing:     Set RstSal = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    SHOW_ERROR Me.Name, "MuestraMayor1"
End Function



Private Sub pBuscarAsiento()
    Dim xfrm As New SGI2_funciones.formularios
    xfrm.AsientoBuscar xCon
    Set xfrm = Nothing
End Sub




Sub CargarResumen1()
    On Error GoTo error
    Dim RstRes As New ADODB.Recordset
    Dim A&
    Dim xTotal1, xTotal2, xTotal3, xTotal4 As Double
    Dim xAcumulado(7) As Double
    Dim nSQLAjuste  As String
    
    Frame5.Left = 3413
    Frame5.Top = 2685
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Value = 0
    
    Frame5.Visible = True
    Label3.Caption = "Procesando Resumen"
    Fg3.Rows = Fg3.FixedRows
    
    DoEvents
    
    Dim nSQLCuenta As String
    Dim nSQL As String
    '--------------------------
    nSQLCuenta = ""
    '--SI AGREGA CUENTAS AS GRID, GENERAR EL FILTRO A CONCATENAR A LA CONSULTA
    For A = 1 To Fg2.Rows - 1
        If Trim(Fg2.TextMatrix(A, 1)) <> "" Then
            nSQLCuenta = nSQLCuenta + " con_planctas.cuenta Like '" & Trim(Fg2.TextMatrix(A, 1)) & "%' OR "
        End If
    Next A
    If nSQLCuenta <> "" Then nSQLCuenta = " AND (" + Left(nSQLCuenta, Len(nSQLCuenta) - 3) + ") "
    '--------------------------

    '--para ajuste por diferencia de cambio
    nSQLAjuste = " AND (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    '-----------------------------------------------


    '**************************************************************
    'LEYENDA:
    'SI: Saldos Iniciales
    'MP: Movimientos del Periodo
    'SM: Sumas del Mayor
    'SA: Saldos Al
    'CB: Cuentas de Balance
    'CT: Cuentas de Transferencia
    'GN: Ganancias por Naturaleza
    'GF: Ganancias por Funcion

    
    '--19/04/09
    '--se cambia los saldos iniciales solo debera de mostrar debe o harer
    '-- IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol) AS SIDebSol,
    '-- IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol) AS SIHabSol,
    
    nSQL = "SELECT con_planctas.id as idcue, con_planctas.cuenta, con_planctas.descripcion AS descri , con_planctas.tipsal , " _
        + vbCr + " IIf(((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol)))>0,((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))),0) AS SIDebSol, " _
        + vbCr + " IIf(((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol)))>0,((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))),0) AS SIHabSol, " _
        + vbCr + " IIf(MovPeriodo.DebSol Is Null,0,MovPeriodo.DebSol) AS MPDeb, " _
        + vbCr + " IIf(MovPeriodo.HabSol Is Null,0,MovPeriodo.HabSol) AS MPHab, " _
        + vbCr + " [SIDeb]+[MPDeb] AS SMDeb,  " _
        + vbCr + " [SIHab]+[MPHab] AS SMHab, " _
        + vbCr + " IIf((SMDeb-SMHab)>0,(SMDeb-SMHab),0) AS SADeb, " _
        + vbCr + " IIf((SMHab-SMDeb)>0,(SMHab-SMDeb),0) AS SAHab, " _
        + vbCr + " con_planctas.iddes,con_planctas.iddes2,con_planctas.id AS IdCta "
    
    If NulosN(TxtIdMon.Text) = 2 Then
        nSQL = Replace(nSQL, "DebSol", "DebDol")
        nSQL = Replace(nSQL, "HabSol", "HabDol")
        
    End If
    '--movimientos del periodo
    nSQL = nSQL _
        + vbCr + " FROM (con_planctas LEFT JOIN " _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
        + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id=con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc=con_tc.fecha " _
        + vbCr + " WHERE (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLCuenta & nSQLAjuste _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS MovPeriodo ON con_planctas.id = MovPeriodo.IdCta) " _
        + vbCr + " Left Join "
    
    '--saldos iniciales
    nSQL = nSQL _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
        + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE (((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "'))) " & nSQLCuenta & nSQLAjuste _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS SaldosIni "




    nSQLAjuste = nSQLAjuste & " AND (  (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))  OR  (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')   )"


    nSQL = nSQL _
        + vbCr + " ON con_planctas.id = SaldosIni.IdCta " _
        + vbCr + " WHERE con_planctas.id In (SELECT con_diario.idcue FROM con_diario " & IIf(nSQLCuenta <> "", "WHERE " & Mid(nSQLCuenta, 5), "") & " ) " & nSQLCuenta _
        + vbCr + " ORDER BY con_planctas.cuenta; "
    
    '--si seleccionar por periodo
    If opt_fecha(1).Value = True Then
        nSQL = Replace(nSQL, "(((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))", "con_diario.idmes>=" & mMesIni & " And con_diario.idmes <= " & mMesFin)
        nSQL = Replace(nSQL, "(con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')", "con_diario.idmes < " & mMesIni)
        
    End If


'**************************************************************
    
    
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo SALIR:
    RstTmp.Filter = adFilterNone
    RstTmp.Sort = "cuenta"
    fra_msg.Visible = True
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    If RstTmp.RecordCount > 0 Then ProgressBar1.Max = RstTmp.RecordCount
    
    Label3.Caption = "Procesando Resumen"
    For A = 1 To RstTmp.RecordCount
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        '-----------------------------------------------
        ProgressBar1.Value = A
        Fg3.Rows = Fg3.Rows + 1
        Fg3.TextMatrix(A + 1, 1) = NulosC(RstTmp("cuenta"))
        
        Fg3.TextMatrix(A + 1, 2) = NulosC(RstTmp("descri"))
        
        Fg3.TextMatrix(A + 1, 3) = Format(NulosN(RstTmp("SIDeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 4) = Format(NulosN(RstTmp("SIHab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 5) = Format(NulosN(RstTmp("MPDeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 6) = Format(NulosN(RstTmp("MPHab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 7) = Format(NulosN(RstTmp("SMDeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 8) = Format(NulosN(RstTmp("SMHab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 9) = Format(NulosN(RstTmp("SADeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 10) = Format(NulosN(RstTmp("SAHab")), FORMAT_MONTO)
        
        
        xAcumulado(0) = xAcumulado(0) + NulosN(Fg3.TextMatrix(A + 1, 3)) '--saldeb
        xAcumulado(1) = xAcumulado(1) + NulosN(Fg3.TextMatrix(A + 1, 4)) '--salhab
        xAcumulado(2) = xAcumulado(2) + NulosN(Fg3.TextMatrix(A + 1, 5)) '--movdeb
        xAcumulado(3) = xAcumulado(3) + NulosN(Fg3.TextMatrix(A + 1, 6)) '--movhab
        xAcumulado(4) = xAcumulado(4) + NulosN(Fg3.TextMatrix(A + 1, 7)) '--maydeb
        xAcumulado(5) = xAcumulado(5) + NulosN(Fg3.TextMatrix(A + 1, 8)) '--mayhab
        xAcumulado(6) = xAcumulado(6) + NulosN(Fg3.TextMatrix(A + 1, 9)) '--deudor
        xAcumulado(7) = xAcumulado(7) + NulosN(Fg3.TextMatrix(A + 1, 10)) '--acreedor
        
        RstTmp.MoveNext
        If RstTmp.EOF = True Then Exit For
    Next A
    
    Fg3.Rows = Fg3.Rows + 2
    
    Fg3.TextMatrix(Fg3.Rows - 1, 2) = "TOTAL =>"
    '-------------------------------
    Dim Col&
    
    Fg3.AutoSizeMode = flexAutoSizeColWidth
    
    For A = 0 To UBound(xAcumulado())
        Fg3.TextMatrix(Fg3.Rows - 1, 3 + A) = Format(xAcumulado(A), FORMAT_MONTO)
        FORMATO_CELDA Fg3, Fg3.Rows - 1, 3 + A, , True
        Fg3.AutoSize 3 + A
    Next A
    Erase xAcumulado()
    '-------------------------------
    If RstTmp.RecordCount > 0 Then
        GRID_COLOR_FONDO Fg3, 2, 3, Fg3.Rows - 3, 4, RGB(255, 255, 236)
        GRID_COLOR_FONDO Fg3, 2, 7, Fg3.Rows - 3, 8, RGB(255, 255, 236)
        
    End If
        GRID_COLOR_FONDO Fg3, Fg3.Rows - 2, 1, Fg3.Rows - 1, Fg3.Cols - 1, RGB(231, 254, 224)
    
SALIR:
    Set RstRes = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    
    MsgBox "El Mayor se terminó de procesar con éxito", vbInformation, xTitulo
    
    Exit Sub
error:
    Frame5.Visible = False
    Set RstRes = Nothing
    fra_msg.Visible = False
    SHOW_ERROR Me.Name, "CargarResumen"
End Sub




