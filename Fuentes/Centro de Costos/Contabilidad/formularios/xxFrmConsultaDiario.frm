VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmConsultaDiario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Diario "
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
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
            Picture         =   "FrmConsultaDiario.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":277E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":2A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":2E2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
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
            Object.ToolTipText     =   "Ver Asientos Descuadrados"
            ImageIndex      =   13
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
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   795
      Left            =   2760
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   105
         TabIndex        =   15
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
         X1              =   0
         X2              =   5910
         Y1              =   780
         Y2              =   765
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   5940
         Y1              =   15
         Y2              =   30
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   5925
         X2              =   5925
         Y1              =   -15
         Y2              =   915
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
         Caption         =   "Procesando Diario"
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
         TabIndex        =   17
         Top             =   90
         Width           =   4020
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
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6300
      Left            =   0
      TabIndex        =   9
      Top             =   1275
      Width           =   11970
      _cx             =   21114
      _cy             =   11112
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
      Caption         =   "      Diario     | Resumen Cuentas "
      Align           =   0
      CurrTab         =   1
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
         Height          =   5880
         Left            =   45
         TabIndex        =   11
         Top             =   45
         Width           =   11880
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   5835
            Left            =   15
            TabIndex        =   12
            Top             =   30
            Width           =   11850
            _cx             =   20902
            _cy             =   10292
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConsultaDiario.frx":31BC
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
         Height          =   5880
         Left            =   -12525
         TabIndex        =   10
         Top             =   45
         Width           =   11880
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5820
            Left            =   15
            TabIndex        =   13
            Top             =   30
            Width           =   11775
            _cx             =   20770
            _cy             =   10266
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
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConsultaDiario.frx":326E
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
   Begin VB.Frame Frame4 
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
      Height          =   870
      Left            =   8175
      TabIndex        =   23
      Top             =   390
      Width           =   2040
      Begin VB.OptionButton OptLibro 
         Caption         =   "Por Libro"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   270
         Value           =   -1  'True
         Width           =   1110
      End
      Begin VB.OptionButton OptTodo 
         Caption         =   "Todos los Libros"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   555
         Width           =   1575
      End
      Begin VB.Label LblIdMes 
         AutoSize        =   -1  'True
         Caption         =   "LblIdMes"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1725
         TabIndex        =   27
         Top             =   525
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label LblIdLibro 
         AutoSize        =   -1  'True
         Caption         =   "LblIdLibro"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1725
         TabIndex        =   26
         Top             =   300
         Visible         =   0   'False
         Width           =   690
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "[ Moneda ]"
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
      Height          =   870
      Left            =   10275
      TabIndex        =   28
      Top             =   390
      Width           =   1575
      Begin VB.OptionButton OptDolares 
         Caption         =   "Dolares"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   555
         Width           =   900
      End
      Begin VB.OptionButton OptSoles 
         Caption         =   "Soles"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   29
         Top             =   270
         Value           =   -1  'True
         Width           =   900
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
      Height          =   870
      Left            =   60
      TabIndex        =   20
      Top             =   390
      Width           =   1965
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Periodo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   22
         Top             =   555
         Width           =   1125
      End
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Fecha"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   21
         Top             =   270
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   2070
      TabIndex        =   4
      Top             =   390
      Width           =   6060
      Begin VB.CommandButton cmd_periodo 
         Height          =   255
         Index           =   1
         Left            =   4455
         Picture         =   "FrmConsultaDiario.frx":343D
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   975
         Width           =   285
      End
      Begin VB.CommandButton cmd_periodo 
         Height          =   255
         Index           =   0
         Left            =   2340
         Picture         =   "FrmConsultaDiario.frx":37BF
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   975
         Width           =   285
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   1200
         _ExtentX        =   2117
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
         Valor           =   "05/12/2007"
      End
      Begin VB.CommandButton CmdBusProv 
         Height          =   240
         Left            =   5535
         Picture         =   "FrmConsultaDiario.frx":3B41
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   195
         Width           =   240
      End
      Begin VB.TextBox TxtLibro 
         Height          =   300
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "TxtLibro"
         Top             =   165
         Width           =   4845
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   3075
         TabIndex        =   2
         Top             =   480
         Width           =   1200
         _ExtentX        =   2117
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
         Valor           =   "05/12/2007"
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
         Left            =   3075
         TabIndex        =   32
         Top             =   945
         Width           =   1710
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
         Left            =   960
         TabIndex        =   19
         Top             =   945
         Width           =   1710
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   8
         Top             =   570
         Width           =   135
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Libro"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   240
         Width           =   345
      End
   End
End
Attribute VB_Name = "FrmConsultaDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xNumPag As Integer
Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE

Dim xAcumulado(2, 1) As Double
'--xAcumulado(0,?):: Acumulado por Asiento  ?::0=debe sol; 1::haber sol; 2::debe dol;  3::haber dol
'--xAcumulado(1,?):: Acumulado por libro
'--xAcumulado(2,?):: Acumulado general
Dim mMesIni As Integer
Dim mMesFin As Integer


Private Sub cmd_periodo_Click(Index As Integer)
    If Index = 0 Then
        mMesIni = SeleccionaMes(xCon)
        lbl_periodo(0).Caption = Busca_Codigo(mMesIni, "id", "descripcion", "con_meses", "N", xCon)
    Else
        mMesFin = SeleccionaMes(xCon)
        lbl_periodo(1).Caption = Busca_Codigo(mMesFin, "id", "descripcion", "con_meses", "N", xCon)
    End If
End Sub

Private Sub pDescuadrados()
    If OptLibro.Value = True Then
        If TxtLibro.Text = "" Then
            MsgBox "No ha especificado el libro a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtLibro.SetFocus
            Exit Sub
        End If
    End If
    
    If opt_fecha(0).Value = True Then '--solo pro fecha
        If NulosC(TxtFchIni.Valor) = "" Or NulosC(TxtFchFin.Valor) = "" Then
            MsgBox "El rango de fechas del periodo a consultar es invalido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If
        
        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio del periodo no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If
    Else
        If mMesIni > mMesFin Then
            MsgBox "El periodo de inicio debe ser inferior o igual al periodo final", vbExclamation, xTitulo
            cmd_periodo(0).SetFocus
            Exit Sub
        End If
    End If
    
    Me.MousePointer = vbHourglass
    If opt_fecha(0).Value = True Then '--por fecha
        If OptLibro.Value = True Then
            FrmDescuadrados.RECIBE_LINK_FRM TxtFchIni.Valor, TxtFchFin.Valor, 0, 0, True, NulosN(LblIdLibro.Caption)
        Else
            FrmDescuadrados.RECIBE_LINK_FRM TxtFchIni.Valor, TxtFchFin.Valor, 0, 0
        End If
    Else '--por periodo
        If OptLibro.Value = True Then
            FrmDescuadrados.RECIBE_LINK_FRM Date, Date, mMesIni, mMesFin, False, NulosN(LblIdLibro.Caption)
        Else
            FrmDescuadrados.RECIBE_LINK_FRM Date, Date, mMesIni, mMesFin, False, 0
        End If
    End If
    
    FrmDescuadrados.Show
    Me.MousePointer = vbDefault
End Sub


Private Sub CmdBusProv_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SqlCad = "SELECT * FROM mae_libros  where activo = -1 ORDER BY descripcion "
    
    xform.Titulo = "Buscando Libro Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtLibro.Text = NulosC(xRs("descripcion"))
        LblIdLibro.Caption = NulosC(xRs("id"))
        If TxtFchIni.Visible = True Then
            TxtFchIni.SetFocus
        Else
            cmd_periodo(0).SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub pExportar()
    Dim xFun As New SGI2_funciones.formularios
    Dim Rst As New ADODB.Recordset
    
    If TabOne1.CurrTab = 0 Then
        If Fg1.Rows = 1 Then
            MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        'ExportarEcelDiario
        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "LIBRO DIARIO", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Por Fecha", "diario.xls"  ', Rst, ""
        Set xFun = Nothing
    End If
    If TabOne1.CurrTab = 1 Then
        If Fg2.Rows = 2 Then
            MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
'        ExportarExcelResumen
        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "RESUMEN DEL DIARIO", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Por Fecha", "diario.xls"  ', Rst, ""
        Set xFun = Nothing
    End If
End Sub

Private Sub pImprimir()
    If Me.TabOne1.CurrTab = 0 Then
        If Fg1.Rows = 1 Then
            MsgBox "No hay registros para imprimir", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        FrmPrintDiario.Show
    Else
        Dim xMoneda As String
        Dim xPrint As New SGI2_funciones.formularios
        Dim nPeriodo   As String
        
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

        If OptSoles.Value = True Then
            xMoneda = "Nuevos Soles"
        Else
            xMoneda = "Dolares Americanos"
        End If
        Me.MousePointer = vbHourglass
        xPrint.Imprimir_x_VSFlexGrid Fg2, "LIBRO DIARIO ", "(Expresado en " + xMoneda + ")", nPeriodo, False, True
        Set xPrint = Nothing
        Me.MousePointer = vbDefault
    End If
    
End Sub

Private Sub pConsultar()
    If OptLibro.Value = True Then
        If TxtLibro.Text = "" Then
            MsgBox "No ha especificado el libro a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtLibro.SetFocus
            Exit Sub
        End If
    End If
    ''''''''''''
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Sub
    End If
    
    If opt_fecha(0).Value = True Then '--por fecha
        If NulosC(TxtFchIni.Valor) = "" Then
            MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If
        If NulosC(TxtFchFin.Valor) = "" Then
            MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchFin.SetFocus
            Exit Sub
        End If
        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If
        If (Year(TxtFchIni.Valor) <> Year(TxtFchFin.Valor)) Then
            MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        ElseIf Year(TxtFchIni.Valor) <> CStr(AnoTra) Then
            MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If
    Else '--por periodo
        If mMesIni > mMesFin Then
            MsgBox "El periodo de inicio debe ser inferior o igual al periodo final", vbExclamation, xTitulo
            cmd_periodo(0).SetFocus
            Exit Sub
        End If
    End If
    
    '''''''''''
    Erase xAcumulado()

    '''''''''''
    BAND_INTERRUMPIR = False
    'pConfigurarGrilla
    
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    If OptTodo.Value = True Then
        RST_Busq Rst, "SELECT * FROM mae_libros where activo = -1  ORDER BY id", xCon
    Else
        RST_Busq Rst, "SELECT * FROM mae_libros WHERE id = " & NulosN(LblIdLibro.Caption) & "", xCon
    End If
    Me.TabOne1.CurrTab = 0
    
    Frame5.Visible = True
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Fg1.Rows = 2
        For A = 1 To Rst.RecordCount
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo SALIR:
            '-----------------------------------------------
            Fg1.Rows = Fg1.Rows + 1
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, "LIBRO:   " + UCase(Trim(Rst("descripcion") & "")), flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 1
            
            Label3.Caption = "Procesando: " + Trim(Rst("descripcion"))
            
            Fg1.Rows = Fg1.Rows + 1
            If OptTodo.Value = True Then
                ProcesarDiario Rst("id"), "Libro " + UCase(Trim(Rst("descripcion"))), 1
            Else
                ProcesarDiario Val(LblIdLibro.Caption), "Libro " + Trim(Rst("descripcion") & ""), 1
            End If
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    
    '---------
    '--MOSTRAR LOS ACUMULADOS POR LIBRO
    'If OptLibro.Value = False And (xAcumulado(2, 0) <> 0 Or xAcumulado(2, 1) <> 0) Then
    If (xAcumulado(2, 0) <> 0 Or xAcumulado(2, 1) <> 0) Then
        Fg1.Rows = Fg1.Rows + 2
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = "Total Gen.==>"
        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(xAcumulado(2, 0), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xAcumulado(2, 1), FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 7
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10

    End If
    Erase xAcumulado()
    Me.TabOne1.CurrTab = 1
    '--SI SE NTERRUMPE EL PROCESO => SALIR
    If BAND_INTERRUMPIR = True Then GoTo SALIR:
    '-----------------------------------------------
    Label3.Caption = "Procesando Resumen "
    ProcesarResumen NulosN(LblIdLibro.Caption)
    Frame5.Visible = False
    Exit Sub
    
SALIR:
    Frame5.Visible = False
    Erase xAcumulado()
    If BAND_INTERRUMPIR = True Then
        MsgBox "La consulta fue interrumpida", vbInformation, xTitulo
    End If
End Sub


Sub ProcesarResumen(IDLIBRO As Integer)
    Dim RstRes As New ADODB.Recordset
    Dim N_SQL_WHERE, N_SQL_WHERE1, N_SQL_SALDO, N_SQL_SALDO1 As String
    Dim N_SQL As String
    Dim xAcumulado(0, 1) As Double
    Erase xAcumulado()

    If IDLIBRO <> 0 Then
        N_SQL = "SELECT con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion, Sum(con_diario.impdebsol) AS debe, Sum(con_diario.imphabsol) AS haber " _
            & " FROM con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion, " _
            & " con_diario.idlib, con_diario.fchasi HAVING (((con_diario.idlib)=" & IDLIBRO & ") AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))) ORDER BY con_planctas.cuenta"
    Else
        N_SQL = "SELECT con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion, Sum(con_diario.impdebsol) AS debe, Sum(con_diario.imphabsol) AS haber " _
            & " FROM con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion, " _
            & " con_diario.fchasi HAVING (((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))) " _
            & " ORDER BY con_planctas.cuenta"
    End If
    
    RST_Busq RstRes, N_SQL, xCon
    
    Erase xAcumulado()
    If RstRes.State = 0 Then GoTo SALIR
    If RstRes.BOF = True Or RstRes.EOF = True Or RstRes.RecordCount = 0 Then GoTo SALIR
    RstRes.MoveFirst
    ProgressBar1.Min = 1
    If RstRes.RecordCount <> 0 Then
        If RstRes.RecordCount > 1 Then ProgressBar1.Max = RstRes.RecordCount
    End If
    Do While Not RstRes.EOF
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        ProgressBar1.Value = RstRes.Bookmark
        '-----------------------------------------------
        DoEvents
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = RstRes("cuenta") & ""
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = RstRes("descripcion") & ""

        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(NulosN(RstRes.Fields("debe")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(RstRes.Fields("haber")), FORMAT_MONTO)
        
        xAcumulado(0, 0) = xAcumulado(0, 0) + NulosN(Format(NulosN(RstRes.Fields("debe")), FORMAT_MONTO))
        xAcumulado(0, 1) = xAcumulado(0, 1) + NulosN(Format(NulosN(RstRes.Fields("haber")), FORMAT_MONTO))
        
        If NulosN(RstRes.Fields("debe")) = 0 And NulosN(RstRes.Fields("haber")) = 0 Then
            Fg2.RemoveItem Fg2.Rows - 1
        End If
        RstRes.MoveNext
        
        If RstRes.EOF = True Then
            Fg2.Rows = Fg2.Rows + 2
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = "TOTAL ==>"
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            FORMATO_CELDA Fg2, Fg2.Rows - 1, 2
            FORMATO_CELDA Fg2, Fg2.Rows - 1, 3
            FORMATO_CELDA Fg2, Fg2.Rows - 1, 4
            
        End If
    Loop
    
    Exit Sub
SALIR:
    MsgBox "El Diario se terminó de procesar con éxito", vbInformation, xTitulo
    Erase xAcumulado()
    Set RstRes = Nothing

End Sub

Sub ProcesarResumen2(IDLIBRO As Integer)
    Dim RstRes As New ADODB.Recordset
    Dim N_SQL_WHERE, N_SQL_WHERE1, N_SQL_SALDO, N_SQL_SALDO1 As String
    Dim N_SQL As String
    Dim xAcumulado(0, 1) As Double
    Erase xAcumulado()
    '--xAcumulado(0,?):: Acumulado por Asiento  ?::0=debe sol; 1::haber sol; 2::debe dol;  3::haber dol
    '--xAcumulado(1,?):: Acumulado por libro
    '--xAcumulado(2,?):: Acumulado general
    
    '---DEL RESUMEN
    If NulosN(IDLIBRO) <> 0 And OptLibro.Value = True Then
        N_SQL_WHERE = " AND ( con_diario.idlib = " + CStr(IDLIBRO) + ")"
        N_SQL_WHERE1 = " AND ( con_diario1.idlib = " + CStr(IDLIBRO) + ")"
    End If

    N_SQL_SALDO = ""
    If opt_fecha(0).Value = True Then '--por fecha
        If CDate(Me.TxtFchIni.Valor) = CDate("01/01/" + AnoTra) Then
            N_SQL_WHERE = Replace(N_SQL_WHERE, "AND", "")
            N_SQL_SALDO1 = " OR ( (con_diario1.fchasi) IS NULL " + N_SQL_WHERE1 + " ) "
        Else
            N_SQL_SALDO = " (con_diario.fchasi) IS NOT NULL "
            N_SQL_SALDO1 = " AND ( (con_diario1.fchasi) IS NOT NULL ) "
        End If
    Else '--por intervalo
        N_SQL_WHERE = Replace(N_SQL_WHERE, "AND", "")
        N_SQL_SALDO1 = " "
    End If
        
    If N_SQL_SALDO <> "" Or N_SQL_WHERE <> "" Then
        N_SQL_WHERE = " WHERE " + N_SQL_SALDO + N_SQL_WHERE
    End If
    
    N_SQL = "SELECT con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion, "
    If Me.OptSoles = True Then
        N_SQL = N_SQL _
        + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol=0,con_diario1.impdebsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.impdebdol))))  AS debesol   FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue  WHERE ( con_diario1.fchasi IS NOT NULL AND (con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')) " + N_SQL_WHERE1 + " ) " + N_SQL_SALDO1 + " GROUP BY con_diario1.idcue  HAVING con_diario1.idcue=con_diario.idcue ) AS impdebesol, " _
        + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol=0,con_diario1.imphabsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.imphabdol))))  AS habersol  FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue  WHERE ( con_diario1.fchasi IS NOT NULL AND (con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')) " + N_SQL_WHERE1 + " ) " + N_SQL_SALDO1 + " GROUP BY con_diario1.idcue  HAVING con_diario1.idcue=con_diario.idcue ) AS imphabersol "
    Else
        N_SQL = N_SQL _
        + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol<>0,con_diario1.impdebdol,IIf(con_tc1.impven Is Null Or con_diario1.impdebsol=0,0,(con_diario1.impdebsol/con_tc1.impven)))) AS debedol  FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue  WHERE ( con_diario1.fchasi IS NOT NULL AND (con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')) " + N_SQL_WHERE1 + " ) " + N_SQL_SALDO1 + " GROUP BY con_diario1.idcue HAVING con_diario1.idcue=con_diario.idcue ) AS impdebedol, " _
        + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol<>0,con_diario1.imphabdol,IIf(con_tc1.impven Is Null Or con_diario1.imphabsol=0,0,(con_diario1.imphabsol/con_tc1.impven)))) AS haberdol  FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue  WHERE ( con_diario1.fchasi IS NOT NULL AND (con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')) " + N_SQL_WHERE1 + " ) " + N_SQL_SALDO1 + " GROUP BY con_diario1.idcue HAVING con_diario1.idcue=con_diario.idcue ) AS imphaberdol "
    End If
    N_SQL = N_SQL _
    + vbCr + " FROM con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue " _
    + vbCr + N_SQL_WHERE _
    + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion " _
    + vbCr + " ORDER BY con_planctas.cuenta, con_planctas.descripcion; "

       '--REEMPLAZANDO EL INTERVALO DE FECHA
    If opt_fecha(0).Value = True Then
        N_SQL = Replace(N_SQL, "con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')", " con_diario1.fchasi >=CDate('" + Me.TxtFchIni.Valor + "') And con_diario1.fchasi <= CDate('" + Me.TxtFchFin.Valor + "') ")
    Else
        N_SQL = Replace(N_SQL, "con_diario1.fchasi IS NOT NULL AND ", " ")
        
        N_SQL = Replace(N_SQL, "con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')", " ( con_diario1.idmes >= " & mMesIni & " AND con_diario1.idmes <= " & mMesFin & " ) ")
        
    End If
    N_SQL = Replace(N_SQL, "impdebesol", "debe")
    N_SQL = Replace(N_SQL, "imphabersol", "haber")
    N_SQL = Replace(N_SQL, "impdebedol", "debe")
    N_SQL = Replace(N_SQL, "imphaberdol", "haber")
    
    RST_Busq RstRes, N_SQL, xCon
    
    Erase xAcumulado()
    If RstRes.State = 0 Then GoTo SALIR
    If RstRes.BOF = True Or RstRes.EOF = True Or RstRes.RecordCount = 0 Then GoTo SALIR
    RstRes.MoveFirst
    ProgressBar1.Min = 1
    If RstRes.RecordCount <> 0 Then
        If RstRes.RecordCount > 1 Then ProgressBar1.Max = RstRes.RecordCount
    End If
    Do While Not RstRes.EOF
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        ProgressBar1.Value = RstRes.Bookmark
        '-----------------------------------------------
        DoEvents
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = RstRes("cuenta") & ""
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = RstRes("descripcion") & ""

        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(NulosN(RstRes.Fields("debe")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(RstRes.Fields("haber")), FORMAT_MONTO)
        
        xAcumulado(0, 0) = xAcumulado(0, 0) + NulosN(Format(NulosN(RstRes.Fields("debe")), FORMAT_MONTO))
        xAcumulado(0, 1) = xAcumulado(0, 1) + NulosN(Format(NulosN(RstRes.Fields("haber")), FORMAT_MONTO))
        
        If NulosN(RstRes.Fields("debe")) = 0 And NulosN(RstRes.Fields("haber")) = 0 Then
            Fg2.RemoveItem Fg2.Rows - 1
        End If
        RstRes.MoveNext
        
        If RstRes.EOF = True Then
            Fg2.Rows = Fg2.Rows + 2
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = "TOTAL ==>"
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            FORMATO_CELDA Fg2, Fg2.Rows - 1, 2
            FORMATO_CELDA Fg2, Fg2.Rows - 1, 3
            FORMATO_CELDA Fg2, Fg2.Rows - 1, 4
            
        End If
    Loop
SALIR:
    MsgBox "El Diario se terminó de procesar con éxito", vbInformation, xTitulo
    Erase xAcumulado()
    Set RstRes = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        OptLibro.Value = True
        LimpiaText lbl_periodo
        lbl_periodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
        lbl_periodo(1).Caption = lbl_periodo(0).Caption
        mMesIni = xMes
        mMesFin = xMes
        TabOne1.CurrTab = 0
        TxtLibro.SetFocus
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
    Blanquea
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    OptSoles.Value = True
    Frame2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDNone
    Fg2.Editable = flexEDNone
    
    SetearCuadricula Fg1, 1, xCon
End Sub

Sub Blanquea()
    TxtLibro.Text = ""
    LblIdLibro.Caption = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
End Sub

Private Sub opt_fecha_Click(Index As Integer)
    If Index = 0 Then '--por fecha
        TxtFchFin.Visible = True
        TxtFchIni.Visible = True
        lbl(1).Visible = True
        lbl(0).Caption = "Del"
        lbl(0).Caption = "Al"
        Ocultar cmd_periodo, False
        Ocultar lbl_periodo, False
        
    Else '--por periodo
          
        TxtFchFin.Visible = False
        TxtFchIni.Visible = False
        lbl(0).Caption = "De"
        lbl(1).Caption = "A"
        cmd_periodo(0).Top = 525
        lbl_periodo(0).Top = 495
        cmd_periodo(1).Top = 525
        lbl_periodo(1).Top = 495
        Ocultar cmd_periodo, True
        Ocultar lbl_periodo, True

    End If
End Sub

Private Sub OptLibro_Click()
    If OptLibro.Value = True Then
        CmdBusProv.Enabled = True
        TxtLibro.Enabled = True
    End If
End Sub

Private Sub OptTodo_Click()
    If OptTodo.Value = True Then
        TxtLibro.Text = ""
        LblIdLibro.Caption = ""
        CmdBusProv.Enabled = False
        TxtLibro.Enabled = False
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then pDescuadrados
    If Button.Index = 5 Then pExportar
    If Button.Index = 6 Then pImprimir
    If Button.Index = 7 Then Configurar
    If Button.Index = 9 Then
        Unload Me
    End If
End Sub

Sub Configurar()
    Dim xform As New SGI2_funciones.Varias
    If xform.CambioOpcionLiro(1, xCon) = True Then
        SetearCuadricula Fg1, 1, xCon
        pConsultar
    End If
    Set xform = Nothing
End Sub

Private Sub TxtLibro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtLibro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Sub ExportarEcelDiario()
    Dim A&, B&, xFilas&
    On Error GoTo error
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Add  'Trim(App.Path) + "\RegCompras.xls"
    
    Label3.Caption = "Exportando Documentos"
    
    With objExcel.ActiveSheet
        
        .Cells(1, 2) = NomEmp
        .Cells(1, 13) = Date
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
        If Me.OptSoles.Value = True Then
            nTitulo1 = "(Expresado en Nuevos Soles)"
        Else
            nTitulo1 = "(Expresado en Dolares Americanos)"
        End If
        .Cells(4, 2) = "Libro Diario"
        .Cells(5, 2) = nPeriodo
        .Cells(6, 2) = nTitulo1
        '**********************************
        
        If Fg1.ColWidth(1) <> 0 Then .Columns(2).ColumnWidth = Fg1.ColWidth(1) / 100
        If Fg1.ColWidth(2) <> 0 Then .Columns(3).ColumnWidth = Fg1.ColWidth(2) / 100
        If Fg1.ColWidth(3) <> 0 Then .Columns(4).ColumnWidth = Fg1.ColWidth(3) / 100
        If Fg1.ColWidth(4) <> 0 Then .Columns(5).ColumnWidth = Fg1.ColWidth(4) / 100
        If Fg1.ColWidth(5) <> 0 Then .Columns(6).ColumnWidth = Fg1.ColWidth(5) / 100
        If Fg1.ColWidth(6) <> 0 Then .Columns(7).ColumnWidth = Fg1.ColWidth(6) / 100
        If Fg1.ColWidth(7) <> 0 Then .Columns(8).ColumnWidth = Fg1.ColWidth(7) / 100
        If Fg1.ColWidth(8) <> 0 Then .Columns(9).ColumnWidth = Fg1.ColWidth(8) / 100
        If Fg1.ColWidth(9) <> 0 Then .Columns(10).ColumnWidth = Fg1.ColWidth(9) / 100
        If Fg1.ColWidth(10) <> 0 Then .Columns(11).ColumnWidth = Fg1.ColWidth(10) / 100
        
        xFilas = 8
        For B = 1 To Fg1.Cols - 1
            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(0, B)
        Next B
        
        xFilas = xFilas + 1
        For A = 1 To Fg1.Rows - 1
            DoEvents
            For B = 1 To Fg1.Cols - 1
                If B <= 8 Then
                    If B = 1 Then
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                        If InStr(Fg1.TextMatrix(A, B), "LIBRO:  ") <> 0 Then
                            .Cells(xFilas, 2) = "'" + Fg1.TextMatrix(A, B)
                            GoTo SIG_FIL
                        End If
                    Else
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                    End If
                Else
                    If IsNumeric(Fg1.TextMatrix(A, B)) = True Then
                        .Cells(xFilas, B + 1) = NulosN(Fg1.TextMatrix(A, B))
                    End If
                End If
            Next B

SIG_FIL:
            xFilas = xFilas + 1
        Next A
    End With
    
    MsgBox "El proceso de exportación terminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "ExportarExcelDetalle", , IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
End Sub

Private Sub ExportarExcelResumen()
    On Error GoTo error
    Dim xExport As New SGI2_funciones.formularios
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
    If Me.OptSoles.Value = True Then
        nTitulo1 = "(Expresado en Nuevos Soles)"
    Else
        nTitulo1 = "(Expresado en Dolares Americanos)"
    End If
    
    
    xExport.VSFlexGrid_Exportar_MSExcel xCon, Fg2, "RESUMEN DEL DIARIO", nPeriodo, nTitulo1, "Resumen del Diario"
    Set xExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub


Private Sub BuscarVSFlexGrid()
    On Error GoTo error
    
    If Me.TabOne1.CurrTab <> 0 Then Exit Sub
    Dim xExport As New SGI2_funciones.formularios
    Dim xCampos(3, 3) As String
    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
    xCampos(0, 0) = "Num.Reg.":     xCampos(0, 1) = "1":    xCampos(0, 2) = "C":    xCampos(0, 3) = "-1"
    xCampos(1, 0) = "Fch. Doc":     xCampos(1, 1) = "3":    xCampos(1, 2) = "F":    xCampos(1, 3) = "0"
    xCampos(2, 0) = "Nº Documento": xCampos(2, 1) = "4":    xCampos(2, 2) = "C":    xCampos(2, 3) = "0"
    xCampos(3, 0) = "Nº. Cuenta":    xCampos(3, 1) = "5":    xCampos(3, 2) = "C":    xCampos(3, 3) = "0"
    
    xExport.VSFlexGrid_Buscar Me.hWnd, Fg1, xCampos(), Fg1.Row
    Set xExport = Nothing
    Me.MousePointer = vbDefault
    
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "BuscarVSFlexGrid"
End Sub
    
Function BuscarCampo(IdLib As Integer, Tipo As Integer, IdModulo As Integer, IdDocPro As Integer, IdMovimiento As Integer, IdOriDes As Integer) As ADODB.Recordset
    Dim Rst As New ADODB.Recordset
    Dim SeDetalla As Boolean
    
    
    If IdLib = 1 Then
        'compras
        RST_Busq Rst, "SELECT mae_prov.numruc, mae_prov.nombre AS apenom, mae_documento.abrev, con_tc.impven, com_compras.fchdoc, com_compras.glosa, com_compras.fchdoc as fchope," _
            & " Mid([com_compras]![numreg],1,2)+[mae_libros]![codsun]+Mid([com_compras]![numreg],3,4) AS numreg, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc " _
            & " FROM mae_libros RIGHT JOIN (mae_documento RIGHT JOIN (mae_prov RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) " _
            & " ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) ON mae_libros.id = com_compras.idlib " _
            & " WHERE (((com_compras.id)=" & IdMovimiento & ") AND ((con_tc.idmon)=2))", xCon
    
    End If
    
    If IdLib = 2 Then
        'ventas
        RST_Busq Rst, "SELECT mae_cliente.numruc, mae_cliente.nombre AS apenom, mae_documento.abrev, vta_ventas.fchdoc, con_tc.impven, vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc," _
            & " Mid(vta_ventas!numreg,1,2)+mae_libros!codsun+Mid(vta_ventas!numreg,3,4) AS numreg, vta_ventas.fchdoc AS fchope, vta_ventas.glosa " _
            & " FROM ((mae_libros RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_libros.id = vta_ventas.idlib) " _
            & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
            & " WHERE (((vta_ventas.id)=" & IdMovimiento & ") AND ((con_tc.idmon)=2))", xCon
    End If
    
    If IdLib = 5 Then
        'retenciones
        RST_Busq Rst, "SELECT mae_cliente.numruc, mae_cliente.nombre AS apenom, mae_documento.abrev, vta_ventas.fchdoc, con_tc.impven, " _
            & " [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, Mid([vta_ventas]![numreg],1,2)+[mae_libros]![codsun]+Mid([vta_ventas]![numreg],3,4) AS numreg, " _
            & " vta_ventas.fchdoc AS fchope, vta_ventas.glosa FROM (((((con_retencion LEFT JOIN con_retenciondet ON con_retencion.id = con_retenciondet.id) " _
            & " LEFT JOIN vta_ventas ON con_retenciondet.iddoc = vta_ventas.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) " _
            & " LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_libros " _
            & " ON vta_ventas.idlib = mae_libros.id WHERE (((con_retencion.id)=" & IdMovimiento & ") AND ((con_tc.idmon)=2) AND ((con_retenciondet.iddoc)=" & IdDocPro & "))", xCon
    End If
    
    If IdLib = 6 Then
        If Tipo = 1 Then
            SeDetalla = DetallarModulo(IdOriDes, origen, xCon)
        Else
            SeDetalla = DetallarModulo(IdOriDes, destino, xCon)
        End If
        
        If Tipo = 1 Then
            If SeDetalla = False Then
                Set Rst = Nothing
            Else
                If IdModulo = 6 Then
                    RST_Busq Rst, "SELECT tes_documentos.abrev, [apepat]+[apemat]+'. '+[nom] AS apenom, IIf(IsNull([tes_cajaorigendet].[numser])=-1,[tes_cajaorigendet].[numdoc],[tes_cajaorigendet].[numser]+'-'+[tes_cajaorigendet].[numdoc]) AS numdoc, " _
                        & " tes_caja.fchope AS fchdoc, '' AS numreg, '' AS numruc, tes_caja.glosa, tes_caja.fchope, con_tc.impven FROM ((tes_documentos LEFT JOIN " _
                        & " ((tes_cajaorigendet LEFT JOIN tes_usuarios ON tes_cajaorigendet.idper = tes_usuarios.id) LEFT JOIN pla_empleados ON tes_usuarios.idper = pla_empleados.id) " _
                        & " ON tes_documentos.id = tes_cajaorigendet.tipdoc) LEFT JOIN tes_caja ON tes_cajaorigendet.idtes = tes_caja.id) LEFT JOIN con_tc ON " _
                        & " tes_caja.fchope = con_tc.fecha WHERE (((tes_cajaorigendet.idtes)=" & IdMovimiento & ") AND ((tes_cajaorigendet.idori)=" & IdOriDes & ") AND ((con_tc.idmon)=2))", xCon
                End If
            End If
        End If
        If Tipo = 2 Then
            If SeDetalla = False Then
                'Set Rst = Nothing
                RST_Busq Rst, "SELECT tes_caja.fchope, tes_caja.glosa, '' AS abrev, '' AS apenom, '' AS numdoc, '' AS numruc, '' AS numreg, '' AS fchdoc, con_tc.impven " _
                    & " FROM (tes_caja RIGHT JOIN tes_cajadestino ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha " _
                    & " WHERE (((tes_cajadestino.idtes)=" & IdMovimiento & ") AND ((tes_cajadestino.iddes)=" & IdOriDes & "))", xCon
            Else
                If IdModulo = 1 Then 'procesamos los documentos de compra
                    RST_Busq Rst, "SELECT con_diario.iddocpro, mae_documento.abrev, mae_prov.nombre AS apenom, IIf(IsNull([com_compras]![numser])=-1,[com_compras]![numdoc],[com_compras]![numser]+'-'+[com_compras]![numdoc]) AS numdoc, " _
                        & " com_compras.fchdoc, mae_prov.numruc, Mid([com_compras]![numreg],1,2)+[mae_libros]![codsun]+Mid([com_compras]![numreg],3,4) AS numreg, " _
                        & " tes_caja.fchope, tes_caja.glosa, con_tc.impven FROM ((con_diario LEFT JOIN tes_caja ON con_diario.idmov = tes_caja.id) LEFT JOIN " _
                        & " (((com_compras LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id) " _
                        & " LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON con_diario.iddocpro = com_compras.id) LEFT JOIN con_tc ON " _
                        & " tes_caja.fchope = con_tc.fecha WHERE (((con_diario.iddocpro)=" & IdDocPro & ") AND ((con_diario.idmov)=" & IdMovimiento & ") AND ((con_tc.idmon)=2))", xCon
                End If
                If IdModulo = 7 Then
                    RST_Busq Rst, "SELECT tes_documentos.abrev, IIf(IsNull([tes_cajadestinodet]![numser])=-1,[tes_cajadestinodet]![numdoc],[tes_cajadestinodet]![numser]+'-'+[tes_cajadestinodet]![numdoc]) AS numdoc, " _
                        & " tes_caja.fchope AS fchdoc, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS apenom, '' AS numreg, " _
                        & " [pla_empleados]![numdoc] AS numruc, tes_caja.fchope, tes_caja.glosa, con_tc.impven FROM (((tes_cajadestinodet LEFT JOIN tes_caja ON tes_cajadestinodet.idtes = tes_caja.id) " _
                        & " LEFT JOIN (tes_usuarios LEFT JOIN pla_empleados ON tes_usuarios.idper = pla_empleados.id) ON tes_cajadestinodet.idper = tes_usuarios.id) " _
                        & " LEFT JOIN tes_documentos ON tes_cajadestinodet.tipdoc = tes_documentos.id) LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha " _
                        & " WHERE (((tes_cajadestinodet.idtes)=" & IdMovimiento & ") AND ((tes_cajadestinodet.iddes)=" & IdOriDes & ") AND ((con_tc.idmon)=2))", xCon
                End If
            
                If IdModulo = 2 Then  'procesamos los documentos de venta
                    RST_Busq Rst, "SELECT con_diario.iddocpro, mae_documento.abrev, mae_cliente.nombre AS apenom, IIf(IsNull([vta_ventas]![numser])=-1,[vta_ventas]![numdoc],[vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]) AS numdoc, " _
                        & " mae_cliente.numruc, Mid([vta_ventas]![numreg],1,2)+[mae_libros]![codsun]+Mid([vta_ventas]![numreg],3,4) AS numreg, tes_caja.fchope, " _
                        & " tes_caja.glosa, con_diario.idmov, vta_ventas.fchdoc, con_tc.impven FROM ((((con_diario LEFT JOIN tes_caja ON con_diario.idmov = tes_caja.id) " _
                        & " LEFT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON con_diario.iddocpro = vta_ventas.id) " _
                        & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
                        & " LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha WHERE (((con_diario.iddocpro)=" & IdDocPro & ") AND ((con_diario.idmov)=" & IdMovimiento & ") AND ((con_tc.idmon)=2))", xCon
                End If
            End If
        End If
    End If
    
    Set BuscarCampo = Rst
    Set Rst = Nothing
End Function

Sub ProcesarDiario(IDLIBRO As Integer, N_Libro As String, Tipo As Integer)
    Dim Rst As New ADODB.Recordset
    Dim CadSql As String
    
    If Tipo = 1 Then
        CadSql = "SELECT con_diario.idlib, mae_libros.descripcion AS nomlib, con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion, con_diario.idmov, " _
            & " Format(con_diario!idmes,'00')+mae_libros!codsun+con_diario!numasi AS numreg, 0 AS impven, '' AS fchemi, 'aa' AS tipdoc, 'bb' AS numdoc, " _
            & " con_diario.impdebsol, con_diario.imphabsol, con_diario.impdebdol, con_diario.imphabdol, con_diario.idorides, con_diario.idmod, con_diario.iddocpro, " _
            & " con_diario.tipo, tes_caja.tipmov, mae_tipomov.descripcion AS tipmov2 FROM mae_libros RIGHT JOIN ((tes_caja RIGHT JOIN (con_planctas RIGHT JOIN " _
            & " con_diario ON con_planctas.id = con_diario.idcue) ON tes_caja.id = con_diario.idmov) LEFT JOIN mae_tipomov ON tes_caja.tipmov = mae_tipomov.id) " _
            & " ON mae_libros.id = con_diario.idlib WHERE (((con_diario.idlib)=" & IDLIBRO & ") AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))) ORDER BY Format(con_diario!idmes,'00')+mae_libros!codsun+con_diario!numasi"
    Else
    End If

    RST_Busq Rst, CadSql, xCon
'    Rst.Filter = "numreg  ='02010007'"
    
    '--LIMPIAR ACUMULADO POR LIBRO
    xAcumulado(0, 0) = 0:   xAcumulado(0, 1) = 0
    xAcumulado(1, 0) = 0:   xAcumulado(1, 1) = 0
    '---------------------------------------------------------------------------
    If Rst.State = 0 Then GoTo SALIR
    If Rst.BOF = True Or Rst.EOF = True Or Rst.RecordCount = 0 Then GoTo SALIR
    Rst.MoveFirst
    Dim xAsiento
    Dim HabDol, DebDol  As Double
    
    xAsiento = NulosC(Rst.Fields("numreg"))
    
    Dim Cambiar As Boolean
    Dim RstDatoAdicional As New ADODB.Recordset
    
    If Rst.RecordCount > 1 Then
        ProgressBar1.Max = Rst.RecordCount
    End If
    Do While Not Rst.EOF
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        ProgressBar1.Value = Rst.Bookmark
        '-----------------------------------------------
        Fg1.Rows = Fg1.Rows + 1
        
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst.Fields("numreg"))
        
        Set RstDatoAdicional = Nothing
        Set RstDatoAdicional = BuscarCampo(Rst("idlib"), NulosN(Rst("tipo")), NulosN(Rst("idmod")), NulosN(Rst("iddocpro")), Rst("idmov"), NulosN(Rst("idorides")))
        
        If RstDatoAdicional.State <> 0 Then
            If RstDatoAdicional.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = Format(NulosC(RstDatoAdicional("fchope")), "dd/mm/yy")
                Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstDatoAdicional("glosa"))
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstDatoAdicional("numreg"))
                Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(RstDatoAdicional("abrev"))
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(RstDatoAdicional("fchdoc"), "dd/mm/yy")
                Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(RstDatoAdicional("numdoc"))
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(RstDatoAdicional("numruc"))
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(RstDatoAdicional("apenom"))
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(RstDatoAdicional("impven")), "0.000")
            End If
        End If
        Set RstDatoAdicional = Nothing
        
        Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosC(Rst.Fields("cuenta"))      '****
        Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosC(Rst.Fields("descripcion")) '****
        
        
        Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(Rst.Fields("impdebsol"), FORMAT_MONTO) '****
        Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(Rst.Fields("imphabsol"), FORMAT_MONTO) '****
        
        xAcumulado(0, 0) = xAcumulado(0, 0) + NulosN(Format(NulosN(Rst.Fields("impdebsol")), FORMAT_MONTO))
        xAcumulado(0, 1) = xAcumulado(0, 1) + NulosN(Format(NulosN(Rst.Fields("imphabsol")), FORMAT_MONTO))
       
        Rst.MoveNext
        
        If Rst.EOF = True Then
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = "Total ==>"
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 13) = Fg1.TextMatrix(Fg1.Rows - 1, 14), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 13) = Fg1.TextMatrix(Fg1.Rows - 1, 14), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 13) = Fg1.TextMatrix(Fg1.Rows - 1, 14), &H800000, &HFF)
            
            xAcumulado(1, 0) = xAcumulado(1, 0) + xAcumulado(0, 0)
            xAcumulado(1, 1) = xAcumulado(1, 1) + xAcumulado(0, 1)
            
            Exit Do
        End If
        
        'FechaRegistro
        If xAsiento <> Rst.Fields("numreg") & "" Then
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = "Total ==>"
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 13) = Fg1.TextMatrix(Fg1.Rows - 1, 14), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 13) = Fg1.TextMatrix(Fg1.Rows - 1, 14), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 13) = Fg1.TextMatrix(Fg1.Rows - 1, 14), &H800000, &HFF)
            
            xAcumulado(1, 0) = xAcumulado(1, 0) + xAcumulado(0, 0)
            xAcumulado(1, 1) = xAcumulado(1, 1) + xAcumulado(0, 1)
            
            Fg1.Rows = Fg1.Rows + 1

            xAsiento = Rst.Fields("numreg") & ""
            xAcumulado(0, 0) = 0:   xAcumulado(0, 1) = 0
        End If
                
    Loop
        
    Fg1.Rows = Fg1.Rows + 2
    Fg1.TextMatrix(Fg1.Rows - 1, 12) = "Total " + StrConv(N_Libro, 3) + " ==>"
    Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(xAcumulado(1, 0), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(xAcumulado(1, 1), FORMAT_MONTO)
    '--ACUMULAR LOS TOTALES POR LIBRO
    xAcumulado(2, 0) = xAcumulado(2, 0) + NulosN(Format(xAcumulado(1, 0), FORMAT_MONTO))
    xAcumulado(2, 1) = xAcumulado(2, 1) + NulosN(Format(xAcumulado(1, 1), FORMAT_MONTO))
    '-----------------------------------------------------------
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 12
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 13
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 14
    Exit Sub
SALIR:
    Fg1.Rows = Fg1.Rows + 1
    Set Rst = Nothing
End Sub

Sub ProcesarDiario2(IDLIBRO As Integer, N_Libro As String)
    Dim Rst As New ADODB.Recordset
    
    Dim N_SQL_WHERE As String
    Dim N_SQL As String
    Dim N_SQL_SALDO As String
    
    Frame5.Left = 3413
    Frame5.Top = 2685
    Frame5.Visible = True
    DoEvents
    N_SQL_SALDO = " "
    If NulosN(IDLIBRO) <> 0 Then
        N_SQL_WHERE = " AND ( con_diario.idlib = " + CStr(IDLIBRO) + " ) "
    End If
    
    
    If opt_fecha(0).Value = True Then '--por fecha
        If CDate(Me.TxtFchIni.Valor) = CDate("01/01/" + AnoTra) Then
            N_SQL_SALDO = " OR ( (con_diario.fchasi) IS NULL " + N_SQL_WHERE + " ) "
        Else
            N_SQL_SALDO = " AND (con_diario.fchasi) IS NOT NULL "
        End If
    Else '--por periodo

    End If
    
    N_SQL = "SELECT con_diario.idlib, IIf([con_diario].[idlib]<>3,[mae_libros].[descripcion],[mae_librossub].[descripcion]) AS nomlib, con_diario.idcue AS idcuenta, con_planctas.cuenta, con_planctas.descripcion, con_diario.idmov, " _
        + vbCr + " Format([con_diario]![idmes],'00') & IIf(mae_libros.codsun Is Null OR mae_libros.codsun ='','FF',Format([mae_libros].[codsun],'00')) & Trim([con_diario]![numasi]) AS numreg, " _
        + vbCr + " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,mae_documento.abrev,IIf(con_diario.idlib=2,mae_documento_1.abrev,IIf(con_diario.idlib=3,mae_documento_2.abrev,IIf(con_diario.idlib=4,mae_documento_3.abrev,IIf(con_diario.idlib=5,mae_documento_4.abrev,IIf(con_diario.idlib=6,mae_doccajaban.abrev,IIf(con_diario.idlib=8,'CAN',IIf(con_diario.idlib=9,'PLA',IIf(con_diario.idlib=37,'CAN',IIf(con_diario.idlib=38,mae_doccajaban_1.abrev,IIf(con_diario.idlib=39,mae_documento_5.abrev ,'OTROS LIBROS')))))))))))) AS tipdoc, " _
        + vbCr + " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.fchdoc,IIf(con_diario.idlib=2,vta_ventas.fchdoc,IIf(con_diario.idlib=3,con_proviciones.fchdoc,IIf(con_diario.idlib=4,con_percepcion.fchdoc,IIf(con_diario.idlib=5,con_retencion.fchemi,IIf(con_diario.idlib=6,con_cajabanco.fchope,IIf(con_diario.idlib=8,con_canjes.fchemi,IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,con_letra.fchemi,IIf(con_diario.idlib=38,con_ctasrendir.fchemi,IIf(con_diario.idlib=39,con_devoluciones.fchemi,'OTROS LIBROS')))))))))))) AS fchemi, " _
        + vbCr + " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.numser & '-' & com_compras.numdoc,IIf(con_diario.idlib=2,vta_ventas.numser & '-' & vta_ventas.numdoc,IIf(con_diario.idlib=3,con_proviciones.numser & '-' & con_proviciones.numdoc,IIf(con_diario.idlib=4,con_percepcion.numser & '-' & con_percepcion.numdoc,IIf(con_diario.idlib=5,con_retencion.numser & '-' & con_retencion.numdoc,IIf(con_diario.idlib=6,con_cajabanco.numdoc,IIf(con_diario.idlib=8,[con_canjes].[numser] & '-' & [con_canjes].[numdoc],IIf(con_diario.idlib=9,' ',IIf(con_diario.idlib=37,' ',IIf(con_diario.idlib=38,con_ctasrendir.numdoc,IIf(con_diario.idlib=39,con_devoluciones.numdoc,'OTROS LIBROS')))))))))))) AS numdoc, " _
        + vbCr + " con_tc.impven, "
    If Me.OptSoles = True Then
        N_SQL = N_SQL _
        + vbCr + " IIf(con_diario.impdebdol<>0,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
        + vbCr + " IIf(con_diario.imphabdol<>0,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol "
    Else
        N_SQL = N_SQL _
        + vbCr + " IIf(con_diario.impdebdol<>0,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
        + vbCr + " IIf(con_diario.imphabdol<>0, con_diario.imphabdol, IIf(con_tc.impven Is Null Or con_diario.imphabsol = 0, 0, (con_diario.imphabsol / con_tc.impven))) As imphaberdol "
    End If
    N_SQL = N_SQL _
        + vbCr + " FROM (mae_libros RIGHT JOIN ((((((con_planctas RIGHT JOIN ((((((((((((((con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN com_compras ON con_diario.idmov = com_compras.id) LEFT JOIN vta_ventas ON con_diario.idmov = vta_ventas.id) LEFT JOIN con_retencion ON con_diario.idmov = con_retencion.id) LEFT JOIN con_percepcion ON con_diario.idmov = con_percepcion.id) LEFT JOIN con_proviciones ON con_diario.idmov = con_proviciones.id) LEFT JOIN con_cajabanco ON con_diario.idmov = con_cajabanco.id) LEFT JOIN con_canjes ON con_diario.idmov = con_canjes.id) " _
        + vbCr + " LEFT JOIN mae_doccajaban ON con_cajabanco.iddoc = mae_doccajaban.id) LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_documento AS mae_documento_1 ON vta_ventas.tipdoc = mae_documento_1.id) LEFT JOIN mae_documento AS mae_documento_2 ON con_proviciones.tipdoc = mae_documento_2.id) LEFT JOIN mae_documento AS mae_documento_3 ON con_percepcion.tipdoc = mae_documento_3.id) LEFT JOIN mae_documento AS mae_documento_4 ON con_retencion.iddoc = mae_documento_4.id) ON con_planctas.id = con_diario.idcue) LEFT JOIN con_ctasrendir ON con_diario.idmov = con_ctasrendir.id) LEFT JOIN con_letra ON con_diario.idmov = con_letra.id) LEFT JOIN mae_doccajaban AS mae_doccajaban_1 ON con_ctasrendir.tipdoc = mae_doccajaban_1.id) " _
        + vbCr + " LEFT JOIN con_devoluciones ON con_diario.idmov = con_devoluciones.id) LEFT JOIN mae_documento AS mae_documento_5 ON con_devoluciones.iddoc = mae_documento_5.id) ON mae_libros.id = con_diario.idlib) LEFT JOIN mae_librossub ON (con_proviciones.idlib = mae_librossub.idlib) AND (con_proviciones.idsublib = mae_librossub.id) "
    '***************************************
    If opt_fecha(0).Value = True Then  '--por fecha
        N_SQL = N_SQL + vbCr + " WHERE ( (con_diario.fchasi >=CDate('" + TxtFchIni.Valor + "') And con_diario.fchasi<=CDate('" + TxtFchFin.Valor + "')) "
    Else '--por intervalo
        N_SQL = N_SQL + vbCr + " WHERE ( con_diario.idmes >=  " & mMesIni & " AND con_diario.idmes <= " & mMesFin & " "
    End If
    '***************************************
    N_SQL = N_SQL _
        + vbCr + N_SQL_WHERE + " ) " + N_SQL_SALDO _
        + vbCr + " ORDER BY con_diario.idlib, con_diario.idmes, con_diario.numasi, con_planctas.cuenta;"
    '***************************************
    N_SQL = Replace(N_SQL, "impdebesol", "debe")
    N_SQL = Replace(N_SQL, "imphabersol", "haber")
    N_SQL = Replace(N_SQL, "impdebedol", "debe")
    N_SQL = Replace(N_SQL, "imphaberdol", "haber")
    '***************************************
    RST_Busq Rst, N_SQL, xCon
    '--LIMPIAR ACUMULADO POR LIBRO
    xAcumulado(0, 0) = 0:   xAcumulado(0, 1) = 0
    xAcumulado(1, 0) = 0:   xAcumulado(1, 1) = 0
    '---------------------------------------------------------------------------
    If Rst.State = 0 Then GoTo SALIR
    If Rst.BOF = True Or Rst.EOF = True Or Rst.RecordCount = 0 Then GoTo SALIR
    Rst.MoveFirst
    Dim xAsiento
    Dim HabDol, DebDol  As Double
    
    xAsiento = NulosC(Rst.Fields("numreg"))
    
    Dim Cambiar As Boolean
    If Rst.RecordCount > 1 Then
        ProgressBar1.Max = Rst.RecordCount
    End If
    Do While Not Rst.EOF
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        ProgressBar1.Value = Rst.Bookmark
        '-----------------------------------------------
        Fg1.Rows = Fg1.Rows + 1
        
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst.Fields("numreg"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst.Fields("nomlib"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst.Fields("tipdoc"))
        If IsDate(Rst.Fields("fchemi")) = True Then
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(Rst.Fields("fchemi"), FORMAT_DATE)
        End If
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(Rst.Fields("numdoc"))
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(Rst.Fields("cuenta"))
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(Rst.Fields("descripcion"))
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = Rst.Fields("impven") & ""
        
        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Rst.Fields("debe"), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(Rst.Fields("haber"), FORMAT_MONTO)
        
        xAcumulado(0, 0) = xAcumulado(0, 0) + NulosN(Format(NulosN(Rst.Fields("debe")), FORMAT_MONTO))
        xAcumulado(0, 1) = xAcumulado(0, 1) + NulosN(Format(NulosN(Rst.Fields("haber")), FORMAT_MONTO))
       
        Rst.MoveNext
        
        If Rst.EOF = True Then
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = "Total ==>"
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 9) = Fg1.TextMatrix(Fg1.Rows - 1, 10), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 9) = Fg1.TextMatrix(Fg1.Rows - 1, 10), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 9) = Fg1.TextMatrix(Fg1.Rows - 1, 10), &H800000, &HFF)
            
            xAcumulado(1, 0) = xAcumulado(1, 0) + xAcumulado(0, 0)
            xAcumulado(1, 1) = xAcumulado(1, 1) + xAcumulado(0, 1)
            
            Exit Do
        End If
        
        'FechaRegistro
        If xAsiento <> Rst.Fields("numreg") & "" Then
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = "Total ==>"
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 9) = Fg1.TextMatrix(Fg1.Rows - 1, 10), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 9) = Fg1.TextMatrix(Fg1.Rows - 1, 10), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, IIf(Fg1.TextMatrix(Fg1.Rows - 1, 9) = Fg1.TextMatrix(Fg1.Rows - 1, 10), &H800000, &HFF)
            
            xAcumulado(1, 0) = xAcumulado(1, 0) + xAcumulado(0, 0)
            xAcumulado(1, 1) = xAcumulado(1, 1) + xAcumulado(0, 1)
            
            Fg1.Rows = Fg1.Rows + 1

            xAsiento = Rst.Fields("numreg") & ""
            xAcumulado(0, 0) = 0:   xAcumulado(0, 1) = 0
        End If
                
    Loop
        
    Fg1.Rows = Fg1.Rows + 2
    Fg1.TextMatrix(Fg1.Rows - 1, 7) = "Total " + StrConv(N_Libro, 3) + " ==>"
    Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(xAcumulado(1, 0), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(xAcumulado(1, 1), FORMAT_MONTO)
    '--ACUMULAR LOS TOTALES POR LIBRO
    xAcumulado(2, 0) = xAcumulado(2, 0) + NulosN(Format(xAcumulado(1, 0), FORMAT_MONTO))
    xAcumulado(2, 1) = xAcumulado(2, 1) + NulosN(Format(xAcumulado(1, 1), FORMAT_MONTO))
    '-----------------------------------------------------------
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 7
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 9
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 10
SALIR:
    Fg1.Rows = Fg1.Rows + 1
    Set Rst = Nothing
    
End Sub


