VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultaDiario 
   Caption         =   "Contabilidad - Diario "
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   12510
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
      Height          =   870
      Left            =   8160
      TabIndex        =   25
      Top             =   390
      Width           =   3705
      Begin VB.CommandButton CmdBusMon 
         Height          =   230
         Left            =   1155
         Picture         =   "FrmConsultaDiario.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   420
         Width           =   210
      End
      Begin VB.TextBox TxtIdMon 
         Height          =   300
         Left            =   690
         MaxLength       =   1
         TabIndex        =   27
         Text            =   "TxtIdMon"
         Top             =   390
         Width           =   705
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
         Left            =   1395
         TabIndex        =   29
         Top             =   390
         Width           =   2205
      End
      Begin VB.Label LblTipCam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
         Height          =   195
         Index           =   4
         Left            =   75
         TabIndex        =   28
         Top             =   495
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   3690
      TabIndex        =   6
      Top             =   390
      Width           =   4425
      Begin VB.CommandButton cmd_periodo 
         Height          =   255
         Index           =   0
         Left            =   1845
         Picture         =   "FrmConsultaDiario.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   960
         Width           =   285
      End
      Begin VB.CommandButton cmd_periodo 
         Height          =   255
         Index           =   1
         Left            =   4035
         Picture         =   "FrmConsultaDiario.frx":04B4
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   960
         Width           =   285
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   540
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
         Height          =   230
         Left            =   4110
         Picture         =   "FrmConsultaDiario.frx":0836
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   210
         Width           =   210
      End
      Begin VB.TextBox TxtLibro 
         Height          =   300
         Left            =   540
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "TxtLibro"
         Top             =   165
         Width           =   3825
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   2745
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
         Left            =   540
         TabIndex        =   24
         Top             =   930
         Width           =   1620
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
         Left            =   2745
         TabIndex        =   23
         Top             =   930
         Width           =   1620
      End
      Begin VB.Label LblIdLibro 
         AutoSize        =   -1  'True
         Caption         =   "LblIdLibro"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4020
         TabIndex        =   20
         Top             =   540
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Index           =   1
         Left            =   2445
         TabIndex        =   10
         Top             =   570
         Width           =   135
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Libro"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   345
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
      Left            =   2040
      TabIndex        =   17
      Top             =   390
      Width           =   1620
      Begin VB.OptionButton OptLibro 
         Caption         =   "Por Libro"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   1110
      End
      Begin VB.OptionButton OptTodo 
         Caption         =   "Todos los Libros"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   555
         Width           =   1455
      End
      Begin VB.Label LblIdMes 
         AutoSize        =   -1  'True
         Caption         =   "LblIdMes"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1725
         TabIndex        =   19
         Top             =   525
         Visible         =   0   'False
         Width           =   645
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
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":0968
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":0EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":123E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":13C2
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":1816
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":192E
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":1E72
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":23B6
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":24CA
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":25DE
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":2A32
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":2B9E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":30E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":3400
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":3792
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":3B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDiario.frx":3E3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
            Object.ToolTipText     =   "Buscar Asiento"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Configurar Formatos"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Left            =   2850
      TabIndex        =   11
      Top             =   7890
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   105
         TabIndex        =   12
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
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   90
         Width           =   1530
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
         TabIndex        =   13
         Top             =   90
         Width           =   1530
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
      TabIndex        =   15
      Top             =   390
      Width           =   1935
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Periodo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   16
         Top             =   555
         Width           =   1125
      End
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Fecha"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   3
         Top             =   270
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6300
      Left            =   60
      TabIndex        =   30
      Top             =   1290
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
         Height          =   5880
         Left            =   45
         TabIndex        =   31
         Top             =   45
         Width           =   11880
         _cx             =   20955
         _cy             =   10372
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
         FormatString    =   $"FrmConsultaDiario.frx":4158
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
         Height          =   5880
         Left            =   12615
         TabIndex        =   32
         Top             =   45
         Width           =   11880
         _cx             =   20955
         _cy             =   10372
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
         FormatString    =   $"FrmConsultaDiario.frx":4327
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
Attribute VB_Name = "FrmConsultaDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Modificado : 06/11/08 - Johan Castro
'             Corregir el diario de las operaciones en dolares, no muestran en soles
'             Corregir el diario valorizados en dolares
'Modificado : 26/01/10 - Johan Castro
'             mostrar la fecha inicio y final empiece el 01/01 de año de trabajo



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

Dim mColDebe As Integer '--posicion de la columna debe
Dim mColHaber As Integer '--posicion de la columna haber
Dim mPosRegistro As Integer '--indica la posicion del numero de registro


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
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
   
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_libros  where activo = -1 ORDER BY descripcion ", xCampos(), "Buscando Libro Contable", "descripcion", "descripcion", Principio
    If xRs.State = 1 Then
        TxtLibro.Text = NulosC(xRs("descripcion"))
        LblIdLibro.Caption = NulosC(xRs("id"))
        If TxtFchIni.Visible = True Then
            TxtFchIni.SetFocus
        Else
            cmd_periodo(0).SetFocus
        End If
    End If
    Set xRs = Nothing
End Sub


Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub

    If Me.Height > 3000 Then
        TabOne1.Top = 1300
        TabOne1.Width = Me.Width - 150
        TabOne1.Height = Me.Height - 1700
    End If
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

Private Sub pExportar()
    Dim xFun As New SGI2_funciones.formularios
    Dim Rst As New ADODB.Recordset
    
    If TabOne1.CurrTab = 0 Then
        If Fg1.Rows = 1 Then
            MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        'ExportarEcelDiario
'''        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "LIBRO DIARIO", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en " & LblMoneda.Caption, "Diario - Detalle"        ', Rst, ""
        GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "LIBRO DIARIO", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en " & LblMoneda.Caption
        
        Set xFun = Nothing
    End If
    If TabOne1.CurrTab = 1 Then
        If fg2.Rows = 2 Then
            MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
'        ExportarExcelResumen
        xFun.VSFlexGrid_Exportar_MSExcel xCon, fg2, "RESUMEN DEL DIARIO", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en " & LblMoneda.Caption, "Diario - Resumen"   ', Rst, ""
        Set xFun = Nothing
    End If
End Sub

Private Sub pImprimirDet()
    TabOne1.CurrTab = 0
    If Me.TabOne1.CurrTab = 0 Then
        If Fg1.Rows = 1 Then
            MsgBox "No hay registros para imprimir", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
    End If
    
    Dim nPeriodo   As String
    Dim xMoneda As String
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
    
    xMoneda = LblMoneda.Caption
    
    Dim RstTmp As New ADODB.Recordset
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 1) And ((con_formatostipodet.idformatotipo) = 2) " _
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
    xfrm.Titulo1 = "LIBRO DIARIO " & "(Expresado en " & xMoneda & ")"
    xfrm.Titulo2 = nPeriodo
    xfrm.TamañoFuente = 6
    xfrm.TamañoCabecera = 8
    xfrm.FuenteCabecera = "Courier New"
    xfrm.Posicion_Hoja = Vertical
    xfrm.Tamaño_Hoja = A_4
    xfrm.TextoConsiderar = "LIBRO"
    xfrm.TextoConsiderarAncho = 5
    xfrm.ImprimirArray xCampos, Rst
    Set xfrm = Nothing
    
    
    
    
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
    Me.TabOne1.CurrTab = 1
    
    Frame5.Visible = True
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Fg1.Rows = 2
        fg2.Rows = 1
        DoEvents
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
''''                ProcesarDiario Rst("id"), "Libro " + UCase(Trim(Rst("descripcion"))), 1
                ProcesarDiario4 Rst("id"), "Libro " + Trim(Rst("descripcion") & "")
            Else
            
'                ProcesarDiario2 NulosN(LblIdLibro.Caption), "Libro " + Trim(Rst("descripcion") & "")
'                ProcesarDiario NulosN(LblIdLibro.Caption), "Libro " + Trim(Rst("descripcion") & ""), 1
'                ProcesarDiario3 NulosN(LblIdLibro.Caption), "Libro " + Trim(Rst("descripcion") & "")
                 ProcesarDiario4 NulosN(LblIdLibro.Caption), "Libro " + Trim(Rst("descripcion") & "")
            End If
            
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
            
        Next A
    End If
    
    '---------
    If OptLibro.Value = False And (xAcumulado(2, 0) <> 0 Or xAcumulado(2, 1) <> 0) Then
        Fg1.Rows = Fg1.Rows + 2
        Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total Gen.==>"
        Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xAcumulado(2, 0), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xAcumulado(2, 1), FORMAT_MONTO)

        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1
        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe
        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber
        
        Fg1.Rows = Fg1.Rows + 1
    End If
    
    
    Erase xAcumulado()
    Me.TabOne1.CurrTab = 0
    '--SI SE NTERRUMPE EL PROCESO => SALIR
    If BAND_INTERRUMPIR = True Then GoTo SALIR:
    '-----------------------------------------------
    Label3.Caption = "Procesando Resumen "
'    ProcesarResumen NulosN(LblIdLibro.Caption)
    ProcesarResumen3 NulosN(LblIdLibro.Caption)
'''    ProcesarResumenLibros
    Frame5.Visible = False
    
    '------------------------
    '--ajustando las columnas de acuerdo a los importes
    Fg1.AutoSizeMode = flexAutoSizeColWidth
    Fg1.AutoSize mColDebe
    Fg1.AutoSize mColHaber

    '--ajustando las columnas de acuerdo a los importes
    fg2.AutoSizeMode = flexAutoSizeColWidth
    fg2.AutoSize 3
    fg2.AutoSize 4
    '------------------------
    
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
        fg2.Rows = fg2.Rows + 1
        fg2.TextMatrix(fg2.Rows - 1, 1) = RstRes("cuenta") & ""
        fg2.TextMatrix(fg2.Rows - 1, 2) = RstRes("descripcion") & ""

        fg2.TextMatrix(fg2.Rows - 1, 3) = Format(NulosN(RstRes.Fields("debe")), FORMAT_MONTO)
        fg2.TextMatrix(fg2.Rows - 1, 4) = Format(NulosN(RstRes.Fields("haber")), FORMAT_MONTO)
        
        xAcumulado(0, 0) = xAcumulado(0, 0) + NulosN(Format(NulosN(RstRes.Fields("debe")), FORMAT_MONTO))
        xAcumulado(0, 1) = xAcumulado(0, 1) + NulosN(Format(NulosN(RstRes.Fields("haber")), FORMAT_MONTO))
        
        If NulosN(RstRes.Fields("debe")) = 0 And NulosN(RstRes.Fields("haber")) = 0 Then
            fg2.RemoveItem fg2.Rows - 1
        End If
        RstRes.MoveNext
        
        If RstRes.EOF = True Then
            fg2.Rows = fg2.Rows + 2
            fg2.TextMatrix(fg2.Rows - 1, 2) = "TOTAL ==>"
            fg2.TextMatrix(fg2.Rows - 1, 3) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            fg2.TextMatrix(fg2.Rows - 1, 4) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            FORMATO_CELDA fg2, fg2.Rows - 1, 2
            FORMATO_CELDA fg2, fg2.Rows - 1, 3
            FORMATO_CELDA fg2, fg2.Rows - 1, 4
            
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
    If NulosN(TxtIdMon.Text) = 1 Then
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
        fg2.Rows = fg2.Rows + 1
        fg2.TextMatrix(fg2.Rows - 1, 1) = RstRes("cuenta") & ""
        fg2.TextMatrix(fg2.Rows - 1, 2) = RstRes("descripcion") & ""

        fg2.TextMatrix(fg2.Rows - 1, 3) = Format(NulosN(RstRes.Fields("debe")), FORMAT_MONTO)
        fg2.TextMatrix(fg2.Rows - 1, 4) = Format(NulosN(RstRes.Fields("haber")), FORMAT_MONTO)
        
        xAcumulado(0, 0) = xAcumulado(0, 0) + NulosN(Format(NulosN(RstRes.Fields("debe")), FORMAT_MONTO))
        xAcumulado(0, 1) = xAcumulado(0, 1) + NulosN(Format(NulosN(RstRes.Fields("haber")), FORMAT_MONTO))
        
        If NulosN(RstRes.Fields("debe")) = 0 And NulosN(RstRes.Fields("haber")) = 0 Then
            fg2.RemoveItem fg2.Rows - 1
        End If
        RstRes.MoveNext
        
        If RstRes.EOF = True Then
            fg2.Rows = fg2.Rows + 2
            fg2.TextMatrix(fg2.Rows - 1, 2) = "TOTAL ==>"
            fg2.TextMatrix(fg2.Rows - 1, 3) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            fg2.TextMatrix(fg2.Rows - 1, 4) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            FORMATO_CELDA fg2, fg2.Rows - 1, 2
            FORMATO_CELDA fg2, fg2.Rows - 1, 3
            FORMATO_CELDA fg2, fg2.Rows - 1, 4
            
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
        
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
                
        OptLibro.Value = True
        LimpiaText lbl_periodo
        lbl_periodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
        lbl_periodo(1).Caption = lbl_periodo(0).Caption
        mMesIni = xMes
        mMesFin = xMes
        TabOne1.CurrTab = 0
        
        TxtFchIni.Valor = ""
        TxtFchFin.Valor = ""
    
        TxtLibro.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    ElseIf KeyCode = vbKeyF3 And Shift = 0 Then
        BuscarVSFlexGrid
    ElseIf KeyCode = vbKeyF8 Then
        pConsultar
    End If

End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Blanquea
    TxtFchIni.Valor = CDate("01/01/" & AnoTra)
    TxtFchFin.Valor = CDate("01/01/" & AnoTra)
    
    Fg1.SelectionMode = flexSelectionByRow
    fg2.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDNone
    fg2.Editable = flexEDNone
    
    SetearCuadricula Fg1, 1, xCon, 1, 0, False
    
    '--buscar los registros
    Fg1.AutoSearch = flexSearchFromTop
    fg2.AutoSearch = flexSearchFromTop
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
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
        Ocultar cmd_periodo, False
        Ocultar lbl_periodo, False
        
    Else '--por periodo
          
        TxtFchFin.Visible = False
        TxtFchIni.Visible = False
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
    If Button.Index = 5 Then pBuscarAsiento
    
    If Button.Index = 7 Then
        'pExportar
        ExportarEcelDiario
    End If
    If Button.Index = 8 Then
        If TabOne1.CurrTab = 0 Then
            pImprimirDet
        Else
            pImprimirRes
        End If
    End If
    If Button.Index = 9 Then Configurar
    If Button.Index = 11 Then
        Unload Me
    End If
End Sub

Sub Configurar()
    Dim xform As New SGI2_funciones.Varias
    If xform.CambioOpcionLiro(1, xCon, 1) = True Then
        SetearCuadricula Fg1, 1, xCon, 1, 0, False
        pConsultar
    End If
    Set xform = Nothing
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
        If NulosN(TxtIdMon.Text) = 1 Then
            nTitulo1 = "(Expresado en Nuevos Soles)"
        Else
            nTitulo1 = "(Expresado en Dolares Americanos)"
        End If
        .Cells(4, 2) = "Libro Diario"
        .Cells(5, 2) = nPeriodo
        .Cells(6, 2) = nTitulo1
        '**********************************
        
'        If Fg1.ColWidth(1) <> 0 Then .Columns(2).ColumnWidth = Fg1.ColWidth(1) / 100
'        If Fg1.ColWidth(2) <> 0 Then .Columns(3).ColumnWidth = Fg1.ColWidth(2) / 100
'        If Fg1.ColWidth(3) <> 0 Then .Columns(4).ColumnWidth = Fg1.ColWidth(3) / 100
'        If Fg1.ColWidth(4) <> 0 Then .Columns(5).ColumnWidth = Fg1.ColWidth(4) / 100
'        If Fg1.ColWidth(5) <> 0 Then .Columns(6).ColumnWidth = Fg1.ColWidth(5) / 100
'        If Fg1.ColWidth(6) <> 0 Then .Columns(7).ColumnWidth = Fg1.ColWidth(6) / 100
'        If Fg1.ColWidth(7) <> 0 Then .Columns(8).ColumnWidth = Fg1.ColWidth(7) / 100
'        If Fg1.ColWidth(8) <> 0 Then .Columns(9).ColumnWidth = Fg1.ColWidth(8) / 100
'        If Fg1.ColWidth(9) <> 0 Then .Columns(10).ColumnWidth = Fg1.ColWidth(9) / 100
'        If Fg1.ColWidth(10) <> 0 Then .Columns(11).ColumnWidth = Fg1.ColWidth(10) / 100
        
        xFilas = 8
        For B = 1 To Fg1.Cols - 1
            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(0, B)
        Next B
        
        xFilas = xFilas + 1
        For A = 1 To Fg1.Rows - 1
            DoEvents
            For B = 1 To Fg1.Cols - 1
                If B <= 12 Then
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
    If NulosN(TxtIdMon.Text) = 1 Then
        nTitulo1 = "(Expresado en Nuevos Soles)"
    Else
        nTitulo1 = "(Expresado en Dolares Americanos)"
    End If
    
    
    xExport.VSFlexGrid_Exportar_MSExcel xCon, fg2, "RESUMEN DEL DIARIO", nPeriodo, nTitulo1, "Resumen del Diario"
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
        
        If Tipo = 1 Then '--origen
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
        
        If Tipo = 2 Then '--destino
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
                        & " LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha WHERE (((con_diario.iddocpro)=" & IdDocPro & ") AND ((con_diario.idmov)=" & IdMovimiento & ") AND ((con_tc.idmon)=2)) ", xCon
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
            & " And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))) ORDER BY Format(con_diario!idmes,'00')+mae_libros!codsun+con_diario!numasi,con_planctas.cuenta "
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
        
    If NulosN(TxtIdMon.Text) = 1 Then
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



Private Sub ProcesarDiario3(IDLIBRO As Integer, N_Libro As String)
    '===================================================================================================
    'creado:     20/11/08
    'Propósito:  Mostrar la información del diario
    '
    'Entradas:   IDLIBRO = Código de Libro
    '            N_Libro = Desripción de Libro
    '            Tipo = ?
    '
    'Resultados: Informacion de los diversos libros en pantalla
    '===================================================================================================
    Dim Rst As New ADODB.Recordset
    
    Dim nSQL As String
    Dim nSQLSaldo As String
    Dim nSQLWhere As String
    Dim nSQLCampos As String
    
    Frame5.Left = 3413
    Frame5.Top = 2685
    Frame5.Visible = True
    ProgressBar1.Value = 1
    
    '---
    DoEvents
    
    '----------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------
    nSQL = "SELECT f_det.abrev, f_det.nomcampo " _
     & " FROM con_formatostipo AS f LEFT JOIN con_formatostipodet AS f_det ON (f.idformato = f_det.idformato) AND (f.id = f_det.idformatotipo) " _
     & " WHERE (((f.idformato)=1) AND ((f.defecto)=-1) AND ((f_det.mostrar)=-1)) " _
     & " ORDER BY f_det.orden; "
    
    RST_Busq Rst, nSQL, xCon
    
    If Rst.RecordCount = 0 Then
        MsgBox "El formato que desea Consultar no tiene columnas para visualizar" & vbCr & "Para continuar seleccione las columnas que desea visualizar", vbExclamation, xTitulo
        BAND_INTERRUMPIR = True
        Exit Sub
    End If
    
    Rst.MoveFirst
    nSQLCampos = ""
    Do While Not Rst.EOF
        nSQLCampos = nSQLCampos & Rst("nomcampo") & ","
        Rst.MoveNext
    Loop
    nSQLCampos = Mid(nSQLCampos, 1, Len(nSQLCampos) - 1)
    
    Set Rst = Nothing
    
    '----------------------------------------------------------------------------------
    '--generando el filtro
        
    If opt_fecha(0).Value = True Then  '--por fecha
        nSQLWhere = " WHERE ( (con_diario.idlib = " & IDLIBRO & " AND con_diario.fchasi BETWEEN CDate('" & TxtFchIni.Valor & "') And CDate('" + TxtFchFin.Valor + "')) "
    Else '--por intervalo
        nSQLWhere = " WHERE ( (con_diario.idlib = " & IDLIBRO & " AND con_diario.idmes >= " & mMesIni & " AND con_diario.idmes <= " & mMesFin & ") "
    End If
    
    '--para los saldos
    If opt_fecha(0).Value = True Then '--por fecha
        If CDate(Me.TxtFchIni.Valor) = CDate("01/01/" + AnoTra) Then
            nSQLWhere = nSQLWhere & "OR (con_diario.idlib = " & IDLIBRO & " AND con_diario.fchasi IS NULL) )"
        Else
            nSQLWhere = nSQLWhere & " )"
        End If
    Else '--por periodo
        If mMesIni = 1 Then
            nSQLWhere = nSQLWhere & "OR (con_diario.idlib = " & IDLIBRO & " AND con_diario.fchasi IS NULL) )"
        Else
            nSQLWhere = nSQLWhere & " )"
        End If
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''SI HAY DOS MONEDAS EN UN MISMO DIA EN CON_TC HABILITAR LA SIGUIENTE LINEA DE CODIGO
''''''    nSQLWhere = nSQLWhere & " AND (con_tc.idmon=2 OR con_tc.idmon IS NULL ) "
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    
    '----------------------------------------------------------------------------------
    '--generando la consulta
    
    Select Case IDLIBRO
    
    Case 1, 40
         '1: --compras
         '40:--honoharios
         
        nSQL = "SELECT con_diario.idmov, com_compras.idpro, con_diario.idcue, com_compras.idmon,mae_libros.descripcion as libdesc, com_compras.tipdoc, mae_prov.numruc, mae_prov.nombre AS apenom, mae_documento.abrev AS tdocdesc, mae_documento.codsun AS tdocsun, com_compras.fchdoc, con_tc.impven AS tc, IIf([com_compras]![numser] Is Null Or [com_compras]![numser]='','',[com_compras]![numser] & '-') & [com_compras]![numdoc] AS numdoc, Mid([com_compras]![numreg],1,2)+[mae_libros]![codsun]+Mid([com_compras]![numreg],3,4) AS registro, Mid([com_compras]![numreg],1,2)+[mae_libros]![codsun]+Mid([com_compras]![numreg],3,4) AS registroref, com_compras.fchdoc AS fchope, com_compras.glosa, mae_libros.codsun as libsun, CDbl([con_diario].[numasi]) AS corr, IIf([com_compras]![numser] Is Null Or [com_compras]![numser]='','',[com_compras]![numser] & '-') & [com_compras]![numdoc] AS docsustenta,  " _
            + vbCr + " con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2, con_diario.imphabdol, IIf(con_tc.impven Is Null Or con_diario.imphabsol = 0, 0, (con_diario.imphabsol / con_tc.impven))) As imphaberdol, " _
            + vbCr + " '' as RefTDoc, '' as RefNumDoc " _
            + vbCr + " FROM (mae_prov RIGHT JOIN ((((mae_libros RIGHT JOIN com_compras ON mae_libros.id = com_compras.idlib) LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN (con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) ON com_compras.id = con_diario.idmov) ON mae_prov.id = com_compras.idpro) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id "
        If IDLIBRO = 40 Then
            nSQL = Replace(nSQL, "com_compras", "com_honorarios")
        End If
        
    Case 2 '--ventas
        nSQL = "SELECT con_diario.idmov, vta_ventas.idcli, con_diario.idcue, con_diario.idmon,mae_libros.descripcion as libdesc, vta_ventas.tipdoc, mae_cliente.numruc, mae_cliente.nombre AS apenom, mae_documento.abrev AS tdocdesc, mae_documento.codsun AS tdocsun, vta_ventas.fchdoc, con_tc.impven AS tc, IIf([vta_ventas]![numser] Is Null Or [vta_ventas]![numser]='','',[vta_ventas]![numser] & '-') & [vta_ventas]![numdoc] AS numdoc, Mid([vta_ventas]![numreg],1,2)+[mae_libros]![codsun]+Mid([vta_ventas]![numreg],3,4) AS registro, Mid([vta_ventas]![numreg],1,2)+[mae_libros]![codsun]+Mid([vta_ventas]![numreg],3,4) AS registroref, vta_ventas.fchdoc AS fchope, vta_ventas.glosa, mae_libros.codsun as libsun, CDbl([con_diario].[numasi]) AS corr, IIf([vta_ventas]![numser] Is Null Or [vta_ventas]![numser]='','',[vta_ventas]![numser] & '-') & [vta_ventas]![numdoc] AS docsustenta,  " _
            + vbCr + " con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2, con_diario.imphabdol, IIf(con_tc.impven Is Null Or con_diario.imphabsol = 0, 0, (con_diario.imphabsol / con_tc.impven))) As imphaberdol, " _
            + vbCr + " '' as RefTDoc, '' as RefNumDoc " _
            + vbCr + " FROM ((((mae_libros RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_libros.id = vta_ventas.idlib) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN (con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) ON vta_ventas.id = con_diario.idmov) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id "
            
    Case 3 '--diario
        nSQL = "SELECT con_diario.idmov, '' AS idpro, con_diario.idcue, con_proviciones.idmon, mae_libros.descripcion AS libdesc, con_proviciones.tipdoc, '' AS numruc, '' AS apenom, mae_documento.abrev AS tdocdesc, mae_documento.codsun AS tdocsun, con_proviciones.fchdoc, con_tc.impven AS tc, IIf([con_proviciones]![numser] Is Null Or [con_proviciones]![numser]='','',[con_proviciones]![numser] & '-') & [con_proviciones]![numdoc] AS numdoc, Mid([con_proviciones]![numreg],1,2)+[mae_libros]![codsun]+Mid([con_proviciones]![numreg],3,4) AS registro, Mid([con_proviciones]![numreg],1,2)+[mae_libros]![codsun]+Mid([con_proviciones]![numreg],3,4) AS registroref, con_proviciones.fchdoc AS fchope, con_proviciones.glosa, mae_libros.codsun AS libsun, CDbl([con_diario].[numasi]) AS corr, IIf([con_proviciones]![numser] Is Null Or [con_proviciones]![numser]='','',[con_proviciones]![numser] & '-') & [con_proviciones]![numdoc] AS docsustenta, " _
            + vbCr + " con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc " _
            + vbCr + " FROM (mae_libros RIGHT JOIN ((((con_proviciones LEFT JOIN mae_librossub ON con_proviciones.idsublib = mae_librossub.id) LEFT JOIN (con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) ON con_proviciones.id = con_diario.idmov) LEFT JOIN mae_documento ON con_proviciones.tipdoc = mae_documento.id) LEFT JOIN con_tc ON con_proviciones.fchdoc = con_tc.fecha) ON mae_libros.id = con_proviciones.idlib) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id "

    Case 4, 5
        '--4:Igv Percepciones
        '--5:Igv Retenciones

        nSQLWhere = Replace(nSQLWhere, "WHERE", " AND ")
    
            '--retencion venta
        nSQL = "SELECT DISTINCT con_diario.idmov, IIf(con_diario.iddocpro=0,mae_cliente.id,mae_cliente_1.id) AS idclipro, con_diario.idcue, " _
            + vbCr + " IIf(con_diario.iddocpro=0,con_diario.idmon,mae_moneda_2.id) AS idmon, mae_libros.descripcion AS libdesc, " _
            + vbCr + " IIf(con_diario.iddocpro=0,mae_documento.id,mae_documento_2.id) AS tipdoc, IIf(con_diario.iddocpro=0,mae_cliente.numruc,mae_cliente_1.numruc) AS numruc, IIf(con_diario.iddocpro=0,mae_cliente.nombre,mae_cliente_1.nombre) AS apenom, IIf(con_diario.iddocpro=0,mae_documento.abrev,mae_documento_2.abrev) AS tdocdesc, IIf(con_diario.iddocpro=0,mae_documento.codsun,mae_documento_2.codsun) AS tdocsun, IIf(con_diario.iddocpro=0,con_retencion.fchemi,vta_ventas.fchdoc) AS fchdoc, con_tc.impven AS tc, " _
            + vbCr + " IIf(con_diario.iddocpro=0,IIf(con_retencion!numser Is Null Or con_retencion!numser='','',con_retencion!numser & '-') & con_retencion!numdoc,IIf(vta_ventas!numser Is Null Or vta_ventas!numser='','',vta_ventas!numser & '-') & vta_ventas!numdoc) AS numdoc, " _
            + vbCr + " Left(con_retencion.numreg,2) & Format(mae_libros.codsun,'00') & Right(con_retencion.numreg,4) AS registro, " _
            + vbCr + " IIf(con_diario.iddocpro=0,Left(con_retencion.numreg,2) & Format(mae_libros.codsun,'00') & Right(con_retencion.numreg,4),Left(vta_ventas.numreg,2) & Format(mae_libros_2.codsun,'00') & Right(vta_ventas.numreg,4)) AS registroref, " _
            + vbCr + " con_retencion.fchemi AS fchope, con_retencion.glosa, IIf(con_diario.iddocpro=0,mae_libros.codsun,mae_libros_2.codsun) AS libsun, IIf(con_diario.iddocpro=0,CDbl(con_diario.numasi),IIf(vta_ventas.numreg Is Null,0,CDbl(Right(vta_ventas.numreg,4)))) AS corr, IIf(con_diario.iddocpro=0,IIf(con_retencion!numser Is Null Or con_retencion!numser='','',con_retencion!numser & '-') & con_retencion!numdoc,IIf(vta_ventas!numser Is Null Or vta_ventas!numser='','',vta_ventas!numser & '-') & vta_ventas!numdoc) AS docsustenta, " _
            + vbCr + " con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
            + vbCr + " IIf(con_diario.iddocpro=0,mae_moneda.simbolo,mae_moneda_2.simbolo) AS simbolo, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc " _
            + vbCr + " FROM (((((mae_documento RIGHT JOIN (((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN vta_ventas ON con_diario.iddocpro = vta_ventas.id) RIGHT JOIN (con_retencion LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) ON con_diario.idmov = con_retencion.id) ON mae_documento.id = con_retencion.iddoc) LEFT JOIN mae_moneda AS mae_moneda_2 ON vta_ventas.idmon = mae_moneda_2.id) LEFT JOIN mae_documento AS mae_documento_2 ON vta_ventas.tipdoc = mae_documento_2.id) LEFT JOIN mae_libros AS mae_libros_2 ON vta_ventas.idlib = mae_libros_2.id) LEFT JOIN mae_cliente ON con_retencion.idpro = mae_cliente.id) LEFT JOIN mae_cliente AS mae_cliente_1 ON vta_ventas.idcli = mae_cliente_1.id " _
            + vbCr + " WHERE (((con_retencion.tipo)=2)) "
            
        nSQL = nSQL + nSQLWhere
        
        '--retencion compra
            
        nSQL = nSQL & vbCr & " Union" _
            + vbCr + " SELECT DISTINCT con_diario.idmov, IIf(con_diario.iddocpro=0,mae_prov.id,mae_prov_1.id) AS idclipro, con_diario.idcue, " _
            + vbCr + " IIf(con_diario.iddocpro=0,con_diario.idmon,mae_moneda_1.id) AS idmon, mae_libros.descripcion AS libdesc, " _
            + vbCr + " IIf(con_diario.iddocpro=0,mae_documento.id,mae_documento_1.id) AS tipdoc, IIf(con_diario.iddocpro=0,mae_prov.numruc,mae_prov_1.numruc) AS numruc, IIf(con_diario.iddocpro=0,mae_prov.nombre,mae_prov_1.nombre) AS apenom, IIf(con_diario.iddocpro=0,mae_documento.abrev,mae_documento_1.abrev) AS tdocdesc, IIf(con_diario.iddocpro=0,mae_documento.codsun,mae_documento_1.codsun) AS tdocsun, IIf(con_diario.iddocpro=0,con_retencion.fchemi,com_compras.fchdoc) AS fchdoc, con_tc.impven AS tc, " _
            + vbCr + " IIf(con_diario.iddocpro=0,IIf(con_retencion!numser Is Null Or con_retencion!numser='','',con_retencion!numser & '-') & con_retencion!numdoc,IIf(com_compras!numser Is Null Or com_compras!numser='','',com_compras!numser & '-') & com_compras!numdoc) AS numdoc, " _
            + vbCr + " Left(con_retencion.numreg,2) & Format(mae_libros.codsun,'00') & Right(con_retencion.numreg,4) AS registro, " _
            + vbCr + " IIf(con_diario.iddocpro=0,Left(con_retencion.numreg,2) & Format(mae_libros.codsun,'00') & Right(con_retencion.numreg,4),Left(com_compras.numreg,2) & Format(mae_libros_1.codsun,'00') & Right(com_compras.numreg,4)) AS registroref, con_retencion.fchemi AS fchope, con_retencion.glosa, IIf(con_diario.iddocpro=0,mae_libros.codsun,mae_libros_1.codsun) AS libsun, IIf(con_diario.iddocpro=0,CDbl(con_diario.numasi),IIf(com_compras.numreg Is Null,0,CDbl(Right(com_compras.numreg,4)))) AS corr, IIf(con_diario.iddocpro=0,IIf(con_retencion!numser Is Null Or con_retencion!numser='','',con_retencion!numser & '-') & con_retencion!numdoc,IIf(com_compras!numser Is Null Or com_compras!numser='','',com_compras!numser & '-') & com_compras!numdoc) AS docsustenta, " _
            + vbCr + " con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, IIf(con_diario.iddocpro=0,mae_moneda.simbolo,mae_moneda_1.simbolo) AS simbolo, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc " _
            + vbCr + " FROM (((((mae_documento RIGHT JOIN (((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN com_compras ON con_diario.iddocpro = com_compras.id) RIGHT JOIN (con_retencion LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) ON con_diario.idmov = con_retencion.id) ON mae_documento.id = con_retencion.iddoc) LEFT JOIN mae_moneda AS mae_moneda_1 ON com_compras.idmon = mae_moneda_1.id) LEFT JOIN mae_documento AS mae_documento_1 ON com_compras.tipdoc = mae_documento_1.id) LEFT JOIN mae_libros AS mae_libros_1 ON com_compras.idlib = mae_libros_1.id) LEFT JOIN mae_prov ON con_retencion.idpro = mae_prov.id) LEFT JOIN mae_prov AS mae_prov_1 ON com_compras.idpro = mae_prov_1.id " _
            + vbCr + " WHERE (((con_retencion.tipo)=1)) "
            
        nSQL = nSQL + nSQLWhere
    
        If IDLIBRO = 4 Then
            nSQL = Replace(nSQL, "con_retencion", "con_percepcion")
            nSQL = Replace(nSQL, "con_percepcion.idpro", "con_percepcion.idcli")
            nSQL = Replace(nSQL, "con_percepcion.fchemi", "con_percepcion.fchdoc")
            nSQL = Replace(nSQL, "con_percepcion.iddoc", "con_percepcion.tipdoc")
        End If
    
        
    Case 6 '--caja y bancos
'        Exit Sub

    
        nSQLWhere = Replace(nSQLWhere, "WHERE", " AND ")
        nSQLWhere = Replace(nSQLWhere, "con_diario", "diario")
            
    '--ingresos y egresos(cuando no es detallado) -- origen (haber)
    nSQL = "SELECT tes_caja.id AS idmov, con_diario.idcue, mae_tipomov.descripcion AS Tipo1, tes_caja.idmon, mae_libros.descripcion AS libdesc, " _
            + vbCr + " '' AS tipdoc, '' AS numruc, '' AS apenom, tes_documentos.abrev AS tdocdesc, '' AS tdocsun, tes_caja.fchope AS fchdoc, con_tc.impven AS tc, " _
            + vbCr + " IIf([tes_cajaorigendet].[numser]='' Or [tes_cajaorigendet].[numser] Is Null,[tes_cajaorigendet].[numdoc],[tes_cajaorigendet].[numser] & '-' & [tes_cajaorigendet].[numdoc]) AS numdoc, " _
            + vbCr + " Mid([tes_caja].[numreg],1,2)+[mae_libros].[codsun]+Mid([tes_caja].[numreg],3,4) AS registro, " _
            + vbCr + " Mid([tes_caja].[numreg],1,2)+[mae_libros].[codsun]+Mid([tes_caja].[numreg],3,4) AS registroref, " _
            + vbCr + " tes_caja.fchope, tes_caja.glosa, mae_libros.codsun AS libsun, CDbl([con_diario].[numasi]) AS corr, IIf(Trim([tes_cajaorigendet].[numser])='' Or [tes_cajaorigendet].[numser] Is Null,'',[tes_cajaorigendet].[numser] & '-' & [tes_cajaorigendet].[numdoc]) AS docsustenta, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc " _
            + vbCr + " FROM (mae_libros RIGHT JOIN (((tes_caja LEFT JOIN ((con_diario LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) ON tes_caja.id = con_diario.idmov) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN mae_tipomov ON tes_caja.tipmov = mae_tipomov.id) ON mae_libros.id = con_diario.idlib) INNER JOIN (tes_cajaori LEFT JOIN (tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) ON (tes_cajaori.idmod = tes_cajaorigendet.idmod) AND (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes " _
            + vbCr + " WHERE ( tes_cajaorigendet.idtes is null and con_diario.tipo=1 ) " & nSQLWhere
            '+ vbCr + " WHERE ( con_diario.iddocpro=0 and con_diario.tipo=1 ) " & nSQLWhere
            
            
    '--ingresos y egresos(cuando es detallado) - origen (haber)
        nSQL = nSQL & vbCr & " Union All" _
            + vbCr + " SELECT tes_caja.id AS idmov, con_diario.idcue, mae_tipomov.descripcion AS Tipo1, tes_caja.idmon, mae_libros.descripcion AS libdesc, " _
            + vbCr + " '' AS tipdoc, '' AS numruc, '' AS apenom, tes_documentos.abrev AS tdocdesc, '' AS tdocsun, tes_caja.fchope AS fchdoc, con_tc.impven AS tc, " _
            + vbCr + " IIf([tes_cajaorigendet].[numser]='' Or [tes_cajaorigendet].[numser] Is Null,[tes_cajaorigendet].[numdoc],[tes_cajaorigendet].[numser] & '-' & [tes_cajaorigendet].[numdoc]) AS numdoc, " _
            + vbCr + " Mid([tes_caja].[numreg],1,2)+[mae_libros].[codsun]+Mid([tes_caja].[numreg],3,4) AS registro, Mid([tes_caja].[numreg],1,2)+[mae_libros].[codsun]+Mid([tes_caja].[numreg],3,4) AS registroref, " _
            + vbCr + " tes_caja.fchope, tes_caja.glosa, mae_libros.codsun AS libsun, CDbl([con_diario].[numasi]) AS corr, " _
            + vbCr + " IIf(Trim([tes_cajaorigendet].[numser])='' Or [tes_cajaorigendet].[numser] Is Null,'',[tes_cajaorigendet].[numser] & '-' & [tes_cajaorigendet].[numdoc]) AS docsustenta, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc " _
            + vbCr + " FROM ((mae_libros RIGHT JOIN (((tes_caja LEFT JOIN (((con_diario LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) LEFT JOIN tes_origen ON con_diario.idorides = tes_origen.id) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) ON tes_caja.id = con_diario.idmov) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN mae_tipomov ON tes_caja.tipmov = mae_tipomov.id) ON mae_libros.id = con_diario.idlib) INNER JOIN (tes_cajaori LEFT JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id " _
            + vbCr + " WHERE ( tes_cajaorigendet.idtes is not null and con_diario.tipo=1 ) " & nSQLWhere
            '+ vbCr + " WHERE ( con_diario.iddocpro<>0 and con_diario.tipo=1 ) " & nSQLWhere
    
    '--ingresos y egresos(cuando no es detallado) -- destino (debe)
        nSQL = nSQL & vbCr & " Union All" _
            + vbCr + " SELECT tes_caja.id AS idmov, con_diario.idcue, mae_tipomov.descripcion AS Tipo1, tes_caja.idmon, mae_libros.descripcion AS libdesc, " _
            + vbCr + " '' AS tipdoc, '' AS numruc, '' AS apenom, IIf([tes_cajadestinodet].[tipdoc] Is Null,'',[tes_documentos].[abrev]) AS tdocdesc, '' AS tdocsun, tes_caja.fchope AS fchdoc, con_tc.impven AS tc, " _
            + vbCr + " IIf([tes_cajadestinodet].[numdoc] Is Null,'',IIf(Trim([tes_cajadestinodet].[numser])='' Or [tes_cajadestinodet].[numser] Is Null,[tes_cajadestinodet].[numdoc],[tes_cajadestinodet].[numser] & '-' & [tes_cajadestinodet].[numdoc])) AS numdoc, " _
            + vbCr + " Mid([tes_caja].[numreg],1,2)+[mae_libros].[codsun]+Mid([tes_caja].[numreg],3,4) AS registro, " _
            + vbCr + " Mid([tes_caja].[numreg],1,2)+[mae_libros].[codsun]+Mid([tes_caja].[numreg],3,4) AS registroref, " _
            + vbCr + " tes_caja.fchope, tes_caja.glosa, mae_libros.codsun AS libsun, CDbl([con_diario].[numasi]) AS corr, IIf(Trim([tes_cajadestinodet].[numser])='' Or [tes_cajadestinodet].[numser] Is Null,'',[tes_cajadestinodet].[numser] & '-' & [tes_cajadestinodet].[numdoc]) AS docsustenta, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc " _
            + vbCr + " FROM mae_libros RIGHT JOIN (((((tes_caja LEFT JOIN ((con_diario LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) ON tes_caja.id = con_diario.idmov) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN mae_tipomov ON tes_caja.tipmov = mae_tipomov.id) INNER JOIN tes_cajadestino ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN (tes_cajadestinodet LEFT JOIN tes_documentos ON tes_cajadestinodet.tipdoc = tes_documentos.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON mae_libros.id = con_diario.idlib " _
            + vbCr + " WHERE (tes_cajadestinodet.idtes is null AND con_diario.tipo=2) " & nSQLWhere
    
        '--ingresos y egresos(cuando es detallado) - destino (debe)
            
        nSQL = nSQL & vbCr & " Union All" _
            + vbCr + " SELECT tes_caja.id AS idmov, diario.idcue, mae_tipomov.descripcion AS Tipo1, tes_caja.idmon, mae_libros.descripcion AS ibdesc, " _
            + vbCr + " IIF(destdet.iddoc=0,destdet.tipdoc, IIf(destdet.idmod=1,com.tipdoc,IIf(destdet.idmod=2,vta.tipdoc,IIf(destdet.idmod=8,bol.iddoc,IIf(destdet.idmod=9,hon.tipdoc,IIf(destdet.idmod=10,reem.idtipdoc,-1)))))) AS tipdoc, " _
            + vbCr + " IIF(destdet.iddoc=0,IIF(destdet.idmod=1 OR destdet.idmod=9 OR destdet.idmod=10 ,prov0.numruc,IIF(destdet.idmod=2,cli0.numruc,  IIF(destdet.idmod=8,emp0.numdoc,'Por Vincular')) ),IIf(destdet.idmod=1,prov1.numruc,IIf(destdet.idmod=2,cli2.numruc,IIf(destdet.idmod=8,emp3.numdoc,IIf(destdet.idmod=9,prov4.numruc,IIf(destdet.idmod=10,prov5.numruc,'Por Vincular')))))) AS numruc, " _
            + vbCr + " IIF(destdet.iddoc=0,IIF(destdet.idmod=1 OR destdet.idmod=9 OR destdet.idmod=10 ,prov0.nombre,IIF(destdet.idmod=2,cli0.nombre,  IIF(destdet.idmod=8,emp0.nombre,'Por Vincular')) ),IIf(destdet.idmod=1,prov1.nombre,IIf(destdet.idmod=2,cli2.nombre,IIf(destdet.idmod=8,emp3.nombre,IIf(destdet.idmod=9,prov4.nombre,IIf(destdet.idmod=10,prov5.nombre,'Por Vincular')))))) AS apenom, " _
            + vbCr + " IIF(destdet.iddoc=0,doc0.abrev,IIf(destdet.idmod=1,doc1.abrev,IIf(destdet.idmod=2,doc2.abrev,IIf(destdet.idmod=8,doc3.abrev,IIf(destdet.idmod=9,doc4.abrev,IIf(destdet.idmod=10,doc5.abrev,'Por Vincular')))))) AS tdocdesc, " _
            + vbCr + " IIF(destdet.iddoc=0,doc0.codsun,IIf(destdet.idmod=1,doc1.codsun,IIf(destdet.idmod=2,doc2.codsun,IIf(destdet.idmod=8,doc3.codsun,IIf(destdet.idmod=9,doc4.codsun,IIf(destdet.idmod=10,doc5.codsun,'Por Vincular')))))) AS tdocsun, " _
            + vbCr + " IIF(destdet.iddoc=0,destdet.fchdoc,IIf(destdet.idmod=1,com.fchdoc,IIf(destdet.idmod=2,vta.fchdoc,IIf(destdet.idmod=8,bol.fchdoc,IIf(destdet.idmod=9,hon.fchdoc,IIf(destdet.idmod=10,reem.fchdoc,Null)))))) AS fchdoc, " _
            + vbCr + " con_tc.impven AS tc, " _
            + vbCr + " IIF(destdet.iddoc=0,IIf(destdet.numser Is Null Or destdet.numser='','',destdet.numser & '-') & destdet.numdoc,IIf(destdet.idmod=1,IIf(com.numser Is Null Or com.numser='','',com.numser & '-') & com.numdoc,IIf(destdet.idmod=2,IIf(vta.numser Is Null Or vta.numser='','',vta.numser & '-') & vta.numdoc,IIf(destdet.idmod=8,IIf(bol.numser Is Null Or bol.numser='','',bol.numser & '-') & bol.numdoc,IIf(destdet.idmod=9,IIf(hon.numser Is Null Or hon.numser='','',hon.numser & '-') & hon.numdoc,IIf(destdet.idmod=10,IIf(reem.numser Is Null Or reem.numser='','',reem.numser & '-') & reem.numdoc,'Por Vincular')))))) AS numdoc, " _
            + vbCr + " Mid(tes_caja.numreg,1,2)+mae_libros.codsun+Mid(tes_caja.numreg,3,4) AS registro, " _
            + vbCr + " IIf(destdet.idmod=1,Mid(com.numreg,1,2)+lib1.codsun+Mid(com.numreg,3,4),IIf(destdet.idmod=2,Mid(vta.numreg,1,2)+lib2.codsun+Mid(vta.numreg,3,4),IIf(destdet.idmod=8,Mid(bol.numreg,1,2)+lib3.codsun+Mid(bol.numreg,3,4),IIf(destdet.idmod=9,Mid(hon.numreg,1,2)+lib4.codsun+Mid(hon.numreg,3,4),IIf(destdet.idmod=10,'-','Por Vincular'))))) AS registroref, " _
            + vbCr + " tes_caja.fchope, tes_caja.glosa, " _
            + vbCr + " IIF(destdet.iddoc=0,null,IIf(destdet.idmod=1,lib1.codsun,IIf(destdet.idmod=2,lib2.codsun,IIf(destdet.idmod=8,lib3.codsun,IIf(destdet.idmod=9,lib4.codsun,IIf(destdet.idmod=10,'','Por Vincular')))))) AS libsun, " _
            + vbCr + " IIF(destdet.iddoc=0,IIf(tes_caja.numreg Is Null,0,CDbl(Right(tes_caja.numreg,4))),IIf(destdet.idmod=1,IIf(com.numreg Is Null,0,CDbl(Right(com.numreg,4))),IIf(destdet.idmod=2,IIf(vta.numreg Is Null,0,CDbl(Right(vta.numreg,4))),IIf(destdet.idmod=8,IIf(bol.numreg Is Null,0,CDbl(Right(bol.numreg,4))),IIf(destdet.idmod=9,IIf(hon.numreg Is Null,0,CDbl(Right(hon.numreg,4))),IIf(destdet.idmod=10,10,-1)))))) AS corr, " _
            + vbCr + " IIF(destdet.iddoc=0,IIf(destdet.numser Is Null Or destdet.numser='','',destdet.numser & '-') & destdet.numdoc,IIf(destdet.idmod=1,IIf(com.numser Is Null Or com.numser='','',com.numser & '-') & com.numdoc,IIf(destdet.idmod=2,IIf(vta.numser Is Null Or vta.numser='','',vta.numser & '-') & vta.numdoc,IIf(destdet.idmod=8,IIf(bol.numser Is Null Or bol.numser='','',bol.numser & '-') & bol.numdoc,IIf(destdet.idmod=9,IIf(hon.numser Is Null Or hon.numser='','',hon.numser & '-') & hon.numdoc,IIf(destdet.idmod=10,IIf(reem.numser Is Null Or reem.numser='','',reem.numser & '-') & reem.numdoc,'Por Vincular')))))) AS docsustenta, " _
            + vbCr + " con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
            + vbCr + " IIf(diario.idmon=2,IIf(con_tc.impven Is Null,0,diario.impdebdol*con_tc.impven),diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(diario.idmon=2,IIf(con_tc.impven Is Null,0,diario.imphabdol*con_tc.impven),diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(diario.idmon=2,diario.impdebdol,IIf(con_tc.impven Is Null Or diario.impdebsol=0,0,(diario.impdebsol/con_tc.impven))) AS impdebedol, " _
            + vbCr + " IIf(diario.idmon=2,diario.imphabdol,IIf(con_tc.impven Is Null Or diario.imphabsol=0,0,(diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc "
            nSQL = nSQL _
            + vbCr + " FROM (((((mae_documento AS doc1 RIGHT JOIN ((com_reembolsables AS reem RIGHT JOIN ((((((((((tes_caja LEFT JOIN mae_tipomov ON tes_caja.tipmov = mae_tipomov.id) INNER JOIN tes_cajadestinodet AS destdet ON tes_caja.id = destdet.idtes) INNER JOIN (mae_libros RIGHT JOIN (((con_diario AS diario LEFT JOIN con_planctas ON diario.idcue = con_planctas.id) LEFT JOIN mae_moneda ON diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON diario.fchdoc = con_tc.fecha) ON mae_libros.id = diario.idlib) ON (destdet.corr = diario.iddocpro) AND (destdet.idtes = diario.idmov)) LEFT JOIN (mae_prov AS prov1 RIGHT JOIN (com_compras AS com " _
            + vbCr + " LEFT JOIN mae_libros AS lib1 ON com.idlib = lib1.id) ON prov1.id = com.idpro) ON destdet.iddoc = com.id) LEFT JOIN (((mae_cliente AS cli2 RIGHT JOIN vta_ventas AS vta ON cli2.id = vta.idcli) LEFT JOIN mae_libros AS lib2 ON vta.idlib = lib2.id) LEFT JOIN mae_documento AS doc2 ON vta.tipdoc = doc2.id) ON destdet.iddoc = vta.id) LEFT JOIN (((pla_empleados AS emp3 RIGHT JOIN pla_boleta AS bol ON emp3.id = bol.idemp) LEFT JOIN mae_libros AS lib3 ON bol.idlib = lib3.id) LEFT JOIN mae_documento AS doc3 ON bol.iddoc = doc3.id) ON destdet.iddoc = bol.id) LEFT JOIN com_honorarios AS hon ON destdet.iddoc = hon.id) LEFT JOIN mae_libros AS lib4 ON hon.idlib = lib4.id) " _
            + vbCr + " LEFT JOIN mae_documento AS doc4 ON hon.tipdoc = doc4.id) LEFT JOIN mae_prov AS prov4 ON hon.idpro = prov4.id) ON reem.id = destdet.iddoc) LEFT JOIN mae_prov AS prov5 ON reem.idpro = prov5.id) ON doc1.id = com.tipdoc) LEFT JOIN mae_documento AS doc5 ON reem.idtipdoc = doc5.id) LEFT JOIN pla_empleados AS emp0 ON destdet.idper = emp0.id) LEFT JOIN mae_prov AS prov0 ON destdet.idper = prov0.id) LEFT JOIN mae_cliente AS cli0 ON destdet.idper = cli0.id) LEFT JOIN mae_documento AS doc0 ON destdet.tipdoc = doc0.id " _
            + vbCr + " WHERE (destdet.idtes Is Not Null AND diario.tipo=2) " & nSQLWhere

            
    
    Case 8 '--Canjes de Facturas
        
        nSQL = "SELECT con_diario.idmov, IIf([con_canjesdet].[tipo]=1,[mae_cliente].[id],[mae_prov].[id]) AS idpro, con_diario.idcue, con_canjes.idmon, " _
            + vbCr + " mae_libros.descripcion AS libdesc, IIf([con_canjesdet].[tipo]=1,vta_ventas.tipdoc,com_compras.tipdoc) AS tipdoc, IIf([con_canjesdet].[tipo]=1,[mae_cliente].[numruc],[mae_prov].[numruc]) AS numruc, IIf([con_canjesdet].[tipo]=1,[mae_cliente].[nombre],[mae_prov].[nombre]) AS apenom, IIf([con_canjesdet].[tipo]=1,mae_documento_1.abrev,mae_documento.abrev) AS tdocdesc, IIf([con_canjesdet].[tipo]=1,mae_documento_1.codsun,mae_documento.codsun) AS tdocsun, IIf([con_canjesdet].[tipo]=1,vta_ventas.fchdoc,com_compras.fchdoc) AS fchdoc, con_tc.impven AS tc, IIf(con_canjes!numser Is Null Or con_canjes!numser='','',con_canjes!numser & '-') & con_canjes!numdoc AS numdoc, " _
            + vbCr + " Mid(con_canjes!numreg,1,2)+mae_libros!codsun+Mid(con_canjes!numreg,3,4) AS registro, " _
            + vbCr + " IIf([con_canjesdet].[tipo]=1, Mid(vta_ventas!numreg,1,2)+mae_libros_2!codsun+Mid(vta_ventas!numreg,3,4), Mid(com_compras!numreg,1,2)+mae_libros_1!codsun+Mid(com_compras!numreg,3,4)) AS registroref, " _
            + vbCr + " con_canjes.fchemi AS fchope, con_canjes.glosa, mae_libros.codsun AS libsun, CDbl(con_diario.numasi) AS corr, IIf([con_canjesdet].[tipo]=1,IIf(vta_ventas!numser Is Null Or vta_ventas!numser='','',vta_ventas!numser & '-') & vta_ventas!numdoc,IIf(com_compras!numser Is Null Or com_compras!numser='','',com_compras!numser & '-') & com_compras!numdoc) AS docsustenta, " _
            + vbCr + " con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
            + vbCr + " IIf([con_canjesdet].[tipo]=1,mae_moneda_2.simbolo,mae_moneda_1.simbolo) AS simbolo, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc " _
            + vbCr + " FROM mae_libros INNER JOIN (((((mae_documento AS mae_documento_1 RIGHT JOIN ((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) INNER JOIN ((((con_canjesdet LEFT JOIN com_compras ON con_canjesdet.iddoc = com_compras.id) LEFT JOIN vta_ventas ON con_canjesdet.iddoc = vta_ventas.id) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id) ON (con_diario.idmov = con_canjesdet.idcan) AND (con_diario.iddocpro = con_canjesdet.iddoc)) INNER JOIN con_canjes ON con_diario.idmov = con_canjes.id) " _
            + vbCr + " LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) ON mae_documento_1.id = vta_ventas.tipdoc) LEFT JOIN mae_moneda AS mae_moneda_1 ON com_compras.idmon = mae_moneda_1.id) LEFT JOIN mae_moneda AS mae_moneda_2 ON vta_ventas.idmon = mae_moneda_2.id) LEFT JOIN mae_libros AS mae_libros_2 ON vta_ventas.idlib = mae_libros_2.id) LEFT JOIN mae_libros AS mae_libros_1 ON com_compras.idlib = mae_libros_1.id) ON mae_libros.id = con_diario.idlib "
        
    Case 9 '--Planilla de Pago
        nSQL = "SELECT con_diario.idmov, pla_boleta.idemp AS idpro, con_diario.idcue, pla_boleta.idmon, mae_libros.descripcion AS libdesc, pla_boleta.iddoc AS tipdoc, pla_empleados.numdoc AS numruc, pla_empleados.nombre AS apenom, mae_documento.abrev AS tdocdesc, mae_documento.codsun AS tdocsun, pla_boleta.fchdoc, con_tc.impven AS tc, " _
            + vbCr + " IIf(pla_boleta!numser Is Null Or pla_boleta!numser='','',pla_boleta!numser & '-') & pla_boleta!numdoc AS numdoc, Mid(pla_boleta!numreg,1,2)+mae_libros!codsun+Mid(pla_boleta!numreg,3,4) AS registro, Mid(pla_boleta!numreg,1,2)+mae_libros!codsun+Mid(pla_boleta!numreg,3,4) AS registroref, " _
            + vbCr + " pla_boleta.fchdoc AS fchope, pla_boleta.glosa, mae_libros.codsun AS libsun, CDbl(con_diario.numasi) AS corr, IIf(pla_boleta!numser Is Null Or pla_boleta!numser='','',pla_boleta!numser & '-') & pla_boleta!numdoc AS docsustenta, " _
            + vbCr + " con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc " _
            + vbCr + " FROM pla_empleados INNER JOIN (((pla_boleta LEFT JOIN mae_documento ON pla_boleta.iddoc = mae_documento.id) LEFT JOIN mae_libros ON pla_boleta.idlib = mae_libros.id) LEFT JOIN (((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON pla_boleta.id = con_diario.idmov) ON pla_empleados.id = pla_boleta.idemp "
        
'    Case 37 '--Canjes de Letras
'
'    Case 38 '--Entregas a Rendir a Cuenta
'
'    Case 39 '--Rendición de Cuentas
    Case 41 '--Liquidacion Gasto Debito
        nSQL = "SELECT con_diario.idmov, vta_gastodebito.idcli, con_diario.idcue, con_diario.idmon, mae_libros.descripcion AS libdesc, vta_gastodebito.tipdoc, mae_cliente.numruc, mae_cliente.nombre AS apenom, mae_documento.abrev AS tdocdesc, mae_documento.codsun AS tdocsun, vta_gastodebito.fchemi AS fchdoc, con_tc.impven AS tc, " _
            + vbCr + " IIf([vta_gastodebito]![numser] Is Null Or [vta_gastodebito]![numser]='','',[vta_gastodebito]![numser] & '-') & [vta_gastodebito]![numdoc] AS numdoc, Mid([vta_gastodebito]![numreg],1,2)+[mae_libros]![codsun]+Mid([vta_gastodebito]![numreg],3,4) AS registro, Mid([vta_gastodebito]![numreg],1,2)+[mae_libros]![codsun]+Mid([vta_gastodebito]![numreg],3,4) AS registroref, vta_gastodebito.fchemi AS fchope, vta_gastodebito.glosa, mae_libros.codsun AS libsun, CDbl([con_diario].[numasi]) AS corr, IIf([vta_gastodebito]![numser] Is Null Or [vta_gastodebito]![numser]='','',[vta_gastodebito]![numser] & '-') & [vta_gastodebito]![numdoc] AS docsustenta, " _
            + vbCr + " con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol," _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc " _
            + vbCr + " FROM ((((mae_libros RIGHT JOIN (mae_cliente RIGHT JOIN vta_gastodebito ON mae_cliente.id = vta_gastodebito.idcli) ON mae_libros.id = vta_gastodebito.idlib) LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) LEFT JOIN (con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) ON vta_gastodebito.id = con_diario.idmov) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id "

    Case Else
        
        Exit Sub
        
    End Select
    
    '--cadena completada con el filtro del where
    If IDLIBRO <> 4 And IDLIBRO <> 5 And IDLIBRO <> 6 Then nSQL = nSQL + vbCr + nSQLWhere
    '--------
    
    '--remplazando segun la moneda seleccionada
    If NulosN(TxtIdMon.Text) = 1 Then
        nSQL = Replace(nSQL, "impdebesol", "debe")
        nSQL = Replace(nSQL, "imphabersol", "haber")
    Else
        nSQL = Replace(nSQL, "impdebedol", "debe")
        nSQL = Replace(nSQL, "imphaberdol", "haber")
    End If
    
    '--
    nSQL = "Select " & nSQLCampos & _
            vbCr + " from ( " _
            + vbCr + nSQL _
            + vbCr + ") as diario ORDER BY registro, ctanum ,numdoc"
    
    
    RST_Busq Rst, nSQL, xCon
    '--LIMPIAR ACUMULADO POR LIBRO
    xAcumulado(0, 0) = 0:   xAcumulado(0, 1) = 0
    xAcumulado(1, 0) = 0:   xAcumulado(1, 1) = 0
    '---------------------------------------------------------------------------
    If Rst.State = 0 Then GoTo SALIR
    If Rst.BOF = True Or Rst.EOF = True Or Rst.RecordCount = 0 Then GoTo SALIR
    Rst.MoveFirst
    Dim xAsiento
    Dim HabDol, DebDol  As Double
    Dim mCol& '--indica la posicion del campo
    
    Dim Cambiar As Boolean
    If Rst.RecordCount > 1 Then
        '--obtener el primer registro para evaluar el cambio de registro
        xAsiento = NulosC(Rst.Fields("registro"))
        
        ProgressBar1.Max = Rst.RecordCount
        ProgressBar1.Value = 1
        DoEvents
        
    End If
    Do While Not Rst.EOF
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        ProgressBar1.Value = Rst.Bookmark
        
        '-----------------------------------------------
        Fg1.Rows = Fg1.Rows + 1
        
        For mCol = 0 To Rst.Fields.Count - 1
        
            Select Case LCase(Rst.Fields(mCol).Name)
                Case "libdesc", "registro", "registroref", "glosa", "numruc", "apenom", "tdocdesc", "docsustenta", "ctanum", "ctadesc", "simbolo"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
                Case "fchdoc", "fchope"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_DATE)
                Case "tc"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), "0.000")
                Case "debe"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_MONTO)
                    xAcumulado(0, 0) = xAcumulado(0, 0) + NulosN(Rst.Fields("debe"))
                    mColDebe = mCol + 1
                Case "haber"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_MONTO)
                    xAcumulado(0, 1) = xAcumulado(0, 1) + NulosN(Rst.Fields("haber"))
                    mColHaber = mCol + 1
                    
                Case Else
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
            End Select
            
        Next mCol
        
        Rst.MoveNext
        
        If Rst.EOF = True Then
            Fg1.Rows = Fg1.Rows + 1
            
            If mColDebe > 2 Then Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total ==>"
            Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            If mColDebe > 2 Then FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            
            xAcumulado(1, 0) = xAcumulado(1, 0) + xAcumulado(0, 0)
            xAcumulado(1, 1) = xAcumulado(1, 1) + xAcumulado(0, 1)
            
            Exit Do
        End If
        
        'FechaRegistro
        If xAsiento <> Rst.Fields("registro") & "" Then
            Fg1.Rows = Fg1.Rows + 1
            If mColDebe > 2 Then Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total ==>"
            Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            If mColDebe > 2 Then FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            
            xAcumulado(1, 0) = xAcumulado(1, 0) + xAcumulado(0, 0)
            xAcumulado(1, 1) = xAcumulado(1, 1) + xAcumulado(0, 1)
            
            Fg1.Rows = Fg1.Rows + 1

            xAsiento = Rst.Fields("registro") & ""
            xAcumulado(0, 0) = 0:   xAcumulado(0, 1) = 0
        End If
                
    Loop
        
    Fg1.Rows = Fg1.Rows + 2
    If mColDebe > 2 Then Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total " + StrConv(N_Libro, 3) + " ==>"
    Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xAcumulado(1, 0), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xAcumulado(1, 1), FORMAT_MONTO)
    '--ACUMULAR LOS TOTALES POR LIBRO
    xAcumulado(2, 0) = xAcumulado(2, 0) + NulosN(Format(xAcumulado(1, 0), FORMAT_MONTO))
    xAcumulado(2, 1) = xAcumulado(2, 1) + NulosN(Format(xAcumulado(1, 1), FORMAT_MONTO))
    '-----------------------------------------------------------
    If mColDebe > 2 Then FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1
    FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe
    FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber
SALIR:
    Fg1.Rows = Fg1.Rows + 1
    Set Rst = Nothing
    
End Sub


Private Sub ProcesarResumen3(IDLIBRO As Integer)
    '===================================================================================================
    'creado: 20/11/08
    'Propósito: Mostrar el resumen del diario
    '
    'Entradas:  IDLIBRO = Código de Libro
    '           puede indicar un codigo is IDLIBRO<>0 o todos los libros si IDLIBRO=0
    '
    'Resultados: Informacion resumida del diario
    '===================================================================================================

    Dim RstRes As New ADODB.Recordset
    Dim xAcumulado(0, 1) As Double
    Erase xAcumulado()
    '--xAcumulado(0,?):: Acumulado por Asiento  ?::0=debe sol; 1::haber sol; 2::debe dol;  3::haber dol
    '--xAcumulado(1,?):: Acumulado por libro
    '--xAcumulado(2,?):: Acumulado general
    
    '*********************************************************************************************************************************
    
    Dim nSQL As String
    Dim nSQLLibro As String
    Dim nSQLSaldo As String
    Dim nSQLWhere As String
    Dim nSQLAjuste As String
       
    
    '--Para el libro
    If NulosN(IDLIBRO) <> 0 Then nSQLLibro = "(con_diario.idlib = " & IDLIBRO & ") and"
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) AND "
    '------
        
    If opt_fecha(0).Value = True Then  '--por fecha
        nSQLWhere = " WHERE ( ( " & nSQLAjuste & nSQLLibro & " con_diario.fchasi BETWEEN CDate('" & TxtFchIni.Valor & "') And CDate('" + TxtFchFin.Valor + "')) "
    Else '--por intervalo
        nSQLWhere = " WHERE ( ( " & nSQLAjuste & nSQLLibro & " con_diario.idmes >= " & mMesIni & " AND con_diario.idmes <= " & mMesFin & ") "
    End If
    
    '--para los saldos
    If opt_fecha(0).Value = True Then '--por fecha
        If CDate(Me.TxtFchIni.Valor) = CDate("01/01/" + AnoTra) Then
            'nSQLWhere = nSQLWhere & " OR ( " & nSQLAjuste & nSQLLibro & " con_diario.fchasi IS NULL) )"
            nSQLWhere = nSQLWhere & " )"
        Else
            nSQLWhere = nSQLWhere & " )"
        End If
    Else '--por periodo
        If mMesIni = 1 Then
            'nSQLWhere = nSQLWhere & "OR ( " & nSQLAjuste & nSQLLibro & " con_diario.fchasi IS NULL) )"
            nSQLWhere = nSQLWhere & " )"
        Else
            nSQLWhere = nSQLWhere & " )"
        End If
    End If

              
    nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS impdebesol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS imphabersol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven)))))) AS impdebedol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven)))))) AS imphaberdol " _
        + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
        + vbCr + nSQLWhere _
        + vbCr + " GROUP BY con_planctas.cuenta, con_planctas.descripcion;"
    
    '--remplazando segun la moneda seleccionada
    If NulosN(TxtIdMon.Text) = 1 Then
        nSQL = Replace(nSQL, "impdebesol", "debe")
        nSQL = Replace(nSQL, "imphabersol", "haber")
    Else
        nSQL = Replace(nSQL, "impdebedol", "debe")
        nSQL = Replace(nSQL, "imphaberdol", "haber")
    End If
    
    '*********************************************************************************************************************************
    
    RST_Busq RstRes, nSQL, xCon
    
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
        fg2.Rows = fg2.Rows + 1
        fg2.TextMatrix(fg2.Rows - 1, 1) = RstRes("cuenta") & ""
        fg2.TextMatrix(fg2.Rows - 1, 2) = RstRes("descripcion") & ""

        fg2.TextMatrix(fg2.Rows - 1, 3) = Format(NulosN(RstRes.Fields("debe")), FORMAT_MONTO)
        fg2.TextMatrix(fg2.Rows - 1, 4) = Format(NulosN(RstRes.Fields("haber")), FORMAT_MONTO)
        
        xAcumulado(0, 0) = xAcumulado(0, 0) + NulosN(Format(NulosN(RstRes.Fields("debe")), FORMAT_MONTO))
        xAcumulado(0, 1) = xAcumulado(0, 1) + NulosN(Format(NulosN(RstRes.Fields("haber")), FORMAT_MONTO))
        
        If NulosN(RstRes.Fields("debe")) = 0 And NulosN(RstRes.Fields("haber")) = 0 Then
            fg2.RemoveItem fg2.Rows - 1
        End If
        RstRes.MoveNext
        
        If RstRes.EOF = True Then
            
        End If
        
        
    Loop
    
    If xAcumulado(0, 0) <> 0 Or xAcumulado(0, 1) <> 0 Then
        fg2.Rows = fg2.Rows + 2
        fg2.TextMatrix(fg2.Rows - 1, 2) = "TOTAL ==>"
        fg2.TextMatrix(fg2.Rows - 1, 3) = Format(xAcumulado(0, 0), FORMAT_MONTO)
        fg2.TextMatrix(fg2.Rows - 1, 4) = Format(xAcumulado(0, 1), FORMAT_MONTO)
        
        FORMATO_CELDA fg2, fg2.Rows - 1, 2
        FORMATO_CELDA fg2, fg2.Rows - 1, 3
        FORMATO_CELDA fg2, fg2.Rows - 1, 4
    
    
'        '--cuadrar a la fuerza el detalle con el resumen
        Fg1.TextMatrix(Fg1.Rows - 2, mColDebe) = Format(xAcumulado(0, 0), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 2, mColHaber) = Format(xAcumulado(0, 1), FORMAT_MONTO)
    
        
        If xAcumulado(0, 0) <> xAcumulado(0, 1) Then
            fg2.Rows = fg2.Rows + 1
            fg2.TextMatrix(fg2.Rows - 1, 2) = "DIF ==>"
            If xAcumulado(0, 0) > xAcumulado(0, 1) Then
                    fg2.TextMatrix(fg2.Rows - 1, 4) = Format(xAcumulado(0, 0) - xAcumulado(0, 1), FORMAT_MONTO)
                Else
                    fg2.TextMatrix(fg2.Rows - 1, 3) = Format(xAcumulado(0, 1) - xAcumulado(0, 0), FORMAT_MONTO)
                End If
            FORMATO_CELDA fg2, fg2.Rows - 1, 2, vbRed, True
            FORMATO_CELDA fg2, fg2.Rows - 1, 3, vbRed, True
            FORMATO_CELDA fg2, fg2.Rows - 1, 4, vbRed, True
        End If
        
    
    End If
    
    
    
SALIR:
    Frame5.Visible = False
    MsgBox "El Diario se terminó de procesar con éxito", vbInformation, xTitulo
    Erase xAcumulado()
    Set RstRes = Nothing
End Sub



Private Sub ProcesarDiario4(IDLIBRO As Integer, N_Libro As String)
    '===================================================================================================
    'creado:     11/01/09
    'Propósito:  Mostrar la información del diario
    '
    'Entradas:   IDLIBRO = Código de Libro
    '            N_Libro = Desripción de Libro
    '
    'Resultados: Informacion de los diversos libros en pantalla
    '===================================================================================================
    Dim Rst As New ADODB.Recordset
    
    Dim nSQL As String
    Dim nSQLSaldo As String
    Dim nSQLWhere As String
    Dim nSQLCampos As String
    Dim nSQLAjuste As String
    Dim nSQLLibro As String
    
    Frame5.Left = 3413
    Frame5.Top = 2685
    Frame5.Visible = True
    ProgressBar1.Value = 1
    
    '---
    DoEvents
'    '**********************************************************************************************
    nSQLCampos = fSetearCuadriculaColumna(xCon, 1)
    If nSQLCampos = "" Then Exit Sub
'    '**********************************************************************************************
    '--generando el filtro
    '--Para el libro
    nSQLLibro = " con_diario.idlib = " & IDLIBRO & " AND "
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) AND "
    '------
        
    If opt_fecha(0).Value = True Then  '--por fecha
        nSQLWhere = " WHERE ( ( " & nSQLAjuste & nSQLLibro & " con_diario.fchasi BETWEEN CDate('" & TxtFchIni.Valor & "') And CDate('" + TxtFchFin.Valor + "')) "
    Else '--por intervalo
        nSQLWhere = " WHERE ( ( " & nSQLAjuste & nSQLLibro & " con_diario.idmes >= " & mMesIni & " AND con_diario.idmes <= " & mMesFin & ") "
    End If
    
    '--para los saldos
    If opt_fecha(0).Value = True Then '--por fecha
        If CDate(Me.TxtFchIni.Valor) = CDate("01/01/" + AnoTra) Then
            'nSQLWhere = nSQLWhere & " OR ( " & nSQLAjuste & nSQLLibro & " con_diario.fchasi IS NULL) )"
            nSQLWhere = nSQLWhere & " )"
        Else
            nSQLWhere = nSQLWhere & " )"
        End If
    Else '--por periodo
        If mMesIni = 1 Then
            'nSQLWhere = nSQLWhere & "OR ( " & nSQLAjuste & nSQLLibro & " con_diario.fchasi IS NULL) )"
            nSQLWhere = nSQLWhere & " )"
        Else
            nSQLWhere = nSQLWhere & " )"
        End If
        
        
    End If
    
    

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''SI HAY DOS MONEDAS EN UN MISMO DIA EN CON_TC HABILITAR LA SIGUIENTE LINEA DE CODIGO
''''''    nSQLWhere = nSQLWhere & " AND (con_tc.idmon=2 OR con_tc.idmon IS NULL ) "
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '----------------------------------------------------------------------------------
    '--generando la consulta
   'antes de 07/02/09
'   nSQL = "SELECT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, Format([con_diario].[idmes],'00') AS mes, mae_libros.codsun AS libsun, CDbl(con_diario.numasi) AS corr, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, con_diario.rglosa AS glosa, con_diario.rregistro AS registroref, iif(con_diario.ridtipper=0,tes_documentos.abrev,mae_documento.abrev) AS tdocdesc, con_diario.rfchope AS fchdoc, con_diario.rnumerodoc AS numdoc, " _
'            + vbCr + " IIf([con_diario].[ridtipper]=1,[mae_prov].[numruc],IIf([con_diario].[ridtipper]=2,[mae_cliente].[numruc],IIf([con_diario].[ridtipper]=3,[pla_empleados].[numdoc],''))) AS numruc, " _
'            + vbCr + " IIf([con_diario].[ridtipper]=1,[mae_prov].[nombre],IIf([con_diario].[ridtipper]=2,[mae_cliente].[nombre],IIf([con_diario].[ridtipper]=3,[pla_empleados].[apepat]&' '&[pla_empleados].[apemat]&', '&[pla_empleados].[nom],''))) AS apenom, mae_documento.codsun AS tdocsun, con_tc.impven AS tc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo as moneda, " _
'            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
'            + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
'            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
'             + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_diario.imphabsol=0,0,(con_diario.imphabsol/con_tc.impven))) AS imphaberdol, " _
'            + vbCr + " '' AS RefTDoc, '' AS RefNumDoc " _
'            + vbCr + " FROM (pla_empleados RIGHT JOIN (mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper) ON pla_empleados.id = con_diario.ridper) LEFT JOIN tes_documentos ON con_diario.rtipdoc = tes_documentos.id "


   nSQL = "SELECT Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null,'',mae_libros.codsun) & con_diario.numasi AS registro, Format(con_diario.idmes,'00') AS mes, mae_libros.codsun AS libsun, CDbl(con_diario.numasi) AS corr, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, con_diario.rglosaope as glosaope, con_diario.rglosa AS glosaref, con_diario.rregistro AS registroref, iif(con_diario.ridtipper=0,tes_documentos.abrev,mae_documento.abrev) AS tdocdesc, con_diario.rfchope AS fchdoc, con_diario.rnumerodoc AS numdoc, " _
            + vbCr + " IIf(con_diario.ridtipper=1,mae_prov.numruc,IIf(con_diario.ridtipper=2,mae_cliente.numruc,IIf(con_diario.ridtipper=3,pla_empleados.numdoc,IIf(con_diario.ridtipper=5,mae_bancos.numruc,'')))) AS numruc, " _
            + vbCr + " IIf(con_diario.ridtipper=1,mae_prov.nombre,IIf(con_diario.ridtipper=2,mae_cliente.nombre,IIf(con_diario.ridtipper=3,pla_empleados.apepat & ' ' & pla_empleados.apemat & ', ' & pla_empleados.nom,IIf(con_diario.ridtipper=5,mae_bancos.descripcion,'')))) AS apenom , mae_documento.codsun AS tdocsun, iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo AS monope, mae_moneda_1.simbolo AS monref, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) AS imphaberdol, " _
            + vbCr + " iif(con_diario.rnumerodoc1 is null,'',mae_documento_1.abrev) AS tdocdesc1, con_diario.rnumerodoc1 AS numdoc1, " _
            + vbCr + " tes_documentos_1.abrev AS tdocdesc2, con_diario.rfchope2 AS fchdoc2, con_diario.rnumerodoc2 AS numdoc2,con_diario.ridtipper2, iif(con_diario.ridtipper2<>5,'', mae_bancos_1.numruc ) AS numruc2,iif(con_diario.ridtipper2<>5,'',mae_bancos_1.descripcion ) AS apenom2 " _
            + vbCr + " FROM ((((((pla_empleados RIGHT JOIN (mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper) ON pla_empleados.id = con_diario.ridper) LEFT JOIN tes_documentos ON con_diario.rtipdoc = tes_documentos.id) LEFT JOIN mae_bancos ON con_diario.ridper = mae_bancos.id) LEFT JOIN mae_bancos AS mae_bancos_1 ON con_diario.ridper2 = mae_bancos_1.id) LEFT JOIN mae_documento AS mae_documento_1 ON con_diario.rtipdoc1 = mae_documento_1.id) LEFT JOIN tes_documentos AS tes_documentos_1 ON con_diario.rtipdoc2 = tes_documentos_1.id) LEFT JOIN mae_moneda AS mae_moneda_1 ON con_diario.ridmon = mae_moneda_1.id "

    '--cadena completada con el filtro del where
    nSQL = nSQL + vbCr + nSQLWhere
    '--------
    
    '--remplazando segun la moneda seleccionada
    If NulosN(TxtIdMon.Text) = 1 Then
        nSQL = Replace(nSQL, "impdebesol", "debe")
        nSQL = Replace(nSQL, "imphabersol", "haber")
    Else
        nSQL = Replace(nSQL, "impdebedol", "debe")
        nSQL = Replace(nSQL, "imphaberdol", "haber")
    End If
    
    '--
    nSQL = "Select " & nSQLCampos & _
            vbCr + " from ( " _
            + vbCr + nSQL _
            + vbCr + ") as diario ORDER BY registro, ctanum ,numdoc"
    
    
    RST_Busq Rst, nSQL, xCon
    
    '--obtener la posicione de las columnas debe,haber,saldo
    Dim mColCampo As Integer
    Dim mCol& '--indica la posicion del campo
    
    mCol = 0
    For mColCampo = 0 To Rst.Fields.Count - 1
        mCol = mCol + 1
        Select Case LCase(Rst.Fields(mColCampo).Name)
            Case "debe", "impdebesol", "impdebedol": mColDebe = mCol
            Case "haber", "imphabersol", "imphaberdol": mColHaber = mCol
            Case "registro": mPosRegistro = mCol
        End Select
    Next mColCampo
    
    
    '--LIMPIAR ACUMULADO POR LIBRO
    xAcumulado(0, 0) = 0:   xAcumulado(0, 1) = 0
    xAcumulado(1, 0) = 0:   xAcumulado(1, 1) = 0

    '******************************************************************************************************************
    '******************************************************************************************************************
    
    '--verificando si hay saldos iniciales excepto en apertura
    '--para los saldos
    nSQL = ""
    Dim nSQLFchIni As String
    
''''''''    If opt_fecha(0).Value = True Then '--por fecha
''''''''        If CDate(Me.TxtFchIni.Valor) > CDate("01/01/" + AnoTra) Then
''''''''            nSQLFchIni = " (((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "'))) "
''''''''        End If
''''''''    Else '--por periodo
''''''''        If mMesIni > 1 Then
''''''''            nSQLFchIni = " ((con_diario.fchasi) Is Null Or (con_diario.idmes)<" & mMesIni & " ) "
''''''''        End If
''''''''    End If
    
    If nSQLFchIni <> "" Then
         Dim RstSal As New ADODB.Recordset

        nSQL = "SELECT SaldosIni.libro, " _
            + vbCr + " SUM(IIF(SaldosIni.DebSol-SaldosIni.HabSol>0,SaldosIni.DebSol-SaldosIni.HabSol,0) ) AS SIDeb,  " _
            + vbCr + " SUM(IIF(SaldosIni.HabSol-SaldosIni.DebSol>0,SaldosIni.HabSol-SaldosIni.DebSol,0) ) AS SIHab  "
        
 '+ vbCr + " SUM( IIf(((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol)))>0,((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))),0) ) AS SIDeb, "
 '+ vbCr + " SUM( IIf(((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol)))>0,((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))),0) ) AS SIHab "
        
' + vbCr + " SUM( SaldosIni.DebSol ) AS SIDeb,
' + vbCr + " SUM( SaldosIni.HabSol ) AS SIHab

        If NulosN(TxtIdMon.Text) = 2 Then
            nSQL = Replace(nSQL, "DebSol", "DebDol")
            nSQL = Replace(nSQL, "HabSol", "HabDol")
        End If
        
        '--saldos iniciales
        nSQL = nSQL _
            + vbCr + " FROM ( " _
            + vbCr + " SELECT mae_libros.descripcion AS libro, con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
            + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
            + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
            + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
            + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
            + vbCr + " FROM ((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id " _
            + vbCr + " WHERE " & nSQLLibro & nSQLAjuste & nSQLFchIni _
            + vbCr + " GROUP BY mae_libros.descripcion, con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
            + vbCr + " ORDER BY con_planctas.cuenta " _
            + vbCr + " ) AS SaldosIni GROUP BY SaldosIni.libro "

        RST_Busq RstSal, nSQL, xCon

        If RstSal.RecordCount <> 0 Then
            Fg1.Rows = Fg1.Rows + 1
            
            If mColDebe > 2 Then Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total SI ==>"
            Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(NulosN(RstSal("SIDeb")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(NulosN(RstSal("SIHab")), FORMAT_MONTO)
            
            If mColDebe > 2 Then FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            
            Fg1.Rows = Fg1.Rows + 1
            
            xAcumulado(1, 0) = NulosN(RstSal("SIDeb"))
            xAcumulado(1, 1) = NulosN(RstSal("SIHab"))
           
        End If
    End If
    
    '******************************************************************************************************************
    '******************************************************************************************************************
    
    If Rst.State = 0 Then GoTo SALIR
    If Rst.BOF = True Or Rst.EOF = True Or Rst.RecordCount = 0 Then GoTo SALIR
        
    Rst.MoveFirst
    Dim xAsiento
    Dim HabDol, DebDol  As Double
    
    Dim Cambiar As Boolean
    If Rst.RecordCount > 1 Then
'''''Aplicando Orden
'''        Rst.Sort = "fchope,registro"
    
        '--obtener el primer registro para evaluar el cambio de registro
        xAsiento = NulosC(Rst.Fields("registro"))
        
        ProgressBar1.Max = Rst.RecordCount
        ProgressBar1.Value = 1
        DoEvents
        
        
    End If
    
    
    Do While Not Rst.EOF
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        ProgressBar1.Value = Rst.Bookmark
        
        '-----------------------------------------------
        Fg1.Rows = Fg1.Rows + 1
        
        For mCol = 0 To Rst.Fields.Count - 1
        
            Select Case LCase(Rst.Fields(mCol).Name)
                Case "libdesc", "registro", "registroref", "glosa", "numruc", "apenom", "tdocdesc", "docsustenta", "ctanum", "ctadesc", "simbolo"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
                Case "fchdoc", "fchope"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_DATE)
                Case "tc", "tipcam"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), "0.000")
                Case "debe"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_MONTO)
                    xAcumulado(0, 0) = xAcumulado(0, 0) + NulosN(Rst.Fields("debe"))
                    mColDebe = mCol + 1
                Case "haber"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_MONTO)
                    xAcumulado(0, 1) = xAcumulado(0, 1) + NulosN(Rst.Fields("haber"))
                    mColHaber = mCol + 1
                Case Else
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
                    
            End Select
            
        Next mCol
        
        Rst.MoveNext
        
        If Rst.EOF = True Then
            Fg1.Rows = Fg1.Rows + 1
            
            If mColDebe > 2 Then Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total ==>"
            Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            If mColDebe > 2 Then FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            
            xAcumulado(1, 0) = xAcumulado(1, 0) + xAcumulado(0, 0)
            xAcumulado(1, 1) = xAcumulado(1, 1) + xAcumulado(0, 1)
            
            Exit Do
        End If
        
        'FechaRegistro
        If xAsiento <> Rst.Fields("registro") & "" Then
            Fg1.Rows = Fg1.Rows + 1
            If mColDebe > 2 Then Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total ==>"
            Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xAcumulado(0, 0), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xAcumulado(0, 1), FORMAT_MONTO)
            
            If mColDebe > 2 Then FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber, IIf(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Fg1.TextMatrix(Fg1.Rows - 1, mColHaber), &H800000, &HFF)
            
            xAcumulado(1, 0) = xAcumulado(1, 0) + xAcumulado(0, 0)
            xAcumulado(1, 1) = xAcumulado(1, 1) + xAcumulado(0, 1)
            
            Fg1.Rows = Fg1.Rows + 1

            xAsiento = Rst.Fields("registro") & ""
            xAcumulado(0, 0) = 0:   xAcumulado(0, 1) = 0
        End If
                
    Loop
        
    Fg1.Rows = Fg1.Rows + 2
    If mColDebe > 2 Then Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total " + StrConv(N_Libro, 3) + " ==>"
    Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xAcumulado(1, 0), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xAcumulado(1, 1), FORMAT_MONTO)
    '--ACUMULAR LOS TOTALES POR LIBRO
    xAcumulado(2, 0) = xAcumulado(2, 0) + NulosN(Format(xAcumulado(1, 0), FORMAT_MONTO))
    xAcumulado(2, 1) = xAcumulado(2, 1) + NulosN(Format(xAcumulado(1, 1), FORMAT_MONTO))
    '-----------------------------------------------------------
    If mColDebe > 2 Then FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1
    FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe
    FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber
    
        
SALIR:
    Fg1.Rows = Fg1.Rows + 1
    If Rst.State = 1 Then
        If Rst.RecordCount = 0 Then
            '--eliminando los libros que no tienen movimiento
            If Fg1.Rows > 3 Then
                Fg1.Rows = Fg1.Rows - 3
            End If
        End If
    End If
    Set Rst = Nothing
    
End Sub

Private Sub pBuscarAsiento()
    Dim xfrm As New SGI2_funciones.formularios
    xfrm.AsientoBuscar xCon
    Set xfrm = Nothing
End Sub

Private Sub ProcesarResumen4(IDLIBRO As Integer)
     On Error GoTo error
    Dim RstRes As New ADODB.Recordset
    Dim A&
    Dim xTotal1, xTotal2, xTotal3, xTotal4 As Double
    Dim xAcumulado(7) As Double
    Dim nSQLAjuste  As String
    Dim nSQLIdLibro As String
    Dim nSQLCierre As String '--sentencia sql para no mostrar el cierre
    
    Frame5.Left = 3413
    Frame5.Top = 2685
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Value = 0
    
    pConfigurarGrilla
    
    Frame5.Visible = True
    Label3.Caption = "Procesando Resumen"
    fg2.Rows = fg2.FixedRows
    
    DoEvents
    
    Dim nSQLCuenta As String
    Dim nSQL As String
    '--------------------------
    nSQLCuenta = ""
    '--------------------------
    '**********************************************************************************************
    '--Para el libro
    If NulosN(IDLIBRO) <> 0 Then nSQLIdLibro = " AND (con_diario.idlib = " & IDLIBRO & ") "
    
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " AND (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    '**********************************************************************************************
    nSQLCierre = " AND (con_diario.idmes<>13) "
    
    '-----------------------------------------------
    '**********************************************************************************************


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
        + vbCr + " IIf(((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol)))>0,((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))),0) AS SIDeb, " _
        + vbCr + " IIf(((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol)))>0,((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))),0) AS SIHab, " _
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

    nSQL = nSQL _
        + vbCr + " FROM (con_planctas LEFT JOIN " _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
        + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id=con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc=con_tc.fecha " _
        + vbCr + " WHERE (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLCuenta & nSQLAjuste & nSQLIdLibro _
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
        + vbCr + " WHERE (((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "'))) " & nSQLCuenta & nSQLAjuste & nSQLIdLibro & nSQLCierre _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS SaldosIni "

    nSQLAjuste = nSQLIdLibro & nSQLAjuste & " AND (  (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))  OR  (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')   )"

    nSQL = nSQL _
        + vbCr + " ON con_planctas.id = SaldosIni.IdCta " _
        + vbCr + " WHERE con_planctas.id In (SELECT con_diario.idcue FROM con_diario " & IIf(nSQLIdLibro <> "", "WHERE " & Mid(nSQLIdLibro, 5), "") & " ) " & nSQLCuenta _
        + vbCr + " ORDER BY con_planctas.cuenta; "
    
    '--si seleccionar por periodo
    If opt_fecha(1).Value = True Then
        nSQL = Replace(nSQL, "(((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))", "con_diario.idmes>=" & mMesIni & " And con_diario.idmes <= " & mMesFin)
        nSQL = Replace(nSQL, "(con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')", "con_diario.idmes < " & mMesIni)
        
    End If

    '**************************************************************
    
    RST_Busq RstRes, nSQL, xCon

    If RstRes.State = 0 Then GoTo SALIR:
    RstRes.Filter = adFilterNone
    RstRes.Sort = "cuenta"
    
    If RstRes.BOF = False Or RstRes.EOF = False Or RstRes.RecordCount <> 0 Then RstRes.MoveFirst
    If RstRes.RecordCount > 0 Then ProgressBar1.Max = RstRes.RecordCount
    
    Label3.Caption = "Procesando Resumen"
    For A = 1 To RstRes.RecordCount
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        '-----------------------------------------------
        ProgressBar1.Value = A
        fg2.Rows = fg2.Rows + 1
        fg2.TextMatrix(A + 1, 1) = NulosC(RstRes("cuenta"))
        
        fg2.TextMatrix(A + 1, 2) = NulosC(RstRes("descri"))
        
        fg2.TextMatrix(A + 1, 3) = Format(NulosN(RstRes("SIDeb")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 4) = Format(NulosN(RstRes("SIHab")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 5) = Format(NulosN(RstRes("MPDeb")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 6) = Format(NulosN(RstRes("MPHab")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 7) = Format(NulosN(RstRes("SMDeb")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 8) = Format(NulosN(RstRes("SMHab")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 9) = Format(NulosN(RstRes("SADeb")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 10) = Format(NulosN(RstRes("SAHab")), FORMAT_MONTO)
        
        
        xAcumulado(0) = xAcumulado(0) + NulosN(fg2.TextMatrix(A + 1, 3)) '--saldeb
        xAcumulado(1) = xAcumulado(1) + NulosN(fg2.TextMatrix(A + 1, 4)) '--salhab
        xAcumulado(2) = xAcumulado(2) + NulosN(fg2.TextMatrix(A + 1, 5)) '--movdeb
        xAcumulado(3) = xAcumulado(3) + NulosN(fg2.TextMatrix(A + 1, 6)) '--movhab
        xAcumulado(4) = xAcumulado(4) + NulosN(fg2.TextMatrix(A + 1, 7)) '--maydeb
        xAcumulado(5) = xAcumulado(5) + NulosN(fg2.TextMatrix(A + 1, 8)) '--mayhab
        xAcumulado(6) = xAcumulado(6) + NulosN(fg2.TextMatrix(A + 1, 9)) '--deudor
        xAcumulado(7) = xAcumulado(7) + NulosN(fg2.TextMatrix(A + 1, 10)) '--acreedor
        
        RstRes.MoveNext
        If RstRes.EOF = True Then Exit For
    Next A
    
    fg2.Rows = fg2.Rows + 2
    
    fg2.TextMatrix(fg2.Rows - 1, 2) = "TOTAL =>"
    '-------------------------------
    Dim Col&
    
    fg2.AutoSizeMode = flexAutoSizeColWidth
    
    For A = 0 To UBound(xAcumulado())
        fg2.TextMatrix(fg2.Rows - 1, 3 + A) = Format(xAcumulado(A), FORMAT_MONTO)
        FORMATO_CELDA fg2, fg2.Rows - 1, 3 + A, , True
        fg2.AutoSize 3 + A
    Next A
    Erase xAcumulado()
    '-------------------------------
    If RstRes.RecordCount > 0 Then
        GRID_COLOR_FONDO fg2, 2, 3, fg2.Rows - 3, 4, RGB(255, 255, 236)
        GRID_COLOR_FONDO fg2, 2, 7, fg2.Rows - 3, 8, RGB(255, 255, 236)
        
    End If
        GRID_COLOR_FONDO fg2, fg2.Rows - 2, 1, fg2.Rows - 1, fg2.Cols - 1, RGB(231, 254, 224)
    
SALIR:
    Set RstRes = Nothing
    Frame5.Visible = False
    
    MsgBox "El Mayor se terminó de procesar con éxito", vbInformation, xTitulo
    
    Exit Sub
error:
    Frame5.Visible = False
    Set RstRes = Nothing
'    fra_msg.Visible = False
    SHOW_ERROR Me.Name, "CargarResumen"
End Sub


Private Sub pConfigurarGrilla()
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
                                   
    Dim k&, j&
    
    With fg2
        '-----
        
        .Cols = 11
        .Rows = 2
        .FixedRows = 2
        .FrozenCols = 2
        .RowHeight(0) = 500
        .RowHeight(1) = 350
        UNIR_CELDAS fg2, 0, 1, 0, 2, "Datos de la Cuenta", flexAlignCenterCenter
        UNIR_CELDAS fg2, 0, 3, 0, 4, "Saldos Iniciales", flexAlignCenterCenter
        UNIR_CELDAS fg2, 0, 5, 0, 6, "Movimiento del Periodo", flexAlignCenterCenter
        UNIR_CELDAS fg2, 0, 7, 0, 8, "Sumas del Mayor", flexAlignCenterCenter
        UNIR_CELDAS fg2, 0, 9, 0, 10, "Saldos Finales", flexAlignCenterCenter
        
'        '--DATOS DE FILA
        .TextMatrix(1, 1) = "Nº. Cuenta":       .ColWidth(1) = 1100:       .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(1, 2) = "Descripción":      .ColWidth(2) = 3000:       .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(1, 3) = "Debe":       .ColWidth(3) = 1300:       .ColAlignment(3) = flexAlignRightCenter
        .TextMatrix(1, 4) = "Haber":      .ColWidth(4) = 1300:       .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(1, 5) = "Debe":       .ColWidth(5) = 1320:       .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(1, 6) = "Haber":      .ColWidth(6) = 1320:       .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(1, 7) = "Debe":       .ColWidth(7) = 1320:       .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(1, 8) = "Haber":      .ColWidth(8) = 1320:       .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(1, 9) = "Debe":       .ColWidth(9) = 1320:       .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(1, 10) = "Haber":     .ColWidth(10) = 1320:      .ColAlignment(10) = flexAlignRightCenter
        
        '--AGREGANDO LAS FECHAS EN LA CABECERA
        If opt_fecha(0).Value = True Then
            If IsDate(TxtFchIni.Valor) = True Then UNIR_CELDAS fg2, 0, 3, 0, 4, "Saldos Iniciales" + vbCr + " Al " + CStr(CDate(TxtFchIni.Valor) - 1), flexAlignCenterCenter
            If IsDate(TxtFchFin.Valor) = True Then UNIR_CELDAS fg2, 0, 9, 0, 10, "Saldos Finales" + vbCr + " Al " + CStr(CDate(TxtFchFin.Valor) - 1), flexAlignCenterCenter
        Else
            If IsDate(TxtFchIni.Valor) = True Then UNIR_CELDAS fg2, 0, 3, 0, 4, "Saldos Iniciales" + vbCr + " A " + lbl_periodo(0).Caption, flexAlignCenterCenter
            If IsDate(TxtFchFin.Valor) = True Then UNIR_CELDAS fg2, 0, 9, 0, 10, "Saldos Finales" + vbCr + " A " + lbl_periodo(1).Caption, flexAlignCenterCenter
        End If
    End With
    
    DoEvents
End Sub






Private Sub ProcesarResumenLibros()
    '--21/04/09
    '--mostrar el resumen por libro
    
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
    
    pConfigurarGrilla
    
    Frame5.Visible = True
    Label3.Caption = "Procesando Resumen"
    fg2.Rows = fg2.FixedRows
    
    DoEvents
    
    Dim nSQLCuenta As String
    Dim nSQL As String
    '--------------------------
    nSQLCuenta = ""
    '--------------------------
    '**********************************************************************************************
    
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " AND (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    
    '-----------------------------------------------
    '**********************************************************************************************


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
    
    nSQL = "SELECT mae_libros.id ,mae_libros.descripcion AS descri , " _
        + vbCr + " IIf(((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol)))>0,((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))),0) AS SIDeb, " _
        + vbCr + " IIf(((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol)))>0,((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))),0) AS SIHab, " _
        + vbCr + " IIf(MovPeriodo.DebSol Is Null,0,MovPeriodo.DebSol) AS MPDeb, " _
        + vbCr + " IIf(MovPeriodo.HabSol Is Null,0,MovPeriodo.HabSol) AS MPHab, " _
        + vbCr + " [SIDeb]+[MPDeb] AS SMDeb,  " _
        + vbCr + " [SIHab]+[MPHab] AS SMHab, " _
        + vbCr + " IIf((SMDeb-SMHab)>0,(SMDeb-SMHab),0) AS SADeb, " _
        + vbCr + " IIf((SMHab-SMDeb)>0,(SMHab-SMDeb),0) AS SAHab "
    
    If NulosN(TxtIdMon.Text) = 2 Then
        nSQL = Replace(nSQL, "DebSol", "DebDol")
        nSQL = Replace(nSQL, "HabSol", "HabDol")
        
    End If

    nSQL = nSQL _
        + vbCr + " FROM (mae_libros LEFT JOIN " _
        + vbCr + " ( " _
        + vbCr + "  SELECT mae_libros.id as IdLib, mae_libros.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
        + vbCr + " FROM ((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id " _
        + vbCr + " WHERE (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLCuenta & nSQLAjuste _
        + vbCr + " GROUP BY mae_libros.id,mae_libros.descripcion " _
        + vbCr + " ORDER BY mae_libros.descripcion " _
        + vbCr + " ) AS MovPeriodo ON mae_libros.id = MovPeriodo.IdLib)  " _
        + vbCr + " Left Join "
    
    '--saldos iniciales
    nSQL = nSQL _
        + vbCr + " ( " _
        + vbCr + " SELECT mae_libros.id as IdLib, mae_libros.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
        + vbCr + " FROM ((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id " _
        + vbCr + " WHERE (((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "'))) " & nSQLCuenta & nSQLAjuste _
        + vbCr + " GROUP BY mae_libros.id,mae_libros.descripcion " _
        + vbCr + " ORDER BY mae_libros.descripcion" _
        + vbCr + " ) AS SaldosIni "

    nSQL = nSQL _
        + vbCr + " ON mae_libros.id = SaldosIni.IdLib " _
        + vbCr + " WHERE mae_libros.id In (SELECT con_diario.idlib FROM con_diario  ) " _
        + vbCr + " ORDER BY mae_libros.descripcion; "
    
    '--si seleccionar por periodo
    If opt_fecha(1).Value = True Then
        nSQL = Replace(nSQL, "(((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))", "con_diario.idmes>=" & mMesIni & " And con_diario.idmes <= " & mMesFin)
        nSQL = Replace(nSQL, "(con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')", "con_diario.idmes < " & mMesIni)
        
    End If

    '**************************************************************
    
    RST_Busq RstRes, nSQL, xCon

    If RstRes.State = 0 Then GoTo SALIR:
    RstRes.Filter = adFilterNone
    
    If RstRes.BOF = False Or RstRes.EOF = False Or RstRes.RecordCount <> 0 Then RstRes.MoveFirst
    If RstRes.RecordCount > 0 Then ProgressBar1.Max = RstRes.RecordCount
    
    Label3.Caption = "Procesando Resumen de Libros"
    For A = 1 To RstRes.RecordCount
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        '-----------------------------------------------
        ProgressBar1.Value = A
        fg2.Rows = fg2.Rows + 1
        fg2.TextMatrix(A + 1, 1) = NulosC(RstRes("id"))
        
        fg2.TextMatrix(A + 1, 2) = NulosC(RstRes("descri"))
        
        fg2.TextMatrix(A + 1, 3) = Format(NulosN(RstRes("SIDeb")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 4) = Format(NulosN(RstRes("SIHab")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 5) = Format(NulosN(RstRes("MPDeb")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 6) = Format(NulosN(RstRes("MPHab")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 7) = Format(NulosN(RstRes("SMDeb")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 8) = Format(NulosN(RstRes("SMHab")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 9) = Format(NulosN(RstRes("SADeb")), FORMAT_MONTO)
        fg2.TextMatrix(A + 1, 10) = Format(NulosN(RstRes("SAHab")), FORMAT_MONTO)
        
        
        xAcumulado(0) = xAcumulado(0) + NulosN(fg2.TextMatrix(A + 1, 3)) '--saldeb
        xAcumulado(1) = xAcumulado(1) + NulosN(fg2.TextMatrix(A + 1, 4)) '--salhab
        xAcumulado(2) = xAcumulado(2) + NulosN(fg2.TextMatrix(A + 1, 5)) '--movdeb
        xAcumulado(3) = xAcumulado(3) + NulosN(fg2.TextMatrix(A + 1, 6)) '--movhab
        xAcumulado(4) = xAcumulado(4) + NulosN(fg2.TextMatrix(A + 1, 7)) '--maydeb
        xAcumulado(5) = xAcumulado(5) + NulosN(fg2.TextMatrix(A + 1, 8)) '--mayhab
        xAcumulado(6) = xAcumulado(6) + NulosN(fg2.TextMatrix(A + 1, 9)) '--deudor
        xAcumulado(7) = xAcumulado(7) + NulosN(fg2.TextMatrix(A + 1, 10)) '--acreedor
        
        RstRes.MoveNext
        If RstRes.EOF = True Then Exit For
    Next A
    
    fg2.Rows = fg2.Rows + 2
    
    fg2.TextMatrix(fg2.Rows - 1, 2) = "TOTAL =>"
    '-------------------------------
    Dim Col&
    
    fg2.AutoSizeMode = flexAutoSizeColWidth
    
    For A = 0 To UBound(xAcumulado()) - 2
        fg2.TextMatrix(fg2.Rows - 1, 3 + A) = Format(xAcumulado(A), FORMAT_MONTO)
        FORMATO_CELDA fg2, fg2.Rows - 1, 3 + A, , True
        fg2.AutoSize 3 + A
    Next A
    Erase xAcumulado()
    '-------------------------------
    If RstRes.RecordCount > 0 Then
        GRID_COLOR_FONDO fg2, 2, 3, fg2.Rows - 3, 4, RGB(255, 255, 236)
        GRID_COLOR_FONDO fg2, 2, 7, fg2.Rows - 3, 8, RGB(255, 255, 236)
        
    End If
        GRID_COLOR_FONDO fg2, fg2.Rows - 2, 1, fg2.Rows - 1, fg2.Cols - 1, RGB(231, 254, 224)
    
SALIR:
    Set RstRes = Nothing
    Frame5.Visible = False
    
    MsgBox "El Mayor se terminó de procesar con éxito", vbInformation, xTitulo
    
    Exit Sub
error:
'    Resume
    Frame5.Visible = False
    Set RstRes = Nothing
'    fra_msg.Visible = False
    SHOW_ERROR Me.Name, "CargarResumen"
End Sub


Private Sub pImprimirRes()
    '--imprimir resumen
    TabOne1.CurrTab = 1
    If Me.TabOne1.CurrTab = 1 Then
        If fg2.Rows = 1 Then
            MsgBox "No hay registros para imprimir", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
    End If
    
    Dim nPeriodo   As String
    Dim xMoneda As String
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
    
    xMoneda = LblMoneda.Caption
    
    Dim RstTmp As New ADODB.Recordset
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 1) And ((con_formatostipodet.idformatotipo) = 3) " _
        & " And ((con_formatostipodet.mostrar) = -1)) ORDER BY con_formatostipodet.orden", xCon
    
    Dim xCampos() As String
    Dim xFil, xCol As Double
    
    ReDim xCampos(fg2.Rows - 2, fg2.Cols - 1)
    
    Dim xFila As Double
    xFila = 0
    For xFil = 1 To fg2.Rows - 1
        For xCol = 1 To fg2.Cols - 1
            xCampos(xFila, xCol) = fg2.TextMatrix(xFil, xCol)
        Next xCol
        xFila = xFila + 1
    Next xFil
    
    Rst.MoveFirst
    For A = 1 To Rst.RecordCount
        If NulosC(xCampos(0, A)) = NulosC(Rst("abrev")) Then
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
    xfrm.Titulo1 = "RESUMEN DEL DIARIO " & "(Expresado en " & xMoneda & ")"
    xfrm.Titulo2 = nPeriodo
    xfrm.TamañoFuente = 6
    xfrm.TamañoCabecera = 8
    xfrm.FuenteCabecera = "Courier New"
    xfrm.Posicion_Hoja = Vertical
    xfrm.Tamaño_Hoja = A_4
    xfrm.TextoConsiderar = " "
    xfrm.TextoConsiderarAncho = 5
    xfrm.ImprimirArray xCampos, Rst

    Set xfrm = Nothing
    Set Rst = Nothing
    
End Sub

