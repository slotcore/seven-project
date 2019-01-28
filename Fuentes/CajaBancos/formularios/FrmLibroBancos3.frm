VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmLibroBancos3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Libro Bancos"
   ClientHeight    =   7785
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11895
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12600
      Top             =   1470
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
            Picture         =   "FrmLibroBancos3.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":277E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":2A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":2E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLibroBancos3.frx":31BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   795
      Left            =   12000
      TabIndex        =   47
      Top             =   6840
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   105
         TabIndex        =   48
         Top             =   330
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
      Begin VB.Label Label5 
         Caption         =   "Procesando Libro Bancos"
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
         TabIndex        =   50
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
         TabIndex        =   49
         Top             =   90
         Visible         =   0   'False
         Width           =   1530
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
         X1              =   15
         X2              =   5925
         Y1              =   780
         Y2              =   765
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7440
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   11895
      _cx             =   20981
      _cy             =   13123
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
      Caption         =   "   Consulta   |   Detalle   "
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   0
      Position        =   0
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
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   7020
         Left            =   12540
         TabIndex        =   3
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6960
            Left            =   0
            TabIndex        =   5
            Top             =   60
            Width           =   11805
            _cx             =   20823
            _cy             =   12277
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
            TabOutlineColor =   0
            FrontTabForeColor=   -2147483630
            Caption         =   "   Libro Bancos   |   Conciliación   "
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
            Begin VB.Frame Frame11 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame11"
               Height          =   6930
               Left            =   315
               TabIndex        =   22
               Top             =   15
               Width           =   11475
               Begin VB.Frame Frame1 
                  Height          =   945
                  Left            =   30
                  TabIndex        =   26
                  Top             =   240
                  Width           =   9405
                  Begin VB.CommandButton CmdBusIdBanco 
                     Height          =   240
                     Left            =   3195
                     Picture         =   "FrmLibroBancos3.frx":34D6
                     Style           =   1  'Graphical
                     TabIndex        =   28
                     Top             =   210
                     Width           =   240
                  End
                  Begin VB.TextBox TxtCuenta 
                     Height          =   300
                     Left            =   1035
                     Locked          =   -1  'True
                     TabIndex        =   27
                     Text            =   "TxtCuenta"
                     Top             =   180
                     Width           =   2430
                  End
                  Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
                     Height          =   315
                     Left            =   1035
                     TabIndex        =   30
                     Top             =   510
                     Width           =   1260
                     _ExtentX        =   2223
                     _ExtentY        =   556
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Locked          =   -1  'True
                     Valor           =   "04/09/2009"
                  End
                  Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
                     Height          =   315
                     Left            =   3510
                     TabIndex        =   31
                     Top             =   510
                     Width           =   1260
                     _ExtentX        =   2223
                     _ExtentY        =   556
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Locked          =   -1  'True
                     Valor           =   "04/09/2009"
                  End
                  Begin VB.CommandButton cmdbuscar 
                     Enabled         =   0   'False
                     Height          =   405
                     Left            =   5970
                     Picture         =   "FrmLibroBancos3.frx":3608
                     Style           =   1  'Graphical
                     TabIndex        =   29
                     ToolTipText     =   "Buscar Movimientos de Bancos para conciliar"
                     Top             =   495
                     Width           =   1455
                  End
                  Begin VB.Label LblCtaNombre 
                     Caption         =   "Label7"
                     Height          =   375
                     Left            =   2820
                     TabIndex        =   52
                     Top             =   1020
                     Width           =   1695
                  End
                  Begin VB.Label LblCtaNum 
                     Caption         =   "Label6"
                     Height          =   255
                     Left            =   960
                     TabIndex        =   51
                     Top             =   1050
                     Width           =   1545
                  End
                  Begin VB.Label LblIdMoneda 
                     AutoSize        =   -1  'True
                     Caption         =   "LblIdMoneda"
                     ForeColor       =   &H000000FF&
                     Height          =   195
                     Left            =   7830
                     TabIndex        =   39
                     Top             =   705
                     Visible         =   0   'False
                     Width           =   930
                  End
                  Begin VB.Label LblIBcoCta 
                     AutoSize        =   -1  'True
                     Caption         =   "LblIBcoCta"
                     ForeColor       =   &H000000FF&
                     Height          =   195
                     Left            =   5010
                     TabIndex        =   38
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   780
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     Caption         =   "Fch. Inicio"
                     Height          =   195
                     Index           =   1
                     Left            =   105
                     TabIndex        =   37
                     Top             =   555
                     Width           =   735
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Nº Cuenta"
                     Height          =   195
                     Index           =   0
                     Left            =   105
                     TabIndex        =   36
                     Top             =   225
                     Width           =   840
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     Caption         =   "Fch. Final"
                     Height          =   195
                     Index           =   2
                     Left            =   2535
                     TabIndex        =   35
                     Top             =   555
                     Width           =   690
                  End
                  Begin VB.Label LblBanco 
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "LblBanco"
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
                     Left            =   3510
                     TabIndex        =   34
                     Top             =   180
                     Width           =   3915
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
                     Left            =   7440
                     TabIndex        =   33
                     Top             =   180
                     Width           =   1815
                  End
                  Begin VB.Label LblIdCuentaContable 
                     AutoSize        =   -1  'True
                     Caption         =   "LblIdCuentaContable"
                     ForeColor       =   &H000000FF&
                     Height          =   195
                     Left            =   7815
                     TabIndex        =   32
                     Top             =   510
                     Visible         =   0   'False
                     Width           =   1485
                  End
               End
               Begin VB.Frame Frame3 
                  Caption         =   "[  Ordenado Por  ]"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000040&
                  Height          =   945
                  Left            =   9480
                  TabIndex        =   43
                  Top             =   240
                  Width           =   1965
                  Begin VB.OptionButton OptSel3 
                     Caption         =   "Nº Documento"
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
                     Left            =   90
                     TabIndex        =   46
                     Top             =   690
                     Width           =   1770
                  End
                  Begin VB.OptionButton OptSel1 
                     Caption         =   "Fch. Operación"
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
                     Left            =   90
                     TabIndex        =   45
                     Top             =   210
                     Value           =   -1  'True
                     Width           =   1770
                  End
                  Begin VB.OptionButton OptSel2 
                     Caption         =   "Nº Registro"
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
                     Left            =   90
                     TabIndex        =   44
                     Top             =   450
                     Width           =   1560
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   5685
                  Left            =   45
                  TabIndex        =   25
                  Top             =   1200
                  Width           =   11385
                  _cx             =   20082
                  _cy             =   10028
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
                  BackColorSel    =   64
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
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmLibroBancos3.frx":3A4A
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
               Begin VB.Label Label8 
                  Alignment       =   2  'Center
                  Caption         =   "Detalle de Conciliación"
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   40
                  Top             =   30
                  Width           =   11235
               End
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H0000FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6930
               Left            =   12720
               TabIndex        =   6
               Top             =   15
               Width           =   11475
               Begin VB.CommandButton cmddocspendconc 
                  Caption         =   "Documentos Pend. de Conciliación"
                  Enabled         =   0   'False
                  Height          =   420
                  Left            =   6600
                  TabIndex        =   9
                  Top             =   30
                  Width           =   2715
               End
               Begin VB.CommandButton CmdVerconciliacion 
                  Caption         =   "&Ver Conciliación"
                  Height          =   420
                  Left            =   9480
                  TabIndex        =   8
                  Top             =   30
                  Width           =   1845
               End
               Begin VB.TextBox TxtImpExtracto 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   2520
                  TabIndex        =   7
                  Text            =   "TxtImpExtracto"
                  Top             =   120
                  Width           =   1200
               End
               Begin SizerOneLibCtl.TabOne TabOne3 
                  Height          =   3330
                  Left            =   15
                  TabIndex        =   10
                  Top             =   3600
                  Width           =   11400
                  _cx             =   20108
                  _cy             =   5874
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
                  Caption         =   "Cheques Emitidos y Otros Docs no cobrados (Mes Anterior)|Movimientos en Banco no considerados"
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
                  Begin VB.Frame Frame7 
                     BackColor       =   &H00008000&
                     BorderStyle     =   0  'None
                     Height          =   2910
                     Left            =   -11955
                     TabIndex        =   19
                     Top             =   45
                     Width           =   11310
                     Begin TrueOleDBGrid70.TDBGrid Dg4 
                        Height          =   2835
                        Left            =   30
                        TabIndex        =   42
                        Top             =   30
                        Width           =   11250
                        _ExtentX        =   19844
                        _ExtentY        =   5001
                        _LayoutType     =   4
                        _RowHeight      =   -2147483647
                        _WasPersistedAsPixels=   0
                        Columns(0)._VlistStyle=   0
                        Columns(0)._MaxComboItems=   5
                        Columns(0).Caption=   "Nº Id"
                        Columns(0).DataField=   "idmov"
                        Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                        Columns(1)._VlistStyle=   0
                        Columns(1)._MaxComboItems=   5
                        Columns(1).Caption=   "Nº Reg"
                        Columns(1).DataField=   "registro"
                        Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                        Columns(2)._VlistStyle=   0
                        Columns(2)._MaxComboItems=   5
                        Columns(2).Caption=   "T.D."
                        Columns(2).DataField=   "docabrev"
                        Columns(2).NumberFormat=   "Short Date"
                        Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                        Columns(3)._VlistStyle=   0
                        Columns(3)._MaxComboItems=   5
                        Columns(3).Caption=   "Nº Documento"
                        Columns(3).DataField=   "numerodoc"
                        Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                        Columns(4)._VlistStyle=   0
                        Columns(4)._MaxComboItems=   5
                        Columns(4).Caption=   "Fch  Ope"
                        Columns(4).DataField=   "fchope1"
                        Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                        Columns(5)._VlistStyle=   0
                        Columns(5)._MaxComboItems=   5
                        Columns(5).Caption=   "Medio Pago"
                        Columns(5).DataField=   "descmedpag"
                        Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                        Columns(6)._VlistStyle=   0
                        Columns(6)._MaxComboItems=   5
                        Columns(6).Caption=   "Glosa"
                        Columns(6).DataField=   "glosa"
                        Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                        Columns(7)._VlistStyle=   0
                        Columns(7)._MaxComboItems=   5
                        Columns(7).Caption=   "Total Debe"
                        Columns(7).DataField=   "impdebe1"
                        Columns(7).NumberFormat=   "0.00"
                        Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                        Columns(8)._VlistStyle=   0
                        Columns(8)._MaxComboItems=   5
                        Columns(8).Caption=   "Total Haber"
                        Columns(8).DataField=   "imphaber1"
                        Columns(8).NumberFormat=   "0.00"
                        Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                        Columns(9)._VlistStyle=   20
                        Columns(9)._MaxComboItems=   5
                        Columns(9).Caption=   "Conc"
                        Columns(9).DataField=   "xconc"
                        Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                        Columns(10)._VlistStyle=   0
                        Columns(10)._MaxComboItems=   5
                        Columns(10).Caption=   "Usuario"
                        Columns(10).DataField=   "usuario"
                        Columns(10).NumberFormat=   "0.00"
                        Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                        Columns.Count   =   11
                        Splits(0)._UserFlags=   0
                        Splits(0).Locked=   -1  'True
                        Splits(0).MarqueeStyle=   3
                        Splits(0).RecordSelectorWidth=   265
                        Splits(0)._SavedRecordSelectors=   0   'False
                        Splits(0).DividerColor=   12632256
                        Splits(0).FilterBar=   -1  'True
                        Splits(0).SpringMode=   0   'False
                        Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                        Splits(0)._ColumnProps(0)=   "Columns.Count=11"
                        Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
                        Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                        Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
                        Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                        Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8704"
                        Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
                        Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
                        Splits(0)._ColumnProps(8)=   "Column(1).Width=1482"
                        Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
                        Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1402"
                        Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
                        Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8704"
                        Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
                        Splits(0)._ColumnProps(14)=   "Column(2).Width=794"
                        Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
                        Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=714"
                        Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
                        Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8704"
                        Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
                        Splits(0)._ColumnProps(20)=   "Column(3).Width=2090"
                        Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
                        Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2011"
                        Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
                        Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=8704"
                        Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
                        Splits(0)._ColumnProps(26)=   "Column(4).Width=1455"
                        Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
                        Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1376"
                        Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
                        Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=8705"
                        Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
                        Splits(0)._ColumnProps(32)=   "Column(5).Width=1746"
                        Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
                        Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1667"
                        Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
                        Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=8704"
                        Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
                        Splits(0)._ColumnProps(38)=   "Column(6).Width=6720"
                        Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
                        Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=6641"
                        Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
                        Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=8704"
                        Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
                        Splits(0)._ColumnProps(44)=   "Column(7).Width=2196"
                        Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
                        Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2117"
                        Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
                        Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=8706"
                        Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
                        Splits(0)._ColumnProps(50)=   "Column(8).Width=1931"
                        Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
                        Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1852"
                        Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
                        Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=8706"
                        Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
                        Splits(0)._ColumnProps(56)=   "Column(9).Width=714"
                        Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
                        Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=635"
                        Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
                        Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=513"
                        Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
                        Splits(0)._ColumnProps(62)=   "Column(10).Width=159"
                        Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
                        Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=79"
                        Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
                        Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=8706"
                        Splits(0)._ColumnProps(67)=   "Column(10).Visible=0"
                        Splits(0)._ColumnProps(68)=   "Column(10).Order=11"
                        Splits.Count    =   1
                        PrintInfos(0)._StateFlags=   3
                        PrintInfos(0).Name=   "piInternal 0"
                        PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                        PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                        PrintInfos(0).PageHeaderHeight=   0
                        PrintInfos(0).PageFooterHeight=   0
                        PrintInfos.Count=   1
                        Appearance      =   0
                        ColumnFooters   =   -1  'True
                        DefColWidth     =   0
                        HeadLines       =   1.5
                        FootLines       =   1
                        MultipleLines   =   0
                        CellTipsWidth   =   0
                        DeadAreaBackColor=   12632256
                        RowDividerColor =   12632256
                        RowSubDividerColor=   12632256
                        DirectionAfterEnter=   0
                        MaxRows         =   250000
                        ViewColumnCaptionWidth=   0
                        ViewColumnWidth =   0
                        _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                        _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                        _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                        _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                        _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                        _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                        _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                        _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.fgcolor=&H0&,.bold=0"
                        _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                        _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
                        _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                        _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
                        _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                        _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
                        _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
                        _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
                        _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
                        _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                        _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
                        _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
                        _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
                        _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
                        _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
                        _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
                        _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                        _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
                        _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                        _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
                        _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                        _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                        _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                        _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                        _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
                        _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
                        _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                        _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                        _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                        _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.alignment=0,.locked=-1"
                        _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
                        _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
                        _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
                        _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0,.locked=-1"
                        _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
                        _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
                        _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
                        _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0,.locked=-1"
                        _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
                        _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
                        _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
                        _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=0,.locked=-1"
                        _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
                        _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
                        _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
                        _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2,.locked=-1"
                        _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
                        _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
                        _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
                        _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=0,.locked=-1"
                        _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
                        _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
                        _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
                        _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=78,.parent=13,.alignment=0,.locked=-1"
                        _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=14"
                        _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=15"
                        _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=17"
                        _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13,.alignment=1,.locked=-1"
                        _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
                        _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
                        _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
                        _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1,.locked=-1"
                        _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
                        _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
                        _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
                        _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=82,.parent=13,.alignment=2,.locked=0"
                        _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=79,.parent=14"
                        _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=80,.parent=15"
                        _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=81,.parent=17"
                        _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=1,.locked=-1"
                        _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
                        _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
                        _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
                        _StyleDefs(80)  =   "Named:id=33:Normal"
                        _StyleDefs(81)  =   ":id=33,.parent=0"
                        _StyleDefs(82)  =   "Named:id=34:Heading"
                        _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                        _StyleDefs(84)  =   ":id=34,.wraptext=-1"
                        _StyleDefs(85)  =   "Named:id=35:Footing"
                        _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                        _StyleDefs(87)  =   "Named:id=36:Selected"
                        _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                        _StyleDefs(89)  =   "Named:id=37:Caption"
                        _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
                        _StyleDefs(91)  =   "Named:id=38:HighlightRow"
                        _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                        _StyleDefs(93)  =   "Named:id=39:EvenRow"
                        _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                        _StyleDefs(95)  =   "Named:id=40:OddRow"
                        _StyleDefs(96)  =   ":id=40,.parent=33"
                        _StyleDefs(97)  =   "Named:id=41:RecordSelector"
                        _StyleDefs(98)  =   ":id=41,.parent=34"
                        _StyleDefs(99)  =   "Named:id=42:FilterBar"
                        _StyleDefs(100) =   ":id=42,.parent=33"
                     End
                     Begin VB.Label Label4 
                        AutoSize        =   -1  'True
                        Caption         =   "Total ==>"
                        Height          =   195
                        Index           =   1
                        Left            =   6930
                        TabIndex        =   20
                        Top             =   2640
                        Width           =   675
                     End
                  End
                  Begin VB.Frame Frame8 
                     BackColor       =   &H0000FF00&
                     BorderStyle     =   0  'None
                     Height          =   2910
                     Left            =   45
                     TabIndex        =   11
                     Top             =   45
                     Width           =   11310
                     Begin VB.TextBox TxtTotDeb 
                        Alignment       =   1  'Right Justify
                        Appearance      =   0  'Flat
                        Height          =   300
                        Left            =   7500
                        Locked          =   -1  'True
                        TabIndex        =   16
                        Text            =   "TxtTotDeb"
                        Top             =   2595
                        Width           =   1095
                     End
                     Begin VB.TextBox TxtTotHab 
                        Alignment       =   1  'Right Justify
                        Appearance      =   0  'Flat
                        Height          =   300
                        Left            =   8595
                        Locked          =   -1  'True
                        TabIndex        =   15
                        Text            =   "TxtTotHab"
                        Top             =   2595
                        Width           =   1095
                     End
                     Begin VB.Frame Frame2 
                        Height          =   2565
                        Left            =   9900
                        TabIndex        =   12
                        Top             =   30
                        Width           =   1425
                        Begin VB.CommandButton CmdDelMov 
                           Caption         =   "Eliminar Movimiento"
                           Height          =   645
                           Left            =   75
                           Style           =   1  'Graphical
                           TabIndex        =   14
                           Top             =   1350
                           Width           =   1260
                        End
                        Begin VB.CommandButton CmdAgregaMov 
                           Caption         =   "Agregar Movimiento"
                           Height          =   645
                           Left            =   75
                           Style           =   1  'Graphical
                           TabIndex        =   13
                           Top             =   555
                           Width           =   1260
                        End
                     End
                     Begin VSFlex7Ctl.VSFlexGrid Fg3 
                        Height          =   2565
                        Left            =   15
                        TabIndex        =   17
                        Top             =   15
                        Width           =   9825
                        _cx             =   17330
                        _cy             =   4524
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
                        BackColorSel    =   64
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
                        Rows            =   1
                        Cols            =   6
                        FixedRows       =   1
                        FixedCols       =   1
                        RowHeightMin    =   0
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   0   'False
                        FormatString    =   $"FrmLibroBancos3.frx":3BA8
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
                     Begin VB.Label Label2 
                        AutoSize        =   -1  'True
                        Caption         =   "Total ==>"
                        Height          =   195
                        Left            =   6765
                        TabIndex        =   18
                        Top             =   2640
                        Width           =   675
                     End
                  End
               End
               Begin TrueOleDBGrid70.TDBGrid Dg3 
                  Height          =   3105
                  Left            =   60
                  TabIndex        =   41
                  Top             =   480
                  Width           =   11355
                  _ExtentX        =   20029
                  _ExtentY        =   5477
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "Nº Id"
                  Columns(0).DataField=   "idmov"
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).Caption=   "Nº Reg"
                  Columns(1).DataField=   "registro"
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(2)._VlistStyle=   0
                  Columns(2)._MaxComboItems=   5
                  Columns(2).Caption=   "T.D."
                  Columns(2).DataField=   "docabrev"
                  Columns(2).NumberFormat=   "Short Date"
                  Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(3)._VlistStyle=   0
                  Columns(3)._MaxComboItems=   5
                  Columns(3).Caption=   "Nº Documento"
                  Columns(3).DataField=   "numerodoc"
                  Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(4)._VlistStyle=   0
                  Columns(4)._MaxComboItems=   5
                  Columns(4).Caption=   "Fch  Ope"
                  Columns(4).DataField=   "fchope1"
                  Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(5)._VlistStyle=   0
                  Columns(5)._MaxComboItems=   5
                  Columns(5).Caption=   "Medio Pago"
                  Columns(5).DataField=   "descmedpag"
                  Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(6)._VlistStyle=   0
                  Columns(6)._MaxComboItems=   5
                  Columns(6).Caption=   "Glosa"
                  Columns(6).DataField=   "glosa"
                  Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(7)._VlistStyle=   0
                  Columns(7)._MaxComboItems=   5
                  Columns(7).Caption=   "Total Debe"
                  Columns(7).DataField=   "impdebe1"
                  Columns(7).NumberFormat=   "0.00"
                  Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(8)._VlistStyle=   0
                  Columns(8)._MaxComboItems=   5
                  Columns(8).Caption=   "Total Haber"
                  Columns(8).DataField=   "imphaber1"
                  Columns(8).NumberFormat=   "0.00"
                  Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(9)._VlistStyle=   20
                  Columns(9)._MaxComboItems=   5
                  Columns(9).Caption=   "Conc"
                  Columns(9).DataField=   "xconc"
                  Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(10)._VlistStyle=   0
                  Columns(10)._MaxComboItems=   5
                  Columns(10).Caption=   "Usuario"
                  Columns(10).DataField=   "usuario"
                  Columns(10).NumberFormat=   "0.00"
                  Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   11
                  Splits(0)._UserFlags=   0
                  Splits(0).Locked=   -1  'True
                  Splits(0).MarqueeStyle=   3
                  Splits(0).RecordSelectorWidth=   265
                  Splits(0)._SavedRecordSelectors=   0   'False
                  Splits(0).DividerColor=   12632256
                  Splits(0).FilterBar=   -1  'True
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=11"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8704"
                  Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
                  Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(8)=   "Column(1).Width=1482"
                  Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1402"
                  Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
                  Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8704"
                  Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(14)=   "Column(2).Width=794"
                  Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
                  Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=714"
                  Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
                  Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8704"
                  Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
                  Splits(0)._ColumnProps(20)=   "Column(3).Width=2090"
                  Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
                  Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2011"
                  Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
                  Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=8704"
                  Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
                  Splits(0)._ColumnProps(26)=   "Column(4).Width=1455"
                  Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
                  Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1376"
                  Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
                  Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=8705"
                  Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
                  Splits(0)._ColumnProps(32)=   "Column(5).Width=1746"
                  Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
                  Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1667"
                  Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
                  Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=8704"
                  Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
                  Splits(0)._ColumnProps(38)=   "Column(6).Width=6720"
                  Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
                  Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=6641"
                  Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
                  Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=8704"
                  Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
                  Splits(0)._ColumnProps(44)=   "Column(7).Width=2196"
                  Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
                  Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2117"
                  Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
                  Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=8706"
                  Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
                  Splits(0)._ColumnProps(50)=   "Column(8).Width=1958"
                  Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
                  Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1879"
                  Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
                  Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=8706"
                  Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
                  Splits(0)._ColumnProps(56)=   "Column(9).Width=794"
                  Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
                  Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=714"
                  Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
                  Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=513"
                  Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
                  Splits(0)._ColumnProps(62)=   "Column(10).Width=1588"
                  Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
                  Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=1508"
                  Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
                  Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=8706"
                  Splits(0)._ColumnProps(67)=   "Column(10).Visible=0"
                  Splits(0)._ColumnProps(68)=   "Column(10).Order=11"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   3
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  Appearance      =   0
                  ColumnFooters   =   -1  'True
                  DefColWidth     =   0
                  HeadLines       =   1.5
                  FootLines       =   1
                  Caption         =   "Movimientos en Libros"
                  MultipleLines   =   0
                  CellTipsWidth   =   0
                  DeadAreaBackColor=   12632256
                  RowDividerColor =   12632256
                  RowSubDividerColor=   12632256
                  DirectionAfterEnter=   0
                  MaxRows         =   250000
                  ViewColumnCaptionWidth=   0
                  ViewColumnWidth =   0
                  _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.fgcolor=&H0&,.bold=0"
                  _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
                  _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
                  _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
                  _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
                  _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
                  _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
                  _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
                  _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
                  _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
                  _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
                  _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                  _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
                  _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
                  _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                  _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                  _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                  _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                  _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
                  _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
                  _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                  _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                  _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                  _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
                  _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
                  _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
                  _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
                  _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
                  _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
                  _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
                  _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
                  _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
                  _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
                  _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
                  _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2,.locked=-1"
                  _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
                  _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
                  _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
                  _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
                  _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
                  _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=78,.parent=13,.alignment=0,.locked=-1"
                  _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=14"
                  _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=15"
                  _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=17"
                  _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13,.alignment=1,.locked=-1"
                  _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
                  _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
                  _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
                  _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1,.locked=-1"
                  _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
                  _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
                  _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
                  _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=82,.parent=13,.alignment=2,.locked=0"
                  _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=79,.parent=14"
                  _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=80,.parent=15"
                  _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=81,.parent=17"
                  _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=1,.locked=-1"
                  _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
                  _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
                  _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
                  _StyleDefs(80)  =   "Named:id=33:Normal"
                  _StyleDefs(81)  =   ":id=33,.parent=0"
                  _StyleDefs(82)  =   "Named:id=34:Heading"
                  _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(84)  =   ":id=34,.wraptext=-1"
                  _StyleDefs(85)  =   "Named:id=35:Footing"
                  _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(87)  =   "Named:id=36:Selected"
                  _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(89)  =   "Named:id=37:Caption"
                  _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(91)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(93)  =   "Named:id=39:EvenRow"
                  _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(95)  =   "Named:id=40:OddRow"
                  _StyleDefs(96)  =   ":id=40,.parent=33"
                  _StyleDefs(97)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(98)  =   ":id=41,.parent=34"
                  _StyleDefs(99)  =   "Named:id=42:FilterBar"
                  _StyleDefs(100) =   ":id=42,.parent=33"
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Saldo Estado de Cuenta Bancario"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   21
                  Top             =   165
                  Width           =   2400
               End
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   7020
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6720
            Left            =   15
            TabIndex        =   4
            Top             =   300
            Width           =   11790
            _ExtentX        =   20796
            _ExtentY        =   11853
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Banco"
            Columns(1).DataField=   "banco"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nª Cta. Cte."
            Columns(2).DataField=   "numcue"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "M"
            Columns(3).DataField=   "simbolo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch Inicio"
            Columns(4).DataField=   "fchini"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fch. Final"
            Columns(5).DataField=   "fchfin"
            Columns(5).NumberFormat=   "Short Date"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Total Debe"
            Columns(6).DataField=   "impdeb"
            Columns(6).NumberFormat=   "0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Total Haber"
            Columns(7).DataField=   "imphab"
            Columns(7).NumberFormat=   "0.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Saldo"
            Columns(8).DataField=   "impsal"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=4233"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4154"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=3228"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3149"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=582"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=503"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1852"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1773"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=1773"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1693"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2434"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2355"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=2434"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2355"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=2090"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=2011"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   1.5
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.fgcolor=&H0&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
            _StyleDefs(72)  =   "Named:id=33:Normal"
            _StyleDefs(73)  =   ":id=33,.parent=0"
            _StyleDefs(74)  =   "Named:id=34:Heading"
            _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(76)  =   ":id=34,.wraptext=-1"
            _StyleDefs(77)  =   "Named:id=35:Footing"
            _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(79)  =   "Named:id=36:Selected"
            _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=37:Caption"
            _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(83)  =   "Named:id=38:HighlightRow"
            _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=39:EvenRow"
            _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(87)  =   "Named:id=40:OddRow"
            _StyleDefs(88)  =   ":id=40,.parent=33"
            _StyleDefs(89)  =   "Named:id=41:RecordSelector"
            _StyleDefs(90)  =   ":id=41,.parent=34"
            _StyleDefs(91)  =   "Named:id=42:FilterBar"
            _StyleDefs(92)  =   ":id=42,.parent=33"
         End
         Begin VB.Label LblMes 
            AutoSize        =   -1  'True
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   9000
            TabIndex        =   23
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Conciliación"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   15
            Width           =   11595
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Configurar"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "&Activar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "&Desactivar"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Activar &Todos Registros"
      End
      Begin VB.Menu Menu1_5 
         Caption         =   "Desactivar Todos Re&gistros"
      End
      Begin VB.Menu Menu1_6 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_7 
         Caption         =   "&Limpiar Filtro"
      End
      Begin VB.Menu Menu1_8 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_9 
         Caption         =   "&Exportar MSExcel"
      End
   End
End
Attribute VB_Name = "FrmLibroBancos3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim QueHace As Integer
Dim RstFrm As New ADODB.Recordset

Dim SeEjecuto As Boolean
Dim CaracteresNumericos As String
Dim mIdRegistro&
Dim Ordenar As String

Dim RstDg2 As New ADODB.Recordset '--registros del pediodo conciliados y no conciliados

Dim RstDg3 As New ADODB.Recordset '--registros del periodo pendientes de conciliar
Dim RstDg4 As New ADODB.Recordset '--registros anterior al periodo de busqueda

''--especfica el orden de la lista de la consulta
Dim fOrdenLista2 As Boolean
Dim fOrdenLista3 As Boolean
Dim fOrdenLista4 As Boolean

Dim mTipoGrid As Integer '--2:Dg2 ,3:Dg3; 4:Dg4

Dim mColDebe As Integer '--posicion de la columna debe
Dim mColHaber As Integer '--posicion de la columna haber
Dim mColSaldo As Integer '--posicion de la columna  saldo

Dim mMesActivo As Integer '--indica el mes activo

Dim mPosRegistro As Integer '--indica la posicion del numero de registro
Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim xHorIni As Date

'//////////////////////////////////////////////////////////
Dim oPDF As cPDF
Dim xFilaInicial As Integer
Dim xNumPag As Integer
'//////////////////////////////////////////////////////////
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Sub MuestraSegundoTab()

    Blanquea
    
    Dim rst As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    
    If RstFrm.RecordCount = 0 Then Exit Sub
    
    'Detalle del Documento
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    TabOne2.CurrTab = 0
    TabOne3.CurrTab = 0
    
    TxtFchIni.Valor = NulosC(RstFrm("fchini"))
    TxtFchFin.Valor = NulosC(RstFrm("fchfin"))
    
'    LblMoneda.Caption = RstFrm("moneda")
'    LblBanco.Caption = RstFrm("banco")
'    TxtCuenta.Text = RstFrm("numcta")
'    LblIdMoneda.Caption = RstFrm("idmon")
'    LblIBcoCta.Caption = RstFrm("idbcocta")
'    LblIdCuentaContable.Caption = RstFrm("idcuen")
    
            LblBanco.Caption = NulosC(RstFrm("banco"))
            TxtCuenta.Text = NulosC(RstFrm("numcue"))
            LblMoneda.Caption = NulosC(RstFrm("moneda"))
            
            LblIdCuentaContable.Caption = NulosN(RstFrm("idcuen"))
            LblIBcoCta.Caption = NulosN(RstFrm("idbcocta"))
            LblIdMoneda.Caption = NulosN(RstFrm("idmon"))
            
            LblCtaNum.Caption = NulosC(RstFrm("ctanum"))
            LblCtaNombre.Caption = NulosC(RstFrm("ctanombre"))
                
    
    
    TxtImpExtracto.Text = Format(NulosN(RstFrm("impsalbco")), FORMAT_MONTO)
    
    pCargarLBanco
    
    MostrarMovimientosPendientes
    

    'Visualizamos  OTROS MOVIMIENTOS
    RST_Busq RstDet, " SELECT * FROM tes_concidet WHERE idconc =" & RstFrm("id") & " and movimiento = 3 ORDER BY fchope ", xCon
    
    If RstDet.RecordCount <> 0 Then
        With Fg3
            Fg3.Rows = 1
            Do While Not RstDet.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, 1) = NulosC(RstDet("detalle"))
                .TextMatrix(.Rows - 1, 2) = NulosC(RstDet("fchope"))
                .TextMatrix(.Rows - 1, 3) = NulosC(RstDet("numdoc"))
                .TextMatrix(.Rows - 1, 4) = Format(NulosN(RstDet("impdeb")), FORMAT_MONTO)
                .TextMatrix(.Rows - 1, 5) = Format(NulosN(RstDet("imphab")), FORMAT_MONTO)
                
                RstDet.MoveNext
            Loop
        End With
    End If
    
    TxtTotDeb.Text = Format(GRID_SUMAR_COL(Fg3, 4), FORMAT_MONTO)
    TxtTotHab.Text = Format(GRID_SUMAR_COL(Fg3, 5), FORMAT_MONTO)
    
End Sub

Sub Modificar()
    
    Label8.Caption = "Modificando Conciliación Bancaria"
    
    QueHace = 2
    
    Bloquea True
    
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    TabOne1.TabEnabled(1) = True
    
    TabOne2.CurrTab = 0
    TabOne3.CurrTab = 0
    
    ActivaTool
    
    Dg3.BackColor = vbWhite
    Dg4.BackColor = vbWhite
    
    xHorIni = Time
    
    MuestraSegundoTab
    
    
End Sub

Sub Cancelar()

    Label4(0).Caption = "Consulta de Conciliación Bancaria"
    
    TabOne1.TabEnabled(0) = True
    
    TabOne1.CurrTab = 0


    Dg3.BackColor = &HE0FEFE
    Dg4.BackColor = &HE0FEFE

    Bloquea False
    
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
       
End Sub
Sub Nuevo()
    QueHace = 1
    Blanquea
    Bloquea True
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    TabOne1.TabEnabled(1) = True
    
    TabOne2.CurrTab = 0
    TabOne3.CurrTab = 0
    
    ActivaTool
    
    
    Dg3.BackColor = vbWhite
    Dg4.BackColor = vbWhite
    xHorIni = Time
    TxtCuenta.SetFocus
    Label8.Caption = "Nueva Conciliación Bancaria"
End Sub

Sub Bloquea(band As Boolean)
    
    TxtCuenta.Locked = Not band
    
    TxtFchIni.Locked = Not band
    TxtFchFin.Locked = Not band
    
    cmdbuscar.Enabled = band
    TxtImpExtracto.Enabled = band
    cmddocspendconc.Enabled = band
        
    Dg3.Splits(0).Locked = Not band
    Dg4.Splits(0).Locked = Not band
    
    CmdAgregaMov.Enabled = band
    CmdDelMov.Enabled = band
    
End Sub

Sub Blanquea()
    Fg1.Rows = Fg1.FixedRows
    Fg3.Rows = Fg3.FixedRows
    
    TxtCuenta.Text = ""
    LblBanco.Caption = ""
    LblMoneda.Caption = ""
    LblIdCuentaContable.Caption = ""
    LblIBcoCta.Caption = ""
    LblIdMoneda.Caption = ""
    TxtImpExtracto.Text = ""
    
    TxtTotHab.Text = "0.00"
    TxtTotDeb.Text = "0.00"

    
    Set RstDg2 = Nothing
    Set RstDg3 = Nothing
    Set RstDg4 = Nothing
    
    Set Dg3.DataSource = Nothing
    Set Dg4.DataSource = Nothing
    '--MOVIMIENTOS NO CONSIDERADOS EN EL PERIODO
    Dg3.Columns("impdebe1").FooterText = "0.00"
    Dg3.Columns("imphaber1").FooterText = "0.00"
    '--MOVIMIENTOS DE PERIODOS ANTERIORES
    Dg4.Columns("impdebe1").FooterText = "0.00"
    Dg4.Columns("imphaber1").FooterText = "0.00"
    '--LIMPIAR LOS FILTROS
    TDB_FiltroLimpiar Dg3
    TDB_FiltroLimpiar Dg4

End Sub

Sub ActivaTool()
    '--ACTIVAR TOOLBAR
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub


Function Grabar() As Boolean
    '--VALIDAR EL INGRESO DEL REGISTRO

    If fValidar() = False Then Exit Function
    
    '---LIMPIAR LOS FILTROS QUE SE LE HAYAN HECHO
    TDB_FiltroLimpiar Dg3
    TDB_FiltroLimpiar Dg4
    
    '---Otros movimientos no considerados
    Dim A As Integer
    
    TabOne2.CurrTab = 1
    TabOne3.CurrTab = 1
    For A = 1 To Fg3.Rows - 1
    
        If NulosC(Fg3.TextMatrix(A, 1)) = "" Then
            MsgBox "Falta agregar la Glosa a un registro en Movimientos no Considerados en Banco" & vbCr & "Caso contrario elimine dicho registro", vbInformation, xTitulo
            Fg3.Row = A
            Fg3.Col = 1
            Fg3.SetFocus
            Exit Function
            
        End If
        If IsDate(Fg3.TextMatrix(A, 2)) = False Then
            MsgBox "Falta agregar la Fecha a un registro en Movimientos no Considerados en Banco" & vbCr & "Caso contrario elimine dicho registro", vbInformation, xTitulo
            Fg3.Row = A
            Fg3.Col = 2
            Fg3.SetFocus
            Exit Function
            
        End If
        If NulosN(Fg3.TextMatrix(A, 4)) + NulosN(Fg3.TextMatrix(A, 5)) = 0 Then
            MsgBox "Falta agregar el importe a un registro en Movimientos no Considerados en Banco" & vbCr & "Caso contrario elimine dicho registro", vbInformation, xTitulo
            Fg3.Row = A
            
            Fg3.SetFocus
            Exit Function
            
        End If
    Next A
        
    
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    
    Dim xId As Double

    On Error GoTo LaCague:
    
    Me.MousePointer = vbHourglass
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("tes_conci", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM tes_conci ", xCon
        RstCab.AddNew
    Else
        xId = NulosN(RstFrm!id)
        
        RST_Busq RstCab, "SELECT  * FROM tes_conci WHERE id =" & xId & "", xCon
        
        '--ACTUALIZAR COMO NO CONCILIADO A TES_CAJA
        xCon.Execute "UPDATE tes_caja INNER JOIN tes_concidet ON tes_caja.id = tes_concidet.idmov " _
        & " SET tes_caja.conciliado = 0 WHERE (((tes_concidet.idconc)=" & xId & ") AND ((tes_concidet.conciliado)=-1));"
        
        '--ELIMINAMOS EL DETALLE DE LO CONCILIADO
        xCon.Execute "DELETE * FROM tes_concidet WHERE idconc = " & xId & ""
        
    End If
        
    RST_Busq RstDet, "SELECT TOP 1 * FROM tes_concidet", xCon
     
    mIdRegistro = xId
    
    RstCab("id") = xId
    RstCab("descripcion") = NulosC(LblBanco)
    RstCab("idmon") = NulosN(LblIdMoneda)
    RstCab("idbcocta") = NulosN(LblIBcoCta)
    RstCab("fchini") = CDate(TxtFchIni.Valor)
    RstCab("fchfin") = CDate(TxtFchFin.Valor)
    
    RstCab("impdeb") = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe))
    RstCab("imphab") = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, mColHaber))
    RstCab("impsal") = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, mColSaldo))
    RstCab("impsalbco") = NulosN(TxtImpExtracto.Text)
        
    RstCab.Update

    '---------------------------------------------------------------------------------------------------
    
    '--GRABAR LOS MOVIMIENTOS DEL PERIODO
    
    RstDg3.Filter = ""
    If RstDg3.RecordCount <> 0 Then
        '--se utiliza campo xinicio para que no considere la conciliacion de los mov de otros periodos
        RstDg3.Filter = "xconc=-1 and xinicio=0"
        
        Do While Not RstDg3.EOF
            
            RstDet.AddNew
            RstDet("idconc") = xId
            RstDet("idmov") = RstDg3("idmov")
            RstDet("impdeb") = NulosN(RstDg3("impdebe1"))
            RstDet("imphab") = NulosN(RstDg3("imphaber1"))
            RstDet("conciliado") = -1
            RstDet("movimiento") = 1
            RstDet.Update
            
            RstDg3.MoveNext
        Loop
    End If
    

    '---------------------------------------------------------------------------------------------------
    '--GRABAR LOS MOVIMIENTOS DE PERIODOS ANTERIORES
    If RstDg4.RecordCount <> 0 Then
        RstDg4.Filter = ""
        
        Do While Not RstDg4.EOF
            
            RstDet.AddNew
            RstDet("idconc") = xId
            RstDet("idmov") = RstDg4("idmov")
            RstDet("impdeb") = NulosN(RstDg4("impdebe1"))
            RstDet("imphab") = NulosN(RstDg4("imphaber1"))
            RstDet("conciliado") = NulosN(RstDg4("xconc"))
            RstDet("movimiento") = 2 '--movimientos anteriores
            
            RstDet.Update
            
            RstDg4.MoveNext
        Loop
    End If
    
    '---------------------------------------------------------------------------------------------------
    ' GRABAR OTROS MOVIMIENTOS NO CONSIDERADOS EN TESORERIA
    If Fg3.Rows - 1 <> 0 Then
        For A = 1 To Fg3.Rows - 1
            RstDet.AddNew
            RstDet("idconc") = xId
            RstDet("idmov") = 0
            RstDet("detalle") = NulosC(Fg3.TextMatrix(A, 1))
            If IsDate(Fg3.TextMatrix(A, 2)) = True Then RstDet("fchope") = CDate(Fg3.TextMatrix(A, 2))
            RstDet("numdoc") = NulosC(Fg3.TextMatrix(A, 3))
            RstDet("movimiento") = 3 '--otros movimientos
            RstDet("impdeb") = NulosN(Fg3.TextMatrix(A, 4))
            RstDet("imphab") = NulosN(Fg3.TextMatrix(A, 5))
''            RstDet.Cancel
            RstDet.Update
         
        Next
    End If

    '--ACTUALIZAR LOS MOVIMIENTOS CONCILIADOS, CONSIDERAR COMO CONCILIADO LOS MOVIMIENTOS ANULADOS O EXTRAVIADOS PARA CHEQUES
        
    xCon.Execute "UPDATE tes_caja INNER JOIN tes_concidet ON tes_caja.id = tes_concidet.idmov SET tes_caja.conciliado = -1 " _
        + vbCr + " WHERE (((tes_concidet.idconc)=" & xId & ") AND ((tes_concidet.conciliado)=-1) AND (([tes_concidet].[impdeb]+[tes_concidet].[imphab])<>0)) OR " _
        + vbCr + " (((tes_concidet.idconc)=" & xId & ") AND ((tes_concidet.conciliado)=-1) AND (([tes_concidet].[impdeb]+[tes_concidet].[imphab])=0) AND ((tes_caja.glosa) Like '%ANULADO%' Or (tes_caja.glosa) Like '%EXTRAVIADO%'));"

    '----------------------------------------------------------------------------------------------------------

    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    Set RstCab = Nothing
    Set RstDet = Nothing
    
    xCon.CommitTrans
    
    Grabar = True
    Me.MousePointer = vbDefault
    Exit Function
    
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Me.MousePointer = vbDefault
    MsgBox "No se a grabado el registro por el motivo sgte:" & vbCr & Err.Description & vbCr & Err.Source, vbCritical, xTitulo
    Err.Clear
    
    
End Function



Private Sub CmdAgregaMov_Click()

If QueHace = 3 Then Exit Sub
    Fg3.Rows = Fg3.Rows + 1
    
    
    
    With Fg3
        .Select Fg3.Rows - 1, 1, Fg3.Rows - 1, 1
    End With
End Sub

Private Sub cmdbuscar_Click()
    Dim rs As New ADODB.Recordset
    
    If QueHace = 3 Then Exit Sub

    If fValidar() = False Then Exit Sub


    If QueHace = 1 Then

    RST_Busq rs, "SELECT TOP 1 *  FROM Tes_conci WHERE idbcocta  =" & NulosN(LblIBcoCta) & " and fchini = CDate('" & TxtFchIni.Valor & "')  ORDER BY ID DESC ", xCon

        If rs.RecordCount <> 0 Then
            'Validamos las fechas
            
            If CDate(TxtFchIni.Valor) = CDate(rs!fchini) And CDate(TxtFchFin.Valor) = CDate(rs!fchfin) Then
                MsgBox "Existe una consulta con el mismo rango de fecha ya conciliado", vbInformation, Me.Caption
                Exit Sub
            ElseIf CDate(TxtFchIni.Valor) = CDate(rs!fchini) And CDate(TxtFchFin.Valor) < CDate(rs!fchfin) Then
                MsgBox "Existe una consulta con el mismo rango de fecha ya conciliado", vbInformation, Me.Caption
                Exit Sub
            ElseIf CDate(TxtFchIni.Valor) = CDate(rs!fchini) And CDate(TxtFchFin.Valor) > CDate(rs!fchfin) Then
            'MOSTRAMOS LA CONSULTA VALIDANDO LO CONCILIADO  COMPARAMOS LOS IMPORTES Y SI SON DIFERENTES DEBE CONCILIARSE
            'OTRA VEZ
                MostrarMovimientosPendientes
            End If
        Else
            pCargarLBanco
            MostrarMovimientosPendientes
        End If
        
    Else
            MostrarMovimientosPendientes
    End If


End Sub

Private Sub CmdBusIdBanco_Click()
    
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    
    If QueHace = 3 Then Exit Sub
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(5, 4) As String
    
    xCampos(0, 0) = "Banco":            xCampos(0, 1) = "banco":        xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Cta Cte":       xCampos(1, 1) = "numcue":       xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "M":                xCampos(2, 1) = "simbolo":      xCampos(2, 2) = "600":          xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nombre Cuenta":    xCampos(3, 1) = "ctanombre":    xCampos(3, 2) = "2800":         xCampos(3, 3) = "C"
    xCampos(4, 0) = "Cuenta":           xCampos(4, 1) = "ctanum":       xCampos(4, 2) = "1000":         xCampos(4, 3) = "C"
    'xCampos(5, 0) = "Codigo":           xCampos(5, 1) = "id":           xCampos(5, 2) = "800":          xCampos(5, 3) = "N"
    
    '--Lista de bancos con sus respectivas cuentas
    nSQL = "SELECT mae_banconumcta.id, mae_banconumcta.idcuen, mae_bancos.descripcion AS banco, mae_banconumcta.numcue, mae_moneda.descripcion AS moneda, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanombre, mae_banconumcta.idmon, mae_moneda.simbolo " _
        + vbCr + " FROM (mae_bancos RIGHT JOIN (con_planctas RIGHT JOIN mae_banconumcta ON con_planctas.id = mae_banconumcta.idcuen) ON mae_bancos.id = mae_banconumcta.idban) LEFT JOIN mae_moneda ON mae_banconumcta.idmon = mae_moneda.id " _
        + vbCr + " ORDER BY con_planctas.descripcion; "

    '--muestra ventana para seleccionar un registro
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando ", "ctanombre", "ctanombre", Principio
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            
            LblBanco.Caption = NulosC(xRs("banco"))
            TxtCuenta.Text = NulosC(xRs("numcue"))
            LblMoneda.Caption = NulosC(xRs("moneda"))
            
            LblIdCuentaContable.Caption = NulosN(xRs("idcuen"))
            LblIBcoCta.Caption = NulosN(xRs("id"))
            LblIdMoneda.Caption = NulosN(xRs("idmon"))
            
            LblCtaNum.Caption = NulosC(xRs("ctanum"))
            LblCtaNombre.Caption = NulosC(xRs("ctanombre"))
            
            TxtFchIni.SetFocus
        End If
    
    End If
    Set xRs = Nothing

End Sub

Private Sub CmdDelMov_Click()
    If QueHace = 3 Then Exit Sub
    
    If Fg3.Row < 1 Or Fg3.Rows < 1 Then Exit Sub
    
    
    Fg3.RemoveItem Fg3.Row
    TxtTotDeb.Text = Format(GRID_SUMAR_COL(Fg3, 4), FORMAT_MONTO)
    TxtTotHab.Text = Format(GRID_SUMAR_COL(Fg3, 5), FORMAT_MONTO)
    If Fg3.Rows <> 1 Then Fg3.Select Fg3.Rows - 1, 1

End Sub

Private Sub cmddocspendconc_Click()
    
    MostrarMovimientosPendientes

    
End Sub

Private Sub pCargarGrid()
    Dim nSQL  As String
    Dim Rpta As Integer
    
    nSQL = "SELECT Tes_conci.*, mae_moneda.descripcion AS moneda, mae_banconumcta.numcue, con_planctas.descripcion AS ctanombre, mae_banconumcta.idcuen, con_planctas.cuenta AS ctanum, mae_bancos.descripcion AS banco, mae_moneda.simbolo " _
        + vbCr + " FROM (con_planctas INNER JOIN (mae_banconumcta INNER JOIN (mae_moneda INNER JOIN Tes_conci ON mae_moneda.id = Tes_conci.idmon) ON mae_banconumcta.id = Tes_conci.idbcocta) ON con_planctas.id = mae_banconumcta.idcuen) LEFT JOIN mae_bancos ON mae_banconumcta.idban = mae_bancos.id " _
        + vbCr + " ORDER BY Tes_conci.id; "
    
    TDB_FiltroLimpiar Dg1
    Set RstFrm = Nothing
    
    '--cargando datos
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon

    Set Dg1.DataSource = RstFrm
    
    Me.MousePointer = vbDefault
    
    '--posicionar en la pestaña de consulta
    TabOne1.CurrTab = 0
    '************************************************

End Sub


Private Sub CmdVerconciliacion_Click()
    VerConciliacion
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstFrm
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 Then
        VerMovimientos1 IdMenuActivo, RstFrm("id"), xCon
    End If
End Sub

Private Sub Dg3_DblClick()
    If QueHace = 3 Then Exit Sub
    If RstDg3.State = 0 Then Exit Sub
    If RstDg3.RecordCount = 0 Then Exit Sub
    RstDg3("xconc") = Not RstDg3("xconc")
    
End Sub

Private Sub Dg3_FilterChange()
    TDB_FiltroGenerar Dg3, RstDg3
End Sub

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista3 = False Then nOrden = "ASC"
    If fOrdenLista3 = True Then nOrden = "DESC"
    RstDg3.Sort = CStr(Dg3.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista3 = Not fOrdenLista3
    Err.Clear
End Sub

Private Sub Dg3_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If RstDg3.State = 0 Then Exit Sub
    If KeyCode = vbKeySpace Or KeyCode = 13 Then
        If RstDg3.RecordCount <> 0 Then
            RstDg3("xconc") = Not RstDg3("xconc")
        End If
    End If
    If KeyCode = vbKeyF5 Then
        TDB_FiltroLimpiar Dg3
        RstDg3.Filter = ""
        If RstDg3.RecordCount <> 0 Then RstDg3.MoveFirst
    End If
End Sub

Private Sub Dg3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mTipoGrid = 3
    If QueHace <> 3 Then
        If Button = 2 Then PopupMenu Menu1
    Else
        
    End If
End Sub

Private Sub Dg4_DblClick()
    If QueHace = 3 Then Exit Sub
    If RstDg4.State = 0 Then Exit Sub
    If RstDg4.RecordCount = 0 Then Exit Sub
    RstDg4("xconc") = Not RstDg4("xconc")
End Sub

Private Sub Dg4_FilterChange()
    TDB_FiltroGenerar Dg4, RstDg4
End Sub

Private Sub Dg4_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista4 = False Then nOrden = "ASC"
    If fOrdenLista4 = True Then nOrden = "DESC"
    RstDg4.Sort = CStr(Dg4.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista4 = Not fOrdenLista4
    Err.Clear
End Sub

Private Sub Dg4_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If RstDg4.State = 0 Then Exit Sub
    If KeyCode = vbKeySpace Or KeyCode = 13 Then
        If RstDg4.RecordCount <> 0 Then RstDg4("xconc") = Not RstDg4("xconc")
    End If
    If KeyCode = vbKeyF5 Then
        TDB_FiltroLimpiar Dg4
        RstDg4.Filter = ""
        If RstDg4.RecordCount <> 0 Then RstDg4.MoveFirst
    End If
End Sub

Private Sub Dg4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mTipoGrid = 4
    If QueHace <> 3 Then
        If Button = 2 Then PopupMenu Menu1
    Else
        
    End If
End Sub


Private Sub Fg3_CellChanged(ByVal Row As Long, ByVal Col As Long)

    If QueHace = 3 Then Exit Sub
    If Col = 2 Then
        If IsDate(Fg3.TextMatrix(Row, Col)) = False Then
            Fg3.TextMatrix(Row, Col) = ""
        Else
            Fg3.TextMatrix(Row, Col) = CDate(Fg3.TextMatrix(Row, Col))
        End If
    End If
    If Col = 4 Or Col = 5 Then
        If NulosN(Fg3.TextMatrix(Fg3.Row, 4)) > 0 And NulosN(Fg3.TextMatrix(Fg3.Row, 5)) > 0 Then
          MsgBox "Montos Debe y Haber uno de ellos debe ser 0", vbInformation, Me.Caption
        End If
        
        TxtTotDeb.Text = Format(GRID_SUMAR_COL(Fg3, 4), FORMAT_MONTO)
        TxtTotHab.Text = Format(GRID_SUMAR_COL(Fg3, 5), FORMAT_MONTO)
        
    End If

End Sub

Private Sub Fg3_EnterCell()
    If QueHace = 3 Then
        Fg3.Editable = flexEDNone
    Else
        Fg3.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg3_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

  
    If KeyAscii = 13 Then Exit Sub
    
        Select Case Col
            Case 4, 5
                            
                If validar_numero(KeyAscii) = False Then KeyAscii = 0
                
                TxtTotDeb.Text = Format(GRID_SUMAR_COL(Fg3, 4), FORMAT_MONTO)
                TxtTotHab.Text = Format(GRID_SUMAR_COL(Fg3, 5), FORMAT_MONTO)
                
        End Select
        
End Sub

Private Sub Fg3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then CmdDelMov_Click
End Sub



Private Sub Form_Activate()
    
    If SeEjecuto = False Then
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon

        
        '--LIMPIANDO LOS CONTROLES
        TxtCuenta.Text = ""
        LblBanco.Caption = ""
        LblMoneda.Caption = ""
        
        '--ESTABLECIENDO A LA FECHA ACTUAL PARA FACILITAR LA BUSQUEDA
        TxtFchIni.Valor = Date
        TxtFchFin.Valor = Date
        
        TxtFchIni.Valor = ""
        TxtFchFin.Valor = ""
        TxtTotDeb.Text = "0.00"
        TxtTotHab.Text = "0.00"
       
        TxtImpExtracto.CausesValidation = False
        
        
'''        lbl_periodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
'''        lbl_periodo(1).Caption = lbl_periodo(0).Caption
'''        mMesIni = xMes
'''        mMesFin = xMes

        TabOne1.CurrTab = 0

        'TxtFchIni.SetFocus
        
        pCargarGrid
        
        SeEjecuto = True
        
        
    End If
    
    
End Sub
Sub Eliminar()
    Dim Rpta As Integer
    
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay conciliaciones para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    TabOne1.CurrTab = 0
    Rpta = MsgBox("¿ Esta seguro de eliminar la conciliación  ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        
        '--ACTUALIZAR COMO NO CONCILIADO LOS MOVIMIENTOS CONCILIADOS
        xCon.Execute "UPDATE tes_caja INNER JOIN tes_concidet ON tes_caja.id = tes_concidet.idmov " _
        & " SET tes_caja.conciliado = 0 WHERE (((tes_concidet.idconc)=" & RstFrm("id") & ") AND ((tes_concidet.conciliado)=-1));"
        
        '--ELIMINAR LA CONCILIACION
        xCon.Execute "DELETE * FROM tes_concidet  WHERE idconc = " & RstFrm("id") & " "
        xCon.Execute "DELETE * FROM tes_conci WHERE id = " & RstFrm("id") & " "
               
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstFrm("id") & " AND idform = " & IdMenuActivo
               
        MsgBox " se elimino con exito la conciliación ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        
        RstFrm.Requery
        Dg1.Refresh
        If RstFrm.RecordCount = 0 Then
            Rpta = MsgBox("No se han registrado movimientos en el periodo especificado, ¿ Desea agregar uno ahora ?", vbInformation + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then Nuevo
        End If
    End If
End Sub


Sub MostrarMovimientosPendientes()
    '===================================================================================================
    'Creado : xx/xx/09 Por: Johan Castro
    'Propósito: Cargar en Recordset temporales los movimientos de banco para conlicliar
    '
    'Entradas:  Ninguna
    '
    'Resultados: 2 Recorset's temporales
    '               Rst para las operaciones del periodo de consulta
    '               Rst para las operaciones de periodos anteriores
    '
    '===================================================================================================
    
    Dim nSQL  As String
    Dim RstTmp As New ADODB.Recordset '--rst temporal para hacer la consulta y destinarlo al RstDg2,RsDg3,RstDg4

    If fValidar() = False Then Exit Sub
    '----------------------------------------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------------------------------------
    
    '--CARGANDO LAS OPERACIONES PENDIENTES POR CONCILIAR EN EL PERIODO DE BUSQUEDA
    
    '--xconc Campo para identificar mov conciliados
    '--xinicio Campo para identificar si mov esta conciliado en otros periodos(toma valor -1)
    
    If QueHace = 1 Then
        nSQL = "SELECT tes_caja.id as idmov, tes_caja.fchope, tes_caja.fchope & '' AS fchope1, tes_cajaorigendet.idmod, tes_documentos.abrev AS docabrev, tes_documentos.descripcion AS descdoc, tes_caja.glosa, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Mid([tes_caja].[numreg],3) AS registro, tes_cajaorigendet.numser, tes_cajaorigendet.numdoc, " _
            + vbCr + " IIf([tes_caja].[tipmov]=1, IIf([tes_cajaorigendet].[idtes] Is Null,tes_cajaori.importe,[tes_cajaorigendet].[importe]),0) & '' AS ImpDebe1, " _
            + vbCr + " IIf([tes_caja].[tipmov]=2, IIf([tes_cajaorigendet].[idtes] Is Null,tes_cajaori.importe,[tes_cajaorigendet].[importe]),0) & '' AS ImpHaber1, " _
            + vbCr + " tes_origen.descripcion AS descori, con_planctas.cuenta, tes_mediopago.descripcion AS descmedpag, tes_origen.idcuen, " _
            + vbCr + " 0 as xconc ,0 as xinicio, tes_caja.conciliado , tes_caja.idmon, iif(tes_cajaorigendet.numser is null or tes_cajaorigendet.numser='','',tes_cajaorigendet.numser & '-') &   tes_cajaorigendet.numdoc as numerodoc, mae_moneda.simbolo, tes_cajaori.tc AS tipcam,tes_caja.tipmov " _
            + vbCr + " FROM ((tes_caja LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) RIGHT JOIN (((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) LEFT JOIN ((tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) LEFT JOIN tes_mediopago ON tes_cajaorigendet.idmedpag = tes_mediopago.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id " _
            + vbCr + " WHERE (((tes_caja.fchope) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')) AND ((tes_cajaori.idbcocta)=" & NulosN(LblIBcoCta.Caption) & ")) " _
            + vbCr + " ORDER BY tes_caja.fchope; "
    Else
        nSQL = "SELECT tes_caja.id AS idmov, tes_caja.fchope, tes_caja.fchope & '' AS fchope1, tes_cajaorigendet.idmod, tes_documentos.abrev AS docabrev, tes_documentos.descripcion AS descdoc, tes_caja.glosa, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Mid([tes_caja].[numreg],3) AS registro, tes_cajaorigendet.numser, tes_cajaorigendet.numdoc, IIf([tes_caja].[tipmov]=1,IIf([tes_cajaorigendet].[idtes] Is Null,tes_cajaori.importe,[tes_cajaorigendet].[importe]),0) & '' AS ImpDebe1, IIf([tes_caja].[tipmov]=2,IIf([tes_cajaorigendet].[idtes] Is Null,tes_cajaori.importe,[tes_cajaorigendet].[importe]),0) & '' AS ImpHaber1, tes_origen.descripcion AS descori, con_planctas.cuenta, tes_mediopago.descripcion AS descmedpag, tes_origen.idcuen, " _
                + vbCr + " IIf(vista.conc Is Null,0,vista.conc) AS xconc,IIF(xconc=0 and tes_caja.conciliado=-1,-1,0) as xinicio,tes_caja.conciliado , tes_caja.idmon, IIf(tes_cajaorigendet.numser Is Null Or tes_cajaorigendet.numser='','',tes_cajaorigendet.numser & '-') & tes_cajaorigendet.numdoc AS numerodoc, mae_moneda.simbolo, tes_cajaori.tc AS tipcam, tes_caja.tipmov " _
                + vbCr + " FROM ((((tes_caja LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) RIGHT JOIN ((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN ((tes_cajaorigendet " _
                + vbCr + " LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) LEFT JOIN tes_mediopago ON tes_cajaorigendet.idmedpag = tes_mediopago.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN " _
                + vbCr + " ( SELECT tes_concidet.idconc, tes_concidet.idmov, tes_concidet.conciliado AS conc FROM tes_concidet  " _
                + vbCr + " WHERE (((tes_concidet.movimiento)=1) AND ((tes_concidet.idconc)=" & NulosN(RstFrm("id")) & ")) ) as vista  ON tes_caja.id = vista.idmov " _
                + vbCr + " WHERE (((tes_caja.fchope) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')) AND ((tes_cajaori.idbcocta)=" & NulosN(LblIBcoCta.Caption) & ")) " _
                + vbCr + " ORDER BY tes_caja.fchope; "
    
    End If
    '--SI SE MODIFICA LA CONCILIACION =>> NO SE MOSTRAR LOS MOV. QUE YA FUERON CONCILIADOS
''    If QueHace <> 1 Then
''
''        nSQL = Replace(nSQL, "WHERE", "WHERE tes_caja.id NOT IN (  SELECT tes_concidet.idmov  " _
''        + vbCr + " FROM tes_concidet INNER JOIN tes_conci ON tes_concidet.idconc = tes_conci.Id " _
''        + vbCr + " WHERE (((tes_concidet.idconc)<>" & NulosN(RstFrm("id")) & ") AND ((tes_conci.idbcocta)=" & NulosN(LblIBcoCta.Caption) & ") AND ((tes_concidet.conciliado)=-1) AND (([tes_concidet].[impdeb]+[tes_concidet].[imphab])<>0) AND ((tes_concidet.movimiento)=2)) )  AND ")
''
''    End If
''
    

    RST_Busq RstTmp, nSQL, xCon

    
    '--ESTABLECIENDO EL TIPO DE ORDEN SEGUN SELECCION POR USUARIO
    If OptSel1.Value = True Then RstTmp.Sort = "fchope"
    If OptSel2.Value = True Then RstTmp.Sort = "registro"
    If OptSel3.Value = True Then RstTmp.Sort = "numerodoc"
    
    '--DEFINIR LA ESTRURCTURA DEL RST PARA ALMACENAR LA INFORMACION TEMPORAL
    DEFINIR_RST_TMP RstDg3, RstTmp
    
    '--CARGAR LA INFORMACION AL RST
    CARGAR_RST_TMP RstDg3, RstTmp
    
    '--LIMPIANDO EL RST TEMPORAL
    Set RstTmp = Nothing
    
    '--CARGANDO EL GRID
    Set Dg3.DataSource = RstDg3
    
    '--MOSTRAR LOS RESUMENES
    Dg3.Columns("glosa").FooterText = "Total ==>>"
    Dg3.Columns("impdebe1").FooterText = Format(RstRegistroSumar(RstDg3, "impdebe1"), FORMAT_MONTO)
    Dg3.Columns("imphaber1").FooterText = Format(RstRegistroSumar(RstDg3, "imphaber1"), FORMAT_MONTO)
    
    '----------------------------------------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------------------------------------

    '--CARGAMOS LOS MOVIMIENTOS PENDIENTES DE OTROS PERIODOS
            
    '--cuando se agrega un nuevo registro
    If QueHace = 1 Then
        nSQL = "SELECT tes_caja.id as idmov,[tes_caja].[fchope], [tes_caja].[fchope] & '' AS fchope1, tes_cajaorigendet.idmod, tes_documentos.abrev AS docabrev,tes_documentos.descripcion AS descdoc, tes_caja.glosa, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Mid([tes_caja].[numreg],3) AS registro, tes_caja.idmon, " _
            + vbCr + " 0 as xconc ,0 as xinicio, tes_caja.conciliado, tes_cajaorigendet.numser, tes_cajaorigendet.numdoc, " _
            + vbCr + " IIf([tes_cajaorigendet].[idtes] Is Null,tes_cajaori.importe,IIf([tes_caja].[tipmov]=1,[tes_cajaorigendet].[importe],0)) AS ImpDebe1, " _
            + vbCr + " IIf([tes_cajaorigendet].[idtes] Is Null,tes_cajaori.importe,IIf([tes_caja].[tipmov]=2,[tes_cajaorigendet].[importe],0)) AS ImpHaber1, " _
            + vbCr + " tes_origen.descripcion AS descori, con_planctas.cuenta, tes_mediopago.descripcion AS descmedpag, tes_origen.idcuen, iif(tes_cajaorigendet.numser is null or tes_cajaorigendet.numser='','',tes_cajaorigendet.numser & '-') &   tes_cajaorigendet.numdoc as numerodoc, mae_moneda.simbolo, tes_cajaori.tc AS tipcam,tes_caja.tipmov " _
            + vbCr + " FROM ((tes_caja LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) RIGHT JOIN (((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) LEFT JOIN ((tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) LEFT JOIN tes_mediopago ON tes_cajaorigendet.idmedpag = tes_mediopago.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id " _
            + vbCr + " WHERE (((tes_caja.fchope)<CDate('" & TxtFchIni.Valor & "')) AND ((tes_cajaori.idbcocta)=" & NulosN(LblIBcoCta) & ") AND ((tes_caja.idmon)=" & NulosN(LblIdMoneda.Caption) & ") AND ((tes_caja.conciliado)=0)) " _
            + vbCr + " ORDER BY [tes_caja].[fchope] & '';"

    Else
        '--si se modifica el registro considerar lo sgte.
        '--1ra consulta:lista de todos los pendientes sin considerar los conciliados(no se considera como conciliado cuando la suma de impdeb e imphab=0).
        '--2da consulta:lista de solo los conciliados
        
        nSQL = "SELECT tes_caja.id as idmov,[tes_caja].[fchope], [tes_caja].[fchope] & '' AS fchope1, tes_cajaorigendet.idmod, tes_documentos.abrev AS docabrev, tes_documentos.descripcion AS descdoc, tes_caja.glosa, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Mid([tes_caja].[numreg],3) AS registro, tes_caja.idmon, " _
            + vbCr + " 0 AS xconc,0 as xinicio,tes_caja.conciliado, tes_cajaorigendet.numser, tes_cajaorigendet.numdoc, " _
            + vbCr + " IIf([tes_cajaorigendet].[idtes] Is Null,tes_cajaori.importe,IIf([tes_caja].[tipmov]=1,[tes_cajaorigendet].[importe],0)) & '' AS ImpDebe1, " _
            + vbCr + " IIf([tes_cajaorigendet].[idtes] Is Null,tes_cajaori.importe,IIf([tes_caja].[tipmov]=2,[tes_cajaorigendet].[importe],0)) & '' AS ImpHaber1, " _
            + vbCr + " tes_origen.descripcion AS descori, con_planctas.cuenta, tes_mediopago.descripcion AS descmedpag, tes_origen.idcuen, iif(tes_cajaorigendet.numser is null or tes_cajaorigendet.numser='','',tes_cajaorigendet.numser & '-') &   tes_cajaorigendet.numdoc as numerodoc, mae_moneda.simbolo, tes_cajaori.tc AS tipcam,tes_caja.tipmov " _
            + vbCr + " FROM ((tes_caja LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) RIGHT JOIN (((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) LEFT JOIN ((tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) LEFT JOIN tes_mediopago ON tes_cajaorigendet.idmedpag = tes_mediopago.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id " _
            + vbCr + " WHERE (((tes_caja.id) Not In (SELECT tes_concidet.idmov FROM tes_concidet WHERE (((tes_concidet.idconc)=" & NulosN(RstFrm("id")) & ") AND ((tes_concidet.movimiento)=2) AND ((tes_concidet.conciliado)=-1) AND (([tes_concidet].[impdeb]+[tes_concidet].[imphab])<>0)))) AND ((tes_caja.conciliado)=0) AND ((tes_caja.fchope)<CDate('" & TxtFchIni.Valor & "')) AND ((tes_cajaori.idbcocta)=" & NulosN(LblIBcoCta.Caption) & ")) " _
            + vbCr + " UNION " _
            + vbCr + " SELECT tes_caja.id, [tes_caja].[fchope],[tes_caja].[fchope] & '' AS fchope1, tes_cajaorigendet.idmod,tes_documentos.abrev AS docabrev, tes_documentos.descripcion AS descdoc, tes_caja.glosa, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Mid([tes_caja].[numreg],3) AS registro, tes_caja.idmon, " _
            + vbCr + " -1 AS xconc,0 as xinicio, tes_concidet.conciliado, tes_cajaorigendet.numser, tes_cajaorigendet.numdoc, " _
            + vbCr + " tes_concidet.impdeb & '' AS ImpDebe1, tes_concidet.imphab  & '' AS ImpHaber1,  " _
            + vbCr + " tes_origen.descripcion AS descori, con_planctas.cuenta, tes_mediopago.descripcion AS descmedpag, tes_origen.idcuen,  iif(tes_cajaorigendet.numser is null or tes_cajaorigendet.numser='','',tes_cajaorigendet.numser & '-') &   tes_cajaorigendet.numdoc as numerodoc, mae_moneda.simbolo, tes_cajaori.tc AS tipcam,tes_caja.tipmov " _
            + vbCr + " FROM ((tes_caja LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) RIGHT JOIN ((((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) LEFT JOIN tes_concidet ON tes_cajaori.idtes = tes_concidet.idmov) LEFT JOIN ((tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) LEFT JOIN tes_mediopago ON tes_cajaorigendet.idmedpag = tes_mediopago.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id " _
            + vbCr + " WHERE (((tes_concidet.idconc) = " & NulosN(RstFrm("id")) & ") And ((tes_cajaori.idbcocta) = " & NulosN(LblIBcoCta) & ") And ((tes_concidet.conciliado) = -1) And ((tes_concidet.movimiento) = 2))  AND (([tes_concidet].[impdeb]+[tes_concidet].[imphab])<>0) "

    End If
            
    RST_Busq RstTmp, nSQL, xCon
    
    '--ESTABLECIENDO EL TIPO DE ORDEN SEGUN SELECCION POR USUARIO
    If OptSel1.Value = True Then RstTmp.Sort = "fchope"
    If OptSel2.Value = True Then RstTmp.Sort = "registro"
    If OptSel3.Value = True Then RstTmp.Sort = "numerodoc"
    
    '--DEFINIR LA ESTRUCUTURA DEL RST
    DEFINIR_RST_TMP RstDg4, RstTmp
    
    '--CARGAR LOS DATOS AL RST
    CARGAR_RST_TMP RstDg4, RstTmp
    Set RstTmp = Nothing
    '---------------------------------
    
    Set Dg4.DataSource = RstDg4
    
    Dg4.Columns("glosa").FooterText = "Total ==>>"
    Dg4.Columns("impdebe1").FooterText = Format(RstRegistroSumar(RstDg4, "impdebe1"), FORMAT_MONTO)
    Dg4.Columns("imphaber1").FooterText = Format(RstRegistroSumar(RstDg4, "imphaber1"), FORMAT_MONTO)
    
End Sub

Sub Conciliar()

End Sub

Private Sub Form_Load()
    '--CENTRAR EL FORMULARIO
    CentrarFrm Me
    
    QueHace = 3
    SeEjecuto = False
    
    '--ESTABLECER EL GRID POR DEFECTO LA CONSULTA
    TabOne1.CurrTab = 0
    
    CaracteresNumericos = "0123456789.-" & Chr(8)
    
    '--ESTABLECIENDO LOS FORMATOS DE LOS GRID'S
    Dg1.Columns("fchini").NumberFormat = FORMAT_DATE
    Dg1.Columns("fchfin").NumberFormat = FORMAT_DATE
    
    Dg1.Columns("impdeb").NumberFormat = FORMAT_MONTO
    Dg1.Columns("imphab").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impsal").NumberFormat = FORMAT_MONTO

    '--------
    Dg3.BatchUpdates = False
    Dg4.BatchUpdates = False
    
    Dg3.Columns("fchope1").NumberFormat = FORMAT_DATE
    Dg3.Columns("impdebe1").NumberFormat = FORMAT_MONTO
    Dg3.Columns("imphaber1").NumberFormat = FORMAT_MONTO
    
    Dg4.Columns("fchope1").NumberFormat = FORMAT_DATE
    Dg4.Columns("impdebe1").NumberFormat = FORMAT_MONTO
    Dg4.Columns("imphaber1").NumberFormat = FORMAT_MONTO

    SeEjecuto = False
    
    Fg1.Rows = 1
    Fg1.Editable = flexEDNone
    
    '--ESTABLECER EL COLOR DE FONDO DE LOS FRAMES
    Frame2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F
    Frame8.BackColor = &H8000000F
    Frame11.BackColor = &H8000000F
    Frame4.BackColor = &H8000000F
    Frame7.BackColor = &H8000000F
    '--CARGAR EL FORMATO DE LA GRILLA QUE CONTIENE EL LIBRO DIARIO
    SetearCuadricula Fg1, 8, xCon, 1, 0, False



End Sub


Private Sub Menu1_1_Click()
    Select Case mTipoGrid
        Case 3: TDB_SelDesActCheck Dg3, RstDg3, "xconc", "-1"
        Case 4: TDB_SelDesActCheck Dg4, RstDg4, "xconc", "-1"
        Case Else
    End Select
End Sub

Private Sub Menu1_2_Click()
    Select Case mTipoGrid
        Case 3: TDB_SelDesActCheck Dg3, RstDg3, "xconc", "0"
        Case 4: TDB_SelDesActCheck Dg4, RstDg4, "xconc", "0"
        Case Else
    End Select
End Sub

Private Sub menu1_4_Click()
    Select Case mTipoGrid
        Case 3: TDB_TodosDesActCheck Dg3, RstDg3, "xconc", "-1"
        Case 4: TDB_TodosDesActCheck Dg4, RstDg4, "xconc", "-1"
        Case Else
    End Select
End Sub

Private Sub Menu1_5_Click()
    Select Case mTipoGrid
        Case 3: TDB_TodosDesActCheck Dg3, RstDg3, "xconc", "0"
        Case 4: TDB_TodosDesActCheck Dg4, RstDg4, "xconc", "0"
        Case Else
    End Select
End Sub

Private Sub Menu1_7_Click()
    Select Case mTipoGrid
        Case 3: RstDg3.Filter = "": TDB_FiltroLimpiar Dg3
        Case 4: RstDg4.Filter = "": TDB_FiltroLimpiar Dg4
        Case Else
    End Select
End Sub

Private Sub Menu1_9_Click()
    pExportar
End Sub


Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        'Validamos si la cuadricula tiene datos
        If QueHace = 3 Then
            If RstFrm.RecordCount = 0 Then
                MsgBox "No existe información para visualizar", vbInformation, Me.Caption
                'Blanquea
                Exit Sub
            Else
                MuestraSegundoTab
            End If
        End If
    End If
     
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        If RstFrm.RecordCount = 0 Then
            MsgBox "No se han realizado conciliacion para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then Cancelar
    If Button.Index = 6 Then
        If Grabar = True Then
            Cancelar
            RstFrm.Requery
            Dg1.Refresh
            '--------------------------------------------------------------------------
            If RstFrm.RecordCount <> 0 Then
                RstFrm.MoveFirst
                RstFrm.Find "id=" & mIdRegistro
                If RstFrm.EOF = True Then RstFrm.MoveFirst
            End If
            '--------------------------------------------------------------------------
        End If
    End If
    
    If Button.Index = 11 Then pExportar
    If Button.Index = 12 Then pImprimir
    If Button.Index = 13 Then pConfigurar
    
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub



Private Sub TxtCuenta_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 116 Then
        CmdBusIdBanco_Click
    End If
End Sub

Private Sub TxtCuenta_Validate(Cancel As Boolean)
    If NulosC(TxtCuenta.Text) = "" Then
        LblBanco.Caption = ""
        LblMoneda.Caption = ""
        LblIBcoCta.Caption = 0
        LblIdCuentaContable.Caption = 0
        LblIdMoneda.Caption = 0
        LblCtaNombre.Caption = ""
        LblCtaNum.Caption = ""
    End If
End Sub

Private Sub TxtFchFin_LostFocus()
    If QueHace = 3 Then Exit Sub
    cmdbuscar.SetFocus
End Sub

Private Sub TxtImpExtracto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub


Private Sub pExportarRstDg()
    

    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset

    Dim xCampos(10, 3) As String
    
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Nº Reg":       xCampos(0, 1) = "registro":     xCampos(0, 2) = 0:   xCampos(0, 3) = "900"
    xCampos(1, 0) = "T.D.":         xCampos(1, 1) = "docabrev":     xCampos(1, 2) = 0:   xCampos(1, 3) = "350"
    xCampos(2, 0) = "Num. Doc":     xCampos(2, 1) = "numerodoc":    xCampos(2, 2) = 0:   xCampos(2, 3) = "1600"
    xCampos(3, 0) = "Fch.Ope":      xCampos(3, 1) = "fchope1":      xCampos(3, 2) = 1:   xCampos(3, 3) = "900"
    xCampos(4, 0) = "Medio Pago":   xCampos(4, 1) = "descmedpag":   xCampos(4, 2) = 1:   xCampos(4, 3) = "900"
    xCampos(5, 0) = "Glosa":        xCampos(5, 1) = "glosa":        xCampos(5, 2) = 0:   xCampos(5, 3) = "1200"
    xCampos(6, 0) = "M":            xCampos(6, 1) = "simbolo":      xCampos(6, 2) = 1:   xCampos(6, 3) = "500"
    xCampos(7, 0) = "T.C.":         xCampos(7, 1) = "tipcam":       xCampos(7, 2) = 2:   xCampos(7, 3) = "700"
    xCampos(8, 0) = "Imp Debe":     xCampos(8, 1) = "impdebe1":     xCampos(8, 2) = 2:   xCampos(8, 3) = "900"
    xCampos(9, 0) = "Imp Haber":    xCampos(9, 1) = "imphaber1":    xCampos(9, 2) = 2:   xCampos(9, 3) = "900"
    xCampos(10, 0) = "Conc":        xCampos(10, 1) = "xconc":  xCampos(10, 2) = 2:  xCampos(10, 3) = "1100"
    
    Select Case mTipoGrid
        Case 3:
            Set RstTmp = RstDg3.Clone
        Case 4:
            Set RstTmp = RstDg4.Clone
        Case Else
            Exit Sub
    End Select
    
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "CONCILIACION BANCARIA", "Periodo:  " & TxtFchIni.Valor & " a " & TxtFchFin.Valor, "", "Conciliacion Bancaria", RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
    
End Sub

Private Sub TxtImpExtracto_Validate(Cancel As Boolean)
    TxtImpExtracto.Text = Format(TxtImpExtracto.Text, FORMAT_MONTO)
End Sub


Private Sub pCargarLBanco()
    '===================================================================================================
    'Creado : xx/xx/09 Por: Johan Castro
    'Propósito: Muestra el libro mayor de la cuenta seleccionada
    '
    'Entradas:  Ninguna
    '
    'Resultados: Libro Mayor en pantalla
    '
    'Nota:       1.- Seleccionar Nro Cta de Banco
    '            2.- Indicar el periodo de consulta
    '            3.- Clic en boton para buscar
    '===================================================================================================

    Dim RstDet As New ADODB.Recordset
    Dim RstSal As New ADODB.Recordset
    Dim nSQLCampos As String
    Dim mCol As Long
    Dim mColCampo As Integer
    Dim nSQLAjuste As String
    Dim nSQLIdLibro As String
    
    
    Dim RstTmp2 As New ADODB.Recordset
    
    Dim A&, B&, C&
    Dim nSQL As String
    On Error GoTo error

    DoEvents
    
    Dim nSQLCuenta As String
    nSQLCuenta = ""
    '--Sentencia SQL para filtrar el codigo de la cuenta contable
    nSQLCuenta = " and con_planctas.id =" & LblIdCuentaContable.Caption & " "
    '---------
    '--ESTABLECER EL CAMPO A TOTALIZAR EN FUNCION DEL RECORDSET TMP (RstTmp2) , TANTO A SOLES Y DOLARES
    Dim CAMPO_DEBE, CAMPO_HABER, CAMPO_SALDO As String
    
    If NulosN(LblIdMoneda.Caption) = 1 Then
        CAMPO_DEBE = "impdebsol":  CAMPO_HABER = "imphabsol": CAMPO_SALDO = "impsalsol"
    Else
        CAMPO_DEBE = "impdebdol":  CAMPO_HABER = "imphabdol": CAMPO_SALDO = "impsaldol"
    End If
    '-----------------------------------------------------------------------
     
    '**********************************************************************************************
    nSQLIdLibro = "  "
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " (con_diario.ajuste in (0, " & NulosN(LblIdMoneda.Caption) & ") ) AND "
    '-----------------------------------------------
     
    Set RstTmp2 = Nothing
    '**********************************************************************************************
    nSQLCampos = fSetearCuadriculaColumna(xCon, 8)
    If nSQLCampos = "" Then Exit Sub
    nSQLCampos = "idcuenta,tipsal," & nSQLCampos
     '**********************************************************************************************
    '--MOSTRAR EL TOOLBAR PARA MOSTRAR EL INCREMENTO
    Frame6.Left = 3413
    Frame6.Top = 3780
    Frame6.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0
     '**********************************************************************************************

   '--tomar tipo de cambio del diario cuando idlib = bancos y diversos
   nSQL = "SELECT iif(con_diario.idlib=6 and tes_caja.conciliado is not null,tes_caja.conciliado,0) as conciliado,iif(conciliado=-1,'Si','No') as conciliado1, " _
            + vbCr + " con_diario.idcue AS idcuenta,con_planctas.tipsal,Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, Format([con_diario].[idmes],'00') AS mes, mae_libros.codsun AS libsun, CDbl(con_diario.numasi) AS corr, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, con_diario.rglosaope as glosaope, con_diario.rglosa AS glosaref, con_diario.rregistro AS registroref, iif(con_diario.ridtipper=0,tes_documentos.abrev,mae_documento.abrev) AS tdocdesc, con_diario.fchdoc, con_diario.rnumerodoc AS numdoc, " _
            + vbCr + " IIf([con_diario].[ridtipper]=5,[mae_bancos].[numruc],'') AS numruc, " _
            + vbCr + " IIf([con_diario].[ridtipper]=5,[mae_bancos].[descripcion],'') AS apenom , mae_documento.codsun AS tdocsun, iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo AS monope, mae_moneda_1.simbolo AS monref, "
    
    If NulosN(LblIdMoneda.Caption) = 1 Then
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
        + vbCr + " tes_documentos_1.abrev AS tdocdesc2, con_diario.rfchope2 AS fchdoc2, con_diario.rnumerodoc2 AS numdoc2,con_diario.ridtipper2, '' AS numruc2,'' AS apenom2 " _
        + vbCr + " FROM ((((((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc = tes_documentos.id) " _
        + vbCr + " LEFT JOIN mae_bancos ON con_diario.ridper = mae_bancos.id) LEFT JOIN mae_documento AS mae_documento_1 ON con_diario.rtipdoc1 = mae_documento_1.id) LEFT JOIN tes_documentos AS tes_documentos_1 ON con_diario.rtipdoc2 = tes_documentos_1.id) LEFT JOIN mae_moneda AS mae_moneda_1 ON con_diario.ridmon = mae_moneda_1.id) LEFT JOIN tes_caja ON con_diario.idmov = tes_caja.id "
        
''    If opt_fecha(0).Value = True Then
        nSQL = nSQL + vbCr + " WHERE " & nSQLAjuste & " ( con_diario.fchasi >=CDate('" + TxtFchIni.Valor + "') And con_diario.fchasi<=CDate('" + TxtFchFin.Valor + "') ) " _
            + vbCr + " AND ( con_diario.fchasi >=CDate('01/01/" + AnoTra + "') And con_diario.fchasi <= CDate('31/12/" + AnoTra + "') ) "
''    Else
''        nSQL = nSQL + vbCr + " WHERE " & nSQLAjuste & " ( con_diario.idmes >= " & mMesIni & " and con_diario.idmes <= " & mMesFin & " ) and con_diario.año = " & AnoTra & " "
''    End If
        
    nSQL = nSQL + nSQLCuenta + vbCr + " ORDER BY con_planctas.cuenta ASC "

     '**********************************************************************************************
    '--remplazando segun la moneda seleccionada
    
    If NulosN(LblIdMoneda.Caption) = 1 Then
        nSQL = Replace(nSQL, "impdebesol", "debe")
        nSQL = Replace(nSQL, "imphabersol", "haber")
        nSQL = Replace(nSQL, "impsalsol", "saldo")
    Else
        nSQL = Replace(nSQL, "impdebedol", "debe")
        nSQL = Replace(nSQL, "imphaberdol", "haber")
        nSQL = Replace(nSQL, "impsaldol", "saldo")
    End If
        
    nSQL = "Select " & nSQLCampos & _
            vbCr + " from ( " _
            + vbCr + nSQL _
            + vbCr + ") as diario "
     '**********************************************************************************************
     
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    
    RST_Busq RstTmp2, nSQL, xCon
    
    
    '--OBTENER LAS POSICIONES DE LAS COLUMNAS DEBE, HABER Y SALDO
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
    
     xSaldo = 0
     xTotal1 = 0
     xTotal2 = 0
     
     '---------------------------------------------------------------------------------------------------------------
     '---------------------------------------------------------------------------------------------------------------
     
    'OBTENER EL SALDO ANTERIOR DE LA CUENTA
     Set RstSal = Nothing
                                                 
    nSQL = "SELECT con_diario.idcue as idcuenta, con_planctas.cuenta,con_planctas.tipsal, " _
         + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS impdebsol, " _
         + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS imphabsol, " _
         + vbCr + " Sum(IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven)))))) AS impdebdol, " _
         + vbCr + " Sum(IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven)))))) As imphabdol " _
         + vbCr + " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue "
     
     nSQL = nSQL + vbCr + " WHERE ( " & nSQLIdLibro & nSQLAjuste & " con_diario.fchasi Is Null and con_diario.año = " & AnoTra & " ) Or ( " & nSQLIdLibro & nSQLAjuste & " con_diario.fchasi < CDate('" & TxtFchIni.Valor & "')) "
     
     nSQL = nSQL + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.tipsal  " _
         + vbCr + " HAVING con_diario.idcue =" & LblIdCuentaContable.Caption & " "
                     
     RST_Busq RstSal, nSQL, xCon
     
     Fg1.Rows = Fg1.Rows + 1
     
     If RstSal.RecordCount <> 0 Then
         If UCase(RstSal.Fields("tipsal") & "") = "D" Or NulosC(RstSal.Fields("tipsal")) = "" Then
             xSaldo = (NulosN(RstSal(CAMPO_DEBE)) - NulosN(RstSal(CAMPO_HABER)))
         Else
             xSaldo = (NulosN(RstSal(CAMPO_HABER)) - NulosN(RstSal(CAMPO_DEBE)))
         End If
         
         '-----------------------
         If NulosN(RstSal(CAMPO_DEBE)) - NulosN(RstSal(CAMPO_HABER)) > 0 Then
             Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(NulosN(RstSal(CAMPO_DEBE)) - NulosN(RstSal(CAMPO_HABER)), FORMAT_MONTO)
             xTotal1 = NulosN(RstSal(CAMPO_DEBE)) - NulosN(RstSal(CAMPO_HABER))
         Else
             Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(NulosN(RstSal(CAMPO_HABER)) - NulosN(RstSal(CAMPO_DEBE)), FORMAT_MONTO)
             xTotal2 = NulosN(RstSal(CAMPO_HABER)) - NulosN(RstSal(CAMPO_DEBE))
         End If
         '-----------------------
         Fg1.TextMatrix(Fg1.Rows - 1, mColSaldo) = Format(xSaldo, FORMAT_MONTO)
             
    Else
        xSaldo = 0
        Fg1.TextMatrix(Fg1.Rows - 1, mColSaldo) = "0.00"
    End If
    
    Fg1.TextMatrix(Fg1.Rows - 1, 3) = "SALDOS INICIALES =>>"
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, , True
    FORMATO_CELDA Fg1, Fg1.Rows - 1, mColSaldo, , True
     
     '---------------------------------------------------------------------------------------------------------------
     '---------------------------------------------------------------------------------------------------------------
     
    RstTmp2.Filter = adFilterNone
    
    Label5.Caption = "Procesando Cta   :  " + LblBanco.Caption
    DoEvents
    xFila = xFila + 1
    
    If RstTmp2.RecordCount <> 0 Then
     
        RstTmp2.MoveFirst
        
        '--ESTABLECIENDO EL TIPO DE ORDEN SEGUN SELECCION POR USUARIO
        If OptSel1.Value = True Then RstTmp2.Sort = "fchdoc"
        If OptSel2.Value = True Then RstTmp2.Sort = "registro"
        If OptSel3.Value = True Then RstTmp2.Sort = "numdoc"
        
        '--ESTABLECIENDO LA CANTIDAD TOTAL DE REGISTROS A ESCRIB
        ProgressBar1.Max = RstTmp2.RecordCount
        
        '--ESCRIBIR EN EL GRID
        Do While Not RstTmp2.EOF
            DoEvents
             
             ProgressBar1.Value = ProgressBar1.Value + 1
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
                         
                     Case "tc", "tipcam"
                         Fg1.TextMatrix(Fg1.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), "0.000")
                         
                     'Case "debe"
                     Case "debe", "impdebesol", "impdebedol":
                         Fg1.TextMatrix(Fg1.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), FORMAT_MONTO)
                         xTotal1 = xTotal1 + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, mCol))
                                                  
                     Case "haber", "imphabersol", "imphaberdol":
                         Fg1.TextMatrix(Fg1.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), FORMAT_MONTO)
                         xTotal2 = xTotal2 + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, mCol))

                     Case "saldo", "impsalsol", "impsaldol"
                         If UCase(RstTmp2.Fields("tipsal") & "") = "D" Or NulosC(RstTmp2.Fields("tipsal")) = "" Then
                             xSaldo = xSaldo + Format((NulosN(RstTmp2(mColDebe + 1)) - NulosN(RstTmp2(mColHaber + 1))), FORMAT_MONTO)
                         Else
                             xSaldo = xSaldo + Format((NulosN(RstTmp2(mColHaber + 1)) - NulosN(RstTmp2(mColDebe + 1))), FORMAT_MONTO)
                         End If
                         Fg1.TextMatrix(Fg1.Rows - 1, mCol) = Format(xSaldo, FORMAT_MONTO)
                         
                    Case "conciliado1"
                        Fg1.TextMatrix(Fg1.Rows - 1, mCol) = NulosC(RstTmp2.Fields(mColCampo))
                        If UCase(NulosC(RstTmp2.Fields(mColCampo))) = "SI" Then
                            GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HCEFFFE
                        Else
                            GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, RGB(244, 155, 155)
                        End If
                        
                     Case Else
                         Fg1.TextMatrix(Fg1.Rows - 1, mCol) = NulosC(RstTmp2.Fields(mColCampo))
                         
                 End Select
                 
             Next mColCampo
             
             RstTmp2.MoveNext

             xFila = xFila + 1
         Loop
     Else
     
     End If
     
     xTotal1_1 = xTotal1_1 + xTotal1
     xTotal2_1 = xTotal2_1 + xTotal2

    If xTotal1_1 <> 0 Or xTotal2_1 <> 0 Then
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, mColDebe - 1) = "Total "
        Fg1.TextMatrix(Fg1.Rows - 1, mColDebe) = Format(xTotal1_1, FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xTotal2_1, FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, mColSaldo) = Format(xSaldo, FORMAT_MONTO)
         
        ''''Fg1.TextMatrix(Fg1.Rows - 1, mColHaber) = Format(xTotal1_1, FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe - 1, , True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColDebe, , True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColHaber, , True
        FORMATO_CELDA Fg1, Fg1.Rows - 1, mColSaldo, , True
        
    End If
    
'''    '--ajustando las columnas de acuerdo a los importes
'''    Fg1.AutoSizeMode = flexAutoSizeColWidth
'''    Fg1.AutoSize mColDebe
'''    Fg1.AutoSize mColHaber
'''    Fg1.AutoSize mColSaldo
    

Salir:
    Set RstDet = Nothing:     Set RstSal = Nothing
    Frame6.Visible = False
    
    Fg1.FrozenRows = 1
    
    Exit Sub
error:
''''    Resume
    Set RstDet = Nothing:     Set RstSal = Nothing
    Frame6.Visible = False
    
    MsgBox Err.Description & vbCr & Err.Source & vbCr & Err.Number, vbCritical, xTitulo
    ''Resume
    Err.Clear
     
End Sub


Sub pConfigurar()
    '--COFIGURAR LA PRESENTACION DEL LIBRO BANCO
    Dim xForm As New SGI2_funciones.Varias
    TabOne1.CurrTab = 0
    If xForm.CambioOpcionLiro(8, xCon, 1) = True Then
        SetearCuadricula Fg1, 5, xCon, 1, 0, False
    End If
    Set xForm = Nothing
End Sub





Private Sub pExportar()
    '===================================================================================================
    'Creado : 01/05/10 Por: Johan Castro
    'Propósito: Exportar el libro Banco y registros por conciliar
    '
    'Entradas:  Ninguna
    '
    'Resultados: Libro Banco, Movimientos en libro y Mov de periodos anteriores en excel
    '
    'Nota:       1.- Libro Banco clic en boton exportar excel en toolbar
    '            2.- Movimiento en libro y periodos anteriores clic derecho sobre el grid
    '===================================================================================================

    If TabOne1.CurrTab = 0 Then
        MsgBox "Falta seleccionar el registro", vbExclamation, xTitulo
        Exit Sub
        
    ElseIf TabOne1.CurrTab = 1 Then
        If Fg1.Rows = 1 Then
            MsgBox "No hay registros para exportar", vbExclamation, xTitulo
            Exit Sub
        End If
    End If
        
    Dim nTitulo As String
    Dim nTitulo1 As String
    Dim nPeriodo As String
    
    nPeriodo = "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor
    
        
    Dim xFun As New SGI2_funciones.formularios
    
    If TabOne1.CurrTab = 0 Then
        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "DETALLE DEL LIBRO BANCOS", nPeriodo, "Expresado en " & LblMoneda.Caption, "Libro Bancos - Detalle"          ', Rst, ""
    Else
        Dim xCampos(8, 3) As String
        Dim xrst As New ADODB.Recordset
        '0::Nombre a Mostrar;
        '1::nombre de Campo del Rst;
        '2::alineacion(0::derecha, 1::centro, 2::izquierda);
        '3::ancho de columna
        '--obs: el rst puede tener mas columnas solo se consideran los campos del array
        xCampos(0, 0) = "Nº Reg":       xCampos(0, 1) = "registro":    xCampos(0, 2) = 0:   xCampos(0, 3) = "900"
        xCampos(1, 0) = "T.D.":         xCampos(1, 1) = "docabrev":    xCampos(1, 2) = 0:   xCampos(1, 3) = "350"
        xCampos(2, 0) = "Num. Doc":     xCampos(2, 1) = "numerodoc":   xCampos(2, 2) = 0:   xCampos(2, 3) = "1500"
        xCampos(3, 0) = "Fch.Ope.":     xCampos(3, 1) = "fchope1":     xCampos(3, 2) = 1:   xCampos(3, 3) = "950"
        xCampos(4, 0) = "Medio Pago":   xCampos(4, 1) = "descmedpag":  xCampos(4, 2) = 0:   xCampos(4, 3) = "1000"
        xCampos(5, 0) = "Glosa":        xCampos(5, 1) = "glosa":       xCampos(5, 2) = 0:   xCampos(5, 3) = "2000"
        xCampos(6, 0) = "Total Debe":   xCampos(6, 1) = "impdebe1":    xCampos(6, 2) = 2:   xCampos(6, 3) = "1000"
        xCampos(7, 0) = "Total Haber":  xCampos(7, 1) = "imphaber1":   xCampos(7, 2) = 2:   xCampos(7, 3) = "1000"
        xCampos(8, 0) = "Conc.":        xCampos(8, 1) = "xconc":       xCampos(8, 2) = 1:   xCampos(8, 3) = "700"
        
        
        Select Case mTipoGrid
            Case 3 '--Registos del periodo actual
                Set xrst = RstDg3.Clone
                nTitulo = "MOVIMIENTOS EN LIBROS"
            Case 4 '--Registros
                Set xrst = RstDg4.Clone
                nTitulo = "MOVIMIENTOS DE PERIODOS ANTERIORES"
            Case Else
                Set xrst = Nothing
                Exit Sub
        End Select
        xFun.VSFlexGrid_Exportar_MSExcel xCon, , nTitulo, nPeriodo, "Expresado en " & LblMoneda.Caption, "Libro Bancos - Detalle", xrst, xCampos()
        Set xrst = Nothing
        Set xFun = Nothing
    End If
    
End Sub




Private Sub pImprimir()
    Dim xMoneda As String
    Dim nPeriodo As String
    
    
    If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
        nPeriodo = "Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
    Else
        nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
    End If
    
    
    xMoneda = LblMoneda.Caption
    '----
    Dim RstTmp As New ADODB.Recordset
    Dim A As Integer
    Dim rst As New ADODB.Recordset
    RST_Busq rst, "SELECT con_formatostipodet.*,con_formatostipo.rpttitulo, con_formatostipo.rpttamdet, con_formatostipo.rpttamcab " _
        & " FROM con_formatostipodet INNER JOIN con_formatostipo ON (con_formatostipo.id = con_formatostipodet.idformatotipo) AND (con_formatostipodet.idformato = con_formatostipo.idformato) " _
        & " WHERE (((con_formatostipo.idformato)=8) AND ((con_formatostipodet.mostrar)=-1) AND ((con_formatostipo.defecto)=-1)) " _
        & " ORDER BY con_formatostipodet.orden", xCon
    
    Dim xCampos() As String
    Dim xFil, xCol As Double
    Dim xIndice As Integer
    
        
    ReDim xCampos(Fg1.Rows - 2, Fg1.Cols - 1)
    
    Dim xFila As Double
    xFila = 0
    For xFil = 1 To Fg1.Rows - 1
        For xCol = 1 To Fg1.Cols - 1
            xCampos(xFila, xCol) = Fg1.TextMatrix(xFil, xCol)
        Next xCol
        xFila = xFila + 1
    Next xFil
    
    rst.MoveFirst
    For A = 1 To rst.RecordCount
        If NulosC(xCampos(0, A)) = NulosC(rst("abrev")) Then
            If rst("imprimir") = False Then
                xCampos(0, A) = ""
            End If
        End If
        rst.MoveNext
        If rst.EOF = True Then Exit For
    Next A
    
    rst.MoveFirst
    
    Dim xfrm As New eps_librerias.Imprimir
    
    xfrm.Cabecera1 = NomEmp
    xfrm.Cabecera2 = "RUC Nº: " & NumRuc
    xfrm.Fecha = Format(Date, "dd/mm/yyyy")
    xfrm.Titulo1 = NulosC(rst("rpttitulo")) & " (Expresado en " & xMoneda & ")"
    xfrm.Titulo2 = nPeriodo & vbCr & "N° Cta: " & TxtCuenta.Text & "   " & LblBanco.Caption
    xfrm.TamañoFuente = NulosN(rst("rpttamdet"))
    xfrm.TamañoCabecera = NulosN(rst("rpttamcab"))
    xfrm.FuenteCabecera = "Courier New"
    xfrm.Posicion_Hoja = Vertical
    xfrm.Tamaño_Hoja = A_4
    xfrm.TextoConsiderar = " "
    xfrm.TextoConsiderarAncho = 1
    xfrm.ImprimirArray xCampos, rst
    Set xfrm = Nothing
    Set rst = Nothing
    
End Sub



Sub OpcionesPeriodo()
    
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, 27, mMesActivo, fCierrePeriodo, xCon
    '------------------------------------------------------------------------------------------
    
End Sub

Private Function fValidar() As Boolean
    '===================================================================================================
    'Creado : xx/xx/09 Por: Johan Castro
    'Propósito: Validar que las opciones seleccionadas por el usuario sean correctas
    '           Muestra un mensaje si falta algun dato seleccionar o ingresar
    '
    'Entradas:  Ninguna
    '
    'Resultados: Alerta si puede continuar con la consulta
    '
    '===================================================================================================
    '--posicionar en pestaña de libro banco
    TabOne2.CurrTab = 0
    fValidar = True
    If Trim(Me.TxtCuenta.Text) = "" Then
        MsgBox "Seleccione Cuenta", vbInformation, Me.Caption
        TxtCuenta.SetFocus
        fValidar = False
        Exit Function
    End If
    
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "Falta especificar la fecha de inicio", vbExclamation, xTitulo
        TxtFchIni.SetFocus
        fValidar = False
        Exit Function
    End If
    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "Falta especificar la fecha de final", vbExclamation, xTitulo
        TxtFchFin.SetFocus
        fValidar = False
        Exit Function
    End If
    
    If CDate(TxtFchFin.Valor) < CDate(TxtFchIni.Valor) Then
        MsgBox "La fecha de inicio es mayor que la fecha final" & vbCr & "Modifique las fechas", vbExclamation, xTitulo
        TxtFchIni.SetFocus
        fValidar = False
        Exit Function
    End If
    
    If Month(CDate(TxtFchIni.Valor)) <> Month(CDate(TxtFchFin.Valor)) Then
        MsgBox "La conciliacion bancaria debe ser mensual", vbInformation, Me.Caption
        fValidar = False
        Exit Function
    End If
    
End Function

Sub VerConciliacion()
    '===================================================================================================
    'Creado : 29/05/10 Por: Johan Castro
    'Propósito: Imprimir la conciliacion y presentar en pdf
    '
    'Entradas:  Ninguna
    '
    'Resultados: Archivo en pdf
    '
    '===================================================================================================

    Set oPDF = New cPDF
    
    Dim xFila As Long
    Dim vdblSaldoInicial As Double
    Dim vdblMovAbono As Double
    Dim vdblMovCargo As Double
    Dim vdblMovSaldo As Double
    Dim vdblSaldoFinal As Double
    
    Dim vdblTotCargos As Double '--Cargos Registrados y no Registrados
    Dim vdblTotCargos1 As Double '--Operaciones no conciliadas
    Dim vdblTotCargos2 As Double '--Operaciones no registradas en seven
    
    Dim vdblTotAbonos As Double '--Abonos Registrados y no Registrados
    Dim vdblTotAbonos1 As Double '--Operaciones no conciliadas
    Dim vdblTotAbonos2 As Double '--Operaciones no registradas en seven
        
    Dim vDblTotLibroConc As Double '--Saldo segun libros + Total Cargos Pendientes + Total Abonos Pendientes
        
    Dim vintFila As Integer '--Recorre cada fila de la grilla de mov no considerados
    
    '--definir rst temporal para unir operaciones pendientes de conciliacion
    '--periodos anteriores y periodo actual
    Dim xRstTMP As New ADODB.Recordset
    
    '--Inicializar contador de paginas
    xNumPag = 0
    
    If oPDF.PDFCreate(App.Path & "\000001.pdf") = True Then

        oPDF.Fonts.Add "Tit", Times_BoldItalic, WinAnsiEncoding
        oPDF.Fonts.Add "Head", Times_Italic, WinAnsiEncoding
        oPDF.Fonts.Add "Cont", Courier, WinAnsiEncoding
        oPDF.Fonts.Add "CB", Courier_Bold, WinAnsiEncoding
        oPDF.Fonts.Add "Time", Times_Roman, WinAnsiEncoding

        
        CrearCabecera
                
        xFila = 110
        
        vdblSaldoInicial = Fg1.TextMatrix(Fg1.FixedRows, mColSaldo)
        vdblMovAbono = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, mColDebe)) - NulosN(Fg1.TextMatrix(Fg1.FixedRows, mColDebe))
        vdblMovCargo = (-1) * NulosN(Fg1.TextMatrix(Fg1.Rows - 1, mColHaber)) + NulosN(Fg1.TextMatrix(Fg1.FixedRows, mColHaber))
        vdblMovSaldo = vdblMovAbono + vdblMovCargo
        
        oPDF.WTextBox xFila + 0, 30, 9, 300, "SALDOS SEGUN LIBRO BANCO AL " & CDate(CDate(TxtFchIni.Valor) - 1), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila + 9, 30, 9, 300, "(+)MAS TOTAL INGRESOS", "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila + 18, 30, 9, 300, "(-)MENOS TOTAL EGRESOS", "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila + 27, 30, 9, 300, "TOTAL INGRESOS Y EGRESOS DEL MES", "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila + 36, 30, 9, 300, "SALDO SEGUN LIBRO BANCO AL " & CDate(TxtFchFin.Valor), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
        
        '--colocar los datos
        oPDF.WTextBox xFila + 0, 440, 9, 80, Format(vdblSaldoInicial, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila + 9, 350, 9, 80, Format(vdblMovAbono, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila + 18, 350, 9, 80, Format(vdblMovCargo, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
        oPDF.WRectangle xFila + 27, 350, 0, 80, 0, vbBlack
        oPDF.WTextBox xFila + 27, 440, 9, 80, Format(vdblMovSaldo, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
        oPDF.WRectangle xFila + 36, 440, 0, 80, 0, vbBlack
        oPDF.WTextBox xFila + 36, 440, 9, 80, Format(vdblSaldoInicial + vdblMovSaldo, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
                
        xFila = xFila + 50
        '------------------------------------------------------------------------------------------------
        xFila = 165
        '--imprimir las cabeceras
        oPDF.WTextBox xFila, 30, 28, 400, "MAS (+) DOCUMENTOS CONTABILIZADOS PENDIENTES DE COBRO EN BANCO (EGRESOS)", "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
        
        CrearSubCabecera xFila
        '------------------------------------------------------------------------------------------------
        '--mostrar datos de cargos pendientes
        '--
        DEFINIR_RST_TMP xRstTMP, RstDg4
        '--Quitar filtro
        RstDg4.Filter = ""
        RstDg3.Filter = ""
        '--Limpiar los filtros del grid
        TDB_FiltroLimpiar Dg3
        TDB_FiltroLimpiar Dg4
        '--Filtrar los registros de cargos
        RstDg4.Filter = "xconc=0 and tipmov=2 and ImpHaber1<>'0'" '--Otros periodos no conciliados
        RstDg3.Filter = "xconc=0 and tipmov=2 and ImpHaber1<>'0'" '--Periodo actual pendientes de conciliar
        '--aplicando orden
        RstDg4.Sort = "fchope"
        RstDg3.Sort = "fchope"
        '--Agregando datos al rst temporal
        CARGAR_RST_TMP xRstTMP, RstDg4
        CARGAR_RST_TMP xRstTMP, RstDg3
        '--Quitar filtro
        RstDg4.Filter = ""
        RstDg3.Filter = ""
        '--Verificando si hay registros por conciliar
        If xRstTMP.RecordCount <> 0 Then
            xRstTMP.MoveFirst
            xFila = xFila + 2
            xFila = LeerFila(xFila)
            Do While Not xRstTMP.EOF
                oPDF.WTextBox xFila, 30, 6, 40, xRstTMP("registro"), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 70, 6, 210, Left$(NulosC(xRstTMP("glosa")), 48), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack, True
                oPDF.WTextBox xFila, 280, 6, 40, Format(NulosC(xRstTMP("fchope")), FORMAT_DATE), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 310, 6, 25, NulosC(xRstTMP("docabrev")), "Cont", 7, hCenter, hLeft, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 335, 6, 60, NulosC(xRstTMP("numerodoc")), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 395, 6, 50, Format(xRstTMP("ImpHaber1"), FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
                xRstTMP.MoveNext
                If xRstTMP.EOF = False Then
                    xFila = xFila + 9
                    xFila = LeerFila(xFila)
                End If
            Loop
        End If
        '--totalizando
        vdblTotCargos1 = RstRegistroSumar(xRstTMP, "ImpHaber1")
        
        xFila = xFila + 10
        xFila = LeerFila(xFila)
        oPDF.WRectangle xFila, 30, 0, 490, 0, vbBlack
        
        xFila = xFila + 2
        xFila = LeerFila(xFila)
        oPDF.WTextBox xFila, 290, 9, 80, "TOTAL CONTABILIZADO", "Cont", 7, hRight, hLeft, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila, 365, 9, 80, Format(vdblTotCargos1, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
        '------------------------------------------------------------------------------------------------
        '--imprimir las cabeceras
        xFila = xFila + 15
        xFila = LeerFila(xFila)
        oPDF.WTextBox xFila, 30, 28, 400, "MAS (+) DOCUMENTOS AGREGADOS PENDIENTES DE COBRO EN BANCO (EGRESOS)", "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
        
        CrearSubCabecera xFila
        '------------------------------------------------------------------------------------------------
        '--Totalizando
        vdblTotCargos2 = NulosN(TxtTotHab.Text)
        
        xFila = xFila + 2
        xFila = LeerFila(xFila)
        For vintFila = Fg3.FixedRows To Fg3.Rows - 1
            If NulosN(Fg3.TextMatrix(vintFila, 5)) <> 0 Then
                oPDF.WTextBox xFila, 30, 6, 40, "", "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 70, 6, 210, Fg3.TextMatrix(vintFila, 1), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack, True
                oPDF.WTextBox xFila, 280, 6, 40, Format(Fg3.TextMatrix(vintFila, 2), FORMAT_DATE), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 310, 6, 25, "", "Cont", 7, hCenter, hLeft, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 335, 6, 60, Fg3.TextMatrix(vintFila, 3), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 395, 6, 50, Format(Fg3.TextMatrix(vintFila, 5), FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
                xFila = xFila + 9
                xFila = LeerFila(xFila)
            End If
        Next
        
        If vdblTotCargos2 = 0 Then
            xFila = xFila + 10 '--Muestra linea en blanco
        Else
            xFila = xFila + 1
        End If
        xFila = LeerFila(xFila)
        oPDF.WRectangle xFila, 30, 0, 490, 0, vbBlack
        
        xFila = xFila + 2
        xFila = LeerFila(xFila)
        oPDF.WTextBox xFila, 290, 9, 80, "TOTAL AGREGADO", "Cont", 7, hRight, hLeft, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila, 365, 9, 80, Format(vdblTotCargos2, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
        
        '------------------------------------------------------------------------------------------------
        vdblTotCargos = vdblTotCargos1 + vdblTotCargos2
        xFila = xFila + 15
        xFila = LeerFila(xFila)
        oPDF.WTextBox xFila, 335, 9, 80, "TOTAL EGRESOS", "Cont", 7, hRight, hLeft, RGB(0, 0, 128), 0, vbBlack
        oPDF.WRectangle xFila, 440, 0, 80, 0, vbBlack
        oPDF.WTextBox xFila, 440, 9, 80, Format(vdblTotCargos, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
        
        
        '------------------------------------------------------------------------------------------------
        '----- INGRESOS
        '------------------------------------------------------------------------------------------------
        xFila = xFila + 30
        xFila = LeerFila(xFila)
        '--imprimir las cabeceras
        oPDF.WTextBox xFila, 30, 28, 400, "MENOS (-) DOCUMENTOS CONTABILIZADOS PENDIENTES DE PAGO EN BANCO (INGRESOS)", "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
        
        CrearSubCabecera xFila
        '------------------------------------------------------------------------------------------------
        '--Limpiar rst temporal de Cargos
        Set xRstTMP = Nothing
        
        DEFINIR_RST_TMP xRstTMP, RstDg4
        '--Filtrar los registros de abonos
        RstDg4.Filter = "xconc=0 and tipmov=1 and ImpHaber1<>'0'" '--Otros periodos no conciliados
        RstDg3.Filter = "xconc=0 and tipmov=1 and ImpHaber1<>'0'" '--Periodo actual pendientes de conciliar
        '--aplicando orden
        RstDg4.Sort = "fchope"
        RstDg3.Sort = "fchope"
        '--Agregando datos al rst temporal
        CARGAR_RST_TMP xRstTMP, RstDg4
        CARGAR_RST_TMP xRstTMP, RstDg3
        '--Quitar filtro
        RstDg4.Filter = ""
        RstDg3.Filter = ""
        '--Verificando si hay registros por conciliar
        If xRstTMP.RecordCount <> 0 Then
            xRstTMP.MoveFirst
            xFila = xFila + 2
            xFila = LeerFila(xFila)
            Do While Not xRstTMP.EOF
                oPDF.WTextBox xFila, 30, 6, 40, NulosC(xRstTMP("registro")), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 70, 6, 210, Left$(NulosC(xRstTMP("glosa")), 48), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack, True
                oPDF.WTextBox xFila, 280, 6, 40, Format(NulosC(xRstTMP("fchope")), FORMAT_DATE), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 310, 6, 25, NulosC(xRstTMP("docabrev")), "Cont", 7, hCenter, hLeft, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 335, 6, 60, NulosC(xRstTMP("numerodoc")), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 395, 6, 50, Format(xRstTMP("ImpDebe1"), FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
                xRstTMP.MoveNext
                If xRstTMP.EOF = False Then
                    xFila = xFila + 9
                    xFila = LeerFila(xFila)
                End If
            Loop
        End If
        '--totalizando
        vdblTotAbonos1 = RstRegistroSumar(xRstTMP, "ImpDebe1")
        
        xFila = xFila + 10
        xFila = LeerFila(xFila)
        oPDF.WRectangle xFila, 30, 0, 490, 0, vbBlack
        
        xFila = xFila + 2
        xFila = LeerFila(xFila)
        oPDF.WTextBox xFila, 290, 9, 80, "TOTAL CONTABILIZADO", "Cont", 7, hRight, hLeft, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila, 365, 9, 80, Format(vdblTotAbonos1, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
        '------------------------------------------------------------------------------------------------
        '--imprimir las cabeceras
        xFila = xFila + 15
        xFila = LeerFila(xFila)
        oPDF.WTextBox xFila, 30, 28, 400, "MENOS (-) DOCUMENTOS AGREGADOS PENDIENTES DE PAGO EN BANCO (INGRESOS)", "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
        
        CrearSubCabecera xFila
        '------------------------------------------------------------------------------------------------
        '--Totalizando
        vdblTotAbonos2 = NulosN(TxtTotDeb.Text)
        
        xFila = xFila + 2
        xFila = LeerFila(xFila)
        For vintFila = 1 To Fg3.Rows - 1
            If NulosN(Fg3.TextMatrix(vintFila, 4)) <> 0 Then
                oPDF.WTextBox xFila, 30, 6, 40, "", "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 70, 6, 210, Fg3.TextMatrix(vintFila, 1), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack, True
                oPDF.WTextBox xFila, 280, 6, 40, Format(Fg3.TextMatrix(vintFila, 2), FORMAT_DATE), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 310, 6, 25, "", "Cont", 7, hCenter, hLeft, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 335, 6, 60, Fg3.TextMatrix(vintFila, 3), "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
                oPDF.WTextBox xFila, 395, 6, 50, Format(Fg3.TextMatrix(vintFila, 4), FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
                xFila = xFila + 9
                xFila = LeerFila(xFila)
            End If
        Next
        
        If vdblTotAbonos2 = 0 Then
            xFila = xFila + 10 '--Muestra linea en blanco
        Else
            xFila = xFila + 1
        End If
        xFila = LeerFila(xFila)
        oPDF.WRectangle xFila, 30, 0, 490, 0, vbBlack
        
        xFila = xFila + 2
        xFila = LeerFila(xFila)
        oPDF.WTextBox xFila, 290, 9, 80, "TOTAL AGREGADO", "Cont", 7, hRight, hLeft, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila, 365, 9, 80, Format(vdblTotAbonos2, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
        
        '------------------------------------------------------------------------------------------------
        vdblTotAbonos = (-1) * (vdblTotAbonos1 + vdblTotAbonos2)
        xFila = xFila + 15
        xFila = LeerFila(xFila)
        oPDF.WTextBox xFila, 335, 9, 80, "TOTAL INGRESOS", "Cont", 7, hRight, hLeft, RGB(0, 0, 128), 0, vbBlack
        oPDF.WRectangle xFila, 440, 0, 80, 0, vbBlack
        oPDF.WTextBox xFila, 440, 9, 80, Format(vdblTotAbonos, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
                
        '------------------------------------------------------------------------------------------------
        vDblTotLibroConc = vdblSaldoInicial + vdblMovSaldo + vdblTotCargos + vdblTotAbonos
        xFila = xFila + 20
        xFila = LeerFila(xFila)
        oPDF.WTextBox xFila, 30, 9, 200, "TOTAL LIBRO BANCO CONCILIADO", "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila, 440, 9, 80, Format(vDblTotLibroConc, FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
        
        xFila = xFila + 12
        xFila = LeerFila(xFila)
        oPDF.WTextBox xFila, 30, 9, 200, "TOTAL SALDO SEGUN EXTRACTO", "Cont", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila, 440, 9, 80, Format(NulosN(TxtImpExtracto.Text), FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
                
        xFila = xFila + 12
        xFila = LeerFila(xFila)
        oPDF.WRectangle xFila, 440, 0, 80, 0, vbBlack
        oPDF.WTextBox xFila, 335, 9, 80, "DIFERENCIAS", "Cont", 7, hRight, hLeft, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila, 440, 9, 80, Format(Format(NulosN(TxtImpExtracto.Text) - vDblTotLibroConc), FORMAT_MONTO), "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
                   
        xFila = xFila + 80
        xFila = LeerFila(xFila)
        
        '--dejar espacios en blanco para la impresion de las firmas
        If xFila < 130 Then xFila = 150
        
        oPDF.WRectangle xFila, 100, 0, 130, 0, vbBlack
        oPDF.WRectangle xFila, 290, 0, 130, 0, vbBlack
        oPDF.WTextBox xFila, 105, 9, 80, "REALIZADO POR", "Cont", 7, hRight, hLeft, RGB(0, 0, 128), 0, vbBlack
        oPDF.WTextBox xFila, 295, 9, 80, "REVISADO POR", "Cont", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
                
        oPDF.PDFClose
        oPDF.Show
        Set oPDF = Nothing

'        If Opcion = 1 Then
'            Shell ("rundll32.exe url.dll,FileProtocolHandler " & Trim(App.Path) & ("\000001.pdf")), vbMaximizedFocus
'        End If
    Else
        Set oPDF = Nothing
        MsgBox "No se Puede Mostrar Documento  000001.pdf, posiblemente el archivo ya se encuentra abierto", vbCritical, "Error"
    End If

End Sub

Function LeerFila(xFila As Long) As Integer
    '===================================================================================================
    'Creado : 31/05/10 Por: Johan Castro
    'Propósito: Verificar el salto de página
    '
    'Entradas:  xFila=fila actual de la hoja donde se imprimen los datos
    '
    'Resultados: Fila validada para imprimir los datos
    '            Si hay saldo de pagina la fila se reinicia
    '
    '===================================================================================================
    
    If xFila >= 790 Then
        CrearCabecera
        LeerFila = 110 '--reinicio de fila
    Else
        LeerFila = xFila
    End If
End Function

Sub CrearCabecera()
    '===================================================================================================
    'Creado : 30/05/10 Por: Johan Castro
    'Propósito: Crear nueva hoja e Imprimir el encabezado
    '
    'Entradas:  Ninguna
    '
    'Resultados: Nueva hoja con la cabecera impresa
    '
    '===================================================================================================

    Dim xFila As Long
    oPDF.NewPage A4_Vertical ', 525, 675
    xNumPag = xNumPag + 1

    xFila = 19

    oPDF.WTextBox xFila, 450, 5, 30, "PAGINA", "Cont", 5, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 4, 450, 5, 30, "FECHA", "Cont", 5, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack

    oPDF.WTextBox xFila, 480, 5, 50, ": " & Format(xNumPag, "000"), "Cont", 5, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 4, 480, 5, 50, ": " & Format(Date, FORMAT_DATE), "Cont", 5, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack

    xFila = 20
    oPDF.WTextBox xFila, 10, 28, 565, "CONCILIACION BANCARIA", "Tit", 12, hCenter, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 10, 10, 28, 565, "Del " & CDate(TxtFchIni.Valor) & " Al " & CDate(TxtFchFin.Valor), "Tit", 6, hCenter, vMiddle, RGB(0, 0, 128), 0, vbBlack
        
    xFila = xFila + 30
    oPDF.WTextBox xFila, 30, 9, 150, "EMPRESA", "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila, 180, 9, 200, ":  " & NomEmp, "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 7, 30, 9, 150, "RUC", "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 7, 180, 9, 200, ":  " & NumRuc, "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 14, 30, 9, 150, "ENTIDAD BANCARIA", "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 14, 180, 9, 350, ":  " & UCase(LblBanco.Caption), "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 21, 30, 9, 150, "NRO CTA CTE", "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 21, 180, 9, 350, ":  " & TxtCuenta.Text, "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 28, 30, 9, 150, "NRO CTA CONTABLE", "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 28, 180, 9, 350, ":  " & LblCtaNum.Caption, "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 35, 30, 9, 150, "NOMBRE DE CUENTA CONTABLE", "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 35, 180, 9, 350, ":  " & LblCtaNombre.Caption, "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 42, 30, 9, 150, "MONEDA", "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila + 42, 180, 9, 350, ":  " & LblMoneda.Caption, "Cont", 6, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack

    xFila = xFila + 54
   
    oPDF.WRectangle xFila, 30, 1, 490, 0, vbBlack
    
    xFila = xFila + 5

    oPDF.LineStroke
    
End Sub

Private Sub CrearSubCabecera(xFila As Long)
    '===================================================================================================
    'Creado : 01/06/10 Por: Johan Castro
    'Propósito: Imprimir cabecera de subgrupos
    '
    'Entradas:  xFila=Posicion actual de la fila
    '
    'Resultados: Sub grupo impreso
    '
    '===================================================================================================

    '--imprimir linea
    xFila = xFila + 22
    oPDF.WRectangle xFila, 30, 0, 490, 0, vbBlack
    
    xFila = xFila + 2
    oPDF.WTextBox xFila, 30, 9, 40, "Nro. Reg.", "Head", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila, 70, 9, 210, "Glosa", "Head", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila, 280, 9, 40, "Fch. Ope.", "Head", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila, 310, 9, 25, "T.D.", "Head", 7, hCenter, hLeft, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila, 335, 9, 60, "Nro. Documento", "Head", 7, hLeft, vMiddle, RGB(0, 0, 128), 0, vbBlack
    oPDF.WTextBox xFila, 395, 9, 50, "Importe", "Head", 7, hRight, vMiddle, RGB(0, 0, 128), 0, vbBlack
    
    '--imprimir linea
    xFila = xFila + 10
    oPDF.WRectangle xFila, 30, 0, 490, 0, vbBlack
    
End Sub
