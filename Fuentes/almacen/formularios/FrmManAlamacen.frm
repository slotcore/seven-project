VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmManAlamacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacen - Inventario"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManAlamacen.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   26
      Top             =   375
      Width           =   11880
      _cx             =   20955
      _cy             =   12726
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
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   8388608
      Caption         =   "  &Consulta  |   &Detalle  "
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   12525
         TabIndex        =   30
         Top             =   375
         Width           =   11790
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6390
            Left            =   15
            TabIndex        =   33
            Top             =   360
            Width           =   11760
            _cx             =   20743
            _cy             =   11271
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
            FrontTabColor   =   12632256
            BackTabColor    =   -2147483629
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "  Datos Producto | Datos Contables "
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
            Begin VB.Frame Frame4 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   6360
               Left            =   12690
               TabIndex        =   59
               Top             =   15
               Width           =   11415
               Begin VB.CommandButton CmdNetoDomic 
                  Height          =   240
                  Left            =   2595
                  Picture         =   "FrmManAlamacen.frx":2B10
                  Style           =   1  'Graphical
                  TabIndex        =   108
                  Top             =   2160
                  Width           =   240
               End
               Begin VB.TextBox TxtTotPor 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   7440
                  Locked          =   -1  'True
                  TabIndex        =   91
                  TabStop         =   0   'False
                  Text            =   "TxtTotPor"
                  Top             =   5550
                  Width           =   900
               End
               Begin VB.Frame Frame5 
                  Height          =   1560
                  Left            =   8730
                  TabIndex        =   88
                  Top             =   3975
                  Width           =   2055
                  Begin VB.CommandButton CmdDelCenCos 
                     Caption         =   "&Eliminar Cento Costo"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   300
                     TabIndex        =   90
                     Top             =   825
                     Width           =   1470
                  End
                  Begin VB.CommandButton CmdAddCenCos 
                     Caption         =   "&Agregar Cento Costo"
                     Enabled         =   0   'False
                     Height          =   495
                     Left            =   300
                     TabIndex        =   89
                     Top             =   330
                     Width           =   1470
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   1470
                  Left            =   240
                  TabIndex        =   86
                  Top             =   4065
                  Width           =   8430
                  _cx             =   14870
                  _cy             =   2593
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
                  Rows            =   1
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManAlamacen.frx":2C42
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
               Begin VB.CommandButton CmdBusTipVen 
                  Height          =   240
                  Left            =   2595
                  Picture         =   "FrmManAlamacen.frx":2CDC
                  Style           =   1  'Graphical
                  TabIndex        =   83
                  Top             =   3495
                  Width           =   240
               End
               Begin VB.CommandButton CmdBusTipCom 
                  Height          =   240
                  Left            =   2595
                  Picture         =   "FrmManAlamacen.frx":2E0E
                  Style           =   1  'Graphical
                  TabIndex        =   80
                  Top             =   3180
                  Width           =   240
               End
               Begin VB.CommandButton CmdBusSel 
                  Height          =   240
                  Left            =   2595
                  Picture         =   "FrmManAlamacen.frx":2F40
                  Style           =   1  'Graphical
                  TabIndex        =   77
                  Top             =   2715
                  Width           =   240
               End
               Begin VB.CommandButton CmdIdRet 
                  Height          =   240
                  Left            =   2595
                  Picture         =   "FrmManAlamacen.frx":3072
                  Style           =   1  'Graphical
                  TabIndex        =   68
                  Top             =   1185
                  Width           =   240
               End
               Begin VB.TextBox TxtidRet 
                  Height          =   300
                  Left            =   1950
                  Locked          =   -1  'True
                  MaxLength       =   2
                  TabIndex        =   20
                  Text            =   "TxtidRet"
                  Top             =   1155
                  Width           =   915
               End
               Begin VB.CommandButton CmdIdPer 
                  Height          =   240
                  Left            =   2595
                  Picture         =   "FrmManAlamacen.frx":31A4
                  Style           =   1  'Graphical
                  TabIndex        =   67
                  Top             =   1815
                  Width           =   240
               End
               Begin VB.TextBox TxtIdPer 
                  Height          =   300
                  Left            =   1950
                  Locked          =   -1  'True
                  MaxLength       =   2
                  TabIndex        =   22
                  Text            =   "TxtIdPer"
                  Top             =   1785
                  Width           =   915
               End
               Begin VB.CommandButton CmdIdDet 
                  Height          =   240
                  Left            =   2595
                  Picture         =   "FrmManAlamacen.frx":32D6
                  Style           =   1  'Graphical
                  TabIndex        =   66
                  Top             =   1500
                  Width           =   240
               End
               Begin VB.TextBox TxtIdDet 
                  Height          =   300
                  Left            =   1950
                  Locked          =   -1  'True
                  MaxLength       =   2
                  TabIndex        =   21
                  Text            =   "TxtIdDet"
                  Top             =   1470
                  Width           =   915
               End
               Begin VB.CommandButton CmdBusCtaCom 
                  Height          =   240
                  Left            =   3135
                  Picture         =   "FrmManAlamacen.frx":3408
                  Style           =   1  'Graphical
                  TabIndex        =   61
                  Top             =   405
                  Width           =   240
               End
               Begin VB.TextBox TxtCtaCom 
                  Height          =   300
                  Left            =   1950
                  Locked          =   -1  'True
                  MaxLength       =   14
                  TabIndex        =   18
                  Text            =   "TxtCtaCom"
                  Top             =   375
                  Width           =   1455
               End
               Begin VB.CommandButton CmdBusCtaVen 
                  Height          =   240
                  Left            =   3135
                  Picture         =   "FrmManAlamacen.frx":353A
                  Style           =   1  'Graphical
                  TabIndex        =   60
                  Top             =   720
                  Width           =   240
               End
               Begin VB.TextBox TxtCtaVen 
                  Height          =   300
                  Left            =   1950
                  Locked          =   -1  'True
                  MaxLength       =   14
                  TabIndex        =   19
                  Text            =   "TxtCtaVen"
                  Top             =   690
                  Width           =   1455
               End
               Begin VB.TextBox TxtIdSelectivo 
                  Height          =   300
                  Left            =   1950
                  Locked          =   -1  'True
                  MaxLength       =   2
                  TabIndex        =   23
                  Text            =   "TxtIdSelectivo"
                  Top             =   2685
                  Width           =   915
               End
               Begin VB.TextBox TxtIdTipCom 
                  Height          =   300
                  Left            =   1950
                  Locked          =   -1  'True
                  MaxLength       =   2
                  TabIndex        =   24
                  Text            =   "TxtIdTipCom"
                  Top             =   3150
                  Width           =   915
               End
               Begin VB.TextBox TxtIdTipVen 
                  Height          =   300
                  Left            =   1950
                  Locked          =   -1  'True
                  MaxLength       =   2
                  TabIndex        =   25
                  Text            =   "TxtIdTipVen"
                  Top             =   3465
                  Width           =   915
               End
               Begin VB.TextBox TxtIdNetoDomic 
                  Height          =   300
                  Left            =   1950
                  Locked          =   -1  'True
                  MaxLength       =   2
                  TabIndex        =   109
                  Text            =   "TxtIdNetoDomic"
                  Top             =   2115
                  Width           =   915
               End
               Begin VB.Label LblNetoDomic 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblNetoDomic"
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
                  Left            =   2910
                  TabIndex        =   111
                  Top             =   2115
                  Width           =   5760
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Neto Domiciliado"
                  Height          =   195
                  Index           =   24
                  Left            =   225
                  TabIndex        =   110
                  Top             =   2145
                  Width           =   1200
               End
               Begin VB.Label Label1 
                  Caption         =   "Total ==>"
                  Height          =   195
                  Left            =   6450
                  TabIndex        =   92
                  Top             =   5595
                  Width           =   870
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Centro de Costos"
                  Height          =   195
                  Index           =   20
                  Left            =   240
                  TabIndex        =   87
                  Top             =   3810
                  Width           =   1215
               End
               Begin VB.Label LblIdTipVen 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblIdTipVen"
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
                  Left            =   2910
                  TabIndex        =   85
                  Top             =   3465
                  Width           =   5760
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "I.G.V. Venta"
                  Height          =   195
                  Index           =   19
                  Left            =   240
                  TabIndex        =   84
                  Top             =   3495
                  Width           =   870
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "I.G.V. Compra"
                  Height          =   195
                  Index           =   18
                  Left            =   240
                  TabIndex        =   82
                  Top             =   3180
                  Width           =   990
               End
               Begin VB.Label LblIdTipCom 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblIdTipCom"
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
                  Left            =   2910
                  TabIndex        =   81
                  Top             =   3150
                  Width           =   5760
               End
               Begin VB.Label LblSelectivo 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblSelectivo"
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
                  Left            =   2910
                  TabIndex        =   79
                  Top             =   2670
                  Width           =   5760
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "I.S.C."
                  Height          =   195
                  Index           =   17
                  Left            =   240
                  TabIndex        =   78
                  Top             =   2715
                  Width           =   390
               End
               Begin VB.Label LbIdCuentaVen 
                  AutoSize        =   -1  'True
                  Caption         =   "LbIdCuentaVen"
                  Height          =   195
                  Left            =   9390
                  TabIndex        =   76
                  Top             =   750
                  Visible         =   0   'False
                  Width           =   1110
               End
               Begin VB.Label LbIdCuentaCom 
                  AutoSize        =   -1  'True
                  Caption         =   "LbIdCuentaCom"
                  Height          =   195
                  Left            =   9390
                  TabIndex        =   75
                  Top             =   420
                  Visible         =   0   'False
                  Width           =   1140
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Retención"
                  Height          =   195
                  Index           =   15
                  Left            =   240
                  TabIndex        =   74
                  Top             =   1185
                  Width           =   735
               End
               Begin VB.Label LblRetencion 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblRetencion"
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
                  Left            =   2910
                  TabIndex        =   73
                  Top             =   1155
                  Width           =   5760
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Percepción"
                  Height          =   195
                  Index           =   11
                  Left            =   240
                  TabIndex        =   72
                  Top             =   1815
                  Width           =   810
               End
               Begin VB.Label LblPercepcion 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblPercepcion"
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
                  Left            =   2910
                  TabIndex        =   71
                  Top             =   1785
                  Width           =   5760
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Detracción"
                  Height          =   195
                  Index           =   10
                  Left            =   240
                  TabIndex        =   70
                  Top             =   1500
                  Width           =   780
               End
               Begin VB.Label LblDetraccion 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblDetraccion"
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
                  Left            =   2910
                  TabIndex        =   69
                  Top             =   1470
                  Width           =   5760
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Cta. Contable Compras"
                  Height          =   195
                  Index           =   9
                  Left            =   240
                  TabIndex        =   65
                  Top             =   405
                  Width           =   1620
               End
               Begin VB.Label LblNomCtaCom 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblNomCtaCom"
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
                  Left            =   3450
                  TabIndex        =   64
                  Top             =   375
                  Width           =   5760
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Cta. Contable Ventas"
                  Height          =   195
                  Index           =   8
                  Left            =   240
                  TabIndex        =   63
                  Top             =   720
                  Width           =   1500
               End
               Begin VB.Label LblNomCtaVen 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblNomCtaVen"
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
                  Left            =   3450
                  TabIndex        =   62
                  Top             =   690
                  Width           =   5760
               End
            End
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   6360
               Left            =   330
               TabIndex        =   34
               Top             =   15
               Width           =   11415
               Begin VB.TextBox TxtPrecioIni 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   25
                  TabIndex        =   15
                  Text            =   "TxtPrecioIni"
                  Top             =   5670
                  Width           =   1470
               End
               Begin VB.CommandButton CmdBusMatPri 
                  Height          =   240
                  Left            =   2040
                  Picture         =   "FrmManAlamacen.frx":366C
                  Style           =   1  'Graphical
                  TabIndex        =   112
                  Top             =   6015
                  Width           =   240
               End
               Begin VB.CommandButton CmdBusTipMovimiento 
                  Height          =   240
                  Left            =   2040
                  Picture         =   "FrmManAlamacen.frx":379E
                  Style           =   1  'Graphical
                  TabIndex        =   38
                  Top             =   2790
                  Width           =   240
               End
               Begin VB.TextBox TxtStockMax 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   13
                  Text            =   "TxtStockMax"
                  Top             =   5325
                  Width           =   1470
               End
               Begin MSComDlg.CommonDialog CommonDialog1 
                  Left            =   10590
                  Top             =   600
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin VB.CommandButton CmdDelFoto 
                  Caption         =   "Eliminar Imagen"
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   9150
                  TabIndex        =   104
                  Top             =   5970
                  Width           =   1335
               End
               Begin VB.CommandButton CmdAddFoto 
                  Caption         =   "Agregar Imagen"
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   7755
                  TabIndex        =   103
                  Top             =   5970
                  Width           =   1335
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg2 
                  Height          =   1245
                  Left            =   6795
                  TabIndex        =   102
                  Top             =   4650
                  Width           =   4590
                  _cx             =   8096
                  _cy             =   2196
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
                  BackColorSel    =   64
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManAlamacen.frx":38D0
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
               Begin VB.Frame Frame6 
                  Height          =   4545
                  Left            =   6810
                  TabIndex        =   101
                  Top             =   45
                  Width           =   4575
                  Begin VB.Image Image1 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   4290
                     Left            =   75
                     Stretch         =   -1  'True
                     Top             =   165
                     Width           =   4395
                  End
               End
               Begin VB.TextBox TxtDescCaracteristica 
                  Height          =   1095
                  Left            =   90
                  Locked          =   -1  'True
                  MaxLength       =   100
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   8
                  Text            =   "FrmManAlamacen.frx":3972
                  Top             =   3240
                  Width           =   6585
               End
               Begin VB.CommandButton CmdBusFam 
                  Height          =   240
                  Left            =   2040
                  Picture         =   "FrmManAlamacen.frx":398A
                  Style           =   1  'Graphical
                  TabIndex        =   93
                  Top             =   780
                  Width           =   240
               End
               Begin VB.CommandButton CmdBusTipiTem 
                  Height          =   240
                  Left            =   2040
                  Picture         =   "FrmManAlamacen.frx":3ABC
                  Style           =   1  'Graphical
                  TabIndex        =   40
                  Top             =   465
                  Width           =   240
               End
               Begin VB.TextBox TxtTipPro 
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   3
                  TabIndex        =   1
                  Text            =   "TxtTipPro"
                  Top             =   435
                  Width           =   915
               End
               Begin VB.CommandButton CmdBusUnidad 
                  Height          =   240
                  Left            =   2355
                  Picture         =   "FrmManAlamacen.frx":3BEE
                  Style           =   1  'Graphical
                  TabIndex        =   39
                  Top             =   4395
                  Width           =   240
               End
               Begin VB.TextBox TxtDescTecnica 
                  Height          =   480
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   100
                  TabIndex        =   6
                  Text            =   "TxtDescTecnica"
                  Top             =   2265
                  Width           =   5280
               End
               Begin VB.CommandButton CmdBusSubClase 
                  Height          =   240
                  Left            =   2040
                  Picture         =   "FrmManAlamacen.frx":3D20
                  Style           =   1  'Graphical
                  TabIndex        =   37
                  Top             =   1410
                  Width           =   240
               End
               Begin VB.TextBox TxtIdSubClase 
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   3
                  TabIndex        =   4
                  Text            =   "TxtIdSubClase"
                  Top             =   1380
                  Width           =   915
               End
               Begin VB.CommandButton CmdBusClase 
                  Height          =   240
                  Left            =   2040
                  Picture         =   "FrmManAlamacen.frx":3E52
                  Style           =   1  'Graphical
                  TabIndex        =   36
                  Top             =   1095
                  Width           =   240
               End
               Begin VB.TextBox TxtIdClase 
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   3
                  TabIndex        =   3
                  Text            =   "TxtIdClase"
                  Top             =   1065
                  Width           =   915
               End
               Begin VB.TextBox TxtCodPro 
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   0
                  Text            =   "TxtCodPro"
                  Top             =   120
                  Width           =   1770
               End
               Begin VB.TextBox TxtPrecio 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   5025
                  Locked          =   -1  'True
                  MaxLength       =   25
                  TabIndex        =   16
                  Text            =   "TxtPrecio"
                  Top             =   5640
                  Width           =   1470
               End
               Begin VB.TextBox TxtStockMin 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   5025
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   14
                  Text            =   "TxtStockMin"
                  Top             =   5325
                  Width           =   1470
               End
               Begin VB.TextBox TxtDescripcion 
                  Height          =   480
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   100
                  TabIndex        =   5
                  Text            =   "TxtDescripcion"
                  Top             =   1770
                  Width           =   5280
               End
               Begin VB.TextBox TxtUnidad 
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   5
                  TabIndex        =   9
                  Text            =   "TxtUnidad"
                  Top             =   4365
                  Width           =   1245
               End
               Begin VB.TextBox TxtStockAct 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   5025
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   12
                  Text            =   "TxtStockAct"
                  Top             =   5010
                  Width           =   1470
               End
               Begin VB.TextBox TxtStockIni 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   25
                  TabIndex        =   11
                  Text            =   "TxtStockIni"
                  Top             =   5010
                  Width           =   1470
               End
               Begin VB.CommandButton CmdBusMoneda 
                  Height          =   240
                  Left            =   2355
                  Picture         =   "FrmManAlamacen.frx":3F84
                  Style           =   1  'Graphical
                  TabIndex        =   35
                  Top             =   4710
                  Width           =   240
               End
               Begin VB.TextBox TxtIdMon 
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   1
                  TabIndex        =   10
                  Text            =   "TxtIdMon"
                  Top             =   4680
                  Width           =   1245
               End
               Begin VB.TextBox TxtIdFamilia 
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   3
                  TabIndex        =   2
                  Text            =   "TxtIdFamilia"
                  Top             =   750
                  Width           =   915
               End
               Begin VB.TextBox TxtIdTipmov 
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   1
                  TabIndex        =   7
                  Text            =   "TxtIdTipmov"
                  Top             =   2760
                  Width           =   915
               End
               Begin VB.TextBox TxtIdMatPri 
                  Height          =   300
                  Left            =   1395
                  Locked          =   -1  'True
                  MaxLength       =   2
                  TabIndex        =   17
                  Text            =   "TxtidMatPri"
                  Top             =   5985
                  Width           =   915
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Precio Inicial"
                  Height          =   195
                  Index           =   26
                  Left            =   90
                  TabIndex        =   115
                  Top             =   5685
                  Width           =   900
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Materia Prima"
                  Height          =   195
                  Index           =   25
                  Left            =   105
                  TabIndex        =   114
                  ToolTipText     =   "Materia Prima Principal"
                  Top             =   6015
                  Width           =   960
               End
               Begin VB.Label LblMatPri 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblMatPri"
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
                  Left            =   2340
                  TabIndex        =   113
                  Top             =   5985
                  Width           =   4170
               End
               Begin VB.Label LblTipoMovi 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblTipoMovi"
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
                  Left            =   2340
                  TabIndex        =   107
                  Top             =   2760
                  Width           =   4320
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo Movimiento"
                  Height          =   195
                  Index           =   23
                  Left            =   105
                  TabIndex        =   106
                  Top             =   2790
                  Width           =   1170
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Stock Máximo"
                  Height          =   195
                  Index           =   22
                  Left            =   105
                  TabIndex        =   105
                  Top             =   5370
                  Width           =   1005
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Características"
                  Height          =   195
                  Index           =   21
                  Left            =   105
                  TabIndex        =   100
                  Top             =   3030
                  Width           =   1065
               End
               Begin VB.Label LblPrefijo3 
                  Caption         =   "LblPrefijo3"
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   5790
                  TabIndex        =   99
                  Top             =   225
                  Visible         =   0   'False
                  Width           =   780
               End
               Begin VB.Label LblPrefijo2 
                  Caption         =   "LblPrefijo2"
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   4950
                  TabIndex        =   98
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   780
               End
               Begin VB.Label LblPrefijo1 
                  Caption         =   "LblPrefijo1"
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   4170
                  TabIndex        =   97
                  Top             =   195
                  Visible         =   0   'False
                  Width           =   780
               End
               Begin VB.Label LblPrefijo 
                  Caption         =   "LblPrefijo"
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   3465
                  TabIndex        =   96
                  Top             =   195
                  Visible         =   0   'False
                  Width           =   780
               End
               Begin VB.Label LblFamilia 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblFamilia"
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
                  Left            =   2355
                  TabIndex        =   95
                  Top             =   750
                  Width           =   4320
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Familia"
                  Height          =   195
                  Index           =   3
                  Left            =   105
                  TabIndex        =   94
                  Top             =   780
                  Width           =   480
               End
               Begin VB.Label LblIdUnidad 
                  Caption         =   "LblIdUnidad"
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   5685
                  TabIndex        =   58
                  Top             =   15
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo de Item"
                  Height          =   195
                  Index           =   0
                  Left            =   105
                  TabIndex        =   57
                  Top             =   465
                  Width           =   885
               End
               Begin VB.Label LblTipoPro 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblTipoPro"
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
                  Left            =   2355
                  TabIndex        =   56
                  Top             =   435
                  Width           =   4320
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
                  Left            =   2655
                  TabIndex        =   55
                  Top             =   4680
                  Width           =   2490
               End
               Begin VB.Label LblDescUnidad 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblDescUnidad"
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
                  Left            =   2655
                  TabIndex        =   54
                  Top             =   4365
                  Width           =   4035
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Desc. Técnica"
                  Height          =   195
                  Index           =   16
                  Left            =   105
                  TabIndex        =   53
                  Top             =   2310
                  Width           =   1050
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Sub Clase"
                  Height          =   195
                  Index           =   2
                  Left            =   105
                  TabIndex        =   52
                  Top             =   1410
                  Width           =   720
               End
               Begin VB.Label LblSubClase 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblSubClase"
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
                  Left            =   2355
                  TabIndex        =   51
                  Top             =   1380
                  Width           =   4320
               End
               Begin VB.Label Moneda 
                  AutoSize        =   -1  'True
                  Caption         =   "Moneda"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   50
                  Top             =   4740
                  Width           =   585
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Código"
                  Height          =   195
                  Index           =   7
                  Left            =   105
                  TabIndex        =   49
                  Top             =   150
                  Width           =   495
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Clase"
                  Height          =   195
                  Index           =   1
                  Left            =   105
                  TabIndex        =   48
                  Top             =   1095
                  Width           =   390
               End
               Begin VB.Label LblClase 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblClase"
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
                  Left            =   2355
                  TabIndex        =   47
                  Top             =   1065
                  Width           =   4320
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Stock Minimo"
                  Height          =   195
                  Index           =   4
                  Left            =   3735
                  TabIndex        =   46
                  Top             =   5355
                  Width           =   960
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Precio Actual"
                  Height          =   195
                  Index           =   5
                  Left            =   3735
                  TabIndex        =   45
                  Top             =   5685
                  Width           =   945
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Descripción"
                  Height          =   195
                  Index           =   6
                  Left            =   105
                  TabIndex        =   44
                  Top             =   1830
                  Width           =   840
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Unidad"
                  Height          =   195
                  Index           =   12
                  Left            =   105
                  TabIndex        =   43
                  Top             =   4395
                  Width           =   510
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Stock Actual"
                  Height          =   195
                  Index           =   13
                  Left            =   3735
                  TabIndex        =   42
                  Top             =   5055
                  Width           =   915
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Stock Inicial"
                  Height          =   195
                  Index           =   14
                  Left            =   105
                  TabIndex        =   41
                  Top             =   5055
                  Width           =   870
               End
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Item"
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
            Left            =   105
            TabIndex        =   31
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   27
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   28
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "codpro"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripción"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Unidad"
            Columns(3).DataField=   "abreunimed"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Moneda"
            Columns(4).DataField=   "simbolo"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Stock"
            Columns(5).DataField=   "stckact"
            Columns(5).NumberFormat=   "0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Tipo Item"
            Columns(6).DataField=   "desctippro"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Activo"
            Columns(7).DataField=   "xactivo"
            Columns(7).NumberFormat=   "General Number"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=3201"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3122"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=7461"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=7382"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1296"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1217"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1376"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1296"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1667"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1588"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=514"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=3493"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=3413"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=1164"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1085"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Named:id=33:Normal"
            _StyleDefs(69)  =   ":id=33,.parent=0"
            _StyleDefs(70)  =   "Named:id=34:Heading"
            _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   ":id=34,.wraptext=-1"
            _StyleDefs(73)  =   "Named:id=35:Footing"
            _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(75)  =   "Named:id=36:Selected"
            _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=37:Caption"
            _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(79)  =   "Named:id=38:HighlightRow"
            _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=39:EvenRow"
            _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(83)  =   "Named:id=40:OddRow"
            _StyleDefs(84)  =   ":id=40,.parent=33"
            _StyleDefs(85)  =   "Named:id=41:RecordSelector"
            _StyleDefs(86)  =   ":id=41,.parent=34"
            _StyleDefs(87)  =   "Named:id=42:FilterBar"
            _StyleDefs(88)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Items"
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
            Left            =   105
            TabIndex        =   29
            Top             =   30
            Width           =   11610
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
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
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Item"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar Item"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar un Item"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Retirar Item"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir "
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
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
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Eliminar             "
      End
   End
End
Attribute VB_Name = "FrmManAlamacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMANALAMACEN.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : AQUI SE CREAN, MODIFICAN Y ELIMINAN LOS ITEMS Y SE LES ASIGNA LA CUENTA CONTABLE.
'*                  : Y EL CENTRO DE COSTO.
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer              ' VARIABLE QUE ESPECIFICA EN QUE ESTADO SE ENCUENTRA EL FORMULARIO 1 = NUEVO; 2 = MODIFICA; 3 = SOLO LECTURA
Dim SeEjecuto As Boolean            ' VARIABLE QUE ESPECIFICA SI SE EJECUTO EL EVENTO ACTIVATE DEL FORMULARIO, SOLO ES USADO EN ESE EVENTO
Dim RstPro As New ADODB.Recordset   ' RECORDSET QUE ALAMCENARA LA LISTA DE PRODUCTOS REGISTRADOS
Dim RstTem As New ADODB.Recordset   ' RECORDSET TEMPORAL QUE SE USARA PARA ALGUNA OPERACIONES EN EL FORMULARIO
Dim xHorIni As Date                 ' ALMACENA LA HORA DE INICIO
Dim CODIGOTMP As Integer            ' UTIL PARA ALMACENAR LOS CODIGOS DE (TIPO DE ITEM,FAMILIA, CLASE, SUBCLASE)
                                    ' SE ALMACENARA ANTES DE BUSCAR O PRESS ENTER EN LOS TXT'S PARA VERIFICAR SI EL NUEVO REGISTRO TIENE CODIGO DIFERENTE AL TEMPORAL
                                    ' SI SON <>'S SE LIMPIARAN LOS DATOS DEPENDIENTE
                                    ' EJ. TIPO ITEM <> COD_TMP==>> LIMPIAR FAMILIA,CLASE, SUBCLASE

Dim fOrdenLista As Boolean                 ' --especfica el orden de la lista de la consulta
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim mIdRegistro& '--identificador del registro

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES DEL FORMULARIO PARA AGREGAR O MODIFICAR UN
'*                    REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'*****************************************************************************************************
Sub Bloquea()
    TxtCodPro.Locked = Not TxtCodPro.Locked
    TxtTipPro.Locked = Not TxtTipPro.Locked
    TxtIdClase.Locked = Not TxtIdClase.Locked
    TxtIdSubClase.Locked = Not TxtIdSubClase.Locked
    TxtIdFamilia.Locked = Not TxtIdFamilia.Locked
    TxtDescripcion.Locked = Not TxtDescripcion.Locked
    TxtDescTecnica.Locked = Not TxtDescTecnica.Locked
    TxtDescCaracteristica.Locked = Not TxtDescCaracteristica.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtStockIni.Locked = Not TxtStockIni.Locked
    TxtStockMin.Locked = Not TxtStockMin.Locked
    TxtStockMax.Locked = Not TxtStockMax.Locked
    TxtStockAct.Locked = Not TxtStockAct.Locked
    TxtPrecio.Locked = Not TxtPrecio.Locked
    TxtIdSelectivo.Locked = Not TxtIdSelectivo.Locked
    TxtIdTipCom.Locked = Not TxtIdTipCom.Locked
    TxtIdTipVen.Locked = Not TxtIdTipVen.Locked
    TxtIdTipmov.Locked = Not TxtIdTipmov.Locked
    
    TxtCtaCom.Locked = Not TxtCtaCom.Locked
    TxtCtaVen.Locked = Not TxtCtaVen.Locked
    TxtIdMatPri.Locked = Not TxtIdMatPri.Locked
    CmdAddFoto.Enabled = Not CmdAddFoto.Enabled
    CmdDelFoto.Enabled = Not CmdDelFoto.Enabled
    
    CmdAddCenCos.Enabled = Not CmdAddCenCos.Enabled
    CmdDelCenCos.Enabled = Not CmdDelCenCos.Enabled
    TxtPrecioIni.Locked = Not TxtPrecioIni.Locked
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LIMPIA LOS CONTTROLES DEL FORMULARIO PARA EL INGRESO DE NUEVOS DATOS, SE ACTIVA
'*                    CUANDO SE AGREGA UN NUEVO REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'*****************************************************************************************************
Sub Blanquea()
    TxtCodPro.Text = ""
    TxtTipPro.Text = "":            LblTipoPro.Caption = ""
    TxtIdFamilia.Text = "":         LblFamilia.Caption = ""
    TxtIdClase.Text = "":           LblClase.Caption = ""
    TxtIdSubClase.Text = "":        LblSubClase.Caption = ""
    TxtDescripcion.Text = "":
    TxtDescTecnica.Text = ""
    TxtDescCaracteristica.Text = ""
    TxtUnidad.Text = "":            LblDescUnidad.Caption = ""
    TxtIdMon.Text = "":             LblMoneda.Caption = ""
    TxtStockIni.Text = ""
    TxtStockMin.Text = ""
    TxtStockMax.Text = ""
    TxtStockAct.Text = ""
    TxtPrecio.Text = ""
    TxtPrecioIni.Text = ""
    
    TxtCtaCom.Text = "":            LblNomCtaCom.Caption = "":    LbIdCuentaCom.Caption = ""
    TxtCtaVen.Text = "":            LbIdCuentaVen.Caption = "":    LblNomCtaVen.Caption = ""
    TxtIdNetoDomic.Text = "":       LblNetoDomic.Caption = ""
    TxtidRet.Text = "":             LblRetencion.Caption = ""
    TxtIdDet.Text = "":             LblDetraccion.Caption = ""
    TxtIdPer.Text = "":             LblPercepcion.Caption = ""
    TxtIdSelectivo.Text = "":       LblSelectivo.Caption = ""
    TxtIdTipCom.Text = "":          LblIdTipCom.Caption = ""
    TxtIdTipVen.Text = "":          LblIdTipVen.Caption = ""
    TxtIdTipmov.Text = "":          LblTipoMovi.Caption = ""
    TxtTotPor.Text = ""
    TxtIdMatPri.Text = "":          LblMatPri.Caption = ""
End Sub

Private Sub CmdAddCenCos_Click()
    ' AGREGAMOS UN CENTRO DE COSTO, PARA ELLO LLAMAMOS A LA CLASE SeleCentroCosto PARA SELECCIONAR EL O LOS CENTRO DE COSTOS
    ' CORRESPONDIENTE
    If xDeDonde = 2 Then Exit Sub
    
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    Dim Encontro As Boolean
    Dim xFrm As New SGI2_funciones.formularios
    Set Rst = xFrm.SeleCentroCosto(xCon)
    
    If Rst.State = 1 Then
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            Encontro = False
            For A = 1 To Rst.RecordCount
                For B = 1 To Fg1.Rows - 1
                    If Fg1.TextMatrix(B, 4) = Rst("idcencos") Then
                        Encontro = True
                    End If
                Next B
                
                If Encontro = False Then
                    If Fg1.TextMatrix(Fg1.Rows - 1, 1) <> "" Then
                        Fg1.Rows = Fg1.Rows + 1
                    End If
                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("codigo")
                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("descripcion")
                    Fg1.TextMatrix(Fg1.Rows - 1, 4) = Rst("idcencos")
                End If
                
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
    End If
    Set xFrm = Nothing
End Sub

Private Sub CmdAddFoto_Click()
    Fg2.Rows = Fg2.Rows + 1
End Sub

Private Sub CmdBusClase_Click()
    ' BUSCAMOS LA CLASE QUE SE LE ASIGNARA AL ITEM
    If QueHace = 3 Then Exit Sub
    
    If VALIDAR_DATA(2) = False Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_clase.* FROM mae_clase WHERE mae_clase.idfam = " + CStr(Trim(TxtIdFamilia.Text)) + " ORDER BY mae_clase.descripcion ASC ;"
    
    xform.Titulo = "Buscando Clase"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        CODIGOTMP = NulosN(TxtIdClase.Text)
        TxtIdClase.Text = xRs("id")
        LblPrefijo2.Caption = xRs("prefijo")
        LblClase.Caption = xRs("descripcion")
        
        If CODIGOTMP <> 0 And CODIGOTMP <> NulosN(TxtTipPro.Text) Then
            TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
        End If
        TxtIdSubClase.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCtaCom_Click()
    ' BUSCAMOS LA CUENTA CONTABLE QUE SE LE ASIGNARA AL ITEM CUANDO SEA UNA COMPRA
    If xDeDonde = 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "cuenta":             xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":        xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT con_planctas.cuenta & '' as cuenta, con_planctas.descripcion, con_planctas.id " _
        & " From con_planctas ORDER BY con_planctas.cuenta"
    
    xform.Titulo = "Buscando Cuentas Contables"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCtaCom.Text = xRs("cuenta")
        LblNomCtaCom.Caption = xRs("descripcion")
        LbIdCuentaCom.Caption = xRs("id")
        TxtCtaVen.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PROCEDIMIENTO PARA BUSCAR UN REGISTRO EN EL RECORSET RstPro
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'*****************************************************************************************************
Sub Buscar()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "6000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "codpro":         xCampos(1, 2) = "2000":    xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.id From alm_inventario " _
        & " ORDER BY alm_inventario.descripcion"
    
    xform.Titulo = "Buscando Items"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            RstPro.MoveFirst
            RstPro.Find "id = " & xRs("id") & ""
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCtaVen_Click()
    ' BUSCAMOS LA CUENTA CONTABLE QUE SE LE ASIGNARA AL ITEM CUANDO SEA UNA VENTA
    If xDeDonde = 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "cuenta":             xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":        xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id " _
        & " From con_planctas ORDER BY con_planctas.cuenta"
    
    xform.Titulo = "Buscando Cuentas Contables"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCtaVen.Text = xRs("cuenta")
        LblNomCtaVen.Caption = xRs("descripcion")
        LbIdCuentaVen.Caption = xRs("id")
        TxtidRet.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusFam_Click()
    ' BUSCAMOS LA FAMILIA QUE SE LE ASIGNARA AL ITEM
    If QueHace = 3 Then Exit Sub
    If VALIDAR_DATA(1) = False Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_familia.* FROM mae_familia where mae_familia.idtippro = " + CStr(Trim(TxtTipPro.Text)) + " ORDER BY mae_familia.descripcion ASC ; "
    
    xform.Titulo = "Buscando Familia"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        CODIGOTMP = NulosN(TxtIdFamilia.Text)
        
        TxtIdFamilia.Text = xRs("id")
        LblPrefijo1.Caption = xRs("prefijo") & ""
        LblFamilia.Caption = xRs("descripcion") & ""
        
        If CODIGOTMP <> 0 And CODIGOTMP <> NulosN(TxtIdFamilia.Text) Then
            LblFamilia.Caption = "":
            TxtIdClase.Text = "":       LblClase.Caption = ""
            TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
        End If
        
        TxtIdClase.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMatPri_Click()
    'BUSCAMOS LA MATERIA PRIMA PRINCIPAL DEL PRODUCTO QUE SE LE ASIGNARA AL ITEM
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT pro_estacionalidad.* FROM pro_estacionalidad"
    
    xform.Titulo = "Buscando Materia Prima Principal"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMatPri.Text = xRs("id")
        LblMatPri.Caption = xRs("descripcion")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMoneda_Click()
    'BUSCAMOS LA MONEDA QUE SE LE ASIGNARA AL ITEM
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_moneda.* FROM mae_moneda"
    
    xform.Titulo = "Buscando Moneda"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMon.Text = xRs("id")
        LblMoneda.Caption = xRs("descripcion")
        TxtStockIni.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusSel_Click()
    'BUSCAMOS EL IMPUESTO SELECTIVO QUE SE LE APLICARA AL ITEM
    If xDeDonde = 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Descripcion":        xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "3000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Porcentaje":         xCampos(1, 1) = "tasa":           xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    xCampos(2, 0) = "Cuenta":             xCampos(2, 1) = "cuenta":         xCampos(2, 2) = "1500":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Descripcion Cuenta": xCampos(3, 1) = "descuen":        xCampos(3, 2) = "1500":         xCampos(3, 3) = "C"
    
    xform.SQLCad = "SELECT mae_impuestos.id, mae_impuestos.descripcion, mae_impuestos.tasa, mae_impuestos.idcuen, " _
        & " con_planctas.cuenta, con_planctas.descripcion AS descuen FROM mae_impuestos LEFT JOIN con_planctas " _
        & " ON mae_impuestos.idcuen = con_planctas.id"
   
    xform.Titulo = "Buscando Impuestos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdSelectivo.Text = xRs("id")
        LblSelectivo.Caption = xRs("descripcion")
        TxtIdTipCom.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusSubClase_Click()
    'BUSCAMOS LA SUB CLASE QUE SE LE ASIGNARA AL ITEM
    If QueHace = 3 Then Exit Sub

    If VALIDAR_DATA(3) = False Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_subclase.* FROM mae_subclase WHERE mae_subclase.idClas = " + Trim(TxtIdClase.Text) + " ORDER BY mae_subclase.descripcion ASC; "
    
    xform.Titulo = "Buscando Sub Clase"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdSubClase.Text = xRs("id")
        LblPrefijo3.Caption = xRs("prefijo") & ""
        LblSubClase.Caption = xRs("descripcion")
        TxtDescripcion.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipCom_Click()
    ' BUSCAMOS EL TIPO DE COMPRA QUE SE LE ASIGNARA AL ITEM
    If xDeDonde = 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "3000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipocompra.id, mae_tipocompra.descripcion From mae_tipocompra " _
        & " ORDER BY mae_tipocompra.descripcion"
   
    xform.Titulo = "Buscando Tipo de Compra"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdTipCom.Text = xRs("id")
        LblIdTipCom.Caption = xRs("descripcion")
        TxtIdTipVen.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipiTem_Click()
    ' BUSCANOS EL TIPO DE PRODUCTO QUE SE LE ASIGNARA A CADA ITEM
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
    
    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        CODIGOTMP = NulosN(TxtTipPro.Text)
        TxtTipPro.Text = xRs("id")
        LblTipoPro.Caption = xRs("descripcion")
        LblPrefijo.Caption = xRs("prefijo")
        
        If CODIGOTMP <> 0 And CODIGOTMP <> NulosN(TxtTipPro.Text) Then
            TxtIdFamilia.Text = "":     LblFamilia.Caption = "":
            TxtIdClase.Text = "":       LblClase.Caption = ""
            TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
        End If
        
        TxtIdFamilia.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipMovimiento_Click()
    'BUSCA EL TIPO DE MOVIMIENTO QUE SE LE ASIGNARA AL ITEM
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipomovimiento.* FROM mae_tipomovimiento"
    
    xform.Titulo = "Buscando Tipo de Movimiento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdTipmov.Text = xRs("id")
        LblTipoMovi.Caption = xRs("descripcion")
        TxtDescCaracteristica.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipVen_Click()
    'MUESTRA EL TIPO DE VENTA QUE SE APLICARA AL ITEM
    If xDeDonde = 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "3000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipoventa.id, mae_tipoventa.descripcion From mae_tipoventa " _
        & " ORDER BY mae_tipoventa.descripcion"
   
    xform.Titulo = "Buscando Tipo de I.G.V para la Venta"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdTipVen.Text = xRs("id")
        LblIdTipVen.Caption = xRs("descripcion")
        TxtCtaCom.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusUnidad_Click()
    'BUSCAMOS LA UNIDAD DE MEDIDA QUE SE LE ASIGNARA AL ITEM
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Abreviatura":   xCampos(1, 1) = "abrev":          xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_unidades.* FROM mae_unidades"
    
    xform.Titulo = "Buscando Unidades"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblIdUnidad.Caption = xRs("id")
        TxtUnidad.Text = xRs("abrev")
        LblDescUnidad.Caption = xRs("descripcion")
        TxtIdMon.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDelCenCos_Click()
    'ELIMINAMOS EL CENTRO DE COSTO
    If xDeDonde = 2 Then Exit Sub
    Fg1.RemoveItem Fg1.Row
End Sub

Private Sub CmdDelFoto_Click()
    'ELIMINAMOS LA FOTOGRAFIA ASIGNADA AL ITEM
    If Fg2.Rows = 1 Then
        MsgBox "No ha fotografias para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim Rpta As Integer
    Rpta = MsgBox("¿ Esta seguro de eliminar la imagen del item ?, La imagen no podra ser recuperada nuevamente", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Dim xFile As New FileSystemObject
        If NulosN(Fg2.TextMatrix(Fg2.Row, 3)) <> 0 Then
            If xFile.FileExists(NulosC(Fg2.TextMatrix(Fg2.Row, 2))) = True Then
                xFile.DeleteFile NulosC(Fg2.TextMatrix(Fg2.Row, 2))
            End If
            xCon.Execute "DELETE * FROM alm_inventariofoto WHERE idalm = " & RstPro("id") & " AND id = " & NulosN(Fg2.TextMatrix(Fg2.Row, 3)) & ""
            Fg2.RemoveItem Fg2.Row
        Else
            Fg2.RemoveItem Fg2.Row
        End If
        If Fg2.Rows > 1 Then
            Fg2.Select 1, 1, 1, 1
            Image1.Picture = LoadPicture(Fg2.TextMatrix(Fg2.Row, 2))
        Else
            Image1.Picture = LoadPicture("")
        End If
        MsgBox "La imagen fue eliminada con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub CmdIdDet_Click()
    ' BUSCAMOS LA DETRACCION QUE SE LE ASIGNARA AL ITEM
    If xDeDonde = 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_detraccion.* FROM mae_detraccion"
    
    xform.Titulo = "Buscando Detraccion"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdDet.Text = xRs("id")
        LblDetraccion.Caption = xRs("descripcion")
        TxtIdPer.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdIdPer_Click()
    'BUSCAMOS LA PERCEPCION QUE SE LE ASIGNARA AL ITEM
    If xDeDonde = 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_percepcion.* FROM mae_percepcion"
    
    xform.Titulo = "Buscando Percepcion"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdPer.Text = xRs("id")
        LblPercepcion.Caption = NulosC(xRs("descripcion"))
        TxtIdSelectivo.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdIdRet_Click()
    'BUSCAMOS LA RETENCION QUE SE LE ASIGNARA AL ITEM
    If xDeDonde = 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_retencion.* FROM mae_retencion"
    
    xform.Titulo = "Buscando Retenciones"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtidRet.Text = xRs("id")
        LblRetencion.Caption = xRs("descripcion")
        TxtIdDet.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstPro
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstPro.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    ' SI SE HA PRESIONADO LA TECLA F12 MOSTRAMOS LA INFORMACION DE EDICION DEL REGISTRO
    If KeyCode = 123 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstPro("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        CmdAddCenCos_Click
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : HallaTotal
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : HALLA EL TOTAL PARA LOS CENTRO DE COSTOS ASIGNADOS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'*****************************************************************************************************
Sub HallaTotal()
    Dim A As Integer
    Dim xTotal As Double
    
    For A = 1 To Fg1.Rows - 1
        xTotal = xTotal + NulosN(Fg1.TextMatrix(A, 3))
    Next A
    TxtTotPor.Text = Format(xTotal, "0.00")
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 3 Then
        HallaTotal
    End If
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Exit Sub
    If Fg1.Col = 1 Or Fg1.Col = 3 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then
        CmdAddCenCos_Click
    End If
    If KeyCode = 46 Then
        If Fg1.Rows = 1 Then
            MsgBox "No se han especificado centros de costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Fg1.RemoveItem Fg1.Row
        Fg1.Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If xDeDonde = 2 Then Exit Sub
    If Button = 2 Then
        If QueHace = 3 Then Exit Sub
        PopupMenu menu1
    End If
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    'ASIGNAMOS LA IMAGEN AL ITEM
    If Col = 1 Then
        CommonDialog1.Filter = "Archivos jpg (*.jpg)|*.jpg"
        CommonDialog1.ShowOpen
        Fg2.TextMatrix(Fg2.Row, 1) = CommonDialog1.FileName
        Image1.Picture = LoadPicture(Fg2.TextMatrix(Fg2.Row, 1))
    End If
End Sub

Private Sub Fg2_RowColChange()
    ' MOSTRAMOS LA IMAGEN ACTUAL DEL ITEM, CUANDO SE EJECUTE ESTE EVENTO
    If Fg2.Rows <> 1 Then Image1.Picture = LoadPicture(Fg2.TextMatrix(Fg2.Row, 1))
End Sub

Private Sub Form_Activate()
    ' CARGAMOS LOS ITEMS DEL INVENTARIO Y LOS MOSTRAMOS EN LA LA PRIMERA PESTAÑA DEL FORMULARIO, ESTE EVENTO SOLO SE EJECUTARA
    ' UNA SOLA VEZ
    If SeEjecuto = False Then
        Dim Rpta As Integer
        
        IdMenuActivo = xIdMenu
        
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        SeEjecuto = True
        RST_Busq RstPro, "SELECT alm_inventario.*, mae_tipoproducto.descripcion AS desctippro, mae_moneda.descripcion AS descmon, " _
            & " mae_unidades.abrev AS abreunimed, mae_unidades.descripcion AS descunimed, mae_moneda.simbolo, mae_clase.descripcion AS descclase, " _
            & " mae_subclase.descripcion AS descsubclase, mae_familia.descripcion AS descfam, IIf(alm_inventario.activo =-1,'Activo','De Baja') as xactivo " _
            & " FROM (((mae_unidades RIGHT JOIN (mae_tipoproducto RIGHT JOIN (alm_inventario LEFT JOIN mae_moneda ON " _
            & " alm_inventario.idmon = mae_moneda.id) ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) " _
            & " LEFT JOIN mae_clase ON alm_inventario.idclas = mae_clase.id) LEFT JOIN mae_subclase ON alm_inventario.idsubclas = mae_subclase.id) " _
            & " LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id ORDER BY alm_inventario.descripcion", xCon

        Set Dg1.DataSource = RstPro
        
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LIMPIA LOS CONTROLES DEL FORMULARIO PARA EL INGRESO DE UN NUEVO REGSITRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    Blanquea
    Bloquea
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Item"
    TabOne2.CurrTab = 0
    TxtIdTipCom.Text = "1"
    
    Fg1.Rows = 1
    Fg2.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    Fg1.ColComboList(1) = "|..."
    Fg2.ColComboList(1) = "|..."
    
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    
    Image1.Picture = LoadPicture("")
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    
    TxtIdTipCom_Validate True
    xHorIni = Time
    TxtCodPro.SetFocus
End Sub

Private Sub Form_Load()
    ' CARGAMOS EL FORMULARIO
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Frame4.BackColor = &H8000000F
    
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    Fg1.ColWidth(4) = 0
    
    Fg2.ColWidth(2) = 0
    Fg2.ColWidth(3) = 0
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    
    CaracteresNumericos = "0123456789." & Chr(8)
End Sub

Private Sub menu1_1_Click()
    CmdAddCenCos_Click
End Sub

Private Sub menu1_3_Click()
    CmdDelCenCos_Click
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 3 Then Exit Sub
        MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstPro.Requery
            Dg1.Refresh
            Cancelar
            
            If RstPro.RecordCount <> 0 Then
                RstPro.MoveFirst
                RstPro.Find "id=" & mIdRegistro
                If RstPro.EOF = True Then RstPro.MoveFirst
            End If
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then
        TabOne1.CurrTab = 0
        Filtrar
    End If
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstPro.Filter = ""
    End If
    
    If Button.Index = 10 Then
        TabOne1.CurrTab = 0
        Buscar
    End If
    
    If Button.Index = 12 Then pExportar
    
    If Button.Index = 13 Then
        FrmConsultaItems.Show vbModal
    End If
    
    If Button.Index = 15 Then
        Set RstPro = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FILATRO EL RECORSET PRINCIPAL RstPro PARA LA BUSQUEDA DE UN DETERMINADO ITEM
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'*****************************************************************************************************
Sub Filtrar()
    Dim xform As New eps_librerias.FormFiltrar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(5, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "C":         xCampos(0, 3) = "4200"
    xCampos(1, 0) = "Tipo Item":     xCampos(1, 1) = "desctippro":    xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Moneda":        xCampos(2, 1) = "simbolo":       xCampos(2, 2) = "C":         xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Unidad":        xCampos(3, 1) = "abreunimed":    xCampos(3, 2) = "C":         xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Stock":         xCampos(4, 1) = "stckact":       xCampos(4, 2) = "N":         xCampos(4, 3) = "1500"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstPro        'recorset que llena el grid
    Set RstPro = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstPro
    Dg1.Refresh
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCCION
'* Descripcion      : GRABA LOS DATOS EN UN NUEVO REGISTRO, DEVUELVE VERDADERO CUANDO TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'*****************************************************************************************************
Function Grabar() As Boolean
    ' VALIDAMOS QUE TODOS LOS CAMPOS ESTEN LLENADOS CORRECTAMENTE
    If VALIDAR_DATA(4) = False Then Exit Function
    
    If NulosC(TxtIdTipmov.Text) = "" Then
        TabOne2.CurrTab = 0
        MsgBox "No ha especificado el tipo de movimiento para el item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdTipmov.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtDescripcion.Text) = "" Then
        TabOne2.CurrTab = 0
        MsgBox "No ha especificado la descripcion del item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDescripcion.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtUnidad.Text) = "" Then
        MsgBox "No ha especificado la unidad de medida del item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TabOne2.CurrTab = 0
        TxtUnidad.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtIdMon.Text) = "" Then
        MsgBox "No ha especificado la moneda del item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TabOne2.CurrTab = 0
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtPrecioIni.Text) = "" Then
        MsgBox "No ha especificado el precio incial del item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TabOne2.CurrTab = 0
        TxtPrecioIni.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtPrecio.Text) = "" Then
        MsgBox "No ha especificado el precio unitario del item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TabOne2.CurrTab = 0
        TxtPrecio.SetFocus
        Exit Function
    End If
    
    If xDeDonde = 1 Then
        If NulosC(TxtCtaCom.Text) = "" And NulosC(TxtCtaVen.Text) = "" Then
            MsgBox "No ha especificado Cuenta Contable de Compra o Venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TabOne2.CurrTab = 1
            TxtCtaCom.SetFocus
            Exit Function
        End If
    End If
    
    Dim Rpta As Integer
    Dim A As Integer
    
    'ELIMINAMOS LAS FILAS QUE ESTEN EN BLANCO DE LA CUADRICULA DE CENTRO DE COSTOS
    If Fg1.Rows >= 2 Then
        For A = 1 To Fg1.Rows - 1
            If NulosC(Fg1.TextMatrix(A, 1)) = "" Then
                Fg1.RemoveItem A
                A = A - 1
            End If
            If A = Fg1.Rows - 1 Then
                Exit For
            End If
        Next A
    End If
    
    Dim Preguntar As Boolean
    
    Preguntar = True
    If Fg1.Rows - 1 = 0 Then
        ' SI NO SE HA AGREGADO CENTRO DE COSTOS, PREGUNTAMOS SI SE DESEA AGREGAR UN CENTRO DE COSTOS
        Rpta = MsgBox("No ha especificado la distribución de centro de costos para este item, ¿Desea agregarlo ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            TabOne2.CurrTab = 1
            Fg1.SetFocus
            Exit Function
        Else
            Preguntar = False
        End If
    End If
    
    If Preguntar = True Then
        ' VERIFICAMOS QUE LA DISTRIBUCION DEL CENTRO DE COSTO ESTE AL 100%
        For A = 1 To Fg1.Rows - 1
            If NulosC(Fg1.TextMatrix(A, 3)) = "" Then
                MsgBox "No ha especificado el porcentaje de distribución para el Centro de Costo " + Trim(Fg1.TextMatrix(A, 2)), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                TabOne2.CurrTab = 1
                Fg1.SetFocus
                Exit Function
            End If
        Next A
        
        If NulosN(TxtTotPor.Text) <> 100 Then
            MsgBox "La distribución del Centro de Costo no es al 100%, verifique la distribución", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TabOne2.CurrTab = 1
            Fg1.SetFocus
            Exit Function
        End If
    End If
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " el Item", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim RstDetCenCos As New ADODB.Recordset
    Dim xId As Double
    
    On Error GoTo LaCague
    xCon.BeginTrans
    
    If QueHace = 1 Then
        ' SI SE ESTA AGREGANDO UN NUEVO REGISTRO, OBTENEMOS EL ULTIMO ID Y PREPARAMOS LOS RECORDSET PARA GRABAR EL NUEVO REGISTRO
        xId = HallaCodigoTabla("alm_inventario", xCon, "id")
        RST_Busq RstCab, "SELECT * FROM alm_inventario", xCon
        RST_Busq RstDetCenCos, "SELECT * FROM alm_invencencos", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
    Else
        ' SI SE ESTA MODIFICANDO OBTENEMOS EL ID DEL REGISTRO QUE SE ESTA EDITANDO Y CARGAMOS EL RECORSET CON EL REGISTRO ACTUAL PARA
        ' REEMPLAZAR LOS NUEVOS DATOS
        xId = RstPro("id")
        RST_Busq RstCab, "SELECT * FROM alm_inventario WHERE id = " & xId & "", xCon
        xCon.Execute "DELETE * FROM alm_invencencos WHERE idpro = " & xId & ""
        RST_Busq RstDetCenCos, "SELECT * FROM alm_invencencos", xCon
    End If
   
    ' VALIDAMOS EL CODIGO AUTOGENERADO DEL ITEM
    TxtCodPro.Tag = DevuelveCodigo(TxtCodPro.Text, QueHace)
    If Trim(TxtCodPro.Text) = "" Then
        TxtCodPro.Text = TxtCodPro.Tag
    ElseIf Trim(TxtCodPro.Text) <> Trim(TxtCodPro.Tag) Then
        If MsgBox("El código del Item esta por cambiarse" + vbCr + "Código anterior: " + Trim(TxtCodPro.Text) + vbCr + "Nuevo Código:   " + Trim(TxtCodPro.Tag) + vbCr + "Seguro desea cabiar el Código del Item...", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbYes Then
            TxtCodPro.Text = TxtCodPro.Tag
        End If
    End If
    
    mIdRegistro = xId
    
    RstCab("codpro") = NulosC(TxtCodPro.Text)
    RstCab("descripcion") = TxtDescripcion.Text
    RstCab("desctec") = NulosC(TxtDescTecnica.Text)
    RstCab("idunimed") = NulosN(LblIdUnidad.Caption)
    RstCab("preini") = NulosN(TxtPrecioIni.Text)
    RstCab("preuni") = NulosN(TxtPrecio.Text)
    RstCab("tippro") = NulosN(TxtTipPro.Text)
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    
    RstCab("stckini") = NulosN(NulosN(TxtStockIni.Text))
    RstCab("stckact") = NulosN(NulosN(TxtStockAct.Text))
    RstCab("stckmin") = NulosN(NulosN(TxtStockMin.Text))
    RstCab("stckmax") = NulosN(NulosN(TxtStockMax.Text))
    RstCab("idclas") = NulosN(TxtIdClase.Text)
    RstCab("idsubclas") = NulosN(TxtIdSubClase.Text)
    RstCab("idfam") = NulosN(TxtIdFamilia.Text)
       
    RstCab("idcuenta") = NulosN(LbIdCuentaCom.Caption)
    RstCab("idcuentaven") = NulosN(LbIdCuentaVen.Caption)
    
    RstCab("idnetonodomi") = NulosN(TxtIdNetoDomic.Text)
    
    RstCab("idret") = NulosN(TxtidRet.Text)
    RstCab("iddet") = NulosN(TxtIdDet.Text)
    RstCab("idper") = NulosN(TxtIdPer.Text)
    
    RstCab("idimpsel") = NulosN(TxtIdSelectivo.Text)
    RstCab("idtipcom") = NulosN(TxtIdTipCom.Text)
    RstCab("idtipven") = NulosN(TxtIdTipVen.Text)
    
    RstCab("tipo") = NulosN(TxtIdTipmov.Text)
    RstCab("caracteristica") = NulosC(TxtDescCaracteristica.Text)
    If NulosN(TxtIdMatPri.Text) <> 0 Then
        RstCab("idmatpri") = NulosN(TxtIdMatPri.Text)
    End If
    RstCab.Update
    
    'grabamos la distribucion del centro de costos
    If Fg1.Rows > 1 Then
        For A = 1 To Fg1.Rows - 1
            RstDetCenCos.AddNew
            RstDetCenCos("idpro") = xId
            RstDetCenCos("idcencos") = Fg1.TextMatrix(A, 4)
            RstDetCenCos("imppor") = Fg1.TextMatrix(A, 3)
            RstDetCenCos.Update
        Next A
    End If
    
    'grabamos la foto y la copiamos a su ruta final
    Dim Rst As New ADODB.Recordset
    Dim xArchivo, xRuta As String
    Dim xFile As New FileSystemObject
    Dim xIdFoto, xItemFoto As Integer
    
    If Fg2.Rows > 1 Then
        RST_Busq Rst, "SELECT * FROM alm_inventariofoto", xCon
        
        For A = 1 To Fg2.Rows - 1
            xArchivo = ""
            xRuta = ""
            If NulosN(Fg2.TextMatrix(A, 3)) = 0 Then
                Rst.AddNew
                Rst("idalm") = xId
                xItemFoto = xItemFoto + 1
                Fg2.TextMatrix(A, 3) = xItemFoto
                Rst("id") = xItemFoto
                Rst("descripcion") = Fg2.TextMatrix(A, 1)
                
                xArchivo = Format(xId, "0000") + "-" + Format(NulosN(Fg2.TextMatrix(A, 3)), "0000") + ".jpg"
                xRuta = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTAAR", "RUTAS")
                xRuta = xRuta + "items\" + xArchivo
                xFile.CopyFile NulosC(Fg2.TextMatrix(A, 1)), xRuta
                Rst("archivo") = NulosC(xArchivo)
                
                Rst.Update
            Else
                xItemFoto = Fg2.TextMatrix(A, 3)
            End If
        Next A
    End If
    '*************************************************************************************
    '*** SINCRONIZAR BASE DE DATOS - alm_inventario ***'
    If xDeDonde = 2 Then
        SincronizarBD xCon, "alm_inventario", xId, QueHace
    End If
    '*************************************************************************************
    'grabamos los datos de la operacion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    '*************************************************************************************
    
    xCon.CommitTrans
    MsgBox "El item se grabo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDetCenCos = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description), vbCritical, xTitulo
End Function

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELAR EL PROCESO DE AGRGAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
    ActivaTool
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA UN REGISTRO PARA SU MODIFICACION
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    Bloquea
    ActivaTool
    If TabOne1.CurrTab = 0 Then
        Blanquea
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If

    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Item"
    Fg1.ColComboList(1) = "|..."
    Fg2.ColComboList(1) = "|..."
    
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    
    Fg2.Editable = flexEDKbdMouse
    Fg2.SelectionMode = flexSelectionFree
    
    TabOne2.CurrTab = 0
    xHorIni = Time
    TxtCodPro.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA EL REGISTRO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Dim xRs As New ADODB.Recordset
    Dim xCad As String
    
    If xDeDonde = 2 Then Exit Sub '--es unificado
    
    ' VERIFICAMOS QUE NO SE HAYAN HECHO OPERACIONES DE COMPRA CON EL ITEM SELECCIONADO
    If (RstPro("tippro") = 1) Or (RstPro("tippro") = 4) Or (RstPro("tippro") = 2) Or (RstPro("tippro") = 5) Then
        RST_Busq xRs, "SELECT com_comprasdet.iditem From com_comprasdet WHERE com_comprasdet.iditem = " & RstPro("id") & "", xCon
        
        If xRs.RecordCount <> 0 Then
            MsgBox "No se puede eliminar el item seleccionado, esta registrado en una compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set xRs = Nothing
            Exit Sub
        End If
    End If
    
    ' VERIFICAMOS QUE NO SE HAYAN HECHO OPERACION DE VENTA CON EL ITEM SELECCIONADO
    If (RstPro("tippro") = 3) Then
        RST_Busq xRs, "SELECT vta_ventasdet.iditem From vta_ventasdet WHERE (((vta_ventasdet.iditem) = " & RstPro("id") & "))", xCon
        
        If xRs.RecordCount <> 0 Then
            MsgBox "No se puede eliminar el item seleccionado, esta registrado en una venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set xRs = Nothing
            Exit Sub
        End If
    End If
    
    ' SI EL ITEM NO TIENE NINGUNA OPERACION SE PROCEDE A ELIMINAR PREVIA AUTORIZACION DEL USUARIO
    Rpta = MsgBox("¿ Esta seguro de eliminar el item ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM alm_inventario WHERE id = " & RstPro("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstPro("id") & " AND idform = " & IdMenuActivo

        
        MsgBox "El item se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstPro.Requery
        Dg1.Refresh
        Exit Sub
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : DevuelveCodigo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DEVUELVE EL CODIGO AUTOGENERADO DEL ITEM QUE SE ESRA INGRESANDO, GENERA EL
'*                    CODIGO EN FUNCION AL CAMPO prefijo DE LAS TABLAS mae_tipoproducto, mae_familia,
'*                    mae_clase, mae_subclase
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                  ----------------------------------------------------------------------------------
'*                    Codigo    |  STRING           |  CODIGO DEL ITEM QUE SE ESTA MODIFICANDO
'*                    xQueHace  |  INTEGER          |  ESPECIFICA SI SE MODIFICA O AGREGA UN REGISTRO
'* Devuelve         : STRING
'*****************************************************************************************************
Function DevuelveCodigo(xCodigo As String, xQueHace As Integer) As String
    'xQueHace = 1 'esta añadiendo un nuevo item
    'xQueHace = 2 'esta modificando un item
    
    Dim Rst As New ADODB.Recordset

    ' BUSCAMOS SI LA CLASIFICACION DEL ITEM EXISTE PARA GENERAR EL NUEVO CODIGO
    RST_Busq Rst, "SELECT alm_inventario.tippro, alm_inventario.idfam, alm_inventario.idclas, alm_inventario.idsubclas, " _
        & " alm_inventario.codpro From alm_inventario Where (((alm_inventario.tippro) = " & NulosN(TxtTipPro.Text) & ") " _
        & " And ((alm_inventario.idfam) = " & NulosN(TxtIdFamilia.Text) & ") And ((alm_inventario.idclas) = " & NulosN(TxtIdClase.Text) & ") " _
        & " And ((alm_inventario.idsubclas) = " & NulosN(TxtIdSubClase.Text) & ")) ORDER BY alm_inventario.codpro", xCon
    
    If Rst.RecordCount = 0 Then
        ' SI NO EXISTE LA CLASIFICACION DEL ITEM, INICIALIZAMOS LA NUMERACION DEL CODIGO
        DevuelveCodigo = Trim(LblPrefijo.Caption) + Trim(LblPrefijo1.Caption) + Trim(LblPrefijo2.Caption) + Trim(LblPrefijo3.Caption) + "0001"
    Else
        ' SI EXISTE BUSCAMOS EL ULTIMO REGISTRO Y LE SUMAMOS UNO PARA INCREMENTAR EL CONTADOR DEL CODIGO
        Rst.MoveLast
        Dim xLongitud As Integer
        Dim xNumCodigo As String
        Dim xCodigo2 As String
        
        If xQueHace = 1 Then
            xLongitud = Len(Trim(NulosC(Rst("codpro"))))
            xNumCodigo = NulosN(Mid(Rst("codpro"), (xLongitud - 4) + 1, 4))
            xCodigo2 = Format(NulosN(xNumCodigo) + 1, "0000")
            xCodigo2 = Trim(LblPrefijo.Caption) + Trim(LblPrefijo1.Caption) + Trim(LblPrefijo2.Caption) + Trim(LblPrefijo3.Caption) + xCodigo2
        Else
            If NulosC(TxtCodPro.Text) = "" Then
                xLongitud = Len(Trim(Rst("codpro")))
                xNumCodigo = Mid(Rst("codpro"), (xLongitud - 4) + 1, 4)
                xCodigo2 = Format(NulosN(xNumCodigo) + 1, "0000")
                xCodigo2 = Trim(LblPrefijo.Caption) + Trim(LblPrefijo1.Caption) + Trim(LblPrefijo2.Caption) + Trim(LblPrefijo3.Caption) + xCodigo2
            Else
                xCodigo2 = Trim(TxtCodPro.Text)
            End If
        End If
        DevuelveCodigo = xCodigo2
    End If
    Set Rst = Nothing
End Function

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        If ButtonMenu.Index = 1 Then
            Modificar
        End If
        If ButtonMenu.Index = 2 Then
            ActivarItem
        End If
    End If
    
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then
            Eliminar
        End If
        If ButtonMenu.Index = 2 Then
            Retirar
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ActivarItem
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA UN ITEM DESACTIVADO PARA SU VISUALIZACION EN LAS CONSULTAS DEL SISTEMA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivarItem()
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "6000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "codpro":         xCampos(1, 2) = "2000":    xCampos(1, 3) = "C"
    
    ' CARGAMOS LOS ITEMS QUE HAYAN SIDO DESACTIVADOS
    xform.SQLCad = "SELECT alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.id From alm_inventario " _
        & " WHERE activo = 0 ORDER BY alm_inventario.descripcion"
    
    xform.Titulo = "Buscando Items Retirados"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Dim Rpta As Integer
            Rpta = MsgBox("¿Esta seguro de activar el item seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                ' ACTIVAMOS EL ITEM ACTUALIZADO EL CAMPO activo A -1 DE LA TABLA alm_inventario
                xCon.Execute "UPDATE alm_inventario SET alm_inventario.activo = -1 WHERE (((alm_inventario.id)=" & xRs("id") & "))"
                
                xHorIni = Time
                'grabamos los datos de la operacion
                GrabarOperacion xIdUsuario, IdMenuActivo, 2, xHorIni, Time, Date, xCon, CDbl(xRs("id"))
                '*************************************************************************************

                
                MsgBox "El item " + Trim(xRs("descripcion")) + " se activo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                RstPro.Requery
                Dg1.Refresh
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Retirar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DESACTIVA UN ITEM ACTIVO, LOS ITEMS DESACTIVADOS NO PODRAN SER VISUALIZADOS EN
'*                    EL SISTEMA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Retirar()
    Dim Rpta As Integer
    xHorIni = Time
    Rpta = MsgBox("Esta seguro de retirar el item " + Trim(RstPro("descripcion")), vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        ' DESACTIVAMOS EL ITEM ACUTALIZANDO EL CAMPO activo a 0 DE LA TABLA alm_inventario
        xCon.Execute "UPDATE alm_inventario SET alm_inventario.activo = 0 WHERE (((alm_inventario.id)=" & RstPro("id") & "))"
        'grabamos los datos de la operacion
        GrabarOperacion xIdUsuario, IdMenuActivo, 2, xHorIni, Time, Date, xCon, CDbl(RstPro("id"))
        '*************************************************************************************
        
        
        MsgBox "El item " + Trim(RstPro("descripcion")) + " se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstPro.Requery
        Dg1.Refresh
    End If
End Sub

Private Sub TxtCodPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtCtaCom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtCtaCom_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 32 Then
        If QueHace = 3 Then Exit Sub
        TxtCtaCom.Text = ""
        LblNomCtaCom.Caption = ""
        LbIdCuentaCom.Caption = ""
    End If
    
    If KeyCode = 116 Then
        CmdBusCtaCom_Click
    End If
End Sub

Private Sub TxtCtaCom_Validate(Cancel As Boolean)
    If Trim(TxtCtaCom.Text) = "" Then
        LblNomCtaCom.Caption = ""
        LbIdCuentaCom.Caption = ""
        Exit Sub
    End If
    LblNomCtaCom.Caption = Busca_Codigo(NulosC(TxtCtaCom.Text), "cuenta", "descripcion", "con_planctas", "C", xCon)
    If LblNomCtaCom.Caption <> "" Then
        LbIdCuentaCom.Caption = Busca_Codigo(NulosC(TxtCtaCom.Text), "cuenta", "id", "con_planctas", "C", xCon)
    Else
        LbIdCuentaCom.Caption = ""
    End If
End Sub

Private Sub TxtCtaVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtCtaVen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 32 Then
        If QueHace = 3 Then Exit Sub
        TxtCtaVen.Text = ""
        LblNomCtaVen.Caption = ""
        LbIdCuentaVen.Caption = ""
    End If
    
    If KeyCode = 116 Then
        CmdBusCtaVen_Click
    End If
End Sub

Private Sub TxtCtaVen_Validate(Cancel As Boolean)
    If Trim(TxtCtaVen.Text) = "" Then
        LblNomCtaVen.Caption = ""
        LbIdCuentaVen.Caption = ""
        Exit Sub
    End If
    LblNomCtaVen.Caption = Busca_Codigo(NulosC(TxtCtaVen.Text), "cuenta", "descripcion", "con_planctas", "C", xCon)
    If LblNomCtaVen.Caption <> "" Then
        LbIdCuentaVen.Caption = Busca_Codigo(NulosC(TxtCtaVen.Text), "cuenta", "id", "con_planctas", "C", xCon)
    Else
        LbIdCuentaVen.Caption = ""
    End If
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDescTecnica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdClase_Change()
    If Trim(TxtIdClase.Text) = "" Then
        LblClase.Caption = ""
        TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
    End If
End Sub

Private Sub TxtIdClase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtIdClase.Text <> "" Then
            If VALIDAR_DATA(2) = False Then Exit Sub
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_clase WHERE id =" & NulosN(TxtIdClase.Text) & " AND mae_clase.idfam=" + CStr(Trim(TxtIdFamilia.Text)), xCon)
            If RstTem.RecordCount <> 0 Then
                LblClase.Caption = RstTem("descripcion") & ""
                LblPrefijo2.Caption = RstTem("prefijo") & ""
                TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
            Else
                TxtIdClase.Text = "":   LblClase.Caption = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdClase_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusClase_Click
    End If
End Sub

Private Sub TxtIdDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDet_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 32 Then
        If QueHace = 3 Then Exit Sub
        TxtIdDet.Text = ""
        LblDetraccion.Caption = ""
    End If
    
    If KeyCode = 116 Then
        CmdIdDet_Click
    End If
End Sub

Private Sub TxtIdFamilia_Change()
    If Trim(TxtIdFamilia.Text) = "" Then
        TxtIdFamilia.Text = "":     LblFamilia.Caption = ""
        TxtIdClase.Text = "":       LblClase.Caption = ""
        TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
    End If
End Sub

Private Sub TxtIdFamilia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtIdFamilia.Text <> "" Then
            If VALIDAR_DATA(1) = False Then Exit Sub
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_familia WHERE id =" & NulosN(TxtIdFamilia.Text) & " AND mae_familia.idtippro = " + CStr(Trim(TxtTipPro.Text)), xCon)
            If RstTem.RecordCount <> 0 Then
                LblFamilia.Caption = RstTem("descripcion") & ""
                LblPrefijo1.Caption = NulosC(RstTem("prefijo")) & ""
            Else
                TxtIdFamilia.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdFamilia_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusFam_Click
    End If
End Sub

Private Sub TxtidMatPri_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtIdMatPri.Text <> "" Then
            Set RstTem = BuscaConCriterio("SELECT * FROM pro_estacionalidad WHERE id =" & NulosN(TxtIdMatPri.Text) & "", xCon)
            If RstTem.RecordCount <> 0 Then
                LblMatPri.Caption = RstTem("descripcion") & ""
                TxtIdMatPri.Text = RstTem("id") & ""
            Else
                TxtIdMatPri.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtidMatPri_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMatPri_Click
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtIdMon.Text <> "" Then
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_moneda WHERE id =" & NulosN(TxtIdMon.Text) & "", xCon)
            If RstTem.RecordCount <> 0 Then
                LblMoneda.Caption = RstTem("descripcion")
            Else
                TxtIdMon.Text = ""
                LblMoneda.Caption = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMoneda_Click
    End If
End Sub

Private Sub TxtIdPer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdPer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 32 Then
        If QueHace = 3 Then Exit Sub
        TxtIdPer.Text = ""
        LblPercepcion.Caption = ""
    End If
    
    If KeyCode = 116 Then
        CmdIdPer_Click
    End If
End Sub

Private Sub TxtidRet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtidRet_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 32 Then
        If QueHace = 3 Then Exit Sub
        TxtidRet.Text = ""
        LblRetencion.Caption = ""
    End If
    
    If KeyCode = 116 Then
        CmdIdRet_Click
    End If
End Sub


Private Sub TxtIdSelectivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdSelectivo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 32 Then
        If QueHace = 3 Then Exit Sub
        TxtIdSelectivo.Text = ""
        LblSelectivo.Caption = ""
    End If
    
    If KeyCode = 116 Then
        CmdBusSel_Click
    End If
End Sub

Private Sub TxtIdSelectivo_Validate(Cancel As Boolean)
    If NulosC(TxtIdSelectivo.Text) = "" Then Exit Sub
    
    LblSelectivo.Caption = Busca_Codigo(NulosN(TxtIdSelectivo.Text), "id", "descripcion", "mae_impuestos", "N", xCon)
    If LblSelectivo.Caption = "" Then
        TxtIdSelectivo.Text = ""
    End If
End Sub

Private Sub TxtIdSubClase_Change()
    If Trim(TxtIdSubClase.Text) = "" Then
        LblSubClase.Caption = ""
    End If
End Sub

Private Sub TxtIdSubClase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtIdSubClase.Text <> "" Then
            If VALIDAR_DATA(3) = False Then Exit Sub
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_subclase WHERE id =" & NulosN(TxtIdSubClase.Text) & " AND mae_subclase.idClas = " + Trim(TxtIdClase.Text), xCon)
            If RstTem.RecordCount <> 0 Then
                LblSubClase.Caption = RstTem("descripcion")
                LblPrefijo3.Caption = RstTem("prefijo")
            Else
                LblSubClase.Caption = ""
                TxtIdSubClase.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdSubClase_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusSubClase_Click
    End If
End Sub

Private Sub TxtIdTipCom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub TxtIdTipCom_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 32 Then
        If QueHace = 3 Then Exit Sub
        TxtIdTipCom.Text = ""
        LblIdTipCom.Caption = ""
    End If
    
    If KeyCode = 116 Then
        CmdBusTipCom_Click
    End If
End Sub

Private Sub TxtIdTipCom_Validate(Cancel As Boolean)
    If xDeDonde = 2 Then Exit Sub
    If NulosC(TxtIdTipCom.Text) = "" Then
        LblIdTipCom.Caption = ""
        Exit Sub
    End If
    
    LblIdTipCom.Caption = Busca_Codigo(TxtIdTipCom.Text, "id", "descripcion", "mae_tipocompra", "N", xCon)
    If LblIdTipCom.Caption = "" Then
        TxtIdTipCom.Text = ""
    End If
End Sub

Private Sub TxtIdTipmov_Change()
    If Trim(TxtIdTipmov.Text) = "" Then
        LblTipoMovi.Caption = ""
    End If
End Sub

Private Sub TxtIdTipmov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtIdTipmov.Text <> "" Then
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_tipomovimiento WHERE id =" & NulosN(TxtIdTipmov.Text) & "", xCon)
            If RstTem.RecordCount <> 0 Then
                LblTipoMovi.Caption = RstTem("descripcion") & ""
                TxtIdTipmov.Text = RstTem("id") & ""
            Else
                TxtIdTipmov.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdTipmov_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipMovimiento_Click
    End If
End Sub

Private Sub TxtIdTipVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdTipVen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 32 Then
        If QueHace = 3 Then Exit Sub
        TxtIdTipVen.Text = ""
        LblIdTipVen.Caption = ""
    End If
    
    If KeyCode = 116 Then
        CmdBusTipVen_Click
    End If
End Sub

Private Sub TxtIdTipVen_Validate(Cancel As Boolean)
    If NulosC(TxtIdTipVen.Text) = "" Then
        LblIdTipVen.Caption = ""
        Exit Sub
    End If
    
    LblIdTipVen.Caption = Busca_Codigo(TxtIdTipVen.Text, "id", "descripcion", "mae_tipoventa", "N", xCon)
    If LblIdTipVen.Caption = "" Then
        TxtIdTipVen.Text = ""
    End If
End Sub

Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtPrecioIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtStockAct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtStockIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtStockMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtStockMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipPro_Change()
    If Trim(TxtTipPro.Text) = "" Then
        LblTipoPro.Caption = ""
        TxtIdFamilia.Text = "":     LblFamilia.Caption = "":
        TxtIdClase.Text = "":       LblClase.Caption = ""
        TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
    End If
End Sub

Private Sub TxtTipPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtTipPro.Text <> "" Then
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_tipoproducto WHERE id =" & NulosN(TxtTipPro.Text) & "", xCon)
            If RstTem.RecordCount <> 0 Then
                
                LblTipoPro.Caption = RstTem("descripcion") & ""
                LblPrefijo.Caption = RstTem("prefijo") & ""

                TxtIdFamilia.Text = "":     LblFamilia.Caption = "":
                TxtIdClase.Text = "":       LblClase.Caption = ""
                TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
                
            Else
                TxtTipPro.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipiTem_Click
    End If
End Sub

Private Sub TxtUnidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DEL REGISTRO ACTUAL, ESTE EVENTO SE EJECUTA CUANDO EL
'*                    FORMULARIO ESTA EN MODO DE LECTURA O MODIFICAR
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Blanquea
    TabOne2.CurrTab = 0
    TxtCodPro.Text = NulosC(RstPro("codpro"))
    TxtTipPro.Text = RstPro("tippro")
    TxtIdClase.Text = NulosN(RstPro("idclas"))
    TxtIdSubClase.Text = NulosN(RstPro("idsubclas"))
    TxtIdFamilia.Text = NulosN(RstPro("idfam"))
    TxtDescripcion.Text = NulosC(RstPro("descripcion"))
    TxtDescTecnica.Text = NulosC(RstPro("desctec"))
    TxtUnidad.Text = NulosC(RstPro("abreunimed"))
    TxtIdMon.Text = RstPro("idmon")
    TxtIdTipmov.Text = NulosN(RstPro("tipo"))
    If NulosN(RstPro("idmatpri")) <> 0 Then
        TxtIdMatPri.Text = NulosN(RstPro("idmatpri"))
        LblMatPri.Caption = Busca_Codigo(RstPro("idmatpri"), "id", "descripcion", "pro_estacionalidad", "N", xCon)
    Else
        TxtIdMatPri.Text = ""
    End If
    
    TxtStockIni.Text = Format(NulosN(RstPro("stckini")), "0.0000")
    TxtStockMin.Text = Format(NulosN(RstPro("stckmin")), "0.0000")
    TxtStockMax.Text = Format(NulosN(RstPro("stckmax")), "0.0000")
    TxtStockAct.Text = Format(NulosN(RstPro("stckact")), "0.0000")
    TxtPrecioIni.Text = Format(NulosN(RstPro("preini")), "0.0000")
    TxtPrecio.Text = Format(NulosN(RstPro("preuni")), "0.0000")
    TxtDescCaracteristica.Text = NulosC(RstPro("caracteristica"))
    TxtIdTipmov.Text = NulosN(RstPro("tipo"))
    
    If NulosN(RstPro("tipo")) <> 0 Then
        LblTipoMovi.Caption = Busca_Codigo(RstPro("tipo"), "id", "descripcion", "mae_tipomovimiento", "N", xCon)
    End If
    'mostramos los datos contables
    If NulosN(RstPro("idret")) <> 0 Then
        TxtidRet.Text = NulosN(RstPro("idret"))
        LblRetencion.Caption = Busca_Codigo(RstPro("idret"), "id", "descripcion", "mae_retencion", "N", xCon)
    End If
    If NulosN(RstPro("iddet")) <> 0 Then
        TxtIdDet.Text = NulosN(RstPro("iddet"))
        LblDetraccion.Caption = Busca_Codigo(RstPro("iddet"), "id", "descripcion", "mae_detraccion", "N", xCon)
    End If
    If NulosN(RstPro("idper")) <> 0 Then
        TxtIdPer.Text = NulosN(RstPro("idper"))
        LblPercepcion.Caption = NulosC(Busca_Codigo(RstPro("idper"), "id", "descripcion", "mae_percepcion", "N", xCon))
    End If
    
    If NulosN(RstPro("idcuenta")) <> 0 Then
        LbIdCuentaCom.Caption = RstPro("idcuenta")
        TxtCtaCom.Text = Busca_Codigo(RstPro("idcuenta"), "id", "cuenta", "con_planctas", "N", xCon)
        LblNomCtaCom.Caption = Busca_Codigo(RstPro("idcuenta"), "id", "descripcion", "con_planctas", "N", xCon)
    End If
    
    If NulosN(RstPro("idcuentaven")) <> 0 Then
        LbIdCuentaVen.Caption = RstPro("idcuentaven")
        TxtCtaVen.Text = Busca_Codigo(RstPro("idcuentaven"), "id", "cuenta", "con_planctas", "N", xCon)
        LblNomCtaVen.Caption = Busca_Codigo(RstPro("idcuentaven"), "id", "descripcion", "con_planctas", "N", xCon)
    End If
    
    If NulosN(RstPro("idnetonodomi")) <> 0 Then
        TxtIdNetoDomic.Text = RstPro("idtipven")
        LblNetoDomic.Caption = Busca_Codigo(RstPro("idnetonodomi"), "id", "descripcion", "mae_netonodomiciliado", "N", xCon)
    End If
    
    If NulosN(RstPro("idimpsel")) <> 0 Then
        TxtIdSelectivo.Text = RstPro("idimpsel")
        LblSelectivo.Caption = Busca_Codigo(RstPro("idimpsel"), "id", "descripcion", "mae_impuestos", "N", xCon)
    End If

    If NulosN(RstPro("idtipcom")) <> 0 Then
        TxtIdTipCom.Text = RstPro("idtipcom")
        LblIdTipCom.Caption = Busca_Codigo(RstPro("idtipcom"), "id", "descripcion", "mae_tipocompra", "N", xCon)
    End If

    If NulosN(RstPro("idtipven")) <> 0 Then
        TxtIdTipVen.Text = RstPro("idtipven")
        LblIdTipVen.Caption = Busca_Codigo(RstPro("idtipven"), "id", "descripcion", "mae_tipoventa", "N", xCon)
    End If
    
    LblTipoPro.Caption = NulosC(RstPro("desctippro"))
    LblClase.Caption = NulosC(RstPro("descclase"))
    LblSubClase.Caption = NulosC(RstPro("descsubclase"))
    LblFamilia.Caption = NulosC(RstPro("descfam"))
    LblDescUnidad.Caption = NulosC(RstPro("descunimed"))
    LblMoneda.Caption = RstPro("descmon")
    LblIdUnidad.Caption = RstPro("idunimed")
    
    ' cargamos los datos del centro de costos
    If xDeDonde = 1 Then
        ' si el llamado de la libreria viene de inventario grabamos los datos del centro de costo
        Dim Rst As New ADODB.Recordset
        Dim A As Integer
    
        Fg1.Rows = 1
        RST_Busq Rst, "SELECT alm_invencencos.idpro, con_centrocosto.id, con_centrocosto.descripcion, alm_invencencos.imppor, " _
            & " con_centrocosto.codigo FROM alm_invencencos LEFT JOIN con_centrocosto ON alm_invencencos.idcencos = con_centrocosto.id " _
            & " WHERE (((alm_invencencos.idpro)= " & RstPro("id") & "))", xCon
        
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                
                Fg1.TextMatrix(A, 1) = NulosN(Rst("codigo"))
                Fg1.TextMatrix(A, 2) = NulosC(Rst("descripcion"))
                Fg1.TextMatrix(A, 3) = Format(Rst("imppor"), "0.00")
                Fg1.TextMatrix(A, 4) = NulosN(Rst("id"))
                Rst.MoveNext
                If Rst.EOF = True Then
                    Exit For
                End If
            Next A
        End If
        Set Rst = Nothing
    End If
    
    
    ' MOSTRAMOS LOS PREFIJOSPARA LOS CODIGOS
    LblPrefijo.Caption = Busca_Codigo(NulosN(RstPro("tippro")), "id", "prefijo", "mae_tipoproducto", "N", xCon)
    LblPrefijo1.Caption = NulosC(Busca_Codigo(NulosN(RstPro("idfam")), "id", "prefijo", "mae_familia", "N", xCon))
    LblPrefijo2.Caption = NulosC(Busca_Codigo(NulosN(RstPro("idfam")), "id", "prefijo", "mae_clase", "N", xCon))
    LblPrefijo3.Caption = NulosC(Busca_Codigo(NulosN(RstPro("idfam")), "id", "prefijo", "mae_subclase", "N", xCon))
    
    ' CARGAMOS LAS FOTOS
    
    If xDeDonde = 1 Then
        Dim xRuta As String
        xRuta = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTAAR", "RUTAS")
        xRuta = xRuta + "items\"
        
        Set Rst = Nothing
        RST_Busq Rst, "SELECT * FROM alm_inventariofoto WHERE idalm = " & RstPro("id") & " ORDER BY id", xCon
        Fg2.Rows = 1
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                Fg2.Rows = Fg2.Rows + 1
                Fg2.TextMatrix(A, 1) = NulosC(xRuta) & Rst("archivo")
                Fg2.TextMatrix(A, 3) = Rst("id")
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
        If Fg2.Rows = 1 Then
            Image1.Picture = LoadPicture("")
        Else
            Image1.Picture = LoadPicture(NulosC(Fg2.TextMatrix(1, 1)))
        End If
    End If
End Sub

Private Sub TxtUnidad_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusUnidad_Click
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : VALIDAR_DATA
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VALIDA LOS DATOS INGRESADO PARA EL CODIGO AUTOGENERADO, DEVUELVE VERDADERO
'*                    SI LOS DATOS INGRESADOS SON LOS CORRECTOS
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Tipo      |  INTEGER          |  ESPECIFICA EL TIPO DE NIVEL QUE SE ESTA
'*                                                     EVALUANDO
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Private Function VALIDAR_DATA(TIPO As Integer) As Boolean
Select Case TIPO
    Case 0 '--TIPO DE ITEM
    
    Case 1 '--FAMILIA
        If NulosN(Trim(TxtTipPro.Text)) = 0 Then
            MsgBox "Seleccione el Tipo de Item" + vbCr + "Luego Continue Seleccionando la Familia", vbExclamation, xTitulo
            TabOne2.CurrTab = 0
            TxtTipPro.SetFocus
            Exit Function
        End If
    Case 2 '--CLASE
        If NulosN(Trim(TxtTipPro.Text)) = 0 Then
            MsgBox "Seleccione el Tipo de Item" + vbCr + "Luego Continue Seleccionando la Familia", vbExclamation, xTitulo
            TabOne2.CurrTab = 0
            TxtTipPro.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdFamilia.Text)) = 0 Then
            MsgBox "Seleccione la Familia" + vbCr + "Luego Continue Seleccionando la Clase", vbExclamation, xTitulo
            TabOne2.CurrTab = 0
            TxtIdFamilia.SetFocus
            Exit Function
        End If
    Case 3 '--SUB CLASE
        If NulosN(Trim(TxtTipPro.Text)) = 0 Then
            MsgBox "Seleccione el Tipo de Item" + vbCr + "Luego Continue Seleccionando la Familia", vbExclamation, xTitulo
            TabOne2.CurrTab = 0
            TxtTipPro.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdFamilia.Text)) = 0 Then
            MsgBox "Seleccione la Familia" + vbCr + "Luego Continue Seleccionando la Clase", vbExclamation, xTitulo
            TabOne2.CurrTab = 0
            TxtIdFamilia.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdClase.Text)) = 0 Then
            MsgBox "Seleccione la Clase" + vbCr + "Luego Continue Seleccionando la Sub Clase", vbExclamation, xTitulo
            TabOne2.CurrTab = 0
            TxtIdClase.SetFocus
            Exit Function
        End If
        
    Case 4 '--VALIDAR CUANDO SE GRABE O MODIFIQUE EL REGISTRO
        If NulosN(Trim(TxtTipPro.Text)) = 0 Then
            MsgBox "Seleccione el Tipo de Item" + vbCr + "Luego Continue Seleccionando la Familia", vbExclamation, xTitulo
            TabOne2.CurrTab = 0
            TxtTipPro.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdFamilia.Text)) = 0 Then
            MsgBox "Seleccione la Familia" + vbCr + "Luego Continue Seleccionando la Clase", vbExclamation, xTitulo
            TabOne2.CurrTab = 0
            TxtIdFamilia.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdClase.Text)) = 0 Then
            MsgBox "Seleccione la Clase" + vbCr + "Luego Continue Seleccionando la Sub Clase", vbExclamation, xTitulo
            TabOne2.CurrTab = 0
            TxtIdClase.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdSubClase.Text)) = 0 Then
            MsgBox "Seleccione la Sub Clase" + vbCr + "Luego Continue", vbExclamation, xTitulo
            TabOne2.CurrTab = 0
            TxtIdSubClase.SetFocus
            Exit Function
        End If
    End Select
    VALIDAR_DATA = True
End Function

'*****************************************************************************************************
'* Nombre           : pExportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A EXCEL LOS ITEMS REGISTRADOS EN EL SISTEMA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    '--se dajade usar esto por que no exporta otros datos 14/10/10
''    Dim xFun As New eps_librerias.FuncionesDGrid
''    xFun.ExportarDGExcel RstPro, Dg1, "LISTA DE ITEMS"
''    Set xFun = Nothing

    TabOne1.CurrTab = 0

    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset

    Dim xCampos(28, 3) As String

    ' CONFIGURAMOS LOS DATOS QUE SE VAN A EXPORTAR
    ' 0::Nombre a Mostrar;
    ' 1::nombre de Campo del Rst;
    ' 2::alineacion(0::derecha, 1::centro, 2::izquierda);
    ' 3::ancho de columna
    ' -obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Código":           xCampos(0, 1) = "codpro":           xCampos(0, 2) = 0:  xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = 0:  xCampos(1, 3) = "4500"

    xCampos(2, 0) = "Tipo Producto":    xCampos(2, 1) = "tipoproducto":     xCampos(2, 2) = 0:  xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Familia":          xCampos(3, 1) = "familia":          xCampos(3, 2) = 0:  xCampos(3, 3) = "1400"
    xCampos(4, 0) = "Clase":            xCampos(4, 1) = "clase":            xCampos(4, 2) = 0:  xCampos(4, 3) = "900"
    xCampos(5, 0) = "SubClase":         xCampos(5, 1) = "subclase":         xCampos(5, 2) = 0:  xCampos(5, 3) = "1057"

    xCampos(6, 0) = "U.M.":             xCampos(6, 1) = "abreunimed":       xCampos(6, 2) = 1:  xCampos(6, 3) = "450"
    xCampos(7, 0) = "M":                xCampos(7, 1) = "simbolo":          xCampos(7, 2) = 1:  xCampos(7, 3) = "450"
    xCampos(8, 0) = "Prec. Ini":        xCampos(8, 1) = "preini":           xCampos(8, 2) = 2:  xCampos(8, 3) = "1100"
    xCampos(9, 0) = "Prec. Unit":       xCampos(9, 1) = "preuni":           xCampos(9, 2) = 2:  xCampos(9, 3) = "1100"
    xCampos(10, 0) = "Stock Inicial":   xCampos(10, 1) = "stckini":         xCampos(10, 2) = 2: xCampos(10, 3) = "1100"
    xCampos(11, 0) = "Stock Actual":    xCampos(11, 1) = "stckact":         xCampos(11, 2) = 2: xCampos(11, 3) = "1100"
    xCampos(12, 0) = "Stock Min":       xCampos(12, 1) = "stckmin":         xCampos(12, 2) = 2: xCampos(12, 3) = "1100"
    xCampos(13, 0) = "Stock Max":       xCampos(13, 1) = "stckmax":         xCampos(13, 2) = 2: xCampos(13, 3) = "1100"

    xCampos(14, 0) = "Tipo Movimiento": xCampos(14, 1) = "tipomovimiento":  xCampos(14, 2) = 0: xCampos(14, 3) = "1580"
    xCampos(15, 0) = "Neto Domicilado": xCampos(15, 1) = "netodomiciliado": xCampos(15, 2) = 0: xCampos(15, 3) = "1540"

    xCampos(16, 0) = "Nº Cta Compra":   xCampos(16, 1) = "ctanumcompra":    xCampos(16, 2) = 0: xCampos(16, 3) = "1400"
    xCampos(17, 0) = "Cta Compra":      xCampos(17, 1) = "ctacompra":       xCampos(17, 2) = 0: xCampos(17, 3) = "2400"
    xCampos(18, 0) = "Nº Cta Venta":    xCampos(18, 1) = "ctanumventa":     xCampos(18, 2) = 0: xCampos(18, 3) = "1400"
    xCampos(19, 0) = "Cta Venta":       xCampos(19, 1) = "ctaventa":        xCampos(19, 2) = 0: xCampos(19, 3) = "2400"

    xCampos(20, 0) = "Retención":       xCampos(20, 1) = "retencion":       xCampos(20, 2) = 0: xCampos(20, 3) = "1200"
    xCampos(21, 0) = "Detracción":      xCampos(21, 1) = "detraccion":      xCampos(21, 2) = 0: xCampos(21, 3) = "1200"
    xCampos(22, 0) = "Percepción":      xCampos(22, 1) = "percepcion":      xCampos(22, 2) = 0: xCampos(22, 3) = "1200"
    xCampos(23, 0) = "ISC":             xCampos(23, 1) = "isc":             xCampos(23, 2) = 0: xCampos(23, 3) = "1200"

    xCampos(24, 0) = "IGV Compra":      xCampos(24, 1) = "igvcompra":       xCampos(24, 2) = 0:  xCampos(24, 3) = "2700"
    xCampos(25, 0) = "IGV Venta":       xCampos(25, 1) = "igvventa":        xCampos(25, 2) = 0:  xCampos(25, 3) = "2700"
    xCampos(26, 0) = "Activo":          xCampos(26, 1) = "xactivo":          xCampos(26, 2) = 1:  xCampos(26, 3) = "600"

    xCampos(27, 0) = "Cod. Cen Costo":  xCampos(27, 1) = "codcencos":       xCampos(27, 2) = 0:  xCampos(27, 3) = "1000"
    xCampos(28, 0) = "Descripción":     xCampos(28, 1) = "descencos":       xCampos(28, 2) = 0:  xCampos(28, 3) = "3000"

    ' CARGAMOS LOS DATOS QUE SE VAN A EXPORTAR
    nSQL = "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev AS abreunimed, mae_moneda.simbolo, alm_inventario.preini,alm_inventario.preuni, alm_inventario.stckini, " _
        & " alm_inventario.stckact, alm_inventario.stckmin, alm_inventario.stckmax, mae_tipoproducto.descripcion AS tipoproducto, mae_familia.descripcion AS familia, " _
        & " mae_clase.descripcion AS clase, mae_subclase.descripcion AS subclase, con_planctas.cuenta AS ctanumcompra, con_planctas.descripcion AS ctacompra, " _
        & " con_planctas_1.cuenta AS ctanumventa, con_planctas_1.descripcion AS ctaventa, mae_retencion.descripcion AS retencion, mae_detraccion.descripcion AS detraccion, " _
        & " mae_percepcion.descripcion AS percepcion, mae_impuestos.descripcion AS ISC, mae_tipocompra.descripcion AS IGVCompra, mae_tipoventa.descripcion AS IGVVenta, " _
        & " mae_tipomovimiento.descripcion AS tipomovimiento, mae_netonodomiciliado.descripcion AS netodomiciliado, IIf([alm_inventario].[activo]=-1,'Si','No') AS Activo, " _
        & " alm_inventario.idmon, alm_inventario.idunimed, alm_inventario.tippro, alm_inventario.idfam, alm_inventario.idclas, alm_inventario.idsubclas, alm_inventario.idcuenta AS idctacompra, " _
        & " alm_inventario.idcuentaven AS idctaventa, alm_inventario.idret, alm_inventario.iddet, alm_inventario.idper, alm_inventario.idimpsel, alm_inventario.idtipcom, " _
        & " alm_inventario.idtipven, con_centrocosto.codigo AS codcencos, con_centrocosto.descripcion AS descencos, iif(alm_inventario.activo =-1,'Activo','De Baja') as xactivo FROM ((mae_unidades RIGHT JOIN (mae_tipoproducto " _
        & " RIGHT JOIN (((mae_tipoventa RIGHT JOIN (mae_tipocompra RIGHT JOIN ((mae_retencion RIGHT JOIN (mae_percepcion RIGHT JOIN (mae_detraccion RIGHT JOIN " _
        & " ((con_planctas RIGHT JOIN ((((alm_inventario LEFT JOIN mae_moneda ON alm_inventario.idmon = mae_moneda.id) LEFT JOIN mae_clase ON alm_inventario.idclas = mae_clase.id) " _
        & " LEFT JOIN mae_subclase ON alm_inventario.idsubclas = mae_subclase.id) LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) " _
        & " ON con_planctas.id = alm_inventario.idcuenta) LEFT JOIN con_planctas AS con_planctas_1 ON alm_inventario.idcuentaven = con_planctas_1.id) " _
        & " ON mae_detraccion.id = alm_inventario.iddet) ON mae_percepcion.id = alm_inventario.idper) ON mae_retencion.id = alm_inventario.idret) " _
        & " LEFT JOIN mae_impuestos ON alm_inventario.idimpsel = mae_impuestos.id) ON mae_tipocompra.id = alm_inventario.idtipcom) ON mae_tipoventa.id = alm_inventario.idtipven) " _
        & " LEFT JOIN mae_tipomovimiento ON alm_inventario.tipo = mae_tipomovimiento.id) LEFT JOIN mae_netonodomiciliado ON alm_inventario.idnetonodomi = mae_netonodomiciliado.id) " _
        & " ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) LEFT JOIN alm_invencencos ON alm_inventario.id = alm_invencencos.idpro) " _
        & " LEFT JOIN con_centrocosto ON alm_invencencos.idcencos = con_centrocosto.id ORDER BY alm_inventario.descripcion"

    RST_Busq RstTmp, nSQL, xCon
    ' EXPORTAMOS LOS DATOS
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "LISTADO DE ITEMS", "", "", "Listado de Items ", RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
End Sub

Private Sub CmdNetoDomic_Click()
    If xDeDonde = 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    'Dim xform As New EPS_Buscar.Buscar
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_netonodomiciliado.* FROM mae_netonodomiciliado"
    
    xform.Titulo = "Buscando Percepcion"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdNetoDomic.Text = xRs("id")
        LblNetoDomic.Caption = NulosC(xRs("descripcion"))
        TxtIdSelectivo.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub TxtIdNetoDomic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdNetoDomic_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 32 Then
        If QueHace = 3 Then Exit Sub
        TxtIdNetoDomic.Text = ""
        LblNomCtaVen.Caption = ""
    End If
    
    If KeyCode = 116 Then
        CmdIdRet_Click
    End If
End Sub

Private Sub TxtIdNetoDomic_Validate(Cancel As Boolean)
    If Trim(TxtIdNetoDomic.Text) = "" Then
        LblNetoDomic.Caption = ""
        Exit Sub
    End If
    LblNomCtaVen.Caption = Busca_Codigo(NulosN(TxtIdNetoDomic.Text), "id", "descripcion", "mae_netonodomiciliado", "N", xCon)
End Sub
