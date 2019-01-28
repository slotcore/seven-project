VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "aspatextboxfecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVentas 
   Caption         =   "Ventas - Ingreso de Ventas"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   2025
      Left            =   12435
      TabIndex        =   100
      Top             =   5385
      Visible         =   0   'False
      Width           =   6705
      Begin VB.TextBox TxtNewSaldo2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1320
         TabIndex        =   112
         Text            =   "TxtNewSaldo2"
         Top             =   1515
         Width           =   1395
      End
      Begin VB.TextBox TxtSaldo2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1320
         TabIndex        =   110
         Text            =   "TxtSaldo2"
         Top             =   1200
         Width           =   1395
      End
      Begin VB.TextBox TxtCliente2 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1320
         TabIndex        =   108
         Text            =   "TxtCliente2"
         Top             =   780
         Width           =   5280
      End
      Begin VB.TextBox TxtNumDoc2 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1320
         TabIndex        =   104
         Text            =   "TxtNumDoc2"
         Top             =   465
         Width           =   2055
      End
      Begin VB.Frame Frame9 
         Height          =   870
         Left            =   3240
         TabIndex        =   114
         Top             =   1050
         Width           =   3375
         Begin VB.CommandButton Command2 
            Height          =   630
            Left            =   1710
            Picture         =   "FrmVentas.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   180
            Width           =   750
         End
         Begin VB.CommandButton Command1 
            Height          =   630
            Left            =   930
            Picture         =   "FrmVentas.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   180
            Width           =   750
         End
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Saldo"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   113
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   111
         Top             =   1245
         Width           =   405
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   109
         Top             =   825
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   106
         Top             =   510
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizar Saldo del Documento"
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
         Left            =   225
         TabIndex        =   102
         Top             =   90
         Width           =   2730
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   300
         Left            =   30
         Top             =   45
         Width           =   6615
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   1995
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   6690
         X2              =   6690
         Y1              =   15
         Y2              =   2010
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   6690
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   6675
         Y1              =   2010
         Y2              =   2010
      End
   End
   Begin VB.Frame Fraseldoc 
      BorderStyle     =   0  'None
      Caption         =   "Documento"
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
      Height          =   2250
      Left            =   12435
      TabIndex        =   60
      Top             =   -360
      Visible         =   0   'False
      Width           =   5565
      Begin VB.CommandButton CmdBusAlmacen2 
         Height          =   240
         Left            =   3180
         Picture         =   "FrmVentas.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   525
         Width           =   240
      End
      Begin VB.TextBox TxtAlmacen2 
         Height          =   300
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   95
         Text            =   "TxtAlmacen2"
         Top             =   495
         Width           =   2025
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmiAnul 
         Height          =   300
         Left            =   1425
         TabIndex        =   99
         Top             =   1125
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.TextBox TxtNumDocGen 
         Height          =   300
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   103
         Text            =   "TxtNumDocGen"
         Top             =   1755
         Width           =   1335
      End
      Begin VB.CommandButton CmdBusSerGen 
         Height          =   240
         Left            =   2490
         Picture         =   "FrmVentas.frx":0746
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1485
         Width           =   240
      End
      Begin VB.TextBox TxtNumSerGen 
         Height          =   300
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   101
         Text            =   "TxtNumSerGen"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton CmdBusTipDocGen 
         Height          =   240
         Left            =   5160
         Picture         =   "FrmVentas.frx":0878
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   840
         Width           =   240
      End
      Begin VB.TextBox TxtIdDocGen 
         Height          =   300
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   97
         Text            =   "TxtIdDocGen"
         Top             =   810
         Width           =   4005
      End
      Begin VB.Frame Frame7 
         Height          =   1020
         Left            =   3030
         TabIndex        =   93
         Top             =   1065
         Width           =   2400
         Begin VB.CommandButton cmdsalirseldoc 
            Height          =   600
            Left            =   1200
            Picture         =   "FrmVentas.frx":09AA
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   270
            Width           =   720
         End
         Begin VB.CommandButton cmdokseldoc 
            Height          =   600
            Left            =   450
            Picture         =   "FrmVentas.frx":0CB4
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   270
            Width           =   720
         End
      End
      Begin VB.Label LblidAlmacen2 
         AutoSize        =   -1  'True
         Caption         =   "LblidAlmacen2"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3765
         TabIndex        =   98
         Top             =   585
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   94
         Top             =   510
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Documento"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   92
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   165
         TabIndex        =   74
         Top             =   1785
         Width           =   1050
      End
      Begin VB.Label LblIdDocumentoGen 
         AutoSize        =   -1  'True
         Caption         =   "LblIdDocumentoGen"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3765
         TabIndex        =   73
         Top             =   390
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5565
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nº Serie"
         Height          =   195
         Left            =   165
         TabIndex        =   71
         Top             =   1470
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emisión de Documentos Anulados"
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
         Left            =   135
         TabIndex        =   70
         Top             =   105
         Width           =   2880
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Left            =   45
         Top             =   45
         Width           =   5475
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   5550
         X2              =   5550
         Y1              =   15
         Y2              =   2220
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   2235
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   5550
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   68
         Top             =   840
         Width           =   1185
      End
   End
   Begin VB.Frame Fradocsproc 
      BorderStyle     =   0  'None
      Caption         =   "2"
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
      Height          =   3285
      Left            =   12540
      TabIndex        =   22
      Top             =   1980
      Visible         =   0   'False
      Width           =   3705
      Begin VB.CommandButton cmdEliminarOKdocsproc 
         Height          =   630
         Left            =   1380
         Picture         =   "FrmVentas.frx":0FBE
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2535
         Width           =   765
      End
      Begin VB.CommandButton cmdOKdocsproc 
         Height          =   630
         Left            =   600
         Picture         =   "FrmVentas.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2535
         Width           =   750
      End
      Begin VB.CommandButton cmdSalirdocsproc 
         Height          =   630
         Left            =   2355
         Picture         =   "FrmVentas.frx":13CA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2535
         Width           =   750
      End
      Begin VSFlex7Ctl.VSFlexGrid fgdocsproc 
         Height          =   1950
         Left            =   150
         TabIndex        =   23
         Top             =   450
         Width           =   3405
         _cx             =   6006
         _cy             =   3440
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmVentas.frx":16D4
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
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   3660
         Y1              =   3270
         Y2              =   3270
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   3690
         X2              =   3690
         Y1              =   15
         Y2              =   3285
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   3675
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   -15
         Y2              =   3255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos Facturados en pantalla"
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
         Left            =   195
         TabIndex        =   76
         Top             =   90
         Width           =   3075
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Left            =   45
         Top             =   45
         Width           =   3615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7755
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
            Picture         =   "FrmVentas.frx":1759
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":1C9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":202F
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":21B3
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":2607
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":271F
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":2C63
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":31A7
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":32BB
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":33CF
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":3823
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":398F
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas.frx":3ED7
            Key             =   "IMG12"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7470
      Left            =   -15
      TabIndex        =   15
      Top             =   375
      Width           =   11895
      _cx             =   20981
      _cy             =   13176
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
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7050
         Left            =   12540
         TabIndex        =   30
         Top             =   375
         Width           =   11805
         Begin VB.Frame Frame6 
            Height          =   3000
            Left            =   9195
            TabIndex        =   77
            Top             =   3375
            Width           =   2550
            Begin VSFlex7Ctl.VSFlexGrid Fg4 
               Height          =   2445
               Left            =   60
               TabIndex        =   78
               Top             =   420
               Width           =   2430
               _cx             =   4286
               _cy             =   4313
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
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmVentas.frx":41F1
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
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Documentos Adjuntos"
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
               Height          =   210
               Left            =   60
               TabIndex        =   79
               Top             =   180
               Width           =   1860
            End
         End
         Begin VB.CommandButton CmdBusDocRef2 
            Height          =   240
            Left            =   9405
            Picture         =   "FrmVentas.frx":4276
            Style           =   1  'Graphical
            TabIndex        =   126
            Top             =   2370
            Width           =   240
         End
         Begin VB.CommandButton CmdBusDocRef 
            Height          =   240
            Left            =   8040
            Picture         =   "FrmVentas.frx":43A8
            Style           =   1  'Graphical
            TabIndex        =   123
            Top             =   1740
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.TextBox TxtNumDocRef 
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   13
            Text            =   "TxtNumDocRef"
            Top             =   2340
            Width           =   3390
         End
         Begin VB.CommandButton CmdBusIdTipDocRef 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmVentas.frx":44DA
            Style           =   1  'Graphical
            TabIndex        =   120
            Top             =   2370
            Width           =   240
         End
         Begin VB.Frame Frame10 
            Height          =   465
            Left            =   9630
            TabIndex        =   117
            Top             =   1440
            Width           =   2115
            Begin VB.Label LblPeriodo2 
               Alignment       =   2  'Center
               Caption         =   "LblPeriodo2"
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
               Height          =   300
               Left            =   120
               TabIndex        =   118
               Top             =   120
               Width           =   1860
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Motivo de la Nota de Credito ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   645
            Left            =   4935
            TabIndex        =   90
            Top             =   3990
            Visible         =   0   'False
            Width           =   3435
            Begin VB.CommandButton CmdMotNotCre 
               Height          =   240
               Left            =   3060
               Picture         =   "FrmVentas.frx":460C
               Style           =   1  'Graphical
               TabIndex        =   136
               Top             =   300
               Width           =   240
            End
            Begin VB.TextBox TxtDocRef 
               Height          =   300
               Left            =   75
               MaxLength       =   50
               TabIndex        =   91
               Text            =   "TxtDocRef"
               Top             =   255
               Width           =   3270
            End
            Begin VB.Label LblIdConNC 
               Caption         =   "LblIdConNC"
               ForeColor       =   &H000000C0&
               Height          =   210
               Left            =   1815
               TabIndex        =   137
               Top             =   30
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin VB.CommandButton CmdBusAlm 
            Height          =   240
            Left            =   6720
            Picture         =   "FrmVentas.frx":473E
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   795
            Width           =   240
         End
         Begin VB.Frame FraRetencion 
            Caption         =   "[ Retención de 4ta Categoria ]"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   645
            Left            =   5175
            TabIndex        =   64
            Top             =   2820
            Visible         =   0   'False
            Width           =   3435
            Begin VB.OptionButton OptSi 
               Caption         =   "Afecto"
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
               Left            =   525
               TabIndex        =   66
               Top             =   285
               Width           =   885
            End
            Begin VB.OptionButton OptNo 
               Caption         =   "No Afecto"
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
               Left            =   1800
               TabIndex        =   65
               Top             =   285
               Width           =   1440
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "[ Opciones de Descuento]"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   780
            Left            =   8655
            TabIndex        =   75
            Top             =   2685
            Width           =   3105
            Begin VB.CheckBox Check1 
               Caption         =   "Ingresar Neto"
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
               Left            =   705
               TabIndex        =   88
               Top             =   525
               Width           =   1500
            End
            Begin VB.OptionButton OptDes1 
               Caption         =   "Porcentaje"
               Height          =   195
               Left            =   420
               TabIndex        =   87
               Top             =   210
               Width           =   1215
            End
            Begin VB.OptionButton OptDes2 
               Caption         =   "Valor"
               Height          =   195
               Left            =   1965
               TabIndex        =   86
               Top             =   210
               Width           =   735
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   15
               X2              =   3060
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000003&
               Index           =   0
               X1              =   15
               X2              =   3060
               Y1              =   465
               Y2              =   465
            End
         End
         Begin VB.CommandButton CmdBusVen 
            Height          =   240
            Left            =   6720
            Picture         =   "FrmVentas.frx":4870
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   1110
            Width           =   240
         End
         Begin VB.TextBox TxtIdVen 
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   5
            Text            =   "TxtIdVen"
            Top             =   1080
            Width           =   705
         End
         Begin VB.CommandButton CmdBusNumSer 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmVentas.frx":49A2
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   1740
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipItem 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmVentas.frx":4AD4
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   795
            Width           =   240
         End
         Begin VB.CommandButton CmdBusCli 
            Height          =   240
            Left            =   3075
            Picture         =   "FrmVentas.frx":4C06
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1425
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmVentas.frx":4D38
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   1110
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2730
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   8
            Text            =   "TxtNumDoc"
            Top             =   1710
            Width           =   1440
         End
         Begin VB.CommandButton CmdBusCondicion 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmVentas.frx":4E6A
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   2055
            Width           =   240
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   6720
            Picture         =   "FrmVentas.frx":4F9C
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   480
            Width           =   240
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   7
            Text            =   "TxtNumSer"
            Top             =   1710
            Width           =   915
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   4
            Text            =   "TxtTipDoc"
            Top             =   1080
            Width           =   915
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   6
            Text            =   "TxtNumRuc"
            Top             =   1395
            Width           =   1770
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "TxtIdMon"
            Top             =   450
            Width           =   705
         End
         Begin VB.TextBox TxtConPag 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   10
            Text            =   "TxtConPag"
            Top             =   2025
            Width           =   915
         End
         Begin VB.TextBox TxtTipItem 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   2
            Text            =   "TxtTipItem"
            Top             =   765
            Width           =   915
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
            Height          =   300
            Left            =   1575
            TabIndex        =   0
            Top             =   450
            Width           =   1260
            _ExtentX        =   2223
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
            Locked          =   -1  'True
            Valor           =   "03/01/2004"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchVen 
            Height          =   300
            Left            =   6285
            TabIndex        =   11
            Top             =   2025
            Width           =   1215
            _ExtentX        =   2143
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
            Locked          =   -1  'True
            Valor           =   "03/01/2004"
         End
         Begin VB.Frame Fratipven 
            Caption         =   "[ Tipo de Facturacion ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   645
            Left            =   60
            TabIndex        =   48
            Top             =   2820
            Width           =   5070
            Begin VB.OptionButton optconcotizacion 
               Caption         =   "Cotización"
               Enabled         =   0   'False
               Height          =   255
               Left            =   2310
               TabIndex        =   67
               Top             =   285
               Width           =   1215
            End
            Begin VB.CommandButton cmdagregardocs 
               Caption         =   "Adicionar"
               Enabled         =   0   'False
               Height          =   330
               Left            =   3600
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   51
               Top             =   210
               Width           =   1380
            End
            Begin VB.OptionButton optsinguia 
               Caption         =   "Sin Guia"
               Enabled         =   0   'False
               Height          =   270
               Left            =   135
               TabIndex        =   50
               Top             =   285
               Width           =   945
            End
            Begin VB.OptionButton optconguia 
               Caption         =   "Guia"
               Enabled         =   0   'False
               Height          =   270
               Left            =   1350
               TabIndex        =   49
               Top             =   285
               Value           =   -1  'True
               Width           =   825
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2910
            Left            =   60
            TabIndex        =   14
            Top             =   3465
            Width           =   11670
            _cx             =   20585
            _cy             =   5133
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   20
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmVentas.frx":50CE
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
         Begin VB.Frame Frame4 
            Height          =   750
            Left            =   60
            TabIndex        =   31
            Top             =   6300
            Width           =   11700
            Begin VB.CommandButton CmdPreHist 
               Caption         =   "Ver His. Precios"
               Enabled         =   0   'False
               Height          =   495
               Left            =   3645
               Style           =   1  'Graphical
               TabIndex        =   119
               Top             =   165
               Width           =   1170
            End
            Begin VB.CommandButton CmdSel 
               Caption         =   "&Seleccionar Item"
               Enabled         =   0   'False
               Height          =   495
               Left            =   2440
               Style           =   1  'Graphical
               TabIndex        =   89
               Top             =   165
               Width           =   1170
            End
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "&Eliminar Item"
               Enabled         =   0   'False
               Height          =   495
               Left            =   1235
               TabIndex        =   85
               Top             =   165
               Width           =   1170
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "&Agregar Item"
               Enabled         =   0   'False
               Height          =   495
               Left            =   30
               TabIndex        =   84
               Top             =   165
               Width           =   1170
            End
            Begin VB.TextBox txtisc 
               Alignment       =   1  'Right Justify
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
               Left            =   9180
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   36
               TabStop         =   0   'False
               Text            =   "TxtIsc"
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox txtinafecto 
               Alignment       =   1  'Right Justify
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
               Left            =   6285
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   35
               TabStop         =   0   'False
               Text            =   "TxtInafecto"
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox TxtBruto 
               Alignment       =   1  'Right Justify
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
               Left            =   5055
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   34
               TabStop         =   0   'False
               Text            =   "TxtBruto"
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox TxtIGV 
               Alignment       =   1  'Right Justify
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
               Left            =   7605
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   33
               TabStop         =   0   'False
               Text            =   "TxtIGV"
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
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
               Left            =   10380
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   32
               TabStop         =   0   'False
               Text            =   "TxtTotal"
               Top             =   360
               Width           =   1200
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   4920
               X2              =   4920
               Y1              =   90
               Y2              =   810
            End
            Begin VB.Line Line3 
               BorderColor     =   &H80000003&
               Index           =   0
               X1              =   4905
               X2              =   4905
               Y1              =   105
               Y2              =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "I.S.C."
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
               Index           =   3
               Left            =   9180
               TabIndex        =   42
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Inafecto"
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
               Index           =   1
               Left            =   6255
               TabIndex        =   41
               Top             =   120
               Width           =   1140
            End
            Begin VB.Label LblIgvTasa 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "LblIgvTasa"
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
               Left            =   7950
               TabIndex        =   40
               Top             =   120
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Bruto"
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
               Left            =   5070
               TabIndex        =   39
               Top             =   120
               Width           =   885
            End
            Begin VB.Label LblRotulo 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. (         )"
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
               Left            =   7560
               TabIndex        =   38
               Top             =   120
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total"
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
               Index           =   2
               Left            =   10365
               TabIndex        =   37
               Top             =   120
               Width           =   450
            End
         End
         Begin VB.TextBox TxtIdAlm 
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "TxtIdAlm"
            Top             =   765
            Width           =   705
         End
         Begin VB.TextBox TxtIdTipDoc 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "TxtIdTipDoc"
            Top             =   2340
            Width           =   915
         End
         Begin VB.TextBox TxtDocRefCredi 
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   9
            Text            =   "TxtDocRefCredi"
            Top             =   1710
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.Label lblReg 
            Alignment       =   1  'Right Justify
            Caption         =   "lblReg"
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
            Height          =   270
            Left            =   9555
            TabIndex        =   135
            Top             =   30
            Width           =   2190
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Condición de Pago"
            Height          =   195
            Index           =   4
            Left            =   105
            TabIndex        =   134
            Top             =   2070
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   133
            Top             =   1755
            Width           =   1275
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Documento"
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   132
            Top             =   1125
            Width           =   1410
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   7
            Left            =   105
            TabIndex        =   131
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión"
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   130
            Top             =   495
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Item"
            Height          =   195
            Index           =   6
            Left            =   105
            TabIndex        =   129
            Top             =   810
            Width           =   660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Doc. Ref."
            Height          =   195
            Index           =   8
            Left            =   105
            TabIndex        =   128
            ToolTipText     =   "Tipo de Documento de Referencia"
            Top             =   2385
            Width           =   1005
         End
         Begin VB.Label LblIdDocRef2 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDocRef2"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   9735
            TabIndex        =   127
            Top             =   2385
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Referente al Documento"
            Height          =   195
            Left            =   4455
            TabIndex        =   125
            Top             =   1755
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Label LblIdDocRef 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDocRef"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8340
            TabIndex        =   124
            Top             =   1755
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Ref."
            Height          =   195
            Index           =   9
            Left            =   5505
            TabIndex        =   122
            ToolTipText     =   "Documento de Referencia"
            Top             =   2385
            Width           =   690
         End
         Begin VB.Label LblDescTipDocRef 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescTipDocRef"
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
            Left            =   2520
            TabIndex        =   121
            Top             =   2340
            Width           =   2655
         End
         Begin VB.Label LblAlmacen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblAlmacen"
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
            Left            =   7020
            TabIndex        =   83
            Top             =   765
            Width           =   2655
         End
         Begin VB.Label LblIdAlmacen 
            AutoSize        =   -1  'True
            Caption         =   "LblIdAlmacen"
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   3780
            TabIndex        =   82
            Top             =   540
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   11
            Left            =   5580
            TabIndex        =   81
            Top             =   810
            Width           =   615
         End
         Begin VB.Label LblNomVen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomVen"
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
            Left            =   7020
            TabIndex        =   63
            Top             =   1080
            Width           =   4710
         End
         Begin VB.Label Lblvendedor 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   5505
            TabIndex        =   62
            Top             =   1125
            Width           =   690
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Ventas"
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
            TabIndex        =   59
            Top             =   45
            Width           =   11595
         End
         Begin VB.Label LblTipoItem 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipoItem"
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
            Left            =   2520
            TabIndex        =   18
            Top             =   765
            Width           =   2715
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Ven."
            Height          =   195
            Index           =   3
            Left            =   5505
            TabIndex        =   58
            ToolTipText     =   "Fecha de Vencimiento"
            Top             =   2070
            Width           =   690
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "T.C."
            Height          =   195
            Left            =   10125
            TabIndex        =   57
            Top             =   495
            Width           =   300
         End
         Begin VB.Label LblTipoCambio 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipoCambio"
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
            Left            =   10440
            TabIndex        =   56
            Top             =   450
            Width           =   1305
         End
         Begin VB.Label LblCondPag 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCondPag"
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
            Left            =   2520
            TabIndex        =   55
            Top             =   2025
            Width           =   2655
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
            Left            =   7020
            TabIndex        =   20
            Top             =   450
            Width           =   2655
         End
         Begin VB.Label LblNomCli 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomCli"
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
            Left            =   3375
            TabIndex        =   54
            Top             =   1395
            Width           =   4635
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
            Left            =   2520
            TabIndex        =   53
            Top             =   1080
            Width           =   2715
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2550
            Top             =   1815
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   5610
            TabIndex        =   19
            Top             =   495
            Width           =   585
         End
         Begin VB.Label LblIdCliente 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCliente"
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   2865
            TabIndex        =   52
            Top             =   510
            Visible         =   0   'False
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7050
         Left            =   45
         TabIndex        =   27
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6735
            Left            =   30
            TabIndex        =   16
            Top             =   300
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11880
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Reg"
            Columns(0).DataField=   "numreg1"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "TD"
            Columns(1).DataField=   "abrev"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numerodoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Emi"
            Columns(3).DataField=   "fchdoc1"
            Columns(3).NumberFormat=   "Short Date"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Ven."
            Columns(4).DataField=   "fchven1"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Cliente"
            Columns(5).DataField=   "nombre"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "M."
            Columns(6).DataField=   "simbolo"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Imp. Bru."
            Columns(7).DataField=   "impbru1"
            Columns(7).NumberFormat=   "0.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "I.G.V."
            Columns(8).DataField=   "impigv1"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Importe"
            Columns(9).DataField=   "imptotdoc1"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Saldo"
            Columns(10).DataField=   "impsal1"
            Columns(10).NumberFormat=   "0.00"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=661"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=582"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2566"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2487"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1773"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1693"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1773"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1693"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=4815"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=4736"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=609"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=529"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=1508"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1429"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1244"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1164"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(55)=   "Column(9).Width=1508"
            Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1429"
            Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(61)=   "Column(10).Width=1588"
            Splits(0)._ColumnProps(62)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(63)=   "Column(10)._WidthInPix=1508"
            Splits(0)._ColumnProps(64)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(65)=   "Column(10)._ColStyle=514"
            Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
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
            HeadLines       =   1
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
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=78,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=1"
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
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
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
            Left            =   9120
            TabIndex        =   28
            Top             =   30
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Ventas"
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
            TabIndex        =   29
            Top             =   45
            Width           =   11565
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   12840
         X2              =   24645
         Y1              =   375
         Y2              =   7425
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   21
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
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Factura"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Restaurar Factura"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Saldo del Documento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Anular Factura"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Factura"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Emitir Factura Anulada"
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
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Documento"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Documento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Exportar a Excel"
               EndProperty
            EndProperty
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
   Begin VB.Menu menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar Item            "
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Item                "
      End
      Begin VB.Menu menu1_4 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_5 
         Caption         =   "Ver Historico de Precios"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu menu2_1 
         Caption         =   "Agregar Documento"
      End
      Begin VB.Menu menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu2_3 
         Caption         =   "Eliminar Documento"
      End
   End
End
Attribute VB_Name = "FrmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstVent As New ADODB.Recordset
Dim QueHace As Integer
Dim TasaImpuesto As Double
Dim CaracteresNumericos As String
Dim SeEjecuto As Boolean
Dim ValTipCam As Double
Dim xIdCuenTasa As Integer  'codigo de la cuenta contable del impuesto
Dim xCuentaDoc As Integer   'codigo de la cuenta contable del documento
Dim Mostrando As Boolean
Dim swguiafact              '0 No se facturaron, 1 Se facturaron
Dim Agregando As Boolean    'para saber cuando se este agregando datos en el grid de productos
Dim xHorIni As Date


Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim mIdRegistro& '--identificador del registro
Dim mMesActivo As Integer
Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)

Sub ActivarEntorno()
    TabOne1.Enabled = Not TabOne1.Enabled
    Toolbar1.Enabled = Not Toolbar1.Enabled
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    If RstVent.RecordCount = 0 Then
        MsgBox "No hay documentos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    TabOne1.CurrTab = 0
    Rpta = MsgBox("¿ Esta seguro de eliminar el documento seleccionado ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & RstVent("id") & " AND idlib = 2 AND Iddoc = " & RstVent("tipdoc") & ""
               
        If RstVent("oriitem") = 1 Then
            'si el origen del item es igual a 1 Actualizamos el saldo del stock
            Call ActualizarStock("E", RstVent("id"))
        End If
        If RstVent("oriitem") = 2 Then
            'actualizamos a 0 el campo "iddocven" de la tabla vta_guia para poder facturarla con otro numero de factura
            xCon.Execute "UPDATE vta_guia SET vta_guia.iddocven = 0 WHERE (((vta_guia.iddocven)=" & RstVent("id") & "))"
        End If
        
        xCon.Execute "DELETE * FROM vta_ventas WHERE id = " & RstVent("id") & ""
        xCon.Execute "DELETE * FROM con_diario WHERE idlib = 2 AND idmov = " & RstVent("id") & ""
        
        MsgBox RstVent("nomdoc") & " se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstVent.Requery
        Dg1.Refresh
        If RstVent.RecordCount = 0 Then
            Rpta = MsgBox("No se han registrado movimientos en el periodo especificado, ¿ Desea agregar uno ahora ?", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstVent = Nothing
                Unload Me
            End If
        End If
    End If
End Sub

Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    Toolbar1.Buttons(11).Enabled = Not Toolbar1.Buttons(11).Enabled
    
    Toolbar1.Buttons(13).Enabled = Not Toolbar1.Buttons(13).Enabled
    Toolbar1.Buttons(15).Enabled = Not Toolbar1.Buttons(15).Enabled
End Sub

Sub RestaurarFactura()
    'Se restaura una factura anulada
    Dim Rpta As Integer
    
    Rpta = MsgBox("Esta seguro de restaurar la factura Nº " + RstVent("numser") & "-" & RstVent("numdoc"), vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption)
    If Rpta = vbYes Then
        xCon.Execute "UPDATE vta_ventas SET vta_ventas.Anulado = 0, " _
            & " vta_ventas.idcli = 1  " _
            & " WHERE vta_ventas.id =" & RstVent("id") & ""
        
        xCon.Execute "DELETE * FROM vta_ventasdet WHERE vta_ventasdet.idvta =" & RstVent("id") & ""
        RstVent.Requery
        Dg1.Refresh
        MsgBox "La factura se restauro con exito", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
    End If
End Sub

Sub ActualizarStock(TIPO As String, numid As Integer)
    'QueHace 1 ADICIONAR
    'QueHace 2 MODIFICAR
    
    'Tipo S= Salida
    'Tipo E= Extorno por Anular Guia , Eliminar Guia
    Dim RstDet As New ADODB.Recordset
    Dim Rstitem As New ADODB.Recordset
    Dim xcant As Double

   'SI NO ESTA EL ID DEl DOCUMENTO ES UN DOCUMENTO SIN GUIA DE REMISION
    RST_Busq RstDet, "SELECT vta_guia.* From vta_guia WHERE vta_guia.iddocven = " & numid & "", xCon

    If RstDet.RecordCount = 0 Then
        Set RstDet = Nothing

        RST_Busq RstDet, "SELECT vta_ventasdet.* FROM vta_ventasdet WHERE idvta = " & numid & "", xCon
        Do While Not RstDet.EOF
            RST_Busq Rstitem, "SELECT Alm_Inventario.* FROM ALM_Inventario WHERE id = " & RstDet("iditem") & "", xCon
            If Rstitem.RecordCount > 0 Then
                If TIPO = "S" Then
                    Rstitem("stckact") = Rstitem("stckact") - RstDet("canpro")
                ElseIf TIPO = "E" Then
                    Rstitem("stckact") = Rstitem("stckact") + RstDet("canpro")
                End If
                Rstitem.Update
            End If
            
            RstDet.MoveNext
        Loop
    End If
    
    Set RstDet = Nothing
    Set Rstitem = Nothing
End Sub

Sub Anular()
    Dim Rpta As Integer
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    Rpta = MsgBox("¿Esta seguro de anular " & RstVent("nomdoc") & " Nº " & RstVent("numser") & "-" & RstVent("numdoc") + "?", vbYesNo + vbDefaultButton1 + vbQuestion, Me.Caption)
    
    If Rpta = vbYes Then
        xCon.Execute "UPDATE vta_ventas SET vta_ventas.Anulado = -1, " _
            & " vta_ventas.impbru = 0, vta_ventas.impinaf = 0, vta_ventas.impigv = 0,  vta_ventas.impisc = 0,  " _
            & " vta_ventas.impotr = 0, vta_ventas.imptotdoc = 0,  vta_ventas.impsal = 0  " _
            & " WHERE vta_ventas.id = " & RstVent("id") & " "
        
        If RstVent("oriitem") = 1 Then
            'si el origen del item es igual a 1 Actualizamos el saldo del stock
            Call ActualizarStock("E", RstVent("id"))
        End If
        If RstVent("oriitem") = 2 Then
            'actualizamos a 0 el campo "iddocven" d ela tabla vta_guia para poder facturarla con otro numero de factura
            xCon.Execute "UPDATE vta_guia SET vta_guia.iddocven = 0 WHERE (((vta_guia.iddocven)=" & RstVent("id") & "))"
        End If
        
        xCon.Execute "DELETE * FROM vta_ventasdet WHERE vta_ventasdet.idvta = " & RstVent("id") & ""
                
        'ponemos el diario a valor 0
        RST_Busq Rst, "SELECT * FROM con_diario WHERE idlib = 2 AND idmov = " & RstVent("id") & "", xCon
        
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                Rst("impdebsol") = 0
                Rst("imphabsol") = 0
                Rst("impdebdol") = 0
                Rst("imphabdol") = 0
                
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
        Set Rst = Nothing
        
        MsgBox RstVent("nomdoc") & " se anulo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstVent.Requery
        Dg1.Refresh
    End If
End Sub

Sub Cancelar()
    Dim X As Integer
    Bloquea
    Fg1.ColComboList(1) = ""
    Label5.Caption = "Detalle de Venta"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
     
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
       
    'Colocamos en el campo estado 0  de la tabla guia que indica no  esta facturado
    If fgdocsproc.Rows - 1 > 0 Then
        If swguiafact = 0 Then
            For X = 1 To fgdocsproc.Rows - 1
                xCon.Execute " UPDATE vta_guia SET Vta_guia.Estado = 0 WHERE vta_guia.id = " & NulosN(fgdocsproc.TextMatrix(X, 1)) & ""
            Next
            fgdocsproc.Rows = 1
        End If
    End If
    swguiafact = 0
End Sub

Sub Nuevo()
    QueHace = 1
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Label5.Caption = "Agregando Venta"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    OptSi.Value = True
    Fg1.Rows = 1
    Fg4.Rows = 1
    optsinguia.Value = True
    optsinguia_Click
    OptDes1.Value = True

    TxtFchDoc.Valor = Format(Date, "dd/mm/yyyy")
    
    If Check1.Value = 1 Then Check1_Click
    xHorIni = Time
    TxtFchDoc.SetFocus
End Sub

Sub Modificar()
    If RstVent.RecordCount = 0 Then
        MsgBox "No hay Registros para Modificar", vbInformation, Me.Caption
        Exit Sub
    End If
    If NulosN(RstVent("anulado")) = -1 Then
        MsgBox "El Documento de Venta esta Anulado" & vbCr & "No se Puede Modificar", vbInformation, Me.Caption
        Exit Sub
    End If
   
    QueHace = 2
    
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    MuestraSegundoTab
    Label5.Caption = "Modificando Ventas"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    xHorIni = Time
    TxtFchDoc.SetFocus
End Sub

Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    
    If RstVent.RecordCount = 0 Then Exit Sub
    Blanquea
    lblReg.Caption = "Nº Reg. " & NulosC(RstVent("numreg1"))
    
    TxtTipItem.Text = NulosN(RstVent("idtipo"))
    TxtTipDoc.Text = NulosN(RstVent("tipdoc"))
    TxtTipDoc_Validate False
    
    TxtNumRuc.Text = NulosC(RstVent("numruc"))
    TxtNumSer.Text = NulosC(RstVent("numser"))
    TxtNumDoc.Text = NulosC(RstVent("numdoc"))
    If IsDate(RstVent("fchdoc")) = True Then TxtFchDoc.Valor = CDate(RstVent("fchdoc"))
    If IsDate(RstVent("fchven")) = True Then TxtFchVen.Valor = CDate(RstVent("fchven"))
    TxtConPag.Text = NulosN(RstVent("idconpag"))
    TxtIdMon.Text = NulosN(RstVent("idmon"))
    If RstVent("idven") <> 0 Then
        TxtIdVen.Text = NulosN(RstVent("idven"))
        
        Set Rst = BuscaConCriterio("SELECT vta_vendedores.*, UCase(pla_empleados!apepat)+' '+ UCase(pla_empleados!apemat)+', '+pla_empleados!nom AS apenom " _
            & " FROM vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id WHERE vta_vendedores.id = " & NulosN(TxtIdVen.Text) & "", xCon)
        
        If Rst.RecordCount <> 0 Then
            LblNomVen.Caption = Rst("apenom")
        End If
    End If
    
    If NulosN(RstVent("idtipdocref")) <> 0 Then
        TxtIdTipDoc.Text = NulosC(RstVent("idtipdocref"))
        LblDescTipDocRef.Caption = Busca_Codigo(NulosC(RstVent("idtipdocref")), "id", "descripcion", "mae_docreferencia", "N", xCon)
        
        If NulosN(RstVent("idtipdocref")) = 4 Then
            RST_Busq Rst, "SELECT var_ordendespacho.id, [var_ordendespacho]![anno] & [var_ordendespacho]![idaduana] & [var_ordendespacho]![idregimen] & [var_ordendespacho]![numdoc] AS numdoc" _
                & " From var_ordendespacho WHERE (((var_ordendespacho.id)=" & NulosN(RstVent("iddocref2")) & "))", xCon
        End If
        If Rst.RecordCount <> 0 Then
            TxtNumDocRef.Text = Rst("numdoc")
            LblIdDocRef2.Caption = Rst("id")
        End If
        Set Rst = Nothing
    End If
    
    
    TxtBruto.Text = Format(NulosN(RstVent("impbru")), FORMAT_MONTO)
    TxtIGV.Text = Format(NulosN(RstVent("impigv")), FORMAT_MONTO)
    TxtTotal.Text = Format(NulosN(RstVent("imptotdoc")), FORMAT_MONTO)
    txtinafecto.Text = Format(NulosN(RstVent("impinaf")), FORMAT_MONTO)
    txtisc.Text = Format(NulosN(RstVent("impisc")), FORMAT_MONTO)
    
    If NulosN(RstVent("idalm")) <> 0 Then
        TxtIdAlm.Text = Format(RstVent("idalm"), "0")
    Else
        TxtIdAlm.Text = ""
    End If
    
    LblTipoItem.Caption = NulosC(RstVent("desctipcom"))
    LblNomDoc.Caption = NulosC(RstVent("nomdoc"))
    LblNomCli.Caption = NulosC(RstVent("nombre"))
    LblCondPag.Caption = NulosC(RstVent("desccond"))
    TxtNumRuc.Text = NulosC(RstVent("numruc"))
    LblMoneda.Caption = NulosC(RstVent("descmon"))
    LblIdAlmacen.Caption = NulosN(RstVent("idalm"))
    LblAlmacen.Caption = Busca_Codigo(RstVent("idalm"), "id", "descripcion", "alm_almacenes", "N", xCon)
                
    LblIdCliente.Caption = RstVent("idcli")
    xIdCuenTasa = NulosN(RstVent("idcuenvta"))

    If RstVent("idmon") = 1 Then
        LblTipCam.Visible = False
        LblTipoCambio.Visible = False
    Else
        LblTipCam.Visible = True
        LblTipoCambio.Visible = True
        LblTipoCambio.Caption = RstVent("impven")
    End If
        
    If RstVent("oriitem") = 1 Then optsinguia.Value = True: optsinguia_Click
    If RstVent("oriitem") = 2 Then optconguia.Value = True: optconguia_Click
    If RstVent("oriitem") = 3 Then optconcotizacion.Value = True
    
    'MOSTRAMOS EL TIPO DE DESCUENTO APLICADO
    If RstVent("tipdes") = 1 Or NulosN(RstVent("tipdes")) = 0 Then OptDes1.Value = True
    If RstVent("tipdes") = 2 Then OptDes2.Value = True
        
    Frame5.Visible = False
    
    TxtIdTipDoc.Visible = True
    LblDescTipDocRef.Visible = True
    CmdBusIdTipDocRef.Visible = True
    
    TxtNumDocRef.Visible = True
    CmdBusDocRef2.Visible = True
    Label3(9).Visible = True
    Label3(8).Visible = True
    
    If NulosN(RstVent("tipdoc")) = 7 And NulosN(RstVent("iddocref")) <> 0 Then
        TxtIdTipDoc.Visible = False
        LblDescTipDocRef.Visible = False
        CmdBusIdTipDocRef.Visible = False
        
        TxtNumDocRef.Visible = False
        CmdBusDocRef2.Visible = False
        Label3(9).Visible = False
        Label3(8).Visible = False
        
        Frame5.Left = 5175
        Frame5.Top = 2820
        Frame5.Visible = True
        If NulosN(RstVent("idmotnotcre")) <> 0 Then
            LblIdConNC.Caption = NulosN(RstVent("idmotnotcre"))
            TxtDocRef.Text = Busca_Codigo(NulosN(RstVent("idmotnotcre")), "id", "descripcion", "vta_conceptonc", "N", xCon)
        End If
        
        'LA NOTA DE CREDITO HACE REFERENCIA A UNA FACTURA
        TxtDocRefCredi.Visible = True
        Label33.Visible = True
        CmdBusDocRef.Visible = True
        
        LblIdDocRef.Caption = NulosN(RstVent("iddocref"))
        TxtDocRefCredi.Text = Busca_Codigo(NulosN(RstVent("iddocref")), "id", "numser", "vta_ventas", "N", xCon) + "-" + Busca_Codigo(RstVent("iddocref"), "id", "numdoc", "vta_ventas", "N", xCon)
    Else
        TxtDocRefCredi.Visible = False
        Label33.Visible = False
        CmdBusDocRef.Visible = False
    End If
       
    
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer

    'CARGAMOS LAS GUIAS DE LAS FACTURAS
    If optconguia.Value = True Then
        RST_Busq RstDet, "SELECT vta_guia.id, mae_documento.abrev, [vta_guia]![numser]+'-'+[vta_guia]![numdoc] AS numdoc" _
            & " FROM vta_guia LEFT JOIN mae_documento ON vta_guia.tipdoc = mae_documento.id " _
            & " WHERE (((vta_guia.iddocven)=" & RstVent("id") & "))", xCon
        If RstDet.RecordCount <> 0 Then
            
            RstDet.MoveFirst
            For A = 1 To RstDet.RecordCount
                Fg4.Rows = Fg4.Rows + 1
                Fg4.TextMatrix(A, 1) = RstDet("numdoc")
                Fg4.TextMatrix(A, 2) = RstDet("abrev")
                Fg4.TextMatrix(A, 3) = RstDet("id")
                
                RstDet.MoveNext
                If RstDet.EOF = True Then
                    Exit For
                End If
            Next A
        End If
        Set RstDet = Nothing
    End If
     
    'CARGAMOS LOS ITEMS DE LA FACTURA
    Set RstDet = Nothing
    Mostrando = True

    RST_Busq RstDet, "SELECT vta_ventasdet.*, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuentaven, " _
        & " alm_inventario.idtipven FROM mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN vta_ventasdet " _
        & " ON alm_inventario.id = vta_ventasdet.iditem) ON mae_unidades.id = alm_inventario.idunimed " _
        & " WHERE (((vta_ventasdet.idvta)=" & RstVent("id") & "))", xCon
    
    If RstDet.RecordCount <> 0 Then
        Do While Not RstDet.EOF
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstDet("descripcion"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDet("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(RstDet("canpro"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(RstDet("preunibru"), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(RstDet("valdes"), "0.0000")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(RstDet("preuni"), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(RstDet("imptot"), "0.0000")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(RstDet("iditem"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(RstDet("idunimed"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(RstDet("idcuentaven"))
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(RstDet("idtipven"))
            RstDet.MoveNext
        Loop
    End If
    
    Set RstDet = Nothing
    Mostrando = False
    
    Set RstDet = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & NulosN(TxtIdMon) & " and tipope = -1", xCon)
    If RstDet.RecordCount = 1 Then
        xCuentaDoc = RstDet("idcuen")
    End If
    
    Set RstDet = BuscaConCriterio("SELECT mae_impuestos.tasa from mae_impuestos WHERE mae_impuestos.id = 1 ", xCon)
    If RstDet.RecordCount = 1 Then
        TasaImpuesto = NulosN(RstDet("tasa"))
        'LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) + "%"
        LblIgvTasa.Caption = Format(Trim(Str(TasaImpuesto)), "0.00")
    End If
    
    pGridConfigurar
    
    
    Set RstDet = Nothing
End Sub

Sub Bloquea()
    TxtTipItem.Locked = Not TxtTipItem.Locked
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    
    'If QueHace = 1 Then
        TxtNumSer.Locked = Not TxtNumSer.Locked
        TxtNumDoc.Locked = Not TxtNumDoc.Locked
    'End If
    
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
    TxtFchVen.Locked = Not TxtFchVen.Locked
    TxtConPag.Locked = Not TxtConPag.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtIdAlm.Locked = Not TxtIdAlm.Locked
    
    Frame3.Enabled = Not Frame3.Enabled
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
    CmdSel.Enabled = Not CmdSel.Enabled
    CmdPreHist.Enabled = Not CmdPreHist.Enabled
    
    optsinguia.Enabled = Not optsinguia.Enabled
    optconguia.Enabled = Not optconguia.Enabled
    optconcotizacion.Enabled = Not optconcotizacion.Enabled
    
    TxtDocRef.Locked = Not TxtDocRef.Locked
    TxtIdTipDoc.Locked = Not TxtIdTipDoc.Locked
    TxtNumDocRef.Locked = Not TxtNumDocRef.Locked
    
    TxtIdVen.Locked = Not TxtIdVen.Locked
End Sub

Sub Blanquea()
    lblReg.Caption = ""
    
    TxtTipItem.Text = ""
    TxtTipDoc.Text = ""
    TxtNumRuc.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtFchDoc.Valor = ""
    
    TxtFchVen.Valor = ""
    TxtConPag.Text = ""
    TxtIdMon.Text = ""
    TxtDocRef.Text = ""
    TxtDocRefCredi.Text = ""
    
    LblNomDoc.Caption = ""
    LblNomCli.Caption = ""
    LblCondPag.Caption = ""
    LblMoneda.Caption = ""
    LblIdCliente.Caption = ""
    LblTipoItem.Caption = ""
    TxtIdVen = ""
    LblNomVen = ""
    TxtIdAlm.Text = ""
    LblAlmacen.Caption = ""
    
    txtinafecto = ""
    txtisc = ""
    TxtBruto.Text = ""
    TxtIGV.Text = ""
    TxtTotal.Text = ""

    TxtDocRef.Text = ""
    TxtIdTipDoc.Text = ""
    TxtNumDocRef.Text = ""
        
    LblDescTipDocRef.Caption = ""
    LblIdDocRef.Caption = ""
    LblIdDocRef2.Caption = ""
    
    Fg4.Rows = 1
    Fg1.Rows = 1
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Fg1.ColWidth(14) = 1005
        Fg1.ColWidth(15) = 705
        If optconguia = True Then
            Fg1.ColWidth(1) = 2900 '5400 - 1710
        End If
        If optsinguia = True Then
            Fg1.ColWidth(1) = 5400 - 1710
        End If
    Else
        Fg1.ColWidth(14) = 0
        Fg1.ColWidth(15) = 0
        If optconguia = True Then
            Fg1.ColWidth(1) = 3500
        End If
        If optsinguia = True Then
            Fg1.ColWidth(1) = 5400
        End If
    End If
End Sub

Private Sub CmdAddItem_Click()
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Sub
    End If
    
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtTipItem.Text) = "" Then
        MsgBox "No ha especificado el tipo de item a buscar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipItem.SetFocus
        Exit Sub
    End If
    
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = "" Then
        Fg1.Col = 1
        Fg1.Row = Fg1.Rows - 1
'        Fg1.SetFocus
        Fg1_CellButtonClick Fg1.Rows - 1, 1
        Fg1.SetFocus
        Exit Sub
    End If
    
    Fg1.Rows = Fg1.Rows + 1
    '--agregando cantidad por defecto a 1 cuando es servcio
    If NulosN(TxtTipItem.Text) = 5 Then
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = 1
    End If
    
    With Fg1
        .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
    End With
    
    Fg1_CellButtonClick Fg1.Rows - 1, 1
    
    Fg1.SetFocus
End Sub

Private Sub cmdagregardocs_Click()
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Sub
    End If
    
    If NulosC(LblIdCliente.Caption) = "" Then
        MsgBox "No ha especificado el cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Sub
    End If
    
    If optconguia.Value = True Then
        CargarGuia
    End If
    If optconcotizacion.Value = True Then
        CargarCotizacion
    End If
End Sub

Sub CargarCotizacion()
    'Dim xfrm As New EPS_Buscar.Seleccion
    Dim xfrm As New eps_librerias.FormSeleccion
    Dim xCampos(4, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim xRs1 As New ADODB.Recordset
    
    xCampos(0, 0) = "Nº Documento":    xCampos(0, 1) = "numdoc":        xCampos(0, 2) = "1300":   xCampos(0, 3) = "C":     xCampos(0, 4) = "S"
    xCampos(1, 0) = "Fch. Giro":       xCampos(1, 1) = "fchdoc":        xCampos(1, 2) = "1000":   xCampos(1, 3) = "C":     xCampos(1, 4) = "N"
    xCampos(2, 0) = "Forma Pago":      xCampos(2, 1) = "descripcion":   xCampos(2, 2) = "1500":   xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "Moneda":          xCampos(3, 1) = "simbolo":       xCampos(3, 2) = "1200":   xCampos(3, 3) = "C":     xCampos(3, 4) = "N"

    xfrm.SQLCad = "SELECT vta_cotizacion.id, vta_cotizacion.numdoc, vta_cotizacion.idcli, vta_cotizacion.fchdoc, " _
        & " mae_condpago.descripcion, vta_cotizacion.impbru, mae_moneda.simbolo FROM (vta_cotizacion LEFT JOIN mae_condpago " _
        & " ON vta_cotizacion.idconpag = mae_condpago.id) LEFT JOIN mae_moneda ON vta_cotizacion.idmon = mae_moneda.id " _
        & " WHERE (((vta_cotizacion.idcli)=" & LblIdCliente.Caption & ") AND ((vta_cotizacion.idest)=2))"
        
    xfrm.Titulo = "Buscando Cotizaciones"
    
    Set xfrm.Coneccion = xCon
    Set xRs = xfrm.Seleccionar(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount = 0 Then
            Set xRs = Nothing
            Exit Sub
        End If
        Dim xCadWhere As String
        Dim A As Integer
        Dim Rst As New ADODB.Recordset
        
        Fg4.Rows = 1
        xRs.MoveFirst
        
        'CARGAMOS LOS DOCUMENTOS ADJUNTOS Y LO MOSTRAMOS EN LA LISTA DE "DOCUMENTOS ADJUNTOS"
        For A = 1 To xRs.RecordCount
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(A, 1) = xRs("numdoc")
            Fg4.TextMatrix(A, 2) = "" 'xRs("abrev")
            Fg4.TextMatrix(A, 3) = xRs("id")
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
        MostrarItemsCotizacion
    End If
End Sub

Sub MostrarItemsCotizacion()
    Dim A As Integer
    Dim xCadWhere  As String
    Dim Rst As New ADODB.Recordset
    xCadWhere = ""
    
    If Fg4.Rows = 1 Then
        Fg1.Rows = 1
        Exit Sub
    End If
    'CREAMOS LA SENTENCIA WHERE PARA LA CONSULTA SQL
    For A = 1 To Fg4.Rows - 1
        xCadWhere = xCadWhere + "(vta_cotizaciondet.idvta = " & Fg4.TextMatrix(A, 3) & ")"
        If A = Fg4.Rows - 1 Then Exit For
        xCadWhere = xCadWhere + " OR "
    Next A
    
    RST_Busq Rst, "SELECT vta_cotizaciondet.iditem AS id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, Sum(vta_cotizaciondet.canpro) AS campro, " _
        & " vta_cotizaciondet.iditem, alm_inventario.idtipven, alm_inventario.idcuentaven, vta_cotizaciondet.preuni " _
        & " FROM vta_cotizacion LEFT JOIN ((vta_cotizaciondet LEFT JOIN alm_inventario ON vta_cotizaciondet.iditem = alm_inventario.id) LEFT JOIN mae_unidades " _
        & " ON vta_cotizaciondet.idunimed = mae_unidades.id) ON vta_cotizacion.id = vta_cotizaciondet.idvta " _
        & " Where " + xCadWhere _
        & " GROUP BY vta_cotizaciondet.iditem, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idtipven, " _
        & " alm_inventario.idcuentaven, vta_cotizaciondet.preuni", xCon

    Fg1.Rows = 1
    
    Agregando = True
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Busca_Codigo(Rst("id"), "id", "stckact", "alm_inventario", "N", xCon)
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("descripcion") 'DESCRIPCION DEL PRODUCTO
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("abrev")       'ABREVIATURA
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(Rst("preuni"), "0.00")   'PRECIO UNITARIO
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Rst("campro")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(Rst("preuni") * Rst("campro"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Rst("id")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Rst("idcuentaven")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Rst("idtipven")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
    Agregando = False
    HallarTotal
End Sub

Sub CargarGuia()
    'Dim xfrm As New EPS_Buscar.Seleccion
    Dim xfrm As New eps_librerias.FormSeleccion
    Dim xCampos(4, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim xRs1 As New ADODB.Recordset
    
    xCampos(0, 0) = "Nº Documento":    xCampos(0, 1) = "nrodoc":        xCampos(0, 2) = "1500":   xCampos(0, 3) = "C":     xCampos(0, 4) = "S"
    xCampos(1, 0) = "Fch. Giro":       xCampos(1, 1) = "fecgiro":       xCampos(1, 2) = "1000":   xCampos(1, 3) = "C":     xCampos(1, 4) = "N"
    xCampos(2, 0) = "Cliente":         xCampos(2, 1) = "nombre":        xCampos(2, 2) = "2500":   xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "Motivo":          xCampos(3, 1) = "descripcion":   xCampos(3, 2) = "2000":   xCampos(3, 3) = "C":     xCampos(3, 4) = "N"

    xfrm.SQLCad = "SELECT vta_guia.id, vta_guia.fecgiro, [vta_guia]![numser]+'-'+[vta_guia]![numdoc] AS NroDoc, mae_cliente.numruc, mae_cliente.nombre, " _
        & " mae_mottra.descripcion, vta_guia.idcli, mae_documento.abrev FROM mae_mottra RIGHT JOIN ((mae_cliente RIGHT JOIN vta_guia ON " _
        & " mae_cliente.id = vta_guia.idcli) LEFT JOIN mae_documento ON vta_guia.tipdoc = mae_documento.id) ON mae_mottra.id = vta_guia.idmottra " _
        & " WHERE (((vta_guia.idcli)=" & NulosN(LblIdCliente.Caption) & ") AND ((vta_guia.Anulado)=0) AND ((vta_guia.iddocven)=0)) " _
        & " ORDER BY [vta_guia]![numser]+'-'+[vta_guia]![numdoc] DESC"
        
    xfrm.Titulo = "Buscando Guias del Guias"
    
    Set xfrm.Coneccion = xCon
    Set xRs = xfrm.Seleccionar(xCampos)
    
    Fg4.Rows = 1
    If xRs.State = 1 Then
        If xRs.RecordCount = 0 Then
            Set xRs = Nothing
            Exit Sub
        End If
        Dim xCadWhere As String
        Dim A As Integer
        Dim Rst As New ADODB.Recordset
        
        xRs.MoveFirst
        
        'CARGAMOS LOS DOCUMENTOS ADJUNTOS Y LO MOSTRAMOS EN LA LISTA DE "DOCUMENTOS ADJUNTOS"
        For A = 1 To xRs.RecordCount
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(A, 1) = xRs("nrodoc")
            Fg4.TextMatrix(A, 2) = xRs("abrev")
            Fg4.TextMatrix(A, 3) = xRs("id")
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
        
        MostrarItems
        Agregando = False
    End If
    
    HallarTotal
    Set xfrm = Nothing
End Sub

Sub MostrarItems()
    Dim A As Integer
    Dim xCadWhere  As String
    Dim Rst As New ADODB.Recordset
    xCadWhere = ""
    
    If Fg4.Rows = 1 Then
        Fg1.Rows = 1
        Exit Sub
    End If
    'CREAMOS LA SENTENCIA WHERE PARA LA CONSULTA SQL
    For A = 1 To Fg4.Rows - 1
        xCadWhere = xCadWhere + "(vta_guiadet.idgui=" & Fg4.TextMatrix(A, 3) & ")"
        If A = Fg4.Rows - 1 Then Exit For
        xCadWhere = xCadWhere + " OR "
    Next A
    
    RST_Busq Rst, "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, Sum(vta_guiadet.canpro) AS SumaDecanpro, alm_inventario.id, " _
        & " alm_inventario.idtipven, alm_inventario.idcuentaven, alm_inventario.stckact, alm_inventario.idunimed FROM (vta_guiadet LEFT JOIN alm_inventario " _
        & " ON vta_guiadet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON vta_guiadet.idunimed = mae_unidades.id " _
        & " WHERE " & Trim(xCadWhere) _
        & " GROUP BY alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.id, alm_inventario.idtipven, alm_inventario.idcuentaven, " _
        & " alm_inventario.stckact, alm_inventario.idunimed ORDER BY alm_inventario.id", xCon

    'SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, Sum(vta_guiadet.canpro) AS SumaDecanpro, alm_inventario.id, " _
        & " alm_inventario.idtipven, alm_inventario.idcuentaven, alm_inventario.stckact, alm_inventario.idunimed FROM (vta_guiadet LEFT JOIN alm_inventario " _
        & " ON vta_guiadet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON vta_guiadet.idunimed = mae_unidades.id " _
        & " Where " & Trim(xCadWhere) _
        & " GROUP BY alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.id, alm_inventario.idtipven, alm_inventario.idcuentaven, " _
        & " alm_inventario.stckact ORDER BY alm_inventario.id", xCon
    
    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Dim xPrecio As Double
        Agregando = True
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("descripcion") 'DESCRIPCION DEL PRODUCTO
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("abrev")       'ABREVIATURA
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(Rst("SumaDecanpro"), "0.00") 'CANTIDAD
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(1, "0.00")  'PRECIO UNITARIO
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Rst("id")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Rst("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Rst("idcuentaven")
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Rst("idtipven")
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(Rst("stckact"))
            
            If NulosN(LblIdCliente.Caption) <> 0 Then
                xPrecio = UltimoPrecio(NulosN(Rst("id")), NulosN(LblIdCliente.Caption))
            Else
                xPrecio = UltimoPrecio(NulosN(Rst("id")), 0)
            End If
            
            Fg1.TextMatrix(A, 4) = Format((xPrecio), "0.000000")
            Fg1.TextMatrix(A, 6) = Format((xPrecio), "0.000000")
            Fg1.TextMatrix(A, 7) = (xPrecio * NulosN(Fg1.TextMatrix(A, 3)))
            Fg1.TextMatrix(A, 7) = Format(Fg1.TextMatrix(A, 7), "0.00")
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
     Agregando = False
    End If
End Sub

Private Sub CmdBusAlm_Click()
    If QueHace = 3 Then Exit Sub
    
    'Dim xform As New EPS_Buscar.Buscar
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT alm_almacenes.* FROM alm_almacenes"
    
    xform.Titulo = "Buscando Almacenes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblAlmacen.Caption = xRs("descripcion")
        TxtIdAlm.Text = xRs("id")
        TxtTipDoc.SetFocus
        
        If TxtTipDoc.Text <> "" Then
            Dim Rst As New ADODB.Recordset
            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(TxtIdAlm.Text) & "", xCon)
            If Rst.RecordCount <> 0 Then
                TxtNumSer.Text = Rst("numser")
                TxtNumSer_Validate True
            End If
            
            Set Rst = Nothing
        Else
            TxtNumSer.Text = ""
            TxtNumDoc.Text = ""
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusAlmacen2_Click()
    'Dim xform As New EPS_Buscar.Buscar
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT alm_almacenes.* FROM alm_almacenes"
    
    xform.Titulo = "Buscando Almacenes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblidAlmacen2.Caption = xRs("id")
        TxtAlmacen2.Text = xRs("descripcion")
        TxtIdDocGen.SetFocus
        
        If TxtIdDocGen.Text <> "" Then
            Dim Rst As New ADODB.Recordset
            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(LblIdDocumentoGen.Caption) & " AND idalm = " & NulosN(LblidAlmacen2.Caption) & "", xCon)
            If Rst.RecordCount <> 0 Then
                TxtNumSerGen.Text = Rst("numser")
                'TxtNumSerGen_Validate True
            End If
            
            Set Rst = Nothing
        Else
            TxtNumSerGen.Text = ""
            TxtNumDocGen.Text = ""
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusCondicion_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_condpago ORDER BY descripcion"
    
    xform.Titulo = "Buscando Condicion de Pago"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtConPag.Text = xRs("id")
            LblCondPag.Caption = xRs("descripcion")
            If NulosC(TxtFchDoc.Valor) <> "" Then
                TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + xRs("numdia")
            End If
            TxtFchVen.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDocRef_Click()
    If QueHace = 3 Then Exit Sub

    If NulosN(LblIdCliente.Caption) = 0 Then
        MsgBox "No ha especificado el cliente para referenciar este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(6, 4) As String
    
    xCampos(0, 0) = "Tipo. Doc.":       xCampos(0, 1) = "abrev":                xCampos(0, 2) = "1000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Doc.":        xCampos(1, 1) = "fchdoc":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Documento":     xCampos(2, 1) = "numdoc":               xCampos(2, 2) = "1500":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fch. Ven.":        xCampos(3, 1) = "fchven":               xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
    xCampos(4, 0) = "Total":            xCampos(4, 1) = "imptot":               xCampos(4, 2) = "1000":         xCampos(4, 3) = "N"
    xCampos(5, 0) = "Condicion":        xCampos(5, 1) = "descripcion":          xCampos(5, 2) = "1000":         xCampos(5, 3) = "C"
    
    xform.SQLCad = "SELECT mae_documento.abrev, vta_ventas.fchdoc, [vta_ventas]![numser] & '-' & [vta_ventas]![numdoc] AS numdoc, vta_ventas.fchven, " _
        & " mae_cliente.nombre, mae_condpago.descripcion, vta_ventas.id, vta_ventas.imptotdoc, vta_ventas.idcli, vta_ventas.tipdoc " _
        & " FROM ((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) " _
        & " LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id Where (((vta_ventas.idcli) = " & NulosN(LblIdCliente.Caption) & ") And ((vta_ventas.tipdoc) <> 7)) " _
        & " ORDER BY vta_ventas.fchdoc DESC"
    
    xform.Titulo = "Buscando Documentos del Cliente"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtDocRefCredi.Text = xRs("numdoc")
            LblIdDocRef.Caption = xRs("id")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDocRef2_Click()
    If QueHace = 3 Then Exit Sub
    
    If NulosN(TxtIdTipDoc.Text) = 0 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Nº Documento":      xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Emi.":         xCampos(1, 1) = "fchemi":      xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Fch. Ven.":         xCampos(2, 1) = "fchven":      xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Cliente":           xCampos(3, 1) = "nombre":      xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
    
    If NulosN(TxtIdTipDoc.Text) = 4 Then
        'Orden de Despacho
        xform.SQLCad = "SELECT var_ordendespacho.id, var_ordendespacho!anno & var_ordendespacho!idaduana & var_ordendespacho!idregimen & var_ordendespacho!numdoc AS numdoc, " _
            & " mae_cliente.nombre, var_ordendespacho.idcli, var_ordendespacho.fchemi, var_ordendespacho.fchven FROM var_ordendespacho LEFT JOIN mae_cliente " _
            & " ON var_ordendespacho.idcli = mae_cliente.id"
        
        xform.Titulo = "Orden de Despacho"
    End If
    
    If NulosN(TxtIdTipDoc.Text) = 5 Then
        'Orden de pedido
        MsgBox "Opcion no disponible"
        xform.Titulo = "Orden de Produccion"
        Exit Sub
    End If
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numdoc"
    xform.CampoBusca = "numdoc"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumDocRef.Text = xRs("numdoc")
            LblIdDocRef2.Caption = xRs("id")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusIdTipDocRef_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_docreferencia ORDER BY descripcion"
    
    xform.Titulo = "Buscando Tipo de Documento de Referencia"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdTipDoc.Text = xRs("id")
            LblDescTipDocRef.Caption = xRs("descripcion")
            TxtNumDocRef.Text = ""
            LblIdDocRef2.Caption = ""
            TxtNumDocRef.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusNumSer_Click()
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Sub
    End If

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "iddoc":       xCampos(0, 2) = "1500":    xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion": xCampos(1, 2) = "2500":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Serie":          xCampos(2, 1) = "numser":      xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"
'    xCampos(3, 0) = "Nro Documento":  xCampos(3, 1) = "numdoc":      xCampos(3, 2) = "1500":    xCampos(3, 3) = "C"
    
    xform.SQLCad = "SELECT mae_documento.descripcion, mae_series.iddoc, Format([mae_series].[numser],'0000') AS numser, " _
        & " mae_series.numdoc FROM mae_documento INNER JOIN mae_series ON mae_documento.id = mae_series.iddoc " _
        & " WHERE (((mae_series.iddoc)=1))"
        
        
        

    xform.Titulo = "Buscando Series"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numser"
    xform.CampoBusca = "numser"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumSer.Text = Format(xRs("numser"), "0000")
            TxtNumDoc = HallaNumdocVenta(NulosN(TxtTipDoc.Text), TxtNumSer.Text, xCon) ' HallaNumDoc(CLng(TxtTipDoc), CLng(Trim(TxtNumSer.Text)))
        End If
        TxtConPag.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCli_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Cliente":      xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id, mae_cliente.idven From mae_cliente"
    
    xform.Titulo = "Buscando Cliente"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            If xRs.RecordCount <> 0 Then
                TxtNumRuc.Text = xRs("numruc")
                LblNomCli.Caption = xRs("nombre")
                LblIdCliente.Caption = xRs("id")
                TxtIdVen.Text = NulosN(xRs("idven"))
                TxtIdVen_Validate True
                TxtNumSer.SetFocus
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_moneda ORDER BY descripcion"
    
    xform.Titulo = "Buscando Moneda"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdMon.Text = xRs("id")
            LblMoneda.Caption = xRs("descripcion")
            Fg1.SetFocus
        
            If Trim(TxtIdMon.Text) = "1" Then
                LblTipCam.Visible = False
                LblTipoCambio.Visible = False
            Else
                If TxtFchDoc.Valor = "" Then
                    MsgBox "No ha especificado la fecha del documento, no se puede determinar " & Chr(13) _
                        & "la fecha del tipo de cambio para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    
                    TxtIdMon.Text = ""
                    TxtFchDoc.SetFocus
                    Exit Sub
                End If
                LblTipCam.Visible = True
                LblTipoCambio.Visible = True
                'Set xRs = Nothing
                'Set xRs = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = CDATE('" & TxtFchDoc.Valor & "')", xCon)
                'If xRs.RecordCount = 1 Then
                '    LblTipoCambio.Caption = Format(xRs("impven"), "0.000")
                'End If
                LblTipoCambio.Caption = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusSerGen_Click()
    If TxtIdDocGen.Text = "" Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Nombre":         xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Abreviatura":    xCampos(1, 1) = "abrev":            xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Cod. Sunat":     xCampos(2, 1) = "codsun":           xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nº Serie":       xCampos(3, 1) = "numser":           xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
    
    xform.SQLCad = "SELECT alm_numseries.idalm, alm_numseries.idtipdoc, alm_numseries.numser, mae_documento.abrev, mae_documento.codsun, " _
        & " mae_documento.descripcion " _
        & " FROM alm_numseries LEFT JOIN mae_documento ON alm_numseries.idtipdoc = mae_documento.id " _
        & " WHERE (((alm_numseries.idalm)=" & NulosN(LblidAlmacen2.Caption) & ") AND ((alm_numseries.idtipdoc)=" & NulosN(LblIdDocumentoGen.Caption) & "))"
    
    xform.Titulo = "Buscando Series de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numser"
    xform.CampoBusca = "numser"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumSerGen.Text = xRs("numser")
        TxtNumDocGen.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, mae_impuestos.Abrev AS abreimp, " _
        & " mae_impuestos.idcuenvta AS cuentaimp, alm_numseries.numser FROM alm_numseries LEFT JOIN (mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas " _
        & " ON mae_impuestos.idcuenvta = con_planctas.id) ON mae_documento.idimp = mae_impuestos.id) ON alm_numseries.idtipdoc = mae_documento.id " _
        & " WHERE (((alm_numseries.idalm)=" & NulosN(TxtIdAlm.Text) & "))"



    Dim xImpuesto As Double
    
    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumSer.Text = xRs("numser")
            TxtNumSer_Validate True
            TxtTipDoc.Text = xRs("id")
            LblNomDoc.Caption = xRs("descripcion")
            TasaImpuesto = NulosN(xRs("tasa"))
            
            xIdCuenTasa = NulosN(xRs("cuentaimp"))
            LblRotulo.Caption = Trim(NulosC(xRs("abreimp"))) + " (         )"
            LblIgvTasa.Caption = Format(Trim(Str(TasaImpuesto)), "0.00")
            FraRetencion.Caption = "( Afecta : " + NulosC(xRs("descimp")) + ")"
            
            Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & NulosN(TxtIdMon) & " and tipope = -1", xCon)
            If xRs2.RecordCount > 0 Then
                xCuentaDoc = NulosN(xRs2("idcuen"))
            End If
            Set xRs2 = Nothing
            
            Frame5.Visible = False
            
            'si es Recibo por honorarios
            If xRs("id") = 2 Then
                 FraRetencion.Enabled = True
                 Fratipven.Enabled = False
                 FraRetencion.Visible = True
                 FraRetencion.Caption = "Retención de 4ta Categoria" & "10%"
                 txtisc.Enabled = False
                 txtinafecto.Enabled = False
            Else
                 Fratipven.Enabled = True
                 FraRetencion.Enabled = False
                 FraRetencion.Visible = False
                 txtisc.Enabled = True
                 txtinafecto.Enabled = True
            End If
            
            If xRs("id") = 7 Then
                Label33.Visible = True
                TxtDocRefCredi.Visible = True
                CmdBusDocRef.Visible = True
            Else
                Label33.Visible = False
                TxtDocRefCredi.Visible = False
                CmdBusDocRef.Visible = False
            End If
            TxtNumRuc.SetFocus
        End If
    
        If TxtTipDoc.Text <> "" Then
            Dim Rst As New ADODB.Recordset
            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(LblIdAlmacen.Caption) & "", xCon)
            If Rst.RecordCount <> 0 Then
                TxtNumSer.Text = Rst("numser")
                TxtNumSer_Validate True
            End If
            Set Rst = Nothing
        Else
            TxtNumSer.Text = ""
            TxtNumDoc.Text = ""
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDocGen_Click()
   
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    'xform.SQLCad = "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuenvta  as cuentaimp" _
        & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id"
    
    xform.SQLCad = "SELECT DISTINCT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuenvta AS cuentaimp " _
        & " FROM alm_numseries LEFT JOIN (mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id) ON alm_numseries.idtipdoc = mae_documento.id WHERE (((alm_numseries.idalm)=" & NulosN(LblidAlmacen2.Caption) & "))"

    
    'SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, mae_impuestos.Abrev AS abreimp, " _
        & " mae_impuestos.idcuenvta AS cuentaimp, alm_numseries.numser FROM alm_numseries LEFT JOIN (mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas " _
        & " ON mae_impuestos.idcuenvta = con_planctas.id) ON mae_documento.idimp = mae_impuestos.id) ON alm_numseries.idtipdoc = mae_documento.id " _
        & " WHERE (((alm_numseries.idalm)=" & NulosN(LblidAlmacen2.Caption) & "))"

    Dim xImpuesto As Double
    
    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            LblIdDocumentoGen = xRs("id")
            TxtIdDocGen.Text = xRs("descripcion")
            xIdCuenTasa = NulosN(xRs("cuentaimp"))
            
            Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(LblIdDocumentoGen.Caption) & " and mae_documentocta.idmon =" & 1 & " and tipope = -1", xCon)
            If xRs2.RecordCount > 0 Then
                xCuentaDoc = NulosN(xRs2("idcuen"))
            Else
                MsgBox "No se ha encontrado cuenta contable para el documento especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
            Set xRs2 = Nothing
            
            Frame5.Visible = False
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipItem_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
    
    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipItem.Text = NulosN(xRs("id"))
            LblTipoItem = NulosC(xRs("descripcion"))
            TxtIdAlm.SetFocus
        End If
        
        pGridConfigurar
        
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusVen_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Codigo":    xCampos(0, 1) = "id":         xCampos(0, 2) = "800":          xCampos(0, 3) = "N"
    xCampos(1, 0) = "Vendedor":  xCampos(1, 1) = "apenom":     xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Basico":    xCampos(2, 1) = "basico":     xCampos(2, 2) = "1200":         xCampos(2, 3) = "N"
    xCampos(3, 0) = "Comision":  xCampos(3, 1) = "comision":   xCampos(3, 2) = "1200":         xCampos(3, 3) = "N"
    
    xform.SQLCad = "SELECT vta_vendedores.*, UCase([pla_empleados]![apepat]) & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenom " _
                & " FROM vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id"
    
    xform.Titulo = "Buscando Vendedores"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "apenom"
    xform.CampoBusca = "apenom"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        LblNomVen.Caption = xRs("apenom")
        TxtIdVen.Text = xRs("id")
        'TxtAutoriza.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDelItem_Click()
    If QueHace = 3 Then Exit Sub
    If Fg1.Row < 1 Or Fg1.Rows < 1 Then Exit Sub
    
    Fg1.RemoveItem Fg1.Row
    HallarTotal
    If Fg1.Rows <> 1 Then Fg1.Select Fg1.Rows - 1, 1
End Sub

Private Sub cmdEliminarOKdocsproc_Click()
    Dim Rstguia As New ADODB.Recordset
    Dim X As Integer

    If fgdocsproc.Rows - 1 > 0 Then
        If fgdocsproc.Rows - 1 = 1 Then
            fgdocsproc.Rows = 1
            Fg1.Rows = 1
            HallarTotal
            Exit Sub
        Else
            With Me.Fg1
                For X = 1 To Me.Fg1.Rows - 1
                    RST_Busq Rstguia, "Select Vta_GuiaDet.* From Vta_GuiaDet where Vta_GuiaDet.IdGui = " & NulosN(fgdocsproc.TextMatrix(fgdocsproc.Row, 1)) & " and Vta_GuiaDet.IdItem = " & NulosN(Fg1.TextMatrix(X, 6)) & "", xCon
                    
                    If Rstguia.RecordCount > 0 Then
                        .TextMatrix(X, 4) = NulosN(.TextMatrix(X, 4)) - Rstguia("canpro")
                    End If
                Next
                'Colocamos en el campo estado 0  de la tabla guia que indica que no esta facturado
                xCon.Execute " UPDATE vta_guia SET Vta_guia.Estado = 0 WHERE vta_guia.id = " & NulosN(fgdocsproc.TextMatrix(fgdocsproc.Row, 1)) & ""
            End With
                
            fgdocsproc.RemoveItem fgdocsproc.Row
            HallarTotal
        End If
    End If
    Set Rstguia = Nothing
End Sub

Sub CargarRSTCom(xFechaRegistro As String, Mes As Integer)
    Dim DiaIniAño As String
    DiaIniAño = "01/01/" + Trim(AnoTra)
    
    If mMesActivo >= 1 And mMesActivo <= 12 Then
        RST_Busq RstVent, "SELECT vta_ventas.*, IIf(vta_ventas.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, " _
            & " IIf(vta_ventas.Anulado=0,'Facturado','Anulado') AS EstadoVenta, mae_documento.descripcion AS nomdoc, IIf(vta_ventas.anulado=-1,'', mae_condpago.descripcion) AS desccond, " _
            & " mae_documento.abrev, mae_cliente.numruc, mae_moneda.descripcion AS descmon, IIf(vta_ventas.anulado=-1,'',mae_moneda.simbolo) AS simbolo, mae_impuestos.idcuenvta, " _
            & " con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom, IIf(vta_ventas.anulado=-1,'',mae_condpago.abrev) AS conpagabre, " _
            & " Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4) AS numreg1 FROM ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento LEFT JOIN mae_impuestos " _
            & " ON mae_documento.idimp = mae_impuestos.id) RIGHT JOIN (mae_condpago RIGHT JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) " _
            & " ON mae_condpago.id = vta_ventas.idconpag) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) " _
            & " LEFT JOIN mae_tipoproducto ON vta_ventas.idtipo = mae_tipoproducto.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            & " WHERE (((vta_ventas.fchreg)=CDate('" & xFechaRegistro & "')) AND ((vta_ventas.fchdoc)>=CDate('" & DiaIniAño & "'))) ORDER BY vta_ventas!numser+'-'+vta_ventas!numdoc DESC", xCon
    End If
    
    If mMesActivo = 0 Then
        RST_Busq RstVent, "SELECT vta_ventas.*, mae_cliente.nombre, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numerodoc, IIf(vta_ventas.Anulado=0,'Facturado','Anulado') AS EstadoVenta, " _
            & " mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_moneda.descripcion AS descmon, " _
            & " mae_moneda.simbolo, mae_impuestos.idcuenvta, con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom, mae_condpago.abrev AS conpagabre, " _
            & " Mid([vta_ventas].[numreg],1,2)+[mae_libros].[codsun]+Mid([vta_ventas].[numreg],3,4) AS numreg1 FROM ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN " _
            & " ((mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) RIGHT JOIN (mae_condpago RIGHT JOIN (vta_ventas LEFT JOIN con_tc " _
            & " ON vta_ventas.fchdoc = con_tc.fecha) ON mae_condpago.id = vta_ventas.idconpag) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) " _
            & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_tipoproducto ON vta_ventas.idtipo = mae_tipoproducto.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            & " WHERE (((vta_ventas.fchreg)=CDate('" & DiaIniAño & "')) AND ((vta_ventas.fchdoc)<CDate('" & DiaIniAño & "'))) ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] DESC", xCon
    End If
    
    If mMesActivo = 13 Then
        MsgBox "Ha selecionado el mes de Cierre, selecciones meses comprendidos entre Enero y Diciembre", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstVent = Nothing
        Set Dg1.DataSource = Nothing
        Dg1.Refresh
        Exit Sub
    End If
    RstVent.Requery
    Set Dg1.DataSource = RstVent
    Dg1.Refresh
End Sub

Private Sub CmdMotNotCre_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM vta_conceptonc ORDER BY descripcion"
    
    xform.Titulo = "Buscando Concepto Nota de Credito"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtDocRef.Text = xRs("descripcion")
            LblIdConNC.Caption = xRs("id")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub cmdokseldoc_Click()
    If TxtIdDocGen.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdDocGen.SetFocus
        Exit Sub
    End If
    
    If TxtNumSerGen.Text = "" Then
        MsgBox "No ha especificado el numero de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSerGen.SetFocus
        Exit Sub
    End If
    
    If TxtNumDocGen.Text = "" Then
        MsgBox "No ha especificado el numero del documeto a generar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDocGen.SetFocus
        Exit Sub
    End If

    Dim xFecha As String
    xFecha = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)

    If CDate(TxtFchEmiAnul.Valor) < CDate(xFecha) Then
        MsgBox "La fecha del documento no corresponde la periodo contable especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If

    Dim RstCab As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim xId As Double
    Dim xNumAsiento As String

    RST_Busq xRs, "SELECT vta_ventas.tipdoc, vta_ventas.numser, vta_ventas.numdoc From vta_ventas " _
        & " WHERE (((vta_ventas.tipdoc)=" & NulosN(LblIdDocumentoGen.Caption) & ") AND ((vta_ventas.numser)='" & TxtNumSerGen.Text & "') " _
        & " AND ((vta_ventas.numdoc)='" & TxtNumDocGen.Text & "'))", xCon

    If xRs.RecordCount = 1 Then
        Set xRs = Nothing
        MsgBox "El numero de documento que quiere emitir ya existe", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDocGen.SetFocus
        Exit Sub
    End If
    
    xNumAsiento = NuevoNumAsiento(2, mMesActivo, xCon)

On Error GoTo LaCague
    xCon.BeginTrans

    'Validar si el nro de documento existe solo en modo adicionar documento
    RST_Busq RstCab, "SELECT TOP 1 * FROM vta_ventas", xCon
    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    
    xId = HallaCodigoTabla("vta_ventas", xCon, "id")
    RstCab.AddNew
    RstCab("id") = xId
    RstCab("idlib") = 2
    RstCab("idtipo") = 1
    RstCab("tipdoc") = NulosN(LblIdDocumentoGen.Caption)
    RstCab("idcli") = 1
    RstCab("numser") = TxtNumSerGen.Text
    RstCab("numdoc") = TxtNumDocGen.Text
    RstCab("Fchdoc") = TxtFchEmiAnul.Valor
    RstCab("Fchven") = TxtFchEmiAnul.Valor
    RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    RstCab("idconpag") = 0
    RstCab("idmon") = 1
    RstCab("impbru") = 0
    RstCab("impinaf") = 0
    RstCab("impigv") = 0
    RstCab("impisc") = 0
    RstCab("impotr") = 0
    RstCab("imptotdoc") = 0
    RstCab("impsal") = 0
    RstCab("idmon") = 1
    RstCab("numreg") = Format(mMesActivo, "00") + Trim(xNumAsiento)
    RstCab("anulado") = -1
    'Determinamos si es una exportacion
    RstCab("idtipven") = 0 'en el cual puede ser venta afecta o inafecta para el registro de de ventas
                           'se valida por programa ver tabla mae_tipoventa
    RstCab.Update
    
    
    RstDia.AddNew
    'grabamos el documento de venta en la tabla diario
    RstDia("año") = AnoTra
    RstDia("idmes") = mMesActivo
    RstDia("idlib") = 2
    RstDia("iddoc") = NulosN(LblIdDocumentoGen.Caption)
    RstDia("idmov") = xId
    RstDia("numasi") = xNumAsiento
    RstDia("tc") = ValTipCam
    RstDia("idcue") = xCuentaDoc
    
    If TxtIdMon.Text = "1" Then
        RstDia("impdebsol") = 0
        RstDia("impdebdol") = 0
    Else
        RstDia("impdebsol") = 0
        RstDia("impdebdol") = 0
    End If
    RstDia.Update
    
    RstDia.AddNew
    'grabamos el impuesto del documento de venta en la tabla diario
    RstDia("año") = AnoTra
    RstDia("idmes") = mMesActivo
    RstDia("idlib") = 2
    RstDia("iddoc") = NulosN(LblIdDocumentoGen.Caption)
    RstDia("idmov") = xId
    RstDia("numasi") = xNumAsiento
    RstDia("tc") = ValTipCam
    RstDia("idcue") = xIdCuenTasa
    
    If TxtIdMon.Text = "1" Then
        RstDia("impdebsol") = 0
        RstDia("impdebdol") = 0
    Else
        RstDia("impdebsol") = 0
        RstDia("impdebdol") = 0
    End If
    RstDia.Update
        
    xCon.CommitTrans
        
    MsgBox "El documento anulado se genero con éxito" & vbCr & "Registro Nº: " & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDia = Nothing
    RstVent.Requery
    Dg1.Refresh
    cmdsalirseldoc_Click
    Exit Sub
    
LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set xRs = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Private Sub CmdPreHist_Click()
    If Fg1.Rows < 1 Then Exit Sub
    If Fg1.Row < 1 Then
        MsgBox "Seleccione un Registro para ver el Histórico de Precios", vbInformation, xTitulo
        Exit Sub
    End If
    
    Dim xfrm As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    xfrm.PreciosHistoricos xCon, Fg1.TextMatrix(Fg1.Row, 8), False, NulosC(TxtNumRuc.Text)
    Set xfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdSalirdocsproc_Click()
    Fradocsproc.Visible = False
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
End Sub

'Private Sub CmdSalirRef_Click()
'    Toolbar1.Enabled = True
'    TabOne1.Enabled = True
'    fraconsdocref.Visible = False
'End Sub

Private Sub cmdsalirseldoc_Click()
    ActivarEntorno
    Fraseldoc.Visible = False
End Sub

Private Sub Command1_Click()
    Dim Rpta As Integer
    
    Rpta = MsgBox("Esta seguro de modificar el saldo del documento", vbInformation + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        'actualizamos el saldo del documento en vta_ventas
        xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = 222 WHERE (((vta_ventas.id)=" & RstVent("id") & "))"

        
    End If
End Sub

Private Sub Command2_Click()
    ActivarEntorno
    Frame8.Visible = False
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstVent
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstVent.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear

End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 2, NulosN(RstVent("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If NulosC(TxtTipItem.Text) = "" Then
        MsgBox "No ha especificado el tipo de item a buscar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipItem.SetFocus
        Exit Sub
    End If
    
    If optsinguia.Value <> True Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "4800":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Unid.":         xCampos(1, 1) = "abrev":          xCampos(1, 2) = "500":     xCampos(1, 3) = "C"
    xCampos(2, 0) = "Stock":        xCampos(2, 1) = "stckact":        xCampos(2, 2) = "800":     xCampos(2, 3) = "N"
    xCampos(3, 0) = "Código":       xCampos(3, 1) = "codpro":         xCampos(3, 2) = "2000":    xCampos(3, 3) = "C"
    
    
    '*******************************************************************************************
    Dim nSQLId As String
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 8, "alm_inventario.id", " NOT IN ", True)
    If nSQLId <> "" Then nSQLId = " AND " & nSQLId
    '*******************************************************************************************
    '--obs. apareceran solo items de ventas que tengan cuenta contable
    
    xform.SQLCad = "SELECT alm_inventario.*, mae_unidades.descripcion AS descuni, mae_unidades.abrev, mae_percepcion.tasa " _
        & " FROM mae_unidades RIGHT JOIN (mae_percepcion RIGHT JOIN alm_inventario ON mae_percepcion.id = alm_inventario.idper) " _
        & " ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = " & NulosN(TxtTipItem) & " )) " & nSQLId & " AND alm_inventario.idcuentaven <>0 ORDER BY alm_inventario.descripcion"
    
    xform.Titulo = "Buscando Productos"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    
'    xform.FormaBusca = CualquierParte
'    xform.Ordenado = "descripcion"
'    xform.CampoBusca = "descripcion"
    
    xform.Ordenado = "codpro"
    xform.CampoBusca = "codpro"
    
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    Dim A As Integer
    If xRs.State = 1 Then
        If NulosN(xRs("idcuentaven")) = 0 Then
            MsgBox "El item seleccionado no tiene una cuenta contable asignada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set xform = Nothing
            Set xRs = Nothing
            Exit Sub
        End If
        
        If Fg1.Rows <> 1 Then
            For A = 1 To Fg1.Rows - 1
    
                If NulosN(Fg1.TextMatrix(A, 8)) = xRs("id") Then
                    MsgBox "El item seleccionado ya fue agregado a la lista, seleccione otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Set xRs = Nothing
                    
                    A = Fg1.Rows - 1
                    Exit Sub
                End If
            Next A
        End If
        
        If xRs.RecordCount <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 1) = xRs("descripcion")
            Fg1.TextMatrix(Fg1.Row, 2) = NulosN(xRs("abrev"))
            If NulosN(LblIdCliente.Caption) <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 4) = UltimoPrecio(xRs("id"), NulosN(LblIdCliente.Caption)) 'Format(NulosN(xRs("preuni")), "0.0000")
            Else
                Fg1.TextMatrix(Fg1.Row, 4) = UltimoPrecio(xRs("id"), 0)
            End If
            Fg1.TextMatrix(Fg1.Row, 8) = xRs("id")
            Fg1.TextMatrix(Fg1.Row, 9) = NulosN(xRs("idunimed"))
            Fg1.TextMatrix(Fg1.Row, 10) = NulosN(xRs("idcuentaven"))
            Fg1.TextMatrix(Fg1.Row, 11) = NulosN(xRs("idtipven"))
            Fg1.TextMatrix(Fg1.Row, 12) = NulosN(xRs("tasa"))
            Fg1.TextMatrix(Fg1.Row, 13) = NulosN(xRs("stckact"))
        End If
    End If
    '------------
    If Fg1.Row >= 1 Then
        If NulosN(TxtTipItem.Text) = 5 Then
            Fg1.Col = 4
        Else
            Fg1.Col = 3
        End If
    End If
    '------------
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Function UltimoPrecio(IdItem As Integer, IdCliente As Integer) As Double
    Dim Rst As New ADODB.Recordset
    If IdCliente <> 0 Then
        'Si hay un clientes asignado buscamos el ultimo precio de venta al cliente
        RST_Busq Rst, "SELECT vta_ventas.fchdoc, vta_ventasdet.preuni, vta_ventasdet.iditem, vta_ventas.idcli FROM vta_ventas LEFT JOIN vta_ventasdet " _
            & " ON vta_ventas.id = vta_ventasdet.idvta Where (((vta_ventasdet.IdItem) = " & IdItem & ") And ((vta_ventas.idcli) = " & IdCliente & ")) " _
            & " ORDER BY vta_ventas.fchdoc, vta_ventasdet.preuni", xCon
    Else
        'si no hay un cliente especificado buscamos el ultimo precio de venta del item
        RST_Busq Rst, "SELECT vta_ventas.fchdoc, vta_ventasdet.preuni, vta_ventasdet.iditem, vta_ventas.idcli FROM vta_ventas LEFT JOIN vta_ventasdet " _
            & " ON vta_ventas.id = vta_ventasdet.idvta Where (((vta_ventasdet.IdItem) = " & IdItem & ")) ORDER BY vta_ventas.fchdoc, vta_ventasdet.preuni", xCon
    End If
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveLast
        UltimoPrecio = NulosN(Rst("preuni"))
    Else
        Set Rst = Nothing
        RST_Busq Rst, "SELECT * FROM alm_inventario WHERE (id = " & IdItem & ")", xCon
        If Rst.RecordCount <> 0 Then
            UltimoPrecio = 0
        Else
            UltimoPrecio = NulosN(Rst("preuni"))
        End If
    End If
    Set Rst = Nothing
End Function

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim xTotPorDes As Double
    
    If Agregando = True Then Exit Sub
    If Mostrando = True Then Exit Sub
    If Fg1.Row < 0 Then Exit Sub
    If optsinguia.Value = True Then
        If Col = 3 Then
            Fg1.TextMatrix(Fg1.Row, 3) = Format(Fg1.TextMatrix(Fg1.Row, 3), "0.0000")
        End If
    End If
    
    If Col = 4 And NulosN(TxtTipItem.Text) <> 5 Then
        Dim xSaldo As Double
        xSaldo = NulosN(Fg1.TextMatrix(Fg1.Row, 13)) - NulosN(Fg1.TextMatrix(Fg1.Row, 3))
        
        If NulosN(Fg1.TextMatrix(Fg1.Row, 3)) > NulosN(Fg1.TextMatrix(Fg1.Row, 13)) Then
            MsgBox "No hay suficiente stock del producto : " + Fg1.TextMatrix(Fg1.Row, 1) & Chr(13) _
                & "Cantidad Solicitada : " + Trim(Fg1.TextMatrix(Fg1.Row, 3)) + Chr(13) _
                & "Stock Actual  : " + Trim(Format(Fg1.TextMatrix(Fg1.Row, 13), "0.00")) + Chr(13) _
                & "Faltante        : " + Trim(Str(Format(xSaldo, "0.00"))) + Chr(13), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Fg1.TextMatrix(Fg1.Row, 4) = ""
            Exit Sub
        End If
    End If
    
    
    If Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Then
        If OptDes1.Value = True Then
            xTotPorDes = (NulosN(Fg1.TextMatrix(Fg1.Row, 5)) / 100) + 1
        End If
        
        Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.000000")
        Fg1.TextMatrix(Fg1.Row, 5) = Format(Fg1.TextMatrix(Fg1.Row, 5), "0.0000")
        
        If OptDes1.Value = True Then
            Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) / xTotPorDes
        End If
        If OptDes2.Value = True Then
            Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) - NulosN(Fg1.TextMatrix(Fg1.Row, 5))
        End If
        
        Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), "0.000000")
        Fg1.TextMatrix(Fg1.Row, 7) = NulosN(Fg1.TextMatrix(Fg1.Row, 6)) * NulosN(Fg1.TextMatrix(Fg1.Row, 3))
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.00")
        HallarTotal
    End If
    
    If Col = 14 Or Col = 15 Then
        Dim xIgv As Double
        xIgv = (NulosN(LblIgvTasa.Caption) / 100) + 1
        Fg1.TextMatrix(Fg1.Row, 4) = NulosN(Fg1.TextMatrix(Fg1.Row, 14)) / xIgv
        Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.000000")
        
        If NulosN(Fg1.TextMatrix(Fg1.Row, 15)) <> 0 Then
            Dim xPreUni As Double
            xPreUni = NulosN(Fg1.TextMatrix(Fg1.Row, 14)) / NulosN(Fg1.TextMatrix(Fg1.Row, 15))
            Fg1.TextMatrix(Fg1.Row, 4) = xPreUni / xIgv
            Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.000000")
            
            Fg1.TextMatrix(Fg1.Row, 3) = Format(Fg1.TextMatrix(Fg1.Row, 15), "0.000000")
        End If
        
        'hallamos los totales
        If OptDes1.Value = True Then
            If xTotPorDes <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) / xTotPorDes
            Else
                Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 4))
            End If
        End If
        If OptDes2.Value = True Then
            Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) - NulosN(Fg1.TextMatrix(Fg1.Row, 5))
        End If
        
        Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), "0.000000")
        Fg1.TextMatrix(Fg1.Row, 7) = NulosN(Fg1.TextMatrix(Fg1.Row, 6)) * NulosN(Fg1.TextMatrix(Fg1.Row, 3))
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.00")
        HallarTotal
    End If
End Sub

Sub HallarTotal()
    Dim A As Integer
    Dim totalafec As Double
    Dim totalinaf As Double
    
    txtinafecto.Text = "0.00"
    TxtIGV.Text = "0.00"
    txtisc = "0.00"
    TxtTotal.Text = "0.00"
    
    For A = 1 To Fg1.Rows - 1
        If NulosN(TxtTipDoc) = 2 Then 'SI ES RECIBO POR HONORARIOS
            totalafec = totalafec + NulosN(Fg1.TextMatrix(A, 7)) 'venta  gravada
        Else
            If Fg1.TextMatrix(A, 11) = "1" Then 'si es venta gravada
                totalafec = totalafec + NulosN(Fg1.TextMatrix(A, 7)) 'venta  gravada
            Else
                totalinaf = totalinaf + NulosN(Fg1.TextMatrix(A, 7)) 'venta no gravada
            End If
        End If
    Next A
        
    If NulosN(TxtTipDoc) = 1 Then
        TxtIGV.Text = (totalafec * ((TasaImpuesto / 100) + 1)) - totalafec
        TxtTotal.Text = (totalafec * ((TasaImpuesto / 100) + 1)) + totalinaf
    Else
        TxtTotal.Text = (totalafec * ((TasaImpuesto / 100) + 1)) + totalinaf
        If totalafec > 0 Then
            TxtIGV.Text = (totalafec * ((TasaImpuesto / 100) + 1)) - totalafec
        End If
        txtinafecto = totalinaf
    End If
    
    TxtBruto.Text = Format(totalafec, FORMAT_MONTO)
    txtinafecto.Text = Format(totalinaf, FORMAT_MONTO)
    TxtIGV.Text = Format(TxtIGV.Text, FORMAT_MONTO)
    TxtTotal.Text = Format(TxtTotal.Text, FORMAT_MONTO)
End Sub

Private Sub Fg1_EnterCell()
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    If Fg1.Col = 2 Or Fg1.Col = 7 Then
        Fg1.Editable = flexEDNone
    Else
        If Fg1.Col = 1 Or Fg1.Col = 3 Or Fg1.Col = 5 Or Fg1.Col = 6 Or Fg1.Col = 14 Or Fg1.Col = 15 Then
            If optconguia.Value = True Then
                Fg1.Editable = flexEDNone
            Else
                Fg1.Editable = flexEDKbdMouse
            End If
        End If
        If Fg1.Col = 4 Then
            Fg1.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    If KeyAscii = 13 Then Exit Sub
    '--validar los caracteres que se ingresan
    Select Case Col
        Case 1 '--descripcion
        Case 3, 4, 5, 6, 7 '--canpro,preunibru,valdes,preuni,imptot
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case 2 '--abrev
            KeyAscii = 0
        Case 8, 9, 10, 11 '--iditem,idunimed,idcuentaven,idtipven
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If optsinguia.Value = True Then
        If KeyCode = 46 Then CmdDelItem_Click
        If KeyCode = 45 Then CmdAddItem_Click
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    
    If optsinguia.Value = True Then
        If Button = 2 Then PopupMenu menu1
    End If
End Sub

Private Sub Fg4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        cmdagregardocs_Click
    End If
    
    If KeyCode = 46 Then
        If Fg4.Rows = 1 Then Exit Sub
        Fg4.RemoveItem Fg4.Row
        If optconguia.Value = True Then
            MostrarItems
        Else
            MostrarItemsCotizacion
        End If
        HallarTotal
    End If
End Sub

Private Sub Fg4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    
    If Button = 2 Then
        PopupMenu menu2
    End If
End Sub

Private Sub fgdocsproc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then cmdEliminarOKdocsproc_Click
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
'        Dim Rpta As Integer
'        Dim xFchReg As String
        mMesActivo = xMes
        
        pCargarGrid
        
        
'        xFchReg = "01/" + Format(mMesActivo, "00") + "/" + Trim(AnoTra)
'        CargarRSTCom xFchReg, mMesActivo
'
'        Set Dg1.DataSource = RstVent
'        If RstVent.RecordCount = 0 Then
'            Rpta = MsgBox("No se ha registrado ninguna venta, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
'            If Rpta = vbYes Then
'                Nuevo
'            Else
'                mMesActivo = SeleccionaMes(xCon)
'                xFchReg = "01/" + Format(mMesActivo, "00") + "/" + Trim(AnoTra)
'                OpcionesPeriodo
'                CargarRSTCom xFchReg, mMesActivo
'                If RstVent.State <> 0 Then
'                    Set Dg1.DataSource = RstVent
'                    If RstVent.RecordCount = 0 Then
'                        If mMesActivo <> 0 And mMesActivo <> 13 Then
'                            Rpta = MsgBox("No se ha registrado ninguna venta, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
'                            If Rpta = vbYes Then
'                                Nuevo
'                            Else
'                                Unload Me
'                            End If
'                        End If
'                    Else
'                        Dg1.SetFocus
'                    End If
'                End If
'
'            End If
'        Else
'            OpcionesPeriodo
'        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then '--F3 Nuevo
        If fCierrePeriodo = False Then Exit Sub
        If QueHace <> 3 Then Exit Sub
        Nuevo
    End If
    
    If KeyCode = 115 Then '--F4 Modificar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace <> 3 Then Exit Sub
        Modificar
    End If
    
    If KeyCode = 113 Then '--F2 Grabar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace = 3 Then Exit Sub
        If Grabar = True Then
            QueHace = 3
            Set RstVent = Nothing
            Unload Me
        End If
    End If
    
    If KeyCode = 116 Then '--F5 actualizar
        
    
    End If
    
    If KeyCode = 117 Then '--F6 '--cancelar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace = 3 Then Exit Sub
        
        Cancelar
    End If
    

End Sub

Private Sub Form_Load()
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    
    
    Dg1.Columns("fchdoc1").NumberFormat = FORMAT_DATE
    Dg1.Columns("fchven1").NumberFormat = FORMAT_DATE
    
    Dg1.Columns("impbru1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impigv1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("imptotdoc1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impsal1").NumberFormat = FORMAT_MONTO
        
    CaracteresNumericos = "0123456789." & Chr(8)
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    Fg1.ColWidth(13) = 0
    Fg1.ColWidth(14) = 0
    Fg1.ColWidth(15) = 0
    '5400
    Fg4.ColWidth(3) = 0
    fgdocsproc.Rows = 1
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ColComboList(1) = ""
    swguiafact = 0
    LblIgvTasa.Caption = ""
    TxtFchDoc.Valor = Date
    TxtFchVen.Valor = Date
    
    TxtFchDoc.Valor = ""
    TxtFchVen.Valor = ""
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando una compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub menu1_1_Click()
    CmdAddItem_Click
End Sub

Private Sub menu1_3_Click()
    CmdDelItem_Click
End Sub

Private Sub Menu1_5_Click()
    CmdPreHist_Click
End Sub

Private Sub menu2_1_Click()
    cmdagregardocs_Click
End Sub

Private Sub menu2_3_Click()
    If Fg4.Rows = 1 Then Exit Sub
    Fg4.RemoveItem Fg4.Row
    CargarGuia
End Sub

Private Sub optconcotizacion_Click()
    If QueHace <> 3 Then
        cmdagregardocs.Enabled = True
        CmdAddItem.Enabled = False
        CmdDelItem.Enabled = False
    End If
    Frame6.Left = 9195
    Frame6.Top = 3165
    Frame6.Visible = True
    
    Fg1.ColWidth(1) = 5400
    Fg1.Width = 9105
End Sub

Private Sub optconguia_Click()
    If QueHace <> 3 Then
        cmdagregardocs.Enabled = True
        CmdAddItem.Enabled = False
        CmdDelItem.Enabled = False
    End If
    Frame6.Left = 9210
    Frame6.Top = 3165
    Frame6.Visible = True
    
    Fg1.ColWidth(1) = 3500
    Fg1.Width = 9105
End Sub

Private Sub OptDes1_Click()
    If OptDes1.Value = True Then Fg1.TextMatrix(0, 5) = "   % Dscto."
End Sub

Private Sub OptDes2_Click()
    If OptDes2.Value = True Then Fg1.TextMatrix(0, 5) = "Imp. Dscto."
End Sub

Private Sub OptNo_Click()
    If OptNo.Value = True Then HallarTotal
End Sub

Private Sub OptSi_Click()
    If OptSi.Value = True Then HallarTotal
End Sub

Private Sub optsinguia_Click()
    If QueHace <> 3 Then
        cmdagregardocs.Enabled = False
        CmdAddItem.Enabled = True
        CmdDelItem.Enabled = True
    End If
    
    Fg1.ColWidth(1) = 5400
    Fg1.Width = 11670
    Frame6.Visible = False
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If mMesActivo = 0 Then Cancel = 1: Exit Sub
        'Validamos si la cuadricula tiene datos
        If QueHace = 3 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No existe información para visualizar", vbInformation, Me.Caption
                Blanquea
                Exit Sub
            ElseIf NulosN(RstVent("imptotdoc")) = 0 Then
                MsgBox "El Documento de Venta esta Anulado", vbInformation, Me.Caption
                Cancel = 1
                Exit Sub
            Else
                
                MuestraSegundoTab
            End If
        End If
    End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.index = 1 Then Nuevo
    
    If Button.index = 2 Then
        If RstVent.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.index = 3 Then
        If RstVent.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
    
        'Validamos si el documento esta anulado
        If RstVent("Anulado") = -1 Then
            MsgBox RstVent("nomdoc") & " ya fue anulado, seleccione otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Anular
    End If
        
    If Button.index = 5 Then
        If Grabar = True Then
            Cancelar
            RstVent.Requery
            Dg1.Refresh
            '--------------------------------------------------------------------------
            If RstVent.RecordCount <> 0 Then
                RstVent.MoveFirst
                RstVent.Find "id=" & mIdRegistro
                If RstVent.EOF = True Then RstVent.MoveFirst
            End If
            '--------------------------------------------------------------------------
        End If
    End If
    
    If Button.index = 6 Then Cancelar
    
    If Button.index = 8 Then Filtrar
    
    If Button.index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstVent.Filter = ""
    End If
    If Button.index = 10 Then Buscar
    If Button.index = 11 Then CambiarMes
    
'        Dim xFchReg As String
'        Dim Rpta As Integer
'        mMesActivo = SeleccionaMes(xCon)
'        xFchReg = "01/" + Format(mMesActivo, "00") + "/" + Trim(AnoTra)
'        OpcionesPeriodo
'        CargarRSTCom xFchReg, mMesActivo
'
'        If RstVent.State <> 0 Then
'            Set Dg1.DataSource = RstVent
'            If RstVent.RecordCount = 0 Then
'                If mMesActivo <> 0 And mMesActivo <> 13 Then
'                    Rpta = MsgBox("No se ha registrado ninguna venta, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
'                    If Rpta = vbYes Then
'                        Nuevo
'                    Else
'                        Unload Me
'                    End If
'                End If
'            Else
'                Dg1.SetFocus
'            End If
'        End If
'    End If
    
    If Button.index = 13 Then
        Imprimir
    End If
    
    If Button.index = 15 Then
        Set RstVent = Nothing
        Unload Me
    End If
End Sub

Sub OpcionesPeriodo()
    Dim NomMes As String
''    Dim Cerrado As Boolean
    Dim xFechaMes As String
    Dim xFchIni, xFchFin As Date
    Dim Rpta As Integer
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
''    Cerrado = Busca_Codigo(mMesActivo, "id", "cerrado", "con_meses", "N", xCon)
    
''    If Cerrado = True Then
''        Toolbar1.Buttons(1).Visible = False
''        Toolbar1.Buttons(2).Visible = False
''        Toolbar1.Buttons(3).Visible = False
''        Toolbar1.Buttons(4).Visible = False
''    Else
''        Toolbar1.Buttons(1).Visible = True
''        Toolbar1.Buttons(2).Visible = True
''        Toolbar1.Buttons(3).Visible = True
''        Toolbar1.Buttons(4).Visible = True
''    End If
    
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, 2, mMesActivo, fCierrePeriodo, xCon
    '------------------------------------------------------------------------------------------

    
    If mMesActivo <> 0 And mMesActivo <> 13 Then
        xFechaMes = "01/" + Trim(Format(mMesActivo, "00")) + "/" + Trim(Format(Year(Date), "0000"))
        xFchIni = xFechaMes
        xFchFin = Format(HallaDiasMes(CDate(xFechaMes)), "00") + "/" + Mid(xFechaMes, 4, 7)
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.index = 2 Then
        'MODIFICACION DE DOCUMENTOS
        If ButtonMenu.index = 1 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opcion", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If RstVent("anulado") = -1 Then
                MsgBox "No puede modificar " & RstVent("nomdoc") & " anulado proceda a restaurarlo", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
                Exit Sub
            Else
                Modificar
            End If
        End If
        
        'RESTAURAR DOCUMENTOS
        If ButtonMenu.index = 2 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opcion", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If RstVent("anulado") = -1 Then ' SI EL DOCUMENTO ESTA ANULADO
                RestaurarFactura
            End If
        End If
    End If
  
    If ButtonMenu.Parent.index = 3 Then
        If ButtonMenu.index = 1 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opcion", vbInformation, Me.Caption
                Exit Sub
            End If
            Anular
        End If
        If ButtonMenu.index = 2 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opcion", vbInformation, Me.Caption
                Exit Sub
            End If
            
            Eliminar
        End If
        
        If ButtonMenu.index = 3 Then EmitirAnulada
    End If
    
    If ButtonMenu.Parent.index = 13 Then
        If ButtonMenu.index = 1 Then Imprimir
        If ButtonMenu.index = 2 Then
            Exportar
        End If
    End If

End Sub

Sub Exportar()
'    Dim oExport As New SGI2_funciones.formularios
'    Dim Rst As New ADODB.Recordset
'
'    Dim xCampos(11, 3) As String
'
'    '0::Nombre a Mostrar;
'    '1::nombre de Campo del Rst;
'    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
'    '3::ancho de columna
'    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
'    xCampos(0, 0) = "Nº Registro":        xCampos(0, 1) = "numreg1":      xCampos(0, 2) = 2:   xCampos(0, 3) = "500"
'    xCampos(1, 0) = "Tip. Doc.":          xCampos(1, 1) = "abrev":        xCampos(1, 2) = 2:   xCampos(1, 3) = "5000"
'    xCampos(2, 0) = "Nº Documento":       xCampos(2, 1) = "numerodoc":    xCampos(2, 2) = 2:   xCampos(2, 3) = "450"
'    xCampos(3, 0) = "Fch. Emi.":          xCampos(3, 1) = "fchdoc1":      xCampos(3, 2) = 1:   xCampos(3, 3) = "1100"
'    xCampos(4, 0) = "Fch. Ven.":          xCampos(4, 1) = "fchven1":      xCampos(4, 2) = 1:   xCampos(4, 3) = "4500"
'    xCampos(5, 0) = "Cliente":            xCampos(5, 1) = "nombre":       xCampos(5, 2) = 2:   xCampos(5, 3) = "950"
'    xCampos(6, 0) = "Moneda":             xCampos(6, 1) = "simbolo":      xCampos(6, 2) = 1:   xCampos(6, 3) = "860"
'    xCampos(7, 0) = "Imp. Bru":           xCampos(7, 1) = "impbru1":      xCampos(7, 2) = 0:   xCampos(7, 3) = "650"
'    xCampos(8, 0) = "I.G.V.":             xCampos(8, 1) = "impigv1":      xCampos(8, 2) = 0:   xCampos(8, 3) = "650"
'    xCampos(9, 0) = "Importe":            xCampos(9, 1) = "imptotdoc1":   xCampos(9, 2) = 0:   xCampos(9, 3) = "650"
'    xCampos(10, 0) = "Saldo":             xCampos(10, 1) = "impsal1":     xCampos(10, 2) = 0:  xCampos(10, 3) = "650"
'
'
'    Set Rst = RstVent.Clone
'    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "REGISTRO DE VENTAS", LblMes.Caption, "", "ventas", Rst, xCampos
'    Set oExport = Nothing
'    Set Rst = Nothing
'    Dg1.Refresh

    
End Sub

Private Sub TxtAlmacen2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtAlmacen2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusAlmacen2_Click
    End If
End Sub

Private Sub TxtConPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtConPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCondicion_Click
    End If
End Sub

Private Sub TxtConPag_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtConPag.Text) = "" Then
        'SendKeys vbTab
        Exit Sub
    End If
    Dim xRs1 As New ADODB.Recordset

    RST_Busq xRs1, "SELECT * FROM mae_condpago WHERE id = " & NulosN(TxtConPag.Text) & "", xCon

    If xRs1.RecordCount = 0 Then
        TxtConPag.Text = ""
        LblCondPag.Caption = ""
        TxtFchVen.Valor = ""
    Else
        If TxtFchDoc.Valor = "" Then
            MsgBox "No ha especificado la fecha de emision del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtConPag.Text = ""
            LblCondPag.Caption = ""
            Exit Sub
        End If
        LblCondPag.Caption = Trim(xRs1("descripcion"))
        TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + xRs1("numdia")
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdMotNotCre_Click
    End If
End Sub

Private Sub TxtDocRefCredi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDocRefCredi_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDocRef_Click
    End If
    If KeyCode = 46 Then
        TxtDocRefCredi.Text = ""
        LblIdDocRef.Caption = ""
    End If
End Sub

Private Sub TxtIdAlm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdAlm_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtIdAlm.Text) <> "" Then
        LblAlmacen.Caption = Busca_Codigo(NulosN(TxtIdAlm.Text), "id", "descripcion", "alm_almacenes", "N", xCon)
        If LblAlmacen.Caption = "" Then
            TxtIdAlm.Text = ""
        End If
    End If
End Sub

Private Sub TxtIdDocGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDocGen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDocGen_Click
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtIdMon.Text) = "" Then
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    
    'buscamos el codigo de la moneda         digitada
    RST_Busq xRs1, "SELECT * FROM mae_moneda WHERE id = " & NulosN(TxtIdMon.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtIdMon.Text = ""
        LblMoneda.Caption = ""
    Else
        LblMoneda.Caption = Trim(xRs1("descripcion"))
        
        If Trim(TxtIdMon.Text) = "1" Then
            LblTipCam.Visible = False
            LblTipoCambio.Visible = False
        Else
            If TxtFchDoc.Valor = "" Then
                MsgBox "No ha especificado la fecha del documento, no se puede determinar " & Chr(13) _
                    & "la fecha del tipo de cambio para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                
                TxtIdMon.Text = ""
                LblMoneda.Caption = ""
                TxtFchDoc.SetFocus
                Exit Sub
            End If
            
            LblTipCam.Visible = True
            LblTipoCambio.Visible = True
            'Set xRs1 = Nothing
            'Set xRs1 = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = CDATE('" & TxtFchDoc.Valor & "')", xCon)
           '
            'If xRs1.RecordCount = 1 Then
            '    LblTipoCambio.Caption = Format(xRs1("impven"), "0.000")
            '    ValTipCam = xRs1("impven")
            'Else
            '    LblTipoCambio.Caption = "0.00"
            '    ValTipCam = 0
            'End If
            LblTipoCambio.Caption = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
        End If
    End If
    Set xRs1 = Nothing

End Sub

Private Sub TxtIdTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusIdTipDocRef_Click
    End If
End Sub

Private Sub TxtIdTipDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosN(TxtIdTipDoc.Text) = 0 Then
        TxtIdTipDoc.Text = ""
        LblDescTipDocRef.Caption = ""
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    
    RST_Busq xRs1, "SELECT * FROM mae_docreferencia WHERE id = " & Val(TxtIdTipDoc.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtIdTipDoc.Text = ""
        LblDescTipDocRef.Caption = ""
        TxtNumDocRef.Text = ""
        LblIdDocRef2.Caption = ""

    Else
        LblDescTipDocRef.Caption = Trim(xRs1("descripcion"))
        TxtNumDocRef.Text = ""
        LblIdDocRef2.Caption = ""
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtIdVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdVen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusVen_Click
    End If
End Sub

Private Sub TxtIdVen_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    
    If NulosC(TxtIdVen.Text) <> "" Then
        Set RstTmp = BuscaConCriterio("SELECT vta_vendedores.*, UCase(pla_empleados!apepat)+' '+ UCase(pla_empleados!apemat)+', '+pla_empleados!nom AS apenom " _
            & " FROM vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id WHERE vta_vendedores.id = " & NulosN(TxtIdVen.Text) & "", xCon)
        
        If RstTmp.RecordCount <> 0 Then
            LblNomVen.Caption = RstTmp("apenom")
        Else
            TxtIdVen.Text = ""
            LblNomVen.Caption = ""
        End If
    End If

    Set RstTmp = Nothing
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumDoc.Text) <> "" Then
    
        TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
        
        Dim Rst As New ADODB.Recordset
        Dim nSQL As String
        '--ver si existe el numero de doc
        If QueHace <> 1 Then nSQL = " and vta_ventas.id <> " & NulosN(RstVent("id"))
        
        RST_Busq Rst, "SELECT vta_ventas.numser, vta_ventas.numdoc, vta_ventas.fchdoc, mae_cliente.nombre, Left([vta_ventas].[numreg],2) & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & Right([vta_ventas].[numreg],4) AS registro FROM (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            & " WHERE (((vta_ventas.numser)='" & Trim(TxtNumSer.Text) & "') AND ((vta_ventas.numdoc)='" & TxtNumDoc.Text & "'))" & nSQL, xCon
                
        If Rst.RecordCount <> 0 Then
            '--poner el nuevo numero doc
            TxtNumSer_Validate True
            MsgBox "El número de documento de venta ya existe " & vbCr & "Nº Registro: " & NulosC(Rst("registro")) & vbCr & "Fecha Doc.   " & NulosC(Rst("fchdoc")) & vbCr & "Cliente:         " & NulosC(Rst("nombre")) & vbCr & "Será reemplazado por " + Trim(TxtNumSer.Text) + "-" + Trim(TxtNumDoc.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
        Set Rst = Nothing
        
    End If
End Sub

Private Sub TxtNumDocGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        TxtNumDocGen_Validate True
    End If
End Sub

Private Sub TxtNumDocGen_Validate(Cancel As Boolean)
    If TxtNumDocGen.Text <> "" Then
        TxtNumDocGen.Text = Format(TxtNumDocGen.Text, "0000000000")
    End If
End Sub

Private Sub TxtNumDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Fg1.Rows > 1 Then
            Fg1.Col = 1
            Fg1.SetFocus
        Else
            If CmdAddItem.Enabled = True Then CmdAddItem.SetFocus
        End If
    End If
End Sub

Private Sub TxtNumDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDocRef2_Click
    End If
    If KeyCode = 46 Then
        TxtDocRefCredi.Text = ""
        LblIdDocRef.Caption = ""
    End If
End Sub

Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumRuc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCli_Click
    End If
End Sub

Private Sub TxtNumRuc_Validate(Cancel As Boolean)
    If NulosC(TxtNumRuc.Text) = "" Then
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    RST_Busq xRs1, "SELECT * FROM mae_cliente WHERE numruc like '" & TxtNumRuc.Text & "%' ORDER BY numruc", xCon
    If xRs1.RecordCount <> 0 Then
        TxtNumRuc.Text = xRs1("numruc")
        LblNomCli.Caption = xRs1("nombre")
        LblIdCliente.Caption = xRs1("id")
        'TxtIdVen.Text = NulosN(xRs1("idven"))
        'TxtIdVen_Validate True
    Else
        TxtNumRuc.Text = ""
        LblNomCli.Caption = ""
        LblIdCliente.Caption = ""
        TxtIdVen.Text = ""
        Lblvendedor.Caption = ""
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If NulosC(TxtTipDoc.Text) = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.Text = ""
        TxtTipDoc.SetFocus
        TxtNumSer.Text = ""
        Exit Sub
    End If
        
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusNumSer_Click
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    Dim Rstdoc As New ADODB.Recordset
    If NulosC(TxtNumSer.Text) = "" Then
        Exit Sub
    Else
        If QueHace <> 1 Then Exit Sub
        
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        Dim Rst As New ADODB.Recordset
        
        RST_Busq Rst, "SELECT top 1 vta_ventas.numdoc AS numero from vta_ventas WHERE (((vta_ventas.numser)='" & NulosC(TxtNumSer.Text) & "') AND ((vta_ventas.tipdoc)=1))" _
            & " ORDER BY vta_ventas.numdoc DESC ", xCon

        If Rst.RecordCount = 0 Then
            TxtNumDoc.Text = "0000000001"
        Else
            Rst.MoveFirst
            TxtNumDoc.Text = Format(NulosN(Rst("numero")) + 1, "0000000000")
        End If
        Set Rst = Nothing
    End If
End Sub

Private Sub TxtNumSerGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSerGen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusSerGen_Click
    End If
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc_Click
    End If
End Sub

Sub EmitirAnulada()
    TabOne1.CurrTab = 0
    ActivarEntorno
    
    Fraseldoc.Left = 3315
    Fraseldoc.Top = 2505
    TxtAlmacen2.Text = ""
    TxtIdDocGen.Text = ""
    TxtNumSerGen.Text = ""
    TxtNumDocGen.Text = ""
    LblIdDocumentoGen.Caption = ""
    Fraseldoc.Visible = True
    TxtAlmacen2.SetFocus
End Sub

Function Grabar() As Boolean
    Dim A As Integer
    Dim xFchReg As String
    Dim xFchFin As String
    
    xFchReg = "01/" + Format(mMesActivo, "00") + "/" + Trim(AnoTra)
    A = HallaDiasMes(CDate(xFchReg))
    xFchFin = Trim(Str(A)) + "/" + Format(mMesActivo, "00") + "/" + Trim(AnoTra)
    
    If NulosN(TxtTipDoc.Text) <> 0 Then
        If xCuentaDoc = 0 Then
            MsgBox "No se ha asignado una cuenta contable al documento " + LblNomDoc.Caption & Chr(13) _
                & "Asignele una cuenta en el menu Contabilidad opcion Asignar Ctas. Contables a documentos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    End If
    
    If NulosN(TxtTipDoc.Text) <> 0 Then
        If xIdCuenTasa = 0 Then
            MsgBox "El impuesto asignado al documento " + LblNomDoc.Caption & Chr(13) & " no tiene cuenta contable" & Chr(13) _
                & "Asignele una cuenta en el menu Contabilidad opcion Maestro de Impuestos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    End If
    
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, 10)) = 0 Then
            MsgBox "No se le ha asignado una cuenta contable para venta al item : " & Chr(13) _
                & Fg1.TextMatrix(A, 1) & Chr(13) _
                & "Asignele una cuenta en el menu Almacen opcion Mantenimiento Items de Compra y Venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    Next A
    
    If TxtTipItem.Text = "" Then
        MsgBox "No ha especificado el tipo de item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipItem.SetFocus
        Exit Function
    End If
    
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento de venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    
'    If NulosN(TxtTipDoc.Text) = 7 Then
'        If NulosC(TxtDocRef.Text) = "" Then
'            MsgBox "No ha especificado el o los documentos de referencia para la Nota de credito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            TxtDocRef.SetFocus
'            Exit Function
'        End If
'    End If
    
    If TxtNumRuc.Text = "" Then
        MsgBox "No ha especificado cliente de la venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
    If TxtNumSer.Text = "" Or TxtNumDoc.Text = "" Then
        MsgBox "No ha especificado el numero del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If
    
    If TxtFchDoc.Valor = "" Then
        MsgBox "No ha especificado la fecha de emision del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchDoc.SetFocus
        Exit Function
    End If
    
    If CDate(TxtFchDoc.Valor) < CDate(xFchReg) Then
        MsgBox "No se puede grabar este documento en el periodo actual la fecha de emision es menor a : " + xFchReg, vbInformation + vbOKOnly + vbDefaultButton1
        TxtFchDoc.SetFocus
        Exit Function
    End If
    
    If CDate(TxtFchDoc.Valor) > CDate(xFchFin) Then
        MsgBox "No se puede grabar este documento en el periodo actual la fecha de emision es mayor a : " + xFchFin, vbInformation + vbOKOnly + vbDefaultButton1
        TxtFchDoc.SetFocus
        Exit Function
    End If
    
    
    If TxtFchVen.Valor = "" Then
        MsgBox "No ha especificado la fecha de vencimiento del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchVen.SetFocus
        Exit Function
    End If
    
    If CDate(TxtFchVen.Valor) < CVDate(TxtFchDoc.Valor) Then
        MsgBox " La fecha de vencimiento del documento no puede ser menor a la fecha de emision", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchVen.SetFocus
        Exit Function
    End If
    
    If TxtConPag.Text = "" Then
        MsgBox "No ha especificado la condicion de pago del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtConPag.SetFocus
        Exit Function
    End If
    
    If TxtIdMon.Text = "" Then
        MsgBox "No ha especificado la moneda del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdAlm.Text) = 0 Then
        MsgBox "No ha especificado el nombre del almacen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdAlm.SetFocus
        Exit Function
    End If
       
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado items para la venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    If QueHace = 1 Then 'Validamos si existe el numero del documento en modo adicion
        Dim RstCab As New ADODB.Recordset
    
        RST_Busq RstCab, "SELECT * FROM vta_ventas WHERE tipdoc =" & NulosN(TxtTipDoc.Text) & " AND numser ='" & TxtNumSer.Text & "' AND numdoc = '" & TxtNumDoc.Text & "' ", xCon
    
        If RstCab.RecordCount > 0 Then
            MsgBox "El Nro de documento ha sido registrado por otro usuario se grabara con otro numero", vbInformation, Me.Caption
            TxtNumDoc.Text = HallaNumdocVenta(NulosN(TxtTipDoc.Text), TxtNumSer.Text, xCon)
        End If
        Set RstCab = Nothing
    End If
    
    If NulosN(TxtIdTipDoc.Text) <> 0 Then
        If NulosN(LblIdDocRef2.Caption) = 0 Then
            MsgBox "No ha especificado el documento de referencia para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtNumDocRef.SetFocus
            Exit Function
        End If
    End If
    
    If NulosN(TxtTipDoc.Text) = 7 Then
        If NulosC(TxtDocRefCredi.Text) = "" Then
            MsgBox "No ha especificado el documento al que hace referencia la nota de credito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    
    Dim RstDeta2 As New ADODB.Recordset
    Dim RstActPro As New ADODB.Recordset
    
    Dim RstDet As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim xIdCuen As Integer
    Dim xTotal As Double
    Dim xSaldo As Double
    
    Dim xidtipven As String 'Determina si la venta es de tipo exportacion
    Dim xNumAsiento As String
    
    Dim xId As Double
    Dim X As Integer
    Dim P As Integer
    
On Error GoTo LaCague
    swguiafact = 1
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("vta_ventas", xCon, "id")
        xNumAsiento = NuevoNumAsiento(2, mMesActivo, xCon)
        
        RST_Busq RstCab, "SELECT TOP 1 * FROM vta_ventas", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
        
        If NulosN(TxtTipDoc.Text) = 7 Then
            xSaldo = 0
        Else
            xSaldo = NulosN(TxtTotal.Text)
        End If
        
    Else
        xId = RstVent("id")
        RST_Busq RstCab, "SELECT * FROM vta_ventas WHERE id = " & xId & "", xCon
        
        'Eliminamos el stock agregado con la venta
        RST_Busq RstDeta2, "SELECT vta_ventasdet.* From vta_ventasdet WHERE (((vta_ventasdet.idvta)= " & xId & "))", xCon

        If RstDeta2.RecordCount <> 0 Then
            RstDeta2.MoveFirst
            For A = 1 To RstDeta2.RecordCount
                RST_Busq RstActPro, "SELECT alm_inventario.id, alm_inventario.stckact  From alm_inventario WHERE ((alm_inventario.id=" & RstDeta2("iditem") & "))", xCon
                If RstActPro.RecordCount = 1 Then
                    RstActPro("stckact") = RstActPro("stckact") + RstDeta2("canpro")
                    RstActPro.Update
                End If
                Set RstActPro = Nothing
            Next A
        End If
        Set RstDeta2 = Nothing
        
        'Eliminamos el detalle de la venta
        xCon.Execute "DELETE * FROM vta_ventasdet WHERE idvta = " & xId & ""
        
        
        RST_Busq RstDia, "SELECT * FROM con_diario WHERE idmes = " & Format(CDate(TxtFchDoc.Valor), "mm") & " AND " _
            & " idlib = 2 AND idmov = " & xId & " And iddoc = " & NulosN(TxtTipDoc) & "", xCon
            
        If RstDia.RecordCount <> 0 Then
            xNumAsiento = RstDia("numasi")
        Else
            xNumAsiento = NuevoNumAsiento(2, mMesActivo, xCon)
        End If
        
        Set RstDia = Nothing
        
       'Eliminamos el asiento contable
        xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & Format(CDate(TxtFchDoc.Valor), "mm") & " AND " _
            & " idlib = 2 AND idmov = " & xId & " And iddoc = " & NulosN(TxtTipDoc) & ""
            
        
        If NulosN(TxtTipDoc.Text) = 7 Then
            xSaldo = 0
        Else
            xSaldo = NulosN(TxtTotal.Text)
        End If
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM vta_ventasdet", xCon
    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    
    
    '----------------------------------------------------------------------------
    mIdRegistro = xId
    '----------------------------------------------------------------------------
    RstCab("idlib") = 2
    RstCab("idtipo") = NulosN(TxtTipItem.Text)
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("idcli") = NulosN(LblIdCliente.Caption)
    RstCab("numser") = TxtNumSer.Text
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("fchdoc") = TxtFchDoc.Valor
    RstCab("fchven") = TxtFchVen.Valor
    RstCab("idconpag") = NulosN(TxtConPag.Text)
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("impbru") = NulosN(TxtBruto.Text)
    RstCab("impinaf") = NulosN(txtinafecto.Text)
    RstCab("impigv") = NulosN(TxtIGV.Text)
    RstCab("impisc") = NulosN(txtisc.Text)
    RstCab("impotr") = 0 'NulosN(me.txtir Txtotr...Text)
    RstCab("imptotdoc") = NulosN(TxtTotal.Text)
    RstCab("impsal") = xSaldo 'NulosN(TxtTotal.Text)
    RstCab("idalm") = NulosN(TxtIdAlm.Text)
    
    RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    
    If OptDes1.Value = True Then RstCab("tipdes") = 1
    If OptDes2.Value = True Then RstCab("tipdes") = 2
    
    If CONTABILIZAR = True Then
        RstCab("numreg") = Trim(Format(Str(mMesActivo), "00")) + xNumAsiento
    End If
    
    RstCab("idven") = NulosN(TxtIdVen.Text)
    
    If NulosN(TxtTipDoc.Text) = 7 Then
        RstCab("iddocref") = NulosN(LblIdDocRef.Caption)
        RstCab("idmotnotcre") = NulosN(LblIdConNC.Caption)
    End If
    
    RstCab("anulado") = 0
    
    'grabamos el documento de referencia para la venta (orden venta, orden de despacho etc)
    RstCab("idtipdocref") = NulosN(TxtIdTipDoc.Text)
    RstCab("iddocref2") = NulosN(LblIdDocRef2.Caption)
    
    'Determinamos si es una exportacion
    For A = 1 To Fg1.Rows - 1
        xidtipven = NulosN(Fg1.TextMatrix(A, 11))
    Next A
    
    If xidtipven = 2 And NulosN(TxtIGV.Text) = 0 Then 'si esta venta exportacion
        RstCab("idtipven") = 2
    Else
        RstCab("idtipven") = 0 'en el cual puede ser venta afecta o inafecta para el registro de
                               'de ventas se valida por programa ver tabla mae_tipoventa
    End If
    
    If optsinguia.Value = True Then RstCab("oriitem") = 1
    If optconguia.Value = True Then RstCab("oriitem") = 2
    If optconcotizacion.Value = True Then RstCab("oriitem") = 3
    
    'RstCab("docref") = NulosC(TxtDocRef.Text)
    
    'Actualizamos el saldo del documento
    ActualizaSaldoDoc NulosN(LblIdDocRef.Caption), 2, NulosN(TxtTotal.Text)

    RstCab.Update
    
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idvta") = xId
        RstDet("iditem") = NulosN(Fg1.TextMatrix(A, 8))
        RstDet("idunimed") = NulosN(Fg1.TextMatrix(A, 9))
        If NulosN(Fg1.TextMatrix(A, 6)) <> 0 Then
            RstDet("preuni") = NulosN(Fg1.TextMatrix(A, 6))
        Else
            RstDet("preuni") = NulosN(Fg1.TextMatrix(A, 4))
        End If
        RstDet("valdes") = NulosN(Fg1.TextMatrix(A, 5))
        RstDet("canpro") = NulosN(Fg1.TextMatrix(A, 3))
        RstDet("imptot") = NulosN(Fg1.TextMatrix(A, 7))
        RstDet("tasaper") = NulosN(Fg1.TextMatrix(A, 12))
        RstDet("preunibru") = NulosN(Fg1.TextMatrix(A, 4))
        RstDet.Update
    Next A
   
    'ACTUALIZAMOS EL STOCK
    If optsinguia.Value = True Then
        For A = 1 To Fg1.Rows - 1
            xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckact = ( [alm_inventario]![stckact]-" & NulosN(Fg1.TextMatrix(A, 3)) & ")" _
                & " WHERE (((alm_inventario.id)=" & NulosN(Fg1.TextMatrix(A, 8)) & "))"
        Next A
    End If
        
    If CONTABILIZAR = True Then
        '---------------------------------------
        'Grabamos el libro diario del movimiento
        xAño = AnoTra
        RstDia.AddNew
        'grabamos el documento de venta en la tabla diario
        RstDia("año") = xAño
        RstDia("idmes") = mMesActivo
        RstDia("idlib") = 2
        RstDia("iddoc") = NulosN(TxtTipDoc)
        RstDia("idmov") = xId
        RstDia("numasi") = xNumAsiento
        RstDia("tc") = ValTipCam
        RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
        RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
        RstDia("idcue") = xCuentaDoc
        
        If NulosN(TxtTipDoc.Text) <> 7 Then
            If TxtIdMon.Text = "1" Then
                RstDia("impdebsol") = NulosN(TxtTotal.Text)
                RstDia("impdebdol") = 0
            Else
                RstDia("impdebsol") = NulosN(TxtTotal.Text) * NulosN(LblTipoCambio.Caption)
                RstDia("impdebdol") = NulosN(TxtTotal.Text)
            End If
        Else
            If TxtIdMon.Text = "1" Then
                RstDia("imphabsol") = NulosN(TxtTotal.Text)
                RstDia("imphabdol") = 0
            Else
                RstDia("imphabsol") = NulosN(TxtTotal.Text) * NulosN(LblTipoCambio.Caption)
                RstDia("imphabdol") = NulosN(TxtTotal.Text)
            End If
        End If
        RstDia.Update
        
        '-----------------------------------------------------
        'grabamos el impuesto si la operacion esta afecta a el
        If NulosN(TxtIGV) > 0 Then
            RstDia.AddNew
            RstDia("idmes") = mMesActivo
            RstDia("año") = xAño
            RstDia("idlib") = 2
            RstDia("iddoc") = NulosN(TxtTipDoc)
            RstDia("idmov") = xId
            RstDia("numasi") = xNumAsiento
            RstDia("tc") = ValTipCam
            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
            RstDia("idcue") = xIdCuenTasa
            
            If NulosN(TxtTipDoc.Text) <> 7 Then
                If TxtIdMon.Text = "1" Then
                    RstDia("imphabsol") = NulosN(TxtIGV.Text)
                    RstDia("imphabdol") = 0
                Else
                    RstDia("imphabsol") = NulosN(TxtIGV.Text) * NulosN(LblTipoCambio.Caption)
                    RstDia("imphabdol") = NulosN(TxtIGV.Text)
                End If
            Else
                If TxtIdMon.Text = "1" Then
                    RstDia("impdebsol") = NulosN(TxtIGV.Text)
                    RstDia("impdebdol") = 0
                Else
                    RstDia("impdebsol") = NulosN(TxtIGV.Text) * NulosN(LblTipoCambio.Caption)
                    RstDia("impdebdol") = NulosN(TxtIGV.Text)
                End If
            End If
            RstDia.Update
        End If
    
       '********Rutina para que extraer la base imponible sea afecta o inafecta
        Dim xFun As New eps_librerias.FuncionesData
        Dim xCampos(2, 3) As String
        Dim rstdocus As New ADODB.Recordset
    
        xCampos(0, 0) = "cuenta":     xCampos(0, 1) = "C":      xCampos(0, 2) = "12"
        xCampos(1, 0) = "Importe":    xCampos(1, 1) = "D":      xCampos(1, 2) = "2"
        
        Set rstdocus = xFun.CrearRstTMP(xCampos)
        rstdocus.Open
    
        For X = 1 To Fg1.Rows - 1
            xIdCuen = Trim(Fg1.TextMatrix(X, 10))
            xTotal = NulosN(Fg1.TextMatrix(X, 7))
            
            If rstdocus.RecordCount <> 0 Then rstdocus.MoveFirst
            rstdocus.Find ("cuenta ='" & xIdCuen & "'")
            If rstdocus.EOF = True Then
                rstdocus.AddNew
                rstdocus("cuenta") = Trim(Fg1.TextMatrix(X, 10))
                rstdocus("importe") = xTotal
                rstdocus.Update
            Else
                rstdocus("importe") = rstdocus("importe") + xTotal
                rstdocus.Update
            End If
        Next X
        
        '------------------
        'Grabamos el diario
        If rstdocus.RecordCount > 0 Then
            rstdocus.MoveFirst
            
            Do While Not rstdocus.EOF
                RstDia.AddNew
                RstDia("año") = xAño
                RstDia("idmes") = mMesActivo               'LLAVE - CODIGO DEL MES
                RstDia("idlib") = 2                  'LLAVE - CODIGO DEL LIBRO
                RstDia("iddoc") = NulosN(TxtTipDoc.Text)     'LLAVE - CODIGO DEL DOCUMENTO
                RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
                RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
                RstDia("tc") = ValTipCam
                RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
                RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
                RstDia("idcue") = NulosN(rstdocus("cuenta"))
                
                If NulosN(TxtTipDoc.Text) <> 7 Then
                    If TxtIdMon.Text = "1" Then
                        RstDia("imphabsol") = rstdocus("importe")
                        RstDia("imphabdol") = 0
                    Else
                        RstDia("imphabsol") = rstdocus("importe") * NulosN(LblTipoCambio.Caption)
                        RstDia("imphabdol") = rstdocus("importe")
                    End If
                Else
                    If TxtIdMon.Text = "1" Then
                        RstDia("impdebsol") = rstdocus("importe")
                        RstDia("impdebdol") = 0
                    Else
                        RstDia("impdebsol") = rstdocus("importe") * NulosN(LblTipoCambio.Caption)
                        RstDia("impdebdol") = rstdocus("importe")
                    End If
                End If
                RstDia.Update
                rstdocus.MoveNext
            Loop
        End If
    End If
     
   'Actualizamos en el campo Iddocven de la tabla Guias el Id del Documento de Venta para relacionarlo Factura -Guia
    If optconguia.Value = True Then
        For X = 1 To Fg4.Rows - 1
            xCon.Execute " UPDATE vta_guia SET vta_guia.iddocven = " & xId & " WHERE vta_guia.id = " & NulosN(Fg4.TextMatrix(X, 3)) & ""
        Next X
    End If
    
    'Actualizamos en el campo Iddocven de la tabla cotizaciones el Id del Documento de Venta para relacionarlo Factura - Cotizacion
    If optconcotizacion.Value = True Then
        For X = 1 To Fg4.Rows - 1
            xCon.Execute " UPDATE vta_cotizacion SET vta_cotizacion.iddocven = " & xId & ", vta_cotizacion.idest = 3 WHERE vta_cotizacion.id = " & NulosN(Fg4.TextMatrix(X, 3)) & ""
        Next X
    End If
    
    
    'si esta afecto a la detraccion
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT mae_detraccion.id, mae_detraccion.descripcion, mae_detraccion.tasa, alm_inventario.iddet " _
        & " FROM alm_inventario LEFT JOIN mae_detraccion ON alm_inventario.iddet = mae_detraccion.id " _
        & " WHERE ((alm_inventario.id= " & Val(Fg1.TextMatrix(Fg1.Row, 8)) & "))", xCon

    If Rst.RecordCount <> 0 Then
        If Rst("iddet") <> 0 Then
            MsgBox "Se ha detectado que la venta registrada esta afecta al regimen de la Detraccion " + Chr(13) _
                & "Decripcion : " + Rst("descripcion") + Chr(13) _
                & "tasa : " + Format(Rst("tasa"), "0.00") + "%", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                
            Dim RstDeta As New ADODB.Recordset
            Dim xId2 As Integer
            
            If QueHace = 1 Then
                xId2 = HallaCodigoTabla("con_detraccion", xCon, "id")
                RST_Busq RstDeta, "SELECT * FROM con_detraccion", xCon
                RstDeta.AddNew
                RstDeta("id") = xId2
            Else
                RST_Busq RstDeta, "SELECT con_detraccion.* From con_detraccion " _
                    & " WHERE (((con_detraccion.iddoc)=" & xId & "))", xCon
            End If
            
            If RstDeta.RecordCount = 0 Then
                'este procedimiento es solo para cuando se este modificando una compra afecta a la detraccion y no se le haya hecho la detraccion a la hora de ingresar la compra
                xId2 = HallaCodigoTabla("con_detraccion", xCon, "id")
                RstDeta.AddNew
                RstDeta("id") = xId2
            End If
            
            RstDeta("iddet") = Rst("iddet")
            RstDeta("por") = Rst("tasa")
            RstDeta("iddoc") = xId
            RstDeta("idmon") = NulosN(TxtIdMon.Text)
            RstDeta("tipo") = 2   'especificamos que es una venta
            RstDeta("fchmov") = Date
            RstDeta("Glosa") = ""
            RstDeta("imp") = Format((NulosN(TxtTotal.Text) * (Rst("tasa") / 100)), "0.00")
            RstDeta("numdet") = "SIN NUMERO"
            RstDeta.Update
        End If
    End If
    Dim nSQL As String
    nSQL = "UPDATE (vta_ventas INNER JOIN con_diario ON vta_ventas.id = con_diario.idmov) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
        + vbCr + " SET con_diario.fchdoc=vta_ventas.fchdoc, con_diario.idmon=vta_ventas.idmon, con_diario.ridlib = 2, con_diario.ridtipper = 2, con_diario.ridper = [vta_ventas].[idcli], con_diario.rtipdoc = [vta_ventas].[tipdoc], con_diario.rfchope = [vta_ventas].[fchdoc], con_diario.rnumerodoc = IIf([vta_ventas].[numser] Is Null Or [vta_ventas].[numser]='','',[vta_ventas].[numser] & '-') & [vta_ventas].[numdoc], con_diario.rglosaope = [vta_ventas].[glosa] & '', con_diario.rregistro = Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) " _
        + vbCr + " WHERE con_diario.idlib=2 and  con_diario.idmov= " & xId
    
    xCon.Execute nSQL
    
    'Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, 2, QueHace, xHorIni, Time, Date, xCon, CDbl(xId)
    
    xCon.CommitTrans
    
    MsgBox "La " & Trim(LblNomDoc) & " se registró con éxito" & vbCr & "Registro Nº: " & Format(mMesActivo, "00") & "14" & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    
    
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    Grabar = True
    Exit Function
    
LaCague:
   ' Resume
    xCon.RollbackTrans
    Set rstdocus = Nothing
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDia = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Function HallaNumAsiento(Mes As Integer) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_diario.idmes, con_diario.idlib, con_diario.numasi From con_diario " _
        & " WHERE (((con_diario.idmes)=" & Mes & ") AND ((con_diario.idlib)=2)) ORDER BY numasi", xCon
    
    If Rst.RecordCount = 0 Then
        HallaNumAsiento = "0001"
    Else
        Rst.MoveLast
        HallaNumAsiento = Format(NulosN(Rst("numasi")) + 1, "0000")
    End If
    Exit Function
End Function

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtTipDoc.Text) = "" Then
        Exit Sub
    End If
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
       
    RST_Busq xRs, "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, mae_impuestos.Abrev AS abreimp, " _
        & " mae_impuestos.idcuenvta AS cuentaimp, alm_numseries.numser, alm_numseries.idtipdoc " _
        & " FROM alm_numseries LEFT JOIN (mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id) ON alm_numseries.idtipdoc = mae_documento.id " _
        & " WHERE (((alm_numseries.idalm)=" & NulosN(TxtIdAlm.Text) & ") AND ((alm_numseries.idtipdoc)=" & NulosN(TxtTipDoc.Text) & "))", xCon
        
    If xRs.RecordCount = 0 Then
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
    Else
        TxtTipDoc.Text = xRs("id")
        LblNomDoc.Caption = xRs("descripcion")
        TasaImpuesto = NulosN(xRs("tasa"))
        
        xIdCuenTasa = NulosN(xRs("cuentaimp"))
        
        Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & NulosN(TxtIdMon) & " and tipope = -1", xCon)
        If xRs2.RecordCount > 0 Then
            xCuentaDoc = NulosN(xRs2("idcuen"))
        End If
        
        Set xRs2 = Nothing
        
        LblRotulo.Caption = Trim(NulosC(xRs("abreimp"))) + " (         )"
        LblIgvTasa.Caption = Format(Trim(Str(TasaImpuesto)), "0.00")
        
        If xRs("id") = 7 Then
            Label33.Visible = True
            TxtDocRefCredi.Visible = True
            CmdBusDocRef.Visible = True
            FraRetencion.Visible = False
            Frame5.Left = 5175
            Frame5.Top = 2820
            Frame5.Visible = True
            
        Else
            Frame5.Visible = False
            Label33.Visible = False
            TxtDocRefCredi.Visible = False
            CmdBusDocRef.Visible = False
        End If
    End If
    
'    Frame5.Visible = False
    
    'si es Recibo por honorarios
    If NulosN(TxtTipDoc) = 2 Then
         FraRetencion.Enabled = True
         FraRetencion.Visible = True
         Fratipven.Enabled = False
         FraRetencion.Caption = "Retención de 4ta Categoria " & Trim(Str(TasaImpuesto)) + "%"
         txtisc.Enabled = False
         txtinafecto.Enabled = False
    Else
         Fratipven.Enabled = True
         FraRetencion.Enabled = False
         FraRetencion.Visible = False
         txtisc.Enabled = True
         txtinafecto.Enabled = True
    End If
    
'    If NulosN(TxtTipDoc) = 7 Then
'        Frame5.Left = 5175
'        Frame5.Top = 2565
'        TxtDocRef.Text = ""
'        Frame5.Visible = True
'    End If
    
    'buscamos para hallar el numero de serie asignado al almacen
    If TxtTipDoc.Text <> "" Then
        Dim Rst As New ADODB.Recordset
        Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(LblIdAlmacen.Caption) & "", xCon)
        If Rst.RecordCount <> 0 Then
            TxtNumSer.Text = Rst("numser")
            TxtNumSer_Validate True
        End If
        Set Rst = Nothing
    Else
        TxtNumSer.Text = ""
        TxtNumDoc.Text = ""
    End If
    
    Set xRs = Nothing
End Sub

Private Sub TxtTipItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipItem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipItem_Click
    End If
End Sub

Sub Filtrar()
    TabOne1.CurrTab = 0
    Dim xform As New eps_librerias.FormFiltrar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(7, 4) As String
    
    xCampos(0, 0) = "Cliente":         xCampos(0, 1) = "nombre":        xCampos(0, 2) = "C":         xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Fch. Emision":    xCampos(1, 1) = "fchdoc":        xCampos(1, 2) = "F":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Nº Documento":    xCampos(2, 1) = "numerodoc":     xCampos(2, 2) = "C":         xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Tipo Documento":  xCampos(3, 1) = "abrev":         xCampos(3, 2) = "C":         xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Forma de Pago":   xCampos(4, 1) = "desccond":      xCampos(4, 2) = "C":         xCampos(4, 3) = "1500"
    xCampos(5, 0) = "Moneda":          xCampos(5, 1) = "simbolo":       xCampos(5, 2) = "C":         xCampos(5, 3) = "1500"
    xCampos(6, 0) = "Estado":          xCampos(6, 1) = "estadoventa":   xCampos(6, 2) = "C":         xCampos(6, 3) = "1500"
    xCampos(7, 0) = "importe":         xCampos(7, 1) = "imptotdoc":     xCampos(7, 2) = "N":         xCampos(7, 3) = "1500"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstVent
    Set RstVent = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstVent
    Dg1.Refresh
End Sub

Private Sub TxtTipItem_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    If NulosC(TxtTipItem.Text) <> "" Then
        Set RstTmp = BuscaConCriterio("SELECT * FROM mae_tipoproducto WHERE id = " & NulosN(TxtTipItem.Text) & "", xCon)
        If RstTmp.RecordCount <> 0 Then
           LblTipoItem.Caption = RstTmp("descripcion")
        Else
            TxtTipItem.Text = ""
            LblTipoItem.Caption = ""
        End If
    End If
    Set RstTmp = Nothing
    
    pGridConfigurar
    
End Sub


Sub Imprimir()
    Dim RsPDoc As New ADODB.Recordset
    Dim RsPCab As New ADODB.Recordset
    Dim RsPDet As New ADODB.Recordset
    Dim xRsDoc As New ADODB.Recordset
    Dim xRsDet As New ADODB.Recordset
    Dim RstGui As New ADODB.Recordset
    Dim A As Integer
    Dim xCadGuias As String
    
    RST_Busq xRsDoc, "SELECT vta_ventas.*, mae_cliente.nombre, mae_cliente.dir, mae_cliente.numruc, mae_moneda.descripcion AS mon, " _
        & " [vta_conceptonc]![descripcion] & ' : REF A => ' & [mae_documento]![abrev] & ' ' & [vta_ventas_1]![numser] & '-' & [vta_ventas_1]![numdoc] AS docref2 " _
        & " FROM (((mae_moneda RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_moneda.id = vta_ventas.idmon) " _
        & " LEFT JOIN vta_conceptonc ON vta_ventas.idmotnotcre = vta_conceptonc.id) LEFT JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id) " _
        & " LEFT JOIN mae_documento ON vta_ventas_1.tipdoc = mae_documento.id Where (((vta_ventas.id) = " & RstVent("id") & ")) ORDER BY vta_ventas.fchdoc", xCon

    
    '"SELECT vta_ventas.*, mae_cliente.nombre, mae_cliente.dir, mae_cliente.numruc, mae_moneda.descripcion AS mon " _
        & " FROM mae_moneda RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_moneda.id = vta_ventas.idmon " _
        & " WHERE (((vta_ventas.id)=" & RstVent("id") & ")) ORDER BY vta_ventas.fchdoc", xCon
    
    RST_Busq xRsDet, "SELECT vta_ventasdet.*, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuentaven, " _
        & " alm_inventario.idtipven FROM mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN vta_ventasdet " _
        & " ON alm_inventario.id = vta_ventasdet.iditem) ON mae_unidades.id = alm_inventario.idunimed " _
        & " WHERE (((vta_ventasdet.idvta)=" & RstVent("id") & "))", xCon

    RST_Busq RsPDoc, "SELECT * FROM var_plantilladoc WHERE tipdoc = " & xRsDoc("tipdoc") & " ", xCon


    RST_Busq RstGui, "SELECT vta_guia.numdoc From vta_guia WHERE (((vta_guia.iddocven)=" & RstVent("id") & "))" _
        & " ORDER BY [vta_guia]![numser]+'-'+[vta_guia]![numdoc]", xCon

    If RstGui.RecordCount <> 0 Then
        RstGui.MoveFirst
        xCadGuias = ""
        For A = 1 To RstGui.RecordCount
            xCadGuias = xCadGuias + Trim(Str(NulosN(RstGui("numdoc"))))
            RstGui.MoveNext
            If RstGui.EOF = True Then
                Exit For
            End If
            xCadGuias = xCadGuias + ", "
        Next A
    End If
    Set RstGui = Nothing
    
    If RsPDoc.RecordCount = 0 Then
        MsgBox "No se ha definido la plantilla de impresion para este tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set xRsDoc = Nothing
        Set xRsDet = Nothing
        Set RsPDoc = Nothing
        Exit Sub
    End If
    RST_Busq RsPCab, "SELECT * FROM var_plantilladoc WHERE tipdoc = " & RsPDoc("tipdoc") & " ", xCon
    If RsPCab.RecordCount <> 0 Then
        A = RsPCab("id")
        RST_Busq RsPCab, "SELECT * FROM var_plantillacab WHERE idplan = " & A & " ORDER BY item", xCon
        RST_Busq RsPDet, "SELECT * FROM var_plantilladet WHERE idplan = " & A & " ORDER BY item", xCon
    End If
   
    Printer.Font = "Super Draft 15cpi"
    Printer.FontBold = True
    Printer.FontSize = 11
    Printer.ScaleMode = 6
    
    Dim xCam, xFor As String

    'imprime cabezera
    Do While RsPCab.EOF = False
        xCam = RsPCab("campo")
        xFor = NulosC(RsPCab("formato"))
        
        Printer.CurrentX = RsPCab("posx")
        Printer.CurrentY = RsPCab("posy")
        
        If NulosC(UCase(xCam)) <> UCase("x-numeletra") And NulosC(UCase(xCam)) <> UCase("x-numguia") And NulosC(UCase(xCam)) <> UCase("x-docref") Then
            Printer.Print Format((NulosC(xRsDoc(xCam))), xFor)
        Else
            If NulosC(UCase(xCam)) = UCase("x-numeletra") Then
                Printer.Print "Son : "; NumeroLetra(xRsDoc("imptotdoc"), xRsDoc("idmon"))
            End If
            If NulosC(UCase(xCam)) = UCase("x-numguia") Then
                Printer.Print xCadGuias
            End If
            If NulosC(UCase(xCam)) = UCase("x-docref") Then
                Printer.Print "Referente a Factura(s) : "; xRsDoc("docref")
            End If
        End If
        
        RsPCab.MoveNext
    Loop

    'imprime detalle
    Dim Fila As Integer
    
    Fila = RsPDet("posy")
    xRsDet.MoveFirst
    Do While xRsDet.EOF = False
        RsPDet.MoveFirst
        Do While RsPDet.EOF = False
            xCam = RsPDet("campo")
            xFor = NulosC(RsPDet("formato"))
            Printer.CurrentX = RsPDet("posx")
            Printer.CurrentY = Fila
            If xFor = "" Then
                Printer.Print NulosC(xRsDet(xCam))
            Else
                Printer.Print Format((NulosC(xRsDet(xCam))), xFor)
            End If
            RsPDet.MoveNext
        Loop
        Fila = Fila + 4
        
        xRsDet.MoveNext
    Loop
    
    Printer.EndDoc
End Sub

Private Sub CmdSel_Click()
    If NulosC(TxtTipItem.Text) = "" Then
        MsgBox "No ha especificado el tipo de venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipItem.SetFocus
        Exit Sub
    End If
    
    Dim xfrm As New eps_librerias.FormSeleccion
    Dim xCampos(3, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "4000":         xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
    xCampos(1, 0) = "Uni. Med":       xCampos(1, 1) = "abrev":          xCampos(1, 2) = "1000":        xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "Codigo":         xCampos(2, 1) = "codpro":        xCampos(2, 2) = "1200":         xCampos(2, 3) = "C":    xCampos(2, 4) = "S"


    '*******************************************************************************************
    Dim nSQLId As String
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 8, "alm_inventario.id", " NOT IN ", True)
    If nSQLId <> "" Then nSQLId = " AND " & nSQLId
    '*******************************************************************************************

    xfrm.SQLCad = "SELECT alm_inventario.*, mae_unidades.descripcion AS descuni, mae_unidades.abrev, mae_percepcion.tasa " _
        & " FROM mae_unidades RIGHT JOIN (mae_percepcion RIGHT JOIN alm_inventario ON mae_percepcion.id = alm_inventario.idper) " _
        & " ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = " & NulosN(TxtTipItem) & " )) " & nSQLId & " ORDER BY alm_inventario.descripcion"
    
    xfrm.Titulo = "Buscando Productos"
    
    Set xfrm.Coneccion = xCon
    Set xRs = xfrm.Seleccionar(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Dim X As Integer
            Dim Agregar As Boolean
            Dim xPrecio As Double
            Agregar = True
            Mostrando = True
            xRs.MoveFirst
            
            For X = 1 To xRs.RecordCount
                For A = 1 To Fg1.Rows - 1
                    If Fg1.TextMatrix(A, 6) = xRs("id") Then
                        Agregar = False
                    End If
                Next A
                If Agregar = True Then
                    Fg1.Rows = Fg1.Rows + 1
                    xPrecio = 0
                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = xRs("descripcion")
                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRs("abrev")
                    
                    If NulosN(LblIdCliente.Caption) <> 0 Then
                        xPrecio = UltimoPrecio(NulosN(xRs("id")), NulosN(LblIdCliente.Caption))
                    Else
                        xPrecio = UltimoPrecio(NulosN(xRs("id")), 0)
                    End If
                    Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(xPrecio, "0.0000")
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = xRs("id")
                    Fg1.TextMatrix(Fg1.Rows - 1, 9) = xRs("idunimed")
                    Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(xRs("idcuentaven"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(xRs("idtipven"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(xRs("tasa"))
                    Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(xRs("stckact"))
                                        
                End If
                xRs.MoveNext
                If xRs.EOF = True Then
                    Exit For
                End If
                Agregar = True
            Next X
            Mostrando = False
        End If
    End If
    Set xfrm = Nothing
End Sub

Sub ModificarSaldo()
    ActivarEntorno
    Frame8.Top = 2580
    Frame8.Left = 2760
    Frame8.Visible = True
    
    TxtNumDoc2.Text = ""
    TxtCliente2.Text = ""
    TxtSaldo2.Text = ""
    TxtNewSaldo2.Text = ""
    
    TxtNumDoc.Text = RstVent("numerodoc")
    TxtCliente2.Text = RstVent("nombre")
    TxtSaldo2.Text = RstVent("impsal")
End Sub

'*******************************

Private Sub CambiarMes()
    TabOne1.CurrTab = 0
    mMesActivo = SeleccionaMes(xCon)
    pCargarGrid
End Sub

Private Sub pCargarGrid()
    Dim nSQL  As String
    Dim Rpta As Integer
    Dim DiaIniAño  As String
    Dim xFechaRegistro As String
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    LblPeriodo2.Caption = LblMes.Caption
    DiaIniAño = "01/01/" + Trim(AnoTra)
    xFechaRegistro = "01/" + Format(mMesActivo, "00") + "/" + Trim(AnoTra)
    
    If mMesActivo = 0 Then
        nSQL = "SELECT vta_ventas.*, mae_cliente.nombre, IIf(IsNull([vta_ventas]![numser])=-1,[vta_ventas]![numdoc],[vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]) AS numerodoc, IIf(vta_ventas.Anulado=0,'Facturado','Anulado') AS EstadoVenta, " _
            & " mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_moneda.descripcion AS descmon, mae_moneda.simbolo, " _
            & " mae_impuestos.idcuenvta, con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom, mae_condpago.abrev AS conpagabre, Mid([vta_ventas].[numreg],1,2)+[mae_libros].[codsun]+Mid([vta_ventas].[numreg],3,4) AS numreg1, " _
            & " vta_ventas.fchdoc & '' as fchdoc1,vta_ventas.fchven & '' as fchven1,vta_ventas.impbru & '' as impbru1, vta_ventas.impigv & '' as impigv1 ,vta_ventas.imptotdoc & '' as imptotdoc1, vta_ventas.impsal & '' as impsal1, " _
            & " IIF(vta_ventas.anulado=-1,0,IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc])) & '' AS impven1 " _
            & " FROM ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) RIGHT JOIN (mae_condpago RIGHT " _
            & " JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_condpago.id = vta_ventas.idconpag) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) " _
            & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_tipoproducto ON vta_ventas.idtipo = mae_tipoproducto.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            & " WHERE (((vta_ventas.fchdoc)<CDate('" & DiaIniAño & "'))) ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] DESC"
            
    ElseIf mMesActivo < 13 Then
        nSQL = "SELECT vta_ventas.*, IIf(vta_ventas.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, IIf(IsNull([vta_ventas]![numser])=-1,[vta_ventas]![numdoc],[vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]) AS numerodoc, " _
            & " IIf(vta_ventas.Anulado=0,'Facturado','Anulado') AS EstadoVenta, mae_documento.descripcion AS nomdoc, IIf(vta_ventas.anulado=-1,'', mae_condpago.descripcion) AS desccond, " _
            & " mae_documento.abrev, mae_cliente.numruc, mae_moneda.descripcion AS descmon, IIf(vta_ventas.anulado=-1,'',mae_moneda.simbolo) AS simbolo, mae_impuestos.idcuenvta, " _
            & " con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom, IIf(vta_ventas.anulado=-1,'',mae_condpago.abrev) AS conpagabre, " _
            & " Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4) AS numreg1,  " _
            & " vta_ventas.fchdoc & '' as fchdoc1,vta_ventas.fchven & '' as fchven1,vta_ventas.impbru & '' as impbru1, vta_ventas.impigv & '' as impigv1 ,vta_ventas.imptotdoc & '' as imptotdoc1, vta_ventas.impsal & '' as impsal1, " _
            & " IIF(vta_ventas.anulado=-1,0,IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc])) & '' AS impven1 " _
            & " FROM ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento LEFT JOIN mae_impuestos " _
            & " ON mae_documento.idimp = mae_impuestos.id) RIGHT JOIN (mae_condpago RIGHT JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) " _
            & " ON mae_condpago.id = vta_ventas.idconpag) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) " _
            & " LEFT JOIN mae_tipoproducto ON vta_ventas.idtipo = mae_tipoproducto.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            & " WHERE (((vta_ventas.fchreg)=CDate('" & xFechaRegistro & "')) AND ((vta_ventas.fchdoc)>=CDate('" & DiaIniAño & "'))) ORDER BY vta_ventas!numser+'-'+vta_ventas!numdoc DESC"
    Else
        MsgBox "Ha selecionado el mes de Cierre, selecciones meses comprendidos entre Enero y Diciembre", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstVent = Nothing
        Set Dg1.DataSource = Nothing
        Dg1.Refresh
        Exit Sub
    End If
    
    TDB_FiltroLimpiar Dg1
    Set RstVent = Nothing
    
    '--cargando datos
    Me.MousePointer = vbHourglass
    RST_Busq RstVent, nSQL, xCon

    Set Dg1.DataSource = RstVent
    
    Me.MousePointer = vbDefault
    
    OpcionesPeriodo
    TabOne1.CurrTab = 0
    '************************************************
    
    If RstVent.RecordCount = 0 Then
        Rpta = MsgBox("No se ha registrado ninguna venta, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
        If Rpta = vbYes Then
            Nuevo
        Else
'            If MsgBox("Desea consultar Otro periodo", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbYes Then
'                CambiarMes
'                Exit Sub
'            Else
'                Unload Me
'            End If
'
        End If
'    Else
'        OpcionesPeriodo
    End If

    
    '************************************************
    
    
    
End Sub

Sub Buscar()
    TabOne1.CurrTab = 0
    Dim xRs As New ADODB.Recordset
    
    Dim nSQL As String
    Dim xCampos(8, 4) As String
    
    xCampos(0, 0) = "NumReg":        xCampos(0, 1) = "registro":     xCampos(0, 2) = "820":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":      xCampos(1, 2) = "400":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "N°. Documento": xCampos(2, 1) = "numerodoc":  xCampos(2, 2) = "1400":  xCampos(2, 3) = "C"
    xCampos(3, 0) = "FchEmi":        xCampos(3, 1) = "fchdoc":     xCampos(3, 2) = "830":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "FchVenc":       xCampos(4, 1) = "fchven":     xCampos(4, 2) = "830":   xCampos(4, 3) = "C"
    xCampos(5, 0) = "Cliente":       xCampos(5, 1) = "nombre":     xCampos(5, 2) = "2600":  xCampos(5, 3) = "C"
    
    xCampos(6, 0) = "M":             xCampos(6, 1) = "simbolo":    xCampos(6, 2) = "450":    xCampos(6, 3) = "C"
    xCampos(7, 0) = "Importe":         xCampos(7, 1) = "imptotdoc":     xCampos(7, 2) = "850":    xCampos(7, 3) = "N"
    
    nSQL = "SELECT vta_ventas.id,Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4) AS registro, IIf(vta_ventas.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, mae_documento.abrev, IIf(vta_ventas.anulado=-1,'',mae_moneda.simbolo) AS simbolo, format(vta_ventas.fchdoc,'dd/mm/yy') as fchdoc, format(vta_ventas.fchven,'dd/mm/yy') as fchven, vta_ventas.imptotdoc " _
        + vbCr + " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN vta_ventas ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
        + vbCr + " WHERE (((vta_ventas.numreg) Like '" & Format(mMesActivo, "00") & "%')) " _
    

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Compras", "nombre", "nombre", Principio

    If xRs.State = 1 Then
        RstVent.MoveFirst
        RstVent.Find "id = " & xRs("id") & ""
    End If
    
    Set xRs = Nothing
End Sub




Sub ActualizaSaldoDoc(idDocumento As Integer, Tabla As Integer, ImporteRestar As Double)
    '1 = compras
    '2 = Ventas
    '3 = honorarios
    
    Dim Rst As New ADODB.Recordset
    Dim Total As Double
    
    If Tabla = 2 Then
        RST_Busq Rst, "SELECT Sum(tes_cajadestinodet.acuenta) AS total FROM tes_caja LEFT JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes " _
            & " GROUP BY tes_cajadestinodet.iddoc, tes_caja.tipmov HAVING (((tes_cajadestinodet.iddoc)=" & idDocumento & ") AND ((tes_caja.tipmov)=1))", xCon
            
        Total = BuscaImporteDocumento(idDocumento, 1)
    End If
    
    If Rst.RecordCount <> 0 Then
        Total = ((Total - Rst("total")) - ImporteRestar)
    Else
        Total = (Total - ImporteRestar)
    End If
    
    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & Total & " WHERE (((vta_ventas.id)=" & idDocumento & "))"
    Set Rst = Nothing
End Sub


Function BuscaImporteDocumento(idDocumento As Integer, Tabla As Integer) As Double
    '1 = compras
    '2 = Ventas
    '3 = honorarios
    Dim Rst As New ADODB.Recordset
    
    'compras
    If Tabla = 1 Then RST_Busq Rst, "SELECT * FROM vta_ventas WHERE id = " & idDocumento & "", xCon
    
    If Rst.RecordCount <> 0 Then
        BuscaImporteDocumento = Rst("imptotdoc")
    Else
        BuscaImporteDocumento = 0
    End If
    
    Set Rst = Nothing
End Function

Private Sub pGridConfigurar()
    If NulosN(TxtTipItem.Text) = 5 Then
        Fg1.ColWidth(2) = 0
        Fg1.ColWidth(3) = 0
        Fg1.ColWidth(4) = 1100
        Fg1.ColWidth(5) = 1100
        Fg1.ColWidth(7) = 1200
    Else
        Fg1.ColWidth(2) = 435
        Fg1.ColWidth(3) = 855
        Fg1.ColWidth(4) = 930
        Fg1.ColWidth(5) = 960
        Fg1.ColWidth(7) = 1020
    End If
End Sub
