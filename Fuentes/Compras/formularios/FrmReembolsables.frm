VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmReembolsables 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras - Ingreso de Reembolsables"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraReclasifica 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1935
      Left            =   11970
      TabIndex        =   93
      Top             =   120
      Visible         =   0   'False
      Width           =   7050
      Begin VB.CommandButton CmdClAceptar 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   2055
         TabIndex        =   110
         Top             =   1440
         Width           =   1320
      End
      Begin VB.CommandButton CmdClaCancelar 
         Caption         =   "&Salir"
         Height          =   390
         Left            =   3765
         TabIndex        =   109
         Top             =   1440
         Width           =   1320
      End
      Begin VB.TextBox txtTotal1 
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
         Left            =   900
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   106
         TabStop         =   0   'False
         Text            =   "txtTotal1"
         Top             =   390
         Width           =   1440
      End
      Begin VB.CommandButton CmdBusCtaHab 
         Height          =   240
         Left            =   2085
         Picture         =   "FrmReembolsables.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   1095
         Width           =   240
      End
      Begin VB.CommandButton CmdBusCtaDeb 
         Height          =   240
         Left            =   2085
         Picture         =   "FrmReembolsables.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   780
         Width           =   240
      End
      Begin VB.TextBox TxtCtaHab 
         Height          =   300
         Left            =   900
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   98
         Text            =   "TxtCtaHab"
         Top             =   1065
         Width           =   1455
      End
      Begin VB.TextBox TxtCtaDeb 
         Height          =   300
         Left            =   900
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   96
         Text            =   "TxtCtaDeb"
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   60
         TabIndex        =   105
         Top             =   450
         Width           =   360
      End
      Begin VB.Label LblNomCtaHab 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblNomCtaHab"
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
         Left            =   2400
         TabIndex        =   104
         Top             =   1065
         Width           =   4560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Haber"
         Height          =   195
         Index           =   15
         Left            =   90
         TabIndex        =   103
         Top             =   1095
         Width           =   765
      End
      Begin VB.Label LblNomCtaDeb 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblNomCtaDeb"
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
         Left            =   2400
         TabIndex        =   102
         Top             =   750
         Width           =   4560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Debe"
         Height          =   195
         Index           =   14
         Left            =   90
         TabIndex        =   101
         Top             =   780
         Width           =   720
      End
      Begin VB.Label LbIdCuentaDeb 
         AutoSize        =   -1  'True
         Caption         =   "LbIdCuentaDeb"
         Height          =   195
         Left            =   3990
         TabIndex        =   100
         Top             =   510
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label LbIdCuentaHab 
         AutoSize        =   -1  'True
         Caption         =   "LbIdCuentaHab"
         Height          =   195
         Left            =   5310
         TabIndex        =   99
         Top             =   510
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reclasificar Cuenta"
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
         Left            =   120
         TabIndex        =   94
         Top             =   90
         Width           =   1680
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   3
         X1              =   -180
         X2              =   8415
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   3615
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   2
         X1              =   7020
         X2              =   7020
         Y1              =   15
         Y2              =   3645
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   8595
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   300
         Left            =   45
         Top             =   45
         Width           =   8520
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8265
      Top             =   30
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
            Picture         =   "FrmReembolsables.frx":0264
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":07A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":0B3A
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":0CBE
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":1112
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":122A
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":176E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":1CB2
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":1DC6
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":1EDA
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":232E
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":249A
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReembolsables.frx":29E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   3645
      Left            =   12120
      TabIndex        =   26
      Top             =   2310
      Visible         =   0   'False
      Width           =   8610
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   4425
         TabIndex        =   32
         Top             =   3120
         Width           =   1320
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   5790
         TabIndex        =   31
         Top             =   3120
         Width           =   1320
      End
      Begin VB.CommandButton CmdAddCenCos 
         Caption         =   "&Agregar C.C."
         Height          =   390
         Left            =   1500
         TabIndex        =   30
         Top             =   3120
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelCenCos 
         Caption         =   "&Eliminar C.C."
         Height          =   390
         Left            =   2865
         TabIndex        =   29
         Top             =   3120
         Width           =   1320
      End
      Begin VB.TextBox TxtTotPor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6330
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "TxtTotPor"
         Top             =   2670
         Width           =   975
      End
      Begin VB.TextBox TxtTotImp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7305
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "TxtTotImp"
         Top             =   2670
         Width           =   960
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg5 
         Height          =   2190
         Left            =   75
         TabIndex        =   33
         Top             =   465
         Width           =   8460
         _cx             =   14922
         _cy             =   3863
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
         FormatString    =   $"FrmReembolsables.frx":2D74
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
         ShowComboButton =   -1  'True
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
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   8595
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   8595
         X2              =   8595
         Y1              =   15
         Y2              =   3645
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   3615
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   8595
         Y1              =   3630
         Y2              =   3630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detallar Centro de Costos"
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
         Left            =   255
         TabIndex        =   34
         Top             =   90
         Width           =   2190
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   300
         Left            =   45
         Top             =   45
         Width           =   8520
      End
   End
   Begin VB.Frame Frame11 
      BorderStyle     =   0  'None
      Caption         =   "Frame11"
      Height          =   2700
      Left            =   12300
      TabIndex        =   22
      Top             =   390
      Visible         =   0   'False
      Width           =   7320
      Begin VB.CommandButton CmdAcep 
         Caption         =   "&Aceptar"
         Height          =   405
         Left            =   2985
         TabIndex        =   23
         Top             =   2220
         Width           =   1305
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg4 
         Height          =   1710
         Left            =   195
         TabIndex        =   24
         Top             =   465
         Width           =   6900
         _cx             =   12171
         _cy             =   3016
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
         Rows            =   50
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmReembolsables.frx":2E2F
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
         ShowComboButton =   -1  'True
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
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   7305
         Y1              =   2685
         Y2              =   2685
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   7290
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   15
         Y2              =   2670
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   7305
         X2              =   7305
         Y1              =   15
         Y2              =   2685
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   135
         Width           =   1860
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00400000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   330
         Left            =   45
         Top             =   60
         Width           =   7230
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   30
      TabIndex        =   35
      Top             =   360
      Width           =   11880
      _cx             =   20955
      _cy             =   12753
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   61
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6495
            Left            =   30
            TabIndex        =   111
            Top             =   300
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11456
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
            Columns(1).Caption=   "TD"
            Columns(1).DataField=   "tdabrev"
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
            Columns(4).Caption=   "Cliente"
            Columns(4).DataField=   "clinombre"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Proveedor"
            Columns(5).DataField=   "pronombre"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "M."
            Columns(6).DataField=   "moneda"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "T.C."
            Columns(7).DataField=   "impven1"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Imp. Bru."
            Columns(8).DataField=   "impbru1"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "I.G.V."
            Columns(9).DataField=   "impigv1"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Importe"
            Columns(10).DataField=   "imptot1"
            Columns(10).NumberFormat=   "0.00"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Saldo"
            Columns(11).DataField=   "impsal1"
            Columns(11).NumberFormat=   "0.00"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   12
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=12"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=794"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=714"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2646"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2566"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1455"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1376"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=5292"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=5212"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=661"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=582"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=953"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=873"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=1508"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1429"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(55)=   "Column(8).Visible=0"
            Splits(0)._ColumnProps(56)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(57)=   "Column(9).Width=1244"
            Splits(0)._ColumnProps(58)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._WidthInPix=1164"
            Splits(0)._ColumnProps(60)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(61)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(62)=   "Column(9).Visible=0"
            Splits(0)._ColumnProps(63)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(64)=   "Column(10).Width=1879"
            Splits(0)._ColumnProps(65)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(66)=   "Column(10)._WidthInPix=1799"
            Splits(0)._ColumnProps(67)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(68)=   "Column(10)._ColStyle=514"
            Splits(0)._ColumnProps(69)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(70)=   "Column(11).Width=1799"
            Splits(0)._ColumnProps(71)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(72)=   "Column(11)._WidthInPix=1720"
            Splits(0)._ColumnProps(73)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(74)=   "Column(11)._ColStyle=514"
            Splits(0)._ColumnProps(75)=   "Column(11).Order=12"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
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
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=82,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=79,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=80,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=81,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=78,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=74,.parent=13,.alignment=1"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
            _StyleDefs(84)  =   "Named:id=33:Normal"
            _StyleDefs(85)  =   ":id=33,.parent=0"
            _StyleDefs(86)  =   "Named:id=34:Heading"
            _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(88)  =   ":id=34,.wraptext=-1"
            _StyleDefs(89)  =   "Named:id=35:Footing"
            _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(91)  =   "Named:id=36:Selected"
            _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=37:Caption"
            _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(95)  =   "Named:id=38:HighlightRow"
            _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=39:EvenRow"
            _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(99)  =   "Named:id=40:OddRow"
            _StyleDefs(100) =   ":id=40,.parent=33"
            _StyleDefs(101) =   "Named:id=41:RecordSelector"
            _StyleDefs(102) =   ":id=41,.parent=34"
            _StyleDefs(103) =   "Named:id=42:FilterBar"
            _StyleDefs(104) =   ":id=42,.parent=33"
         End
         Begin VB.Label LblMes 
            AutoSize        =   -1  'True
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   8235
            TabIndex        =   64
            Top             =   30
            Width           =   765
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Reembolsables"
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
            TabIndex        =   63
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
            Caption         =   "LblPeriodo"
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
            Left            =   9810
            TabIndex        =   62
            Top             =   0
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   12525
         TabIndex        =   36
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdBusCli 
            Height          =   240
            Left            =   2985
            Picture         =   "FrmReembolsables.frx":2F0C
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   450
            Width           =   240
         End
         Begin VB.CheckBox ChkTC 
            Caption         =   "Check2"
            Enabled         =   0   'False
            Height          =   195
            Left            =   10470
            TabIndex        =   113
            Top             =   1410
            Width           =   195
         End
         Begin VB.TextBox TxtTC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   10680
            TabIndex        =   112
            Text            =   "TxtTC"
            Top             =   1350
            Width           =   1065
         End
         Begin VB.CommandButton CmdBusDocRef2 
            Height          =   240
            Left            =   7890
            Picture         =   "FrmReembolsables.frx":303E
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   2340
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDocRef 
            Height          =   240
            Left            =   1965
            Picture         =   "FrmReembolsables.frx":3170
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   2340
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   5205
            Picture         =   "FrmReembolsables.frx":32A2
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   1095
            Width           =   240
         End
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   2990
            Picture         =   "FrmReembolsables.frx":33D4
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   780
            Width           =   240
         End
         Begin VB.TextBox TxtGlosa 
            Height          =   300
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "TxtGlosa"
            Top             =   1995
            Width           =   10230
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   1965
            Picture         =   "FrmReembolsables.frx":3506
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1390
            Width           =   240
         End
         Begin VB.Frame Frame5 
            Height          =   495
            Left            =   9675
            TabIndex        =   37
            Top             =   -90
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
               TabIndex        =   38
               Top             =   150
               Width           =   1860
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2190
            Left            =   105
            TabIndex        =   10
            Top             =   3180
            Width           =   11610
            _cx             =   20479
            _cy             =   3863
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
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmReembolsables.frx":3638
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
            ShowComboButton =   -1  'True
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
            Height          =   300
            Left            =   1485
            TabIndex        =   1
            Top             =   1065
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
            Valor           =   "22/05/2008"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchPago 
            Height          =   300
            Left            =   10485
            TabIndex        =   6
            Top             =   1680
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
            Valor           =   "22/05/2008"
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2580
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   5
            Text            =   "TxtNumDoc"
            Top             =   1680
            Width           =   1440
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1485
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   4
            Text            =   "TxtNumSer"
            Top             =   1680
            Width           =   915
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   1485
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   2
            Text            =   "TxtIdMon"
            Top             =   1365
            Width           =   750
         End
         Begin VB.Frame Frame4 
            Height          =   1485
            Left            =   105
            TabIndex        =   40
            Top             =   5325
            Width           =   11610
            Begin VB.TextBox TxtRedondeo 
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
               Left            =   8880
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   15
               TabStop         =   0   'False
               Text            =   "TxtRedonde"
               Top             =   555
               Width           =   1230
            End
            Begin VB.TextBox TxtOtros 
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
               Left            =   7575
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   19
               TabStop         =   0   'False
               Text            =   "TxtOtros"
               Top             =   1110
               Width           =   1230
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Presupuesto"
               Enabled         =   0   'False
               Height          =   345
               Left            =   1545
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Presupuesto"
               Top             =   990
               Width           =   1395
            End
            Begin VB.CommandButton CmdPreHist 
               Caption         =   "Ver His. Precios"
               Enabled         =   0   'False
               Height          =   345
               Left            =   135
               Style           =   1  'Graphical
               TabIndex        =   46
               ToolTipText     =   "Historico de Precios"
               Top             =   990
               Width           =   1395
            End
            Begin VB.CheckBox ChkAjusta 
               Caption         =   "Ajustar "
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
               Left            =   10290
               TabIndex        =   44
               Top             =   615
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox TxtIGV3 
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
               Left            =   6255
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   18
               TabStop         =   0   'False
               Text            =   "TxtIGV3"
               Top             =   1110
               Width           =   1230
            End
            Begin VB.TextBox TxtIGV2 
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
               Left            =   4650
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   17
               TabStop         =   0   'False
               Text            =   "TxtIGV2"
               Top             =   1110
               Width           =   1230
            End
            Begin VB.TextBox TxtBruto3 
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
               Left            =   6255
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   13
               TabStop         =   0   'False
               Text            =   "TxtBruto3"
               Top             =   555
               Width           =   1230
            End
            Begin VB.TextBox TxtBruto2 
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
               Left            =   4650
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   12
               TabStop         =   0   'False
               Text            =   "TxtBruto2"
               Top             =   555
               Width           =   1230
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
               Left            =   10215
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   21
               TabStop         =   0   'False
               Text            =   "TxtTotal"
               Top             =   1110
               Width           =   1320
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
               Left            =   3285
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   16
               TabStop         =   0   'False
               Text            =   "TxtIGV"
               Top             =   1110
               Width           =   1230
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
               Left            =   3285
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   11
               TabStop         =   0   'False
               Text            =   "TxtBruto"
               Top             =   555
               Width           =   1230
            End
            Begin VB.TextBox TxtInafecto 
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
               Left            =   7575
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   14
               TabStop         =   0   'False
               Text            =   "TxtInafect"
               Top             =   555
               Width           =   1230
            End
            Begin VB.TextBox TxtISC 
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
               Left            =   8895
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   20
               TabStop         =   0   'False
               Text            =   "TxtISC"
               Top             =   1110
               Width           =   1230
            End
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "Eliminar Item"
               Enabled         =   0   'False
               Height          =   345
               Left            =   1545
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Eliminar Item"
               Top             =   270
               Width           =   1395
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "Agregar Item"
               Enabled         =   0   'False
               Height          =   345
               Left            =   135
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "Agregar Item"
               Top             =   270
               Width           =   1395
            End
            Begin VB.CommandButton CmdDetCenCos 
               Caption         =   "Centro de Costo"
               Enabled         =   0   'False
               Height          =   345
               Left            =   1545
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   "Centro de Costos"
               Top             =   630
               Width           =   1395
            End
            Begin VB.CommandButton CmdSeleccionar 
               Caption         =   "Seleccionar Item"
               Enabled         =   0   'False
               Height          =   345
               Left            =   135
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Seleccionar Items "
               Top             =   630
               Width           =   1395
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Redondeo a"
               Height          =   195
               Index           =   11
               Left            =   8925
               TabIndex        =   108
               Top             =   180
               Width           =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Céntimos"
               Height          =   195
               Index           =   10
               Left            =   8925
               TabIndex        =   107
               Top             =   360
               Width           =   645
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Otros Cargos"
               Height          =   195
               Left            =   7620
               TabIndex        =   86
               Top             =   885
               Width           =   915
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Tasa I.G.V."
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
               Left            =   10140
               TabIndex        =   84
               Top             =   225
               Width           =   990
            End
            Begin VB.Label LblIgvTasa 
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
               Left            =   11130
               TabIndex        =   77
               Top             =   210
               Width           =   960
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   10215
               TabIndex        =   76
               Top             =   555
               Width           =   1320
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "No Gravadas"
               Height          =   195
               Index           =   9
               Left            =   7620
               TabIndex        =   75
               Top             =   360
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Credito Filscal"
               Height          =   195
               Index           =   8
               Left            =   6240
               TabIndex        =   74
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "de Exp. o no Grav"
               Height          =   195
               Index           =   7
               Left            =   4650
               TabIndex        =   73
               Top             =   360
               Width           =   1290
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Grav y Exp"
               Height          =   195
               Index           =   6
               Left            =   3270
               TabIndex        =   72
               Top             =   360
               Width           =   780
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. 3"
               Height          =   195
               Left            =   6240
               TabIndex        =   71
               Top             =   885
               Width           =   540
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. 2"
               Height          =   195
               Left            =   4650
               TabIndex        =   70
               Top             =   885
               Width           =   540
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "A.G. sin derecho"
               Height          =   195
               Index           =   5
               Left            =   6240
               TabIndex        =   69
               Top             =   180
               Width           =   1185
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base Imp. Ope. Grav."
               Height          =   195
               Index           =   4
               Left            =   4650
               TabIndex        =   68
               Top             =   180
               Width           =   1530
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total"
               Height          =   195
               Index           =   2
               Left            =   10215
               TabIndex        =   52
               Top             =   885
               Width           =   360
            End
            Begin VB.Label LblRotulo 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. "
               Height          =   195
               Left            =   3270
               TabIndex        =   51
               Top             =   885
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base. Imp. Ope."
               Height          =   195
               Index           =   0
               Left            =   3270
               TabIndex        =   50
               Top             =   180
               Width           =   1140
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Adquisiciones"
               Height          =   195
               Index           =   1
               Left            =   7620
               TabIndex        =   49
               Top             =   180
               Width           =   975
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000003&
               Index           =   0
               X1              =   3060
               X2              =   3060
               Y1              =   105
               Y2              =   1470
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000005&
               Index           =   1
               X1              =   3075
               X2              =   3075
               Y1              =   105
               Y2              =   1455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "I.S.C."
               Height          =   195
               Index           =   3
               Left            =   9015
               TabIndex        =   48
               Top             =   885
               Width           =   390
            End
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1485
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   0
            Text            =   "TxtNumRuc"
            Top             =   750
            Width           =   1770
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   4725
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   3
            Text            =   "TxtTipDoc"
            Top             =   1065
            Width           =   750
         End
         Begin VB.TextBox TxtDocRef2 
            Height          =   300
            Left            =   6135
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   9
            Text            =   "TxtDocRef2"
            Top             =   2310
            Width           =   2025
         End
         Begin VB.TextBox TxtTipDocRef 
            Height          =   300
            Left            =   1485
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   8
            Text            =   "Txt"
            Top             =   2310
            Width           =   750
         End
         Begin VB.TextBox TxtNumRucCli 
            Height          =   300
            Left            =   1485
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   116
            Text            =   "TxtNumRucCl"
            Top             =   420
            Width           =   1770
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
            Left            =   3255
            TabIndex        =   118
            Top             =   420
            Width           =   5970
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   4
            Left            =   150
            TabIndex        =   117
            Top             =   450
            Width           =   480
         End
         Begin VB.Label LblIdcliente 
            AutoSize        =   -1  'True
            Caption         =   "LblIdcliente"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1680
            TabIndex        =   114
            Top             =   240
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label LblIdDocRef2 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDocRef2"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   7620
            TabIndex        =   92
            Top             =   2400
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Doc. Referencia"
            Height          =   195
            Index           =   13
            Left            =   4680
            TabIndex        =   90
            Top             =   2355
            Width           =   1395
         End
         Begin VB.Label LblTipDocref 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipDocref"
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
            Left            =   2235
            TabIndex        =   89
            Top             =   2310
            Width           =   2325
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip de Doc. Ref."
            Height          =   195
            Index           =   12
            Left            =   150
            TabIndex        =   88
            Top             =   2355
            Width           =   1185
         End
         Begin VB.Label LblIdTipPer 
            AutoSize        =   -1  'True
            Caption         =   "LblIdTipPer"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   0
            TabIndex        =   85
            Top             =   225
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Documento"
            Height          =   195
            Index           =   1
            Left            =   3255
            TabIndex        =   83
            Top             =   1125
            Width           =   1410
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
            Left            =   5490
            TabIndex        =   82
            Top             =   1065
            Width           =   3735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   80
            Top             =   780
            Width           =   735
         End
         Begin VB.Label LblNomPro 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomPro"
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
            Left            =   3255
            TabIndex        =   79
            Top             =   750
            Width           =   5970
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   10
            Left            =   150
            TabIndex        =   67
            Top             =   2025
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   66
            Top             =   1425
            Width           =   585
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   2640
            TabIndex        =   60
            Top             =   240
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2430
            Top             =   1800
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   59
            Top             =   1725
            Width           =   1275
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Reebolsable"
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
            Left            =   90
            TabIndex        =   58
            Top             =   30
            Width           =   11610
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
            Left            =   2250
            TabIndex        =   57
            Top             =   1365
            Width           =   2325
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "T.C."
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
            Left            =   10005
            TabIndex        =   56
            Top             =   1425
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Emisión"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   55
            Top             =   1110
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Pago"
            Height          =   195
            Index           =   8
            Left            =   9615
            TabIndex        =   54
            Top             =   1725
            Width           =   735
         End
         Begin VB.Label LblIdCenCos 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCenCos"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4050
            TabIndex        =   53
            Top             =   285
            Visible         =   0   'False
            Width           =   900
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   65
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
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
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
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu1 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu1_1_1 
         Caption         =   "Agregar Item"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Item"
      End
      Begin VB.Menu menu1_4 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_5 
         Caption         =   "Ver Historico de Precios"
      End
   End
   Begin VB.Menu Opciones 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu Opciones_1 
         Caption         =   "Agregar documentos de entrada"
      End
      Begin VB.Menu Opciones_2 
         Caption         =   "Agregar documentos de entrada registrados"
      End
   End
End
Attribute VB_Name = "FrmReembolsables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstComp As New ADODB.Recordset
Dim QueHace As Integer
Dim TasaImpuesto As Double
Dim CaracteresNumericos As String
Dim CaracteresNumericos2 As String
Dim SeEjecuto As Boolean
Dim ValTipCam As Double
Dim xDescImp As String
Dim xIdCuenTasa As Integer  'codigo de la cuenta contable del impuesto
Dim xCuentaDoc As Integer   'codigo de la cuenta contable del documento
Dim Mostrando As Boolean
Dim RstTmp As New ADODB.Recordset
Dim xFchFin, xFchIni, xFechaMes As String
Dim RstTempISC As New ADODB.Recordset
Dim AgePer As Boolean
Dim AgeRet As Boolean
Dim DetCenCos As Boolean    'especifica si se va a detallar el centro de costos
Dim CodSunatDoc As String   'especifica el codigo de la sunat del documento
Dim xPorIgv  As Double
Dim xHorIni As Date

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim mIdRegistro& '--identificador del registro
Dim Agregando As Boolean
Dim mMesActivo As Integer '--indica el mes activo
Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Sub Eliminar()
Exit Sub
    Dim Rpta As Integer
    
    TabOne1.CurrTab = 0
    
    If RstComp.State = 0 Then Exit Sub
    If RstComp.RecordCount = 0 Then
        MsgBox "No hay Registros de Compras para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    '**********************************************************************************************************************
    '---evaluar si el registro de compras esta vinculado con otros modulos
    Dim nSQL As String
    Dim Rst As New ADODB.Recordset
    Dim xId&
    
    
    xId = RstComp("id")
    '--verificar si esta vincuado con lgd
    nSQL = "SELECT Left([vta_gastodebito].[numreg],2) & [mae_libros].[codsun] & Right([vta_gastodebito].[numreg],4) AS registro, mae_documento.abrev, vta_gastodebito.numser, vta_gastodebito.numdoc, vta_gastodebito.fchemi AS fchdoc " _
    + vbCr + " FROM ((vta_gastodebito LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) INNER JOIN vta_gastodebitodet ON vta_gastodebito.id = vta_gastodebitodet.idlgd) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id " _
    + vbCr + " WHERE (((vta_gastodebitodet.idmod)=10) AND ((vta_gastodebitodet.iddoc)=" & xId & "));"


    
    RST_Busq Rst, nSQL, xCon
    
    If Rst.RecordCount <> 0 Then
        MsgBox "No puede elimiar el registro, esta vinculado con Liquidación Gasto Débito" & vbCr & "Num. Reg. " & Rst("registro") & vbCr & "Num. Doc. " & Rst("numdoc"), vbInformation, xTitulo
        Set Rst = Nothing
        Exit Sub
    End If
    Set Rst = Nothing
    '**********************************************************************************************************************
    
    Rpta = MsgBox("¿Esta seguro de eliminar la compra seleccionada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        
        xCon.Execute "DELETE * FROM com_reembolsables WHERE id = " & xId & ""
        
        
        MsgBox "El reembolsable se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstComp.Requery
        Dg1.Refresh
        
    End If
End Sub


Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Cancelar()
    Bloquea
    Fg1.ColComboList(1) = ""
    Label5.Caption = "Detalle de Compra"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDNone
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
End Sub

Sub Nuevo()
    QueHace = 1
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Label5.Caption = "Agregando Compra"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    Fg1.Rows = 1
    Fg5.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
     
     'CmdDetCenCos.Enabled = True
    If xOrigen = 1 Then
        CargarValoresDefecto
    End If
    
    
    xHorIni = Time

    pGridConfigurar
    TxtNumRuc.SetFocus
End Sub

Sub CargarValoresDefecto()
    TxtFchDoc.Valor = Date
'    TxtTipCom.Text = "1"
'    TxtTipCom_Validate True
    TxtIdMon.Text = 1
    TxtIdMon_Validate True
    TxtTipDoc.Text = "1"
    TxtTipDoc_Validate True
'    TxtConPag.Text = "1"
'    TxtConPag_Validate True
'    TxtIdAlmacen.Text = "1"
'    TxtIdAlmacen_Validate True
    
    'TxtFchVen.Valor = Date
    'OptOpera1.Value = True
    'OptOpera1_Click
End Sub

Sub Modificar()
    QueHace = 2
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    MuestraSegundoTab
    Label5.Caption = "Modificando Compra"
    Fg1.ColComboList(2) = "|..."
    Fg1.ColComboList(5) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    
    xHorIni = Time
    
'    Check1.Value = 0
    
    TxtFchDoc.SetFocus
End Sub

Sub MuestraSegundoTab()
    Blanquea
    If RstComp.EOF = True Or RstComp.BOF = True Or RstComp.RecordCount = 0 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    
    
    
    pGridConfigurar
    
    LblIdProveedor.Caption = NulosN(RstComp("idpro"))
    TxtNumRuc.Text = NulosC(RstComp("pronumruc"))
    LblNomPro.Caption = NulosC(RstComp("pronombre"))
    
    LblIdcliente.Caption = NulosN(RstComp("idcli"))
    TxtNumRucCli.Text = NulosC(RstComp("clinumruc"))
    LblNomCli.Caption = NulosC(RstComp("clinombre"))
    
    TxtTipDoc.Text = NulosN(RstComp("tipdoc"))
        
    TxtNumSer.Text = NulosC(RstComp("numser"))
    TxtNumDoc.Text = NulosC(RstComp("numdoc"))
    If IsDate(RstComp("fchdoc")) = True Then TxtFchDoc.Valor = RstComp("fchdoc")
    
    TxtIdMon.Text = NulosN(RstComp("idmon"))
    TxtGlosa.Text = NulosC(RstComp("glosa"))
    
    'mostramos el documento de referencia de la compra
    If NulosN(RstComp("idtipdocref")) = 0 Then
        TxtTipDocRef.Text = ""
    Else
        TxtTipDocRef.Text = NulosN(RstComp("idtipdocref"))
    End If
    TxtTipDocRef_Validate False
    LblIdDocRef2.Caption = NulosN(RstComp("iddocref2"))
    Dim Rst As New ADODB.Recordset
    If NulosN(TxtTipDocRef.Text) = 1 Then
        RST_Busq Rst, "SELECT com_ordencompra.id, [com_ordencompra]![numser] & '-' & [com_ordencompra]![numdoc] AS numdoc From com_ordencompra " _
            & " WHERE (((com_ordencompra.id)=" & NulosN(LblIdDocRef2.Caption) & "))", xCon
    ElseIf NulosN(TxtTipDocRef.Text) = 2 Then
    ElseIf NulosN(TxtTipDocRef.Text) = 3 Then
    ElseIf NulosN(TxtTipDocRef.Text) = 4 Then
        RST_Busq Rst, "SELECT var_ordendespacho.id, [var_ordendespacho]![anno] & [var_ordendespacho]![idaduana] & [var_ordendespacho]![idregimen] & [var_ordendespacho]![numdoc] AS numdoc" _
            & " From var_ordendespacho WHERE (((var_ordendespacho.id)=" & NulosN(LblIdDocRef2.Caption) & "))", xCon
    End If
    
    If Rst.State <> 0 Then
        If Rst.RecordCount <> 0 Then
            TxtDocRef2.Text = Rst("numdoc")
            LblIdDocRef2.Caption = Rst("id")
        Else
            TxtDocRef2.Text = ""
            LblIdDocRef2.Caption = ""
        End If
    End If
    Set Rst = Nothing
    
    
    
    LblMoneda.Caption = NulosC(RstComp("descmon"))
    
    
    Dim xCambioQueHace As Boolean
    xCambioQueHace = False
    If QueHace = 3 Then
        xCambioQueHace = True
        QueHace = 2
    End If
    TxtNumRuc_Validate True
    TxtTipDoc_Validate True
    
    If xCambioQueHace = True Then
        QueHace = 3
    End If
    
    '--tipo de cambio
    If NulosN(RstComp("tc")) = 0 Then
        ChkTC.Value = 0
        TxtTC.Text = NulosN(RstComp("impven1"))
        TxtTC.BackColor = &H8000000F
        TxtTC.Enabled = False
    Else
        ChkTC.Value = 1
        TxtTC.Text = NulosN(RstComp("tc"))
        TxtTC.BackColor = vbWhite
        TxtTC.Enabled = True
    End If
    If QueHace = 3 Then TxtTC.BackColor = &H8000000F
    
    '---------------------------------------------------

    '--------------------------------------
    TxtTipDoc_Validate True
    
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    
    Agregando = True
    Fg1.Rows = 1
    
''''    RST_Busq RstDet, "SELECT com_comprasdet.*, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuenta, " _
''''        & " con_planctas.ctadesdeb, con_planctas.ctadeshab,  alm_inventario.idnetonodomi,mae_tipocompra.abrev AS tcompra " _
''''        & " FROM (((con_planctas RIGHT JOIN alm_inventario ON con_planctas.id = alm_inventario.idcuenta) RIGHT JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem) LEFT JOIN mae_unidades ON com_comprasdet.idunimed = mae_unidades.id) LEFT JOIN mae_tipocompra ON com_comprasdet.idtipcom = mae_tipocompra.id " _
''''        & " WHERE (((com_comprasdet.idcom)=" & NulosN(RstComp("id")) & "))", xCon
''''
''''    If RstDet.State = 1 Then
''''        If RstDet.RecordCount <> 0 Then
''''            RstDet.MoveFirst
''''            For A = 1 To RstDet.RecordCount
''''                Fg1.Rows = Fg1.Rows + 1
''''
''''                Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstDet("codpro"))
''''
''''                Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDet("descripcion"))
''''                Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstDet("abrev"))
''''                Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosN(RstDet("canpro"))
''''
''''                Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosN(RstDet("idtipcom"))
''''                Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(RstDet("tcompra"))
''''
''''                Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosN(RstDet("preunibru"))
''''
''''
''''                Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(NulosN(RstDet("valdes")))
''''                Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(RstDet("preuni"))
''''                Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(RstDet("imptot"))
''''
''''                Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(RstDet("iditem"))
''''                Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(RstDet("idunimed"))
''''                Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(RstDet("idcuenta"))
''''                Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosN(RstDet("idtipcom"))
''''                Fg1.TextMatrix(Fg1.Rows - 1, 15) = NulosN(RstDet("ctadesdeb"))
''''                Fg1.TextMatrix(Fg1.Rows - 1, 16) = NulosN(RstDet("ctadeshab"))
''''                Fg1.TextMatrix(Fg1.Rows - 1, 19) = NulosN(RstDet("idnetonodomi"))
''''
''''                RstDet.MoveNext
''''                If RstDet.EOF = True Then
''''                    Exit For
''''                End If
''''            Next A
''''        End If
''''    End If
''''    BuscarImpuestos
    
'    AgregarCentroCosto2 True, RstComp("id")
    
    If NulosN(TxtTipDoc.Text) = 2 Then
        'recibo por honorarios
        TxtInafecto.Text = Format(NulosN(RstComp("impina")), FORMAT_MONTO)
        TxtBruto.Text = Format(NulosN(RstComp("impbru")), FORMAT_MONTO)
        TxtIGV.Text = Format(NulosN(RstComp("impigv")), FORMAT_MONTO)
        'TxtTotal.Text = Format(NulosN(RstComp("impbru")) - NulosN(RstComp("impigv")), FORMAT_MONTO)
        TxtTotal.Text = Format(NulosN(RstComp("imptot")), FORMAT_MONTO)
    Else
        TxtInafecto.Text = Format(NulosN(RstComp("impina")), FORMAT_MONTO)
        
        TxtBruto.Text = Format(NulosN(RstComp("impbru")), FORMAT_MONTO)
        TxtBruto2.Text = "0.00"
        TxtBruto3.Text = "0.00"
        
        TxtIGV.Text = Format(NulosN(RstComp("impigv")), FORMAT_MONTO)
        TxtIGV2.Text = "0.00"
        TxtIGV3.Text = "0.00"
        TxtOtros.Text = "0.00"
        TxtISC.Text = "0.00"
        
        TxtTotal.Text = Format(NulosN(RstComp("imptot")), FORMAT_MONTO)
        
        
    End If
    
    'mostramos el centro de costos
    Set RstDet = Nothing
'    RST_Busq RstDet, "SELECT com_comprascosto.*, con_centrocosto.codigo, con_centrocosto.descripcion " _
'        & " FROM con_centrocosto RIGHT JOIN com_comprascosto ON con_centrocosto.id = com_comprascosto.idcencos " _
'        & " WHERE (((com_comprascosto.idcom)=" & RstComp("id") & "))", xCon
'
'    If RstDet.RecordCount <> 0 Then
'        Fg5.Rows = 1
'        'si tiene mas de un centro de costos lo mostramos en otro formulario
'        CmdDetCenCos.Enabled = True
'        CmdDetCenCos.Caption = "Centro Costo"
'        DetCenCos = True
'        RstDet.MoveFirst
'        For A = 1 To RstDet.RecordCount
'            Fg5.Rows = Fg5.Rows + 1
'            Fg5.TextMatrix(A, 1) = RstDet("codigo")
'            Fg5.TextMatrix(A, 2) = RstDet("descripcion")
'            Fg5.TextMatrix(A, 4) = Format(RstDet("impcos"), "0.00")
'            Fg5.TextMatrix(A, 5) = Format(RstDet("idcencos"), "0.00")
'            RstDet.MoveNext
'
'            If RstDet.EOF = True Then Exit For
'        Next A
'    End If
    
    
    Set RstDet = Nothing
    Agregando = False
    
End Sub

Sub CargarIngresoAlmacen(IdCompra As Integer)
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq Rst, "SELECT alm_ingresodoc.iddoc, alm_ingreso.tipmov, alm_ingreso.fching, mae_documento.abrev, alm_ingreso.nombre, alm_ingreso.id, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc " _
        & " FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) RIGHT JOIN alm_ingresodoc ON alm_ingreso.id = alm_ingresodoc.id " _
        & " WHERE (((alm_ingresodoc.iddoc)=" & IdCompra & ") AND ((alm_ingreso.tipmov)=-1))", xCon

    Fg4.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(A, 1) = Rst("fching")
            Fg4.TextMatrix(A, 2) = Rst("abrev")
            Fg4.TextMatrix(A, 3) = Rst("numdoc")
            Fg4.TextMatrix(A, 4) = Rst("nombre")
            Fg4.TextMatrix(A, 5) = Rst("id")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

Sub Bloquea()
    'TxtTipCom.Locked = Not TxtTipCom.Locked
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
    'TxtFchVen.Locked = Not TxtFchVen.Locked
    'TxtFchPago.Locked = Not TxtFchPago.Locked
    'TxtConPag.Locked = Not TxtConPag.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    'TxtIdAlmacen.Locked = Not TxtIdAlmacen.Locked
    
    ChkTC.Enabled = Not ChkTC.Enabled
    
    TxtGlosa.Locked = Not TxtGlosa.Locked
    
    TxtTipDocRef.Locked = Not TxtTipDocRef.Locked
    'TxtDocRef.Locked = Not TxtDocRef.Locked
    
    TxtBruto.Locked = Not TxtBruto.Locked
    TxtBruto2.Locked = Not TxtBruto2.Locked
    TxtBruto3.Locked = Not TxtBruto3.Locked
    
    TxtRedondeo.Locked = Not TxtRedondeo.Locked
    TxtTotal.Locked = Not TxtTotal.Locked
    
    'Frame9.Enabled = Not Frame9.Enabled
    
''    CmdAddItem.Enabled = Not CmdAddItem.Enabled
''    CmdDelItem.Enabled = Not CmdDelItem.Enabled
    
    TxtTC.BackColor = &H8000000F

End Sub

Sub Blanquea()
    
    'TxtNumOrdCom.Text = ""
    
    TxtTipDoc.Text = ""
    TxtNumRuc.Text = ""
    TxtNumRucCli.Text = ""
    
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtFchDoc.Valor = ""
    
    
    
    TxtIdMon.Text = ""
    'TxtCodCenCos.Text = ""
    
    TxtGlosa.Text = ""
    
    'LblCentroCosto.Caption = ""
    LblIdCenCos.Caption = ""
    LblNomDoc.Caption = ""
    LblNomPro.Caption = ""
    LblMoneda.Caption = ""
    
    LblIdProveedor.Caption = ""
    LblIdcliente.Caption = ""
    LblNomPro.Caption = ""
    LblNomCli.Caption = ""
        
    TxtTipDocRef.Text = ""
    TxtDocRef2.Text = ""
    LblTipDocref.Caption = ""
    LblIdDocRef2.Caption = ""
    
    TxtBruto.Text = "0.00"
    TxtBruto2.Text = "0.00"
    TxtBruto3.Text = "0.00"
    TxtIGV.Text = "0.00"
    TxtIGV2.Text = "0.00"
    TxtIGV3.Text = "0.00"
    TxtTotal.Text = "0.00"
    TxtISC.Text = "0.00"
    TxtInafecto.Text = "0.00"
    TxtOtros.Text = "0.00"
    
    TxtRedondeo.Text = "0.00"
    
    '---
    txtTotal1.Text = "0.00"
    TxtCtaDeb.Text = ""
    LblNomCtaDeb.Caption = ""
    LbIdCuentaDeb.Caption = ""

    TxtCtaHab.Text = ""
    LblNomCtaHab.Caption = ""
    LbIdCuentaHab.Caption = ""
    '----
    
End Sub

Private Sub ChkAjusta_Click()
    If ChkAjusta.Value = 1 Then
        TxtBruto.Locked = False
        TxtInafecto.Locked = False
        TxtIGV.Locked = False
        TxtISC.Locked = False
        TxtTotal.Locked = False
    Else
        TxtBruto.Locked = True
        TxtInafecto.Locked = True
        TxtIGV.Locked = True
        TxtISC.Locked = True
        TxtTotal.Locked = True
    End If
End Sub



Private Sub CmdAcep_Click()
    TabOne1.Enabled = True
    Toolbar1.Enabled = True
    
    Frame11.Visible = False
End Sub

Private Sub CmdAceptar_Click()
    If QueHace = 3 Then
        ActivarEntorno
        Frame6.Visible = False
        Exit Sub
    End If
    
    Dim xTot As Double
    If NulosN(TxtInafecto.Text) >= 0 Then
        xTot = NulosN(TxtBruto.Text) + NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text) + NulosN(TxtInafecto.Text)
    Else
        If NulosN(TxtInafecto.Text) < 0 Then
            xTot = NulosN(TxtBruto.Text) + NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text) + Val(TxtInafecto.Text)
        Else
            xTot = Val(TxtBruto.Text) + NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text)
        End If
    End If
    
    If NulosN(Format(xTot, "0.00")) <> NulosN(Format(TxtTotImp.Text, "0.00")) Then
        MsgBox "la distribucion del centro de costo no coincide con el importe del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    'TxtCodCenCos.Text = ""
    'LblCentroCosto.Caption = ""
    LblIdCenCos.Caption = ""
    
    DetCenCos = True
    Frame6.Visible = False
    ActivarEntorno
End Sub

Private Sub CmdBusDocRef2_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Nº Documento":      xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Emi.":         xCampos(1, 1) = "fchemi":      xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Fch. Ven.":         xCampos(2, 1) = "fchven":      xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Proveedor":         xCampos(3, 1) = "nombre":      xCampos(3, 2) = "4000":         xCampos(3, 3) = "C"
    
    If NulosN(TxtTipDocRef.Text) = 1 Then
        'Orden de Compra
        xform.SQLCad = "SELECT com_ordencompra.id, [com_ordencompra]![numser] & '-' & [com_ordencompra]![numdoc] AS numdoc, com_ordencompra.fchemi, " _
            & " com_ordencompra.fchven, mae_prov.nombre FROM mae_prov RIGHT JOIN com_ordencompra ON mae_prov.id = com_ordencompra.idpro " _
            & " Where (((com_ordencompra.idest) = 2)) ORDER BY com_ordencompra.fchemi"
        xform.Titulo = "Orden de Compra"
        
    ElseIf NulosN(TxtTipDocRef.Text) = 2 Then
        'Orden de Produccion
        MsgBox "Opcion no disponible"
        xform.Titulo = "Orden de Produccion"
        Exit Sub
    ElseIf NulosN(TxtTipDocRef.Text) = 3 Then
        'Orden de Mantenimiento
        xform.Titulo = "Orden de Matenimiento"
        MsgBox "Opcion no disponible"
        Exit Sub
    ElseIf NulosN(TxtTipDocRef.Text) = 4 Then
        'Orden de Despacho
        xform.SQLCad = "SELECT var_ordendespacho.id, var_ordendespacho!anno & var_ordendespacho!idaduana & var_ordendespacho!idregimen & var_ordendespacho!numdoc AS numdoc, " _
            & " mae_cliente.nombre, var_ordendespacho.idcli, var_ordendespacho.fchemi, var_ordendespacho.fchven FROM var_ordendespacho LEFT JOIN mae_cliente " _
            & " ON var_ordendespacho.idcli = mae_cliente.id"
        
        xform.Titulo = "Orden de Despacho"
    Else
        Set xform = Nothing
        Set xRs = Nothing
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
            TxtDocRef2.Text = xRs("numdoc")
            LblIdDocRef2.Caption = xRs("id")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusProv_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Proveedor":    xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id, mae_prov.idcondpag, mae_prov.tipper, mae_tipoempresa.descripcion" _
        & " FROM mae_tipoempresa RIGHT JOIN mae_prov ON mae_tipoempresa.id = mae_prov.tipper WHERE (((mae_prov.activo)=-1))"

    'SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id, mae_prov.idcondpag From mae_prov WHERE activo = -1 "
    
    xform.Titulo = "Buscando Proveedor"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumRuc.Text = xRs("numruc")
            LblNomPro.Caption = xRs("nombre")
            LblIdProveedor.Caption = xRs("id")
            
            
            LblIdTipPer.Caption = xRs("tipper")
            
            
            TxtFchDoc.SetFocus
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
    
    xCampos(0, 0) = "Descripción":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Código":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
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
            'Fg1.SetFocus
            TxtTipDoc.SetFocus
            
'            If Trim(TxtIdMon.Text) = "1" Then
'                LblTipCam.Visible = False
'                LblTipoCambio.Visible = False
'            Else
'                If TxtFchDoc.Valor = "" And ChkTC.Value = 0 Then
'                    MsgBox "No ha especificado la fecha del documento, no se puede determinar " & Chr(13) _
'                        & "la fecha del tipo de cambio para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'
'                    TxtIdMon.Text = ""
'                    TxtFchDoc.SetFocus
'                    Exit Sub
'                End If
'                LblTipCam.Visible = True
'                LblTipoCambio.Visible = True
                'Set xRs = Nothing
                'Set xRs = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = CDATE('" & TxtFchDoc.Valor & "')", xCon)
                'If xRs.RecordCount = 1 Then
                    'LblTipoCambio.Caption = Format(xRs("impcom"), "0.000")
                'End If
'                LblTipoCambio.Caption = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
                
'                If ChkTC.Value = 0 Then TxtTC.Text = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
                
'            End If
            xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuen  as cuentaimp" _
        & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id"
    
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
            CodSunatDoc = xRs("codsun")
            TxtTipDoc.Text = xRs("id")
            LblNomDoc.Caption = NulosC(xRs("descripcion"))
            TasaImpuesto = NulosN(xRs("tasa"))
            xDescImp = NulosC(xRs("descripcion"))
            xIdCuenTasa = NulosN(xRs("cuentaimp"))
            LblRotulo = Trim(NulosC(xRs("abreimp"))) + " (       )"
            LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) '+ "%"
            xPorIgv = (TasaImpuesto / 100)
            
            'xCuentaDoc = NulosN(xRs("idcuen"))
            'TxtNumRuc.SetFocus
            'TxtIdAlmacen.SetFocus
            
            
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDocRef_Click()
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
            TxtTipDocRef.Text = xRs("id")
            LblTipDocref.Caption = xRs("descripcion")
            TxtDocRef2.Text = ""
            LblIdDocRef2.Caption = ""
            TxtDocRef2.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCancelar_Click()
    ActivarEntorno
    DetCenCos = False
    Frame6.Visible = False
End Sub

Sub CargarItems()
    Dim A As Integer
    Dim xCadWHERE As String
    Dim Rst As New ADODB.Recordset
    
    For A = 1 To Fg4.Rows - 1
        xCadWHERE = xCadWHERE + "(alm_ingresodet.id = " & Val(Fg4.TextMatrix(A, 5)) & ")"
        If A = Fg4.Rows - 1 Then
            Exit For
        End If
        xCadWHERE = xCadWHERE + " OR "
    Next A
    
    xCadWHERE = "(" + xCadWHERE + ")"
    
    RST_Busq Rst, "SELECT alm_inventario.codpro, mae_unidades.abrev, alm_inventario.descripcion, Sum(alm_ingresodet.cantidad) AS cantidad, " _
        & " con_planctas.ctadesdeb, con_planctas.ctadeshab, alm_inventario.idcuenta, alm_inventario.iddet, alm_inventario.idtipcom, alm_inventario.id, " _
        & " alm_inventario.idunimed, mae_tipocompra.abrev AS tcompra " _
        & " FROM mae_tipocompra RIGHT JOIN (con_planctas RIGHT JOIN ((alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id)  " _
        & " LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON con_planctas.id = alm_inventario.idcuenta) ON mae_tipocompra.id = alm_inventario.idtipcom " _
        & " Where " + xCadWHERE _
        & " GROUP BY alm_inventario.codpro, mae_unidades.abrev, alm_inventario.descripcion, con_planctas.ctadesdeb, con_planctas.ctadeshab, " _
        & " alm_inventario.idcuenta, alm_inventario.iddet, alm_inventario.idtipcom, alm_inventario.id, alm_inventario.idunimed", xCon
    
    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Mostrando = True

        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(Rst("codpro"))
            
            Fg1.TextMatrix(A, 2) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(A, 3) = NulosC(Rst("abrev"))
            Fg1.TextMatrix(A, 4) = NulosN(Rst("cantidad"))
            
            Fg1.TextMatrix(A, 5) = NulosN(Rst("idtipcom"))
            Fg1.TextMatrix(A, 6) = NulosC(Rst("tcompra"))
            
            Fg1.TextMatrix(A, 7) = 0
            
            Fg1.TextMatrix(A, 11) = Rst("id")
            Fg1.TextMatrix(A, 12) = NulosN(Rst("idunimed"))
            Fg1.TextMatrix(A, 13) = NulosN(Rst("idcuenta"))
            Fg1.TextMatrix(A, 14) = NulosN(Rst("idtipcom"))
            Fg1.TextMatrix(A, 15) = NulosN(Rst("ctadesdeb"))
            Fg1.TextMatrix(A, 16) = NulosN(Rst("ctadeshab"))
                    
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        Mostrando = False
    End If
End Sub



Sub ActivarEntorno()
    TabOne1.Enabled = Not TabOne1.Enabled
    Toolbar1.Enabled = Not Toolbar1.Enabled
End Sub

Private Sub CmdDetCenCos_Click()
    'If QueHace = 3 Then Exit Sub
    If ((NulosN(TxtBruto.Text) + NulosN(TxtInafecto.Text)) = 0) Then
        MsgBox "No ha especificado el importe del documento para distribuir el centro de costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If QueHace = 3 Then
        CmdAddCenCos.Enabled = False
        CmdDelCenCos.Enabled = False
        CmdCancelar.Enabled = False
        Fg5.Editable = flexEDNone
        Fg5.SelectionMode = flexSelectionByRow
    Else
        CmdAddCenCos.Enabled = True
        CmdDelCenCos.Enabled = True
        CmdCancelar.Enabled = True
        Fg5.Editable = flexEDKbdMouse
        Fg5.SelectionMode = flexSelectionFree
    End If
    ActivarEntorno
    TxtTotPor.Text = ""
    TxtTotImp.Text = ""
    Frame6.Left = 1545
    Frame6.Top = 2190
    Frame6.Visible = True
    HallarTotCenCos
End Sub


Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstComp.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstComp
End Sub


Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstComp("id")), xCon
    End If
End Sub

Sub AgregarCentroCosto2(CargarGrabado As Boolean, Optional IdCompra As Integer)
    'CargarGrabado = especifica que se levantara un centro de costos que haya sido grabado
    
    Dim Rst As New ADODB.Recordset
    Dim A, B, C, xFila As Integer
    Dim SeEncontro As Boolean
    
    Fg5.Rows = 1
        
    If CargarGrabado = True Then
        RST_Busq Rst, "SELECT com_comprascosto.idcom, com_comprascosto.idcencos, con_centrocosto.codigo, con_centrocosto.descripcion, com_comprascosto.imppor, com_comprascosto.impcos, " _
            & " con_centrocosto.tipo FROM con_centrocosto INNER JOIN com_comprascosto ON con_centrocosto.id = com_comprascosto.idcencos " _
            & " WHERE (((com_comprascosto.idcom)=" & IdCompra & "))", xCon
            
        If Rst.RecordCount <> 0 Then
            Fg5.Rows = 1
            Rst.MoveFirst
            Mostrando = True
            For A = 1 To Rst.RecordCount
                Fg5.Rows = Fg5.Rows + 1
                Fg5.TextMatrix(Fg5.Rows - 1, 1) = NulosC(Rst("codigo"))
                Fg5.TextMatrix(Fg5.Rows - 1, 2) = NulosC(Rst("descripcion"))
                Fg5.TextMatrix(Fg5.Rows - 1, 3) = Format(NulosN(Rst("imppor")), "0.00")
                Fg5.TextMatrix(Fg5.Rows - 1, 4) = Format(NulosN(Rst("impcos")), "0.00")
                Fg5.TextMatrix(Fg5.Rows - 1, 5) = NulosN(Rst("idcencos"))
                
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
            Mostrando = False
        End If
    Else
        For A = 1 To Fg1.Rows - 1
            'buscamos si el item actual tiene centros de costo definido
            RST_Busq Rst, "SELECT alm_invencencos.idpro, alm_invencencos.idcencos, con_centrocosto.codigo, con_centrocosto.descripcion, " _
            & " alm_invencencos.imppor FROM alm_invencencos LEFT JOIN con_centrocosto ON alm_invencencos.idcencos = con_centrocosto.id " _
            & " WHERE (((alm_invencencos.idpro)=" & NulosN(Fg1.TextMatrix(A, 11)) & "))", xCon
            
            If Rst.RecordCount <> 0 Then
                'si tiene centro de costos agregamos a la cuadricula centro de costos
                Rst.MoveFirst
                For B = 1 To Rst.RecordCount
                    'buscamos si el cetro de costo ya fue agregado a la cuadricula
                    SeEncontro = False
                    xFila = 0
                    For C = 1 To Fg5.Rows - 1
                        If Fg5.TextMatrix(C, 5) = Rst("idcencos") Then
                            SeEncontro = True
                            xFila = C
                        End If
                    Next C
                    
                    If SeEncontro = True Then
                        'nos pocisionamos en la fila que contiene el centro de costos y sumamos el valor
                        If Rst("imppor") < 100 Then
                            Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg5.TextMatrix(Fg5.Rows - 1, 4)) + (NulosN(Fg1.TextMatrix(A, 8)) * ((Rst("imppor") / 100) + 1))
                        Else
                            Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg5.TextMatrix(Fg5.Rows - 1, 4)) + NulosN(Fg1.TextMatrix(A, 8))
                        End If
                    Else
                        'agregamos una nueva fila a la cuadricula centro de costos
                        Fg5.Rows = Fg5.Rows + 1
                        Fg5.TextMatrix(Fg5.Rows - 1, 1) = NulosC(Rst("codigo"))
                        Fg5.TextMatrix(Fg5.Rows - 1, 2) = NulosC(Rst("descripcion"))
                        Fg5.TextMatrix(Fg5.Rows - 1, 3) = Format(NulosN(Rst("imppor")), "0.00")
                        Fg5.TextMatrix(Fg5.Rows - 1, 5) = NulosN(Rst("idcencos"))
                        If NulosN(Rst("imppor")) < 100 Then
                            Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg1.TextMatrix(A, 8)) * ((Rst("imppor") / 100) + 1)
                        Else
                            Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosN(Fg1.TextMatrix(A, 8))
                        End If
                        Fg5.TextMatrix(Fg5.Rows - 1, 4) = Format(Fg5.TextMatrix(Fg5.Rows - 1, 4), "0.00")
                    End If
                    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                Next B
            Else
                If NulosN(Fg1.TextMatrix(A, 9)) <> 0 Then
                    'MsgBox "El item " & NulosC(Fg1.TextMatrix(A, 1)) & ", no tiene especificado un centro de costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                End If
            End If
        Next A
            
        If NulosN(TxtBruto.Text) <> 0 Or NulosN(TxtInafecto.Text) <> 0 Then
            For A = 1 To Fg5.Rows - 1
                Fg5.TextMatrix(A, 3) = (NulosN(Fg5.TextMatrix(A, 4)) / (NulosN(TxtBruto.Text) + NulosN(TxtInafecto.Text))) * 100
                Fg5.TextMatrix(A, 3) = Format(Fg5.TextMatrix(A, 3), "0.00")
            Next A
        End If
    End If
    Set Rst = Nothing
End Sub

Sub AgregarCentroCosto(xIdProducto As Integer, xImporte As Double)
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    Dim SeEncontro As Boolean
    
    'buscamos si el producto tiene centro de costo asignado
    RST_Busq Rst, "SELECT alm_invencencos.idpro, alm_invencencos.idcencos, con_centrocosto.codigo, con_centrocosto.descripcion, " _
        & " alm_invencencos.imppor FROM alm_invencencos LEFT JOIN con_centrocosto ON alm_invencencos.idcencos = con_centrocosto.id " _
        & " WHERE (((alm_invencencos.idpro)=" & xIdProducto & "))", xCon
    
    If Rst.RecordCount <> 0 Then
        For A = 1 To Rst.RecordCount
            For B = 1 To Fg5.Rows - 1
                SeEncontro = False
                If Fg5.TextMatrix(B, 5) = Rst("idcencos") Then
                    SeEncontro = True
                    Exit For
                End If
            Next B
            If SeEncontro = False Then
                'si no lo encuentra lo debe de agregar a la lista de centro de costos
                Fg5.Rows = Fg5.Rows + 1
                Fg5.TextMatrix(Fg5.Rows - 1, 1) = Rst("codigo")
                Fg5.TextMatrix(Fg5.Rows - 1, 2) = Rst("descripcion")
                Fg5.TextMatrix(Fg5.Rows - 1, 3) = Format(Rst("imppor"), "0.00")
                If Rst("imppor") = 100 Then
                    Fg5.TextMatrix(Fg5.Rows - 1, 4) = xImporte * 1
                Else
                    Fg5.TextMatrix(Fg5.Rows - 1, 4) = xImporte * ((Rst("imppor") / 100) + 1)
                End If
                Fg5.TextMatrix(Fg5.Rows - 1, 4) = Format(Fg5.TextMatrix(Fg5.Rows - 1, 4), "0.00")
                Fg5.TextMatrix(Fg5.Rows - 1, 5) = Rst("idcencos")
            Else
                'si el centro de costo ya existe, agregarlo al centro de costo ya existente
                MsgBox "Falta hacer esta opcion"
                'Fg5.TextMatrix(Fg5.Rows - 1, 1) = Rst("codigo")
                'Fg5.TextMatrix(Fg5.Rows - 1, 2) = Rst("descripcion")
                'Fg5.TextMatrix(Fg5.Rows - 1, 5) = Rst("idcencos")
            End If
        Next A
    End If
End Sub

Sub BuscarImpuestos()
    If QueHace = 3 Then Exit Sub
    If Fg1.Rows = 1 Then Exit Sub
    'If NulosC(Fg1.TextMatrix(Fg1.Row, 8)) = "" Then Exit Sub
    Dim A As Integer
    Dim xImpSEL, xImpIGV As Double
    
    Dim Rst As New ADODB.Recordset
    
    TxtIGV.Text = "0.00"
    TxtIGV2.Text = "0.00"
    TxtIGV3.Text = "0.00"
    TxtOtros.Text = "0.00"
    
    Set RstTempISC = Nothing
    PreparaRST_ISC
    xImpSEL = 0
    'buscando selectivo
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, 5)) <> 4 And NulosN(Fg1.TextMatrix(A, 5)) <> 0 Then
            If NulosC(Fg1.TextMatrix(A, 2)) <> "" Then
                RST_Busq Rst, "SELECT mae_impuestos.tasa, mae_impuestos.idcuen, con_planctas.cuenta " _
                    & " FROM (alm_inventario LEFT JOIN mae_impuestos ON alm_inventario.idimpsel = mae_impuestos.id) " _
                    & " LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id WHERE " _
                    & " ((alm_inventario.id = " & NulosN(Fg1.TextMatrix(A, 11)) & " ))", xCon
                
                If Rst.RecordCount <> 0 Then
                    If NulosN(Rst("idcuen")) <> 0 Then
                        xImpSEL = xImpSEL + NulosN(Fg1.TextMatrix(A, 10)) * (NulosN(Rst("tasa")) / 100)
                        
                        If RstTempISC.RecordCount = 0 Then
                            RstTempISC.AddNew
                            RstTempISC("idcuen") = NulosN(Rst("idcuen"))
                            RstTempISC("total") = NulosN(RstTempISC("total")) + NulosN(Fg1.TextMatrix(A, 10)) * (NulosN(Rst("tasa")) / 100)
                        Else
                            RstTempISC.MoveFirst
                            RstTempISC.Find "idcuen = " & Rst("idcuen") & ""
                            
                            If RstTempISC.EOF = False Then
                                RstTempISC("idcuen") = NulosN(Rst("idcuen"))
                                RstTempISC("total") = NulosN(RstTempISC("total")) + NulosN(Fg1.TextMatrix(A, 10)) * (NulosN(Rst("tasa")) / 100)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next A
    
    TxtISC.Text = Format(NulosN(xImpSEL), "0.00")
    
    'buscando el impuesto a las ventas
    If NulosN(LblIdTipPer.Caption) <> 3 Then
        xImpIGV = 0
        
        TxtIGV.Text = NulosN(TxtBruto.Text) * (NulosN(TasaImpuesto) / 100)
        TxtIGV2.Text = NulosN(TxtBruto2.Text) * (NulosN(TasaImpuesto) / 100)
        TxtIGV3.Text = NulosN(TxtBruto3.Text) * (NulosN(TasaImpuesto) / 100)
            
        
 
        If NulosN(TxtTipDoc.Text) <> 2 Then
            TxtTotal.Text = NulosN(TxtBruto.Text) + NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text) + NulosN(TxtInafecto.Text) + NulosN(TxtIGV.Text) + NulosN(TxtIGV2.Text) + NulosN(TxtIGV3.Text) + NulosN(TxtISC.Text)
        Else
            TxtTotal.Text = NulosN(TxtBruto.Text) + NulosN(TxtInafecto.Text) - NulosN(TxtIGV.Text)
        End If
    Else
        
        xImpIGV = 0
        Dim xNeto As Double
        Dim xNeto2 As Double
        
        For A = 1 To Fg1.Rows - 1
        
            xNeto = NulosN(Busca_Codigo(NulosN(Fg1.TextMatrix(A, 19)), "id", "neto", "mae_netonodomiciliado", "N", xCon))
            
'            txt_cb(0).Text = Busca_Codigo(NulosN(RstFrm("idcuenta")), "id", "cuenta", "con_planctas", "N", xCon)
            
            
            xNeto2 = Fg1.TextMatrix(A, 10) * (xNeto / 100)
            
            xImpIGV = xNeto2 * 0.3
        Next A
    
        TxtTotal.Text = NulosN(TxtBruto.Text) + NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text) + NulosN(TxtInafecto.Text) - NulosN(TxtOtros.Text)
    
    End If
    
    TxtOtros.Text = Format(xImpIGV, FORMAT_MONTO)
    TxtTotal.Text = Format(TxtTotal.Text, FORMAT_MONTO)

    TxtIGV.Text = Format(TxtIGV.Text, FORMAT_MONTO)
    TxtIGV2.Text = Format(TxtIGV2.Text, FORMAT_MONTO)
    TxtIGV3.Text = Format(TxtIGV3.Text, FORMAT_MONTO)



End Sub

Sub PreparaRST_ISC()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(9, 3) As String

    xCampos(0, 0) = "idcuen":        xCampos(0, 1) = "N":      xCampos(0, 2) = "2"
    xCampos(1, 0) = "Total":         xCampos(1, 1) = "D":      xCampos(1, 2) = "2"
    Set RstTempISC = xFun.CrearRstTMP(xCampos)

    RstTempISC.Open
End Sub

Sub HallarTotal()
    Dim A As Integer
    Dim Total, TotalIna As Double
    Dim xPorcen As Double
    Dim PreDes As Double
    Fg5.Rows = 1
    Dim Valor As Double
    
    Dim sTotal, sBruto1, sBruto2, sBruto3 As Double
    Dim sAfecto, sInafecto As Double
    
    Total = 0
    TotalIna = 0
    sTotal = 0
    sBruto1 = 0
    sBruto2 = 0
    sBruto3 = 0
    For A = 1 To Fg1.Rows - 1
        '--inicializando valores
        sAfecto = 0
        sInafecto = 0
        
        'Se esta aplicando descuento por porcentaje
''''        If OptDes1.Value = True Then
''''            If NulosN(Fg1.TextMatrix(A, 8)) <> 0 Then
''''                xPorcen = ((NulosN(Fg1.TextMatrix(A, 8)) / 100))
''''''                '--afecto
''''''                PreDes = NulosN(Fg1.TextMatrix(A, 6)) - (NulosN(Fg1.TextMatrix(A, 6)) * xPorcen)
''''''                sAfecto = PreDes * NulosN(Fg1.TextMatrix(A, 4))
''''''                '--inafecto
''''''                sInafecto = (NulosN(Fg1.TextMatrix(A, 7)) / xPorcen) * NulosN(Fg1.TextMatrix(A, 4))
''''
''''                sTotal = (NulosN(Fg1.TextMatrix(A, 7)) / xPorcen) * NulosN(Fg1.TextMatrix(A, 4))
''''
''''            Else
''''''                sAfecto = NulosN(Fg1.TextMatrix(A, 6)) * NulosN(Fg1.TextMatrix(A, 4))
''''''                sInafecto = NulosN(Fg1.TextMatrix(A, 7)) * NulosN(Fg1.TextMatrix(A, 4))
''''
''''                sTotal = NulosN(Fg1.TextMatrix(A, 7)) * NulosN(Fg1.TextMatrix(A, 4))
''''            End If
''''        End If
''''        'Se esta aplicando descuento por importe
''''        If OptDes2.Value = True Then
''''            If NulosN(Fg1.TextMatrix(A, 8)) <> 0 Then
''''''                If NulosN(Fg1.TextMatrix(A, 6)) <> 0 Then
''''''                    sAfecto = (NulosN(Fg1.TextMatrix(A, 6)) - NulosN(Fg1.TextMatrix(A, 8))) * NulosN(Fg1.TextMatrix(A, 4))
''''''                End If
''''                If NulosN(Fg1.TextMatrix(A, 7)) <> 0 Then
''''''                    sInafecto = (NulosN(Fg1.TextMatrix(A, 7)) - NulosN(Fg1.TextMatrix(A, 8))) * NulosN(Fg1.TextMatrix(A, 4))
''''                    sTotal = (NulosN(Fg1.TextMatrix(A, 7)) - NulosN(Fg1.TextMatrix(A, 8))) * NulosN(Fg1.TextMatrix(A, 4))
''''                End If
''''            Else
''''''                sAfecto = (NulosN(Fg1.TextMatrix(A, 6)) * NulosN(Fg1.TextMatrix(A, 4)))
''''''                sInafecto = (NulosN(Fg1.TextMatrix(A, 7)) * NulosN(Fg1.TextMatrix(A, 4)))
''''                sTotal = (NulosN(Fg1.TextMatrix(A, 7)) * NulosN(Fg1.TextMatrix(A, 4)))
''''            End If
''''        End If
        
        Select Case NulosN(Fg1.TextMatrix(A, 5))
            Case 1 '--Ope Grav.
                sBruto1 = sBruto1 + sTotal
            Case 2 '--Ope Grav de Export
                sBruto2 = sBruto2 + sTotal
            Case 3 '--Sin derecho Credito Fiscal
                sBruto3 = sBruto3 + sTotal
            Case 4 '--Adquisiones No Gravadas
                TotalIna = TotalIna + sTotal

            Case Else
                
        End Select
        
        
        'AgregarCentroCosto Val(Fg1.TextMatrix(A, 8)), Val(Fg1.TextMatrix(A, 7))
    Next A


    TxtBruto.Text = Format(sBruto1, FORMAT_MONTO)
    TxtBruto2.Text = Format(sBruto2, FORMAT_MONTO)
    TxtBruto3.Text = Format(sBruto3, FORMAT_MONTO)
    
    TxtInafecto.Text = Format(TotalIna, FORMAT_MONTO)
    
    
    
'    AgregarCentroCosto2 False
End Sub

Sub CargarRSTCom()

    Dim nSQL As String
    
    nSQL = "SELECT com_reembolsables.*, [com_reembolsables].[numser] & '-' & [com_reembolsables].[numdoc] AS numerodoc, mae_documento.abrev AS tdabrev, mae_documento.descripcion AS tddesc, mae_moneda.simbolo AS moneda, mae_cliente.numruc AS clinumruc, mae_cliente.nombre AS clinombre, mae_prov.numruc AS pronumruc, mae_prov.nombre AS pronombre, [com_reembolsables].[fchdoc] & '' AS fchdoc1, mae_moneda.descripcion AS descmon, " _
        + vbCr + " [com_reembolsables].[impbru] & '' AS impbru1, [com_reembolsables].[impina] & '' AS impina1, [com_reembolsables].[impigv] & '' AS impigv1, [com_reembolsables].[imptot] & '' AS imptot1, IIf([com_reembolsables].[tc]=0,[con_tc].[impven],[com_reembolsables].[tc]) & '' AS impven1,[com_reembolsables].[impsal] & '' AS impsal1, var_ordendespacho.fchemi AS fchorden,'' as numreg1 " _
        + vbCr + " FROM (((((com_reembolsables LEFT JOIN mae_cliente ON com_reembolsables.idcli = mae_cliente.id) LEFT JOIN mae_prov ON com_reembolsables.idpro = mae_prov.id) LEFT JOIN mae_documento ON com_reembolsables.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON com_reembolsables.idmon = mae_moneda.id) LEFT JOIN con_tc ON com_reembolsables.fchdoc = con_tc.fecha) LEFT JOIN var_ordendespacho ON com_reembolsables.iddocref2 = var_ordendespacho.id " _
        + vbCr + " WHERE (((com_reembolsables.idmes)=" & mMesActivo & " ));"

    RST_Busq RstComp, nSQL, xCon
    
    
End Sub

Private Sub Fg5_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Mostrando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Dim xTot As Double
    xTot = NulosN(TxtBruto.Text) + NulosN(TxtInafecto.Text)
    
    If Col = 3 Then
        If NulosN(Fg5.TextMatrix(Fg5.Row, 3)) > 100 Then
            Fg5.TextMatrix(Fg5.Row, 3) = ""
            Fg5.TextMatrix(Fg5.Row, 4) = ""
            Exit Sub
        End If
        If NulosN(Fg5.TextMatrix(Fg5.Row, 3)) <> 0 Then
            Fg5.TextMatrix(Fg5.Row, 4) = xTot * NulosN(Fg5.TextMatrix(Fg5.Row, 3) / 100)
        End If
    End If
    
    If Col = 4 Then
        If Fg5.TextMatrix(Fg5.Row, 4) > xTot Then Exit Sub
        
        If NulosN(Fg5.TextMatrix(Fg5.Row, 4)) <> 0 And xTot <> 0 Then
            Fg5.TextMatrix(Fg5.Row, 3) = ((NulosN(Fg5.TextMatrix(Fg5.Row, 4)) / xTot) * 100)
            Fg5.TextMatrix(Fg5.Row, 3) = Format(Fg5.TextMatrix(Fg5.Row, 3), "0.00")
        End If
    End If
    
    HallarTotCenCos
End Sub

Sub HallarTotCenCos()
    Dim A As Integer
    Dim TotPor, TotImp As Double
    
    For A = 1 To Fg5.Rows - 1
        TotPor = TotPor + NulosN(Fg5.TextMatrix(A, 3))
        TotImp = TotImp + NulosN(Fg5.TextMatrix(A, 4))
    Next A
    
    TxtTotPor.Text = Format(TotPor, "0.00")
    TxtTotImp.Text = Format(TotImp, "0.00")
End Sub

Private Sub Fg5_EnterCell()
    If QueHace = 3 Then Exit Sub
    If Fg5.Col = 3 Or Fg5.Col = 4 Then
        Fg5.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Dim Rpta As Integer
        Dim Rst As New ADODB.Recordset
        
        mMesActivo = xMes
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        If xOrigen = 1 Then
        
            LblPeriodo2.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
            Nuevo
        Else
            If CONTABILIZAR = True Then
                OpcionesPeriodo
            Else
                RST_Busq RstComp, "SELECT DISTINCT com_compras.*, mae_prov.nombre, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numerodoc, " _
                    & " mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_prov.numruc, " _
                    & " mae_moneda.descripcion AS descmon, mae_moneda.simbolo, mae_tipoproducto.descripcion AS desctipcom, " _
                    & " con_tc.impcom,com_compras.fchdoc & '' as fchdoc1, com_compras.fchven & '' as fchven1, com_compras.imptot & '' as imptot1, com_compras.impsal & '' as impsal1 " _
                    & " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_condpago RIGHT JOIN ((com_compras LEFT " _
                    & " JOIN mae_tipoproducto ON com_compras.idtipo = mae_tipoproducto.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) " _
                    & " ON mae_condpago.id = com_compras.idconpag) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) " _
                    & " ON mae_prov.id = com_compras.idpro", xCon
                                       
            End If
            Set Rst = Nothing
            
            Set Dg1.DataSource = RstComp
        End If
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
        If RstComp.RecordCount = 0 Then Exit Sub
        Modificar
    End If
    
    If KeyCode = 113 Then '--F2 Grabar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace = 3 Then Exit Sub
        If Grabar = True Then
            If xOrigen = 0 Then
                Cancelar
                RstComp.Requery
                Dg1.Refresh
            Else
                QueHace = 3
                Set RstComp = Nothing
                Unload Me
                Exit Sub
            End If
        End If
    End If
    
    If KeyCode = 116 Then '--F5 actualizar
        
    
    End If
    
    If KeyCode = 117 Then '--F6 '--cancelar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace = 3 Then Exit Sub
        If xOrigen = 1 Then
            QueHace = 3
            IdCompraReg = 0
            Unload Me
            Exit Sub
        End If
        
        Cancelar
    End If
    
End Sub

Private Sub Form_Load()
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    
    Dg1.Columns("fchdoc1").NumberFormat = FORMAT_DATE
    
    
    Dg1.Columns("impbru1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impigv1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("imptot1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impsal1").NumberFormat = FORMAT_MONTO
    
    CaracteresNumericos = "0123456789." & Chr(8)
    CaracteresNumericos2 = "0123456789.-" & Chr(8) & Chr(13)
    
    Fg4.ColWidth(5) = 0
    Fg5.ColWidth(5) = 0
    
    
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    Fg1.ColWidth(13) = 0
    Fg1.ColWidth(14) = 0
    Fg1.ColWidth(15) = 0
    Fg1.ColWidth(16) = 0
    Fg1.ColWidth(17) = 0
    Fg1.ColWidth(18) = 0
    Fg1.ColWidth(19) = 0
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    
    
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ColComboList(2) = ""
    Fg1.ColComboList(5) = ""
    If CONTABILIZAR = True Then
        Toolbar1.Buttons(11).Visible = True
        LblPeriodo.Visible = True
        Frame5.Visible = True
    Else
        Toolbar1.Buttons(11).Visible = False
        LblPeriodo.Visible = False
        Frame5.Visible = False
    End If
    
    Fg4.SelectionMode = flexSelectionByRow
    Fg4.Editable = flexEDNone
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    
    '--dar formato a las columnas
'    Fg1.ColFormat(4) = "#,###,##0.0000" '--cantidad
'    Fg1.ColFormat(6) = "0.000000" '--afecto
'    Fg1.ColFormat(7) = "0.000000" '--inafecto
'    Fg1.ColFormat(8) = "0.000000" '--descuento
'    Fg1.ColFormat(9) = "0.000000" '--nvo precio
'    Fg1.ColFormat(10) = "#,###,##0.0000" '--total
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando una compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If xOrigen = 0 Then
            If RstComp.State = 0 Then Exit Sub
            If RstComp.RecordCount = 0 And QueHace <> 1 Then
                Cancel = 1
                Exit Sub
            End If
            If QueHace = 3 Then MuestraSegundoTab
        End If
    End If
End Sub

Sub Filtrar()
    'Dim xForm As New EPS_Buscar.Filtrar
    Dim xform As New eps_librerias.FormFiltrar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(7, 4) As String
   
    xCampos(0, 0) = "Tipo Documento":     xCampos(0, 1) = "abrev":         xCampos(0, 2) = "C":         xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Moneda":             xCampos(1, 1) = "simbolo":       xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Fch. Emi.":          xCampos(2, 1) = "fchdoc":        xCampos(2, 2) = "F":         xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Proveedor":          xCampos(3, 1) = "nombre":      xCampos(3, 2) = "C":         xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Forma Pago":         xCampos(4, 1) = "desccond":      xCampos(4, 2) = "C":         xCampos(4, 3) = "1500"
    xCampos(5, 0) = "Fch. Vencimiento":   xCampos(5, 1) = "fchven":        xCampos(5, 2) = "C":         xCampos(5, 3) = "1500"
    xCampos(6, 0) = "Importe":            xCampos(6, 1) = "imptot":        xCampos(6, 2) = "C":         xCampos(6, 3) = "1500"
    xCampos(7, 0) = "Saldo":              xCampos(7, 1) = "impsal":        xCampos(7, 2) = "C":         xCampos(7, 3) = "1500"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstComp   'recorset que llena el grid
    Set RstComp = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstComp
    Dg1.Refresh
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        If RstComp.State = 0 Then Exit Sub
        If RstComp.RecordCount = 0 Then Exit Sub
        'preguntamos si la compra esta vinculada a una orden de compra
''        If RstComp("idordcom") <> 0 Then
''            ' no se puede modificar una compra que tenga un orden de compra asignada
''            MsgBox "La compra no se puede modificar por tener una Orden de Compra asignada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
''            Exit Sub
''        End If
        Modificar
    End If
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            If xOrigen = 0 Then
                Cancelar
                RstComp.Requery
                Dg1.Refresh
                
                If RstComp.RecordCount <> 0 Then
                    RstComp.MoveFirst
                    RstComp.Find "id=" & mIdRegistro
                    If RstComp.EOF = True Then RstComp.MoveFirst
                End If
                
            Else
                QueHace = 3
                Set RstComp = Nothing
                Unload Me
                Exit Sub
            End If
        End If
    End If
    
    If Button.Index = 6 Then
        If xOrigen = 1 Then
            QueHace = 3
            IdCompraReg = 0
            Unload Me
            Exit Sub
        End If
        
        Cancelar
    End If
    
    If Button.Index = 8 Then
        Filtrar
    End If
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstComp.Filter = ""
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 13 Then pExportar
        
    If Button.Index = 11 Then
        
        mMesActivo = SeleccionaMes(xCon)
        'If xMes = 0 Then xMes = xMesProv
        OpcionesPeriodo
    End If
    
    If Button.Index = 14 Then
        If RstComp("tipdoc") = 4 Then
            Imprimir
        Else
            MsgBox "No puede imprimir este documento, seleccione una liquidación de compras para efectuar esta operación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    
    If Button.Index = 16 Then
        Set RstComp = Nothing
        Unload Me
    End If
End Sub

Sub OpcionesPeriodo()
     Dim NomMes As String
     Dim Rpta  As Integer
     Dim xFechaMes As String
     
     LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)

    
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    TDB_FiltroLimpiar Dg1
    Set RstComp = Nothing
    '------------------------------------------
    
    LblPeriodo.Caption = LblMes.Caption
    LblPeriodo2.Caption = LblMes.Caption
    
    If mMesActivo <> 0 And mMesActivo <> 13 Then
        xFechaMes = "01/" + Trim(Format(mMesActivo, "00")) + "/" + Trim(Format(AnoTra, "0000"))
        xFchIni = xFechaMes
        xFchFin = Format(HallaDiasMes(CDate(xFechaMes)), "00") + "/" + Mid(xFechaMes, 4, 7)
    End If
    
    CargarRSTCom
   
    Set Dg1.DataSource = RstComp
End Sub

Private Sub TxtBruto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtBruto_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosN(TxtBruto.Text) <> 0 Then
        TxtBruto.Text = Format(TxtBruto.Text, FORMAT_MONTO)
        TxtIGV.Text = NulosN(TxtBruto.Text) * xPorIgv
        TxtIGV.Text = Format(TxtIGV.Text, FORMAT_MONTO)
    Else
        TxtIGV.Text = "0.00"
    End If
    BuscarImpuestos
End Sub

Private Sub TxtBruto2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtBruto2_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosN(TxtBruto2.Text) <> 0 Then
        TxtBruto2.Text = Format(TxtBruto2.Text, FORMAT_MONTO)
        TxtIGV2.Text = NulosN(TxtBruto2.Text) * xPorIgv
        TxtIGV2.Text = Format(TxtIGV2.Text, FORMAT_MONTO)
    Else
        TxtIGV2.Text = "0.00"
    End If
    BuscarImpuestos
End Sub

Private Sub TxtBruto3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtBruto3_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosN(TxtBruto3.Text) <> 0 Then
        TxtBruto3.Text = Format(TxtBruto3.Text, FORMAT_MONTO)
        TxtIGV3.Text = NulosN(TxtBruto3.Text) * xPorIgv
        TxtIGV3.Text = Format(TxtIGV3.Text, FORMAT_MONTO)
    Else
        TxtIGV3.Text = "0.00"
    End If
    BuscarImpuestos
End Sub

Private Sub TxtDocRef2_Change()
    If QueHace = 3 Then Exit Sub
    If NulosN(TxtDocRef2.Text) = 0 Then
        LblIdDocRef2.Caption = ""
    End If
End Sub

Private Sub TxtDocRef2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Fg1.Rows = 1 Then
            CmdAddItem.SetFocus
        Else
            Fg1.Row = 1
            Fg1.Col = 1
            Fg1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDocRef2_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 116 Then
        CmdBusDocRef2_Click
    End If
    If KeyCode = 46 Then
        TxtDocRef2.Text = ""
        LblIdDocRef2.Caption = ""
    End If
End Sub

Private Sub TxtGlosa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And QueHace <> 3 Then
        SendKeys vbTab
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
        LblMoneda.Caption = ""
        Exit Sub
    End If
    Dim xRs1 As New ADODB.Recordset
    
    'buscamos el codigo de la moneda         digitada
    RST_Busq xRs1, "SELECT * FROM mae_moneda WHERE id = " & NulosN(TxtIdMon.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtIdMon.Text = ""
        LblMoneda.Caption = ""
    Else
        LblMoneda.Caption = NulosC(xRs1("descripcion"))
        
''        If Trim(TxtIdMon.Text) = "1" Then
''            LblTipCam.Visible = False
''            LblTipoCambio.Visible = False
''        Else
'            If TxtFchDoc.Valor = "" And ChkTC.Value = 0 Then
'                MsgBox "No ha especificado la fecha del documento, no se puede determinar " & Chr(13) _
'                    & "la fecha del tipo de cambio para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'
'                TxtIdMon.Text = ""
'                LblMoneda.Caption = ""
'                TxtFchDoc.SetFocus
'                Exit Sub
'            End If
'            LblTipCam.Visible = True
'            LblTipoCambio.Visible = True
'            LblTipoCambio.Caption = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
'''        End If
    End If
    Set xRs1 = Nothing
    xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
End Sub

Private Sub TxtIGV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIGV2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIGV3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtInafecto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtInafecto_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosN(TxtInafecto.Text) <> 0 Then
    Else
    End If

    BuscarImpuestos
End Sub

Private Sub TxtISC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumDoc.Text) <> "" Then
        If IsNumeric(TxtNumDoc.Text) = True Then
            TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
        End If
        If NulosC(TxtNumDoc.Text) <> "" And NulosC(TxtNumSer.Text) <> "" Then
            If ExisteNumDocCompra = True Then
                Exit Sub
            End If
        End If
    End If
End Sub

Function ExisteNumDocCompra() As Boolean
    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    If QueHace <> 1 Then nSQL = " and com_compras.id <> " & NulosN(RstComp("id"))
    
    RST_Busq Rst, "SELECT com_compras.fchdoc, Left([com_compras].[numreg],2) & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & Right([com_compras].[numreg],4) AS registro FROM com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id WHERE numser = '" & NulosC(TxtNumSer.Text) & "' and numdoc = '" & NulosC(TxtNumDoc.Text) & "' AND idpro = " & NulosN(LblIdProveedor.Caption) & nSQL, xCon
    If Rst.RecordCount = 0 Then
        ExisteNumDocCompra = False
    Else
        MsgBox "El número de documento ingresado ya fue registrado" & vbCr & "Nº Registro: " & NulosC(Rst("registro")) & vbCr & "Fecha Doc.   " & NulosC(Rst("fchdoc")) & vbCr & "Ingrese Otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.Text = ""

        ExisteNumDocCompra = True
    End If
    Set Rst = Nothing
End Function

Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumRuc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub TxtNumRuc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If TxtNumRuc.Text = "" Then Exit Sub
    Dim xRs1 As New ADODB.Recordset
    
    RST_Busq xRs1, "SELECT mae_prov.id, mae_prov.numruc, mae_prov.nombre, mae_tipoempresa.descripcion, mae_prov.tipper, mae_prov.idcondpag FROM mae_tipoempresa RIGHT JOIN mae_prov " _
        & " ON mae_tipoempresa.id = mae_prov.tipper WHERE (((mae_prov.numruc) Like '" & TxtNumRuc.Text & "%'))", xCon
    'SELECT * FROM mae_prov WHERE numruc like '" & TxtNumRuc.Text & "%' ORDER BY numruc", xCon
    
    If xRs1.RecordCount <> 0 Then
        TxtNumRuc.Text = xRs1("numruc")
        LblNomPro.Caption = xRs1("nombre")
        LblIdProveedor.Caption = xRs1("id")
        

    Else
        TxtNumRuc.Text = ""
        LblNomPro.Caption = ""
        LblIdProveedor.Caption = ""
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumSer.Text) <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        
        If NulosC(TxtNumDoc.Text) <> "" And NulosC(TxtNumSer.Text) <> "" Then
            If ExisteNumDocCompra = True Then
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub TxtRedondeo_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    TxtRedondeo.Text = Format(TxtRedondeo.Text, FORMAT_MONTO)
'    BuscarImpuestos
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

Function Grabar() As Boolean
    Dim A, B, Rpta As Integer
    Exit Function
'''    If NulosN(TxtTipDoc.Text) <> 0 Then
'''        If xCuentaDoc = 0 Then
'''            MsgBox "No se ha asignado una cuenta contable al documento " + LblNomDoc.Caption & Chr(13) _
'''                & "Asignele una cuenta en el menu Contabilidad opcion Asignar Ctas. Contables a documentos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''            Exit Function
'''        End If
'''
'''        If xIdCuenTasa = 0 Then
'''            MsgBox "El impuesto asignado al documento " + LblNomDoc.Caption & Chr(13) & " no tiene cuenta contable" & Chr(13) _
'''                & "Asignele una cuenta en el menu Contabilidad opcion Maestro de Impuestos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''            Exit Function
'''        End If
'''    End If
'''
'''    Dim Rst As New ADODB.Recordset
'''
'''    For A = 1 To Fg1.Rows - 1
'''        '--validamos cuando sea diferente a servicios
'''        If NulosN(TxtTipCom.Text) <> 5 Then
'''
'''            'validamos que el precio ingresado este en un rango de precios especificado
'''            RST_Busq Rst, "SELECT * FROM com_precios WHERE idpro = " & NulosN(Fg1.TextMatrix(A, 11)) & "", xCon
'''            If Rst.RecordCount <> 0 Then
'''                If NulosN(Fg1.TextMatrix(A, 7)) > NulosN(Rst("pretop")) Then
'''                    Set Rst = Nothing
'''                    'buscamos una autorizacion de ingreso para el precio del proveedor
'''                    RST_Busq Rst, "SELECT com_preciosdet.idpro, com_preciosdet.fecreg, com_preciosdet.idprov, com_preciosdet.precio" _
'''                        & " From com_preciosdet  " _
'''                        & " WHERE (((com_preciosdet.idpro)=" & NulosN(Fg1.TextMatrix(A, 10)) & ") AND ((com_preciosdet.fecreg)=CDate('" & Format(TxtFchDoc.Valor, "dd/mm/yyyy") & " ')) " _
'''                        & " AND ((com_preciosdet.idprov)=" & NulosN(LblIdProveedor.Caption) & "))", xCon
'''
'''                    If Rst.RecordCount = 0 Then
'''                        'si no encontramos una autorizacion de precio para el proveedor en el dia de la operacion se rechaza
'''                        MsgBox "El precio ingresado para el item " + NulosC(Fg1.TextMatrix(A, 2)) & Chr(13) _
'''                            & "excede el precio fijado por el administrador de precios, verifique el precio fijado" & Chr(13) _
'''                            & "en el modulo de Compras opcion  Fijar Precios de Compra a Item", vbCritical + vbOKOnly + vbDefaultButton1, xTitulo
'''                        Set Rst = Nothing
'''                        Exit Function
'''                    Else
'''                        If NulosN(Fg1.TextMatrix(A, 7)) > NulosN(Rst("precio")) Then
'''                            'si el precio ingresado es aun mayor que el precio autorizado se rechaza la compra
'''                            MsgBox "El precio ingresado para el item " + NulosC(Fg1.TextMatrix(A, 2)) & Chr(13) _
'''                                & "excede el precio fijado por el administrador de precios, verifique el precio fijado" & Chr(13) _
'''                                & "en el modulo de Compras opcion  Fijar Precios de Compra a Item", vbCritical + vbOKOnly + vbDefaultButton1, xTitulo
'''                            Set Rst = Nothing
'''                            Exit Function
'''                        End If
'''                    End If
'''                End If
'''            End If
'''
'''            Set Rst = Nothing
'''
'''            'validamos que el ingreso de items no exceda el stock maximo
'''            If (OptOpera1.Value = True) Or (OptOpera2.Value = True) Then
'''                RST_Busq Rst, "SELECT * FROM alm_inventario WHERE id = " & NulosN(Fg1.TextMatrix(A, 11)) & "", xCon
'''
'''                If Rst.RecordCount <> 0 Then
'''                    If (NulosN(Rst("stckact")) + NulosN(Fg1.TextMatrix(A, 4))) > NulosN(Rst("stckmax")) Then
'''                        Rpta = MsgBox("La cantidad sumada al stock actual del item " & NulosC(Fg1.TextMatrix(A, 2)) & Chr(13) _
'''                            & "sobrepasa el Stock Maximo asignado ¿Esta seguro de agregar la cantidad especificada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
'''                        If Rpta = vbNo Then
'''                            Set Rst = Nothing
'''                            Exit Function
'''                        End If
'''                    End If
'''                End If
'''            End If
'''        End If
'''
'''        'validamos la cuenta contable del item
'''        If NulosN(Fg1.TextMatrix(A, 13)) = 0 Then
'''            MsgBox "No se le ha asignado una Cuenta Contable al item : " & Chr(13) _
'''                & Fg1.TextMatrix(A, 1) & Chr(13) _
'''                & "Asígnele una cuenta en el menu Almacén opción Mantenimiento Items de Compra y Venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''            Exit Function
'''        End If
'''    Next A
'''
'''
'''    If TxtTipCom.Text = "" Then
'''        MsgBox "No ha especificado el Tipo de Compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''        TxtTipCom.SetFocus
'''        Exit Function
'''    End If
'''
'''    If TxtNumRuc.Text = "" Then
'''        MsgBox "No ha especificado proveedor de la compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''        TxtNumRuc.SetFocus
'''        Exit Function
'''    End If
'''
'''    If TxtNumSer.Text = "" Or TxtNumDoc.Text = "" Then
'''        MsgBox "No ha especificado el numero del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''        TxtNumSer.SetFocus
'''        Exit Function
'''    End If
'''
'''    If TxtFchDoc.Valor = "" Then
'''        MsgBox "No ha especificado la fecha de emision del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''        TxtFchDoc.SetFocus
'''        Exit Function
'''    End If
'''
'''    If TxtFchVen.Valor = "" Then
'''        MsgBox "No ha especificado la fecha de vencimiento del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''        TxtFchVen.SetFocus
'''        Exit Function
'''    End If
'''
'''    If TxtConPag.Text = "" Then
'''        MsgBox "No ha especificado la condicion de pago del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''        TxtConPag.SetFocus
'''        Exit Function
'''    End If
'''
'''    If TxtIdMon.Text = "" Then
'''        MsgBox "No ha especificado la moneda del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''        TxtIdMon.SetFocus
'''        Exit Function
'''    End If
'''
'''    'If DetCenCos = False Then
'''    '    If TxtCodCenCos.Text = "" Then
'''    '        MsgBox "No ha especificado el centro de costos al que pertenece la compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''    '        TxtCodCenCos.SetFocus
'''     '       Exit Function
'''     '   End If
'''    'Else
'''
'''
''''ACTIVAR SE MODIFICO PARA QUE CORRA EN SAVAR
''''        If Fg5.Rows = 1 Then
''''            MsgBox "No ha especificado el centro de costo detallado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
''''            Exit Function
''''        End If
'''    'End If
'''
'''    If Fg1.Rows = 1 Then
'''        MsgBox "No ha especificado items para la compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''        '------------
'''        Fg1.Col = 1
'''        '------------
'''        Fg1.SetFocus
'''        Exit Function
'''    End If
'''
'''    'verificamos que la fecha de vencimiento no sea menor a la fecha de vencimiento
'''    If CDate(TxtFchDoc.Valor) > CDate(xFchFin) Then
'''        MsgBox "La fecha de vencimiento del documento no puede ser menor a la fecha de emision", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''        TxtFchVen.SetFocus
'''        Exit Function
'''    End If
'''
'''    If NulosN(TxtTipCom.Text) <> 5 Then
'''        If NulosC(TxtIdAlmacen.Text) = "" Then
'''            MsgBox "No ha especificado el almacen de destino de la compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''            TxtIdAlmacen.SetFocus
'''            Exit Function
'''        End If
'''    End If
'''    'verificamos que la fecha de vencimiento no sea mayor al periodo contable
'''    If CDate(TxtFchVen.Valor) > (CDate(xFchFin)) Then
'''        If NulosC(TxtFchPago.Valor) = "" Then
'''            MsgBox "No puede registrar este documento en el mes de " + Trim(LblPeriodo.Caption) + ", la fecha de " & Chr(13) _
'''                & "vencimiento es mayor a la fecha del periodo, para registrar este documento en el periodo" & Chr(13) _
'''                & "actual ingrese la fecha de pago menor o igual a la fecha de cierre del periodo ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''            Exit Function
'''        End If
'''    End If
'''
'''    'VERIFICAMOS QUE LOS ITEMS IGRESADOS SON LOS CORRECTOS
'''
'''    'VERIFICAMOS QUE NO EXISTAS FILAS SIN ITEMAS
'''    For A = 1 To Fg1.Rows - 1
'''        If NulosC(Fg1.TextMatrix(A, 2)) = "" Then
'''            Fg1.RemoveItem A
'''        End If
'''    Next A
'''
'''    If Fg1.Rows <> 1 Then
'''        For A = 1 To Fg1.Rows - 1
'''            If NulosN(Fg1.TextMatrix(A, 4)) = 0 Then
'''                If NulosN(TxtTipCom.Text) <> 5 Then
'''                    MsgBox "No ha especificado la cantidad para el item : " + Trim(Fg1.TextMatrix(A, 2)), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''                    Fg1.Col = 4: Fg1.Row = A
'''                    Fg1.SetFocus
'''                    Exit Function
'''                Else
'''                    '--cantidad por defecto cuando sea tipo de item = servicios
'''                    Fg1.TextMatrix(A, 4) = 1
'''                End If
'''            End If
'''        Next A
'''    Else
'''        MsgBox "No se ha especificado ningún item para esta compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''    End If
'''
'''    'verificamos que el total de items sea igual al total de los totales
'''
'''    A = NulosN(Format(GRID_SUMAR_COL(Fg1, 10), "0.00")) + NulosN(TxtRedondeo.Text)
'''
'''    B = NulosN(TxtBruto.Text) + NulosN(TxtBruto2.Text) + NulosN(TxtBruto3.Text) + NulosN(TxtInafecto.Text)
'''
'''    If Round(A, 2) <> Round(B, 2) Then
'''        MsgBox "El monto del detalle del documento no coincide con la sumatoria de los totales" & vbCr & "Diferencia: " & Format(B - A, FORMAT_MONTO), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''        TxtBruto.SetFocus
'''        Exit Function
'''    End If
'''
'''
'''
'''
'''    Dim RstDeta2 As New ADODB.Recordset
'''    Dim RstActPro As New ADODB.Recordset
'''    Dim RstCab As New ADODB.Recordset
'''    Dim RstDet As New ADODB.Recordset
'''    Dim RstDia As New ADODB.Recordset
'''    Dim RstCosto As New ADODB.Recordset
'''    Dim RstReclasifica As New ADODB.Recordset
'''
'''
'''    Dim xIdCuen, xId As Integer
'''    Dim xTotal As Double
'''    Dim xNumAsiento As String
'''    Dim xSaldo As Double '--indica el saldo actual del documento
'''
'''    On Error GoTo LaCague
'''
'''    xCon.BeginTrans
'''
'''    Me.MousePointer = vbHourglass
'''
'''    If QueHace = 1 Then
'''        xId = HallaCodigoTabla("com_compras", xCon, "id")
'''
'''        xNumAsiento = NuevoNumAsiento(1, mMesActivo, xCon)
'''
'''        RST_Busq RstCab, "SELECT TOP 1 * FROM com_compras", xCon
'''
'''        RstCab.AddNew
'''        RstCab("id") = xId
'''        IdCompraReg = xId
'''
'''        If NulosN(TxtTipDoc.Text) = 7 Then
'''            xSaldo = 0
'''        Else
'''            xSaldo = NulosN(TxtTotal.Text)
'''        End If
'''    Else
'''
'''        xId = RstComp("id")
'''
'''        RST_Busq RstCab, "SELECT * FROM com_compras WHERE id = " & xId & "", xCon
'''
'''        '------------------------------------------
'''        'eliminamos el sotck agregado con la compra
'''        If NulosN(TxtTipCom.Text) <> 5 Then
'''            RST_Busq RstDeta2, "SELECT com_comprasdet.* From com_comprasdet WHERE (((com_comprasdet.idcom)=" & xId & "))", xCon
'''
'''            If RstDeta2.RecordCount <> 0 Then
'''                RstDeta2.MoveFirst
'''                For A = 1 To RstDeta2.RecordCount
'''                    RST_Busq RstActPro, "SELECT alm_inventario.id, alm_inventario.stckact  From alm_inventario WHERE ((alm_inventario.id=" & RstDeta2("iditem") & "))", xCon
'''                    If RstActPro.RecordCount = 1 Then
'''                        RstActPro("stckact") = RstActPro("stckact") - RstDeta2("canpro")
'''                        RstActPro.Update
'''                    End If
'''                    Set RstActPro = Nothing
'''                Next A
'''            End If
'''            Set RstDeta2 = Nothing
'''        End If
'''        '----------------------------------
'''        'eliminamos el detalle de la compra
'''        xCon.Execute "DELETE * FROM com_comprasdet WHERE idcom = " & xId & ""
'''
'''
'''        Set RstDia = Nothing
'''
'''        '------------------------------
'''        'eliminamos el asiento contable
'''        If mMesActivo = 0 Then
'''            xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & mMesActivo & " AND idlib = 36 AND idmov = " & xId & ""
'''        Else
'''            xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & mMesActivo & " AND idlib = 1 AND idmov = " & xId & ""
'''        End If
'''
'''
'''
'''        '------------------------------
'''        'eliminamos el centro de costos
'''        xCon.Execute "DELETE * FROM com_comprascosto WHERE idcom = " & xId & ""
'''        'eliminamos si se reclasifica la cuenta
'''        xCon.Execute "DELETE * FROM com_comprasreclasifica WHERE idcom = " & xId & ""
'''
'''
'''        xNumAsiento = Mid(RstComp("numreg"), 3, 4)
'''
'''        '-------------------------------------------------------------
'''        'Borramos los flag de las tablas alm_ingreso y com_ordencompra
'''        If OptOpera3 = True Then
'''            'actualizamos campo idfac en la tabla alm_igreso a 0 para que se vuelva a procesar
'''            'xCon.Execute "UPDATE alm_ingreso SET alm_ingreso.idfac = 0 WHERE (((alm_ingreso.idfac)=" & RstComp("id") & "))"
'''            xCon.Execute "DELETE * FROM alm_ingresodoc WHERE iddoc = " & RstComp("id") & " "
'''        End If
'''
'''        If OptOpera2 = True Then
'''            'actualizamos campo idfac en la tabla com_ordencompra a 0 para que se vuelva a procesar
'''            xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idfac = 0 WHERE (((com_ordencompra.idfac)=" & RstComp("id") & "))"
'''        End If
'''
'''        '************************************************************************************************************************
'''        '--obtener el ultimo saldo
'''        Dim nSQL As String
'''
'''        nSQL = "SELECT " & NulosN(TxtTotal.Text) & " AS imptotori, " & NulosN(TxtIdMon.Text) & " AS idmonori,  Sum(det.imptotsol) AS totsol, Sum(det.imptotdol) AS totdol, " _
'''            & " IIf([idmonori]=1,[imptotori]-[totsol],[imptotori]-[totdol]) AS saldo FROM " _
'''            & " (SELECT tes_cajadestinodet.iddoc, Mid([numreg],1,2) & '01' & Mid([numreg],3,4) AS registro, 'Tesoreria - Egresos' AS modulo, tes_caja.idmon, " _
'''            & " con_tc.impven AS tipcam, tes_cajadestinodet.acuenta AS imptotal, IIf([tes_caja]![idmon]=1,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]*[con_tc]![impven]) AS imptotsol, " _
'''            & " IIf([tes_caja]![idmon]=2,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]/[con_tc]![impven]) AS imptotdol " _
'''            & " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) INNER JOIN tes_cajadestino ON tes_caja.id = tes_cajadestino.idtes) " _
'''            & " INNER JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) " _
'''            & " LEFT JOIN (tes_cajaori LEFT JOIN (tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) " _
'''            & " ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes " _
'''            & " Where (((tes_cajadestinodet.idmod) = 1) And ((tes_caja.tipmov) = 2)) And tes_cajadestinodet.iddoc = " & xId & "  " _
'''            & " Union " _
'''            & " SELECT con_canjesdet.iddoc, Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null Or mae_libros.codsun='','FF',mae_libros.codsun) & con_diario.numasi AS registro, " _
'''            & " 'Tesorería - Canje de documentos' AS modulo, con_canjes.idmon, con_tc.impven AS tipcam, con_diario.imphabsol AS imptotal, " _
'''            & " IIf(con_canjes.idmon=1,con_diario.imphabsol,IIf(con_tc.impven Is Null Or con_tc.impven=0 Or con_diario.imphabdol Is Null " _
'''            & " Or con_diario.imphabdol=0,0,con_diario.imphabdol*con_tc.impven)) AS imptotsol, IIf(con_canjes.idmon=2,con_diario.imphabdol," _
'''            & " IIf(con_tc.impven Is Null Or con_tc.impven=0 Or con_diario.imphabsol Is Null Or con_diario.imphabsol=0,0,con_diario.imphabsol/con_tc.impven)) AS imptotdol " _
'''            & " FROM (((con_canjes LEFT JOIN con_diario ON con_canjes.id = con_diario.idmov) LEFT JOIN con_tc ON con_canjes.fchemi = con_tc.fecha) " _
'''            & " LEFT JOIN mae_libros ON con_canjes.idlib = mae_libros.id) LEFT JOIN con_canjesdet ON (con_diario.idmov = con_canjesdet.idcan) " _
'''            & " AND (con_diario.iddocpro = con_canjesdet.iddoc) Where (((con_diario.idlib) = 8) And ((con_canjesdet.Tipo) = 2)) And con_canjesdet.iddoc = " & xId & ""
'''
'''        nSQL = nSQL & " Union " + vbCr + " SELECT com_compras.iddocref, Mid([numreg],1,2) & '01' & Mid([numreg],3,4) AS registro, 'Compras - Nota Credito' AS modulo, com_compras.idmon, " _
'''            & " con_tc.impven AS tipcam, com_compras.imptot, IIf([com_compras]![idmon]=1,[com_compras]![imptot],[com_compras]![imptot]*[con_tc]![impven]) AS imptotsol, " _
'''            & " IIf([com_compras]![idmon]=2,[com_compras]![imptot],[com_compras]![imptot]/[con_tc]![impven]) AS imptotdol FROM com_compras LEFT JOIN con_tc " _
'''            & " ON com_compras.fchdoc = con_tc.fecha Where (((com_compras.iddocref) = " & xId & ")) " _
'''            & " Union " _
'''            & " SELECT con_devolucionesdet.idcom AS iddoc, Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF'," _
'''            & " [mae_libros].[codsun]) & [con_diario].[numasi] AS registro, 'Tesorería - Rendición de Cuenta' AS modulo, con_devoluciones.idmon, con_tc.impven AS tipcam, " _
'''            & " con_diario.impdebsol AS imptotal, IIf([con_devoluciones].[idmon]=1,[con_diario].[impdebsol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or " _
'''            & " [con_diario].[impdebdol] Is Null Or [con_diario].[impdebdol]=0,0,[con_diario].[impdebdol]*[con_tc].[impven])) AS imptotsol, " _
'''            & " IIf([con_devoluciones].[idmon]=2,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebsol] Is Null " _
'''            & " Or [con_diario].[impdebsol]=0,0,[con_diario].[impdebsol]/[con_tc].[impven])) AS imptotdol FROM mae_libros RIGHT JOIN ((con_devoluciones LEFT JOIN " _
'''            & " con_tc ON con_devoluciones.fchemi = con_tc.fecha) RIGHT JOIN (con_devolucionesdet INNER JOIN con_diario ON (con_devolucionesdet.idcom = con_diario.iddocpro) " _
'''            & " AND (con_devolucionesdet.id = con_diario.idmov)) ON con_devoluciones.id = con_devolucionesdet.id) ON mae_libros.id = con_diario.idlib " _
'''            & " Where (((con_diario.idlib) = 39)) And con_devolucionesdet.idcom = " & xId & " ) as det"
'''
'''
'''        RST_Busq Rst, nSQL, xCon
'''
'''        xSaldo = NulosN(TxtTotal.Text)
'''
'''        If Rst.RecordCount <> 0 Then
'''            If NulosN(TxtTipDoc.Text) = 7 Then
'''                xSaldo = 0
'''            Else
'''                If IsNull(Rst("saldo")) = False Then xSaldo = NulosN(Rst("saldo"))
'''            End If
'''        Else
'''            xSaldo = 0
'''        End If
'''        Set Rst = Nothing
'''        '************************************************************************************************************************
'''    End If
'''
'''    RST_Busq RstDet, "SELECT TOP 1 * FROM com_comprasdet", xCon
'''    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
'''    RST_Busq RstCosto, "SELECT TOP 1 * FROM com_comprascosto", xCon
'''    RST_Busq RstReclasifica, "SELECT TOP 1 * FROM com_comprasreclasifica", xCon
'''
'''    mIdRegistro = xId
'''
'''    RstCab("idlib") = 1
'''    RstCab("idtipo") = NulosN(TxtTipCom.Text)
'''    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
'''    RstCab("idpro") = NulosN(LblIdProveedor.Caption)
'''    RstCab("numser") = TxtNumSer.Text
'''    RstCab("numdoc") = TxtNumDoc.Text
'''    RstCab("fchdoc") = TxtFchDoc.Valor
'''    RstCab("fchven") = TxtFchVen.Valor
'''    If IsDate(TxtFchPago.Valor) = True Then RstCab("fchpag") = TxtFchPago.Valor
'''    RstCab("idconpag") = NulosN(TxtConPag.Text)
'''    RstCab("idmon") = NulosN(TxtIdMon.Text)
'''    RstCab("impbru") = NulosN(TxtBruto.Text)
'''    RstCab("impbru2") = NulosN(TxtBruto2.Text)
'''    RstCab("impbru3") = NulosN(TxtBruto3.Text)
'''    RstCab("impina") = NulosN(TxtInafecto.Text)
'''    RstCab("impigv") = NulosN(TxtIGV.Text)
'''    RstCab("impigv2") = NulosN(TxtIGV2.Text)
'''    RstCab("impigv3") = NulosN(TxtIGV3.Text)
'''    RstCab("otroscargos") = NulosN(TxtOtros.Text)
'''    RstCab("imptot") = NulosN(TxtTotal.Text)
'''    RstCab("glosa") = NulosC(TxtGlosa.Text)
'''
'''    If NulosN(TxtTipDocRef.Text) <> 0 Then
'''        RstCab("idtipdocref") = NulosN(TxtTipDocRef.Text)
'''        RstCab("iddocref2") = NulosN(LblIdDocRef2.Caption)
'''    Else
'''        RstCab("idtipdocref") = 0
'''        RstCab("iddocref2") = 0
'''    End If
'''
'''    RstCab("impsal") = xSaldo
'''
'''    RstCab("impisc") = NulosN(TxtISC.Text)
'''
'''    If NulosN(TxtTipCom.Text) <> 5 Then
'''        RstCab("idalm") = NulosN(TxtIdAlmacen.Text)
'''    End If
'''
'''    'documento al que hace referencia en caso de ser nota de credito
'''    RstCab("iddocref") = NulosN(LblIdDocRef.Caption)
'''
'''    'Actualizamos el saldo del documento
'''    If NulosN(TxtTipCom.Text) = 7 Then
'''
'''        ActualizaSaldoDoc NulosN(LblIdDocRef.Caption), 1, NulosN(TxtTotal.Text)
'''
'''    End If
'''
'''    RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''
'''    If CONTABILIZAR = True Then
'''        RstCab("numreg") = Format(Trim(Str(mMesActivo)), "00") + xNumAsiento
'''    End If
'''
'''    'grabamos el tipo de descuento
'''    If OptDes1.Value = True Then
'''        RstCab("tipdes") = 1
'''    End If
'''    If OptDes2.Value = True Then
'''        RstCab("tipdes") = 2
'''    End If
'''
'''    If OptSi.Value = True Then
'''        RstCab("afecto") = -1
'''    Else
'''        RstCab("afecto") = 0
'''    End If
'''
'''    'especificamos como en que contexto se esta haciendo la compra
'''    If OptOpera1.Value = True Then RstCab("tipcom") = 1  'Compra normal
'''    If OptOpera3.Value = True Then RstCab("tipcom") = 2  'Compra vinculada con documentos de entrada
'''    If OptOpera2.Value = True Then RstCab("tipcom") = 3  'Compra vinculada con Orden de Compra
'''
'''    '--redondeo a centimos
'''    RstCab("impred") = NulosN(TxtRedondeo.Text)
'''    '--tipo de cambio
'''    RstCab("tc") = NulosN(TxtTC.Text)
'''
'''
'''    RstCab.Update
'''
'''    'Grabamos los items de la compra
'''    For A = 1 To Fg1.Rows - 1
'''        RstDet.AddNew
'''        RstDet("idcom") = xId
'''        RstDet("iditem") = NulosN(Fg1.TextMatrix(A, 11))
'''        RstDet("idunimed") = NulosN(Fg1.TextMatrix(A, 12))
'''        RstDet("canpro") = NulosN(Fg1.TextMatrix(A, 4))
'''        RstDet("preunibru") = NulosN(Fg1.TextMatrix(A, 7)) 'precio bruto afecto
'''        RstDet("preunibruina") = 0
'''        RstDet("valdes") = NulosN(Fg1.TextMatrix(A, 8))
'''        RstDet("preuni") = NulosN(Fg1.TextMatrix(A, 9))
'''        RstDet("imptot") = NulosN(Fg1.TextMatrix(A, 10))
'''        '---
'''        RstDet("idtipcom") = NulosN(Fg1.TextMatrix(A, 5))
'''
'''        RstDet.Update
'''
'''        If NulosN(TxtTipCom.Text) = 1 Or NulosN(TxtTipCom.Text) = 4 Or NulosN(TxtTipCom.Text) = 2 Then
'''            RST_Busq RstActPro, "SELECT alm_inventario.id, alm_inventario.stckact FROM alm_inventario WHERE (((alm_inventario.id)=" & NulosN(Fg1.TextMatrix(A, 9)) & "))", xCon
'''
'''            If RstActPro.RecordCount = 1 Then
'''                RstActPro("stckact") = NulosN(RstActPro("stckact")) + NulosN(Fg1.TextMatrix(A, 4))
'''                RstActPro.Update
'''            End If
'''            Set RstActPro = Nothing
'''        End If
'''
'''    Next A
'''
'''
'''    '--grabamos si se reclasifica la cuenta
'''    If NulosN(LbIdCuentaDeb.Caption) <> 0 And NulosN(LbIdCuentaHab.Caption) <> 0 Then
'''        RstReclasifica.AddNew
'''        RstReclasifica("idcom") = xId
'''        RstReclasifica("idcuendeb") = NulosN(LbIdCuentaDeb.Caption)
'''        RstReclasifica("idcuenhab") = NulosN(LbIdCuentaHab.Caption)
'''        RstReclasifica("imptot") = NulosN(txtTotal1.Text)
'''        RstReclasifica.Update
'''    End If
'''
'''    '-------------------------------------------------------------------------------------------------------------------------------------------
'''    'Actualizamos los documentos relacionados con la factura
'''    If OptOpera3.Value = True Then
'''        If Fg4.Rows <> 1 Then
'''            For A = 1 To Fg4.Rows - 1
'''                'actualizamos el flag de los partes de entrada para saber con que documento de compra se valorizaran
'''                'xCon.Execute "UPDATE alm_ingreso SET alm_ingreso.idfac = " & xId & " WHERE (((alm_ingreso.id)=" & NulosN(Fg4.TextMatrix(A, 5)) & "))"
'''                xCon.Execute "INSERT INTO alm_ingresodoc (id, iddoc) values (" & NulosN(Fg4.TextMatrix(A, 5)) & "," & xId & ")"
'''            Next A
'''        End If
'''    End If
'''
'''    If OptOpera2.Value = True Then
'''        If Fg4.Rows <> 3 Then
'''            For A = 1 To Fg4.Rows - 1
'''                'actualizamos el flag de las ordenes de compra para saber con que documento ingresaron las ordenes de compra
'''                xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idfac = " & xId & " WHERE (((com_ordencompra.id)=" & NulosN(Fg4.TextMatrix(A, 5)) & "))"
'''            Next A
'''        End If
'''    End If
'''
'''    '----------------------------------------------------------------------------------------------------------------
'''    '----------------------------------------------------------------------------------------------------------------
'''    'grabamos el centro de costos
''''''    If Fg5.Rows > 1 Then
''''''        For A = 1 To Fg5.Rows - 1
''''''            RstCosto.AddNew
''''''            RstCosto("idcom") = xId
''''''            RstCosto("idcencos") = NulosN(Fg5.TextMatrix(A, 5))
''''''            RstCosto("imppor") = NulosN(Fg5.TextMatrix(A, 3))
''''''            RstCosto("impcos") = NulosN(Fg5.TextMatrix(A, 4))
''''''            RstCosto.Update
''''''        Next A
''''''    End If
'''
'''    xCon.Execute "insert into com_comprascosto (idcom,idcencos,imppor,impcos) " _
'''        & " SELECT com_comprasdet.idcom, alm_invencencos.idcencos, alm_invencencos.imppor, [com_comprasdet].[imptot]*([alm_invencencos].[imppor]/100) AS impcos " _
'''        & " FROM com_comprasdet INNER JOIN alm_invencencos ON com_comprasdet.iditem = alm_invencencos.idpro " _
'''        & " WHERE (((com_comprasdet.idcom)= " & xId & " ));"
'''    '----------------------------------------------------------------------------------------------------------------
'''    '----------------------------------------------------------------------------------------------------------------
'''    'En caso de estar vinculada a una orden de compra actualizamos la orden de compra "3 = PROCESADA"
'''''''''    xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idest = 3 WHERE (((com_ordencompra.id)=" & Val(TxtNumOrdCom.Text) & "))"
'''
'''    If CONTABILIZAR = True Then
'''        '---------------------------------------
'''        'Grabamos el libro diario del movimiento
'''        '-------------------------------------------------------------------------
'''        'grabamos a facturas por pagar Plan de cuentas 42.1 o dependiendo del caso
'''        RstDia.AddNew
'''        RstDia("año") = AnoTra
'''        RstDia("idmes") = mMesActivo
'''        RstDia("idlib") = 1
'''        RstDia("idmov") = xId
'''        RstDia("numasi") = xNumAsiento
'''        RstDia("tc") = ValTipCam
'''        RstDia("idcue") = xCuentaDoc
'''        RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''        RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''        If NulosN(TxtTipDoc.Text) <> 0 Then
'''            If NulosN(TxtTipDoc.Text) <> 7 Then
'''                'cuando se factura u otro comprabante excepto nota de credito hace su asiento norma
'''                If TxtIdMon.Text = "1" Then
'''                    RstDia("imphabsol") = NulosN(TxtTotal.Text)
'''                    RstDia("imphabdol") = 0
'''                Else
'''                    RstDia("imphabsol") = NulosN(TxtTotal.Text) * NulosN(TxtTC.Text)
'''                    RstDia("imphabdol") = NulosN(TxtTotal.Text)
'''                End If
'''            Else
'''                'cuando sea nota de credito hace el asiento inverso al de una venta
'''                If TxtIdMon.Text = "1" Then
'''                    RstDia("impdebsol") = NulosN(TxtTotal.Text)
'''                    RstDia("impdebdol") = 0
'''                Else
'''                    RstDia("impdebsol") = NulosN(TxtTotal.Text) * NulosN(TxtTC.Text)
'''                    RstDia("impdebdol") = NulosN(TxtTotal.Text)
'''                End If
'''            End If
'''        End If
'''        RstDia.Update
'''
'''        '-----------------------------------------------------
'''        'grabamos el impuesto si la operacion esta afecta a el
'''        If NulosN(TxtIGV.Text) <> 0 Then
'''                xIdCuenTasa = fCtaImpuestoTipoCompra(1)
'''                If xIdCuenTasa = 0 Then GoTo LaCague
'''                '-------------------------------------
'''                RstDia.AddNew
'''                RstDia("año") = AnoTra
'''                RstDia("idmes") = mMesActivo
'''                If mMesActivo = 0 Then
'''                    RstDia("idlib") = 36
'''                Else
'''                    RstDia("idlib") = 1
'''                End If
'''                RstDia("idmov") = xId
'''                RstDia("numasi") = xNumAsiento
'''                RstDia("tc") = ValTipCam
'''
'''                RstDia("idcue") = xIdCuenTasa
'''                RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''                RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''
'''                'si el tipo de l proveedor es diferente a no domiciliado
'''                If NulosN(LblIdTipPer.Caption) <> 3 Then
'''                    If NulosN(TxtTipDoc.Text) <> 0 Then
'''                        If NulosN(TxtTipDoc.Text) <> 7 And NulosN(TxtTipDoc.Text) <> 2 Then
'''                            If TxtIdMon.Text = "1" Then
'''                                RstDia("impdebsol") = NulosN(TxtIGV.Text)
'''                                RstDia("impdebdol") = 0
'''                            Else
'''                                RstDia("impdebsol") = NulosN(TxtIGV.Text) * NulosN(TxtTC.Text)
'''                                RstDia("impdebdol") = NulosN(TxtIGV.Text)
'''                            End If
'''                        Else
'''                            If TxtIdMon.Text = "1" Then
'''                                RstDia("imphabsol") = NulosN(TxtIGV.Text)
'''                                RstDia("imphabdol") = 0
'''                            Else
'''                                RstDia("imphabsol") = NulosN(TxtIGV.Text) * NulosN(TxtTC.Text)
'''                                RstDia("imphabdol") = NulosN(TxtIGV.Text)
'''                            End If
'''                        End If
'''                    End If
'''                Else
'''                    If TxtIdMon.Text = "1" Then
'''                        RstDia("imphabsol") = NulosN(TxtIGV.Text)
'''                        RstDia("imphabdol") = 0
'''                    Else
'''                        RstDia("imphabsol") = NulosN(TxtIGV.Text) * NulosN(TxtTC.Text)
'''                        RstDia("imphabdol") = NulosN(TxtIGV.Text)
'''                    End If
'''                End If
'''                RstDia.Update
'''            Else
'''            End If
'''        End If
'''
'''        '***********************************************************************
'''        'grabamos el impuesto si la operacion esta no afecta a el
'''        If NulosN(TxtIGV2.Text) <> 0 Then
'''            xIdCuenTasa = fCtaImpuestoTipoCompra(2)
'''            If xIdCuenTasa = 0 Then GoTo LaCague
'''            '-------------------------------------
'''            RstDia.AddNew
'''            RstDia("año") = AnoTra
'''            RstDia("idmes") = mMesActivo
'''            If mMesActivo = 0 Then
'''                RstDia("idlib") = 36
'''            Else
'''                RstDia("idlib") = 1
'''            End If
'''            RstDia("idmov") = xId
'''            RstDia("numasi") = xNumAsiento
'''            RstDia("tc") = ValTipCam
'''            RstDia("idcue") = xIdCuenTasa
'''            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''
'''            'si el tipo de l proveedor es diferente a no domiciliado
'''            If NulosN(LblIdTipPer.Caption) <> 3 Then
'''                If NulosN(TxtTipDoc.Text) <> 0 Then
'''                    If NulosN(TxtTipDoc.Text) <> 7 And NulosN(TxtTipDoc.Text) <> 2 Then
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("impdebsol") = NulosN(TxtIGV2.Text)
'''                            RstDia("impdebdol") = 0
'''                        Else
'''                            RstDia("impdebsol") = NulosN(TxtIGV2.Text) * NulosN(TxtTC.Text)
'''                            RstDia("impdebdol") = NulosN(TxtIGV2.Text)
'''                        End If
'''                    Else
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("imphabsol") = NulosN(TxtIGV2.Text)
'''                            RstDia("imphabdol") = 0
'''                        Else
'''                            RstDia("imphabsol") = NulosN(TxtIGV2.Text) * NulosN(TxtTC.Text)
'''                            RstDia("imphabdol") = NulosN(TxtIGV2.Text)
'''                        End If
'''                    End If
'''                End If
'''            Else
'''                If TxtIdMon.Text = "1" Then
'''                    RstDia("imphabsol") = NulosN(TxtIGV2.Text)
'''                    RstDia("imphabdol") = 0
'''                Else
'''                    RstDia("imphabsol") = NulosN(TxtIGV2.Text) * NulosN(TxtTC.Text)
'''                    RstDia("imphabdol") = NulosN(TxtIGV2.Text)
'''                End If
'''            End If
'''            RstDia.Update
'''
'''        End If
'''
'''        'grabamos el impuesto si la operacion sin derecho a credito fiscal
'''        If NulosN(TxtIGV3.Text) <> 0 Then
'''            xIdCuenTasa = fCtaImpuestoTipoCompra(3)
'''            If xIdCuenTasa = 0 Then GoTo LaCague
'''            '-------------------------------------
'''            RstDia.AddNew
'''            RstDia("año") = AnoTra
'''            RstDia("idmes") = mMesActivo
'''            If mMesActivo = 0 Then
'''                RstDia("idlib") = 36
'''            Else
'''                RstDia("idlib") = 1
'''            End If
'''            RstDia("idmov") = xId
'''            RstDia("numasi") = xNumAsiento
'''            RstDia("tc") = ValTipCam
'''            RstDia("idcue") = xIdCuenTasa
'''            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''
'''            'si el tipo de l proveedor es diferente a no domiciliado
'''            If NulosN(LblIdTipPer.Caption) <> 3 Then
'''                If NulosN(TxtTipDoc.Text) <> 0 Then
'''                    If NulosN(TxtTipDoc.Text) <> 7 And NulosN(TxtTipDoc.Text) <> 2 Then
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("impdebsol") = NulosN(TxtIGV3.Text)
'''                            RstDia("impdebdol") = 0
'''                        Else
'''                            RstDia("impdebsol") = NulosN(TxtIGV3.Text) * NulosN(TxtTC.Text)
'''                            RstDia("impdebdol") = NulosN(TxtIGV3.Text)
'''                        End If
'''                    Else
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("imphabsol") = NulosN(TxtIGV3.Text)
'''                            RstDia("imphabdol") = 0
'''                        Else
'''                            RstDia("imphabsol") = NulosN(TxtIGV3.Text) * NulosN(TxtTC.Text)
'''                            RstDia("imphabdol") = NulosN(TxtIGV3.Text)
'''                        End If
'''                    End If
'''                End If
'''            Else
'''                If TxtIdMon.Text = "1" Then
'''                    RstDia("imphabsol") = NulosN(TxtIGV3.Text)
'''                    RstDia("imphabdol") = 0
'''                Else
'''                    RstDia("imphabsol") = NulosN(TxtIGV3.Text) * NulosN(TxtTC.Text)
'''                    RstDia("imphabdol") = NulosN(TxtIGV3.Text)
'''                End If
'''            End If
'''            RstDia.Update
'''
'''        End If
'''
'''        '***********************************************************************
'''
'''
'''        '***********************************************************************
'''
'''        'grabamos el impuesto si la operacion a sujeto a no domiciliado
'''        If NulosN(TxtOtros.Text) <> 0 And NulosN(TxtTipDoc.Text) = 107 Then
'''            RstDia.AddNew
'''            RstDia("año") = AnoTra
'''            RstDia("idmes") = mMesActivo
'''            RstDia("idlib") = 1
'''            RstDia("idmov") = xId
'''            RstDia("numasi") = xNumAsiento
'''            RstDia("tc") = ValTipCam
'''            RstDia("idcue") = xIdCuenTasa
'''            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''            If NulosN(TxtTipDoc.Text) <> 0 Then
'''                If NulosN(TxtTipDoc.Text) <> 7 Then
'''                    'cuando se factura u otro comprabante excepto nota de credito hace su asiento norma
'''                    If TxtIdMon.Text = "1" Then
'''                        RstDia("imphabsol") = NulosN(TxtOtros.Text)
'''                        RstDia("imphabdol") = 0
'''                    Else
'''                        RstDia("imphabsol") = NulosN(TxtOtros.Text) * NulosN(TxtTC.Text)
'''                        RstDia("imphabdol") = NulosN(TxtOtros.Text)
'''                    End If
'''                Else
'''                    'cuando sea nota de credito hace el asiento inverso al de una venta
'''                    If TxtIdMon.Text = "1" Then
'''                        RstDia("impdebsol") = NulosN(TxtOtros.Text)
'''                        RstDia("impdebdol") = 0
'''                    Else
'''                        RstDia("impdebsol") = NulosN(TxtOtros.Text) * NulosN(TxtTC.Text)
'''                        RstDia("impdebdol") = NulosN(TxtOtros.Text)
'''                    End If
'''                End If
'''            End If
'''            RstDia.Update
'''        End If
'''
'''    '***********************************************************************
'''
'''
'''
'''
'''        'Dim Rst As New ADODB.Recordset
'''        'grabamos el imponible en function a los items de la factura
'''        Set Rst = Nothing
'''        RST_Busq Rst, "SELECT com_comprasdet.idcom, alm_inventario.idcuenta, Sum(com_comprasdet.imptot) AS SumaDeimptot FROM alm_inventario INNER JOIN com_comprasdet " _
'''            & " ON alm_inventario.id = com_comprasdet.iditem GROUP BY com_comprasdet.idcom, alm_inventario.idcuenta HAVING (((com_comprasdet.idcom)=" & xId & "))", xCon
'''
'''        If Rst.RecordCount <> 0 Then
'''            Rst.MoveFirst
'''            For A = 1 To Rst.RecordCount
'''                RstDia.AddNew
'''                RstDia("año") = AnoTra
'''                RstDia("idmes") = mMesActivo               'LLAVE - CODIGO DEL MES
'''                If mMesActivo = 0 Then
'''                    RstDia("idlib") = 36                 'LLAVE - CODIGO DEL LIBRO
'''                Else
'''                    RstDia("idlib") = 1
'''                End If
'''                RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
'''                RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
'''                RstDia("tc") = ValTipCam
'''                RstDia("idcue") = NulosN(Rst("idcuenta")) 'xIdCuen
'''                RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''                RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''                If NulosN(TxtTipDoc.Text) <> 0 Then
'''                    If NulosN(TxtTipDoc.Text) <> 7 Then
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("impdebsol") = NulosN(Rst("SumaDeimptot"))
'''                            RstDia("impdebdol") = 0
'''                        Else
'''                            RstDia("impdebsol") = NulosN(Rst("SumaDeimptot")) * NulosN(TxtTC.Text)
'''                            RstDia("impdebdol") = NulosN(Rst("SumaDeimptot"))
'''                        End If
'''                    Else
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("imphabsol") = NulosN(Rst("SumaDeimptot"))
'''                            RstDia("imphabdol") = 0
'''                        Else
'''                            RstDia("imphabsol") = NulosN(Rst("SumaDeimptot")) * NulosN(TxtTC.Text)
'''                            RstDia("imphabdol") = NulosN(Rst("SumaDeimptot"))
'''                        End If
'''                    End If
'''                End If
'''                RstDia.Update
'''
'''                Rst.MoveNext
'''                If Rst.EOF = True Then Exit For
'''            Next A
'''        End If
'''
'''        '***************************************************************************************************************
'''        'grabamos el selectivo en funcion a los items de la factura
'''        Set Rst = Nothing
'''
'''        RST_Busq Rst, "SELECT com_comprasdet.idcom, mae_impuestos.idcuen, Sum([com_comprasdet].[imptot]*([mae_impuestos].[tasa]/100)) AS total " _
'''            & " FROM (alm_inventario INNER JOIN mae_impuestos ON alm_inventario.idimpsel = mae_impuestos.id) INNER JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem " _
'''            & " WHERE (((com_comprasdet.idcom)=" & xId & ") AND ((com_comprasdet.idtipcom)<>4)) " _
'''            & " GROUP BY com_comprasdet.idcom, mae_impuestos.idcuen, alm_inventario.idimpsel ", xCon
'''
'''        If Rst.RecordCount <> 0 Then
'''            Rst.MoveFirst
'''
'''            Do While Not Rst.EOF
'''                RstDia.AddNew
'''                RstDia("año") = AnoTra
'''                RstDia("idmes") = mMesActivo               'LLAVE - CODIGO DEL MES
'''                RstDia("idlib") = 1                  'LLAVE - CODIGO DEL LIBRO
'''                RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
'''                RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
'''                RstDia("tc") = ValTipCam
'''                RstDia("idcue") = NulosN(Rst("idcuen"))
'''                RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''                RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''                If NulosN(TxtTipDoc.Text) <> 0 Then
'''                    If NulosN(TxtTipDoc.Text) <> 7 Then
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("impdebsol") = NulosN(Rst("total"))
'''                            RstDia("impdebdol") = 0
'''                        Else
'''                            RstDia("impdebsol") = NulosN(Rst("total")) * NulosN(TxtTC.Text)
'''                            RstDia("impdebdol") = NulosN(Rst("total"))
'''                        End If
'''                    Else
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("imphabsol") = NulosN(Rst("total"))
'''                            RstDia("imphabdol") = 0
'''                        Else
'''                            RstDia("imphabsol") = NulosN(Rst("total")) * NulosN(TxtTC.Text)
'''                            RstDia("imphabdol") = NulosN(Rst("total"))
'''                        End If
'''                    End If
'''                End If
'''                RstDia.Update
'''
'''                Rst.MoveNext
'''            Loop
'''        End If
'''        Set Rst = Nothing
'''
'''        '***************************************************************************************************************
'''        '--redondeo a centimos
'''        If NulosN(TxtRedondeo.Text) <> 0 Then
'''            Dim CtaRedondeo As Long
'''            CtaRedondeo = fCtaRedondeo()
'''            If CtaRedondeo = 0 Then GoTo LaCague
'''            '-----------------------------------
'''            RstDia.AddNew
'''            RstDia("año") = AnoTra
'''            RstDia("idmes") = mMesActivo
'''            If mMesActivo = 0 Then
'''                RstDia("idlib") = 36
'''            Else
'''                RstDia("idlib") = 1
'''            End If
'''            RstDia("idmov") = xId
'''            RstDia("numasi") = xNumAsiento
'''            RstDia("tc") = ValTipCam
'''            RstDia("idcue") = CtaRedondeo
'''            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''
'''            '---si es perdida
'''            If NulosN(TxtRedondeo.Text) > 0 Then
'''                If NulosN(TxtTipDoc.Text) <> 0 Then
'''                    If NulosN(TxtTipDoc.Text) <> 7 Then
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("impdebsol") = NulosN(TxtRedondeo.Text)
'''                            RstDia("impdebdol") = 0
'''                        Else
'''                            RstDia("impdebsol") = NulosN(TxtRedondeo.Text) * NulosN(TxtTC.Text)
'''                            RstDia("impdebdol") = NulosN(TxtRedondeo.Text)
'''                        End If
'''                    Else
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("imphabsol") = NulosN(TxtRedondeo.Text)
'''                            RstDia("imphabdol") = 0
'''                        Else
'''                            RstDia("imphabsol") = NulosN(TxtRedondeo.Text) * NulosN(TxtTC.Text)
'''                            RstDia("imphabdol") = NulosN(TxtRedondeo.Text)
'''                        End If
'''                    End If
'''                End If
'''            Else
'''                If NulosN(TxtTipDoc.Text) <> 0 Then
'''                    If NulosN(TxtTipDoc.Text) = 7 Then
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("impdebsol") = Abs(NulosN(TxtRedondeo.Text))
'''                            RstDia("impdebdol") = 0
'''                        Else
'''                            RstDia("impdebsol") = Abs(NulosN(TxtRedondeo.Text)) * NulosN(TxtTC.Text)
'''                            RstDia("impdebdol") = Abs(NulosN(TxtRedondeo.Text))
'''                        End If
'''                    Else
'''                        If TxtIdMon.Text = "1" Then
'''                            RstDia("imphabsol") = Abs(NulosN(TxtRedondeo.Text))
'''                            RstDia("imphabdol") = 0
'''                        Else
'''                            RstDia("imphabsol") = Abs(NulosN(TxtRedondeo.Text)) * NulosN(TxtTC.Text)
'''                            RstDia("imphabdol") = Abs(NulosN(TxtRedondeo.Text))
'''                        End If
'''                    End If
'''                End If
'''            End If
'''            RstDia.Update
'''        End If
'''
'''
'''        '***************************************************************************************************************
'''        If NulosN(LbIdCuentaDeb.Caption) <> 0 And NulosN(LbIdCuentaHab.Caption) <> 0 Then
'''            '---------------------------------------
'''            'Grabamos si se reclasifica la cuenta
'''            '---------------------------------------
'''            'grabamos a facturas por pagar Plan de cuentas 42.1 o dependiendo del caso
'''            RstDia.AddNew
'''            RstDia("año") = AnoTra
'''            RstDia("idmes") = mMesActivo
'''            RstDia("idlib") = 1
'''            RstDia("idmov") = xId
'''            RstDia("numasi") = xNumAsiento
'''            RstDia("tc") = ValTipCam
'''            RstDia("idcue") = NulosN(LbIdCuentaDeb.Caption)
'''            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''
'''            If TxtIdMon.Text = "1" Then
'''                RstDia("impdebsol") = NulosN(txtTotal1.Text)
'''                RstDia("impdebdol") = 0
'''            Else
'''                RstDia("impdebsol") = NulosN(txtTotal1.Text) * NulosN(TxtTC.Text)
'''                RstDia("impdebdol") = NulosN(txtTotal1.Text)
'''            End If
'''            RstDia.Update
'''            '-----------------------------------------------------
'''            RstDia.AddNew
'''            RstDia("año") = AnoTra
'''            RstDia("idmes") = mMesActivo
'''            RstDia("idlib") = 1
'''            RstDia("idmov") = xId
'''            RstDia("numasi") = xNumAsiento
'''            RstDia("tc") = ValTipCam
'''            RstDia("idcue") = NulosN(LbIdCuentaHab.Caption)
'''            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''            RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''
'''            If TxtIdMon.Text = "1" Then
'''                RstDia("imphabsol") = NulosN(txtTotal1.Text)
'''                RstDia("imphabdol") = 0
'''            Else
'''                RstDia("imphabsol") = NulosN(txtTotal1.Text) * NulosN(TxtTC.Text)
'''                RstDia("imphabdol") = NulosN(txtTotal1.Text)
'''            End If
'''            RstDia.Update
'''
'''        End If
'''        '***************************************************************************************************************
'''
'''
'''        '---------------------------------
'''        'grabamos los asientos automaticos
'''        'grabamos la cuenta de destino debe
'''        Set Rst = Nothing
''''' consulta para generar asientos automaticos en resumen
'''        RST_Busq Rst, "SELECT com_comprasdet.idcom, con_planctas.ctadesdeb, Sum(com_comprasdet.imptot) AS SumaDeimptot FROM con_planctas RIGHT JOIN (alm_inventario " _
'''            & " INNER JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem) ON con_planctas.id = alm_inventario.idcuenta GROUP BY com_comprasdet.idcom, " _
'''            & " con_planctas.ctadesdeb HAVING (((com_comprasdet.idcom)=" & xId & "))", xCon
''''' consulta para generar asientos automaticos en detalle
'''''        RST_Busq Rst, "SELECT com_comprasdet.idcom, con_planctas.ctadesdeb, com_comprasdet.imptot AS SumaDeimptot FROM con_planctas RIGHT JOIN (alm_inventario " _
'''''            & " INNER JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem) ON con_planctas.id = alm_inventario.idcuenta " _
'''''            & " WHERE (((com_comprasdet.idcom)=" & xId & "))", xCon
'''
'''        If Rst.RecordCount <> 0 Then
'''            Rst.MoveFirst
'''            For A = 1 To Rst.RecordCount
'''                If Rst("ctadesdeb") <> 0 Then
'''                    RstDia.AddNew
'''                    RstDia("año") = AnoTra
'''                    RstDia("idmes") = mMesActivo               'LLAVE - CODIGO DEL MES
'''                    RstDia("idlib") = 1                  'LLAVE - CODIGO DEL LIBRO
'''                    RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
'''                    RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
'''                    RstDia("tc") = ValTipCam
'''                    RstDia("idcue") = NulosN(Rst("ctadesdeb")) 'xIdCuen
'''                    RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''                    RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''                    If NulosN(TxtTipDoc.Text) <> 0 Then
'''                        If NulosN(TxtTipDoc.Text) <> 7 Then
'''                            If TxtIdMon.Text = "1" Then
'''                                RstDia("impdebsol") = NulosN(Rst("SumaDeimptot"))
'''                                RstDia("impdebdol") = 0
'''                            Else
'''                                RstDia("impdebsol") = NulosN(Rst("SumaDeimptot")) * NulosN(TxtTC.Text)
'''                                RstDia("impdebdol") = NulosN(Rst("SumaDeimptot"))
'''                            End If
'''                        Else
'''                            If TxtIdMon.Text = "1" Then
'''                                RstDia("imphabsol") = NulosN(Rst("SumaDeimptot"))
'''                                RstDia("imphabdol") = 0
'''                            Else
'''                                RstDia("imphabsol") = NulosN(Rst("SumaDeimptot")) * NulosN(TxtTC.Text)
'''                                RstDia("imphabdol") = NulosN(Rst("SumaDeimptot"))
'''                            End If
'''                        End If
'''                    End If
'''                    RstDia.Update
'''                End If
'''
'''                Rst.MoveNext
'''                If Rst.EOF = True Then Exit For
'''            Next A
'''        End If
'''
'''        'grabamos la cuenta de destino haber
'''        Set Rst = Nothing
''''''--destino automatico en resumen
'''        RST_Busq Rst, "SELECT com_comprasdet.idcom, con_planctas.ctadeshab, Sum(com_comprasdet.imptot) AS SumaDeimptot FROM con_planctas RIGHT JOIN (alm_inventario " _
'''            & " INNER JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem) ON con_planctas.id = alm_inventario.idcuenta GROUP BY com_comprasdet.idcom, " _
'''            & " con_planctas.ctadeshab HAVING (((com_comprasdet.idcom)=" & xId & "))", xCon
''''''--destino automatico en detalle
'''''        RST_Busq Rst, "SELECT com_comprasdet.idcom, con_planctas.ctadeshab, com_comprasdet.imptot AS SumaDeimptot FROM con_planctas RIGHT JOIN (alm_inventario " _
'''''            & " INNER JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem) ON con_planctas.id = alm_inventario.idcuenta " _
'''''            & " WHERE (((com_comprasdet.idcom)=" & xId & "))", xCon
'''
'''        If Rst.RecordCount <> 0 Then
'''            Rst.MoveFirst
'''            For A = 1 To Rst.RecordCount
'''                If Rst("ctadeshab") <> 0 Then
'''                    RstDia.AddNew
'''                    RstDia("año") = AnoTra
'''                    RstDia("idmes") = mMesActivo               'LLAVE - CODIGO DEL MES
'''                    RstDia("idlib") = 1                  'LLAVE - CODIGO DEL LIBRO
'''                    RstDia("idmov") = xId                'LLAVE - CODIGO DEL MOVIMIENTO
'''                    RstDia("numasi") = xNumAsiento       'LLAVE - NUMERO DE ASIENTO
'''                    RstDia("tc") = ValTipCam
'''                    RstDia("idcue") = NulosN(Rst("ctadeshab")) 'xIdCuen
'''                    RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''                    RstDia("fchdoc") = CDate(TxtFchDoc.Valor)
'''                    If NulosN(TxtTipDoc.Text) <> 0 Then
'''                        If NulosN(TxtTipDoc.Text) <> 7 Then
'''                            If TxtIdMon.Text = "1" Then
'''                                RstDia("imphabsol") = NulosN(Rst("SumaDeimptot"))
'''                                RstDia("imphabdol") = 0
'''                            Else
'''                                RstDia("imphabsol") = NulosN(Rst("SumaDeimptot")) * NulosN(TxtTC.Text)
'''                                RstDia("imphabdol") = NulosN(Rst("SumaDeimptot"))
'''                            End If
'''                        Else
'''                            If TxtIdMon.Text = "1" Then
'''                                RstDia("impdebsol") = NulosN(Rst("SumaDeimptot"))
'''                                RstDia("impdebdol") = 0
'''                            Else
'''                                RstDia("impdebsol") = NulosN(Rst("SumaDeimptot")) * NulosN(TxtTC.Text)
'''                                RstDia("impdebdol") = NulosN(Rst("SumaDeimptot"))
'''                            End If
'''                        End If
'''                    End If
'''                    RstDia.Update
'''                End If
'''
'''                Rst.MoveNext
'''                If Rst.EOF = True Then Exit For
'''            Next A
'''        End If
'''
'''
'''    '**************************************************************************************************************
'''    RST_Busq Rst, "SELECT mae_detraccion.id, mae_detraccion.descripcion, mae_detraccion.tasa, alm_inventario.iddet " _
'''        & " FROM alm_inventario LEFT JOIN mae_detraccion ON alm_inventario.iddet = mae_detraccion.id " _
'''        & " WHERE ((alm_inventario.id= " & NulosN(Fg1.TextMatrix(Fg1.Row, 11)) & "))", xCon
'''
'''    If Rst.RecordCount <> 0 Then
'''        If Rst("iddet") <> 0 Then
'''            MsgBox "Se ha detectado que la compra registrada esta afecta al regimen de la Detraccion " + Chr(13) _
'''                & "Decripcion : " + Rst("descripcion") + Chr(13) _
'''                & "tasa : " + Format(Rst("tasa"), "0.00") + "%", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''
'''            Dim RstDeta As New ADODB.Recordset
'''            Dim xId2 As Integer
'''
'''            If QueHace = 1 Then
'''                xId2 = HallaCodigoTabla("con_detraccion", xCon, "id")
'''                RST_Busq RstDeta, "SELECT TOP 1 * FROM con_detraccion", xCon
'''                RstDeta.AddNew
'''                RstDeta("id") = xId2
'''            Else
'''                RST_Busq RstDeta, "SELECT con_detraccion.* From con_detraccion " _
'''                    & " WHERE (((con_detraccion.iddoc)=" & xId & "))", xCon
'''            End If
'''
'''            If RstDeta.RecordCount = 0 Then
'''                'este procedimiento es solo para cuando se este modificando una compra afecta a la detraccion y no se le haya hecho la detraccion a la hora de ingresar la compra
'''                xId2 = HallaCodigoTabla("con_detraccion", xCon, "id")
'''                RstDeta.AddNew
'''                RstDeta("id") = xId2
'''            End If
'''
'''            RstDeta("iddet") = NulosN(Rst("iddet"))
'''            RstDeta("por") = NulosN(Rst("tasa"))
'''            RstDeta("iddoc") = xId
'''            RstDeta("idmon") = NulosN(TxtIdMon.Text)
'''            RstDeta("tipo") = 1
'''            RstDeta("fchmov") = Date
'''            RstDeta("Glosa") = ""
'''            RstDeta("imp") = Format((NulosN(TxtTotal.Text) * (Rst("tasa") / 100)), "0.00")
'''            RstDeta("numdet") = "SIN NUMERO"
'''            RstDeta.Update
'''        End If
'''    End If
'''
'''    '-----------------------------------------------------------------------------------------------------------
'''    '--grabar datos adicionales en el diario
'''    nSQL = "UPDATE ((com_compras INNER JOIN con_diario ON com_compras.id = con_diario.idmov) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha  " _
'''        + vbCr + " SET con_diario.fchdoc=com_compras.fchdoc, con_diario.idmon=com_compras.idmon, con_diario.ridlib = 1, con_diario.ridtipper = 1, con_diario.ridper = [com_compras].[idpro], con_diario.rtipdoc = [com_compras].[tipdoc], con_diario.rfchope = [com_compras].[fchdoc], con_diario.rnumerodoc = IIf([com_compras].[numser] Is Null Or [com_compras].[numser]='','',[com_compras].[numser] & '-') & [com_compras].[numdoc], con_diario.rglosaope = [com_compras].[glosa], con_diario.rregistro = Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4),con_diario.tc = IIf([com_compras].[tc]=0,[con_tc].[impven],[com_compras].[tc]), con_diario.aplicatc = IIf([com_compras].[tc]=0,0,-1) " _
'''        + vbCr + " WHERE (((con_diario.idlib)=1) AND ((con_diario.idmov)=" & xId & ")); "
'''
'''    xCon.Execute nSQL
'''
'''    'grabamos el movimiento en la tabla var_edicion
'''    GrabarOperacion xIdUsuario, 1, QueHace, xHorIni, Time, Date, xCon, CDbl(xId)
'''
'''    '-----------------------------------------------------------------------------------------------------------
'''
'''    xCon.CommitTrans
'''    Me.MousePointer = vbDefault
'''    MsgBox "La compra se registró con éxito" & vbCr & "Registro Nº: " & Format(mMesActivo, "00") & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''
'''    Set RstDeta = Nothing
'''    Set RstCab = Nothing
'''    Set RstDet = Nothing
'''    Set RstDia = Nothing
'''    Set RstCosto = Nothing
'''    Set RstReclasifica = Nothing
'''    Grabar = True
'''
'''    Exit Function
'''
'''LaCague:
''''    Resume
'''    xCon.RollbackTrans
'''    Me.MousePointer = vbDefault
'''    Set RstDeta = Nothing
'''    Set RstCab = Nothing
'''    Set RstDet = Nothing
'''    Set RstDia = Nothing
'''    Set RstCosto = Nothing
'''    Set RstReclasifica = Nothing
'''    If Err.Number <> 0 Then MsgBox "No se pudo guardar el registro por el siguiente motivo :" & vbCr & Trim(Err.Description)
'''    Err.Clear
End Function

Function HallaNumAsiento(Mes As Integer) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_diario.idmes, con_diario.idlib, con_diario.numasi From con_diario " _
        & " WHERE (((con_diario.idmes)=" & Mes & ") AND ((con_diario.idlib)=1)) ORDER BY numasi", xCon
    
    If Rst.RecordCount = 0 Then
        HallaNumAsiento = "0001"
    Else
        Rst.MoveLast
        HallaNumAsiento = Format(NulosN(Rst("numasi")) + 1, "0000")
    End If
    Exit Function
End Function

'Sub MuestraOrden(IdOrdenCompra As Integer)
'    Dim RstOrd As New ADODB.Recordset
'    Dim RstDet As New ADODB.Recordset
'    Dim A As Integer
'
'    RST_Busq RstOrd, "SELECT mae_prov.nombre AS nomprov, mae_prov.numruc, mae_tipoproducto.descripcion AS desctipcom, mae_estadoordcom.descripcion AS descestado, " _
'        & " mae_moneda.simbolo AS moneda, mae_condpago.descripcion AS descconpag, mae_documento.descripcion AS desctipdoc, mae_moneda.descripcion AS descmon, " _
'        & " com_ordencompra.* FROM  mae_estadoordcom RIGHT JOIN (mae_prov RIGHT JOIN (mae_tipoproducto RIGHT JOIN (mae_moneda RIGHT JOIN ((com_ordencompra " _
'        & " LEFT JOIN mae_condpago ON com_ordencompra.idconpag = mae_condpago.id) LEFT JOIN mae_documento ON com_ordencompra.idtipdoc = mae_documento.id) " _
'        & " ON mae_moneda.id = com_ordencompra.idmon) ON mae_tipoproducto.id = com_ordencompra.idtippro) ON mae_prov.id = com_ordencompra.idpro) " _
'        & " ON mae_estadoordcom.id = com_ordencompra.idest Where (((com_ordencompra.id) = " & IdOrdenCompra & ")) ORDER BY com_ordencompra.fchemi DESC", xCon
'
'    If RstOrd.RecordCount <> 0 Then
'        TxtNumOrdCom.Text = Format(RstOrd("id"), "000000")
'        TxtTipCom.Text = RstOrd("idtippro")
'        TxtIdMon.Text = RstOrd("idmon")
'        TxtTipDoc.Text = RstOrd("idtipdoc")
'        TxtNumRuc.Text = RstOrd("numruc")
'        TxtConPag.Text = RstOrd("idconpag")
'
'        LblTipoCompra.Caption = RstOrd("desctipcom")
'        LblMoneda.Caption = RstOrd("descmon")
'        LblNomDoc.Caption = RstOrd("desctipdoc")
'        LblNomPro.Caption = RstOrd("nomprov")
'        LblCondPag.Caption = RstOrd("descconpag")
'
'        LblIdProveedor.Caption = RstOrd("idpro")
'    End If
'
'    TasaImpuesto = HallaDatosImpuestoDocumento(Val(TxtTipDoc.Text), "tasa")
'    xIdCuenTasa = HallaDatosImpuestoDocumento(Val(TxtTipDoc.Text), "cuentaimp") 'NulosN(xRs("cuentaimp"))
'    xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
'    LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) + "%"
'    Set RstOrd = Nothing
'
'    'mostramos el detalle de la orden de compra
'    RST_Busq RstDet, "SELECT alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuenta, " _
'        & " alm_inventario.idtipcom, com_ordencompradet.*, con_planctas.ctadesdeb, con_planctas.ctadeshab " _
'        & " FROM con_planctas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN com_ordencompradet " _
'        & " ON alm_inventario.id = com_ordencompradet.iditem) ON mae_unidades.id = com_ordencompradet.idunimed) " _
'        & " ON con_planctas.id = alm_inventario.idcuenta WHERE (((com_ordencompradet.idcom)=" & Val(TxtNumOrdCom.Text) & "))", xCon
'
'    Mostrando = True
'    Fg1.Rows = 1
'    If RstDet.RecordCount <> 0 Then
'        RstDet.MoveFirst
'        For A = 1 To RstDet.RecordCount
'            Fg1.Rows = Fg1.Rows + 1
'            Fg1.TextMatrix(A, 1) = RstDet("descripcion")
'            Fg1.TextMatrix(A, 2) = RstDet("abrev")
'            Fg1.TextMatrix(A, 3) = Format(RstDet("preuni"), "0.000000")
'            Fg1.TextMatrix(A, 4) = Format(RstDet("canpro"), "0.00")
'            Fg1.TextMatrix(A, 5) = Format(RstDet("imptot"), "0.00")
'            Fg1.TextMatrix(A, 6) = RstDet("iditem")
'            Fg1.TextMatrix(A, 7) = RstDet("idunimed")
'            Fg1.TextMatrix(A, 8) = RstDet("idcuenta")
'            Fg1.TextMatrix(A, 9) = RstDet("idtipcom")
'            Fg1.TextMatrix(A, 10) = NulosN(RstDet("ctadesdeb"))
'            Fg1.TextMatrix(A, 11) = NulosN(RstDet("ctadeshab"))
'
'            RstDet.MoveNext
'            If RstDet.EOF = True Then
'                Exit For
'            End If
'        Next A
'    End If
'    BuscarImpuestos
'    HallarTotal
'
'    Mostrando = False
'End Sub

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtTipDoc.Text) = "" Then
        LblNomDoc.Caption = ""
        Exit Sub
    End If
    Dim xRs As New ADODB.Recordset
    
    RST_Busq xRs, "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuen as cuentaimp " _
        & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id WHERE mae_documento.id  = " & NulosN(TxtTipDoc.Text) & "", xCon
    
    If xRs.RecordCount = 0 Then
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
    Else
        TxtTipDoc.Text = xRs("id")
        LblNomDoc.Caption = xRs("descripcion")
        
    End If
    
    Set xRs = Nothing
    
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

    RST_Busq xRsDoc, "SELECT com_compras.fchdoc, mae_prov.nombre, mae_prov.numdoc, com_compras.imptot, com_compras.tipdoc, com_compras.idmon, " _
        & " mae_prov.dir FROM mae_prov RIGHT JOIN com_compras " _
        & " ON mae_prov.id = com_compras.idpro Where (((com_compras.id) = " & RstComp("id") & "))", xCon
    
    RST_Busq xRsDet, "SELECT com_comprasdet.idcom, alm_inventario.descripcion, mae_unidades.abrev, com_comprasdet.canpro, com_comprasdet.preuni, " _
        & " com_comprasdet.imptot FROM mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem) " _
        & " ON mae_unidades.id = com_comprasdet.idunimed WHERE (((com_comprasdet.idcom)=" & RstComp("id") & "))", xCon

    RST_Busq RsPDoc, "SELECT * FROM var_plantilladoc WHERE tipdoc = " & xRsDoc("tipdoc") & " ", xCon

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
                Printer.Print "Son : "; NumeroLetra(xRsDoc("imptot"), xRsDoc("idmon"))
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


'*******************************

Sub Buscar()
    TabOne1.CurrTab = 0
    Dim xRs As New ADODB.Recordset
    
    Dim nSQL As String
    Dim xCampos(8, 4) As String
    
    xCampos(0, 0) = "N°Reg":        xCampos(0, 1) = "numreg":     xCampos(0, 2) = "820":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":      xCampos(1, 2) = "400":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "N°. Documento": xCampos(2, 1) = "numerodoc":  xCampos(2, 2) = "1400":  xCampos(2, 3) = "C"
    xCampos(3, 0) = "FchEmi":       xCampos(3, 1) = "fchdoc":     xCampos(3, 2) = "830":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "FchVenc":      xCampos(4, 1) = "fchven":     xCampos(4, 2) = "830":   xCampos(4, 3) = "C"
    xCampos(5, 0) = "Proveedor":     xCampos(5, 1) = "nombre":     xCampos(5, 2) = "2600":  xCampos(5, 3) = "C"
    
    xCampos(6, 0) = "M":             xCampos(6, 1) = "simbolo":    xCampos(6, 2) = "450":    xCampos(6, 3) = "C"
    xCampos(7, 0) = "Importe":         xCampos(7, 1) = "imptot":     xCampos(7, 2) = "850":    xCampos(7, 3) = "N"
    
    nSQL = "SELECT com_compras.id,Mid([com_compras].[numreg],1,2)+[mae_libros].[codsun]+Mid([com_compras].[numreg],3,4) AS numreg, mae_prov.nombre, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numerodoc, mae_documento.abrev, format(com_compras.fchdoc,'dd/mm/yy') as fchdoc, format(com_compras.fchven,'dd/mm/yy') as fchven, mae_prov.numruc, mae_moneda.simbolo, com_compras.imptot, com_compras.impsal " _
        + vbCr + " FROM (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id " _
        + vbCr + " WHERE ((month(com_compras.numreg) =" & mMesActivo & ")) " _
        + vbCr + " ORDER BY com_compras.numreg DESC;"


    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Compras", "nombre", "nombre", Principio

    If xRs.State = 1 Then
        RstComp.MoveFirst
        RstComp.Find "id = " & xRs("id") & ""
    End If
    
    Set xRs = Nothing
End Sub

Private Sub TxtTipDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDocRef_Click
    End If
End Sub

Private Sub TxtTipDocRef_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosN(TxtTipDocRef.Text) = 0 Then
        LblTipDocref.Caption = ""
        TxtDocRef2.Text = ""
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    
    RST_Busq xRs1, "SELECT * FROM mae_docreferencia WHERE id = " & NulosN(TxtTipDocRef.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtTipDocRef.Text = ""
        LblTipDocref.Caption = ""
    Else
        LblTipDocref.Caption = NulosC(xRs1("descripcion"))
        TxtDocRef2.Text = ""
        LblIdDocRef2.Caption = ""
    End If
    Set xRs1 = Nothing
End Sub

Sub ActualizaSaldoDoc(idDocumento As Double, Tabla As Integer, ImporteRestar As Double)
    '1 = compras
    '2 = Ventas
    '3 = honorarios
    
    Dim Rst As New ADODB.Recordset
    Dim Total As Double
    
    If Tabla = 1 Then
        RST_Busq Rst, "SELECT Sum(tes_cajadestinodet.acuenta) AS total FROM tes_caja LEFT JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes " _
            & " GROUP BY tes_cajadestinodet.iddoc, tes_caja.tipmov HAVING tes_cajadestinodet.idmod=1 and tes_cajadestinodet.idori <>2 (((tes_cajadestinodet.iddoc)=" & idDocumento & ") AND ((tes_caja.tipmov)=2))", xCon
            
        Total = BuscaImporteDocumento(idDocumento, 1)
        
    End If
    
    
    If Rst.RecordCount <> 0 Then
        Total = ((Total - NulosN(Rst("total"))) - ImporteRestar)
    Else
        Total = (Total - ImporteRestar)
    End If
    
    xCon.Execute "UPDATE com_compras SET com_compras.impsal = " & Total & " WHERE (((com_compras.id)=" & idDocumento & "))"
    Set Rst = Nothing
    
End Sub

Function BuscaImporteDocumento(idDocumento As Double, Tabla As Integer) As Double
    '1 = compras
    '2 = Ventas
    '3 = honorarios
    Dim Rst As New ADODB.Recordset
    
    'compras
    If Tabla = 1 Then RST_Busq Rst, "SELECT * FROM com_compras WHERE id = " & idDocumento & "", xCon
    
    If Rst.RecordCount <> 0 Then
        BuscaImporteDocumento = NulosN(Rst("imptot"))
    Else
        BuscaImporteDocumento = 0
    End If
    
    Set Rst = Nothing
End Function

Private Sub pGridConfigurar()
    
        Fg1.ColWidth(2) = 3855
        Fg1.ColWidth(3) = 0
        Fg1.ColWidth(4) = 0
        Fg1.ColWidth(6) = 1100
        Fg1.ColWidth(7) = 1100
        Fg1.ColWidth(10) = 1300
        If Fg1.Rows > 1 Then Fg1.TextMatrix(Fg1.Rows - 1, 4) = 1

End Sub

'*********************

Private Sub CmdBusCtaDeb_Click()
    If xDeDonde = 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    'Dim xform As New EPS_Buscar.Buscar
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
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
        TxtCtaDeb.Text = xRs("cuenta")
        LblNomCtaDeb.Caption = xRs("descripcion")
        LbIdCuentaDeb.Caption = xRs("id")
        TxtCtaHab.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub



Private Sub CmdBusCtaHab_Click()
    If xDeDonde = 2 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    'Dim xform As New EPS_Buscar.Buscar
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
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
        TxtCtaHab.Text = xRs("cuenta")
        LblNomCtaHab.Caption = xRs("descripcion")
        LbIdCuentaHab.Caption = xRs("id")
        CmdClAceptar.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub



Private Sub TxtCtaDeb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtCtaDeb_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 32 Then
        If QueHace = 3 Then Exit Sub
        TxtCtaDeb.Text = ""
        LblNomCtaDeb.Caption = ""
        LbIdCuentaDeb.Caption = ""
    End If
    
    If KeyCode = 116 Then
        CmdBusCtaDeb_Click
    End If
End Sub

Private Sub TxtCtaDeb_Validate(Cancel As Boolean)
    If Trim(TxtCtaDeb.Text) = "" Then
        LblNomCtaDeb.Caption = ""
        LbIdCuentaDeb.Caption = ""
        Exit Sub
    End If
    LblNomCtaDeb.Caption = Busca_Codigo(NulosC(TxtCtaDeb.Text), "cuenta", "descripcion", "con_planctas", "C", xCon)
    If LblNomCtaDeb.Caption <> "" Then
        LbIdCuentaDeb.Caption = Busca_Codigo(NulosC(TxtCtaDeb.Text), "cuenta", "id", "con_planctas", "C", xCon)
    Else
        LbIdCuentaDeb.Caption = ""
    End If
End Sub

Private Sub TxtCtaHab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtCtaHab_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 32 Then
        If QueHace = 3 Then Exit Sub
        TxtCtaHab.Text = ""
        LblNomCtaHab.Caption = ""
        LbIdCuentaHab.Caption = ""
    End If
    
    If KeyCode = 116 Then
        CmdBusCtaHab_Click
    End If
End Sub


Private Sub TxtCtaHab_Validate(Cancel As Boolean)
    If Trim(TxtCtaHab.Text) = "" Then
        LblNomCtaHab.Caption = ""
        LbIdCuentaHab.Caption = ""
        Exit Sub
    End If
    LblNomCtaHab.Caption = Busca_Codigo(NulosC(TxtCtaHab.Text), "cuenta", "descripcion", "con_planctas", "C", xCon)
    If LblNomCtaHab.Caption <> "" Then
        LbIdCuentaHab.Caption = Busca_Codigo(NulosC(TxtCtaHab.Text), "cuenta", "id", "con_planctas", "C", xCon)
    Else
        LbIdCuentaHab.Caption = ""
    End If
    
    
End Sub




'*********************
Private Function fCtaImpuestoTipoCompra(Tipo As Integer) As Long
    '--buscar la cuenta contable del tipo de compra en mae_tipocompra
    Dim rstBusq As New ADODB.Recordset
    Dim nSQL As String
    
    nSQL = "SELECT mae_tipocompra.id, mae_tipocompra.descripcion, mae_tipocompra.idcuen FROM mae_tipocompra WHERE (((mae_tipocompra.id)=" & Tipo & "));"
    RST_Busq rstBusq, nSQL, xCon
    If rstBusq.RecordCount <> 0 Then
        If NulosN(rstBusq("idcuen")) = 0 Then
            MsgBox "El concepto " & NulosC(rstBusq("descripcion")) & " no tiene Cuenta Contable Asignada" & vbCr & "Asignele una Cuenta para continuar", vbExclamation, xTitulo
            fCtaImpuestoTipoCompra = 0
        Else
            fCtaImpuestoTipoCompra = NulosN(rstBusq("idcuen"))
        End If
    Else
        MsgBox "No existe el concepto con Codigo = " & Tipo & " en mae_tipocompra", vbExclamation, xTitulo
        fCtaImpuestoTipoCompra = 0
    End If
    
    Set rstBusq = Nothing
End Function


Private Function fCtaRedondeo() As Long
    '--buscar la cuenta contable del tipo de compra en mae_tipocompra
    Dim rstBusq As New ADODB.Recordset
    Dim nSQL As String
    
    nSQL = "SELECT mae_redondeo.* FROM mae_redondeo WHERE (((mae_redondeo.idmod)=1) AND ((mae_redondeo.idmon) = " & NulosN(TxtIdMon.Text) & "));"
    
    RST_Busq rstBusq, nSQL, xCon
    If rstBusq.RecordCount <> 0 Then
        If NulosN(TxtRedondeo.Text) < NulosN(rstBusq("impmin")) Or NulosN(TxtRedondeo.Text) > NulosN(rstBusq("impmax")) Then
            MsgBox "El importe de Redondeo a Céntimos no esta en el rango permitido", vbExclamation, xTitulo
            fCtaRedondeo = 0
        Else
            If NulosN(TxtRedondeo.Text) > 0 Then '--perdida
                fCtaRedondeo = NulosN(rstBusq("idcuenper"))
            Else '--ganancia
                fCtaRedondeo = NulosN(rstBusq("idcuengan"))
            End If
            '----
            If fCtaRedondeo = 0 Then
                MsgBox "Falta Configurar la Cuenta Contable para el Redondeo a Céntimos", vbExclamation, xTitulo
                fCtaRedondeo = 0
            End If
        End If
    Else
        MsgBox "No existe la Cuenta Contable para el Redondeo a Céntimos", vbExclamation, xTitulo
        fCtaRedondeo = 0
    End If
    
    Set rstBusq = Nothing
End Function





Private Sub ChkTC_Click()
    If QueHace = 3 Then Exit Sub
    If ChkTC.Value = 0 Then
        TxtTC.BackColor = &H8000000F
        TxtTC.Enabled = False
        If IsDate(TxtFchDoc.Valor) = True Then
            TxtTC.Text = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
        Else
            MsgBox "Falta especificar la Fecha de emision", vbInformation, xTitulo
            TxtFchDoc.SetFocus
            Exit Sub
        End If
    Else
        TxtTC.Enabled = True
        TxtTC.BackColor = vbWhite
        TxtTC.SetFocus
    End If
End Sub


Private Sub pExportar()
    
    TabOne1.CurrTab = 0

    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset

    Dim xCampos(17, 3) As String
    
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":           xCampos(0, 2) = 2:   xCampos(0, 3) = "500"
    xCampos(1, 0) = "Nº Reg":       xCampos(1, 1) = "numreg1":      xCampos(1, 2) = 0:   xCampos(1, 3) = "0"
    xCampos(2, 0) = "R.U.C.":       xCampos(2, 1) = "pronumruc":       xCampos(2, 2) = 0:   xCampos(2, 3) = "1200"
    xCampos(3, 0) = "Proveedor":    xCampos(3, 1) = "pronombre":       xCampos(3, 2) = 0:   xCampos(3, 3) = "3290"
    xCampos(4, 0) = "T.D.":         xCampos(4, 1) = "tdabrev":        xCampos(4, 2) = 0:   xCampos(4, 3) = "350"
    xCampos(5, 0) = "Num. Doc":     xCampos(5, 1) = "numerodoc":    xCampos(5, 2) = 0:   xCampos(5, 3) = "1600"
    xCampos(6, 0) = "Fch.Emi":      xCampos(6, 1) = "fchdoc1":      xCampos(6, 2) = 1:   xCampos(6, 3) = "900"
    xCampos(7, 0) = "Fch. Venc":    xCampos(7, 1) = "fchven":      xCampos(7, 2) = 1:   xCampos(7, 3) = "900"
    xCampos(8, 0) = "Glosa":        xCampos(8, 1) = "glosa":        xCampos(8, 2) = 0:   xCampos(8, 3) = "2000"
    xCampos(9, 0) = "M":            xCampos(9, 1) = "moneda":      xCampos(9, 2) = 1:   xCampos(9, 3) = "500"
    xCampos(10, 0) = "T.C.":        xCampos(10, 1) = "impven1":     xCampos(10, 2) = 2:  xCampos(10, 3) = "700"
'    xCampos(11, 0) = "Imp Bru1":    xCampos(11, 1) = "impbru":      xCampos(11, 2) = 2:  xCampos(11, 3) = "900"
'    xCampos(12, 0) = "Imp Bru2":    xCampos(12, 1) = "impbru2":     xCampos(12, 2) = 2:  xCampos(12, 3) = "900"
'    xCampos(13, 0) = "Imp Bru3":    xCampos(13, 1) = "impbru3":     xCampos(13, 2) = 2:  xCampos(13, 3) = "900"
    xCampos(11, 0) = "Imp Inaf":    xCampos(11, 1) = "impina":      xCampos(11, 2) = 2:  xCampos(11, 3) = "900"
'    xCampos(15, 0) = "Descuento":   xCampos(15, 1) = "impdesc":     xCampos(15, 2) = 2:  xCampos(15, 3) = "900"
'    xCampos(16, 0) = "Imp ISC":     xCampos(16, 1) = "impisc":      xCampos(16, 2) = 2:  xCampos(16, 3) = "900"
'    xCampos(17, 0) = "Imp Igv1":    xCampos(17, 1) = "impigv":      xCampos(17, 2) = 2:  xCampos(17, 3) = "900"
'    xCampos(18, 0) = "Imp Igv2":    xCampos(18, 1) = "impigv2":     xCampos(18, 2) = 2:  xCampos(18, 3) = "900"
'    xCampos(19, 0) = "Imp Igv3":    xCampos(19, 1) = "impigv3":     xCampos(19, 2) = 2:  xCampos(19, 3) = "900"
'    xCampos(20, 0) = "Imp Otros":   xCampos(20, 1) = "otroscargos": xCampos(20, 2) = 2:  xCampos(20, 3) = "900"
    xCampos(12, 0) = "Imp Total":   xCampos(12, 1) = "imptot":      xCampos(12, 2) = 2:  xCampos(12, 3) = "1000"
    xCampos(13, 0) = "Imp Saldo":   xCampos(13, 1) = "impsal":      xCampos(13, 2) = 2:  xCampos(13, 3) = "1000"
    
    
    xCampos(14, 0) = "Ruc Cliente": xCampos(14, 1) = "clinumruc":       xCampos(14, 2) = 0:  xCampos(14, 3) = "1200"
    xCampos(15, 0) = "Cliente":     xCampos(15, 1) = "clinombre":       xCampos(15, 2) = 0:  xCampos(15, 3) = "3000"
    xCampos(16, 0) = "Orden":       xCampos(16, 1) = "numerodocref":    xCampos(16, 2) = 0:  xCampos(16, 3) = "1600"
    xCampos(17, 0) = "Fch Orden":   xCampos(17, 1) = "fchorden":        xCampos(17, 2) = 1:  xCampos(17, 3) = "900"
    
    
    
    Set RstTmp = RstComp.Clone
    
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "LISTADO DE COMPRAS", "Periodo " & LblMes.Caption, "", "Listado de Compras - " & LblMes.Caption, RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
    
    
End Sub




'--------------------------------


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
                TxtNumRucCli.Text = xRs("numruc")
                LblNomCli.Caption = xRs("nombre")
                LblIdcliente.Caption = xRs("id")
                TxtNumSer.SetFocus
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub TxtNumRucCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumRucCli_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCli_Click
    End If
End Sub

Private Sub TxtNumRucCli_Validate(Cancel As Boolean)
    If NulosC(TxtNumRucCli.Text) = "" Then
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    RST_Busq xRs1, "SELECT * FROM mae_cliente WHERE numruc like '" & TxtNumRucCli.Text & "%' ORDER BY numruc", xCon
    If xRs1.RecordCount <> 0 Then
        TxtNumRucCli.Text = xRs1("numruc")
        LblNomCli.Caption = xRs1("nombre")
        LblIdcliente.Caption = xRs1("id")
    Else
        TxtNumRucCli.Text = ""
        LblNomCli.Caption = ""
        LblIdcliente.Caption = ""
    End If
    Set xRs1 = Nothing
End Sub





'---------------------------------
