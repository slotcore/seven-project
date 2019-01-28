VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultaProduccion2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unificado - Consultar Produccion"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmProgresoTot 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1875
      Left            =   4590
      TabIndex        =   47
      Top             =   3990
      Visible         =   0   'False
      Width           =   5655
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   800
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar3 
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1500
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   54
         Top             =   1290
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   53
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando"
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
         TabIndex        =   50
         Top             =   90
         Width           =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   4
         X1              =   0
         X2              =   5610
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   3
         X1              =   5640
         X2              =   5640
         Y1              =   15
         Y2              =   1830
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   20
         X2              =   20
         Y1              =   0
         Y2              =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   5
         X1              =   0
         X2              =   5610
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Productos Terminados ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   105
         TabIndex        =   52
         Top             =   400
         Width           =   2145
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   60
         Top             =   30
         Width           =   5550
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Productos Intermedios ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   105
         TabIndex        =   51
         Top             =   1110
         Width           =   2145
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "&H80000009&"
      Height          =   7110
      Left            =   10770
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   11835
      Begin VB.CommandButton CmdSalir 
         Height          =   555
         Left            =   10815
         Picture         =   "FrmConsultaProduccion2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   6350
         Width           =   735
      End
      Begin VB.CommandButton CmdPrin 
         Height          =   555
         Left            =   10050
         Picture         =   "FrmConsultaProduccion2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exportar MSExcel"
         Top             =   6350
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   990
         TabIndex        =   6
         ToolTipText     =   "Agrandar columnas"
         Top             =   6350
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   5
         ToolTipText     =   "Reducir columnas"
         Top             =   6350
         Visible         =   0   'False
         Width           =   735
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg7 
         Height          =   5820
         Left            =   60
         TabIndex        =   9
         Top             =   420
         Width           =   11640
         _cx             =   20532
         _cy             =   10266
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   128
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsultaProduccion2.frx":0E14
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
         BackStyle       =   0  'Transparent
         Caption         =   "Todas las Empresas"
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
         Left            =   210
         TabIndex        =   46
         Top             =   120
         Width           =   2595
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   11745
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   0
         X1              =   11760
         X2              =   11760
         Y1              =   0
         Y2              =   6990
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   6645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   11745
         Y1              =   7000
         Y2              =   7000
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   330
         Left            =   30
         Top             =   45
         Width           =   11700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consolidado de insumos"
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
         Left            =   105
         TabIndex        =   10
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame FrmProgreso 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1065
      Left            =   3300
      TabIndex        =   42
      Top             =   2640
      Visible         =   0   'False
      Width           =   5625
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   90
         TabIndex        =   43
         Top             =   750
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando Datos"
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
         Left            =   105
         TabIndex        =   44
         Top             =   75
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   5610
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   2
         X1              =   5610
         X2              =   5610
         Y1              =   15
         Y2              =   1050
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   5610
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   30
         Top             =   30
         Width           =   5550
      End
      Begin VB.Label LblProcesa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   105
         TabIndex        =   45
         Top             =   350
         Width           =   1575
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame15"
      Height          =   285
      Left            =   6060
      TabIndex        =   11
      Top             =   7110
      Width           =   5625
      Begin VB.Shape Shape4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   180
         Left            =   3195
         Top             =   45
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "= Faltante de Produccion"
         Height          =   195
         Left            =   3840
         TabIndex        =   13
         Top             =   45
         Width           =   1785
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   180
         Left            =   900
         Top             =   45
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "= Sobre Produccion"
         Height          =   195
         Left            =   1545
         TabIndex        =   12
         Top             =   45
         Width           =   1410
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8235
      Top             =   -30
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
            Picture         =   "FrmConsultaProduccion2.frx":0F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":1482
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":15DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":196E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":1AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":1F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":205E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":25A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":2AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":2BFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":2D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":3162
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion2.frx":32CE
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
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mostrar plan  de abastecimiento unificado"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7155
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   11895
      _cx             =   20981
      _cy             =   12621
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
      BackTabColor    =   -2147483637
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   8388608
      Caption         =   "Tab01  | Tab02 | Tab03| Tab04| Tab05"
      Align           =   0
      CurrTab         =   1
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
      Flags(2)        =   2
      Flags(3)        =   2
      Flags(4)        =   2
      Begin VB.Frame frmContenedor 
         BorderStyle     =   0  'None
         Caption         =   "frmContenedor"
         Height          =   6735
         Index           =   4
         Left            =   13140
         TabIndex        =   36
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   4
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
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
            Begin VB.Frame frmContTerm 
               BorderStyle     =   0  'None
               Caption         =   "frmContTerm"
               Height          =   6315
               Index           =   4
               Left            =   15
               TabIndex        =   40
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid FgTerm 
                  Height          =   6030
                  Index           =   4
                  Left            =   0
                  TabIndex        =   41
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Rows            =   50
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion2.frx":3816
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
            Begin VB.Frame frmContInter 
               BorderStyle     =   0  'None
               Caption         =   "frmContInter"
               Height          =   6315
               Index           =   4
               Left            =   12360
               TabIndex        =   38
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid FgInter 
                  Height          =   6030
                  Index           =   4
                  Left            =   0
                  TabIndex        =   39
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Rows            =   50
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion2.frx":396E
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
      End
      Begin VB.Frame frmContenedor 
         BorderStyle     =   0  'None
         Caption         =   "frmContenedor"
         Height          =   6735
         Index           =   3
         Left            =   12840
         TabIndex        =   30
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   3
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
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
            Begin VB.Frame frmContTerm 
               BorderStyle     =   0  'None
               Caption         =   "frmContTerm"
               Height          =   6315
               Index           =   3
               Left            =   15
               TabIndex        =   34
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid FgTerm 
                  Height          =   6030
                  Index           =   3
                  Left            =   0
                  TabIndex        =   35
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Rows            =   50
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion2.frx":3AE8
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
            Begin VB.Frame frmContInter 
               BorderStyle     =   0  'None
               Caption         =   "frmContInter"
               Height          =   6315
               Index           =   3
               Left            =   12360
               TabIndex        =   32
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid FgInter 
                  Height          =   6030
                  Index           =   3
                  Left            =   0
                  TabIndex        =   33
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Rows            =   50
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion2.frx":3C40
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
      End
      Begin VB.Frame frmContenedor 
         BorderStyle     =   0  'None
         Caption         =   "frmContenedor"
         Height          =   6735
         Index           =   2
         Left            =   12540
         TabIndex        =   24
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   2
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
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
            Begin VB.Frame frmContTerm 
               BorderStyle     =   0  'None
               Caption         =   "frmContTerm"
               Height          =   6315
               Index           =   2
               Left            =   15
               TabIndex        =   28
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid FgTerm 
                  Height          =   6030
                  Index           =   2
                  Left            =   0
                  TabIndex        =   29
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Rows            =   50
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion2.frx":3DBA
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
            Begin VB.Frame frmContInter 
               BorderStyle     =   0  'None
               Caption         =   "frmContInter"
               Height          =   6315
               Index           =   2
               Left            =   12360
               TabIndex        =   26
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid FgInter 
                  Height          =   6030
                  Index           =   2
                  Left            =   0
                  TabIndex        =   27
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Rows            =   50
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion2.frx":3F12
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
      End
      Begin VB.Frame frmContenedor 
         BorderStyle     =   0  'None
         Caption         =   "frmContenedor"
         Height          =   6735
         Index           =   1
         Left            =   45
         TabIndex        =   18
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   1
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
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
            Begin VB.Frame frmContTerm 
               BorderStyle     =   0  'None
               Caption         =   "frmContTerm"
               Height          =   6315
               Index           =   1
               Left            =   15
               TabIndex        =   22
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid FgTerm 
                  Height          =   6030
                  Index           =   1
                  Left            =   0
                  TabIndex        =   23
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Rows            =   2
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion2.frx":408C
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
            Begin VB.Frame frmContInter 
               BorderStyle     =   0  'None
               Caption         =   "frmContInter"
               Height          =   6315
               Index           =   1
               Left            =   12360
               TabIndex        =   20
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid FgInter 
                  Height          =   6030
                  Index           =   1
                  Left            =   0
                  TabIndex        =   21
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Rows            =   2
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion2.frx":41E4
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
      End
      Begin VB.Frame frmContenedor 
         BorderStyle     =   0  'None
         Caption         =   "frmContenedor"
         Height          =   6735
         Index           =   0
         Left            =   -12450
         TabIndex        =   2
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
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
            Begin VB.Frame frmContInter 
               BorderStyle     =   0  'None
               Caption         =   "frmContInter"
               Height          =   6315
               Index           =   0
               Left            =   15
               TabIndex        =   16
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid FgInter 
                  Height          =   6030
                  Index           =   0
                  Left            =   0
                  TabIndex        =   17
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Rows            =   2
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion2.frx":435C
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
            Begin VB.Frame frmContTerm 
               BorderStyle     =   0  'None
               Caption         =   "frmContTerm"
               Height          =   6315
               Index           =   0
               Left            =   -12330
               TabIndex        =   14
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid FgTerm 
                  Height          =   6030
                  Index           =   0
                  Left            =   0
                  TabIndex        =   15
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Rows            =   2
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion2.frx":44D6
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
      End
   End
End
Attribute VB_Name = "FrmConsultaProduccion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCONSULTAPRODUCCION
'* Tipo             : MODULO
'* Descripcion      : FORMULARIO QUE PERMITE CONSULTAR EL PLAN DE PRODUCCION ACTIVO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Enum Devolver
    Receta = 1
    Cantidad = 2
End Enum

Dim RstInsumos As New ADODB.Recordset      ' RECORDSET QUE ALMACENARA LOS INSUMOS
Dim SeEjecuto As Boolean                   ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLO VEZ
Dim xCon1 As New ADODB.Connection          ' CONECCION A LA BASE DE DATOS
Dim xCon2 As New ADODB.Connection          ' CONECCION A LA BASE DE DATOS
Dim cargoTerminado(5) As Boolean
Dim cargoIntermedio(5) As Boolean
Dim xRuta As String
Dim RstEmp As New ADODB.Recordset

Private Sub CmdPrin_Click()
    On Error GoTo error
    
    Dim oExport As New SGI2_funciones.Formularios
    
    Me.MousePointer = vbHourglass
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg7, "Unificado de Producción", "Periodo: " & AnoTra, "", "Unificado de Producción"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
    
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub

Private Sub CmdSalir_Click()
    Frame2.Visible = False
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
End Sub

Private Sub iniciarCampos()
    Dim A As Integer
        Dim xIndex As Integer
        Set xCon1 = AbrirConecciones(AP_RUTABD + "data.mdb")

        RST_Busq RstEmp, "SELECT mae_empresa.* From mae_empresa WHERE (((mae_empresa.anotra)=" & AnoTra & ") AND ((mae_empresa.activo)=-1))", xCon1

        If RstEmp.RecordCount <> 0 Then
            xIndex = 0
            RstEmp.MoveFirst
            'Se rellena los datos de las Empresas
            For A = 1 To RstEmp.RecordCount
                TabOne1.TabCaption(xIndex) = " " & Trim(RstEmp("abrevia")) & " "
                TabOne1.TabVisible(xIndex) = True
                cargoIntermedio(xIndex) = False
                cargoTerminado(xIndex) = False
                
                FgTerm(xIndex).AllowUserResizing = flexResizeColumns
                FgTerm(xIndex).AutoSearch = flexSearchFromTop
                FgTerm(xIndex).ExplorerBar = flexExSortShowAndMove
                FgTerm(xIndex).SelectionMode = flexSelectionByRow
                FgTerm(xIndex).ForeColorSel = &H80000005
                FgTerm(xIndex).BackColorSel = &H80&
                
                FgInter(xIndex).AllowUserResizing = flexResizeColumns
                FgInter(xIndex).AutoSearch = flexSearchFromTop
                FgInter(xIndex).ExplorerBar = flexExSortShowAndMove
                FgInter(xIndex).SelectionMode = flexSelectionByRow
                FgInter(xIndex).ForeColorSel = &H80000005
                FgInter(xIndex).BackColorSel = &H80&
                
                If RstEmp.EOF = True Then
                    Exit For
                End If
                xIndex = xIndex + 1
                RstEmp.MoveNext
            Next A
                RstEmp.MoveFirst
                xRuta = AP_RUTABD + Trim(RstEmp("ruta"))
                Set xCon2 = Nothing
                Set xCon2 = AbrirConecciones(xRuta)
                Me.Refresh
                CargaTerminados 0
        End If
        TabOne1.CurrTab = 0
        TabOne2(0).CurrTab = 0
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        iniciarCampos
    End If
End Sub

Private Sub configurarFG(tipo As Integer, Indice As Integer)
    Select Case tipo
        Case 0
            FgTerm(Indice).AllowUserResizing = flexResizeColumns
            FgTerm(Indice).AutoSearch = flexSearchFromTop
            FgTerm(Indice).ExplorerBar = flexExSortShowAndMove
            FgTerm(Indice).SelectionMode = flexSelectionByRow
            FgTerm(Indice).ForeColorSel = &H80000005
            FgTerm(Indice).BackColorSel = &H80&
            FgTerm(Indice).RowHeight(0) = 300
            FgTerm(Indice).ColWidth(0) = 0
            FgTerm(Indice).TextMatrix(0, 1) = "id"
            FgTerm(Indice).ColWidth(1) = 0
            FgTerm(Indice).TextMatrix(0, 2) = "idmae"
            FgTerm(Indice).ColWidth(2) = 0
            FgTerm(Indice).TextMatrix(0, 3) = "Producto"
            FgTerm(Indice).ColWidth(3) = 5500
            FgTerm(Indice).TextMatrix(0, 4) = "Unidad"
            FgTerm(Indice).TextMatrix(0, 5) = "Programado"
            FgTerm(Indice).TextMatrix(0, 6) = "Stock Ini."
            FgTerm(Indice).TextMatrix(0, 7) = "Producido"
            FgTerm(Indice).TextMatrix(0, 8) = "Total"
            FgTerm(Indice).TextMatrix(0, 9) = "Diferencia"
        Case 1
            FgInter(Indice).AllowUserResizing = flexResizeColumns
            FgInter(Indice).AutoSearch = flexSearchFromTop
            FgInter(Indice).ExplorerBar = flexExSortShowAndMove
            FgInter(Indice).SelectionMode = flexSelectionByRow
            FgInter(Indice).ForeColorSel = &H80000005
            FgInter(Indice).BackColorSel = &H80&
            FgInter(Indice).RowHeight(0) = 300
            FgInter(Indice).ColWidth(0) = 0
            FgInter(Indice).TextMatrix(0, 1) = "id"
            FgInter(Indice).ColWidth(1) = 0
            FgInter(Indice).TextMatrix(0, 2) = "idmae"
            FgInter(Indice).ColWidth(2) = 0
            FgInter(Indice).TextMatrix(0, 3) = "Producto"
            FgInter(Indice).ColWidth(3) = 5500
            FgInter(Indice).TextMatrix(0, 4) = "Unidad"
            FgInter(Indice).TextMatrix(0, 5) = "Programado"
            FgInter(Indice).TextMatrix(0, 6) = "Stock Ini."
            FgInter(Indice).TextMatrix(0, 7) = "Producido"
            FgInter(Indice).TextMatrix(0, 8) = "Total"
            FgInter(Indice).TextMatrix(0, 9) = "Diferencia"
    End Select
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CargaTerminados
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LA LISTA DE PRODUCTOS TERMINADOS DEL PLAN DE PRODUCCION ACTIVO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Indice       |  Integer   |  INDICE DEL CONTROL Fg5
'* DEVUELVE         :
'*****************************************************************************************************
Sub CargaTerminados(Indice As Integer)
    Dim A, B As Integer
    
    Dim RstTmp As New ADODB.Recordset
    Dim RstPro As New ADODB.Recordset
    Dim RstTmpAux As New ADODB.Recordset
    
    Dim cSQL As String
    Dim salIni As Double
    Dim xTotal As Double
    Dim xDiferencia As Double
    
    'Se consulta los datos del Plan de Produccion Activo
    cSQL = "SELECT ges_plaprod.id, ges_plaprod.fchini, ges_plaprod.fchfin, ges_plaprod.activo " _
        + vbCr + "From ges_plaprod " _
        + vbCr + "WHERE (((ges_plaprod.activo)=-1))"
        
    RST_Busq RstTmpAux, cSQL, xCon2
    
    cSQL = "SELECT alm_inventario.id, alm_inventario.idmae, alm_inventario.descripcion, mae_unidades.abrev, Sum(ges_plaproddet.cantidad) AS SumaDecantidad" _
        + vbCr + "FROM ges_plaprod LEFT JOIN (mae_unidades RIGHT JOIN (ges_plaproddet LEFT JOIN alm_inventario ON ges_plaproddet.codpro = alm_inventario.id) ON mae_unidades.id = alm_inventario.idunimed) ON ges_plaprod.id = ges_plaproddet.idpv Where (((ges_plaproddet.idmes) <> 13)) " _
        + vbCr + "GROUP BY alm_inventario.id, alm_inventario.idmae, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckini, alm_inventario.idunimed, ges_plaprod.activo" _
        + vbCr + "HAVING (((ges_plaprod.activo)=-1))"
    
    RST_Busq RstPro, cSQL, xCon2
    
    Set FgTerm(Indice).DataSource = RstPro.DataSource
    FgTerm(Indice).Cols = 10
    If RstPro.RecordCount <> 0 Then
        FrmProgreso.Visible = True
        ProgressBar1.Max = RstPro.RecordCount
        ProgressBar2.Max = RstPro.RecordCount
        LblProcesa = "Procesando Productos Terminados"
        FrmProgreso.Refresh
        
        RstPro.MoveFirst
        For A = 1 To FgTerm(Indice).Rows - 1
            If Not RstEmp.EOF Then Label6 = RstEmp("abrevia")
            Label6 = Label6 & " - " & FgTerm(Indice).TextMatrix(A, 3)
            Label11 = FgTerm(Indice).TextMatrix(A, 3)
            ProgressBar1.Value = A
            ProgressBar2.Value = A
            'Se consulta el saldo actual hasta un dia antes del nuevo plan
            'Dim xTot As Double
            salIni = SaldoActual(FgTerm(Indice).TextMatrix(A, 1), "01/01/" & AnoTra, CDate(RstTmpAux("fchini") - 1), xCon2)

            FgTerm(Indice).TextMatrix(A, 6) = Format(salIni, FORMAT_MONTO)

            Set RstTmp = Nothing

            cSQL = "SELECT Tabla001.SumaDecantidad AS totProd, Tabla002.SumaDecantidad AS prodParc, [Tabla001].[SumaDecantidad]-[Tabla002].[SumaDecantidad] AS dif " _
                + vbCr + "FROM [SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro WHERE (((pro_produccion.dia)>=CDate('01/01/2010') And (pro_produccion.dia)<CDate('26/10/2010'))) GROUP BY pro_producciondet.iditem ;]. AS Tabla002 RIGHT JOIN [SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad FROM pro_producciondet GROUP BY pro_producciondet.iditem;]. AS Tabla001 ON Tabla002.iditem = Tabla001.iditem " _
                + vbCr + "Where (((Tabla001.iditem) = " & FgTerm(Indice).TextMatrix(A, 1) & "))"
            
            RST_Busq RstTmp, cSQL, xCon2
            
            If RstTmp.RecordCount <> 0 Then
                If RstTmp("dif") <> "" Then
'                    'Se realiza la diferencia del total producido
'                    'menos lo producido hasta el primer dia de la programacion
                    FgTerm(Indice).TextMatrix(A, 7) = Format(RstTmp("dif"), FORMAT_MONTO)
                Else
                    FgTerm(Indice).TextMatrix(A, 7) = Format(NulosN(RstTmp("totProd")), FORMAT_MONTO)
                End If
            Else
                FgTerm(Indice).TextMatrix(A, 7) = "0.00"
            End If

            xTotal = NulosN(FgTerm(Indice).TextMatrix(A, 6)) + NulosN(FgTerm(Indice).TextMatrix(A, 7))
            FgTerm(Indice).TextMatrix(A, 8) = xTotal 'Format(xTotal, FORMAT_MONTO)
            xDiferencia = NulosN(FgTerm(Indice).TextMatrix(A, 5)) - NulosN(FgTerm(Indice).TextMatrix(A, 8))
            FgTerm(Indice).TextMatrix(A, 9) = Format(xDiferencia, FORMAT_MONTO)
'
            With FgTerm(Indice)
                .Select A, 9, A, 9
                .FillStyle = flexFillRepeat
                If NulosN(FgTerm(Indice).TextMatrix(A, 9)) <= 0 Then
                    .CellForeColor = &HFF0000
                Else
                    .CellForeColor = &HFF&
                End If
            End With
        Next A
    End If

    With FgTerm(Indice)
        .Select 1, 4, FgTerm(Indice).Rows - 1, 4
        .FillStyle = flexFillRepeat
        .CellBackColor = &HDDFFFF
    End With

    With FgTerm(Indice)
        .Select 1, 5, FgTerm(Indice).Rows - 1, 9
        .FillStyle = flexFillRepeat
        .CellBackColor = &HE0FEE7
        .Select 1, 1, 1, 1
    End With
    FrmProgreso.Visible = False
    
    Label6 = ""
    configurarFG 0, Indice
    cargoTerminado(Indice) = True
End Sub


Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CargarIntermedios
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LA LISTA DE PRODUCTOS INTERMEDIOS DEL PLAN DE PRODUCCION ACTIVO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Indice       |  Integer   |  INDICE DEL CONTROL Fg6
'* DEVUELVE         :
'*****************************************************************************************************
Sub CargarIntermedios(Indice As Integer)
    Dim A, B As Integer
    Dim RstTmp As New ADODB.Recordset
    Dim RstPro As New ADODB.Recordset
    Dim RstTmpAux As New ADODB.Recordset
    
    Set RstPro = Nothing
    Dim cSQL As String
    Dim salIni As Double
    Dim xTotal As Double
    Dim xDiferencia As Double
    
    'Se consulta los datos del Plan de Produccion Activo
    cSQL = "SELECT ges_plaprod.id, ges_plaprod.fchini, ges_plaprod.fchfin, ges_plaprod.activo " _
        + vbCr + "From ges_plaprod " _
        + vbCr + "WHERE (((ges_plaprod.activo)=-1))"
    
    RST_Busq RstTmpAux, cSQL, xCon2
    
    cSQL = "SELECT alm_inventario.id, alm_inventario.idmae, alm_inventario.descripcion, mae_unidades.abrev, Sum(ges_plaproddet2.cantidad) AS SumaDecantidad, alm_inventario.stckini " _
        + vbCr + "FROM mae_unidades RIGHT JOIN (ges_plaprod LEFT JOIN (ges_plaproddet2 LEFT JOIN alm_inventario ON ges_plaproddet2.codpro = alm_inventario.id) ON ges_plaprod.id = ges_plaproddet2.idpv) ON mae_unidades.id = alm_inventario.idunimed " _
        + vbCr + "Where (((ges_plaproddet2.idmes) <> 13)) " _
        + vbCr + "GROUP BY alm_inventario.id, alm_inventario.idmae, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckini, ges_plaprod.activo " _
        + vbCr + "HAVING (((ges_plaprod.activo)=-1))"
        
    RST_Busq RstPro, cSQL, xCon2
    
    Set FgInter(Indice).DataSource = RstPro.DataSource
    FgInter(Indice).Cols = 10
    If RstPro.RecordCount <> 0 Then
    
        FrmProgreso.Visible = True
        ProgressBar1.Max = RstPro.RecordCount
        ProgressBar3.Max = RstPro.RecordCount
        LblProcesa = "Procesando Productos Intermedios"
        FrmProgreso.Refresh
        
        RstPro.MoveFirst
        For A = 1 To FgInter(Indice).Rows - 1
            ProgressBar1.Value = A
            ProgressBar3.Value = A
            If Not RstEmp.EOF Then Label10 = RstEmp("abrevia")
            Label10 = Label10 & " - " & FgInter(Indice).TextMatrix(A, 3)
            Label11 = FgInter(Indice).TextMatrix(A, 3)
            'Se consulta el saldo actual hasta un dia antes del nuevo plan
            salIni = SaldoActual(FgInter(Indice).TextMatrix(A, 1), "01/01/" & AnoTra, CDate(RstTmpAux("fchini") - 1), xCon2)
            
            FgInter(Indice).TextMatrix(A, 6) = Format(salIni, FORMAT_MONTO)
            
            
            Set RstTmp = Nothing

            cSQL = "SELECT Tabla001.SumaDecantidad AS totProd, Tabla002.SumaDecantidad AS prodParc, [Tabla001].[SumaDecantidad]-[Tabla002].[SumaDecantidad] AS dif " _
                + vbCr + "FROM [SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro WHERE (((pro_produccion.dia)>=CDate('01/01/2010') And (pro_produccion.dia)<CDate('26/10/2010'))) GROUP BY pro_producciondet.iditem ;]. AS Tabla002 RIGHT JOIN [SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad FROM pro_producciondet GROUP BY pro_producciondet.iditem;]. AS Tabla001 ON Tabla002.iditem = Tabla001.iditem " _
                + vbCr + "Where (((Tabla001.iditem) = " & FgInter(Indice).TextMatrix(A, 1) & "))"
            
            RST_Busq RstTmp, cSQL, xCon2
            
            If RstTmp.RecordCount <> 0 Then
                If RstTmp("dif") <> "" Then
'                    'Se realiza la diferencia del total producido
'                    'menos lo producido hasta el primer dia de la programacion
                    FgInter(Indice).TextMatrix(A, 7) = Format(RstTmp("dif"), FORMAT_MONTO)
                Else
                    FgInter(Indice).TextMatrix(A, 7) = Format(NulosN(RstTmp("totProd")), FORMAT_MONTO)
                End If
            Else
                FgInter(Indice).TextMatrix(A, 7) = "0.00"
            End If
            
            xTotal = NulosN(FgInter(Indice).TextMatrix(A, 6)) + NulosN(FgInter(Indice).TextMatrix(A, 7))
            FgInter(Indice).TextMatrix(A, 8) = Format(xTotal, FORMAT_MONTO)
            
            xDiferencia = NulosN(FgInter(Indice).TextMatrix(A, 8)) - NulosN(FgInter(Indice).TextMatrix(A, 5))
            FgInter(Indice).TextMatrix(A, 9) = Format(xDiferencia, FORMAT_MONTO)
            
            With FgInter(Indice)
                .Select A, 9, A, 9
                .FillStyle = flexFillRepeat
                If NulosN(FgInter(Indice).TextMatrix(A, 9)) > 0 Then
                    .CellForeColor = &HFF0000
                Else
                    .CellForeColor = &HFF&
                End If
            End With
            
            RstPro.MoveNext
            
            If RstPro.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    With FgInter(Indice)
        .Select 1, 4, FgInter(Indice).Rows - 1, 4
        .FillStyle = flexFillRepeat
        .CellBackColor = &HDDFFFF
    End With

    With FgInter(Indice)
        .Select 1, 5, FgInter(Indice).Rows - 1, 9
        .FillStyle = flexFillRepeat
        .CellBackColor = &HE0FEE7
        .Select 1, 1, 1, 1
    End With
    
    Label10 = ""
    configurarFG 1, Indice
    cargoIntermedio(Indice) = True
    FrmProgreso.Visible = False
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : BuscaReceta
'* Tipo             : FUNCION
'* Descripcion      : BUSCA UNA RECETA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    CodigoItem   |  Integer   |  ESPECIFICA EL ID DEL PRODUCTO
'*                    QueDevuelve  |  Devolver  |  ESPECIFICA EL VALOR DEL TIPO Devolver DEFINIDO
'* DEVUELVE         :
'*****************************************************************************************************
Function BuscaReceta(CodigoItem As Integer, QueDevuelve As Devolver) As Variant
    Dim xRst As New ADODB.Recordset
    
    RST_Busq xRst, "SELECT pro_receta.iditem, pro_receta.codrec, pro_recetains.iditem, pro_recetains.canpro, alm_inventario.descripcion FROM alm_inventario RIGHT JOIN " _
        & " (pro_receta INNER JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) ON alm_inventario.id = pro_recetains.iditem " _
        & " WHERE (((pro_receta.iditem)=" & CodigoItem & ") AND ((pro_receta.prirec)=1) AND ((alm_inventario.tippro)=1))", xCon2

    If xRst.RecordCount <> 0 Then
        If QueDevuelve = Cantidad Then
            BuscaReceta = NulosC(xRst("codrec"))
        Else
            BuscaReceta = NulosN(xRst("canpro"))
        End If
    Else
        BuscaReceta = 0
    End If
    Set xRst = Nothing
End Function

Private Sub procesarTodasEmpresas()
    Dim A As Integer
    Dim xIndex As Integer
    Dim xRuta As String
        
    If RstEmp.RecordCount <> 0 Then
        xIndex = 0
        RstEmp.MoveFirst
        frmProgresoTot.Left = 3300
        frmProgresoTot.Top = 2640
        frmProgresoTot.Visible = True
        frmProgresoTot.Refresh
        For A = 1 To RstEmp.RecordCount
            xRuta = AP_RUTABD + Trim(RstEmp("ruta"))
            Set xCon2 = Nothing
            Set xCon2 = AbrirConecciones(xRuta)
            If Not cargoTerminado(xIndex) Then CargaTerminados xIndex
            If Not cargoIntermedio(xIndex) Then CargarIntermedios xIndex
            
            RstEmp.MoveNext
            
            If RstEmp.EOF = True Then
                Exit For
            End If
            xIndex = xIndex + 1
        Next A
        frmProgresoTot.Visible = False
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : VerUnificado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LA INFORMACION UNIFICADA DEL PROGRAMA DE PRODUCCION, PARA ELLO JALA EL
'*                    PROGRAMA DE PRODUCCION DE LAS DEMAS BASES DE DATO ACTIVAS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub VerUnificado()
    Dim A As Integer
    Dim Total As Double
    Dim xIndex As Integer
    
    procesarTodasEmpresas
    
    TabOne1.Enabled = False
    Toolbar1.Enabled = False
    Frame2.Left = 0
    Frame2.Top = 350
    Frame2.Visible = True
    
    PreparaRST
    
    xIndex = 0
    
    For A = 1 To 5
        Dim B As Integer
        If TabOne1.TabVisible(xIndex) = True Then
            For B = 1 To FgTerm(xIndex).Rows - 1
                RstInsumos.Filter = adFilterNone
                If RstInsumos.RecordCount <> 0 Then
                    RstInsumos.MoveFirst
                End If
                
                RstInsumos.Filter = "cod_item = '" & FgTerm(xIndex).TextMatrix(B, 0) & "'"
                If RstInsumos.RecordCount = 0 Then
                    RstInsumos.AddNew
                    RstInsumos("descripcion") = FgTerm(xIndex).TextMatrix(B, 3)
                    RstInsumos("unimed") = FgTerm(xIndex).TextMatrix(B, 4)
                    RstInsumos("programado") = Format(NulosN(FgTerm(xIndex).TextMatrix(B, 5)), FORMAT_MONTO)
                    RstInsumos("stckini") = Format(NulosN(FgTerm(xIndex).TextMatrix(B, 6)), FORMAT_MONTO)
                    RstInsumos("producido") = Format(NulosN(FgTerm(xIndex).TextMatrix(B, 7)), FORMAT_MONTO)
                    RstInsumos("total") = Format(NulosN(FgTerm(xIndex).TextMatrix(B, 8)), FORMAT_MONTO)
                    'RstInsumos("porprod") = NulosN(FgTerm(xIndex).TextMatrix(B, 8))
                    RstInsumos("saldo") = Format(NulosN(FgTerm(xIndex).TextMatrix(B, 9)), FORMAT_MONTO)
                    RstInsumos("cod_item") = FgTerm(xIndex).TextMatrix(B, 1)
                Else
                    If RstInsumos.RecordCount = 1 Then
                        RstInsumos("programado") = RstInsumos("programado") + NulosN(FgTerm(xIndex).TextMatrix(B, 5))
                        'RstInsumos("porprod") = RstInsumos("porprod") + NulosN(FgTerm(xIndex).TextMatrix(B, 8))
                        RstInsumos("saldo") = RstInsumos("saldo") + NulosN(FgTerm(xIndex).TextMatrix(B, 9))
                    Else
                        'este error nunca debe de ocurrir
                        MsgBox "Hay mas de un items con el mismo codigo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    End If
                End If
            Next B
            
            For B = 1 To FgInter(xIndex).Rows - 1
                RstInsumos.Filter = adFilterNone
                If RstInsumos.RecordCount <> 0 Then
                    RstInsumos.MoveFirst
                End If
                RstInsumos.Filter = "cod_item = '" & FgInter(xIndex).TextMatrix(B, 0) & "'"
                If RstInsumos.RecordCount = 0 Then
                    RstInsumos.AddNew
                    RstInsumos("descripcion") = FgInter(xIndex).TextMatrix(B, 3)
                    RstInsumos("unimed") = FgInter(xIndex).TextMatrix(B, 4)
                    RstInsumos("programado") = Format(NulosN(FgInter(xIndex).TextMatrix(B, 5)), FORMAT_MONTO)
                    RstInsumos("stckini") = Format(NulosN(FgInter(xIndex).TextMatrix(B, 6)), FORMAT_MONTO)
                    RstInsumos("producido") = Format(NulosN(FgInter(xIndex).TextMatrix(B, 7)), FORMAT_MONTO)
                    RstInsumos("total") = Format(NulosN(FgInter(xIndex).TextMatrix(B, 8)), FORMAT_MONTO)
                    'RstInsumos("porprod") = NulosN(FgInter(xIndex).TextMatrix(B, 8))
                    RstInsumos("saldo") = Format(NulosN(FgInter(xIndex).TextMatrix(B, 9)), FORMAT_MONTO)
                    RstInsumos("cod_item") = FgInter(xIndex).TextMatrix(B, 1)
                Else
                    If RstInsumos.RecordCount = 1 Then
                        RstInsumos("programado") = RstInsumos("programado") + NulosN(FgInter(xIndex).TextMatrix(B, 5))
                        RstInsumos("saldo") = RstInsumos("saldo") + NulosN(FgInter(xIndex).TextMatrix(B, 9))
                    Else
                        MsgBox "Hay mas de un items con el mismo codigo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    End If
                End If
            Next B
        End If
        xIndex = xIndex + 1
    Next A
    
    RstInsumos.Filter = adFilterNone
    RstInsumos.Sort = "descripcion"
    RstInsumos.MoveFirst
    Fg7.Rows = 1
    For A = 1 To RstInsumos.RecordCount
        Fg7.Rows = Fg7.Rows + 1
        Fg7.TextMatrix(A, 1) = RstInsumos("descripcion")
        Fg7.TextMatrix(A, 2) = RstInsumos("cod_item")
        Fg7.TextMatrix(A, 3) = RstInsumos("unimed")
        Fg7.TextMatrix(A, 4) = Format(RstInsumos("programado"), FORMAT_MONTO)
        Fg7.TextMatrix(A, 5) = Format(RstInsumos("stckini"), FORMAT_MONTO)
        Fg7.TextMatrix(A, 6) = Format(RstInsumos("producido"), FORMAT_MONTO)
        Fg7.TextMatrix(A, 7) = Format(RstInsumos("total"), FORMAT_MONTO)
        Fg7.TextMatrix(A, 8) = (NulosN(RstInsumos("programado")) - NulosN(RstInsumos("total")))
        
        With Fg7
            .Select A, 8, A, 8
            .FillStyle = flexFillRepeat
            If NulosN(Fg7.TextMatrix(A, 8)) <= 0 Then
                .CellForeColor = &HFF0000
            Else
                .CellForeColor = &HFF&
            End If
        End With
        Fg7.TextMatrix(A, 8) = Abs(NulosN(NulosN(Fg7.TextMatrix(A, 8))))
        Fg7.TextMatrix(A, 8) = Format(NulosN(Fg7.TextMatrix(A, 8)), FORMAT_MONTO)
        RstInsumos.MoveNext
        If RstInsumos.EOF = True Then
            Exit For
        End If
    Next A
    
    With Fg7
        .Select 1, 4, Fg7.Rows - 1, 4
        .FillStyle = flexFillRepeat
        .CellBackColor = &HDDFFFF
    End With

    With Fg7
        .Select 1, 5, Fg7.Rows - 1, 7
        .FillStyle = flexFillRepeat
        .CellBackColor = &HE0FEE7
        .Select 1, 1, 1, 1
    End With
    
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : PreparaRST
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA UN RECORDSET TEMPORAL
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub PreparaRST()
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(9, 3) As String

    xCampos(0, 0) = "cod_item":     xCampos(0, 1) = "C":      xCampos(0, 2) = "16"
    xCampos(1, 0) = "unimed":       xCampos(1, 1) = "C":      xCampos(1, 2) = "4"
    xCampos(2, 0) = "descripcion":  xCampos(2, 1) = "C":      xCampos(2, 2) = "200"
    xCampos(3, 0) = "programado":   xCampos(3, 1) = "N":      xCampos(3, 2) = "2"
    xCampos(4, 0) = "stckini":      xCampos(4, 1) = "N":      xCampos(4, 2) = "2"
    xCampos(5, 0) = "producido":    xCampos(5, 1) = "N":      xCampos(5, 2) = "2"
    xCampos(6, 0) = "total":        xCampos(6, 1) = "N":      xCampos(6, 2) = "2"
    xCampos(7, 0) = "porprod":      xCampos(7, 1) = "N":      xCampos(7, 2) = "2"
    xCampos(8, 0) = "saldo":        xCampos(8, 1) = "N":      xCampos(8, 2) = "2"
    
    Set RstInsumos = xFun.CrearRstTMP(xCampos)
    RstInsumos.Open
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    Dim A As Integer
    
    RST_Busq RstEmp, "SELECT mae_empresa.* From mae_empresa WHERE (((mae_empresa.anotra)=" & AnoTra & ") AND ((mae_empresa.activo)=-1))", xCon1
    If Not RstEmp.RecordCount Then
        RstEmp.MoveFirst
        For A = 0 To NewTab
            xRuta = AP_RUTABD + Trim(RstEmp("ruta"))
            Set xCon2 = Nothing
            Set xCon2 = AbrirConecciones(xRuta)
            RstEmp.MoveNext
        Next A
        If NewTab = 1 And Not cargoTerminado(NewTab) Then CargaTerminados NewTab
    End If
End Sub

Private Sub TabOne2_Switch(Index As Integer, OldTab As Integer, NewTab As Integer, Cancel As Integer)
    Dim A As Integer
    
    RST_Busq RstEmp, "SELECT mae_empresa.* From mae_empresa WHERE (((mae_empresa.anotra)=" & AnoTra & ") AND ((mae_empresa.activo)=-1))", xCon1
    RstEmp.MoveFirst
    For A = 0 To Index
        xRuta = AP_RUTABD + Trim(RstEmp("ruta"))
        Set xCon2 = Nothing
        Set xCon2 = AbrirConecciones(xRuta)
        RstEmp.MoveNext
    Next A
    If NewTab = 1 And Not cargoIntermedio(Index) Then CargarIntermedios Index
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then VerUnificado
    
    If Button.Index = 3 Then
        Unload Me
    End If
End Sub
