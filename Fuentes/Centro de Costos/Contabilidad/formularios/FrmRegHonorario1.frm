VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRegHonorario1 
   Caption         =   "Contabilidad - Renta de 4ta"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1245
      Left            =   3180
      TabIndex        =   33
      Top             =   7890
      Visible         =   0   'False
      Width           =   5010
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   120
         TabIndex        =   34
         Top             =   630
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
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
         TabIndex        =   35
         Top             =   105
         Width           =   1665
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
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   4995
         X2              =   4995
         Y1              =   30
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
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5010
         Y1              =   1230
         Y2              =   1230
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
   Begin SizerOneLibCtl.TabOne TabOne2 
      Height          =   1275
      Left            =   30
      TabIndex        =   2
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
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1185
         Left            =   345
         TabIndex        =   11
         Top             =   45
         Width           =   11490
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
            Left            =   9840
            TabIndex        =   29
            Top             =   0
            Width           =   1635
            Begin VB.Label LblNumreg 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblNumreg"
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
               TabIndex        =   31
               Top             =   630
               Width           =   1440
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nº Registros :"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   30
               Top             =   390
               Width           =   975
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
            Left            =   6870
            TabIndex        =   24
            Top             =   0
            Width           =   2910
            Begin VB.OptionButton OptSort4 
               Caption         =   "Fch. Emision y Nº de Documento"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   28
               Top             =   930
               Width           =   2730
            End
            Begin VB.OptionButton OptSort3 
               Caption         =   "Nº Registro"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   27
               Top             =   700
               Width           =   2010
            End
            Begin VB.OptionButton OptSort1 
               Caption         =   "Fecha  de Emisión"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   26
               Top             =   240
               Width           =   2010
            End
            Begin VB.OptionButton OptSort2 
               Caption         =   "Nº de Documento"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   25
               Top             =   470
               Width           =   2010
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
            TabIndex        =   21
            Top             =   0
            Width           =   2700
            Begin VB.TextBox txtBancarizacion 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1650
               TabIndex        =   32
               Text            =   "txtBancarizacion"
               Top             =   450
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.OptionButton OptOpc11 
               Caption         =   "Todos los Comprobantes"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Width           =   2070
            End
            Begin VB.OptionButton OptOpc22 
               Caption         =   "Bancarización"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   22
               Top             =   555
               Width           =   1920
            End
         End
         Begin VB.Frame Frame1 
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
            ForeColor       =   &H00400000&
            Height          =   570
            Left            =   30
            TabIndex        =   16
            Top             =   0
            Width           =   4095
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   735
               TabIndex        =   17
               Top             =   240
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
               TabIndex        =   18
               Top             =   240
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
               Caption         =   "Hasta"
               Height          =   195
               Index           =   2
               Left            =   2145
               TabIndex        =   20
               Top             =   285
               Width           =   420
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   19
               Top             =   270
               Width           =   465
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
            TabIndex        =   12
            Top             =   570
            Width           =   4095
            Begin VB.CommandButton CmdBusMon 
               Height          =   230
               Left            =   495
               Picture         =   "FrmRegHonorario1.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   180
               MaxLength       =   1
               TabIndex        =   14
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
               TabIndex        =   15
               Top             =   240
               Width           =   3135
            End
         End
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   12825
         TabIndex        =   3
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
            TabIndex        =   4
            Top             =   0
            Width           =   9390
            Begin VB.CommandButton CmdBusCliPro 
               Enabled         =   0   'False
               Height          =   240
               Left            =   8640
               Picture         =   "FrmRegHonorario1.frx":0132
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   210
               Width           =   210
            End
            Begin VB.OptionButton OptSel2 
               Caption         =   "Seleccionar"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1170
               TabIndex        =   6
               Top             =   270
               Width           =   1140
            End
            Begin VB.OptionButton OptSel1 
               Caption         =   "Todos"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   5
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
               TabIndex        =   8
               Text            =   "TxtCliPro"
               Top             =   180
               Width           =   5475
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Proveedor"
               Height          =   195
               Index           =   2
               Left            =   2610
               TabIndex        =   10
               Top             =   270
               Width           =   735
            End
            Begin VB.Label LblIdCliPro 
               Caption         =   "LblIdCliPro"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   2280
               TabIndex        =   9
               Top             =   150
               Visible         =   0   'False
               Width           =   750
            End
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   -120
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
            Picture         =   "FrmRegHonorario1.frx":0264
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":07A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":0B3A
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":0CBE
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":1112
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":122A
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":176E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":1CB2
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":1DC6
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":1EDA
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":232E
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":249A
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":29E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":2CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":308E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegHonorario1.frx":3420
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
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
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
            Object.ToolTipText     =   "Exportar a PDT"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6015
      Left            =   0
      TabIndex        =   1
      Top             =   1650
      Width           =   11880
      _cx             =   20955
      _cy             =   10610
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
      Rows            =   3
      Cols            =   16
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmRegHonorario1.frx":37B2
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
Attribute VB_Name = "FrmRegHonorario1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FrmRegHonorario.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : EMITE EL LIBRO REGISTRO DE DE HONORARIOS, EN FUNCION A CRITERIOS ESPECIFICADOS POR EL
'                     USUARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 21/10/09
'* VERSION          : 1.0
'Modificado         : 10/02/10 - Johan Castro
'                     Agregar filtro por prestador de servicio

'*****************************************************************************************************
Option Explicit

Dim SeEjecuto As Boolean      ' VERIFICA QUE EL EVENTO ACTIVATE SE HAYA EJECUTADO
Dim xFormatoActual As Integer ' ESPECIFICA EL ID DEL FORMATO ACTUAL

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
    
        TxtCliPro.Text = ""
        
        TabOne2.CurrTab = 0
        
        txtBancarizacion.Text = "0.00"
        
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
    
    
        SeEjecuto = True
        TxtFchIni.SetFocus
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    OptOpc11.Value = True
    OptSort3.Value = True
    LblNumreg.Caption = ""
    
    Dim xRs As New ADODB.Recordset
    
    ' SELECCIONAMOS EL FORMATO ACTUAL PARA LA IMPRESION DEL LIBRO "REGISTRO DE COMPRAS"
    RST_Busq xRs, "SELECT con_formatostipo.* From con_formatostipo WHERE (((con_formatostipo.defecto)=-1) AND ((con_formatostipo.idformato)=4))", xCon

    xFormatoActual = xRs("id")
    Set xRs = Nothing
    SetearCuadricula Fg1, 4, xCon, 1, xFormatoActual, False
    
    '--buscar los registros
    Fg1.AutoSearch = flexSearchFromTop
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
End Sub

'*****************************************************************************************************
'* Nombre           : MostrarRegistros
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL REGISTRO DE HONORARIOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MostrarRegistros()
    Dim Rst As New ADODB.Recordset
    Dim A As Long
    Dim nSQL As String
    Dim nSQLCampos As String '--Relacion de campos a mostrar, obtenido de tabla: con_formatostipodet
    Dim nSQLFiltro As String '--Sentencia que indica el filtro a la consulta
    Dim nSQLOrden As String '--Sentencia que indica el orden de la consulta
    
    '--limpiar datos
    LblNumreg.Caption = "0"
    Fg1.Rows = 2
    DoEvents
    '----
    
    Me.MousePointer = vbHourglass
    
    '--obtener el orden de presentacion de los campos
    nSQLCampos = fSetearCuadriculaColumna(xCon, 4)
    '--verificar si hay campos seleccionados para mostrar el reporte
    If nSQLCampos = "" Then Exit Sub
    
    
    '--verificar si hay filtro por cliente
    If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLFiltro = " and com_honorarios.idpro = " & NulosN(LblIdCliPro.Caption) & " "

'''de baja
'''     If OptOpc22.Value = True Then
'''        ' mostramos solo los de bancarizacion en Soles
'''        If TxtIdMon.Text = 1 Then nSQLFiltro = nSQLFiltro & " and IIf([com_honorarios]![idmon]=2,[imptot],IIF(tipcam =0,0,[imptot]/tipcam)) > " & NulosN(txtBancarizacion.Text)
'''        ' mostramos solo los de bancarizacion en Dolares
'''        If TxtIdMon.Text = 2 Then nSQLFiltro = nSQLFiltro & " and IIf([com_honorarios]![idmon]=2,[imptot],IIF(tipcam =0,0,[imptot]/tipcam)) > " & NulosN(txtBancarizacion.Text)
'''    End If
    
   
     ' ORDENA POR FECHA DE DOCUMENTO
    If OptSort1.Value = True Then nSQLOrden = "Order by com_honorarios.fchdoc"
    ' ORDENA POR NUMERO DE DOCUMENTO
    If OptSort2.Value = True Then nSQLOrden = "Order by IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc])"
    ' ORDENA POR NUMERO DE REGISTRO
    If OptSort3.Value = True Then nSQLOrden = "Order by Mid([com_honorarios]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_honorarios]![numreg],3,4)"
    ' ORDENA POR FECHA DE DOCUMENTO Y NUMERO DE DOCUMENTO
    If OptSort4.Value = True Then nSQLOrden = "Order by com_honorarios.fchdoc,IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc])"
    
    
    
    If TxtIdMon.Text = 1 Then
        ' SI LA CONSULTA ES EN SOLES
        nSQL = "SELECT Mid([com_honorarios]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_honorarios]![numreg],3,4) AS registro, com_honorarios.fchreg, " _
            & " com_honorarios.fchdoc, com_honorarios.fchven, mae_moneda.simbolo, com_honorarios.numser, com_honorarios.numdoc, mae_dociden.codsun, mae_prov.numruc, " _
            & " mae_prov.nombre, con_tc.impven, IIf([com_honorarios].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[com_honorarios].[tc]) AS tipcam, " _
            & " IIf([com_honorarios]![idmon]=1,[impbru],[impbru]*tipcam) AS bruto, " _
            & " IIf([com_honorarios]![idmon]=1,[impigv],[impigv]*tipcam) AS impuesto, " _
            & " 0 AS otrasret, " _
            & " IIf([com_honorarios]![idmon]=1,[imptot],[imptot]*tipcam) AS total, " _
            & " com_honorarios.glosa, " _
            & " IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc]) AS numdoc2, " _
            & " Mid([com_honorarios]![numreg],1,2) AS idmes,mae_condpago.descripcion AS condpag, 2 as tipcom, iif(impuesto=0,0,1) as apliret " _
            & " FROM ((((mae_dociden RIGHT JOIN (com_honorarios LEFT JOIN mae_prov ON com_honorarios.idpro = mae_prov.id) ON mae_dociden.id = mae_prov.idtipdoc) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) " _
            & " LEFT JOIN mae_moneda ON com_honorarios.idmon = mae_moneda.id) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN mae_condpago ON com_honorarios.idconpag = mae_condpago.id " _
            & " WHERE (((com_honorarios.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (com_honorarios.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
            & " AND (con_tc.idmon=2 OR con_tc.idmon IS NULL) AND ((Mid([com_honorarios]![numreg],1,2))<>'00')) " & nSQLFiltro

    ElseIf TxtIdMon.Text = 2 Then
    
        ' SI LA CONSULTA ES EN DOLARES
        nSQL = "SELECT Mid([com_honorarios]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_honorarios]![numreg],3,4) AS registro, com_honorarios.fchreg, " _
            & " com_honorarios.fchdoc, com_honorarios.fchven, mae_moneda.simbolo, com_honorarios.numser, com_honorarios.numdoc, mae_dociden.codsun, mae_prov.numruc, " _
            & " mae_prov.nombre, con_tc.impven, IIf([com_honorarios].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[com_honorarios].[tc]) AS tipcam, " _
            & " IIf([com_honorarios]![idmon]=2,[impbru],IIF(tipcam =0,0,[impbru]/tipcam)) AS bruto, " _
            & " IIf([com_honorarios]![idmon]=2,[impigv],IIF(tipcam =0,0,[impigv]/tipcam)) AS impuesto, " _
            & " 0 AS otrasret, " _
            & " IIf([com_honorarios]![idmon]=2,[imptot],IIF(tipcam =0,0,[imptot]/tipcam)) AS total, " _
            & " com_honorarios.glosa, " _
            & " IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc]) AS numdoc2, " _
            & " Mid([com_honorarios]![numreg],1,2) AS idmes,mae_condpago.descripcion AS condpag, 2 as tipcom, iif(impuesto=0,0,1) as apliret " _
            & " FROM ((((mae_dociden RIGHT JOIN (com_honorarios LEFT JOIN mae_prov ON com_honorarios.idpro = mae_prov.id) ON mae_dociden.id = mae_prov.idtipdoc) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) " _
            & " LEFT JOIN mae_moneda ON com_honorarios.idmon = mae_moneda.id) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN mae_condpago ON com_honorarios.idconpag = mae_condpago.id " _
            & " WHERE (((com_honorarios.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (com_honorarios.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
            & " AND (con_tc.idmon=2 OR con_tc.idmon IS NULL) AND ((Mid([com_honorarios]![numreg],1,2))<>'00')) " & nSQLFiltro

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
            Case "bruto":       ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impuesto":    ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "otrasret":    ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "total":       ArrCampos(mCol) = mColCampo: mCol = mCol + 1
        End Select
    Next mColCampo
       
    If OptOpc11.Value = True Then Rst.Filter = adFilterNone                   ' mostramos todos los registros
    If OptOpc22.Value = True Then
        If TxtIdMon.Text = 1 Then Rst.Filter = "total > " & NulosN(txtBancarizacion.Text)   ' mostramos solo los de bancarizacion en Soles
        If TxtIdMon.Text = 2 Then Rst.Filter = "total > " & NulosN(txtBancarizacion.Text)   ' mostramos solo los de bancarizacion en Dolares
    End If
    
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
                
                Case "bruto", "impuesto", "otrasret", "total"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_MONTO)
                                    
                Case "tdpersun"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
                Case Else
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
            End Select
            
        Next mCol
                
        '--verificar si monto=cero y no sea anulado =>> pintar la fila para que muestre una alerta al usuario
        If RstRegistroBuscaCampo(Rst, "total") = True Then
            If NulosN(Rst("total")) = 0 Then
                GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HC0C0FF
            End If
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

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub

    If Me.Height > 3000 Then
        Fg1.Top = 1650
        Fg1.Width = Me.Width - 150
        Fg1.Height = Me.Height - 2100
    End If
End Sub

Private Sub OptOpc11_Click()
    txtBancarizacion.Visible = False
End Sub

Private Sub OptOpc22_Click()
    txtBancarizacion.Visible = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
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
        '--verificar si hay bancarizacion
        If OptOpc22.Value = True Then
            If NulosN(txtBancarizacion.Text) = 0 Then
                MsgBox "Falta especificar la base de la bancarizacion expresado en " & LblMoneda.Caption, vbInformation, xTitulo
                txtBancarizacion.SetFocus
                Exit Sub
            End If
        End If
                
        MostrarRegistros
        
    End If
    
    If Button.Index = 3 Then
        If Fg1.Rows = 2 Then
            MsgBox "No hay registro que exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Dim xFun As New SGI2_funciones.formularios
        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "REGISTRO DE RETECIONES 4TA CATEGORIA", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Por Fecha", "renta4ta.xls"   ', Rst, ""
        Set xFun = Nothing
    End If
    
    If Button.Index = 4 Then
'        Dim xMoneda As String
'        Dim nPeriodo As String
'        Dim xPrint As New SGI2_funciones.formularios
'
'        xmoneda=LblMoneda.Caption
'
'        nPeriodo = "Del " & TxtFchIni.Valor & " Al " & TxtFchFin.Valor
'        Me.MousePointer = vbHourglass
'        xPrint.Imprimir_x_VSFlexGrid Fg1, "REGISTRO RENTA 4ta ", "(Expresado en " + xMoneda + ")", nPeriodo, False, True
'        Set xPrint = Nothing
'        Me.MousePointer = vbDefault

        IMPRIMIR
    End If
    
    If Button.Index = 5 Then Configurar
    
    If Button.Index = 6 Then ExportarPDT
        
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ExportarPDT
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A UN ARCHIVO PLANO EL LIBRO DE HONORARIOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ExportarPDT()
    Dim Rst As New ADODB.Recordset
    Dim NomArch, xCad As String
    Dim A As Long
    Dim B As Long
    Dim nSQL As String
    Dim nSQLCampos As String '--Relacion de campos a mostrar, obtenido de tabla: con_formatostipodet


    
    If Fg1.Rows = 2 Then
        MsgBox "No se ha mostrado ninguna retencion, haga click en el boton"
    End If

    NomArch = "0601" & AnoTra & Format(TxtFchIni.Valor, "mm") & NumRUC & ".4ta"
   
    Open Trim(App.Path) + "\" + NomArch For Output As #1

'    For A = 2 To fg1.Rows - 1
'        xCad = ""
'        xCad = xCad + fg1.TextMatrix(A, 7) + "|"                        ' tipo de documento de identidad del proveedor
'        xCad = xCad + fg1.TextMatrix(A, 8) + "|"                        ' numero de documento del proveedor
'        xCad = xCad + "2" + "|"                                         ' tipo documento de la compra
'        xCad = xCad + Format(fg1.TextMatrix(A, 5), "0000") + "|"        ' numro de serioe
'        xCad = xCad + Format(fg1.TextMatrix(A, 6), "00000000") + "|"    ' numero de documento
'        xCad = xCad + Format(fg1.TextMatrix(A, 11), "0.00") + "|"       ' monto total del servicio
'        xCad = xCad + Format(fg1.TextMatrix(A, 2), "dd/mm/yyyy") + "|"  ' fecha de emision
'        xCad = xCad + Format(fg1.TextMatrix(A, 3), "dd/mm/yyyy") + "|"  ' fecha de pago
'
'        If NulosN(fg1.TextMatrix(A, 12)) = 0 Then
'            xCad = xCad + "0" + "|"  ' especifica si se aplicat retencion de 4ta
'        Else
'            xCad = xCad + "1" + "|"  ' especifica si se aplicat retencion de 4ta
'        End If
'        Print #1, Trim(xCad)
'    Next A
    
    
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''

    DoEvents
    '----
    
    Me.MousePointer = vbHourglass
    
    '--Consulta para obtener lista de campos a exportar
    nSQL = "SELECT con_formatostipodet.nomcampo " _
        + vbCr + " From con_formatostipodet " _
        + vbCr + " Where (((con_formatostipodet.idformato) = 4) And ((con_formatostipodet.idformatotipo) = 3) And ((con_formatostipodet.mostrar) = -1)) " _
        + vbCr + " ORDER BY con_formatostipodet.orden; "
    
    RST_Busq Rst, nSQL, xCon
    
    '--obtener el orden de presentacion de los campos
    
    nSQLCampos = RstRegistroGenerarId(Rst, "nomcampo", "", "", False)
    '--limpiar rst
    Set Rst = Nothing
    
    '--verificar si hay campos seleccionados para mostrar el reporte
    If nSQLCampos = "" Then Exit Sub
    
    '--Obtener la consulta de seleccion
    nSQL = GenerarConsulta()
    
    '--armar la sentencia SQL
    nSQL = "Select " & nSQLCampos & _
            vbCr + " from ( " _
            + vbCr + nSQL _
            + vbCr + ") as consulta "
    '--ejecutar la consulta
    RST_Busq Rst, nSQL, xCon
    
    '--Salir si hay error en la consulta
    If Rst.State = 0 Then GoTo LaCague
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Rst.RecordCount <> 0 Then Rst.MoveFirst
    
    Do While Not Rst.EOF
        xCad = ""
        For B = 0 To Rst.Fields.Count - 1
            
            Select Case LCase(Rst(B).Name)
                Case "codsun", "numruc", "tipcom", "apliret"
                    xCad = xCad + NulosC(Rst(B))
                Case "numser":
                    xCad = xCad + Format(NulosC(Rst(B)), "0000")
                Case "numdoc":
                    xCad = xCad + Format(NulosC(Rst(B)), "00000000")
                Case "bruto":
                    xCad = xCad + Format(NulosN(Rst(B)), "0.00")
                Case "fchdoc", "fchven":
                    xCad = xCad + Format(NulosC(Rst(B)), "dd/mm/yyyy")
                Case Else
                    xCad = xCad + NulosC(Rst(B))
            End Select
            xCad = xCad + "|"
        Next B
        '--escribir en el bloc de notas
        Print #1, Trim(xCad)
        '--siguiente registro
        Rst.MoveNext
    Loop
    
    Close #1
    
    Set Rst = Nothing
    MsgBox "Los honorarios se exportaron para el PDT con exito : " & Trim(App.Path) + "\" + NomArch, vbInformation + vbOKCancel + vbDefaultButton1, xTitulo
    
LaCague:

    Set Rst = Nothing
    
    '--restablecer cursor
    Me.MousePointer = vbDefault
    
End Sub



'*****************************************************************************************************
'* Nombre           : Configurar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL REGISTRO DE HONORARIOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Configurar()
    Dim xform As New SGI2_funciones.Varias
    
    If xform.CambioOpcionLiro(4, xCon, 1) = True Then
    
        Dim xRs As New ADODB.Recordset
        
        ' SELECCIONAMOS EL FORMATO ACTUAL PARA LA IMPRESION DEL LIBRO "REGISTRO DE COMPRAS"
        RST_Busq xRs, "SELECT con_formatostipo.* From con_formatostipo WHERE (((con_formatostipo.defecto)=-1) AND ((con_formatostipo.idformato)=4))", xCon
    
        xFormatoActual = xRs("id")
        Set xRs = Nothing
        
        SetearCuadricula Fg1, 4, xCon, 1, xFormatoActual, False
        
        ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
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
        MostrarRegistros
    End If
    Set xform = Nothing
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
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    
    ' SELECCIONAMOS EL FORMATO ACTUAL PARA LA IMPRESION DEL LIBRO "REGISTRO DE COMPRAS"
    RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 4) And ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ") " _
        & " And ((con_formatostipodet.mostrar) = -1)) ORDER BY con_formatostipodet.orden", xCon

    Dim xCampos() As String
    Dim xFil, xCol As Double
    
    ReDim xCampos(Fg1.Rows - 2, Fg1.Cols - 1)
    'ReDim xCampos(Fg1.Rows - 2, Rst.RecordCount)
    
    Dim xFila As Double
    xFila = 0
    For xFil = 1 To Fg1.Rows - 1
        For xCol = 1 To Fg1.Cols - 1
            xCampos(xFila, xCol) = Fg1.TextMatrix(xFil, xCol)
        Next xCol
        xFila = xFila + 1
    Next xFil
    
    ' BLANQUEAMOS LAS TITULOS DE LAS COLUMNAS QUE NO SE VAN A IMPRIMIR
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
    ' CABECERA DEL REPORTE
    xfrm.Cabecera1 = NomEmp                                                   ' NOMBRE DE LA EMPRESA
    xfrm.Cabecera2 = "RUC Nº: " & NumRUC                                      ' NUMERO DE RUC DE LA EMPRESA
    xfrm.Fecha = Format(Date, "dd/mm/yyyy")                                   ' FECHA DE EMISION DEL REPORTE
    xfrm.Titulo1 = "REGISTRO RENTA DE 4ta " & "(Expresado en " & xMoneda & ")"  ' TITULO DEL REPORTE
    xfrm.Titulo2 = nPeriodo                                                   ' SEGUNDO TITULO DEL REPORTE
    xfrm.TamañoFuente = 6                                                     ' TAMAÑO DE LA FUENTE DEL REPORTE
    xfrm.TamañoCabecera = 8                                                   ' TAMAÑO DE LA FUENTE DE LA CABECERA DEL REPORTE
    xfrm.FuenteCabecera = "Courier New"                                       ' ESTABLECE LA FUENTE DE LA CABECERA
    xfrm.Posicion_Hoja = Horizontal                                           ' ESTABLE LA PREENTACION DE LA HOJA
    xfrm.Tamaño_Hoja = A_4                                                    ' ESTABLECE EL TAMAÑO DE LA HOJA
    
    xfrm.ImprimirArray xCampos, Rst
    Set xfrm = Nothing
    
End Sub



'***************************************************************************************************************************************

Private Sub CmdBusCliPro_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xform.Titulo = "Buscando Proveedores"
    xform.SqlCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov WHERE mae_prov.tipper = 1 ORDER BY mae_prov.nombre"
    xCampos(0, 0) = "Proveedor":       xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"

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


Private Sub txtBancarizacion_Validate(Cancel As Boolean)
    If NulosN(txtBancarizacion.Text) <> 0 Then
        txtBancarizacion.Text = Format(txtBancarizacion.Text, FORMAT_MONTO)
    Else
        txtBancarizacion.Text = "0.00"
    End If
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


Sub MostrarRegistros_xxx()
    '--dejado de usar el 15/10/10 porque esta consulta muestra datos fijos segun una columna establecida, si el usuario cambia el orden de la presentacion
    'de la consulta los datos que se presenten no coincidiran con la cabecera; El cambio consiste en hacer que la consulta se sincronice con la configuracion del reporte.
    
    'se modifica las sgtes linea de codigo
    'Form_Load , Configurar
    'SetearCuadricula Fg1, 4, xCon, 1, xFormatoActual, True
    'pImprimir
    'RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 4) And ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ")) " _
                & " ORDER BY con_formatostipodet.orden", xCon
    
    'por lo sgte
    'Form_Load , Configurar
    'SetearCuadricula Fg1, 4, xCon, 1, xFormatoActual, False
    'pImprimir
    'RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet " _
            & " Where (((con_formatostipodet.idformato) = 4) And ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ") " _
            & " And ((con_formatostipodet.mostrar) = -1)) ORDER BY con_formatostipodet.orden", xCon
       

        
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim nSQLProv As String
    
    '--verificar si hay filtro por cliente
    If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLProv = " and com_honorarios.idpro = " & NulosN(LblIdCliPro.Caption) & " "
    
    '--limpiar datos
    LblNumreg.Caption = "0"
    Fg1.Rows = 2
    DoEvents
    '----
    
    Me.MousePointer = vbHourglass
    
    If TxtIdMon.Text = 1 Then
        ' SI LA CONSULTA ES EN SOLES
        RST_Busq Rst, "SELECT Mid([com_honorarios]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_honorarios]![numreg],3,4) AS numreg, com_honorarios.fchreg, " _
            & " com_honorarios.fchdoc, com_honorarios.fchven, mae_moneda.simbolo, com_honorarios.numser, com_honorarios.numdoc, mae_dociden.codsun, mae_prov.numruc, " _
            & " mae_prov.nombre, con_tc.impven, IIf([com_honorarios].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[com_honorarios].[tc]) AS tipcam, " _
            & " IIf([com_honorarios]![idmon]=1,[impbru],[impbru]*tipcam) AS bruto, " _
            & " IIf([com_honorarios]![idmon]=1,[impigv],[impigv]*tipcam) AS impuesto, " _
            & " 0 AS otrasret, " _
            & " IIf([com_honorarios]![idmon]=1,[imptot],[imptot]*tipcam) AS total, " _
            & " com_honorarios.glosa, " _
            & " IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc]) AS numdoc2, " _
            & " Mid([com_honorarios]![numreg],1,2) AS idmes " _
            & " FROM (((mae_dociden RIGHT JOIN (com_honorarios LEFT JOIN mae_prov ON com_honorarios.idpro = mae_prov.id) ON mae_dociden.id = mae_prov.idtipdoc) " _
            & " LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) LEFT JOIN mae_moneda ON com_honorarios.idmon = mae_moneda.id) LEFT JOIN mae_libros " _
            & " ON com_honorarios.idlib = mae_libros.id WHERE (((com_honorarios.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (com_honorarios.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
            & " AND (con_tc.idmon=2 OR con_tc.idmon IS NULL) AND ((Mid([com_honorarios]![numreg],1,2))<>'00')) " & nSQLProv, xCon

    ElseIf TxtIdMon.Text = 2 Then
    
        ' SI LA CONSULTA ES EN DOLARES
        RST_Busq Rst, "SELECT Mid([com_honorarios]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_honorarios]![numreg],3,4) AS numreg, com_honorarios.fchreg, " _
            & " com_honorarios.fchdoc, com_honorarios.fchven, mae_moneda.simbolo, com_honorarios.numser, com_honorarios.numdoc, mae_dociden.codsun, mae_prov.numruc, " _
            & " mae_prov.nombre, con_tc.impven, IIf([com_honorarios].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[com_honorarios].[tc]) AS tipcam, " _
            & " IIf([com_honorarios]![idmon]=2,[impbru],IIF(tipcam =0,0,[impbru]/tipcam)) AS bruto, " _
            & " IIf([com_honorarios]![idmon]=2,[impbru],IIF(tipcam =0,0,[impigv]/tipcam)) AS impuesto, " _
            & " 0 AS otrasret, " _
            & " IIf([com_honorarios]![idmon]=2,[impbru],IIF(tipcam =0,0,[imptot]/tipcam)) AS total, " _
            & " com_honorarios.glosa, " _
            & " IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc]) AS numdoc2, " _
            & " Mid([com_honorarios]![numreg],1,2) AS idmes " _
            & " FROM (((mae_dociden RIGHT JOIN (com_honorarios LEFT JOIN mae_prov ON com_honorarios.idpro = mae_prov.id) ON mae_dociden.id = mae_prov.idtipdoc) " _
            & " LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) LEFT JOIN mae_moneda ON com_honorarios.idmon = mae_moneda.id) LEFT JOIN mae_libros " _
            & " ON com_honorarios.idlib = mae_libros.id WHERE (((com_honorarios.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (com_honorarios.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
            & " AND (con_tc.idmon=2 OR con_tc.idmon IS NULL) AND ((Mid([com_honorarios]![numreg],1,2))<>'00')) " & nSQLProv, xCon

    End If
    
    If OptOpc11.Value = True Then Rst.Filter = adFilterNone                ' mostramos todos los registros
    If OptOpc22.Value = True Then
        If TxtIdMon.Text = 1 Then Rst.Filter = "total > 3500"              ' mostramos solo los de bancarizacion en Soles
        If TxtIdMon.Text = 2 Then Rst.Filter = "total > 1000"            ' mostramos solo los de bancarizacion en Dolares
    End If
    
    If OptSort1.Value = True Then Rst.Sort = "fchdoc"                      ' ORDENA POR FECHA DE DOCUMENTO
    If OptSort2.Value = True Then Rst.Sort = "numdoc2"                     ' ORDENA POR NUMERO DE DOCUMENTO
    If OptSort3.Value = True Then Rst.Sort = "numreg"                      ' ORDENA POR NUMERO DE REGISTRO
    If OptSort4.Value = True Then Rst.Sort = "fchdoc,numdoc2"              ' ORDENA POR FECHA DE DOCUMENTO Y NUMERO DE DOCUMENTO
    
    If Rst.RecordCount <> 0 Then
        ' MOSTRAMOS LOS DATOS DEL REGISTRO DE HONORARIOS EN EL CONTROL Fg1
        LblNumreg.Caption = Rst.RecordCount
        Rst.MoveFirst
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("numreg")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Format(NulosC(Rst("fchdoc")), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(NulosC(Rst("fchven")), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(Rst("simbolo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(Rst("numser"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(Rst("numdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(Rst("codsun"))
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(Rst("numruc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(Rst("nombre"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(Rst("tipcam")), "0.000")
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(NulosN(Rst("bruto")), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(NulosN(Rst("impuesto")), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("otrasret")), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(NulosN(Rst("total")), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 15) = NulosC(Rst("glosa"))
            
            '--verificar si monto=cero y no sea anulado =>> pintar la fila para que muestre una alerta al usuario
            If NulosN(Rst("total")) = 0 And InStr(LCase(Rst("nombre")), "anulado") = 0 Then
                GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HC0C0FF
            End If
            '---------------------------------------------------------------------------------------
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    
    ' Calculando los totales
    Fg1.Rows = Fg1.Rows + 1
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H800000, False, , "TOTAL =>>"
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H800000, False, , Format(GRID_SUMAR_COL(Fg1, 11), FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H800000, False, , Format(GRID_SUMAR_COL(Fg1, 12), FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H800000, False, , Format(GRID_SUMAR_COL(Fg1, 13), FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, &H800000, False, , Format(GRID_SUMAR_COL(Fg1, 14), FORMAT_MONTO)
    
    '--restablecer cursor
    Me.MousePointer = vbDefault
End Sub




Sub ExportarPDT_BKP()
    Dim Rst As New ADODB.Recordset
    Dim NomArch, xCad As String
    Dim A As Long
    Dim B As Long
    
    If Fg1.Rows = 2 Then
        MsgBox "No se ha mostrado ninguna retencion, haga click en el boton"
    End If

    NomArch = "0601" & AnoTra & Format(TxtFchIni.Valor, "mm") & NumRUC & ".4ta"
   
    Open Trim(App.Path) + "\" + NomArch For Output As #1

'    For A = 2 To fg1.Rows - 1
'        xCad = ""
'        xCad = xCad + fg1.TextMatrix(A, 7) + "|"                        ' tipo de documento de identidad del proveedor
'        xCad = xCad + fg1.TextMatrix(A, 8) + "|"                        ' numero de documento del proveedor
'        xCad = xCad + "2" + "|"                                         ' tipo documento de la compra
'        xCad = xCad + Format(fg1.TextMatrix(A, 5), "0000") + "|"        ' numro de serioe
'        xCad = xCad + Format(fg1.TextMatrix(A, 6), "00000000") + "|"    ' numero de documento
'        xCad = xCad + Format(fg1.TextMatrix(A, 11), "0.00") + "|"       ' monto total del servicio
'        xCad = xCad + Format(fg1.TextMatrix(A, 2), "dd/mm/yyyy") + "|"  ' fecha de emision
'        xCad = xCad + Format(fg1.TextMatrix(A, 3), "dd/mm/yyyy") + "|"  ' fecha de pago
'
'        If NulosN(fg1.TextMatrix(A, 12)) = 0 Then
'            xCad = xCad + "0" + "|"  ' especifica si se aplicat retencion de 4ta
'        Else
'            xCad = xCad + "1" + "|"  ' especifica si se aplicat retencion de 4ta
'        End If
'        Print #1, Trim(xCad)
'    Next A
    
    For A = 2 To Fg1.Rows - 2
        xCad = ""
        For B = 1 To Fg1.Cols - 1
            
            Select Case LCase(Fg1.TextMatrix(1, B))
                Case "tip. doc. idem.":         xCad = xCad + Fg1.TextMatrix(A, B)                          ' tipo de documento de identidad del proveedor
                Case "nº doc ident.":           xCad = xCad + Fg1.TextMatrix(A, B)                          ' numero de documento del proveedor
                Case "tipo compra":             xCad = xCad + Fg1.TextMatrix(A, B)                          ' tipo documento de la compra
                Case "nº serie":                xCad = xCad + Format(Fg1.TextMatrix(A, B), "0000")          ' numro de serioe
                Case "nº documento":            xCad = xCad + Format(Fg1.TextMatrix(A, B), "00000000")      ' numero de documento
                Case "total honorarios":        xCad = xCad + Format(Fg1.TextMatrix(A, B), "0.00")          ' monto total del servicio
                Case "fch. doc.":               xCad = xCad + Format(Fg1.TextMatrix(A, B), "dd/mm/yyyy")    ' fecha de emision
                Case "fch. ven.":               xCad = xCad + Format(Fg1.TextMatrix(A, B), "dd/mm/yyyy")    ' fecha de pago
                Case "aplica":                  xCad = xCad + Fg1.TextMatrix(A, B)
                Case Else
                    xCad = xCad + Fg1.TextMatrix(A, B)
            End Select
            xCad = xCad + "|"
        Next B
        Print #1, Trim(xCad)
    Next A
    
    
    Close #1
    MsgBox "Los honorarios se exportaron para el PDT con exito : " & Trim(App.Path) + "\" + NomArch, vbInformation + vbOKCancel + vbDefaultButton1, xTitulo
End Sub



Private Function GenerarConsulta() As String
    '===================================================================================================
    'Creado : 16/12/10 Por: Johan Castro
    'Propósito: Generar la consulta de seleccion para mostrar en pantalla
    '
    'Entradas:  Ninguno
    '
    'Resultados: Setencia SQL listo a usar
    '
    '===================================================================================================
    
    '--
    Dim nSQL As String
    Dim nSQLFiltro As String '--Sentencia que indica el filtro a la consulta
    Dim nSQLOrden As String '--Sentencia que indica el orden de la consulta
    
    '--verificar si hay filtro por cliente
    If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLFiltro = " and com_honorarios.idpro = " & NulosN(LblIdCliPro.Caption) & " "
        
     If OptOpc22.Value = True Then
        ' mostramos solo los de bancarizacion en Soles
        If TxtIdMon.Text = 1 Then nSQLFiltro = nSQLFiltro & " and IIf([com_honorarios]![idmon]=2,[imptot],IIF(tipcam =0,0,[imptot]/tipcam)) > " & NulosN(txtBancarizacion.Text)
        ' mostramos solo los de bancarizacion en Dolares
        If TxtIdMon.Text = 2 Then nSQLFiltro = nSQLFiltro & " and IIf([com_honorarios]![idmon]=2,[imptot],IIF(tipcam =0,0,[imptot]/tipcam)) > " & NulosN(txtBancarizacion.Text)
    End If
    
   
     ' ORDENA POR FECHA DE DOCUMENTO
    If OptSort1.Value = True Then nSQLOrden = "Order by com_honorarios.fchdoc"
    ' ORDENA POR NUMERO DE DOCUMENTO
    If OptSort2.Value = True Then nSQLOrden = "Order by IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc])"
    ' ORDENA POR NUMERO DE REGISTRO
    If OptSort3.Value = True Then nSQLOrden = "Order by Mid([com_honorarios]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_honorarios]![numreg],3,4)"
    ' ORDENA POR FECHA DE DOCUMENTO Y NUMERO DE DOCUMENTO
    If OptSort4.Value = True Then nSQLOrden = "Order by com_honorarios.fchdoc,IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc])"
    
    
    If TxtIdMon.Text = 1 Then
        ' SI LA CONSULTA ES EN SOLES
        nSQL = "SELECT Mid([com_honorarios]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_honorarios]![numreg],3,4) AS registro, com_honorarios.fchreg, " _
            & " com_honorarios.fchdoc, com_honorarios.fchven, mae_moneda.simbolo, com_honorarios.numser, com_honorarios.numdoc, mae_dociden.codsun, mae_prov.numruc, " _
            & " mae_prov.nombre, con_tc.impven, IIf([com_honorarios].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[com_honorarios].[tc]) AS tipcam, " _
            & " IIf([com_honorarios]![idmon]=1,[impbru],[impbru]*tipcam) AS bruto, " _
            & " IIf([com_honorarios]![idmon]=1,[impigv],[impigv]*tipcam) AS impuesto, " _
            & " 0 AS otrasret, " _
            & " IIf([com_honorarios]![idmon]=1,[imptot],[imptot]*tipcam) AS total, " _
            & " com_honorarios.glosa, " _
            & " IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc]) AS numdoc2, " _
            & " Mid([com_honorarios]![numreg],1,2) AS idmes,mae_condpago.descripcion AS condpag, 2 as tipcom, iif(impuesto=0,0,1) as apliret " _
            & " FROM ((((mae_dociden RIGHT JOIN (com_honorarios LEFT JOIN mae_prov ON com_honorarios.idpro = mae_prov.id) ON mae_dociden.id = mae_prov.idtipdoc) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) " _
            & " LEFT JOIN mae_moneda ON com_honorarios.idmon = mae_moneda.id) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN mae_condpago ON com_honorarios.idconpag = mae_condpago.id " _
            & " WHERE (((com_honorarios.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (com_honorarios.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
            & " AND (con_tc.idmon=2 OR con_tc.idmon IS NULL) AND ((Mid([com_honorarios]![numreg],1,2))<>'00')) " & nSQLFiltro

    ElseIf TxtIdMon.Text = 2 Then
    
        ' SI LA CONSULTA ES EN DOLARES
        nSQL = "SELECT Mid([com_honorarios]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_honorarios]![numreg],3,4) AS registro, com_honorarios.fchreg, " _
            & " com_honorarios.fchdoc, com_honorarios.fchven, mae_moneda.simbolo, com_honorarios.numser, com_honorarios.numdoc, mae_dociden.codsun, mae_prov.numruc, " _
            & " mae_prov.nombre, con_tc.impven, IIf([com_honorarios].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[com_honorarios].[tc]) AS tipcam, " _
            & " IIf([com_honorarios]![idmon]=2,[impbru],IIF(tipcam =0,0,[impbru]/tipcam)) AS bruto, " _
            & " IIf([com_honorarios]![idmon]=2,[impigv],IIF(tipcam =0,0,[impigv]/tipcam)) AS impuesto, " _
            & " 0 AS otrasret, " _
            & " IIf([com_honorarios]![idmon]=2,[imptot],IIF(tipcam =0,0,[imptot]/tipcam)) AS total, " _
            & " com_honorarios.glosa, " _
            & " IIf(IsNull([com_honorarios]![numser])=-1,[com_honorarios]![numdoc],[com_honorarios]![numser] & '-' & [com_honorarios]![numdoc]) AS numdoc2, " _
            & " Mid([com_honorarios]![numreg],1,2) AS idmes,mae_condpago.descripcion AS condpag, 2 as tipcom, iif(impuesto=0,0,1) as apliret " _
            & " FROM ((((mae_dociden RIGHT JOIN (com_honorarios LEFT JOIN mae_prov ON com_honorarios.idpro = mae_prov.id) ON mae_dociden.id = mae_prov.idtipdoc) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) " _
            & " LEFT JOIN mae_moneda ON com_honorarios.idmon = mae_moneda.id) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN mae_condpago ON com_honorarios.idconpag = mae_condpago.id " _
            & " WHERE (((com_honorarios.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (com_honorarios.fchreg)<=CDate('" & TxtFchFin.Valor & "')) " _
            & " AND (con_tc.idmon=2 OR con_tc.idmon IS NULL) AND ((Mid([com_honorarios]![numreg],1,2))<>'00')) " & nSQLFiltro

    End If
    
    GenerarConsulta = nSQL
    
End Function
