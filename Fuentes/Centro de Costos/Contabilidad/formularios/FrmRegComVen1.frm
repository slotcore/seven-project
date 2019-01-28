VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRegComVen1 
   Caption         =   "Contabilidad - Registro de Compras"
   ClientHeight    =   8940
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   12315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1245
      Left            =   60
      TabIndex        =   38
      Top             =   7680
      Visible         =   0   'False
      Width           =   5010
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   120
         TabIndex        =   39
         Top             =   630
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   582
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
         TabIndex        =   40
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
            Picture         =   "FrmRegComVen1.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":277E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":2A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegComVen1.frx":2E2A
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
      Width           =   12315
      _ExtentX        =   21722
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
   Begin SizerOneLibCtl.TabOne TabOne2 
      Height          =   1275
      Left            =   0
      TabIndex        =   1
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
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   12825
         TabIndex        =   8
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
            TabIndex        =   14
            Top             =   0
            Width           =   9390
            Begin VB.OptionButton OptSel1 
               Caption         =   "Todos"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   17
               Top             =   270
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.OptionButton OptSel2 
               Caption         =   "Seleccionar"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1170
               TabIndex        =   16
               Top             =   270
               Width           =   1140
            End
            Begin VB.CommandButton CmdBusCliPro 
               Enabled         =   0   'False
               Height          =   240
               Left            =   8640
               Picture         =   "FrmRegComVen1.frx":31BC
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   210
               Width           =   210
            End
            Begin VB.TextBox TxtCliPro 
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   3405
               Locked          =   -1  'True
               TabIndex        =   18
               Text            =   "TxtCliPro"
               Top             =   180
               Width           =   5475
            End
            Begin VB.Label LblIdCliPro 
               Caption         =   "LblIdCliPro"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   2280
               TabIndex        =   20
               Top             =   150
               Visible         =   0   'False
               Width           =   750
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Proveedor"
               Height          =   195
               Index           =   2
               Left            =   2610
               TabIndex        =   19
               Top             =   270
               Width           =   735
            End
         End
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
            TabIndex        =   9
            Top             =   570
            Width           =   5085
            Begin VB.CommandButton CmdBusTipDoc 
               Height          =   240
               Left            =   735
               Picture         =   "FrmRegComVen1.frx":32EE
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   270
               Width           =   240
            End
            Begin VB.TextBox TxtTipDoc 
               Height          =   300
               Left            =   90
               MaxLength       =   3
               TabIndex        =   11
               Text            =   "TxtTipDoc"
               Top             =   240
               Width           =   915
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
               TabIndex        =   13
               Top             =   240
               Width           =   3975
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "T.D."
               Height          =   195
               Index           =   1
               Left            =   2340
               TabIndex        =   12
               Top             =   330
               Visible         =   0   'False
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1185
         Left            =   345
         TabIndex        =   2
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
            Left            =   9780
            TabIndex        =   35
            Top             =   0
            Width           =   1695
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nº Registros :"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   37
               Top             =   390
               Width           =   975
            End
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
               TabIndex        =   36
               Top             =   630
               Width           =   1440
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
            Left            =   6765
            TabIndex        =   30
            Top             =   0
            Width           =   2940
            Begin VB.OptionButton OptSort2 
               Caption         =   "Nº de Documento"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   34
               Top             =   470
               Width           =   1800
            End
            Begin VB.OptionButton OptSort1 
               Caption         =   "Fecha  de Emisión"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   33
               Top             =   240
               Width           =   1800
            End
            Begin VB.OptionButton OptSort3 
               Caption         =   "Nº Registro"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   32
               Top             =   700
               Width           =   1650
            End
            Begin VB.OptionButton OptSort4 
               Caption         =   "Fch. Emisión y Nº de Documento"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   31
               Top             =   930
               Width           =   2670
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
            TabIndex        =   25
            Top             =   0
            Width           =   2520
            Begin VB.TextBox txtBancarizacion 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1560
               TabIndex        =   41
               Text            =   "txtBancarizacion"
               Top             =   780
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.OptionButton OptOpc44 
               Caption         =   "Percepción"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   29
               Top             =   680
               Width           =   1875
            End
            Begin VB.OptionButton OptOpc33 
               Caption         =   "Detracción"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   460
               Width           =   1125
            End
            Begin VB.OptionButton OptOpc22 
               Caption         =   "Bancarización"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   27
               Top             =   900
               Width           =   1350
            End
            Begin VB.OptionButton OptOpc11 
               Caption         =   "Todas las compras"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   1920
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
            TabIndex        =   4
            Top             =   570
            Width           =   4095
            Begin VB.CommandButton CmdBusMon 
               Height          =   230
               Left            =   495
               Picture         =   "FrmRegComVen1.frx":3420
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   180
               MaxLength       =   1
               TabIndex        =   6
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
               TabIndex        =   7
               Top             =   240
               Width           =   3135
            End
         End
         Begin VB.Frame Frame5 
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
            ForeColor       =   &H00800000&
            Height          =   585
            Left            =   30
            TabIndex        =   3
            Top             =   0
            Width           =   4095
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   735
               TabIndex        =   21
               Top             =   210
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
               TabIndex        =   22
               Top             =   210
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
               Caption         =   "Desde"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   24
               Top             =   270
               Width           =   465
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
               Height          =   195
               Index           =   2
               Left            =   2145
               TabIndex        =   23
               Top             =   255
               Width           =   420
            End
         End
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6000
      Left            =   30
      TabIndex        =   42
      Top             =   1650
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
         FormatString    =   $"FrmRegComVen1.frx":3552
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
         FormatString    =   $"FrmRegComVen1.frx":378A
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
Attribute VB_Name = "FrmRegComVen1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FrmRegComVen.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : EMITE EL LIBRO REGISTRO DE COMPRAS, EN FUNCION A CRITERIOS ESPECIFICADOS POR EL
'                     USUARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 21/10/09
'* VERSION          : 1.0

'Modificado         : 08/02/10 - Johan Castro
'                     Agregar filtro por proveedor y Tipo de Documento
'                     26/04/10 - Johan Castro
'                     Agregar pestaña que muestre resumen por documento
'                     Agregar para imprimir y exportar a MSExcel el resumen por documento
'                     21/07/10 - Johan Castro
'                     Hacer que la consulta este sincronizado con la tabla con_formatostipodet
'                     en relacion al orden de los campos
'*****************************************************************************************************
Option Explicit

Dim SeEjecuto As Boolean               ' ESPECIFICA SI SE EJECUTO EL EVENTO ACTIVATE
Dim xNumPag As Integer                 ' ALMACENA EL NUMERO DE PAGINA
Dim xTotal1, xTotal2, xTotal3, xTotal4, xTotal5, xTotal6 As Double
Dim xCadOrd As String                  ' Cadenas de ordenacion para las consultas
Dim xFormatoActual As Integer          ' indica el id del formato actual que se mostrara en la cuadricula

'*****************************************************************************************************
'* Nombre           : MostrarCompras
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GENERA EL REGISTRO DE COMPRAS EN FUNCION A LAS CONDICIONES APLICADAS POR EL
'*                    USUARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MostrarCompras()
'--06-08-2010
'--01/06/11 Agregar campo baseigv en las sentencias SQL para mostrar en reporte
'--         Mostrar en pantalla el T.C. en formato a 3 decimales
'--20/12/11 Modificar consulta se registro de compras. Eliminar subconsulta de detraccion y mostrar solo campos con valores vacios, Agregar campo "Id" al final de los campos para hacer el filtro con detracc
'--         Definir Recordset RstSpot para la detraccion
'--         Escribir en grid considerando el recordset RstSpot

    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLFiltro As String '--Sentencia que indica el filtro a la consulta
    Dim nSQLCampos As String '--Relacion de campos a mostrar, obtenido de tabla: con_formatostipodet
    Dim RstSpot As New ADODB.Recordset
    
    '--obtener el orden de presentacion de los campos
    nSQLCampos = fSetearCuadriculaColumna(xCon, 2)
    '--verificar si hay campos seleccionados para mostrar el reporte
    If nSQLCampos = "" Then Exit Sub
    nSQLCampos = nSQLCampos & ",id "
    '--verificar si hay filtro por proveedor
    If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLFiltro = " and com_compras.idpro = " & NulosN(LblIdCliPro.Caption) & " "
    
    '--verificar si hay filtro por documento
    If NulosN(TxtTipDoc.Text) <> 0 Then nSQLFiltro = nSQLFiltro & " and com_compras.tipdoc = " & NulosN(TxtTipDoc.Text) & " "

    Me.MousePointer = vbHourglass
    DoEvents
    '--

    If TxtIdMon.Text = 1 Then
        
        nSQL = "SELECT com_compras.id, Left(com_compras.numreg,2)& mae_libros.codsun& Right(com_compras.numreg,4) AS registro, mae_dociden.codsun AS tdpersun, mae_prov.numruc as numruc1, mae_prov.nombre as nombre1, mae_documento.codsun AS tdsun, mae_documento.abrev, " _
            + vbCr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc2, com_compras.numser, com_compras.numdoc, com_compras.fchdoc, com_compras.fchrecep,com_compras.fchpag, mae_condpago.abrev AS condpag, com_compras.fchven, '' AS anodua, com_compras.glosa, mae_moneda.simbolo ,com_compras.tasaigv, " _
            + vbCr + " IIf(com_compras.tc=0,IIf(con_tc.impven Is Null,0,con_tc.impven),com_compras.tc) AS tipcam, IIf(com_compras.tipdoc=7,-1,1) AS xreal, " _
            + vbCr + " xreal * IIf(com_compras.idmon=1,com_compras.impbru,com_compras.impbru*tipcam)  AS impbru1_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=1,com_compras.impbru2,com_compras.impbru2*tipcam) AS impbru2_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=1,com_compras.impbru3,com_compras.impbru3*tipcam) AS impbru3_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=1,com_compras.impina,com_compras.impina*tipcam) AS impina_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=1,com_compras.impisc,com_compras.impisc*tipcam) AS impisc_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=1,com_compras.impigv,com_compras.impigv*tipcam) AS impigv1_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=1,com_compras.impigv2,com_compras.impigv2*tipcam) AS impigv2_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=1,com_compras.impigv3,com_compras.impigv3*tipcam) AS impigv3_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=1,com_compras.otroscargos,com_compras.otroscargos*tipcam) AS impotros_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=1,com_compras.imptot,com_compras.imptot*tipcam) AS imptot_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=1,com_compras.impdesc,com_compras.impdesc*tipcam) AS impdesc_c, " _
            + vbCr + " '' as cpncnd ,ref1.*, '' as spotnum, null as spotfchpag, '' as spotimp, '' as spotglosa " _
            + vbCr + " FROM ((mae_prov LEFT JOIN mae_dociden ON mae_prov.idtipdoc = mae_dociden.id) RIGHT JOIN (mae_moneda RIGHT JOIN (mae_condpago RIGHT JOIN (((com_compras LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) " _
            + vbCr + " LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_condpago.id = com_compras.idconpag) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro ) " _
            + vbCr + " LEFT JOIN " _
            + vbCr + " (SELECT com_compras_1.id AS iddoc, com_compras.id AS refiddoc, Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4) AS refregistro, mae_documento.abrev AS refabrev, mae_documento.codsun AS reftdsun, com_compras.fchdoc AS reffchdoc, com_compras.numser AS refnumser, com_compras.numdoc AS refnumdoc, mae_moneda.simbolo AS refsimbolo, com_compras.imptot AS refimptot, com_compras.glosa AS refglosa " _
            + vbCr + " FROM com_compras AS com_compras_1 INNER JOIN (((com_compras LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON com_compras.idmon = mae_moneda.id) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON com_compras_1.iddocref = com_compras.id " _
            + vbCr + " WHERE com_compras_1.iddocref<>0 ) AS ref1  ON com_compras.id=ref1.iddoc " _
            + vbCr + " WHERE (((com_compras.fchreg) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')) AND ((Mid(com_compras.numreg,1,2))<>'00')) " & nSQLFiltro
    
    Else
    
        nSQL = "SELECT com_compras.id, Left(com_compras.numreg,2)& mae_libros.codsun& Right(com_compras.numreg,4) AS registro, mae_dociden.codsun AS tdpersun, mae_prov.numruc as numruc1, mae_prov.nombre as nombre1, mae_documento.codsun AS tdsun, mae_documento.abrev, " _
            + vbCr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc2, com_compras.numser, com_compras.numdoc, com_compras.fchdoc, com_compras.fchrecep, com_compras.fchpag, mae_condpago.abrev AS condpag, com_compras.fchven, '' AS anodua, com_compras.glosa, mae_moneda.simbolo ,com_compras.tasaigv, " _
            + vbCr + " IIf(com_compras.tc=0,IIf(con_tc.impven Is Null,0,con_tc.impven),com_compras.tc) AS tipcam, IIf(com_compras.tipdoc=7,-1,1) AS xreal, " _
            + vbCr + " xreal * IIf(com_compras.idmon=2,com_compras.impbru,IIF(tipcam=0,0,com_compras.impbru/tipcam))  AS impbru1_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=2,com_compras.impbru2,IIF(tipcam=0,0,com_compras.impbru2/tipcam)) AS impbru2_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=2,com_compras.impbru3,IIF(tipcam=0,0,com_compras.impbru3/tipcam)) AS impbru3_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=2,com_compras.impina,IIF(tipcam=0,0,com_compras.impina/tipcam)) AS impina_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=2,com_compras.impisc,IIF(tipcam=0,0,com_compras.impisc/tipcam)) AS impisc_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=2,com_compras.impigv,IIF(tipcam=0,0,com_compras.impigv/tipcam)) AS impigv1_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=2,com_compras.impigv2,IIF(tipcam=0,0,com_compras.impigv2/tipcam)) AS impigv2_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=2,com_compras.impigv3,IIF(tipcam=0,0,com_compras.impigv3/tipcam)) AS impigv3_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=2,com_compras.otroscargos,IIF(tipcam=0,0,com_compras.otroscargos/tipcam)) AS impotros_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=2,com_compras.imptot,IIF(tipcam=0,0,com_compras.imptot/tipcam)) AS imptot_c, " _
            + vbCr + " xreal * IIf(com_compras.idmon=2,com_compras.impdesc,IIF(tipcam=0,0,com_compras.impdesc/tipcam)) AS impdesc_c, " _
            + vbCr + " '' as cpncnd ,ref1.*, '' as spotnum, null as spotfchpag, '' as spotimp, '' as spotglosa " _
            + vbCr + " FROM ((mae_prov LEFT JOIN mae_dociden ON mae_prov.idtipdoc = mae_dociden.id) RIGHT JOIN (mae_moneda RIGHT JOIN (mae_condpago RIGHT JOIN (((com_compras LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) " _
            + vbCr + " LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_condpago.id = com_compras.idconpag) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro ) " _
            + vbCr + " LEFT JOIN " _
            + vbCr + " (SELECT com_compras_1.id AS iddoc, com_compras.id AS refiddoc, Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4) AS refregistro, mae_documento.abrev AS refabrev, mae_documento.codsun AS reftdsun, com_compras.fchdoc AS reffchdoc, com_compras.numser AS refnumser, com_compras.numdoc AS refnumdoc, mae_moneda.simbolo AS refsimbolo, com_compras.imptot AS refimptot, com_compras.glosa AS refglosa " _
            + vbCr + " FROM com_compras AS com_compras_1 INNER JOIN (((com_compras LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON com_compras.idmon = mae_moneda.id) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON com_compras_1.iddocref = com_compras.id " _
            + vbCr + " WHERE com_compras_1.iddocref<>0 ) AS ref1  ON com_compras.id=ref1.iddoc " _
            + vbCr + " WHERE (((com_compras.fchreg) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')) AND ((Mid(com_compras.numreg,1,2))<>'00')) " & nSQLFiltro
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
    
    '**************************************************************************************************
    '--consulta para obtener constancia de Spot
    nSQL = "SELECT con_detraccion.iddoc AS spotiddoc, con_detraccion.fchpag AS spotfchpag, con_detraccion.numdet AS spotnum, con_detraccion.[imp] AS spotimp, con_detraccion.glosa AS spotglosa " _
        + vbCr + " FROM com_compras INNER JOIN con_detraccion ON com_compras.id = con_detraccion.iddoc " _
        + vbCr + " WHERE (((con_detraccion.tipo)=1) AND ((com_compras.fchreg) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) "
    
    RST_Busq RstSpot, nSQL, xCon
    '**************************************************************************************************
        
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
            Case "impigv2_c":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impigv3_c":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impotros_c":  ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "imptot_c":    ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "impdesc_c":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "spotimp":     ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            Case "refimptot":   ArrCampos(mCol) = mColCampo: mCol = mCol + 1
            
        End Select
        
    Next mColCampo
    '**************************************************************************************************
    
    If OptOpc11.Value = True Then Rst.Filter = adFilterNone                   ' mostramos todos los registros
    If OptOpc22.Value = True Then
        If TxtIdMon.Text = 1 Then Rst.Filter = "imptot_c > " & NulosN(txtBancarizacion.Text)   ' mostramos solo los de bancarizacion en Soles
        If TxtIdMon.Text = 2 Then Rst.Filter = "imptot_c > " & NulosN(txtBancarizacion.Text)   ' mostramos solo los de bancarizacion en Dolares
    End If
    
    
    '--Aplicar orden
    If OptSort1.Value = True Then Rst.Sort = "fchdoc"
    If OptSort2.Value = True Then Rst.Sort = "numdoc"
    If OptSort3.Value = True Then Rst.Sort = "registro"
    If OptSort4.Value = True Then Rst.Sort = "fchdoc,numdoc"
    
    LblNumreg.Caption = 0
    
    If OptOpc33.Value = False Then LblNumreg.Caption = Rst.RecordCount
    
    Do While Not Rst.EOF
    
        DoEvents
        
''        ProgressBar1.Value = Rst.Bookmark
        
        '-----------------------------------------------
        RstSpot.Filter = ""
        RstSpot.Filter = "spotiddoc = " & Rst("id")
        If OptOpc33.Value = True Then ' mostramos solo los detraccion
            If RstSpot.RecordCount = 0 Then GoTo AvanzaDetra
            LblNumreg.Caption = NulosN(LblNumreg.Caption) + 1
        End If
        '--------
        
        
        Fg1.Rows = Fg1.Rows + 1
        '--La última columna "id" no se considera para mostrar en pantalla
        For mCol = 0 To Rst.Fields.Count - 2
        
            Select Case LCase(Rst.Fields(mCol).Name)
                Case "fchdoc", "fchven", "fchrecep", "reffchdoc"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_DATE)
                
                Case "tipcam"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), "0.000")
                
                Case "impbru1_c", "impbru2_c", "impbru3_c", "impina_c", "impisc_c", "impigv1_c", "impigv2_c", "impigv3_c", "imptot_c", "impdesc_c", "refimptot"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), FORMAT_MONTO)
                    
                Case "impotros_c"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(Rst.Fields(mCol), "-" & FORMAT_MONTO)
                
                Case "tdpersun"
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
                    
                '--Constancia Spot--
                Case "spotfchpag"
                    If RstSpot.RecordCount <> 0 Then Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(RstSpot.Fields(LCase(Rst.Fields(mCol).Name)), FORMAT_DATE)

                Case "spotimp"
                    If RstSpot.RecordCount <> 0 Then Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = Format(RstSpot.Fields(LCase(Rst.Fields(mCol).Name)), FORMAT_MONTO)

                Case "spotnum", "spotglosa"
                    If RstSpot.RecordCount <> 0 Then Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(RstSpot.Fields(LCase(Rst.Fields(mCol).Name)))
                '----
                    
                Case Else
                    Fg1.TextMatrix(Fg1.Rows - 1, mCol + 1) = NulosC(Rst.Fields(mCol))
            End Select
            
        Next mCol
                
        '--verificar si monto=cero y no sea anulado =>> pintar la fila para que muestre una alerta al usuario
        If NulosN(Rst("imptot_c")) = 0 And InStr(LCase(Rst("nombre1")), "anulado") = 0 Then
            GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HC0C0FF
        End If
        
AvanzaDetra:
            
        Rst.MoveNext
    Loop
    
    '**************************************************************************************************
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
    '**************************************************************************************************
    
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
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, &H800000, False, , "TOTAL ==>"

    xFila = 2
    For A = 1 To Fg1.Rows - 2
        If Fg1.TextMatrix(xFila, 5) = "07" Then
            xTotal1 = xTotal1 - Abs(NulosN(Fg1.TextMatrix(xFila, 13)))
            xTotal2 = xTotal2 - Abs(NulosN(Fg1.TextMatrix(xFila, 14)))
            xTotal3 = xTotal3 - Abs(NulosN(Fg1.TextMatrix(xFila, 15)))
            xTotal4 = xTotal4 - Abs(NulosN(Fg1.TextMatrix(xFila, 16)))
            
            xIGV1 = xIGV1 - Abs(NulosN(Fg1.TextMatrix(xFila, 19)))
            xIGV2 = xIGV2 - Abs(NulosN(Fg1.TextMatrix(xFila, 20)))
            xIGV3 = xIGV3 - Abs(NulosN(Fg1.TextMatrix(xFila, 21)))
            
            xISC = xISC - Abs(NulosN(Fg1.TextMatrix(xFila, 18)))
            xOTros = xOTros - Abs(NulosN(Fg1.TextMatrix(xFila, 22)))
            xTotalTot = xTotalTot - Abs(NulosN(Fg1.TextMatrix(xFila, 23)))
        Else
            xTotal1 = xTotal1 + NulosN(Fg1.TextMatrix(xFila, 13))
            xTotal2 = xTotal2 + NulosN(Fg1.TextMatrix(xFila, 14))
            xTotal3 = xTotal3 + NulosN(Fg1.TextMatrix(xFila, 15))
            xTotal4 = xTotal4 + NulosN(Fg1.TextMatrix(xFila, 16))
            
            xIGV1 = xIGV1 + NulosN(Fg1.TextMatrix(xFila, 19))
            xIGV2 = xIGV2 + NulosN(Fg1.TextMatrix(xFila, 20))
            xIGV3 = xIGV3 + NulosN(Fg1.TextMatrix(xFila, 21))
            
            xISC = xISC + NulosN(Fg1.TextMatrix(xFila, 18))
            xOTros = xOTros + NulosN(Fg1.TextMatrix(xFila, 22))
            xTotalTot = xTotalTot + NulosN(Fg1.TextMatrix(xFila, 23))
        End If
        
        xFila = xFila + 1
    Next A
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H800000, False, , Format(xTotal1, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 14, &H800000, False, , Format(xTotal2, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 15, &H800000, False, , Format(xTotal3, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &H800000, False, , Format(xTotal4, FORMAT_MONTO)
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 19, &H800000, False, , Format(xIGV1, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &H800000, False, , Format(xIGV2, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 21, &H800000, False, , Format(xIGV3, FORMAT_MONTO)
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &H800000, False, , Format(xISC, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 22, &H800000, False, , Format(xOTros, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 23, &H800000, False, , Format(xTotalTot, FORMAT_MONTO)
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
    
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
        
        TxtCliPro.Text = ""
        
        TabOne2.CurrTab = 0
        
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
    
        txtBancarizacion.Text = "0.00"
    
        OptSort3.Value = True
        SeEjecuto = True
        TxtFchIni.SetFocus
        
        '--enfocar en la pestaña del detalle
        TabOne1.CurrTab = 0
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    LblNumreg.Caption = 0
    Dim xRs As New ADODB.Recordset
    
    ' SELECCIONAMOS EL FORMATO ACTUAL PARA LA IMPRESION DEL LIBRO "REGISTRO DE COMPRAS"
    RST_Busq xRs, "SELECT con_formatostipo.* From con_formatostipo WHERE (((con_formatostipo.defecto)=-1) AND ((con_formatostipo.idformato)=2))", xCon

    xFormatoActual = xRs("id")
    Set xRs = Nothing
    
    '--dar formato al detalle
    SetearCuadricula Fg1, 2, xCon, 1, xFormatoActual, False
    
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    OptOpc11.Value = True
    
    OptSort3.Value = True
    
    '--cargar el formato del resumen
    SetearCuadricula fg2, 2, xCon, 1, 3
    
    '--buscar los registros
    Fg1.AutoSearch = flexSearchFromTop
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
End Sub

'*****************************************************************************************************
'* Nombre           : ExportarPDT
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA AL PDT EL REGISTRO DE COMPRAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ExportarPDT()
    Dim xCol, xFil As Integer
    Dim xCad As String
    Dim NomArch As String
    NomArch = "0621" + NumRUC + AnoTra + Format(TxtFchIni.Valor, "mm") + ".txt"
    Open Trim(App.Path) + "\" + NomArch For Output As #1
    
    Dim Rst As New ADODB.Recordset
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    For xFil = 1 To Fg1.Rows - 1
    
        If NulosC(Fg1.TextMatrix(xFil, 1)) <> "" Then
            xCad = ""
            xCad = Fg1.TextMatrix(xFil, 6) + "|"
            Set Rst = BuscaConCriterio("SELECT * FROM mae_prov WHERE numruc = '" & Trim(Fg1.TextMatrix(xFil, 6)) & "'", xCon)
            xCad = xCad + Rst("apepro1") + "|"
            xCad = xCad + Rst("apepro2") + "|"
            xCad = xCad + Trim(Trim(Rst("nompro1")) + " " + Trim(Rst("nompro2"))) + "|"
            xCad = xCad + Mid(Fg1.TextMatrix(xFil, 5), 2, 3) + "|"
            xCad = xCad + Mid(Fg1.TextMatrix(xFil, 5), 8, 8) + "|"
            xCad = xCad + Format(Fg1.TextMatrix(xFil, 2), "dd/mm/yyyy") + "|"
            xCad = xCad + Fg1.TextMatrix(xFil, 9) + "|"
            If NulosN(Fg1.TextMatrix(xFil, 10)) = 0 Then
                xCad = xCad + "0" + "|"
            Else
                xCad = xCad + "1" + "|"
            End If
            xCad = xCad + "10" + "|||"
            Print #1, Trim(xCad)
        End If
    Next xFil
    
    MsgBox "Los registro de Honorarios se exporta con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Close #1
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub

    If Me.Height > 3000 Then
        TabOne1.Top = 1700
        TabOne1.Width = Me.Width - 150
        TabOne1.Height = Me.Height - 2100
    End If

End Sub

Private Sub OptOpc11_Click()
    txtBancarizacion.Visible = False
End Sub

Private Sub OptOpc22_Click()
    txtBancarizacion.Visible = True
    txtBancarizacion.SetFocus
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
        
        '--verificar que este ingresado la base para mostrar los registros que cumplen con la bancarizacion
'        If OptOpc22.Value = True Then
'            If NulosN(txtBancarizacion.Text) = 0 Then
'                MsgBox "Falta especificar la base de la bancarizacion expresado en " & LblMoneda.Caption, vbInformation, xTitulo
'                txtBancarizacion.SetFocus
'                Exit Sub
'            End If
'        End If
        
        MostrarCompras
        
        MostrarComprasResumen
    End If
    
    If Button.Index = 3 Then
        If Fg1.Rows = 2 Then
            MsgBox "No hay registro que exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Dim xFun As New SGI2_funciones.formularios
        
        If TabOne1.CurrTab = 0 Then     '--imprimir el detalle
            xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "LIBRO COMPRAS", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Por Fecha", "Registro de Compras"    ', Rst, ""
            
        Else                            '--imprimir el resumen
            xFun.VSFlexGrid_Exportar_MSExcel xCon, fg2, "RESUMEN - LIBRO COMPRAS", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Por Fecha", "Registro de Compras"   ', Rst, ""
        End If
        
        
        Set xFun = Nothing
    End If
    
    If Button.Index = 4 Then IMPRIMIR
    
    If Button.Index = 5 Then
        Configurar
    End If
    
    If Button.Index = 7 Then
        Unload Me
    End If
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
    Dim A As Long
    Dim Rst As New ADODB.Recordset
    
    Dim xCampos() As String
    Dim xFil, xCol As Double
    Dim xFila As Double
    
    ' SELECCIONAMOS EL FORMATO ACTUAL PARA LA IMPRESION DEL LIBRO "REGISTRO DE COMPRAS"
    
    If TabOne1.CurrTab = 0 Then
        '--verificar si hay registros para imprimir
        If Fg1.Rows <= Fg1.FixedRows Then
            MsgBox "No hay registros para imprimir", vbInformation, xTitulo
            Exit Sub
        End If
        
        RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet " _
            & " Where (((con_formatostipodet.idformato) = 2) And ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ") " _
            & " And ((con_formatostipodet.mostrar) = -1)) ORDER BY con_formatostipodet.orden", xCon
        
'        RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 2) And " _
'            & " ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ")) ORDER BY con_formatostipodet.orden", xCon
    
        ReDim xCampos(Fg1.Rows - 2, Fg1.Cols - 1)
        
        xFila = 0
        '--asignando el nombre de los campos
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
        
        RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 2) And " _
            & " ((con_formatostipodet.idformatotipo) = 3)) ORDER BY con_formatostipodet.orden", xCon
    
        ReDim xCampos(fg2.Rows - 2, fg2.Cols - 1)
        
        xFila = 0
        '--asignando el nombre de los campos
        For xFil = 1 To fg2.Rows - 1
            For xCol = 1 To fg2.Cols - 1
                xCampos(xFila, xCol) = fg2.TextMatrix(xFil, xCol)
            Next xCol
            xFila = xFila + 1
        Next xFil
        
    End If
    
    ' BLANQUEAMOS LAS TITULOS DE LAS COLUMNAS QUE NO SE VAN A IMPRIMIR
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
    ' CABECERA DEL REPORTE
    xfrm.Cabecera1 = NomEmp                                                   ' NOMBRE DE LA EMPRESA
    xfrm.Cabecera2 = "RUC Nº: " & NumRUC                                      ' NUMERO DE RUC DE LA EMPRESA
    xfrm.Fecha = Format(Date, "dd/mm/yyyy")                                   ' FECHA DE EMISION DEL REPORTE
    xfrm.Titulo1 = "REGISTRO DE COMPRAS " & "(Expresado en " & xMoneda & ")"  ' TITULO DEL REPORTE
    xfrm.Titulo2 = nPeriodo                                                   ' SEGUNDO TITULO DEL REPORTE
    xfrm.TamañoFuente = 6                                                     ' TAMAÑO DE LA FUENTE DEL REPORTE
    xfrm.TamañoCabecera = 8                                                   ' TAMAÑO DE LA FUENTE DE LA CABECERA DEL REPORTE
    xfrm.FuenteCabecera = "Courier New"                                       ' ESTABLECE LA FUENTE DE LA CABECERA
    
    If TabOne1.CurrTab = 0 Then                     ' ESPECIFICA LA ORIENTACION DE LA JOHA
        xfrm.Posicion_Hoja = Horizontal '--detalle
    Else
        xfrm.Posicion_Hoja = Vertical   '--resumen
    End If
    
    xfrm.Tamaño_Hoja = A_4                                                    ' ESTABLECE EL TAMAÑO DE LA HOJA
    xfrm.ImprimirArray xCampos, Rst
    Set xfrm = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Configurar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL REGISTRO DE COMPRAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Configurar()
    Dim xform As New SGI2_funciones.Varias
    If xform.CambioOpcionLiro(2, xCon, 1) = True Then
    
        Dim xRs As New ADODB.Recordset
        ' SELECCIONAMOS EL FORMATO ACTUAL PARA LA IMPRESION DEL LIBRO "REGISTRO DE COMPRAS"
        RST_Busq xRs, "SELECT con_formatostipo.* From con_formatostipo WHERE (((con_formatostipo.defecto)=-1) AND ((con_formatostipo.idformato)=3))", xCon
    
        xFormatoActual = xRs("id")
        
        Set xRs = Nothing
        
        SetearCuadricula Fg1, 2, xCon, 1, xFormatoActual, False
            
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
        MostrarCompras
    End If
    Set xform = Nothing
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
    & " FROM com_compras INNER JOIN mae_documento ON com_compras.tipdoc = mae_documento.id " _
    & " WHERE (((com_compras.fchreg) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " _
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
        & " FROM com_compras INNER JOIN mae_documento ON com_compras.tipdoc = mae_documento.id " _
        & " WHERE (((com_compras.fchreg) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) and com_compras.tipdoc = " & NulosN(TxtTipDoc.Text) & " " _
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





'***************************************************************************************************************************************

Private Sub CmdBusCliPro_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xform.Titulo = "Buscando Proveedores"
    xform.SqlCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov ORDER BY mae_prov.nombre"
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
'* Nombre           : MostrarComprasResumen
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GENERA EL RESUMEN DEL REGISTRO DE COMPRAS EN FUNCION A LAS CONDICIONES APLICADAS POR EL
'*                    USUARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MostrarComprasResumen()
    Dim Rst As New ADODB.Recordset
    Dim SqlCad As String
    Dim A As Long
    Dim nSQLFiltro As String '--Sentencia que indica el filtro a la consulta

    '--verificar si se puede mostrar los datos, esto dependera que esta la configuracion del grid en la base de datos
    If fg2.Cols = 1 Then
        Exit Sub
    End If
    
    '--verificar si hay filtro por cliente
    If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLFiltro = " and com_compras.idpro = " & NulosN(LblIdCliPro.Caption) & " "
    
    '--verificar si hay filtro por documento
    If NulosN(TxtTipDoc.Text) <> 0 Then nSQLFiltro = nSQLFiltro & " and com_compras.tipdoc = " & NulosN(TxtTipDoc.Text) & " "
    
    
    Me.MousePointer = vbHourglass
    DoEvents
    '--
    
    If TxtIdMon.Text = 1 Then
        ' SI EL REGISTRO DE COMPRAS SE VISUALIZA EN SOLES
        SqlCad = "SELECT CONSULTA1.tdocnombre, CONSULTA1.abrev, CONSULTA1.tipdoc, Sum(CONSULTA1.impbru_c) AS impbru_c1, Sum(CONSULTA1.impbru2_c) AS impbru2_c1, Sum(CONSULTA1.impbru3_c) AS impbru3_c1, " _
            & vbCr & " Sum(CONSULTA1.impina_c) AS impina_c1, Sum(CONSULTA1.impisc_c) AS impisc_c1, Sum(CONSULTA1.impigv_c) AS impigv_c1, Sum(CONSULTA1.impigv2_c) AS impigv2_c1, Sum(CONSULTA1.impigv3_c) AS impigv3_c1, Sum(CONSULTA1.otros_c) AS otros_c1, Sum(CONSULTA1.imptot_c) AS imptot_c1 " _
            & " FROM " _
            & vbCr & " (SELECT com_compras.id, Mid(com_compras!numreg,1,2)+mae_libros!codsun+Mid(com_compras!numreg,3,4) AS numeg, com_compras.fchdoc, com_compras.numser+'-'+com_compras.numdoc as numdoc2, " _
            & " com_compras.fchven, com_compras.fchreg, '' AS anodua, mae_documento.codsun AS tipdoc, com_compras.numser, com_compras.numdoc, mae_dociden.codsun AS tdi, mae_prov.numruc, " _
            & vbCr & " mae_prov.nombre, con_tc.impven , IIf([com_compras].[tc]=0,IIF([con_tc].[impven] is null,0,[con_tc].[impven]),[com_compras].[tc]) AS tipcam, mae_moneda.simbolo AS moneda, con_tc.idmon, " _
            & vbCr & " FORMAT(IIf([com_compras]![idmon]=1,[impbru],[com_compras]![impbru]*tipcam),'0.00') AS impbru_c, " _
            & vbCr & " FORMAT(IIf([com_compras]![idmon]=1,[impbru2],[com_compras]![impbru2]*tipcam),'0.00') AS impbru2_c, " _
            & vbCr & " FORMAT(IIf([com_compras]![idmon]=1,[impbru3],[com_compras]![impbru3]*tipcam),'0.00') AS impbru3_c, " _
            & vbCr & " FORMAT(IIf([com_compras]![idmon]=1,[impina],[com_compras]![impina]*tipcam),'0.00') AS impina_c, " _
            & vbCr & " FORMAT(IIf([com_compras]![idmon]=1,[impisc],[com_compras]![impisc]*tipcam),'0.00') AS impisc_c,  " _
            & vbCr & " FORMAT(IIf([com_compras]![idmon]=1,[impigv],[com_compras]![impigv]*tipcam),'0.00') AS impigv_c, " _
            & vbCr & " FORMAT(IIf([com_compras]![idmon]=1,[impigv2],[com_compras]![impigv2]*tipcam),'0.00') AS impigv2_c, " _
            & vbCr & " FORMAT(IIf([com_compras]![idmon]=1,[impigv3],[com_compras]![impigv3]*tipcam),'0.00') AS impigv3_c, " _
            & vbCr & " FORMAT(IIf([com_compras]![idmon]=1,[otroscargos],[com_compras]![otroscargos]*tipcam),'0.00') AS otros_c, " _
            & vbCr & " FORMAT(IIf([com_compras]![idmon]=1,[imptot],[com_compras]![imptot]*tipcam),'0.00') AS imptot_c, " _
            & vbCr & " mae_documento.descripcion AS tdocnombre, mae_documento.abrev " _
            & vbCr & " FROM (mae_prov LEFT JOIN mae_dociden ON mae_prov.idtipdoc = mae_dociden.id) RIGHT JOIN (mae_moneda RIGHT JOIN (((com_compras LEFT JOIN mae_documento " _
            & " ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc " _
            & " ON com_compras.fchdoc = con_tc.fecha) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            & vbCr & " WHERE ( con_tc.idmon=2 OR con_tc.idmon IS NULL) AND ( com_compras.fchreg >=CDate('" & TxtFchIni.Valor & "') And com_compras.fchreg <=CDate('" & TxtFchFin.Valor & "') ) AND Mid(com_compras.numreg,1,2)<>'00' " & nSQLFiltro & ") AS CONSULTA1 " _
            & vbCr & " GROUP BY CONSULTA1.tdocnombre, CONSULTA1.abrev, CONSULTA1.tipdoc " _
            & vbCr & " ORDER BY CONSULTA1.tipdoc "
            
    
    ElseIf TxtIdMon.Text = 2 Then
    
        ' SI EL REGISTRO DE COMPRAS SE VISUALIZA EN DOLARES
        SqlCad = "SELECT CONSULTA1.tdocnombre, CONSULTA1.abrev, CONSULTA1.tipdoc, Sum(CONSULTA1.impbru_c) AS impbru_c1, Sum(CONSULTA1.impbru2_c) AS impbru2_c1, Sum(CONSULTA1.impbru3_c) AS impbru3_c1, " _
            & vbCr & " Sum(CONSULTA1.impina_c) AS impina_c1, Sum(CONSULTA1.impisc_c) AS impisc_c1, Sum(CONSULTA1.impigv_c) AS impigv_c1, Sum(CONSULTA1.impigv2_c) AS impigv2_c1, Sum(CONSULTA1.impigv3_c) AS impigv3_c1, Sum(CONSULTA1.otros_c) AS otros_c1, Sum(CONSULTA1.imptot_c) AS imptot_c1 " _
            & " FROM " _
            & vbCr & "(SELECT com_compras.id, Mid(com_compras!numreg,1,2)+mae_libros!codsun+Mid(com_compras!numreg,3,4) AS numeg, com_compras.fchdoc, com_compras.numser+'-'+com_compras.numdoc as numdoc2, " _
            & " com_compras.fchven, com_compras.fchreg , '' AS anodua, mae_documento.codsun AS tipdoc, com_compras.numser, com_compras.numdoc, mae_dociden.codsun AS tdi, mae_prov.numruc, " _
            & " mae_prov.nombre, con_tc.impven, IIf([com_compras].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[com_compras].[tc]) AS tipcam, mae_moneda.simbolo AS moneda, con_tc.idmon, " _
            & vbCr & " FORMAT(IIF([com_compras]![idmon]=2,[impbru], IIF(tipcam=0,0,[com_compras]![impbru]/tipcam)),'0.00') AS impbru_c, " _
            & vbCr & " FORMAT(IIF([com_compras]![idmon]=2,[impbru2],IIF(tipcam=0,0,[com_compras]![impbru2]/tipcam)),'0.00') AS impbru2_c, " _
            & vbCr & " FORMAT(IIF([com_compras]![idmon]=2,[impbru3],IIF(tipcam=0,0,[com_compras]![impbru3]/tipcam)),'0.00') AS impbru3_c, " _
            & vbCr & " FORMAT(IIF([com_compras]![idmon]=2,[impina],IIF(tipcam=0,0,[com_compras]![impina]/tipcam)),'0.00') AS impina_c, " _
            & vbCr & " FORMAT(IIF([com_compras]![idmon]=2,[impisc],IIF(tipcam=0,0,[com_compras]![impisc]/tipcam)),'0.00') AS impisc_c,  " _
            & vbCr & " FORMAT(IIF([com_compras]![idmon]=2,[impigv],IIF(tipcam=0,0,[com_compras]![impigv]/tipcam)),'0.00') AS impigv_c, " _
            & vbCr & " FORMAT(IIF([com_compras]![idmon]=2,[impigv2],IIF(tipcam=0,0,[com_compras]![impigv2]/tipcam)),'0.00') AS impigv2_c, " _
            & vbCr & " FORMAT(IIF([com_compras]![idmon]=2,[impigv3],IIF(tipcam=0,0,[com_compras]![impigv3]/tipcam)),'0.00') AS impigv3_c, " _
            & vbCr & " FORMAT(IIF([com_compras]![idmon]=2,[otroscargos],IIF(tipcam=0,0,[com_compras]![otroscargos]/tipcam)),'0.00') AS otros_c, " _
            & vbCr & " FORMAT(IIF([com_compras]![idmon]=2,[imptot],IIF(tipcam=0,0,[com_compras]![imptot]/tipcam)),'0.00') AS imptot_c, " _
            & vbCr & " mae_documento.descripcion AS tdocnombre, mae_documento.abrev " _
            & vbCr & " FROM (mae_prov LEFT JOIN mae_dociden ON mae_prov.idtipdoc = mae_dociden.id) RIGHT JOIN (mae_moneda RIGHT JOIN (((com_compras LEFT JOIN mae_documento " _
            & " ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc " _
            & " ON com_compras.fchdoc = con_tc.fecha) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            & vbCr & " WHERE ( con_tc.idmon=2 OR con_tc.idmon IS NULL) AND ( com_compras.fchreg >=CDate('" & TxtFchIni.Valor & "') And com_compras.fchreg <=CDate('" & TxtFchFin.Valor & "') ) AND Mid(com_compras.numreg,1,2)<>'00' " & nSQLFiltro & ") AS CONSULTA1 " _
            & vbCr & " GROUP BY CONSULTA1.tdocnombre, CONSULTA1.abrev, CONSULTA1.tipdoc " _
            & vbCr & " ORDER BY CONSULTA1.tipdoc "
            
    End If

    
    RST_Busq Rst, SqlCad, xCon
    
    '--salir si hay error en la consulta
    If Rst.State = 0 Then GoTo LaCague

    If Rst.RecordCount <> 0 Then
        ' IMPRIMIMOS LOS DATOS DEL RECORDSET EN EL CONTROL Fg1
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            fg2.Rows = fg2.Rows + 1
            fg2.TextMatrix(fg2.Rows - 1, 1) = NulosC(Rst("tdocnombre"))
            fg2.TextMatrix(fg2.Rows - 1, 2) = NulosC(Rst("abrev"))
            '--VERIFICAMOS SI ES NOTA DE CREDITO
            If NulosC(Rst("tipdoc")) = "07" Then
                fg2.TextMatrix(fg2.Rows - 1, 3) = Format(NulosN(Rst("impbru_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 4) = Format(NulosN(Rst("impbru2_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 5) = Format(NulosN(Rst("impbru3_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 6) = Format(NulosN(Rst("impina_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 7) = "-0.00" 'descuentos obtenidos
                fg2.TextMatrix(fg2.Rows - 1, 8) = Format(NulosN(Rst("impisc_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 9) = Format(NulosN(Rst("impigv_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 10) = Format(NulosN(Rst("impigv2_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 11) = Format(NulosN(Rst("impigv3_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 12) = Format(NulosN(Rst("otros_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 13) = Format(NulosN(Rst("imptot_c1")), "-" & FORMAT_MONTO)
            Else
                fg2.TextMatrix(fg2.Rows - 1, 3) = Format(NulosN(Rst("impbru_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 4) = Format(NulosN(Rst("impbru2_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 5) = Format(NulosN(Rst("impbru3_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 6) = Format(NulosN(Rst("impina_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 7) = "0.00" 'descuentos obtenidos
                fg2.TextMatrix(fg2.Rows - 1, 8) = Format(NulosN(Rst("impisc_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 9) = Format(NulosN(Rst("impigv_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 10) = Format(NulosN(Rst("impigv2_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 11) = Format(NulosN(Rst("impigv3_c1")), FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 12) = Format(NulosN(Rst("otros_c1")), "-" & FORMAT_MONTO)
                fg2.TextMatrix(fg2.Rows - 1, 13) = Format(NulosN(Rst("imptot_c1")), FORMAT_MONTO)
            End If
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    
    '--mostrar totales
    fg2.Rows = fg2.Rows + 1
    Dim xCol As Integer
    FORMATO_CELDA fg2, fg2.Rows - 1, 1, &H800000, False, , "TOTAL ==>"
    For xCol = 3 To 13
        FORMATO_CELDA fg2, fg2.Rows - 1, xCol, &H800000, False, , Format(NulosN(GRID_SUMAR_COL(fg2, xCol, fg2.FixedCols, fg2.Rows - 2)), FORMAT_MONTO)
    Next
    '------
LaCague:
    Set Rst = Nothing
        
    '--restablecer cursor
    Me.MousePointer = vbDefault
    
End Sub




Sub MostrarCompras_xxxx()
    '--dejado de usar el 15/10/10 porque esta consulta muestra datos fijos segun una columna establecida, si el usuario cambia el orden de la presentacion
    'de la consulta los datos que se presenten no coincidiran con la cabecera; El cambio consiste en hacer que la consulta se sincronice con la configuracion del reporte.
    
    'se modifica las sgtes linea de codigo
    'Form_Load , Configurar
    'SetearCuadricula Fg1, 2, xCon, 1, xFormatoActual, True
    'pImprimir
    'RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet Where (((con_formatostipodet.idformato) = 2) And ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ")) " _
                & " ORDER BY con_formatostipodet.orden", xCon
    
    'por lo sgte
    'Form_Load , Configurar
    'SetearCuadricula Fg1, 2, xCon, 1, xFormatoActual, False
    'pImprimir
    'RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet " _
            & " Where (((con_formatostipodet.idformato) = 2) And ((con_formatostipodet.idformatotipo) = " & xFormatoActual & ") " _
            & " And ((con_formatostipodet.mostrar) = -1)) ORDER BY con_formatostipodet.orden", xCon
       
    
    Dim Rst As New ADODB.Recordset
    Dim SqlCad As String
    Dim A As Integer
    
    Dim nSQLProv As String
    Dim nSQLTipDoc As String
    
    '--verificar si hay filtro por cliente
    If NulosN(LblIdCliPro.Caption) <> 0 Then nSQLProv = " and com_compras.idpro = " & NulosN(LblIdCliPro.Caption) & " "
    
    '--verificar si hay filtro por documento
    If NulosN(TxtTipDoc.Text) <> 0 Then nSQLTipDoc = " and com_compras.tipdoc = " & NulosN(TxtTipDoc.Text) & " "
    
    '--limpiar datos
    Fg1.Rows = 2
    LblNumreg.Caption = 0
    Me.MousePointer = vbHourglass
    DoEvents
    '--
    
    If TxtIdMon.Text = 1 Then
        ' SI EL REGISTRO DE COMPRAS SE VISUALIZA EN SOLES
        SqlCad = "SELECT CONSULTA1.*, CONSULTA2.spotnum, CONSULTA2.spotfecha, CONSULTA3.factipdoc, CONSULTA3.facfchdoc, CONSULTA3.facnumser, CONSULTA3.facnumdoc" _
            & " FROM "
        SqlCad = SqlCad + " ((SELECT com_compras.id, Mid(com_compras!numreg,1,2)+mae_libros!codsun+Mid(com_compras!numreg,3,4) AS numeg, com_compras.fchdoc, com_compras.numser+'-'+com_compras.numdoc as numdoc2, " _
            & " com_compras.fchven, '' AS anodua, mae_documento.codsun AS td, com_compras.numser, com_compras.numdoc, mae_dociden.codsun AS tdi, mae_prov.numruc, " _
            & " mae_prov.nombre, con_tc.impven , IIf([com_compras].[tc]=0,IIF([con_tc].[impven] is null,0,[con_tc].[impven]),[com_compras].[tc]) AS tipcam, mae_moneda.simbolo AS moneda, con_tc.idmon, " _
            & " IIf([com_compras]![idmon]=1,[impbru],[com_compras]![impbru]*tipcam) AS impbru_c, " _
            & " IIf([com_compras]![idmon]=1,[impbru2],[com_compras]![impbru2]*tipcam) AS impbru2_c, " _
            & " IIf([com_compras]![idmon]=1,[impbru3],[com_compras]![impbru3]*tipcam) AS impbru3_c, " _
            & " IIf([com_compras]![idmon]=1,[impina],[com_compras]![impina]*tipcam) AS impina_c, " _
            & " IIf([com_compras]![idmon]=1,[impisc],[com_compras]![impisc]*tipcam) AS impisc_c,  " _
            & " IIf([com_compras]![idmon]=1,[impigv],[com_compras]![impigv]*tipcam) AS impigv_c, " _
            & " IIf([com_compras]![idmon]=1,[impigv2],[com_compras]![impigv2]*tipcam) AS impigv2_c, " _
            & " IIf([com_compras]![idmon]=1,[impigv3],[com_compras]![impigv3]*tipcam) AS impigv3_c, " _
            & " IIf([com_compras]![idmon]=1,[otroscargos],[com_compras]![otroscargos]*tipcam) AS otros_c, " _
            & " IIf([com_compras]![idmon]=1,[imptot],[com_compras]![imptot]*tipcam) AS imptot_c, com_compras.fchreg " _
            & " FROM (mae_prov LEFT JOIN mae_dociden ON mae_prov.idtipdoc = mae_dociden.id) RIGHT JOIN (mae_moneda RIGHT JOIN (((com_compras LEFT JOIN mae_documento " _
            & " ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc " _
            & " ON com_compras.fchdoc = con_tc.fecha) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            & " WHERE ( con_tc.idmon=2 OR con_tc.idmon IS NULL)  " & nSQLTipDoc & nSQLProv & ") AS CONSULTA1 " _
            & " LEFT JOIN " _
            & " (SELECT con_detraccion.iddoc, con_detraccion.numdet AS spotnum, con_detraccion.fchpag AS spotfecha FROM con_detraccion " _
            & " WHERE (((con_detraccion.tipo)=1))) AS CONSULTA2 ON CONSULTA1.id = CONSULTA2.iddoc) " _
            & " LEFT JOIN " _
            & " (SELECT com_compras.id, mae_documento.codsun AS factipdoc, com_compras_1.fchdoc AS facfchdoc, com_compras_1.numser AS facnumser, " _
            & " com_compras_1.numdoc AS facnumdoc FROM com_compras LEFT JOIN (com_compras AS com_compras_1 LEFT JOIN mae_documento " _
            & " ON com_compras_1.tipdoc = mae_documento.id) ON com_compras.iddocref = com_compras_1.id "
            
        SqlCad = SqlCad + " WHERE (((com_compras.tipdoc)=7))) AS CONSULTA3 ON CONSULTA1.id = CONSULTA3.id " _
            & " WHERE (((CONSULTA1.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (CONSULTA1.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((Mid([numeg],1,2))<>'00'));"
    
    ElseIf TxtIdMon.Text = 2 Then
    
        ' SI EL REGISTRO DE COMPRAS SE VISUALIZA EN DOLARES
        SqlCad = "SELECT CONSULTA1.*, CONSULTA2.spotnum, CONSULTA2.spotfecha, CONSULTA3.factipdoc, CONSULTA3.facfchdoc, CONSULTA3.facnumser, CONSULTA3.facnumdoc" _
            & " FROM "
        
        SqlCad = SqlCad + "((SELECT com_compras.id, Mid(com_compras!numreg,1,2)+mae_libros!codsun+Mid(com_compras!numreg,3,4) AS numeg, com_compras.fchdoc, com_compras.numser+'-'+com_compras.numdoc as numdoc2, " _
            & " com_compras.fchven, '' AS anodua, mae_documento.codsun AS td, com_compras.numser, com_compras.numdoc, mae_dociden.codsun AS tdi, mae_prov.numruc, " _
            & " mae_prov.nombre, con_tc.impven, IIf([com_compras].[tc]=0,IIF([con_tc].[impven] IS NULL,0,[con_tc].[impven]),[com_compras].[tc]) AS tipcam, mae_moneda.simbolo AS moneda, con_tc.idmon, " _
            & " IIF([com_compras]![idmon]=2,[impbru], IIF(tipcam=0,0,[com_compras]![impbru]/tipcam)) AS impbru_c, " _
            & " IIF([com_compras]![idmon]=2,[impbru2],IIF(tipcam=0,0,[com_compras]![impbru2]/tipcam)) AS impbru2_c, " _
            & " IIF([com_compras]![idmon]=2,[impbru3],IIF(tipcam=0,0,[com_compras]![impbru3]/tipcam)) AS impbru3_c, " _
            & " IIF([com_compras]![idmon]=2,[impina],IIF(tipcam=0,0,[com_compras]![impina]/tipcam)) AS impina_c, " _
            & " IIF([com_compras]![idmon]=2,[impisc],IIF(tipcam=0,0,[com_compras]![impisc]/tipcam)) AS impisc_c,  " _
            & " IIF([com_compras]![idmon]=2,[impigv],IIF(tipcam=0,0,[com_compras]![impigv]/tipcam)) AS impigv_c, " _
            & " IIF([com_compras]![idmon]=2,[impigv2],IIF(tipcam=0,0,[com_compras]![impigv2]/tipcam)) AS impigv2_c, " _
            & " IIF([com_compras]![idmon]=2,[impigv3],IIF(tipcam=0,0,[com_compras]![impigv3]/tipcam)) AS impigv3_c, " _
            & " IIF([com_compras]![idmon]=2,[otroscargos],IIF(tipcam=0,0,[com_compras]![otroscargos]/tipcam)) AS otros_c, " _
            & " IIF([com_compras]![idmon]=2,[imptot],IIF(tipcam=0,0,[com_compras]![imptot]/tipcam)) AS imptot_c, com_compras.fchreg " _
            & " FROM (mae_prov LEFT JOIN mae_dociden ON mae_prov.idtipdoc = mae_dociden.id) RIGHT JOIN (mae_moneda RIGHT JOIN (((com_compras LEFT JOIN mae_documento " _
            & " ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc " _
            & " ON com_compras.fchdoc = con_tc.fecha) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            & " WHERE (con_tc.idmon=2 OR con_tc.idmon IS NULL) " & nSQLTipDoc & nSQLProv & ") AS CONSULTA1 " _
            & " LEFT JOIN " _
            & " (SELECT con_detraccion.iddoc, con_detraccion.numdet AS spotnum, con_detraccion.fchpag AS spotfecha FROM con_detraccion " _
            & " WHERE (((con_detraccion.tipo)=1))) AS CONSULTA2 ON CONSULTA1.id = CONSULTA2.iddoc) " _
            & " LEFT JOIN " _
            & " (SELECT com_compras.id, mae_documento.codsun AS factipdoc, com_compras_1.fchdoc AS facfchdoc, com_compras_1.numser AS facnumser, " _
            & " com_compras_1.numdoc AS facnumdoc FROM com_compras LEFT JOIN (com_compras AS com_compras_1 LEFT JOIN mae_documento " _
            & " ON com_compras_1.tipdoc = mae_documento.id) ON com_compras.iddocref = com_compras_1.id "
        SqlCad = SqlCad + "WHERE (((com_compras.tipdoc)=7))) AS CONSULTA3 ON CONSULTA1.id = CONSULTA3.id " _
            & " WHERE (((CONSULTA1.fchreg)>=CDate('" & TxtFchIni.Valor & "') And (CONSULTA1.fchreg)<=CDate('" & TxtFchFin.Valor & "')) AND ((Mid([numeg],1,2))<>'00'));"
    End If

    
    RST_Busq Rst, SqlCad, xCon
    
    If OptOpc11.Value = True Then Rst.Filter = adFilterNone                   ' mostramos todos los registros
    If OptOpc22.Value = True Then
        If TxtIdMon.Text = 1 Then Rst.Filter = "imptot_c > 3500"           ' mostramos solo los de bancarizacion en Soles
        If TxtIdMon.Text = 2 Then Rst.Filter = "imptot_c > 1000"           ' mostramos solo los de bancarizacion en Dolares
    End If
    If OptOpc33.Value = True Then Rst.Filter = "spotnum<>null"                ' mostramos solo los detraccion
    
    If OptSort1.Value = True Then Rst.Sort = "fchdoc"
    If OptSort2.Value = True Then Rst.Sort = "numdoc2"
    If OptSort3.Value = True Then Rst.Sort = "numeg"
    If OptSort4.Value = True Then Rst.Sort = "fchdoc,numdoc"
    
    If Rst.RecordCount <> 0 Then
        ' IMPRIMIMOS LOS DATOS DEL RECORDSET EN EL CONTROL Fg1
        LblNumreg.Caption = Rst.RecordCount
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("numeg")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Format(Rst("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(NulosC(Rst("fchven")), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(Rst("anodua"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(Rst("td"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(Rst("numser"))
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Rst("numdoc")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(Rst("tdi"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(Rst("numruc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Rst("nombre")
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(NulosN(Rst("tipcam")), "0.000")
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosC(Rst("moneda"))
            
            If Fg1.TextMatrix(Fg1.Rows - 1, 5) = "07" Then
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("impbru_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(NulosN(Rst("impbru2_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(NulosN(Rst("impbru3_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(NulosN(Rst("impina_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 17) = "-0.00" 'descuentos obtenidos
                Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(NulosN(Rst("impisc_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(NulosN(Rst("impigv_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(NulosN(Rst("impigv2_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(NulosN(Rst("impigv3_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 22) = Format(NulosN(Rst("otros_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 23) = Format(NulosN(Rst("imptot_c")), "-" & FORMAT_MONTO)
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(NulosN(Rst("impbru_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(NulosN(Rst("impbru2_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(NulosN(Rst("impbru3_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(NulosN(Rst("impina_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 17) = "0.00" 'descuentos obtenidos
                Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(NulosN(Rst("impisc_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(NulosN(Rst("impigv_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(NulosN(Rst("impigv2_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(NulosN(Rst("impigv3_c")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 22) = Format(NulosN(Rst("otros_c")), "-" & FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 23) = Format(NulosN(Rst("imptot_c")), FORMAT_MONTO)
            End If
            Fg1.TextMatrix(Fg1.Rows - 1, 24) = "  " 'CP y NC y ND Sujetas a Retencion
            Fg1.TextMatrix(Fg1.Rows - 1, 25) = NulosC(Rst("spotnum"))
            Fg1.TextMatrix(Fg1.Rows - 1, 26) = Format(NulosC(Rst("spotfecha")), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 27) = Format(NulosC(Rst("facfchdoc")), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 28) = Format(NulosC(Rst("factipdoc")), "00")
            Fg1.TextMatrix(Fg1.Rows - 1, 29) = NulosC(Rst("facnumser"))
            Fg1.TextMatrix(Fg1.Rows - 1, 30) = NulosC(Rst("facnumdoc"))
            
            '--verificar si monto=cero y no sea anulado =>> pintar la fila para que muestre una alerta al usuario
            If NulosN(Rst("imptot_c")) = 0 And InStr(LCase(Rst("nombre")), "anulado") = 0 Then
                GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HC0C0FF
            End If
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    SumarColumna
    
    '--restablecer cursor
    Me.MousePointer = vbDefault
    
End Sub



