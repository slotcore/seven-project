VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCtaCte3 
   Caption         =   "Caja y Bancos - Cuenta Corriente (Cliente, Proveedor)"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne2 
      Height          =   1545
      Left            =   30
      TabIndex        =   5
      Top             =   360
      Width           =   11850
      _cx             =   20902
      _cy             =   2725
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
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1455
         Left            =   345
         TabIndex        =   7
         Top             =   45
         Width           =   11460
         Begin VB.Frame Frame3 
            Caption         =   "[  Seleccionar  ]"
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
            Height          =   900
            Left            =   4110
            TabIndex        =   8
            Top             =   -30
            Width           =   5100
            Begin VB.CommandButton CmdBusCliPro 
               Enabled         =   0   'False
               Height          =   240
               Left            =   4770
               Picture         =   "FrmCtaCte3.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   480
               Width           =   210
            End
            Begin VB.OptionButton OptSel2 
               Caption         =   "Seleccionar"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1500
               TabIndex        =   10
               Top             =   240
               Width           =   1140
            End
            Begin VB.OptionButton OptSel1 
               Caption         =   "Todos"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   9
               Top             =   240
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.TextBox TxtCliPro 
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   165
               Locked          =   -1  'True
               TabIndex        =   27
               Text            =   "TxtCliPro"
               Top             =   450
               Width           =   4845
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cliente"
               Height          =   195
               Index           =   0
               Left            =   2580
               TabIndex        =   29
               Top             =   210
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Label LblIdCliPro 
               Caption         =   "LblIdCliPro"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   3090
               TabIndex        =   28
               Top             =   240
               Visible         =   0   'False
               Width           =   750
            End
         End
         Begin VB.Frame Frame12 
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
            Height          =   900
            Left            =   30
            TabIndex        =   35
            Top             =   -30
            Width           =   1515
            Begin VB.OptionButton OptFch 
               Caption         =   "Fch Reg"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   37
               Top             =   435
               Value           =   -1  'True
               Width           =   1125
            End
            Begin VB.OptionButton OptFch 
               Caption         =   "Fch Doc"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   36
               Top             =   210
               Width           =   1140
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Tipo Reporte ]"
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
            Height          =   900
            Left            =   1620
            TabIndex        =   34
            Top             =   -30
            Width           =   2445
            Begin VB.OptionButton opt4ta 
               Caption         =   "Honorarios"
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
               Left            =   60
               TabIndex        =   40
               Top             =   660
               Width           =   2160
            End
            Begin VB.OptionButton OptProvee 
               Caption         =   "Proveedor"
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
               Left            =   60
               TabIndex        =   39
               Top             =   435
               Width           =   1230
            End
            Begin VB.OptionButton OptCliente 
               Caption         =   "Cliente"
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
               Left            =   60
               TabIndex        =   38
               Top             =   210
               Value           =   -1  'True
               Width           =   960
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "[  Seleccionar Estado ]"
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
            Height          =   1485
            Left            =   9240
            TabIndex        =   30
            Top             =   -30
            Width           =   2220
            Begin VB.CheckBox chk_descuadrado 
               Caption         =   "Descuadrados"
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
               Left            =   150
               TabIndex        =   41
               ToolTipText     =   "Mostrar� solo los documentos cuyo saldo final es negativo"
               Top             =   1140
               Value           =   1  'Checked
               Width           =   1545
            End
            Begin VB.OptionButton OptPen 
               Caption         =   "Pendientes"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   210
               Value           =   -1  'True
               Width           =   1110
            End
            Begin VB.OptionButton OptCan 
               Caption         =   "Cancelados"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   32
               Top             =   465
               Width           =   1350
            End
            Begin VB.OptionButton OptTodos 
               Caption         =   "Todos"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   720
               Width           =   900
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00FFFFFF&
               X1              =   120
               X2              =   2070
               Y1              =   1040
               Y2              =   1040
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   2070
               Y1              =   1020
               Y2              =   1020
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "[Seleccionar Fecha]"
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
            TabIndex        =   21
            Top             =   870
            Width           =   4035
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   705
               TabIndex        =   22
               Top             =   225
               Width           =   1245
               _ExtentX        =   2196
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
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
               Height          =   300
               Left            =   2670
               TabIndex        =   23
               Top             =   225
               Width           =   1245
               _ExtentX        =   2196
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
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
               Height          =   195
               Index           =   2
               Left            =   2145
               TabIndex        =   25
               Top             =   330
               Width           =   420
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
               Height          =   195
               Index           =   1
               Left            =   105
               TabIndex        =   24
               Top             =   330
               Width           =   465
            End
         End
         Begin VB.Frame Frame6 
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
            Height          =   585
            Left            =   4110
            TabIndex        =   16
            Top             =   870
            Width           =   4035
            Begin VB.CommandButton CmdBusMon 
               Height          =   240
               Left            =   1185
               Picture         =   "FrmCtaCte3.frx":0132
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   720
               MaxLength       =   1
               TabIndex        =   18
               Text            =   "TxtIdMon"
               Top             =   240
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
               Left            =   1425
               TabIndex        =   20
               Top             =   240
               Width           =   2490
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda"
               Height          =   195
               Index           =   4
               Left            =   90
               TabIndex        =   19
               Top             =   330
               Width           =   585
            End
         End
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   12795
         TabIndex        =   6
         Top             =   45
         Width           =   11460
         Begin VB.Frame Fra_Orden 
            Caption         =   "[Aplicar Orden]"
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
            Height          =   1035
            Left            =   150
            TabIndex        =   42
            Top             =   90
            Width           =   1725
            Begin VB.OptionButton Opt_Orden 
               Caption         =   "N� Documento"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   45
               Top             =   270
               Width           =   1395
            End
            Begin VB.OptionButton Opt_Orden 
               Caption         =   "N�. Registro"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   44
               Top             =   510
               Width           =   1395
            End
            Begin VB.OptionButton Opt_Orden 
               Caption         =   "Fecha Doc."
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   43
               Top             =   750
               Width           =   1395
            End
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Otros"
            Height          =   345
            Left            =   10140
            TabIndex        =   15
            Top             =   720
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Frame Frame8 
            Caption         =   "[ Documentos de Apertura ]"
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
            Height          =   1035
            Left            =   2010
            TabIndex        =   11
            Top             =   90
            Width           =   2670
            Begin VB.OptionButton OptAperturaSolo 
               Caption         =   "Ver solo Apertura"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   225
               TabIndex        =   14
               Top             =   750
               Width           =   2070
            End
            Begin VB.OptionButton OptAperturaSin 
               Caption         =   "No incluir Apertura"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   225
               TabIndex        =   13
               Top             =   510
               Width           =   1830
            End
            Begin VB.OptionButton OptAperturaCon 
               Caption         =   "Incluir Apertura"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   225
               TabIndex        =   12
               Top             =   270
               Value           =   -1  'True
               Width           =   1710
            End
         End
      End
   End
   Begin VB.Frame fraBarra 
      BorderStyle     =   0  'None
      Caption         =   "FrmConsultaDiario"
      Height          =   780
      Left            =   60
      TabIndex        =   1
      Top             =   7650
      Visible         =   0   'False
      Width           =   6285
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   150
         TabIndex        =   2
         Top             =   315
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   -15
         Y2              =   900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   6270
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   1
         X1              =   -75
         X2              =   6500
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   6270
         X2              =   6270
         Y1              =   -30
         Y2              =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Documentos"
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
         Left            =   195
         TabIndex        =   4
         Top             =   75
         Width           =   2130
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   4605
         TabIndex        =   3
         Top             =   75
         Width           =   1530
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3285
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":0264
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":07A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":0B3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":0C94
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":1026
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":11AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":15FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":1716
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":1C5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":219E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":22B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":23C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":281A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte3.frx":2986
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5655
      Left            =   0
      TabIndex        =   46
      Top             =   1920
      Width           =   11880
      _cx             =   20955
      _cy             =   9975
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
      FrontTabColor   =   14215660
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "      Detalle     |      Resumen     "
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
         Height          =   5235
         Left            =   45
         TabIndex        =   47
         Top             =   45
         Width           =   11790
         _cx             =   20796
         _cy             =   9234
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
         ForeColorSel    =   16777215
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   13
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCtaCte3.frx":2ECE
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
         Height          =   5235
         Left            =   12525
         TabIndex        =   48
         Top             =   45
         Width           =   11790
         _cx             =   20796
         _cy             =   9234
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCtaCte3.frx":3055
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
Attribute VB_Name = "FrmCtaCte3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'--modificado 30/12/09 por Johan Castro
'       considerar los canjes de las notas de debito en letras(idlib=37)
'--modificado 13/01/10 por Johan Castro
'       rpt Ventas:mostrar NC cuando no esten vinculado a un documento de referencia
'--modificado 19/01/10 por Johan Castro
'       considerar el filtro por numdoc,registro,fecha
'--modificado 11/02/10 por Johan Castro
'       considerar ajuste por diferencia de cambio a proveedores,honorarios,cliente
'--modificado 14/05/10 por Johan Castro
'       considerar filtro por intervalo de fechas
'       considerar filtro por documentos de apertura
'--modificado 21/05/10 por Johan Castro
'       no mostrar las nc de ventas cuando son anulados
'--Modificado: 15/06/11 Johan Castro
'--     Dar flexibilidad al formulario para definir el tama�o del mismo

Option Explicit
Dim RstCta As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE




Private Sub CmdBusCliPro_Click()
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tama�o     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    If OptCliente.Value = True Then
        xForm.Titulo = "Buscando Clientes"
        xForm.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente ORDER BY mae_cliente.nombre"
        xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    ElseIf OptProvee.Value = True Then
        xForm.Titulo = "Buscando Proveedores"
        xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov ORDER BY mae_prov.nombre"
        xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    ElseIf opt4ta.Value = True Then
        xForm.Titulo = "Buscando Prestador de Servicio"
        xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov WHERE mae_prov.tipper = 1 ORDER BY mae_prov.nombre"
        xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
        
    End If

    xCampos(1, 0) = "N� R.U.C.":     xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"

    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCliPro.Text = xRs("nombre")
        LblIdCliPro.Caption = xRs("id")
        TxtFchIni.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub



Private Sub Command1_Click()
    FrmCtaCteOtros.Show
    FrmCtaCteOtros.SetFocus
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        pConfigurarGrilla
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
        
        TabOne2.CurrTab = 0
        
        SeEjecuto = True
        
        '--colocar la fecha del primer dia del a�o de trabajo
        TxtFchIni.Valor = CDate("01/01/" & AnoTra)
        
        '--verificar si el a�o de trabajo es igual al a�o actual
        If NulosC(Year(Date)) < AnoTra Then
            TxtFchFin.Valor = CDate("31/12/" & AnoTra)
        Else
            TxtFchFin.Valor = Date
        End If
        
        '--enfocar el cursor en la fecha inicial
        TxtFchIni.SetFocus
    End If
End Sub

Sub CargarCli(IdCliPro)
    Dim rst As New ADODB.Recordset
    Dim Rstabo As New ADODB.Recordset
    Dim A, B, xFila As Integer
    Dim TotDebe, TotHaber As Double
    Dim TotGralDebe, TotGralHaber As Double
    Dim xNomPro As String
    Dim Cambio As Boolean
    Dim nSQL As String
'    On Error GoTo error
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Seleccione una Moneda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    BAND_INTERRUMPIR = False
    pConfigurarGrilla
    '--------------------------
    fraBarra.Left = 2798
    fraBarra.Top = 2925
    
    ProgressBar1.Max = 1
    ProgressBar1.Value = 0
    fraBarra.Visible = True
    fraBarra.Refresh
    
    
    Dim nSQLWhere As String
    Dim nCampoMuestra As String '--indica el campo que se mostrara esta en funcion de la moneda seleccionada
    nSQLWhere = ""
    If OptCliente.Value = True Then '--ventas
        If IdCliPro <> 0 Then nSQLWhere = " and vta_ventas.idcli = " & IdCliPro & " "
        nSQL = "SELECT vta_ventas.id,IIf([vta_ventas]![anulado]=-1,' ',[mae_cliente]![numruc]) AS numruc, IIf([vta_ventas]![anulado]=-1,'Anulado',[mae_cliente]![nombre]) AS nombre, IIf([vta_ventas].[numreg] Is Null Or [vta_ventas].[numreg]='','',Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4)) AS registro, 'Ventas' AS libro, mae_documento.codsun,mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc2, vta_ventas.fchdoc, vta_ventas.fchven, mae_moneda.simbolo, IIf(vta_ventas.tc=0,con_tc.impven,vta_ventas.tc) AS tipcam AS tipcam, " _
            + vbCr + " vta_ventas.idmon,vta_ventas.imptotdoc AS imptotal,vta_ventas.impsal, " _
            + vbCr + " IIf([vta_ventas].[imptotdoc] Is Null,0,IIf([vta_ventas].[idmon]=1,[vta_ventas].[imptotdoc],IIf([con_tc].[impven] Is Null,0,[vta_ventas].[imptotdoc]*[con_tc].[impven]))) AS imptotsol, " _
            + vbCr + " IIf([vta_ventas].[imptotdoc] Is Null,0,IIf([vta_ventas].[idmon]=2,[vta_ventas].[imptotdoc],IIf([con_tc].[impven] Is Null,0,[vta_ventas].[imptotdoc]/[con_tc].[impven]))) AS imptotdol " _
            + vbCr + " FROM ((((vta_ventas LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id " _
            + vbCr + " WHERE vta_ventas.fchdoc <= CDate('" & TxtFchIni.Valor & "') " & nSQLWhere _
            + vbCr + " ORDER BY IIf([vta_ventas]![anulado]=-1,'Anulado',[mae_cliente]![nombre]), [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc];"
    
    ElseIf OptProvee.Value = True Then '--compras
        If IdCliPro <> 0 Then nSQLWhere = " and com_compras.idpro = " & IdCliPro & " "
        
        nSQL = "SELECT  com_compras.id,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_compras.numreg Is Null Or com_compras.numreg='','',Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4)) AS registro, 'Compras' AS libro, mae_documento.codsun,mae_documento.abrev, com_compras!numser+'-'+com_compras!numdoc AS numdoc2, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, con_tc.impven AS tipcam, " _
            + vbCr + " com_compras.idmon,com_compras.imptot AS imptotal,com_compras.impsal, " _
            + vbCr + " IIf([com_compras].[imptot] Is Null,0,IIf([com_compras].[idmon]=1,[com_compras].[imptot],IIf([con_tc].[impven] Is Null,0,[com_compras].[imptot]*[con_tc].[impven]))) AS imptotsol, " _
            + vbCr + " IIf([com_compras].[imptot] Is Null,0,IIf([com_compras].[idmon]=2,[com_compras].[imptot],IIf([con_tc].[impven] Is Null,0,[com_compras].[imptot]/[con_tc].[impven]))) AS imptotdol " _
            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = com_compras.idmon) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
            + vbCr + " WHERE (com_compras.fchdoc <=CDate('" & TxtFchIni.Valor & "') )" & nSQLWhere & "AND ( com_compras.tipdoc <> 7) " _
            + vbCr + " ORDER BY mae_prov!nombre, com_compras.fchdoc;"
        
'WHERE (((com_compras.fchdoc)<=CDate('17/12/2008')) AND ((com_compras.idpro)=1176) AND ((com_compras.tipdoc)<>7))
    Else
        
        Exit Sub
    End If
    
    If NulosN(TxtIdMon.Text) = 1 Then
        nCampoMuestra = "imptotsol"
    ElseIf NulosN(TxtIdMon.Text) = 2 Then
        nCampoMuestra = "imptotdol"
    Else
        fraBarra.Visible = False
        MsgBox "Por el momento no se puede expresar en " & LblMoneda.Caption, vbInformation, xTitulo
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    '--ejecutar la conulta
    RST_Busq rst, nSQL, xCon
    '--filtrar lo que se va mostrar
    If chk_descuadrado.Value = 0 Then
        '--obs. si selecciona la opcion todos no hace el fintro
        If OptPen.Value = True Then rst.Filter = "impsal > 0" ' FILTRAMOS LOS PENDIENTE
        If OptCan.Value = True Then rst.Filter = "impsal <= 0" ' FILTRAMOS LOS CANCELADOS
    End If
    If rst.RecordCount = 0 Then
        MsgBox "No hay documentos del cliente seleccionado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fraBarra.Visible = False
        Set rst = Nothing
        Exit Sub
    End If
    ProgressBar1.Max = rst.RecordCount
    
    Dim xSaldoDoc As Double
    Dim xFilaIni&
    Dim xColor&
    
    Me.MousePointer = vbHourglass
     
    xColor = 0
    If rst.RecordCount <> 0 Then
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo Salir:

        rst.MoveFirst
        xSaldoDoc = 0
        xNomPro = NulosC(rst("nombre"))
        xFila = Fg1.FixedRows
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(xFila, 1) = "N� R.U.C. :"
        Fg1.TextMatrix(xFila, 2) = NulosC(rst("numruc"))
        Fg1.TextMatrix(xFila, 4) = NulosC(rst("nombre"))
        '*****resumen
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(rst("numruc"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(rst("nombre"))
        '******
        xFilaIni = xFila
        With Fg1
            .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H800000
            .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            .FillStyle = flexFillRepeat
            .CellFontBold = True
        End With
    
        TotDebe = 0
        TotHaber = 0
        
        Cambio = False
        
        Dim mRowIni As Integer

        For A = 1 To rst.RecordCount    '--GRUPO DE CLIENTE/PROVEEDOR
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo Salir:
            ProgressBar1.Value = A
            
            xSaldoDoc = 0
            
            If NulosC(rst("nombre")) <> xNomPro Then
                DoEvents
                Cambio = True
                xNomPro = NulosC(rst("nombre"))
                Fg1.Rows = Fg1.Rows + 1
                xFila = xFila + 1
                Fg1.TextMatrix(xFila, 4) = "TOTAL -->"
                Fg1.TextMatrix(xFila, 10) = Format(TotDebe, FORMAT_MONTO)
                Fg1.TextMatrix(xFila, 11) = Format(TotHaber, FORMAT_MONTO)
                Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
                
                '*****resumen
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotDebe, FORMAT_MONTO)
                Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotHaber, FORMAT_MONTO)
                Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
                '******

                
                With Fg1
                    .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
                    .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                    .FillStyle = flexFillRepeat
                    .CellFontBold = True
                End With
                
                '----MOSTRAR SOLO DESCUADRADOS ---------
                If chk_descuadrado.Value = 1 Then
                    If NulosN(Fg1.TextMatrix(xFila, 12)) = 0 Then
                        GRID_DELETE Fg1, Fg1.Rows - 2, Fg1.Rows - 1, e_Fila
                        Fg1.Rows = Fg1.Rows + 1
                        xFila = Fg1.Rows - 1
                    Else
                        Fg1.Rows = Fg1.Rows + 2
                        xFila = xFila + 2
                    End If
                    '---del resumen
                    If NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 5)) = 0 Then
                        GRID_DELETE Fg2, Fg2.Rows - 1, Fg2.Rows - 1, e_Fila
                    End If
                    '---------------
                Else
                    Fg1.Rows = Fg1.Rows + 2
                    xFila = xFila + 2
                End If
                '---------------------------------------------------------
                TotGralHaber = TotGralHaber + TotHaber
                TotGralDebe = TotGralDebe + TotDebe
                
                TotHaber = 0
                TotDebe = 0
                '---------------------------------------------------------
                Fg1.TextMatrix(xFila, 1) = "N� R.U.C. :"
                Fg1.TextMatrix(xFila, 2) = NulosC(rst("numruc"))
                Fg1.TextMatrix(xFila, 4) = xNomPro
                
                '*****resumen
                Fg2.Rows = Fg2.Rows + 1
                Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(rst("numruc"))
                Fg2.TextMatrix(Fg2.Rows - 1, 2) = xNomPro
                '******

                
                With Fg1
                    .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H800000
                    .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                    .FillStyle = flexFillRepeat
                    .CellFontBold = True
                End With
            Else
                Cambio = False
            End If

            Fg1.Rows = Fg1.Rows + 1
            xFila = xFila + 1
            xFilaIni = xFila
            
            Fg1.TextMatrix(xFila, 1) = NulosC(rst("registro"))
            
            Fg1.TextMatrix(xFila, 2) = NulosC(rst("libro"))
            Fg1.TextMatrix(xFila, 3) = NulosC(rst("codsun"))
            Fg1.TextMatrix(xFila, 4) = NulosC(rst("numdoc2"))
            Fg1.TextMatrix(xFila, 5) = Format(rst("fchdoc"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 6) = Format(rst("fchven"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 7) = NulosC(rst("simbolo"))
            Fg1.TextMatrix(xFila, 8) = Format(NulosN(rst("imptotal")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 9) = Format(NulosN(rst("tipcam")), "###0.##0") & ""
            
            Fg1.TextMatrix(xFila, 10) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO)
            '--saldo
            Fg1.TextMatrix(xFila, 12) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO)
            
            xSaldoDoc = NulosN(rst("impsal"))
            TotDebe = TotDebe + NulosN(rst(nCampoMuestra))
            
            
            If OptCliente.Value = True Then
                'Buscamos los abonos del cliente
                'Retenciones UNION Caja y Bancos UNION Canje de documentos UNION Canje de Letra
                nSQL = "SELECT DISTINCT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, 'Retenciones' AS libro, mae_documento.codsun,mae_documento.abrev, [con_retencion]![numser]+'-'+[con_retencion]![numdoc] AS numdoc, con_retencion.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_retenciondet.impret AS imptotal, IIf([con_retenciondet].[impret] Is Null,0,IIf([con_retencion].[idmon]=1,[con_retenciondet].[impret],IIf([con_tc].[impven] Is Null,0,[con_retenciondet].[impret]*[con_tc].[impven]))) AS imptotsol, IIf([con_retenciondet].[impret] Is Null,0,IIf([con_retencion].[idmon]=2,[con_retenciondet].[impret],IIf([con_tc].[impven] Is Null OR [con_tc].[impven]=0,0,[con_retenciondet].[impret]/[con_tc].[impven]))) AS imptotdol " _
                    + vbCr + " FROM mae_moneda RIGHT JOIN ((((con_diario RIGHT JOIN (con_retencion LEFT JOIN mae_documento ON con_retencion.iddoc = mae_documento.id) ON con_diario.idmov = con_retencion.id) LEFT JOIN con_tc ON con_retencion.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) LEFT JOIN con_retenciondet ON con_retencion.id = con_retenciondet.id) ON mae_moneda.id = con_retencion.idmon " _
                    + vbCr + " WHERE (((con_diario.idlib) = 5) And ((con_retencion.tipo) = 2) And ((con_retenciondet.iddoc) = " & rst("id") & ")); " _
                    + vbCr + " UNION " _
                    + vbCr + " SELECT Mid([numreg],1,2)+'01'+Mid([numreg],3,4) AS registro, 'Caja Bancos' AS libro, '' AS codsun, tes_documentos.abrev, " _
                    + vbCr + " IIf([tes_cajaorigendet]![numser]<>'',[tes_cajaorigendet]![numser]+'-'+[tes_cajaorigendet]![numdoc],[tes_cajaorigendet]![numdoc]) AS numdoc, " _
                    + vbCr + " tes_caja.fchope AS fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, tes_cajadestinodet.acuenta AS imptotal, " _
                    + vbCr + " IIf([tes_caja]![idmon]=1,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]*[con_tc]![impven]) AS imptotsol, " _
                    + vbCr + " IIf([tes_caja]![idmon]=2,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]/[con_tc]![impven]) AS imptotdol " _
                    + vbCr + " FROM (((((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) " _
                    + vbCr + " INNER JOIN (tes_cajadestino INNER JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND " _
                    + vbCr + " (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes) INNER JOIN tes_cajaori " _
                    + vbCr + " ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND " _
                    + vbCr + " (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id " _
                    + vbCr + " WHERE (((tes_cajadestinodet.idmod)=2) AND ((tes_cajadestinodet.iddoc)=" & rst("id") & ") AND ((tes_caja.tipmov)=1))" _
                    + vbCr + " UNION" _
                    + vbCr + " SELECT DISTINCT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, 'Canje de documentos' AS libro, '99' AS codsun, 'CAN' AS abrev, [con_canjes].[numser] & '-' & [con_canjes].[numdoc] AS numdoc, con_canjes.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_canjesdet.impcan AS imptotal, IIf([con_canjesdet].[impcan] Is Null,0,IIf([con_canjes].[idmon]=1,[con_canjesdet].[impcan],IIf([con_tc].[impven] Is Null,0,[con_canjesdet].[impcan]*[con_tc].[impven]))) AS imptotsol, IIf([con_canjesdet].[impcan] Is Null,0,IIf([con_canjes].[idmon]=2,[con_canjesdet].[impcan],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0,0,[con_canjesdet].[impcan]/[con_tc].[impven]))) AS imptotdol " _
                    + vbCr + " FROM ((((con_canjes LEFT JOIN con_diario ON con_canjes.id = con_diario.idmov) LEFT JOIN con_tc ON con_canjes.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_canjes.idlib = mae_libros.id) LEFT JOIN con_canjesdet ON con_canjes.id = con_canjesdet.idcan) LEFT JOIN mae_moneda ON con_canjes.idmon = mae_moneda.id " _
                    + vbCr + " WHERE (((con_diario.idlib)=8) AND ((con_canjesdet.iddoc)=" & rst("id") & " and con_canjesdet.tipo=1)); " _
                    + vbCr + " UNION " _
                    + vbCr + " SELECT DISTINCT Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null Or mae_libros.codsun='','FF',mae_libros.codsun) & con_diario.numasi AS registro, 'Canje de Letra' AS libro, '100' AS codsun, 'LE' AS abrev, con_letradet.numlet AS numdoc, con_letradet.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_diario.imphabsol AS imptotal, IIf([con_letra].[idmon]=1,[con_diario].[imphabsol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[imphabdol] Is Null Or [con_diario].[imphabdol]=0,0,[con_diario].[imphabdol]*[con_tc].[impven])) AS imptotsol, IIf([con_letra].[idmon]=2,[con_diario].[imphabdol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[imphabsol] Is Null Or [con_diario].[imphabsol]=0,0,[con_diario].[imphabsol]/[con_tc].[impven])) AS imptotdol " _
                    + vbCr + " FROM (((con_letra LEFT JOIN con_tc ON con_letra.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_letra.idlib = mae_libros.id) LEFT JOIN mae_moneda ON con_letra.idmon = mae_moneda.id) RIGHT JOIN ((con_letradet LEFT JOIN con_diario ON (con_letradet.corr = con_diario.correlativo) AND (con_letradet.idlet = con_diario.idmov)) LEFT JOIN con_letradoc ON con_diario.iddocpro = con_letradoc.iddoc) ON con_letra.id = con_letradet.idlet " _
                    + vbCr + " WHERE con_letra.tiplet=2 AND (((con_letradoc.iddoc)=" & rst("id") & " ) AND ((con_diario.idlib)=37));"
                    
            Else
            
                'Buscamos los abonos al proveedor
                'Caja y Bancos UNION Canje de documentos UNION Canje de Letra UNION Rendici�n de Cuenta
                'nSQL = "SELECT Mid([numreg],1,2)+'01'+Mid([numreg],3,4) AS registro, 'Caja y Bancos' AS libro, '' AS codsun, tes_documentos.abrev, " _
                    & " IIf([tes_cajaorigendet]![numser]<>'',[tes_cajaorigendet]![numser]+'-'+[tes_cajaorigendet]![numdoc],[tes_cajaorigendet]![numdoc]) AS numdoc, " _
                    & " tes_caja.fchope AS fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, tes_cajadestinodet.acuenta AS imptotal, " _
                    & " IIf([tes_caja]![idmon]=1,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]*[con_tc]![impven]) AS imptotsol, " _
                    & " IIf([tes_caja]![idmon]=2,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]/[con_tc]![impven]) AS imptotdol " _
                    & " FROM (((((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) " _
                    & " INNER JOIN (tes_cajadestino INNER JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) " _
                    & " AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN tes_cajaori " _
                    & " ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND " _
                    & " (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id  " _
                    & " WHERE (((tes_cajadestinodet.idmod)=1) AND ((tes_caja.tipmov)=2) AND ((tes_cajadestinodet.iddoc)=" & Rst("id") & ")) " _
                    + vbCr + " UNION " _
                    + vbCr + " SELECT Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null Or mae_libros.codsun='','FF',mae_libros.codsun) & con_diario.numasi AS registro, 'Canje de documentos' AS libro, '99' AS codsun, 'CAN' AS abrev,con_canjes.numser & '-' & con_canjes.numdoc AS numdoc, con_canjes.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_diario.imphabsol AS imptotal, IIf(con_canjes.idmon=1,con_diario.imphabsol,IIf(con_tc.impven Is Null Or con_tc.impven=0 Or con_diario.imphabdol Is Null Or con_diario.imphabdol=0,0,con_diario.imphabdol*con_tc.impven)) AS imptotsol, IIf(con_canjes.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_tc.impven=0 Or con_diario.imphabsol Is Null Or con_diario.imphabsol=0,0,con_diario.imphabsol/con_tc.impven)) AS imptotdol " _
                    + vbCr + " FROM ((((con_canjes LEFT JOIN con_diario ON con_canjes.id = con_diario.idmov) LEFT JOIN con_tc ON con_canjes.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_canjes.idlib = mae_libros.id) LEFT JOIN mae_moneda ON con_canjes.idmon = mae_moneda.id) LEFT JOIN con_canjesdet ON (con_diario.iddocpro = con_canjesdet.iddoc) AND (con_diario.idmov = con_canjesdet.idcan) " _
                    + vbCr + " WHERE (((con_diario.idlib) =8) And ((con_canjesdet.iddoc) = " & Rst("id") & ") And con_canjesdet.tipo = 2); " _
                    + vbCr + " UNION " _
                    + vbCr + " SELECT DISTINCT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, 'Canje de Letra' AS libro, '100' AS codsun,'LE' AS abrev, con_letradet.numlet AS numdoc, con_letradet.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_diario.impdebsol AS imptotal, IIf([con_letra].[idmon]=1,[con_diario].[impdebsol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebdol] Is Null Or [con_diario].[impdebdol]=0,0,[con_diario].[impdebdol]*[con_tc].[impven])) AS imptotsol, IIf([con_letra].[idmon]=2,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebsol] Is Null Or [con_diario].[impdebsol]=0,0,[con_diario].[impdebsol]/[con_tc].[impven])) AS imptotdol " _
                    + vbCr + " FROM (((con_letra LEFT JOIN con_tc ON con_letra.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_letra.idlib = mae_libros.id) LEFT JOIN mae_moneda ON con_letra.idmon = mae_moneda.id) RIGHT JOIN ((con_letradet LEFT JOIN con_diario ON (con_letradet.corr = con_diario.correlativo) AND (con_letradet.idlet = con_diario.idmov)) LEFT JOIN con_letradoc ON con_diario.iddocpro = con_letradoc.iddoc) ON con_letra.id = con_letradet.idlet " _
                    + vbCr + " WHERE con_letra.tiplet=1 AND (((con_letradoc.iddoc)=" & Rst("id") & " ) AND ((con_diario.idlib)=37));" _
                    + vbCr + " UNION " _
                    + vbCr + " SELECT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, 'Rendici�n de Cuenta' AS libro, '101' AS codsun,'REN' AS abrev, con_devoluciones.numdoc, con_devoluciones.fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, con_diario.impdebsol AS imptotal, IIf([con_devoluciones].[idmon]=1,[con_diario].[impdebsol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebdol] Is Null Or [con_diario].[impdebdol]=0,0,[con_diario].[impdebdol]*[con_tc].[impven])) AS imptotsol, IIf([con_devoluciones].[idmon]=2,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebsol] Is Null Or [con_diario].[impdebsol]=0,0,[con_diario].[impdebsol]/[con_tc].[impven])) AS imptotdol " _
                    + vbCr + " FROM ((con_devoluciones LEFT JOIN con_tc ON con_devoluciones.fchemi = con_tc.fecha) LEFT JOIN mae_moneda ON con_devoluciones.idmon = mae_moneda.id) INNER JOIN (mae_libros INNER JOIN (con_devolucionesdet INNER JOIN con_diario ON (con_devolucionesdet.idcom = con_diario.iddocpro) AND (con_devolucionesdet.id = con_diario.idmov)) ON mae_libros.id = con_diario.idlib) ON con_devoluciones.id = con_devolucionesdet.id " _
                    + vbCr + " WHERE (((con_devolucionesdet.idcom)=" & Rst("id") & " ) AND ((con_diario.idlib)=38));"
                    
                    
                nSQL = "SELECT Mid([numreg],1,2)+'01'+Mid([numreg],3,4) AS registro, 'Caja y Bancos' AS libro, '' AS codsun, tes_documentos.abrev, " _
                    & " IIf([tes_cajaorigendet]![numser]<>'',[tes_cajaorigendet]![numser]+'-'+[tes_cajaorigendet]![numdoc],[tes_cajaorigendet]![numdoc]) AS numdoc,  " _
                    & " tes_caja.fchope AS fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, tes_cajadestinodet.acuenta AS imptotal,  " _
                    & " IIf([tes_caja]![idmon]=1,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]*[con_tc]![impven]) AS imptotsol,  " _
                    & " IIf([tes_caja]![idmon]=2,[tes_cajadestinodet]![acuenta],[tes_cajadestinodet]![acuenta]/[con_tc]![impven]) AS imptotdol  " _
                    & " FROM (((((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha)  " _
                    & " INNER JOIN (tes_cajadestino INNER JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) " _
                    & " AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN tes_cajaori " _
                    & " ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND " _
                    & " (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id  " _
                    & " WHERE (((tes_cajadestinodet.idmod)=1) AND ((tes_caja.tipmov)=2) AND ((tes_cajadestinodet.iddoc)=" & rst("id") & ")) " _
                    & " Union " _
                    & " SELECT Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null Or mae_libros.codsun='','FF',mae_libros.codsun) & con_diario.numasi AS registro, " _
                    & " 'Canje de documentos' AS libro, '99' AS codsun, 'CAN' AS abrev,con_canjes.numser & '-' & con_canjes.numdoc AS numdoc, con_canjes.fchemi, " _
                    & " mae_moneda.simbolo, con_tc.impven AS tipcam, con_diario.imphabsol AS imptotal, IIf(con_canjes.idmon=1,con_diario.imphabsol, " _
                    & " IIf(con_tc.impven Is Null Or con_tc.impven=0 Or con_diario.imphabdol Is Null Or con_diario.imphabdol=0,0,con_diario.imphabdol*con_tc.impven)) AS imptotsol, " _
                    & " IIf(con_canjes.idmon=2,con_diario.imphabdol,IIf(con_tc.impven Is Null Or con_tc.impven=0 Or con_diario.imphabsol Is Null Or " _
                    & " con_diario.imphabsol=0,0,con_diario.imphabsol/con_tc.impven)) AS imptotdol FROM ((((con_canjes LEFT JOIN con_diario ON con_canjes.id = con_diario.idmov) " _
                    & " LEFT JOIN con_tc ON con_canjes.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_canjes.idlib = mae_libros.id) LEFT JOIN mae_moneda " _
                    & " ON con_canjes.idmon = mae_moneda.id) LEFT JOIN con_canjesdet ON (con_diario.iddocpro = con_canjesdet.iddoc) AND (con_diario.idmov = con_canjesdet.idcan) " _
                    & " WHERE (((con_diario.idlib) =8) And ((con_canjesdet.iddoc) = " & rst("id") & ") And con_canjesdet.tipo = 2) "
                nSQL = nSQL & " Union " _
                    & " SELECT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, " _
                    & " 'Rendici�n de Cuenta' AS libro, '101' AS codsun,'REN' AS abrev, con_devoluciones.numdoc, con_devoluciones.fchemi, mae_moneda.simbolo, " _
                    & " con_tc.impven AS tipcam, con_diario.impdebsol AS imptotal, IIf([con_devoluciones].[idmon]=1,[con_diario].[impdebsol],IIf([con_tc].[impven] Is Null " _
                    & " Or [con_tc].[impven]=0 Or [con_diario].[impdebdol] Is Null Or [con_diario].[impdebdol]=0,0,[con_diario].[impdebdol]*[con_tc].[impven])) AS imptotsol, " _
                    & " IIf([con_devoluciones].[idmon]=2,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebsol] Is Null " _
                    & " Or [con_diario].[impdebsol]=0,0,[con_diario].[impdebsol]/[con_tc].[impven])) AS imptotdol FROM ((con_devoluciones LEFT JOIN con_tc " _
                    & " ON con_devoluciones.fchemi = con_tc.fecha) LEFT JOIN mae_moneda ON con_devoluciones.idmon = mae_moneda.id) INNER JOIN (mae_libros " _
                    & " INNER JOIN (con_devolucionesdet INNER JOIN con_diario ON (con_devolucionesdet.idcom = con_diario.iddocpro) AND (con_devolucionesdet.id = con_diario.idmov)) " _
                    & " ON mae_libros.id = con_diario.idlib) ON con_devoluciones.id = con_devolucionesdet.id  WHERE (((con_devolucionesdet.idcom)=" & rst("id") & " ) " _
                    & " AND ((con_diario.idlib)=38))"
                nSQL = nSQL & " Union " _
                    & " SELECT Mid([com_compras]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_compras]![numreg],3,4) AS registro, mae_libros.descripcion AS libro, " _
                    & " mae_libros.codsun, '' AS abrev, [com_compras]![numser] & '-' & [com_compras]![numdoc] AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, " _
                    & " con_tc.impven AS timpcam, com_compras.imptot AS imptotal, IIf([com_compras]![idmon]=1,[com_compras]![imptot],[com_compras]![imptot]*[con_tc]![impven]) AS imptotsol, " _
                    & " IIf([com_compras]![idmon]=2,[com_compras]![imptot],[com_compras]![imptot]/[con_tc]![impven]) AS imptotdol FROM mae_moneda RIGHT JOIN " _
                    & " ((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) " _
                    & " ON mae_moneda.id = com_compras.idmon WHERE (((com_compras.iddocref)=" & rst("id") & "))"

                '    & " SELECT Mid([com_compras]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_compras]![numreg],3,4) AS registro, " _
                    & " mae_libros.descripcion, mae_libros.codsun, [com_compras]![numser] & '-' & [com_compras]![numdoc] AS numdoc, com_compras.fchdoc AS fchemi, " _
                    & " mae_moneda.simbolo, con_tc.impven AS timpcam, com_compras.imptot AS imptotal, IIf([com_compras]![idmon]=1,[com_compras]![imptot], " _
                    & " [com_compras]![imptot]*[con_tc]![impven]) AS imptotsol, IIf([com_compras]![idmon]=2,[com_compras]![imptot],[com_compras]![imptot]/[con_tc]![impven]) AS imptotdol " _
                    & " FROM (mae_moneda RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_moneda.id = com_compras.idmon) " _
                    & " LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha WHERE (((com_compras.iddocref)=" & Rst("id") & "))"

'                    & " SELECT DISTINCT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, " _
'                    & " 'Canje de Letra' AS libro, '100' AS codsun,'LE' AS abrev, con_letradet.numlet AS numdoc, con_letradet.fchemi, mae_moneda.simbolo, " _
'                    & " con_tc.impven AS tipcam, con_diario.impdebsol AS imptotal, IIf([con_letra].[idmon]=1,[con_diario].[impdebsol],IIf([con_tc].[impven] Is Null " _
'                    & " Or [con_tc].[impven]=0 Or [con_diario].[impdebdol] Is Null Or [con_diario].[impdebdol]=0,0,[con_diario].[impdebdol]*[con_tc].[impven])) AS imptotsol, " _
'                    & " IIf([con_letra].[idmon]=2,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_tc].[impven]=0 Or [con_diario].[impdebsol] Is Null " _
'                    & " Or [con_diario].[impdebsol]=0,0,[con_diario].[impdebsol]/[con_tc].[impven])) AS imptotdol  FROM (((con_letra LEFT JOIN con_tc ON con_letra.fchemi = con_tc.fecha) " _
'                    & " LEFT JOIN mae_libros ON con_letra.idlib = mae_libros.id) LEFT JOIN mae_moneda ON con_letra.idmon = mae_moneda.id) RIGHT JOIN ((con_letradet " _
'                    & " LEFT JOIN con_diario ON (con_letradet.corr = con_diario.correlativo) AND (con_letradet.idlet = con_diario.idmov)) LEFT JOIN con_letradoc " _
'                    & " ON con_diario.iddocpro = con_letradoc.iddoc) ON con_letra.id = con_letradet.idlet  WHERE con_letra.tiplet=1 AND (((con_letradoc.iddoc)=" & Rst("id") & " ) " _
'                    & " AND ((con_diario.idlib)=37)) " _
'                    & " Union " _

            End If
            
            Set Rstabo = Nothing
            RST_Busq Rstabo, nSQL, xCon
            If Rstabo.RecordCount <> 0 Then
                Rstabo.MoveFirst
                Rstabo.Sort = "fchemi ASC"
                Do While Not Rstabo.EOF
                    '--SI SE NTERRUMPE EL PROCESO => SALIR
'                    DoEvents
                    If BAND_INTERRUMPIR = True Then GoTo Salir:
                    Fg1.Rows = Fg1.Rows + 1
                    xFila = xFila + 1
                    
                    Fg1.TextMatrix(xFila, 1) = NulosC(Rstabo("registro"))
                    
                    Fg1.TextMatrix(xFila, 2) = NulosC(Rstabo("libro"))
                    Fg1.TextMatrix(xFila, 3) = NulosC(Rstabo("codsun"))
                    Fg1.TextMatrix(xFila, 4) = NulosC(Rstabo("numdoc"))
                    Fg1.TextMatrix(xFila, 5) = Format(Rstabo("fchemi"), FORMAT_DATE)
                    Fg1.TextMatrix(xFila, 7) = NulosC(Rstabo("simbolo"))
                    Fg1.TextMatrix(xFila, 8) = Format(NulosN(Rstabo("imptotal")), FORMAT_MONTO)
                    Fg1.TextMatrix(xFila, 9) = Format(NulosN(Rstabo("tipcam")), "####.###")
                    
                    Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                    Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                    TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
                    Rstabo.MoveNext
                Loop
            End If
            
            '---ACTUALIZANDO EL SALDO AL DOCUMENTO
            If xSaldoDoc <> NulosN(Fg1.TextMatrix(xFila, 12)) And NulosN(rst("idmon")) = NulosN(TxtIdMon.Text) Then
                If OptCliente.Value = True Then     '--VENTAS
                    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & NulosN(Fg1.TextMatrix(xFila, 12)) & " WHERE (((vta_ventas.id)=" & rst("id") & "))"
                Else                                '--COMPRAS
                    xCon.Execute "UPDATE com_compras SET com_compras.impsal = " & NulosN(Fg1.TextMatrix(xFila, 12)) & " WHERE (((com_compras.id)=" & rst("id") & "))"
                End If
            End If
            '----MOSTRAR SOLO DESCUADRADOS ---------
            If chk_descuadrado.Value = 1 Then
                If NulosN(Fg1.TextMatrix(xFila, 12)) >= 0 Then
                    GRID_DELETE Fg1, Fg1.Rows - 1 - Rstabo.RecordCount, Fg1.Rows - 1, e_Fila
                    '*********************************************
                    If Rstabo.RecordCount <> 0 Then
                        Rstabo.MoveFirst
                        Do While Not Rstabo.EOF
                            TotHaber = TotHaber - NulosN(Rstabo(nCampoMuestra))
                            Rstabo.MoveNext
                        Loop
                    End If
                    TotDebe = TotDebe - NulosN(rst(nCampoMuestra))
                    '*********************************************
                    xFila = Fg1.Rows - 1
                    mRowIni = -1
                Else
                    mRowIni = 0
                End If
            End If
            '---------------------------------------------------------
            
            rst.MoveNext
            If rst.EOF = True Then
                Exit For
            End If
            If mRowIni = 0 Then
                If xColor = 0 Then
                    GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 1, Fg1.Cols - 1, &H80000005
                    xColor = 1
                Else
                    GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HE0DCDA
                    xColor = 0
                End If
            End If
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        xFila = xFila + 1
        Fg1.TextMatrix(xFila, 4) = "TOTAL -->"

        Fg1.TextMatrix(xFila, 10) = Format(TotDebe, FORMAT_MONTO)
        Fg1.TextMatrix(xFila, 11) = Format(TotHaber, FORMAT_MONTO)
        Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
        '*****resumen
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotDebe, FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotHaber, FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
        '******

        With Fg1
            .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
            .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            .FillStyle = flexFillRepeat
            .CellFontBold = True
        End With
        '----MOSTRAR SOLO DESCUADRADOS ---------
        If chk_descuadrado.Value = 1 Then
            If NulosN(Fg1.TextMatrix(xFila, 12)) = 0 Then
                GRID_DELETE Fg1, Fg1.Rows - 2, Fg1.Rows - 1, e_Fila
                xFila = Fg1.Rows - 1
            End If
            '--del resumen
            If NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 5)) = 0 Then
                GRID_DELETE Fg2, Fg2.Rows - 1, Fg2.Rows - 1, e_Fila
            End If
            '---------------
        End If
        
        '---------------------------------------------------------

        If TotGralDebe <> 0 Or TotGralHaber <> 0 Then
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = "TOTAL GRAL -->"
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(TotGralDebe, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(TotGralHaber, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
            With Fg1
                .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
                .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                .FillStyle = flexFillRepeat
                .CellFontBold = True
            End With
            '*****resumen
            Fg2.Rows = Fg2.Rows + 2
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = "TOTAL GRAL -->"
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotGralDebe, FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotGralHaber, FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
            With Fg2
                .Cell(flexcpForeColor, Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1) = &H80000008
                .Select Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1
                .FillStyle = flexFillRepeat
                .CellFontBold = True
            End With
            '******
    End If
    
    End If
    If mRowIni = 0 Then
        If xColor = 0 Then
            GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 2, Fg1.Cols - 1, &H80000005
            xColor = 1
        Else
            GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 2, Fg1.Cols - 1, &HE0DCDA
            xColor = 0
        End If
    End If
    Set rst = Nothing
    Set Rstabo = Nothing
    fraBarra.Visible = False
    Me.MousePointer = vbDefault
    MsgBox "La Consulta fue se realiz� Correctamente", vbInformation, xTitulo
    Exit Sub
Salir:
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set rst = Nothing
    Set Rstabo = Nothing
    MsgBox "La Consulta fue Interrumpida", vbInformation, xTitulo
error:
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set rst = Nothing
    Set Rstabo = Nothing
    SHOW_ERROR Me.Name, "CargarCli"
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
    TxtCliPro.Text = ""
    TxtFchIni.Valor = ""
    TxtFchIni.Valor = Date
    LblMoneda.Caption = ""
    TxtIdMon.Text = ""
    SeEjecuto = False
    Fg1.AutoSearch = flexSearchFromTop
    Fg2.AutoSearch = flexSearchFromTop
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
  
    If Me.Height > 3000 Then
        TabOne1.Top = 1920
        TabOne1.Width = Me.Width - 150
        TabOne1.Height = Me.Height - 2320
    End If
End Sub

Private Sub opt4ta_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    Label1(0).Caption = "Prestador de Servicio"

End Sub

Private Sub OptCan_Click()
'    chk_descuadrado.Enabled = True
End Sub

Private Sub OptCliente_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    Label1(0).Caption = "Cliente"
End Sub


Private Sub OptPen_Click()
    chk_descuadrado.Value = 0
''    chk_descuadrado.Enabled = False
End Sub

Private Sub OptProvee_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    Label1(0).Caption = "Proveedor"
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

Private Sub OptTodos_Click()
'    chk_descuadrado.Enabled = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then pExportar
    If Button.Index = 4 Then pImprimir
    If Button.Index = 6 Then
        Set RstCta = Nothing
        Unload Me
    End If
End Sub

Sub pExportar()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    Dim nPeriodo As String
    Dim nTitulo1 As String
    
    nPeriodo = "Al  " + CStr(TxtFchIni.Valor)

    nTitulo1 = "(Expresado en " & LblMoneda.Caption & ")"
    
    If TabOne1.CurrTab = 0 Then '--detalle
        '''oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Cuenta Corriente - " + IIf(OptCliente.Value = True, "Cliente", "Proveedor"), nPeriodo, nTitulo1, "Cuenta Corriente An�lisis"
        
        GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "Cuenta Corriente - " + IIf(OptCliente.Value = True, "Cliente", "Proveedor"), "Expresado en " & LblMoneda.Caption, "Cuenta Corriente An�lisis"

    Else
        oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg2, "Resumen de Cuenta Corriente - " + IIf(OptCliente.Value = True, "Cliente", "Proveedor"), nPeriodo, nTitulo1, "Cuenta Corriente An�lisis"
    End If
    
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportar"
End Sub

Private Sub TxtCliPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCliPro_Click
    End If
End Sub

'***********************************************************************************************
'------------CAMBIOS AL 020108

Private Sub CmdBusMon_Click()
    
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tama�o     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":   xCampos(1, 1) = "id":              xCampos(1, 2) = "500":      xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_moneda ORDER BY descripcion ;", xCampos(), "Buscando Moneda", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    TxtIdMon.Text = xRs("id") & ""
    LblMoneda.Caption = xRs("descripcion") & ""
    
Salir:
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

Private Sub BuscarVSFlexGrid()
    On Error GoTo error
    
    Dim oExport As New SGI2_funciones.formularios
    Dim xCampos(4, 3) As String
    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
    xCampos(0, 0) = "N�.Registro":      xCampos(0, 1) = "1":    xCampos(0, 2) = "C":    xCampos(0, 3) = "-1"
    xCampos(1, 0) = "Origen":           xCampos(1, 1) = "2":    xCampos(1, 2) = "C":    xCampos(1, 3) = "0"
    xCampos(2, 0) = "N� Documento":     xCampos(2, 1) = "4":    xCampos(2, 2) = "C":    xCampos(2, 3) = "0"
    xCampos(3, 0) = Label1(0):          xCampos(3, 1) = "4":    xCampos(3, 2) = "C":    xCampos(3, 3) = "0"
    xCampos(4, 0) = "Fch.Emi.":         xCampos(4, 1) = "5":    xCampos(4, 2) = "F":    xCampos(4, 3) = "0"
    
    oExport.VSFlexGrid_Buscar Me.hWnd, Fg1, xCampos(), Fg1.Row
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "BuscarVSFlexGrid"
End Sub

Private Sub pConfigurarGrilla()
    Dim A As Integer
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    With Fg1
        '-----
        .Rows = 2
        .Cols = 17
        .FixedRows = 2
        .FrozenCols = 0
        .RowHeight(0) = 250
        .ColWidth(0) = 200
        UNIR_CELDAS Fg1, 0, 1, 0, 9, "DATOS DEL DOCUMENTO", flexAlignCenterCenter
        FORMATO_CELDA Fg1, 0, 1, vbBlack, True, &HD8E9EC
        If Trim(LblMoneda.Caption) = "" Then
            UNIR_CELDAS Fg1, 0, 10, 0, 12, "IMPORTES", flexAlignCenterCenter
        Else
            UNIR_CELDAS Fg1, 0, 10, 0, 12, "IMPORTES EN " & UCase(LblMoneda.Caption), flexAlignCenterCenter
        End If
        FORMATO_CELDA Fg1, 0, 10, vbBlack, True, &HD8E9EC
        
        UNIR_CELDAS Fg1, 0, 13, 0, 14, "REFERENCIA", flexAlignCenterCenter
        FORMATO_CELDA Fg1, 0, 13, vbBlack, True, &HD8E9EC
        
        .ColWidth(1) = 350
'
        .TextMatrix(1, 1) = "N� Registro":  .ColWidth(1) = 900:   .ColAlignment(1) = flexAlignLeftCenter:     .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 2) = "Origen":       .ColWidth(2) = 1200:   .ColAlignment(2) = flexAlignLeftCenter:     .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "T.D.":         .ColWidth(3) = 450:    .ColAlignment(3) = flexAlignLeftCenter:     .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 4) = "N�.Documento": .ColWidth(4) = 1600:   .ColAlignment(4) = flexAlignLeftCenter:     .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 5) = "Fch.Emi.":     .ColWidth(5) = 800:    .ColAlignment(5) = flexAlignCenterBottom:   .Row = 1: .Col = 5: .CellAlignment = flexAlignCenterBottom
        .TextMatrix(1, 6) = "Fch.Ven.":     .ColWidth(6) = 800:    .ColAlignment(6) = flexAlignCenterBottom:   .Row = 1: .Col = 6: .CellAlignment = flexAlignCenterBottom
        .TextMatrix(1, 7) = "M":            .ColWidth(7) = 450:    .ColAlignment(7) = flexAlignLeftCenter:    .Row = 1: .Col = 7: .CellAlignment = flexAlignCenterBottom
        
        .TextMatrix(1, 8) = "Imp":          .ColWidth(8) = 900:    .ColAlignment(8) = flexAlignRightCenter:   .Row = 1: .Col = 8: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 9) = "T.C.":         .ColWidth(9) = 500:    .ColAlignment(9) = flexAlignRightCenter:    .Row = 1: .Col = 9: .CellAlignment = flexAlignRightCenter
        '----------------------
        .TextMatrix(1, 10) = "Debe":       .ColWidth(10) = 1150:  .ColAlignment(10) = flexAlignRightCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 11) = "Haber":       .ColWidth(11) = 1150:  .ColAlignment(11) = flexAlignRightCenter:   .Row = 1: .Col = 11: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 12) = "Saldo":       .ColWidth(12) = 1150:  .ColAlignment(12) = flexAlignRightCenter:   .Row = 1: .Col = 12: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 13) = "N�.Documento":       .ColWidth(13) = 1400:  .ColAlignment(13) = flexAlignLeftCenter:   .Row = 1: .Col = 13: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 14) = "Glosa":       .ColWidth(14) = 3500:  .ColAlignment(14) = flexAlignLeftCenter:   .Row = 1: .Col = 14: .CellAlignment = flexAlignLeftCenter
        
        
        .TextMatrix(1, 15) = "N�. Cuenta":       .ColWidth(15) = 0:  .ColAlignment(15) = flexAlignLeftCenter:   .Row = 1: .Col = 15: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 16) = "Descripci�n":       .ColWidth(16) = 0:  .ColAlignment(16) = flexAlignLeftCenter:   .Row = 1: .Col = 16: .CellAlignment = flexAlignLeftCenter
        
        
        .SelectionMode = flexSelectionByRow
    End With
    
    With Fg2
        '-----
        .Rows = 1
        .Cols = 6
        .FixedRows = 1
        .FrozenCols = 0
        .RowHeight(0) = 250
        .ColWidth(0) = 200:
        .TextMatrix(0, 1) = "R.U.C.":   .ColWidth(1) = 1200:  .ColAlignment(1) = flexAlignCenterCenter: .Row = 0: .Col = 1: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 2) = "Nombres":  .ColWidth(2) = 5500:  .ColAlignment(2) = flexAlignLeftCenter:   .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        '----------------------
        .TextMatrix(0, 3) = "Debe":    .ColWidth(3) = 1300:  .ColAlignment(3) = flexAlignRightCenter:  .Row = 0: .Col = 3: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 4) = "Haber":    .ColWidth(4) = 1300:  .ColAlignment(4) = flexAlignRightCenter:  .Row = 0: .Col = 4: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(0, 5) = "Saldo":    .ColWidth(5) = 1300:  .ColAlignment(5) = flexAlignRightCenter:  .Row = 0: .Col = 5: .CellAlignment = flexAlignRightCenter
        For A = 1 To .Cols - 1
            FORMATO_CELDA Fg2, 0, A, vbBlack, True, &HD8E9EC
        Next
        .SelectionMode = flexSelectionByRow
    End With
    TabOne1.CurrTab = 0
    DoEvents
End Sub


Private Sub pImprimir()

    On Error GoTo error
    
        Dim oPrint As New SGI2_funciones.formularios
        Dim nPeriodo As String
        Dim nTitulo As String
        Dim nTitulo1 As String
        Dim nTipo As String
        nPeriodo = "Al  " + CStr(TxtFchIni.Valor)
        nTitulo1 = "(Expresado en " & LblMoneda.Caption & ")"
        If OptCliente.Value = True Then
            nTipo = "Cliente"
        ElseIf OptProvee.Value = True Then
            nTipo = "Proveedor"
        Else
            nTipo = "Prestador de Servicio"
        End If
    Me.MousePointer = vbHourglass
    
    If TabOne1.CurrTab = 0 Then
        nTitulo = "Detalle de Cuenta Corriente - " + nTipo
        oPrint.Imprimir_x_VSFlexGrid Fg1, nTitulo, nTitulo1, nPeriodo, True, True
    Else
        nTitulo = "Resumen de Cuenta Corriente - " + nTipo
        oPrint.Imprimir_x_VSFlexGrid Fg2, nTitulo, nTitulo1, nPeriodo, True, True
    End If
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

Sub CargarCli2(IdCliPro)
    '===================================================================================================
    'creado: xx/xx/xx Enrique Pollongo
    'Prop�sito: Generar la consulta a nivel de detalle
    '
    'Entradas:  IdCliPro=Codigo del proveedor, cliente, prestador de servicio OPCIONAL=0
    '
    'Resultados: Consulta segun parametros indicados
    '
    'Modificado: 25/10/10 Johan Castro
    '           Modificar presentaci�n de Formulario.
    '           Mostrar NC que no tienen referencia, figuren como pendiente.
    '           Unir Ruc y Razon Social en una columna combinada, Util para la impresi�n mendiante formato
    '           09/03/11 Johan Castro
    '           Cuando el tipo de documento sea NC y no tenga documento al que haga referencia debe mostrarse
    '           en debe para compras y en haber para ventas
    '           29/04/11 Johan Castro
    '           Modificar las consultasde ventas,compras,honorarios;campo imptotal en apertura tomara valor de
    '           campo imptotori. Mostrar solo registros cuyo campo imptotal sea diferente a cero
    '           17/11/11 Johan Castro
    '           Modificar consulta de ventas mostrar nc que figuren en tesoreria-ingreso y tesoreria-egresos [(cobranzas-destino) o (canjes-origen)]
    '           Modificar consulta de compras mostrar nc que figuren en tesoreria-egresos y tesoreria-ingresos[(cobranzas-destino) o (canjes-origen)]
    '           Modificar consulta de diario agregar campos(idlib, tipmov,tipo, rtipdoc) para ventas y compras
    '           Modificar consulta de diario con nc para no mostar nc que esten en tesoreria(ingresos, egresos) para ventas y compras
    '           Mostrar los descuadrados solo cuando opttodos este seleccionado
    '===================================================================================================

    
    Dim rst As New ADODB.Recordset
    Dim Rstabo As New ADODB.Recordset
    Dim A, B, xFila As Long
    Dim TotDebe, TotHaber As Double
    Dim TotGralDebe, TotGralHaber As Double
    Dim xNomPro As String '--Razon social del proveedor, cliente, prestador de servicio
    Dim Cambio As Boolean
    Dim nSQL As String
    Dim sSaldoFinal As Double '--indica el saldo final por cada documento
        
    'On Error GoTo error
    
    '--posicionar en vista de inicio
    TabOne2.CurrTab = 0
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Seleccione una Moneda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    BAND_INTERRUMPIR = False
    pConfigurarGrilla
    '--------------------------
    fraBarra.Left = 2798
    fraBarra.Top = 2925
    
    TabOne1.CurrTab = 1
    
    ProgressBar1.Max = 1
    ProgressBar1.Value = 0
    fraBarra.Visible = True
    fraBarra.Refresh
    DoEvents
    
    
    Dim nSQLWhere As String '--almacenara la condicion de la consulta
    Dim nCampoMuestra As String '--indica el campo que se mostrara esta en funcion de la moneda seleccionada
    Dim nSQLAjuste  As String
    Dim nSQLApertura As String '--filtro para documentos de apertura
    Dim nSQLFecha As String '--filtro por intervalo de fechas
    
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " and (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    '------
    nSQLWhere = ""
        
    '--aplicar filtro por fecha
    If OptFch(0).Value = True Then '--x fecha de documento
        nSQLFecha = " and ( vta_ventas.fchdoc between CDate('" & TxtFchIni.Valor & "') and CDate('" & TxtFchFin.Valor & "') )"
    ElseIf OptFch(1).Value = True Then '--x fecha de registro
        nSQLFecha = " and ( vta_ventas.fchreg between CDate('" & TxtFchIni.Valor & "') and CDate('" & TxtFchFin.Valor & "') )"
    End If

    If OptCliente.Value = True Then '--ventas
        '--
        If IdCliPro <> 0 Then
            nSQLWhere = " and vta_ventas.idcli = " & IdCliPro & " "
            '--documentos de apertura
            If OptAperturaCon.Value = True Then nSQLApertura = " or (vta_ventas.numreg='000001' " & nSQLWhere & " ) "
        Else
            If OptAperturaCon.Value = True Then nSQLApertura = " or vta_ventas.numreg='000001' "
        End If
            
        If OptAperturaSin.Value = True Then nSQLApertura = " and vta_ventas.numreg<>'000001' "
        If OptAperturaSolo.Value = True Then nSQLApertura = " and vta_ventas.numreg='000001' "
        '-----------------------------
        '--Listado de facturacion, se incluye nc que esten den tesoreria origen ingreso
        
        nSQL = "SELECT vta_ventas.tipdoc,vta_ventas.id,IIf(vta_ventas!anulado=-1,' ',mae_cliente!numruc) AS numruc, IIf(vta_ventas!anulado=-1,'Anulado',mae_cliente!nombre) AS nombre, IIf(vta_ventas.numreg Is Null Or vta_ventas.numreg='',mae_libros.codsun,Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4)) AS registro," _
            + vbCr + " 'Ventas' AS libro, mae_documento.codsun,mae_documento.abrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc2, vta_ventas.fchdoc, vta_ventas.fchven, mae_moneda.simbolo, " _
            + vbCr + " iif(vta_ventas.tc is null or vta_ventas.tc=0,con_tc.impven , vta_ventas.tc) AS tipcam,vta_ventas.idmon, " _
            + vbCr + " IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc) AS imptotal,vta_ventas.impsal, " _
            + vbCr + " IIf(imptotal=0,0,IIf(vta_ventas.idmon=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
            + vbCr + " IIf(imptotal=0,0,IIf(vta_ventas.idmon=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
            + vbCr + " vta_ventas.glosa as glosaope " _
            + vbCr + " FROM ( ((((vta_ventas LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) " _
            + vbCr + "         LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id " _
            + vbCr + "       ) LEFT JOIN " _
            + vbCr + " ( SELECT tes_cajaorigendet.iddoc FROM tes_cajaorigendet INNER JOIN vta_ventas ON tes_cajaorigendet.iddoc = vta_ventas.id WHERE (((tes_cajaorigendet.idmod)=2) AND ((vta_ventas.tipdoc)=7)) GROUP BY tes_cajaorigendet.iddoc  " _
            + vbCr + "  ) as tes " _
            + vbCr + " ON vta_ventas.id = tes.iddoc " _
            + vbCr + " WHERE ( (vta_ventas.tipdoc<>7 AND vta_ventas.anulado=0 AND IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc)<>0) OR " _
            + vbCr + "        (vta_ventas.tipdoc=7 AND vta_ventas.anulado=0 AND IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc)<>0 AND vta_ventas.iddocref=0) OR " _
            + vbCr + "        (tes.iddoc is not null) ) " & nSQLFecha & nSQLWhere & nSQLApertura _
            + vbCr + " ORDER BY IIf(vta_ventas.anulado=-1,'Anulado',mae_cliente!nombre), vta_ventas!numser+'-'+vta_ventas!numdoc;"
    
    
    ElseIf OptProvee.Value = True Then '--compras
        '--Sentencia SQL para filtrar el proveedor
        If IdCliPro <> 0 Then
            nSQLWhere = " and com_compras.idpro = " & IdCliPro & " "
            '--documentos de apertura
            If OptAperturaCon.Value = True Then nSQLApertura = " or (com_compras.numreg='000001' " & nSQLWhere & " ) "
        Else
            If OptAperturaCon.Value = True Then nSQLApertura = " or com_compras.numreg='000001' "
        End If
        
        '--documentos de apertura
        If OptAperturaSin.Value = True Then nSQLApertura = " and com_compras.numreg<>'000001' "
        If OptAperturaSolo.Value = True Then nSQLApertura = " and com_compras.numreg='000001' "
        
        '--reemplazar filtro de fecha "vta_ventas "por compras "com_compras"
        nSQLFecha = Replace(nSQLFecha, "vta_ventas", "com_compras")
        '------------
        
        nSQL = "SELECT  com_compras.tipdoc,com_compras.id,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_compras.numreg Is Null Or com_compras.numreg='',mae_libros.codsun ,Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4)) AS registro, " _
            + vbCr + " 'Compras' AS libro, mae_documento.codsun,mae_documento.abrev, iif(com_compras!numser is null or com_compras!numser ='','',com_compras!numser  +'-' ) + com_compras!numdoc AS numdoc2, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, " _
            + vbCr + " iif(com_compras.tc is null or com_compras.tc=0,con_tc.impven , com_compras.tc) AS tipcam,com_compras.idmon, " _
            + vbCr + " IIf(com_compras.numreg='000001',com_compras.imptotori,com_compras.imptot) AS imptotal, com_compras.impsal, " _
            + vbCr + " IIf(imptotal=0,0,IIf(com_compras.idmon=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
            + vbCr + " IIf(imptotal=0,0,IIf(com_compras.idmon=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
            + vbCr + " com_compras.glosa as glosaope " _
            + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = com_compras.idmon) " _
            + vbCr + "         LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
            + vbCr + "       ) LEFT JOIN " _
            + vbCr + " ( SELECT tes_cajaorigendet.iddoc FROM tes_cajaorigendet INNER JOIN com_compras ON tes_cajaorigendet.iddoc = com_compras.id WHERE (((tes_cajaorigendet.idmod)=1) AND ((com_compras.tipdoc)=7)) " _
            + vbCr + "  ) as tes ON com_compras.id=tes.iddoc " _
            + vbCr + " WHERE ( (IIf(com_compras.numreg='000001',com_compras.imptotori,com_compras.imptot)<>0) AND " _
            + vbCr + "          (com_compras.tipdoc <>7 or com_compras.tipdoc=7 AND com_compras.iddocref=0) OR " _
            + vbCr + "          (tes.iddoc is not null) ) " & nSQLFecha & nSQLWhere & nSQLApertura & "" _
            + vbCr + " ORDER BY mae_prov!nombre, com_compras.fchdoc "
                      
        '--percepciones
        '--Sentencia SQL para filtrar el proveedor de la percepcion
        If IdCliPro <> 0 Then
            nSQLWhere = " and con_percepcion.idcli = " & IdCliPro & " "
            '--documentos de apertura
            If OptAperturaCon.Value = True Then nSQLApertura = " or (con_percepcion.numreg='000001' " & nSQLWhere & " ) "
        Else
            If OptAperturaCon.Value = True Then nSQLApertura = " or con_percepcion.numreg='000001' "
        End If
        
        '--documentos de apertura
        If OptAperturaSin.Value = True Then nSQLApertura = " and con_percepcion.numreg<>'000001' "
        If OptAperturaSolo.Value = True Then nSQLApertura = " and con_percepcion.numreg='000001' "
        
        '--reemplazar filtro de fecha "com_compras "por percepcion "con_percepcion"
        nSQLFecha = Replace(nSQLFecha, "com_compras", "con_percepcion")
        '------------
        
        nSQL = nSQL + vbCr + " Union " _
            + vbCr + "SELECT con_percepcion.tipdoc,con_percepcion.id & '' AS id, mae_prov.numruc, mae_prov.nombre, Mid(con_percepcion!numreg,1,2)+mae_libros.codsun+Mid(con_percepcion!numreg,3,4) AS registro, 'Percepciones' AS libro, mae_documento.codsun, mae_documento.abrev, [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc2, con_percepcion.fchdoc ,con_percepcion.fchdoc as fchven, mae_moneda.simbolo, " _
            + vbCr + " con_tc.impven AS tipcam, con_percepcion.idmon, con_percepcion.imptotper AS imptotal,con_percepcion.impsal, " _
            + vbCr + " IIf(imptotal=0,0,IIf([con_percepcion].[idmon]=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
            + vbCr + " IIf(imptotal=0,0,IIf([con_percepcion].[idmon]=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
            + vbCr + " con_percepcion.glosa AS glosaope " _
            + vbCr + " FROM (mae_moneda RIGHT JOIN (((con_percepcion LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id) LEFT JOIN mae_libros ON con_percepcion.idlib = mae_libros.id) LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) ON mae_moneda.id = con_percepcion.idmon) LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha " _
            + vbCr + " WHERE (((con_percepcion.tipo)=1)) " & nSQLFecha & nSQLWhere & nSQLApertura

        '--tabla visual que permitira dar un orden a la consulta
        nSQL = "SELECT tab.* FROM ( " & nSQL & " ) AS tab ORDER BY tab.nombre,tab.numdoc2 "

    
    ElseIf opt4ta.Value = True Then '--honorarios
        '--Sentencia SQL para filtrar el prestador de servicio
        If IdCliPro <> 0 Then
            nSQLWhere = " and com_honorarios.idpro = " & IdCliPro & " "
            '--documentos de apertura
            If OptAperturaCon.Value = True Then nSQLApertura = " or (com_honorarios.numreg='000001' " & nSQLWhere & " ) "
        Else
            If OptAperturaCon.Value = True Then nSQLApertura = " or com_honorarios.numreg='000001' "
        End If
        '--documentos de apertura
        
        If OptAperturaSin.Value = True Then nSQLApertura = " and com_honorarios.numreg<>'000001' "
        If OptAperturaSolo.Value = True Then nSQLApertura = " and com_honorarios.numreg='000001' "
        
        '--reemplazar filtro de fecha "vta_ventas "por compras "com_compras"
        nSQLFecha = Replace(nSQLFecha, "vta_ventas", "com_honorarios")
        '------------
        nSQL = "SELECT com_honorarios.tipdoc, com_honorarios.id,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_honorarios.numreg Is Null Or com_honorarios.numreg='',mae_libros.codsun ,Left(com_honorarios.numreg,2) & mae_libros.codsun & Right(com_honorarios.numreg,4)) AS registro, 'Honorario' AS libro, mae_documento.codsun,mae_documento.abrev, com_honorarios!numser+'-'+com_honorarios!numdoc AS numdoc2, com_honorarios.fchdoc, com_honorarios.fchven, mae_moneda.simbolo, " _
            + vbCr + " iif(com_honorarios.tc is null or com_honorarios.tc=0,con_tc.impven , com_honorarios.tc) AS tipcam,com_honorarios.idmon, " _
            + vbCr + " IIf(com_honorarios.numreg='000001',com_honorarios.imptotori,com_honorarios.imptot) AS imptotal, com_honorarios.impsal, " _
            + vbCr + " IIf(imptotal=0,0,IIf([com_honorarios].[idmon]=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
            + vbCr + " IIf(imptotal=0,0,IIf([com_honorarios].[idmon]=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
            + vbCr + " com_honorarios.glosa as glosaope " _
            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_honorarios LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) ON mae_documento.id = com_honorarios.tipdoc) ON mae_prov.id = com_honorarios.idpro) ON mae_moneda.id = com_honorarios.idmon) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha " _
            + vbCr + " WHERE  (IIf([com_honorarios].[numreg]='000001',[com_honorarios].[imptotori],[com_honorarios].[imptot])<>0) and ( com_honorarios.tipdoc <> 7) " & nSQLFecha & nSQLWhere & nSQLApertura _
            + vbCr + " ORDER BY mae_prov!nombre, com_honorarios.fchdoc;"

    Else
    
        Exit Sub

    End If
    
    
    '--indicar el campo a mostrar segun la moneda seleccionada
    If NulosN(TxtIdMon.Text) = 1 Then
        nCampoMuestra = "imptotsol"
    ElseIf NulosN(TxtIdMon.Text) = 2 Then
        nCampoMuestra = "imptotdol"
    Else
        fraBarra.Visible = False
        MsgBox "Por el momento no se puede expresar en " & LblMoneda.Caption, vbInformation, xTitulo
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    '--ejecutar la conulta
    RST_Busq rst, nSQL, xCon
    
    '--filtrar lo que se va mostrar
    If chk_descuadrado.Value = 0 Then
        '--obs. si selecciona la opcion todos no hace el fintro
        If OptPen.Value = True Then rst.Filter = "impsal > 0" ' FILTRAMOS LOS PENDIENTE
        If OptCan.Value = True Then rst.Filter = "impsal <= 0" ' FILTRAMOS LOS CANCELADOS
    End If
    
    If rst.RecordCount = 0 Then
        MsgBox "No hay documentos del cliente seleccionado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fraBarra.Visible = False
        Set rst = Nothing
        Exit Sub
    End If
    
    '--aplicando orden
    
    '--si muestra todos los clientes,proveedores
    If OptSel1.Value = True Then
        If Opt_Orden(0).Value = True Then '--numero doc
            rst.Sort = "nombre,numdoc2"
        ElseIf Opt_Orden(1).Value = True Then '--registro
            rst.Sort = "nombre,registro"
        ElseIf Opt_Orden(2).Value = True Then '--fecha doc
            rst.Sort = "nombre,fchdoc,numdoc2"
        Else
            rst.Sort = "nombre,numdoc2,fchdoc"
        End If
    Else
        If Opt_Orden(0).Value = True Then '--numero doc
            rst.Sort = "numdoc2"
        ElseIf Opt_Orden(1).Value = True Then '--registro
            rst.Sort = "registro"
        ElseIf Opt_Orden(2).Value = True Then '--fecha doc
            rst.Sort = "fchdoc,numdoc2"
        Else
            rst.Sort = "numdoc2,fchdoc"
        End If
    End If
    '-------------------------------------
    
    ProgressBar1.Max = rst.RecordCount
    
    Dim xSaldoDoc As Double
    Dim xFilaIni&
    Dim xFilaIniGrupo As Long '--almacena la fila de inicio de grupo proveedor/ciente/prestador servicio
    Dim xColor&
    
    Me.MousePointer = vbHourglass
     
    xColor = 0
    If rst.RecordCount <> 0 Then
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo Salir:

        rst.MoveFirst
        xSaldoDoc = 0
        xNomPro = NulosC(rst("nombre"))
        xFila = Fg1.FixedRows
        
        '------------------------------------------------------------------------
        '--colocar datos del grupo(cliente, proveedor o prestador de servicio
        '--detalle
        Fg1.Rows = Fg1.Rows + 1

        GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 12, "N� R.U.C. : " & RellenarBlancos(NulosC(rst("numruc")), 12, 1) & "  " & xNomPro, flexAlignLeftCenter, True, , , , True
        xFilaIniGrupo = Fg1.Rows - 1
        
        '--resumen
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(rst("numruc"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(rst("nombre"))
        '------------------------------------------------------------------------
        
        xFilaIni = xFila
        
        'dar formato a la fila
        With Fg1
            .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H800000
            .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            .FillStyle = flexFillRepeat
            .CellFontBold = True
        End With
    
        TotDebe = 0
        TotHaber = 0
        
        Cambio = False
        
        Dim mRowIni As Integer
        '--considerar los id de libros
        '         1=Compras
        '         2=Ventas
        '         5=Igv Retenciones;
        '         6=Bancos;
        '         8=Canjes de Facturas
        '         39=Rendici�n de Cuentas
        '         40=Registro de Honorarios Profesionales
        
        
        
        '--cargando los abonos
        If OptCliente.Value = True Then
            If IdCliPro <> 0 Then nSQLWhere = " and con_diario.ridper = " & IdCliPro & " "
            
            nSQL = "SELECT con_diario.rregistro, Format(con_diario.idmes,'00') & mae_libros.codsun & Format(con_diario.numasi,'0000') AS registro, mae_libros.descripcion AS libro, '' AS codsun, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, IIf(con_diario.aplicatc=0,con_tc.impven,con_diario.tc) AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, IIf(con_diario.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, IIf(con_diario.idmon=2,imptotal, iif(tipcam=0,0,imptotal/tipcam)) AS imptotdol, con_diario.ridper,con_diario.rnumerodoc AS numdoc2,con_diario.rglosaope, con_diario.iddoc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, con_diario.idlib, con_diario.tipmov, con_diario.tipo, con_diario.rtipdoc " _
                + vbCr + " FROM ((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) INNER JOIN con_planctas ON con_diario.idcue = con_planctas.id " _
                + vbCr + " WHERE (((con_diario.idlib) In (5,6,8,37,44)) AND ((con_diario.ridlib)=2)) " & nSQLAjuste & nSQLWhere

            '--unido a referencias de nota de credito
            If IdCliPro <> 0 Then nSQLWhere = " and vta_ventas.idcli = " & IdCliPro & " "
            '--el tipo de cambio de la NC se obtendra del documento de referencia, en caso de no tener ingresado manualmente
            nSQL = nSQL + vbCr + " Union All " _
                + vbCr + "SELECT Left(vta_ventas_1.numreg,2) & mae_libros_1.codsun & Right(vta_ventas_1.numreg,4) AS rregistro, Mid(vta_ventas!numreg,1,2) & mae_libros!codsun & Mid(vta_ventas!numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_libros.codsun, mae_documento.abrev, vta_ventas!numser & '-' & vta_ventas!numdoc AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, IIf(vta_ventas_1.tc=0,con_tc.impven,vta_ventas_1.tc) AS tipcam, vta_ventas.imptotdoc AS imptotdocal, IIf(vta_ventas!idmon=1,vta_ventas!imptotdoc,vta_ventas!imptotdoc*tipcam) AS imptotdocsol, " _
                + vbCr + " IIf(vta_ventas!idmon=2,vta_ventas!imptotdoc,iif(tipcam=0,0,vta_ventas!imptotdoc/tipcam) ) AS imptotdocdol, vta_ventas.idcli AS ridper,  vta_ventas_1.numser & '-' & vta_ventas_1.numdoc AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc ,2 as idlib, 0 as tipmov, 0 as tipo, vta_ventas.tipdoc as rtipdoc  " _
                + vbCr + " FROM ( (((mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) " _
                + vbCr + "         INNER JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id) LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_ventas.idmon = mae_documentocta.idmon) AND (vta_ventas.tipdoc = mae_documentocta.iddoc) " _
                + vbCr + "       ) LEFT JOIN " _
                + vbCr + " ( SELECT tes_cajaorigendet.iddoc FROM tes_cajaorigendet INNER JOIN vta_ventas ON tes_cajaorigendet.iddoc = vta_ventas.id WHERE (((tes_cajaorigendet.idmod)=2) AND ((vta_ventas.tipdoc)=7)) GROUP BY tes_cajaorigendet.iddoc " _
                + vbCr + "   ) as tes ON vta_ventas.id = tes.iddoc " _
                + vbCr + " WHERE vta_ventas.anulado=0 and vta_ventas.tipdoc=7 and vta_ventas.iddocref <> 0 and mae_documentocta.tipope =-1 and  tes.iddoc is null " & nSQLWhere
        
            RST_Busq Rstabo, nSQL, xCon
            
        ElseIf OptProvee.Value = True Then
            '--buscar de bancos, canjes de documementos , rendicion de cuenta
            
            If IdCliPro <> 0 Then nSQLWhere = " and con_diario.ridper = " & IdCliPro & " "
            nSQL = "SELECT con_diario.rregistro, Format(con_diario.idmes,'00') & mae_libros.codsun & Format(con_diario.numasi,'0000') AS registro, mae_libros.descripcion AS libro, '' AS codsun, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, IIf(con_diario.aplicatc=0,con_tc.impven,con_diario.tc) AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, IIf(con_diario.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, IIf(con_diario.idmon=2 ,imptotal,iif(tipcam=0,0, imptotal/tipcam)) AS imptotdol, con_diario.ridper,con_diario.rnumerodoc AS numdoc2,con_diario.rglosaope,con_diario.iddoc , con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, con_diario.idlib,con_diario.tipmov,  con_diario.tipo, con_diario.rtipdoc " _
                + vbCr + " FROM (((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN tes_cajadestinodet ON (con_diario.idmov = tes_cajadestinodet.idtes) AND (con_diario.iddocpro = tes_cajadestinodet.corr)) LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id " _
                + vbCr + " WHERE (((con_diario.idlib) In (6,8,39,44)) AND ((con_diario.ridlib) in (1,4))) " & nSQLAjuste & nSQLWhere

            '--unido a referencias de nota de credito
            If IdCliPro <> 0 Then nSQLWhere = " and com_compras.idpro = " & IdCliPro & " "
           nSQL = nSQL + vbCr + " Union All " _
                + vbCr + "SELECT Left(com_compras_1.numreg,2) & mae_libros_1.codsun & Right(com_compras_1.numreg,4) AS rregistro, Mid(com_compras!numreg,1,2) & mae_libros!codsun & Mid(com_compras!numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_libros.codsun, mae_documento.abrev, com_compras!numser & '-' & com_compras!numdoc AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, IIf(com_compras_1.tc=0,con_tc.impven,com_compras_1.tc) AS tipcam, com_compras.imptot AS imptotal, IIf(com_compras!idmon=1,com_compras!imptot,com_compras!imptot*tipcam) AS imptotsol, IIf(com_compras!idmon=2,com_compras!imptot,iif(tipcam=0,0,com_compras!imptot/tipcam)) AS imptotdol, com_compras.idpro AS ridper,  com_compras_1.numser & '-' & com_compras_1.numdoc AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, 1 as idlib, 0 as tipmov, 0 as tipo, com_compras.tipdoc as rtipdoc  " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN com_compras AS com_compras_1 ON com_compras.iddocref = com_compras_1.id) LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) " _
                + vbCr + "       LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (com_compras.idmon = mae_documentocta.idmon) AND (com_compras.tipdoc = mae_documentocta.iddoc) " _
                + vbCr + "        ) LEFT JOIN " _
                + vbCr + " ( SELECT tes_cajaorigendet.iddoc FROM tes_cajaorigendet INNER JOIN com_compras ON tes_cajaorigendet.iddoc = com_compras.id WHERE (((tes_cajaorigendet.idmod)=1) AND ((com_compras.tipdoc)=7)) " _
                + vbCr + "   ) as tes ON com_compras.id=tes.iddoc  " _
                + vbCr + " WHERE com_compras.iddocref Is Not Null And com_compras.iddocref<>0 and mae_documentocta.tipope=0 and tes.iddoc is null " & nSQLWhere
            
            RST_Busq Rstabo, nSQL, xCon
            
        ElseIf opt4ta.Value = True Then
            If IdCliPro <> 0 Then nSQLWhere = " and con_diario.ridper = " & IdCliPro & " "
            
            nSQL = "SELECT con_diario.rregistro, Format(con_diario.idmes,'00') & mae_libros.codsun & Format(con_diario.numasi,'0000') AS registro, mae_libros.descripcion AS libro, '' AS codsun, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, IIf(con_diario.aplicatc=0,con_tc.impven,con_diario.tc) AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, IIf(con_diario.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, IIf(con_diario.idmon=2,imptotal,iif(tipcam=0,0,imptotal/tipcam)) AS imptotdol, con_diario.ridper,con_diario.rnumerodoc AS numdoc2,con_diario.rglosaope,con_diario.iddoc , con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, con_diario.idlib, con_diario.tipmov, con_diario.tipo, con_diario.rtipdoc  " _
                + vbCr + " FROM ((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id " _
                + vbCr + " WHERE (((con_diario.idlib) In (6,8,39,44)) AND ((con_diario.ridlib)=40)) " & nSQLAjuste & nSQLWhere
            
            RST_Busq Rstabo, nSQL, xCon
            
        End If
                       
        
        '--------------
        rst.MoveFirst
        For A = 1 To rst.RecordCount    '--GRUPO DE CLIENTE/PROVEEDOR
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo Salir:
            ProgressBar1.Value = A
            
            xSaldoDoc = 0
            
            If NulosC(rst("nombre")) <> xNomPro Then
                DoEvents
                Cambio = True
                xNomPro = NulosC(rst("nombre"))
                Fg1.Rows = Fg1.Rows + 1
                xFila = xFila + 1
                Fg1.TextMatrix(xFila, 4) = "TOTAL -->"
                
                '--acumulando los totales por grupo
                TotDebe = NulosN(GRID_SUMAR_COL(Fg1, 10, xFilaIniGrupo, Fg1.Rows - 2))
                TotHaber = NulosN(GRID_SUMAR_COL(Fg1, 11, xFilaIniGrupo, Fg1.Rows - 2))
                
                Fg1.TextMatrix(xFila, 10) = Format(TotDebe, FORMAT_MONTO)
                Fg1.TextMatrix(xFila, 11) = Format(TotHaber, FORMAT_MONTO)
                If OptCliente.Value = True Then
                    Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
                Else
                    Fg1.TextMatrix(xFila, 12) = Format(TotHaber - TotDebe, FORMAT_MONTO)
                End If
                
                '*****resumen
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotDebe, FORMAT_MONTO)
                Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotHaber, FORMAT_MONTO)
                
                If OptCliente.Value = True Then
                    Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
                Else
                    Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotHaber - TotDebe, FORMAT_MONTO)
                End If
                '******
                
                With Fg1
                    .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
                    .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                    .FillStyle = flexFillRepeat
                    .CellFontBold = True
                End With
                
                '----MOSTRAR SOLO DESCUADRADOS ---------
                If chk_descuadrado.Value = 1 Then
                    If NulosN(Fg1.TextMatrix(xFila, 12)) = 0 Then
                        GRID_DELETE Fg1, Fg1.Rows - 2, Fg1.Rows - 1, e_Fila
                        Fg1.Rows = Fg1.Rows + 1
                        xFila = Fg1.Rows - 1
                    Else
                        Fg1.Rows = Fg1.Rows + 2
                        xFila = xFila + 2
                    End If
                    '---del resumen
                    If NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 5)) = 0 Then
                        GRID_DELETE Fg2, Fg2.Rows - 1, Fg2.Rows - 1, e_Fila
                    End If
                    '---------------
                Else
                    Fg1.Rows = Fg1.Rows + 2
                    xFila = xFila + 2
                End If
                '---------------------------------------------------------
                TotGralHaber = TotGralHaber + TotHaber
                TotGralDebe = TotGralDebe + TotDebe
                
                TotHaber = 0
                TotDebe = 0
                '---------------------------------------------------------

                GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 12, "N� R.U.C. : " & RellenarBlancos(NulosC(rst("numruc")), 12, 1) & "  " & xNomPro, flexAlignLeftCenter, True, , , , True
                xFilaIniGrupo = Fg1.Rows - 1
                
                '*****resumen
                Fg2.Rows = Fg2.Rows + 1
                Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(rst("numruc"))
                Fg2.TextMatrix(Fg2.Rows - 1, 2) = xNomPro
                '******

                
                With Fg1
                    .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H800000
                    .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                    .FillStyle = flexFillRepeat
                    .CellFontBold = True
                End With
            Else
                Cambio = False
            End If

            Fg1.Rows = Fg1.Rows + 1
            xFila = xFila + 1
            xFilaIni = xFila
            
            Fg1.TextMatrix(xFila, 1) = NulosC(rst("registro"))
            
            Fg1.TextMatrix(xFila, 2) = NulosC(rst("libro"))
            Fg1.TextMatrix(xFila, 3) = NulosC(rst("abrev"))
            Fg1.TextMatrix(xFila, 4) = NulosC(rst("numdoc2"))
            Fg1.TextMatrix(xFila, 5) = Format(rst("fchdoc"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 6) = Format(rst("fchven"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 7) = NulosC(rst("simbolo"))
            Fg1.TextMatrix(xFila, 8) = Format(NulosN(rst("imptotal")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 9) = Format(NulosN(rst("tipcam")), "###0.##0") & ""
            
            If OptCliente.Value = True Then
                If NulosN(rst("tipdoc")) <> 7 Then
                    Fg1.TextMatrix(xFila, 10) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO)
                    TotDebe = TotDebe + NulosN(rst(nCampoMuestra))
                Else
                    Fg1.TextMatrix(xFila, 11) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO) '--saldo
                    TotHaber = TotHaber + NulosN(rst(nCampoMuestra))
                End If
            Else
            
                If rst("tipdoc") <> 7 Then
                    Fg1.TextMatrix(xFila, 11) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO) '--saldo
                    TotHaber = TotHaber + NulosN(rst(nCampoMuestra))
                Else
                    Fg1.TextMatrix(xFila, 10) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO)
                    TotDebe = TotDebe + NulosN(rst(nCampoMuestra))
                End If
                
                
            End If
            
            Fg1.TextMatrix(xFila, 12) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO)
            
            Fg1.TextMatrix(xFila, 13) = NulosC(rst("numdoc2"))
            Fg1.TextMatrix(xFila, 14) = NulosC(rst("glosaope"))
            
            xSaldoDoc = NulosN(rst("impsal"))
            
            
            '-------------------------------------------------------------
            '--filtrar los movimientos de las provisiones para proceder a obtener el saldo actual
                        
            '--ventas
            If OptCliente.Value = True Then
                Rstabo.Filter = "rregistro = '" & rst("registro") & "' and iddoc= " & rst("id")
            '--compras
            ElseIf OptProvee.Value = True Then '--proveedor
                Rstabo.Filter = "rregistro = '" & rst("registro") & "' and iddoc= " & rst("id")
            '--honorarios
            ElseIf opt4ta.Value = True Then '--honorarios
                Rstabo.Filter = "rregistro = '" & rst("registro") & "' and iddoc= " & rst("id")
            End If

            '-------------------------------------------------------------
            
            
            If Rstabo.RecordCount <> 0 Then
                Rstabo.MoveFirst
                '--ordenar el rst para mostrar el detalle
                '--NOta: el primer orden es importante pues indica que el ajuste por diferencia de cambio se mostrara en la ultima posicion del detalle,
                '--      esto indicara que se esta aplicando un ajuste al documento para mostrar el saldo a cero.
                Rstabo.Sort = "libro desc,fchemi ASC"
                
                Do While Not Rstabo.EOF
                    '--SI SE NTERRUMPE EL PROCESO => SALIR
'                    DoEvents
                    If BAND_INTERRUMPIR = True Then GoTo Salir:
                    Fg1.Rows = Fg1.Rows + 1
                    xFila = xFila + 1
                    
                    Fg1.TextMatrix(xFila, 1) = NulosC(Rstabo("registro"))
                    
                    Fg1.TextMatrix(xFila, 2) = NulosC(Rstabo("libro"))
                    Fg1.TextMatrix(xFila, 3) = NulosC(Rstabo("abrev"))
                    Fg1.TextMatrix(xFila, 4) = NulosC(Rstabo("numdoc"))
                    Fg1.TextMatrix(xFila, 5) = Format(Rstabo("fchemi"), FORMAT_DATE)
                    Fg1.TextMatrix(xFila, 7) = NulosC(Rstabo("simbolo"))
                    Fg1.TextMatrix(xFila, 8) = Format(NulosN(Rstabo("imptotal")), FORMAT_MONTO)
                    Fg1.TextMatrix(xFila, 9) = Format(NulosN(Rstabo("tipcam")), "####.###")
                    
                    '--verificar si el libro es de ajuste por dif de cambio
                    If InStr(LCase(Rstabo("libro")), "ajuste") <> 0 Then
                        
                        If OptCliente.Value = True Then
                            '--verificar si es perdida o ganancia
                            If NulosN(Fg1.TextMatrix(xFila - 1, 12)) > 0 Then
                                Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
                                '--obteniendo utlimo saldo
                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                
                            Else
                                Fg1.TextMatrix(xFila, 10) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                TotDebe = TotDebe + NulosN(Rstabo(nCampoMuestra))
                                '--obteniendo utlimo saldo
                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) + NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                
                            End If
                        
                        Else
                            '--verificar si es perdida o ganancia
                            If NulosN(Fg1.TextMatrix(xFila - 1, 12)) < 0 Then
                                Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
                                '--obteniendo utlimo saldo
                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) + NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                
                            Else
                                Fg1.TextMatrix(xFila, 10) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                TotDebe = TotDebe + NulosN(Rstabo(nCampoMuestra))
                                '--obteniendo utlimo saldo
                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                
                            End If
                        End If

                    Else
                    
                        If OptCliente.Value = True Then
                            '--verificar si es egreso-destino(Devoluciones por pago en exceso del cliente)
                            If NulosN(Rstabo("tipmov")) = 2 And NulosN(Rstabo("tipo")) = 2 Then
                                Fg1.TextMatrix(xFila, 10) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                '--verificar si no es nota credito
                                If NulosN(Rstabo("rtipdoc")) <> 7 Then
                                    TotDebe = TotDebe + NulosN(Rstabo(nCampoMuestra))
                                    '--obteniendo utlimo saldo
                                    Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) + NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                Else
                                    TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
                                    '--obteniendo utlimo saldo
                                    Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                End If
                                
                            Else
                                '--Ingresos, tambien se muestra egresos-origen x canje de documentos
                                '--verificar si no es nota credito
                                If NulosN(Rstabo("rtipdoc")) <> 7 Then
                                    Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                    TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
                                Else
                                    '--verificar si nota credito hace referencia a un documento o proviene de bancos
                                    If NulosN(Rstabo("idlib")) = 2 Then
                                        Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                        TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
                                    Else
                                        Fg1.TextMatrix(xFila, 10) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                        TotDebe = TotDebe + NulosN(Rstabo(nCampoMuestra))
                                    End If
                                End If
                                '--obteniendo utlimo saldo
                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                            
                            End If
                                                    
                        ElseIf OptProvee.Value = True Then
                        
''                            If Rstabo("iddoc") = 1125 Then
''                                MsgBox ""
''                            End If
                        
                            '--verificar si es ingreso-destino(Devoluciones por pago en exceso al proveedor)
                            If NulosN(Rstabo("tipmov")) = 1 And NulosN(Rstabo("tipo")) = 2 Then
                                Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                '--verificar si no es nota credito
                                If NulosN(Rstabo("rtipdoc")) <> 7 Then
                                    TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
                                    '--obteniendo utlimo saldo
                                    Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) + NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                Else
                                    TotDebe = TotDebe + NulosN(Rstabo(nCampoMuestra))
                                    '--obteniendo utlimo saldo
                                    Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                
                                End If
                            Else
                                '--verificar si no es nota credito
                                If NulosN(Rstabo("rtipdoc")) <> 7 Then
                                    Fg1.TextMatrix(xFila, 10) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                    TotDebe = TotDebe + NulosN(Rstabo(nCampoMuestra))
                                Else
                                    '--verificar si nota credito hace referencia a un documento o proviene de bancos
                                    If NulosN(Rstabo("idlib")) = 1 Then
                                        Fg1.TextMatrix(xFila, 10) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                        TotDebe = TotDebe + NulosN(Rstabo(nCampoMuestra))
                                    Else
                                        Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                                        TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
                                    End If
                                    
                                    
                                End If
                                '--obteniendo utlimo saldo
                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                            
                            
                            End If

                        Else
                        
                        
                            Fg1.TextMatrix(xFila, 10) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                            TotDebe = TotDebe + NulosN(Rstabo(nCampoMuestra))
                            
                            '--obteniendo utlimo saldo
                            Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                        
                        
                        
                        
                        End If
                        
                    End If
                   
                    Fg1.TextMatrix(xFila, 13) = NulosC(Rstabo("numdoc2"))
                    Fg1.TextMatrix(xFila, 14) = NulosC(Rstabo("rglosaope"))
                                        
                    Rstabo.MoveNext
                Loop
                
            End If
            
            '---ACTUALIZANDO EL SALDO AL DOCUMENTO
            '--solo se actualizara el saldo si el documento esta en la moneda de consulta
            '--considerar actualizar el saldo si es ajuste por diferencia de cambio sin importar la moneda
            If (xSaldoDoc <> NulosN(Fg1.TextMatrix(xFila, 12)) And NulosN(rst("idmon")) = NulosN(TxtIdMon.Text)) Or InStr(LCase(Fg1.TextMatrix(xFila, 2)), "ajuste") <> 0 Then
                
                '--obtener el ultimo saldo del documento
                If InStr(LCase(Fg1.TextMatrix(xFila, 2)), "ajuste") <> 0 Then
                    sSaldoFinal = 0
                Else
                    sSaldoFinal = NulosN(Fg1.TextMatrix(xFila, 12))
                End If
                '--------------------------------------------------------
                
                If OptCliente.Value = True Then     '--VENTAS
                    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & sSaldoFinal & " WHERE (((vta_ventas.id)=" & rst("id") & "))"
                    
                ElseIf OptProvee.Value = True Then                                '--COMPRAS
                    If LCase(NulosC(rst("libro"))) = "percepciones" Then
                        xCon.Execute "UPDATE con_percepcion SET con_percepcion.impsal = " & sSaldoFinal & " WHERE (((con_percepcion.id)=" & rst("id") & "))"
                    Else
                        xCon.Execute "UPDATE com_compras SET com_compras.impsal = " & sSaldoFinal & " WHERE (((com_compras.id)=" & rst("id") & "))"
                    End If
                    
                    
                Else
                    xCon.Execute "UPDATE com_honorarios SET com_honorarios.impsal = " & sSaldoFinal & " WHERE (((com_honorarios.id)=" & rst("id") & "))"
                    
                End If
            End If
            
            '----MOSTRAR SOLO DESCUADRADOS ---------
            If chk_descuadrado.Value = 1 And OptTodos.Value = True Then
                If NulosN(Fg1.TextMatrix(xFila, 12)) >= 0 Then
                'TabOne1.CurrTab = 0
                    GRID_DELETE Fg1, Fg1.Rows - 1 - Rstabo.RecordCount, Fg1.Rows - 1, e_Fila
                    '*********************************************
                    If Rstabo.RecordCount <> 0 Then
                        Rstabo.MoveFirst
                        Do While Not Rstabo.EOF
                            If OptCliente.Value = True Then
                                TotHaber = TotHaber - NulosN(Rstabo(nCampoMuestra))
                            Else
                                TotDebe = TotDebe - NulosN(Rstabo(nCampoMuestra))
                            End If
                            Rstabo.MoveNext
                        Loop
                    End If
                    If OptCliente.Value = True Then
                        TotDebe = TotDebe - NulosN(rst(nCampoMuestra))
                    Else
                        TotHaber = TotHaber - NulosN(rst(nCampoMuestra))
                    End If
                    
                    '*********************************************
                    xFila = Fg1.Rows - 1
                    mRowIni = -1
                Else
                    mRowIni = 0
                End If
            End If
            '---------------------------------------------------------
            
            rst.MoveNext
            If rst.EOF = True Then
                Exit For
            End If
            If mRowIni = 0 Then
                If xColor = 0 Then
                    GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 1, Fg1.Cols - 1, &H80000005
                    xColor = 1
                Else
                    GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HE0DCDA
                    xColor = 0
                End If
            End If
            

        Next A
        
        
        
        
        Fg1.Rows = Fg1.Rows + 1
        xFila = xFila + 1
        Fg1.TextMatrix(xFila, 4) = "TOTAL -->"


        '--acumulando los totales por grupo
        TotDebe = NulosN(GRID_SUMAR_COL(Fg1, 10, xFilaIniGrupo, Fg1.Rows - 2))
        TotHaber = NulosN(GRID_SUMAR_COL(Fg1, 11, xFilaIniGrupo, Fg1.Rows - 2))

        '---------------------------------------------------------
        TotGralHaber = TotGralHaber + TotHaber
        TotGralDebe = TotGralDebe + TotDebe
        '---------------------------------------------------------
        
        Fg1.TextMatrix(xFila, 10) = Format(TotDebe, FORMAT_MONTO)
        Fg1.TextMatrix(xFila, 11) = Format(TotHaber, FORMAT_MONTO)
                
        If OptCliente.Value = True Then
            Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
        Else
            Fg1.TextMatrix(xFila, 12) = Format(TotHaber - TotDebe, FORMAT_MONTO)
        End If
        
        '*****resumen
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotDebe, FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotHaber, FORMAT_MONTO)
        
        If OptCliente.Value = True Then
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
        Else
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotHaber - TotDebe, FORMAT_MONTO)
        End If
        
        
        
        '******

        With Fg1
            .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
            .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            .FillStyle = flexFillRepeat
            .CellFontBold = True
        End With
        '----MOSTRAR SOLO DESCUADRADOS ---------
        If chk_descuadrado.Value = 1 Then
            If NulosN(Fg1.TextMatrix(xFila, 12)) = 0 Then
                GRID_DELETE Fg1, Fg1.Rows - 2, Fg1.Rows - 1, e_Fila
                xFila = Fg1.Rows - 1
            End If
            '--del resumen
            If NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 5)) = 0 Then
                GRID_DELETE Fg2, Fg2.Rows - 1, Fg2.Rows - 1, e_Fila
            End If
            '---------------
        End If
        
        '---------------------------------------------------------

        If TotGralDebe <> 0 Or TotGralHaber <> 0 Then
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = "TOTAL GRAL -->"
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(TotGralDebe, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(TotGralHaber, FORMAT_MONTO)
            
            If OptCliente.Value = True Then
                Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(TotGralHaber - TotGralDebe, FORMAT_MONTO)
            End If
            
            With Fg1
                .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
                .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                .FillStyle = flexFillRepeat
                .CellFontBold = True
            End With
            '*****resumen
            Fg2.Rows = Fg2.Rows + 2
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = "TOTAL GRAL -->"
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotGralDebe, FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotGralHaber, FORMAT_MONTO)
            
            If OptCliente.Value = True Then
                Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
            Else
                Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotGralHaber - TotGralDebe, FORMAT_MONTO)
            End If
            
            With Fg2
                .Cell(flexcpForeColor, Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1) = &H80000008
                .Select Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1
                .FillStyle = flexFillRepeat
                .CellFontBold = True
            End With
            '******
    End If
    
    End If
    If mRowIni = 0 Then
        If xColor = 0 Then
            GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 2, Fg1.Cols - 1, &H80000005
            xColor = 1
        Else
            GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 2, Fg1.Cols - 1, &HE0DCDA
            xColor = 0
        End If
    End If
    
    '--ajustar totales
    '--detalle
    Fg1.AutoSizeMode = flexAutoSizeColWidth
    Fg1.AutoSize 8
    Fg1.AutoSize 10
    Fg1.AutoSize 11
    Fg1.AutoSize 12
    '--resumen
    Fg2.AutoSizeMode = flexAutoSizeColWidth
    Fg2.AutoSize 3
    Fg2.AutoSize 4
    Fg2.AutoSize 5
    '----------------------------------------------------------------------
    
    Set rst = Nothing
    Set Rstabo = Nothing
    fraBarra.Visible = False
    Me.MousePointer = vbDefault
    MsgBox "La Consulta fue se realiz� Correctamente", vbInformation, xTitulo
    Exit Sub
Salir:
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set rst = Nothing
    Set Rstabo = Nothing
    MsgBox "La Consulta fue Interrumpida", vbInformation, xTitulo
    Exit Sub
error:
    
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set rst = Nothing
    Set Rstabo = Nothing
    SHOW_ERROR Me.Name, "CargarCli2"
    
End Sub

Private Sub pConsultar()

    If OptSel1.Value = True Then CargarCli2 0
    If OptSel2.Value = True Then
        TabOne2.CurrTab = 0
        If NulosC(TxtCliPro.Text) = "" Then
            If OptCliente.Value = True Then
                MsgBox "No ha especificado el cliente a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
            If OptProvee.Value = True Then
                MsgBox "No ha especificado el proveedor a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
            If opt4ta.Value = True Then
                MsgBox "No ha especificado el prestador de servicio a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
            TxtCliPro.SetFocus
            Exit Sub
        End If
        CargarCli2 NulosN(LblIdCliPro.Caption)
    End If

End Sub
