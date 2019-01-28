VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCtaCteOtros 
   Caption         =   "Caja y Bancos - Cuenta Corriente Otros"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   12510
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne2 
      Height          =   1275
      Left            =   30
      TabIndex        =   5
      Top             =   360
      Width           =   11910
      _cx             =   21008
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
      Caption         =   "Inicio| Mas"
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
         Caption         =   "Frame7"
         Height          =   1185
         Left            =   345
         TabIndex        =   7
         Top             =   45
         Width           =   11520
         Begin VB.Frame FraReem 
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
            Height          =   525
            Left            =   6870
            TabIndex        =   39
            Top             =   690
            Width           =   2160
            Begin VB.OptionButton OptReem1 
               Caption         =   "Bancos"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   41
               Top             =   210
               Value           =   -1  'True
               Width           =   960
            End
            Begin VB.OptionButton OptReem2 
               Caption         =   "Lgd"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1380
               TabIndex        =   40
               Top             =   240
               Width           =   690
            End
         End
         Begin VB.Frame Frame2 
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
            Height          =   1125
            Left            =   9870
            TabIndex        =   11
            Top             =   0
            Width           =   1650
            Begin VB.OptionButton OptPen 
               Caption         =   "Pendientes"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   225
               TabIndex        =   14
               Top             =   240
               Value           =   -1  'True
               Width           =   1350
            End
            Begin VB.OptionButton OptCan 
               Caption         =   "Cancelados"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   225
               TabIndex        =   13
               Top             =   510
               Width           =   1350
            End
            Begin VB.OptionButton OptTodos 
               Caption         =   "Todos"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   225
               TabIndex        =   12
               Top             =   750
               Width           =   1350
            End
         End
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
            Height          =   825
            Left            =   8280
            TabIndex        =   8
            Top             =   0
            Width           =   1560
            Begin VB.OptionButton OptSel2 
               Caption         =   "Seleccionar"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
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
               Top             =   510
               Value           =   -1  'True
               Width           =   840
            End
         End
         Begin VB.OptionButton Opt4ta 
            Caption         =   "Prestador de Servicio"
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
            Left            =   8820
            TabIndex        =   38
            Top             =   510
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.OptionButton OptProv 
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
            Left            =   8820
            TabIndex        =   37
            Top             =   240
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.OptionButton Optcli 
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
            Left            =   8820
            TabIndex        =   36
            Top             =   30
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.CheckBox chk_descuadrado 
            Caption         =   "Descuadrados"
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   8310
            TabIndex        =   35
            ToolTipText     =   "Mostrará solo los documentos cuyo saldo final es negativo"
            Top             =   960
            Width           =   1545
         End
         Begin VB.OptionButton OptLGD 
            Caption         =   "Liquidacion LGD/LGC"
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
            Left            =   6060
            TabIndex        =   34
            Top             =   230
            Width           =   2190
         End
         Begin VB.OptionButton OptReem 
            Caption         =   "Reembolsables"
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
            Left            =   6060
            TabIndex        =   33
            Top             =   0
            Value           =   -1  'True
            Width           =   1740
         End
         Begin VB.OptionButton OptLetra 
            Caption         =   "Letras"
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
            Left            =   6060
            TabIndex        =   32
            Top             =   460
            Width           =   900
         End
         Begin VB.OptionButton OptPlaLetra 
            Caption         =   "Planilla de Letras"
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
            Left            =   6060
            TabIndex        =   31
            Top             =   690
            Width           =   1890
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
            Height          =   660
            Left            =   2160
            TabIndex        =   20
            Top             =   480
            Width           =   3795
            Begin VB.CommandButton CmdBusMon 
               Height          =   240
               Left            =   1185
               Picture         =   "FrmCtaCteOtros.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   300
               Width           =   210
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   720
               MaxLength       =   1
               TabIndex        =   22
               Text            =   "TxtIdMon"
               Top             =   270
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
               TabIndex        =   24
               Top             =   270
               Width           =   2310
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda"
               Height          =   195
               Index           =   4
               Left            =   90
               TabIndex        =   23
               Top             =   360
               Width           =   585
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Hasta el dia ]"
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
            Height          =   660
            Left            =   60
            TabIndex        =   17
            Top             =   480
            Width           =   2085
            Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
               Height          =   315
               Left            =   660
               TabIndex        =   18
               Top             =   270
               Width           =   1350
               _ExtentX        =   2381
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
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fecha"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   19
               Top             =   315
               Width           =   450
            End
         End
         Begin VB.CommandButton CmdBusCliPro 
            Enabled         =   0   'False
            Height          =   240
            Left            =   5715
            Picture         =   "FrmCtaCteOtros.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   135
            Width           =   210
         End
         Begin VB.TextBox TxtCliPro 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "TxtCliPro"
            Top             =   105
            Width           =   4335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   26
            Top             =   195
            Width           =   480
         End
         Begin VB.Label LblIdCliPro 
            Caption         =   "LblIdCliPro"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   4590
            TabIndex        =   25
            Top             =   30
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   5985
            X2              =   5985
            Y1              =   60
            Y2              =   1140
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            Index           =   3
            X1              =   5985
            X2              =   5985
            Y1              =   60
            Y2              =   990
         End
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   12855
         TabIndex        =   6
         Top             =   45
         Width           =   11520
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
            Height          =   975
            Left            =   180
            TabIndex        =   27
            Top             =   120
            Width           =   1755
            Begin VB.OptionButton Opt_Orden 
               Caption         =   "N° Documento"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   30
               Top             =   240
               Width           =   1425
            End
            Begin VB.OptionButton Opt_Orden 
               Caption         =   "N°. Registro"
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   29
               Top             =   480
               Width           =   1215
            End
            Begin VB.OptionButton Opt_Orden 
               Caption         =   "Fecha Doc."
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   28
               Top             =   720
               Width           =   1185
            End
         End
      End
   End
   Begin VB.Frame fraBarra 
      BorderStyle     =   0  'None
      Caption         =   "FrmConsultaDiario"
      Height          =   780
      Left            =   2760
      TabIndex        =   1
      Top             =   3525
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
      Width           =   12510
      _ExtentX        =   22066
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
               Picture         =   "FrmCtaCteOtros.frx":0264
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":07A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":0B3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":0C94
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":1026
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":11AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":15FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":1716
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":1C5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":219E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":22B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":23C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":281A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCteOtros.frx":2986
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5865
      Left            =   0
      TabIndex        =   42
      Top             =   1650
      Width           =   11910
      _cx             =   21008
      _cy             =   10345
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
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   5445
         Left            =   -12465
         TabIndex        =   43
         Top             =   45
         Width           =   11820
         _cx             =   20849
         _cy             =   9604
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
         FormatString    =   $"FrmCtaCteOtros.frx":2ECE
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
         Height          =   5445
         Left            =   45
         TabIndex        =   44
         Top             =   45
         Width           =   11820
         _cx             =   20849
         _cy             =   9604
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
         FormatString    =   $"FrmCtaCteOtros.frx":3055
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
Attribute VB_Name = "FrmCtaCteOtros"
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

'modificado 15/01/10 por johan castro
'           agregar el filtro po orden
'modificado 26/01/10 por johan castro
'           agregar liquidacion de gasto debito a analisis
'           por espacon lo pongo en opcion "mas"
'modificado 08/09/10 por Johan Castro
'       agregar analisis por reembolsable
'modificado 15/09/10 por Johan Castro
'       En analisis por reembolsable agregar opcion
'       Ver x Bancos => Actualiza saldo
'       Ver x Lgd => No actualiza saldo
'modificado 22/03/11 por Johan Castro
'       Modificar filtro de consulta para dar cancelaciones de lgd, letras
'Modificado: 15/06/11 Johan Castro
'       Dar flexibilidad al formulario para definir el tamaño del mismo

Option Explicit
Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE



Private Sub CmdBusCliPro_Click()
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value Then
        xForm.Titulo = "Buscando Clientes"
        xForm.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente ORDER BY mae_cliente.nombre"
        xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    ElseIf OptProv.Value = True Then
        xForm.Titulo = "Buscando Proveedores"
        xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov ORDER BY mae_prov.nombre"
        xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    ElseIf opt4ta.Value = True Then
        xForm.Titulo = "Buscando Prestador de Servicio"
        xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov WHERE mae_prov.tipper = 1 ORDER BY mae_prov.nombre"
        xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    ElseIf OptPlaLetra.Value = True Then
        xForm.Titulo = "Buscando Banco"
        xForm.SQLCad = "SELECT mae_bancos.numruc, mae_bancos.descripcion as nombre, mae_bancos.id From mae_bancos  ORDER BY mae_bancos.descripcion"
        xCampos(0, 0) = "Banco":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    
    ElseIf OptReem.Value = True Then
        xForm.Titulo = "Buscando Proveedores"
        xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov ORDER BY mae_prov.nombre"
        xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    
    End If

    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"

    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCliPro.Text = xRs("nombre")
        LblIdCliPro.Caption = xRs("id")
        TxtFecha.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        pConfigurarGrilla
        TxtIdMon.Text = 1
        TabOne2.CurrTab = 0
        TxtIdMon_Validate False
        SeEjecuto = True
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
    TxtCliPro.Text = ""
    TxtFecha.Valor = ""
    TxtFecha.Valor = Date
    LblMoneda.Caption = ""
    TxtIdMon.Text = ""
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
End Sub

Private Sub optmod_Click(Index As Integer)
    Select Case Index
        Case 0
            TxtCliPro.Text = ""
            LblIdCliPro.Caption = ""
            Label1(0).Caption = "Cliente"
        
        Case 1
            TxtCliPro.Text = ""
            LblIdCliPro.Caption = ""
            Label1(0).Caption = "Proveedor"
        
        Case 2
        
            TxtCliPro.Text = ""
            LblIdCliPro.Caption = ""
            Label1(0).Caption = "Prestador de Servicio"
    End Select
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
  
    If Me.Height > 3000 Then
        TabOne1.Top = 1920
        TabOne1.Width = Me.Width - 150
        TabOne1.Height = Me.Height - 2320
    End If
End Sub

Private Sub OptCan_Click()
    chk_descuadrado.Enabled = True
End Sub



Private Sub OptLetra_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    Label1(0).Caption = "Cliente"
    FraReem.Visible = False
End Sub

Private Sub OptLGD_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    Label1(0).Caption = "Cliente"
    FraReem.Visible = False
End Sub

Private Sub OptPen_Click()
    chk_descuadrado.Value = 0
    chk_descuadrado.Enabled = False
End Sub

Private Sub OptPlaLetra_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    Label1(0).Caption = "Banco"
    FraReem.Visible = False
End Sub

Private Sub OptReem_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    Label1(0).Caption = "Proveedor"
    FraReem.Visible = True
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
    chk_descuadrado.Enabled = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then pExportar
    If Button.Index = 4 Then pImprimir
    If Button.Index = 6 Then
        Unload Me
    End If
End Sub

Sub pExportar()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    Dim nPeriodo As String
    Dim nTitulo1 As String
    Dim nTipo As String
        
    nPeriodo = "Al  " + CStr(TxtFecha.Valor)

    nTitulo1 = "(Expresado en " & LblMoneda.Caption & ")"
    
    If Optcli.Value = True Then
        nTipo = "Cliente"
    ElseIf OptProv.Value = True Then
        nTipo = "Proveedor"
    ElseIf opt4ta.Value Then
        nTipo = "Prestador de Servicio"
        
    ElseIf OptLGD.Value = True Then
        nTipo = "Cliente"
    End If
    
    If TabOne1.CurrTab = 0 Then '--detalle
        '''oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Cuenta Corriente - "  & nTipo, nPeriodo, nTitulo1, "Cuenta Corriente Análisis"
        
        GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "Cuenta Corriente - " & nTipo, "Expresado en " & LblMoneda.Caption, "Cuenta Corriente Análisis"

    Else
        oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg2, "Resumen de Cuenta Corriente - " & nTipo, nPeriodo, nTitulo1, "Cuenta Corriente Análisis"
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
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
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
    xCampos(0, 0) = "Nº.Registro":      xCampos(0, 1) = "1":    xCampos(0, 2) = "C":    xCampos(0, 3) = "-1"
    xCampos(1, 0) = "Origen":           xCampos(1, 1) = "2":    xCampos(1, 2) = "C":    xCampos(1, 3) = "0"
    xCampos(2, 0) = "Nº Documento":     xCampos(2, 1) = "4":    xCampos(2, 2) = "C":    xCampos(2, 3) = "0"
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
        .Cols = 15
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
        .TextMatrix(1, 1) = "N° Registro":  .ColWidth(1) = 900:   .ColAlignment(1) = flexAlignLeftCenter:     .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 2) = "Origen":       .ColWidth(2) = 1200:   .ColAlignment(2) = flexAlignLeftCenter:     .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "T.D.":         .ColWidth(3) = 450:    .ColAlignment(3) = flexAlignLeftCenter:     .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 4) = "N°.Documento": .ColWidth(4) = 1600:   .ColAlignment(4) = flexAlignLeftCenter:     .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 5) = "Fch.Emi.":     .ColWidth(5) = 800:    .ColAlignment(5) = flexAlignCenterBottom:   .Row = 1: .Col = 5: .CellAlignment = flexAlignCenterBottom
        .TextMatrix(1, 6) = "Fch.Ven.":     .ColWidth(6) = 800:    .ColAlignment(6) = flexAlignCenterBottom:   .Row = 1: .Col = 6: .CellAlignment = flexAlignCenterBottom
        .TextMatrix(1, 7) = "M":            .ColWidth(7) = 450:    .ColAlignment(7) = flexAlignLeftCenter:    .Row = 1: .Col = 7: .CellAlignment = flexAlignCenterBottom
        
        .TextMatrix(1, 8) = "Imp":          .ColWidth(8) = 900:    .ColAlignment(8) = flexAlignRightCenter:   .Row = 1: .Col = 8: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 9) = "T.C.":         .ColWidth(9) = 500:    .ColAlignment(9) = flexAlignRightCenter:    .Row = 1: .Col = 9: .CellAlignment = flexAlignRightCenter
        '----------------------
        .TextMatrix(1, 10) = "Debe":       .ColWidth(10) = 1150:  .ColAlignment(10) = flexAlignRightCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 11) = "Haber":       .ColWidth(11) = 1150:  .ColAlignment(11) = flexAlignRightCenter:   .Row = 1: .Col = 11: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 12) = "Saldo":       .ColWidth(12) = 1150:  .ColAlignment(12) = flexAlignRightCenter:   .Row = 1: .Col = 12: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 13) = "N°.Documento":       .ColWidth(13) = 1400:  .ColAlignment(13) = flexAlignLeftCenter:   .Row = 1: .Col = 13: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 14) = "Glosa":       .ColWidth(14) = 3500:  .ColAlignment(14) = flexAlignLeftCenter:   .Row = 1: .Col = 14: .CellAlignment = flexAlignLeftCenter
        
        
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
        nPeriodo = "Al  " + CStr(TxtFecha.Valor)
        nTitulo1 = "(Expresado en " & LblMoneda.Caption & ")"
        If Optcli.Value = True Then
            nTipo = "Cliente"
        ElseIf OptProv.Value = True Then
            nTipo = "Proveedor"
        ElseIf opt4ta.Value Then
            nTipo = "Prestador de Servicio"
            
        ElseIf OptLGD.Value = True Then
            nTipo = "Cliente"
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
    
    Dim rst As New ADODB.Recordset
    Dim Rstabo As New ADODB.Recordset
    Dim A, B, xFila As Long
    Dim TotDebe, TotHaber As Double
    Dim TotGralDebe, TotGralHaber As Double
    Dim xNomPro As String
    Dim Cambio As Boolean
    Dim nSQL As String
    
    Dim sSaldoFinal As Double '--indica el saldo final por cada documento
    
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
    
    TabOne1.CurrTab = 1
    
    ProgressBar1.Max = 1
    ProgressBar1.Value = 0
    fraBarra.Visible = True
    fraBarra.Refresh
    DoEvents
    
    
    Dim nSQLWhere As String '--almacenara la condicion de la consulta
    Dim nCampoMuestra As String '--indica el campo que se mostrara esta en funcion de la moneda seleccionada
    Dim nSQLAjuste  As String
    
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " and (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    
    nSQLWhere = ""

    If Optcli.Value = True Then '--ventas
'        If IdCliPro <> 0 Then nSQLWhere = " and vta_ventas.idcli = " & IdCliPro & " "
'        nSQL = "SELECT vta_ventas.id,IIf([vta_ventas]![anulado]=-1,' ',[mae_cliente]![numruc]) AS numruc, IIf([vta_ventas]![anulado]=-1,'Anulado',[mae_cliente]![nombre]) AS nombre, IIf([vta_ventas].[numreg] Is Null Or [vta_ventas].[numreg]='',[mae_libros].[codsun],Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4)) AS registro, 'Ventas' AS libro, mae_documento.codsun,mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc2, vta_ventas.fchdoc, vta_ventas.fchven, mae_moneda.simbolo, " _
'            + vbCr + " iif(vta_ventas.tc is null or vta_ventas.tc=0,con_tc.impven , vta_ventas.tc) AS tipcam, " _
'            + vbCr + " vta_ventas.idmon,vta_ventas.imptotdoc AS imptotal,vta_ventas.impsal, " _
'            + vbCr + " IIf(imptotal=0,0,IIf([vta_ventas].[idmon]=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
'            + vbCr + " IIf(imptotal=0,0,IIf([vta_ventas].[idmon]=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
'            + vbCr + " vta_ventas.glosa as glosaope " _
'            + vbCr + " FROM ((((vta_ventas LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id " _
'            + vbCr + " WHERE ( (vta_ventas.tipdoc <>7 and  vta_ventas.anulado =0) or (vta_ventas.tipdoc=7 AND vta_ventas.anulado=0 AND vta_ventas.iddocref=0) ) and vta_ventas.fchdoc <= CDate('" & TxtFecha.Valor & "') " & nSQLWhere _
'            + vbCr + " ORDER BY IIf([vta_ventas]![anulado]=-1,'Anulado',[mae_cliente]![nombre]), [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc];"
    
    ElseIf OptProv.Value = True Then '--compras
'        '--Sentencia SQL para filtrar el proveedor
'        If IdCliPro <> 0 Then nSQLWhere = " and com_compras.idpro = " & IdCliPro & " "
'
'        nSQL = "SELECT  com_compras.id,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_compras.numreg Is Null Or com_compras.numreg='',mae_libros.codsun ,Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4)) AS registro, 'Compras' AS libro, mae_documento.codsun,mae_documento.abrev, iif(com_compras!numser is null or com_compras!numser ='','',com_compras!numser  +'-' ) + com_compras!numdoc AS numdoc2, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, " _
'            + vbCr + " iif(com_compras.tc is null or com_compras.tc=0,con_tc.impven , com_compras.tc) AS tipcam, " _
'            + vbCr + " com_compras.idmon,com_compras.imptot AS imptotal,com_compras.impsal, " _
'            + vbCr + " IIf(imptotal=0,0,IIf([com_compras].[idmon]=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
'            + vbCr + " IIf(imptotal=0,0,IIf([com_compras].[idmon]=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
'            + vbCr + " com_compras.glosa as glosaope " _
'            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = com_compras.idmon) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
'            + vbCr + " WHERE (com_compras.fchdoc <=CDate('" & TxtFecha.Valor & "') ) " & nSQLWhere & " AND ( com_compras.tipdoc <> 7) " _
'            + vbCr + " ORDER BY mae_prov!nombre, com_compras.fchdoc "
'
'        '--percepciones
'        '--Sentencia SQL para filtrar el proveedor de la percepcion
'        If IdCliPro <> 0 Then nSQLWhere = " and con_percepcion.idcli = " & IdCliPro & " "
'
'        nSQL = nSQL + vbCr + " Union " _
'            + vbCr + "SELECT con_percepcion.id & '' AS id, mae_prov.numruc, mae_prov.nombre, Mid(con_percepcion!numreg,1,2)+mae_libros.codsun+Mid(con_percepcion!numreg,3,4) AS registro, 'Percepciones' AS libro, mae_documento.codsun, mae_documento.abrev, [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc2, con_percepcion.fchdoc ,con_percepcion.fchdoc as fchven, mae_moneda.simbolo, " _
'            + vbCr + " con_tc.impven AS tipcam, con_percepcion.idmon, con_percepcion.imptotper AS imptotal,con_percepcion.impsal, " _
'            + vbCr + " IIf(imptotal=0,0,IIf([con_percepcion].[idmon]=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
'            + vbCr + " IIf(imptotal=0,0,IIf([con_percepcion].[idmon]=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
'            + vbCr + " con_percepcion.glosa AS glosaope " _
'            + vbCr + " FROM (mae_moneda RIGHT JOIN (((con_percepcion LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id) LEFT JOIN mae_libros ON con_percepcion.idlib = mae_libros.id) LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) ON mae_moneda.id = con_percepcion.idmon) LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha " _
'            + vbCr + " WHERE (((con_percepcion.tipo)=1)) and con_percepcion.fchdoc <= CDate('" & TxtFecha.Valor & "')" & nSQLWhere
'
'        '--tabla vistual que permitira dar un orden a la consulta
'        nSQL = "SELECT tab.* FROM ( " & nSQL & " ) AS tab ORDER BY tab.nombre,tab.numdoc2 "

    
    ElseIf opt4ta.Value = True Then '--honorarios
'        '--Sentencia SQL para filtrar el prestador de servicio
'        If IdCliPro <> 0 Then nSQLWhere = " and com_honorarios.idpro = " & IdCliPro & " "
'
'        nSQL = "SELECT  com_honorarios.id,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_honorarios.numreg Is Null Or com_honorarios.numreg='',mae_libros.codsun ,Left(com_honorarios.numreg,2) & mae_libros.codsun & Right(com_honorarios.numreg,4)) AS registro, 'Honorario' AS libro, mae_documento.codsun,mae_documento.abrev, com_honorarios!numser+'-'+com_honorarios!numdoc AS numdoc2, com_honorarios.fchdoc, com_honorarios.fchven, mae_moneda.simbolo, " _
'            + vbCr + " iif(com_honorarios.tc is null or com_honorarios.tc=0,con_tc.impven , com_honorarios.tc) AS tipcam, " _
'            + vbCr + " com_honorarios.idmon,com_honorarios.imptot AS imptotal,com_honorarios.impsal, " _
'            + vbCr + " IIf(imptotal=0,0,IIf([com_honorarios].[idmon]=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
'            + vbCr + " IIf(imptotal=0,0,IIf([com_honorarios].[idmon]=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
'            + vbCr + " com_honorarios.glosa as glosaope " _
'            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_honorarios LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) ON mae_documento.id = com_honorarios.tipdoc) ON mae_prov.id = com_honorarios.idpro) ON mae_moneda.id = com_honorarios.idmon) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha " _
'            + vbCr + " WHERE (com_honorarios.fchdoc <=CDate('" & TxtFecha.Valor & "') )" & nSQLWhere & "AND ( com_honorarios.tipdoc <> 7) " & nSQLWhere _
'            + vbCr + " ORDER BY mae_prov!nombre, com_honorarios.fchdoc;"
'
    
    
    ElseIf OptLetra.Value = True Then
        
        If IdCliPro <> 0 Then nSQLWhere = " and let_letra.idclipro = " & IdCliPro & " "
        
        
        nSQL = "SELECT let_letradet.corr AS id, mae_cliente.numruc, mae_cliente.nombre, Left([let_letra].[numreg],2) & [mae_libros].[codsun] & Right([let_letra].[numreg],4) AS registro, 'Letras' AS libro, mae_documento.abrev, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS numdoc2, " _
        + vbCr + " let_letradet.fchemi AS fchdoc, let_letradet.fchven, mae_moneda.simbolo, IIf([let_letra].[tc]=0,[con_tc].[impven],[let_letra].[tc]) AS tipcam, let_letra.idmon, let_letradet.implet AS imptotal, let_letradet.impsal, IIf(let_letra.idmon=1,let_letradet.implet,let_letradet.implet*tipcam) AS imptotsol, IIf(let_letra.idmon=2,let_letradet.implet,IIf(tipcam=0,0,let_letradet.implet/tipcam)) AS imptotdol, let_letra.glosa AS glosaope " _
        + vbCr + " FROM mae_moneda RIGHT JOIN (((((mae_cliente RIGHT JOIN let_letra ON mae_cliente.id = let_letra.idclipro) LEFT JOIN mae_documento ON let_letra.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON let_letra.idlib = mae_libros.id) LEFT JOIN con_tc ON let_letra.fchemi = con_tc.fecha) INNER JOIN let_letradet ON let_letra.id = let_letradet.idlet) ON mae_moneda.id = let_letra.idmon " _
        + vbCr + " WHERE (((let_letradet.fchemi)<=CDate('" & TxtFecha.Valor & "'))) " & nSQLWhere _
        + vbCr + " ORDER BY mae_cliente.nombre, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser];"
    
    ElseIf OptPlaLetra.Value = True Then
        
        If IdCliPro <> 0 Then nSQLWhere = " and mae_banconumcta.idban = " & IdCliPro & " "
        
        nSQL = "SELECT let_planilla.id, mae_bancos.numruc, mae_bancos.descripcion AS nombre, Left([let_planilla].[numreg],2) & [mae_libros].[codsun] & Right([let_planilla].[numreg],4) AS registro, 'Planilla letra' AS libro, mae_documento.abrev, let_planilla.numdoc AS numdoc2, " _
        + vbCr + " let_planilla.fchemi AS fchdoc, '' AS fchven, mae_moneda.simbolo, IIf([let_planilla].[anulado]=-1,0,IIf([let_planilla].[tc]=0,[con_tc].[impven],[let_planilla].[tc])) AS tipcam, let_planilla.idmon, let_planilla.imptot AS imptotal,  let_planilla.impsal, IIf(let_planilla.idmon=1,let_planilla.imptot,let_planilla.imptot*tipcam) AS imptotsol, IIf(let_planilla.idmon=2,let_planilla.imptot,IIf(tipcam=0,0,let_planilla.imptot/tipcam)) AS imptotdol , let_planilla.glosa AS glosaope " _
        + vbCr + " FROM mae_documento RIGHT JOIN (mae_bancos RIGHT JOIN ((((let_planilla LEFT JOIN mae_moneda ON let_planilla.idmon = mae_moneda.id) LEFT JOIN mae_banconumcta ON let_planilla.idbcocta = mae_banconumcta.id) LEFT JOIN mae_libros ON let_planilla.idlib = mae_libros.id) LEFT JOIN con_tc ON let_planilla.fchemi = con_tc.fecha) ON mae_bancos.id = mae_banconumcta.idban) ON mae_documento.id = let_planilla.tipdoc " _
        + vbCr + " WHERE (((let_planilla.fchemi)<=CDate('" & TxtFecha.Valor & "'))) " & nSQLWhere _
        + vbCr + " ORDER BY mae_bancos.descripcion, let_planilla.numdoc;"

            
    
    ElseIf OptLGD.Value = True Then
    
        If IdCliPro <> 0 Then nSQLWhere = " and vta_gastodebito.idcli = " & IdCliPro & " "
        nSQL = "SELECT vta_gastodebito.id,IIf([vta_gastodebito]![anulado]=-1,' ',[mae_cliente]![numruc]) AS numruc, IIf([vta_gastodebito]![anulado]=-1,'Anulado',[mae_cliente]![nombre]) AS nombre, IIf([vta_gastodebito].[numreg] Is Null Or [vta_gastodebito].[numreg]='',[mae_libros].[codsun],Left([vta_gastodebito].[numreg],2) & [mae_libros].[codsun] & Right([vta_gastodebito].[numreg],4)) AS registro, 'Lgd' AS libro, mae_documento.codsun,mae_documento.abrev, [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc] AS numdoc2, vta_gastodebito.fchemi as fchdoc, vta_gastodebito.fchven, mae_moneda.simbolo, " _
            + vbCr + " iif(vta_gastodebito.tc is null or vta_gastodebito.tc=0,con_tc.impven , vta_gastodebito.tc) AS tipcam, " _
            + vbCr + " vta_gastodebito.idmon,vta_gastodebito.imptot AS imptotal,vta_gastodebito.impsal, " _
            + vbCr + " IIf(imptotal=0,0,IIf([vta_gastodebito].[idmon]=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
            + vbCr + " IIf(imptotal=0,0,IIf([vta_gastodebito].[idmon]=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
            + vbCr + " vta_gastodebito.glosa as glosaope " _
            + vbCr + " FROM ((((vta_gastodebito LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id " _
            + vbCr + " WHERE ( (vta_gastodebito.tipdoc <>126 and  vta_gastodebito.anulado =0) ) and vta_gastodebito.fchemi <= CDate('" & TxtFecha.Valor & "') " & nSQLWhere _
            + vbCr + " ORDER BY IIf([vta_gastodebito]![anulado]=-1,'Anulado',[mae_cliente]![nombre]), [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc];"
    
    
    ElseIf OptReem.Value = True Then
     
        '--Sentencia SQL para filtrar el proveedor
        If IdCliPro <> 0 Then nSQLWhere = " and com_reembolsables.idpro = " & IdCliPro & " "
        
        nSQL = "SELECT  com_reembolsables.id,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_reembolsables.numreg Is Null Or com_reembolsables.numreg='',mae_libros.codsun ,Left(com_reembolsables.numreg,2) & mae_libros.codsun & Right(com_reembolsables.numreg,4)) AS registro, 'Reembolsables' AS libro, mae_documento.codsun,mae_documento.abrev, iif(com_reembolsables!numser is null or com_reembolsables!numser ='','',com_reembolsables!numser  +'-' ) + com_reembolsables!numdoc AS numdoc2, com_reembolsables.fchdoc, com_reembolsables.fchven, mae_moneda.simbolo, " _
            + vbCr + " iif(com_reembolsables.tc is null or com_reembolsables.tc=0,con_tc.impven , com_reembolsables.tc) AS tipcam, " _
            + vbCr + " com_reembolsables.idmon,com_reembolsables.imptot AS imptotal,com_reembolsables.impsal, " _
            + vbCr + " IIf(imptotal=0,0,IIf([com_reembolsables].[idmon]=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
            + vbCr + " IIf(imptotal=0,0,IIf([com_reembolsables].[idmon]=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
            + vbCr + " com_reembolsables.glosa as glosaope " _
            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_reembolsables LEFT JOIN mae_libros ON com_reembolsables.idlib = mae_libros.id) ON mae_documento.id = com_reembolsables.tipdoc) ON mae_prov.id = com_reembolsables.idpro) ON mae_moneda.id = com_reembolsables.idmon) LEFT JOIN con_tc ON com_reembolsables.fchdoc = con_tc.fecha " _
            + vbCr + " WHERE (com_reembolsables.fchdoc <=CDate('" & TxtFecha.Valor & "') ) " & nSQLWhere & " AND ( com_reembolsables.tipdoc <> 7) " _
            + vbCr + " ORDER BY mae_prov!nombre, com_reembolsables.fchdoc "
        
        '--tabla vistual que permitira dar un orden a la consulta
        nSQL = "SELECT tab.* FROM ( " & nSQL & " ) AS tab ORDER BY tab.nombre,tab.numdoc2 "
     
    Else
        MsgBox "Pendiente de desarrollo", vbInformation, xTitulo
        
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
    
    '--------------------------------------------------------------------------------------------
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
    '--------------------------------------------------------------------------------------------
    
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
        
        '------------------------------------------------------------------------
        '--colocar datos del grupo(cliente, proveedor o prestador de servicio
        '--detalle
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(xFila, 1) = "Nº R.U.C. :"
        Fg1.TextMatrix(xFila, 2) = NulosC(rst("numruc"))
        Fg1.TextMatrix(xFila, 4) = NulosC(rst("nombre"))
        '--resumen
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(rst("numruc"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(rst("nombre"))
        '------------------------------------------------------------------------
        '******
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
        '         37=Letras
        '         39=Rendición de Cuentas
        '         40=Registro de Honorarios Profesionales
        '         41=Liquidacion de Gasto Debito
        
        '--cargando los abonos
        If Optcli.Value = True Then
'            If IdCliPro <> 0 Then nSQLWhere = " and con_diario.ridper = " & IdCliPro & " "
'
'            nSQL = "SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, '' AS codsun, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, IIf([con_diario].[aplicatc]=0,[con_tc].[impven],[con_diario].[tc]) AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, IIf(con_diario.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, IIf(con_diario.idmon=2,imptotal, iif(tipcam=0,0,imptotal/tipcam)) AS imptotdol, con_diario.ridper,con_diario.rnumerodoc AS numdoc2,con_diario.rglosaope, con_diario.iddoc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc " _
'                + vbCr + " FROM ((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) INNER JOIN con_planctas ON con_diario.idcue = con_planctas.id " _
'                + vbCr + " WHERE (((con_diario.idlib) In (5,6,8,37,44)) AND ((con_diario.ridlib)=2)) " & nSQLAjuste & nSQLWhere
'
'            '--unido a referencias de nota de credito
'            If IdCliPro <> 0 Then nSQLWhere = " and vta_ventas.idcli = " & IdCliPro & " "
'
'            nSQL = nSQL + vbCr + " Union All " _
'                + vbCr + "SELECT Left([vta_ventas_1].[numreg],2) & [mae_libros_1].[codsun] & Right([vta_ventas_1].[numreg],4) AS rregistro, Mid([vta_ventas]![numreg],1,2) & [mae_libros]![codsun] & Mid([vta_ventas]![numreg],3,4) AS registro, mae_libros.descripcion AS libro, mae_libros.codsun, mae_documento.abrev, [vta_ventas]![numser] & '-' & [vta_ventas]![numdoc] AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, IIf([vta_ventas_1].[tc]=0,[con_tc].[impven],[vta_ventas_1].[tc]) AS tipcam, vta_ventas.imptotdoc AS imptotdocal, IIf([vta_ventas]![idmon]=1,[vta_ventas]![imptotdoc],[vta_ventas]![imptotdoc]*tipcam) AS imptotdocsol, IIf([vta_ventas]![idmon]=2,[vta_ventas]![imptotdoc],iif(tipcam=0,0,[vta_ventas]![imptotdoc]/tipcam) ) AS imptotdocdol, vta_ventas.idcli AS ridper,  [vta_ventas_1].[numser] & '-' & [vta_ventas_1].[numdoc] AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc " _
'                 + vbCr + " FROM (((mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) INNER JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id) LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_ventas.idmon = mae_documentocta.idmon) AND (vta_ventas.tipdoc = mae_documentocta.iddoc) " _
'                 + vbCr + " WHERE vta_ventas.anulado=0 and vta_ventas.tipdoc=7 and vta_ventas.iddocref <> 0 and mae_documentocta.tipope =-1 " & nSQLWhere
'
'
'            RST_Busq Rstabo, nSQL, xCon
            
        ElseIf OptProv.Value = True Then
'            '--buscar de bancos, canjes de documentos , rendicion de cuenta
'
'            If IdCliPro <> 0 Then nSQLWhere = " and con_diario.ridper = " & IdCliPro & " "
'            nSQL = "SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, '' AS codsun, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, IIf([con_diario].[aplicatc]=0,[con_tc].[impven],[con_diario].[tc]) AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, IIf(con_diario.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, IIf(con_diario.idmon=2 ,imptotal,iif(tipcam=0,0, imptotal/tipcam)) AS imptotdol, con_diario.ridper,con_diario.rnumerodoc AS numdoc2,con_diario.rglosaope,con_diario.iddoc , con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc " _
'                + vbCr + " FROM (((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN tes_cajadestinodet ON (con_diario.idmov = tes_cajadestinodet.idtes) AND (con_diario.iddocpro = tes_cajadestinodet.corr)) LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id " _
'                + vbCr + " WHERE (((con_diario.idlib) In (6,8,39,44)) AND ((con_diario.ridlib) in (1,4))) " & nSQLAjuste & nSQLWhere
'
'            '--unido a referencias de nota de credito
'            If IdCliPro <> 0 Then nSQLWhere = " and com_compras.idpro = " & IdCliPro & " "
'
'           nSQL = nSQL + vbCr + " Union All " _
'                + vbCr + "SELECT Left([com_compras_1].[numreg],2) & [mae_libros_1].[codsun] & Right([com_compras_1].[numreg],4) AS rregistro, Mid([com_compras]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_compras]![numreg],3,4) AS registro, mae_libros.descripcion AS libro, mae_libros.codsun, mae_documento.abrev, [com_compras]![numser] & '-' & [com_compras]![numdoc] AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, IIf([com_compras_1].[tc]=0,[con_tc].[impven],[com_compras_1].[tc]) AS tipcam, com_compras.imptot AS imptotal, IIf([com_compras]![idmon]=1,[com_compras]![imptot],[com_compras]![imptot]*tipcam) AS imptotsol, IIf([com_compras]![idmon]=2,[com_compras]![imptot],iif(tipcam=0,0,[com_compras]![imptot]/tipcam)) AS imptotdol, com_compras.idpro AS ridper,  [com_compras_1].[numser] & '-' & [com_compras_1].[numdoc] AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc  " _
'                + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN com_compras AS com_compras_1 ON com_compras.iddocref = com_compras_1.id) LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (com_compras.idmon = mae_documentocta.idmon) AND (com_compras.tipdoc = mae_documentocta.iddoc) " _
'                + vbCr + " WHERE com_compras.iddocref Is Not Null And com_compras.iddocref<>0 and mae_documentocta.tipope=0 " & nSQLWhere
'
'            RST_Busq Rstabo, nSQL, xCon
            
        ElseIf opt4ta.Value = True Then
'            If IdCliPro <> 0 Then nSQLWhere = " and con_diario.ridper = " & IdCliPro & " "
'
'            nSQL = "SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, '' AS codsun, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, IIf(con_diario.idmon=1,imptotal,imptotal*[con_tc]![impven]) AS imptotsol, IIf(con_diario.idmon=2 And [con_tc]![impven] Is Not Null,imptotal,imptotal/[con_tc]![impven]) AS imptotdol, con_diario.ridper,con_diario.rnumerodoc AS numdoc2,con_diario.rglosaope,IIf([con_diario].[idlib]=6,[tes_cajadestinodet].[iddoc],[con_diario].[iddocpro]) AS iddoc " _
'                + vbCr + " FROM ((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN tes_cajadestinodet ON (con_diario.iddocpro = tes_cajadestinodet.corr) AND (con_diario.idmov = tes_cajadestinodet.idtes) " _
'                + vbCr + " WHERE (((con_diario.idlib) In (6,8,39,44)) AND ((con_diario.ridlib)=40)) " & nSQLAjuste & nSQLWhere
'
'            RST_Busq Rstabo, nSQL, xCon
            
        '--letras
        ElseIf OptLetra.Value = True Then
            If IdCliPro <> 0 Then nSQLWhere = " and con_diario.ridper = " & IdCliPro & " "
            
            nSQL = "SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, '' AS codsun, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, IIf([con_diario].[aplicatc]=0,[con_tc].[impven],[con_diario].[tc]) AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, IIf(con_diario.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, IIf(con_diario.idmon=2,imptotal,iif(tipcam=0,0,imptotal/tipcam)) AS imptotdol, con_diario.ridper,con_diario.rnumerodoc AS numdoc2,con_diario.rglosaope,con_diario.iddoc , con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc " _
                + vbCr + " FROM ((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id " _
                + vbCr + " WHERE (((con_diario.idlib) In (6,8,39,42)) AND ((con_diario.ridlib)=37)) " & nSQLAjuste & nSQLWhere
            
            RST_Busq Rstabo, nSQL, xCon
            
        '--planilla de letras
        ElseIf OptPlaLetra.Value = True Then
            If IdCliPro <> 0 Then nSQLWhere = " and con_diario.ridper = " & IdCliPro & " "
            
            nSQL = "SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, '' AS codsun, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, IIf(con_diario.idmon=1,imptotal,imptotal*[con_tc]![impven]) AS imptotsol, IIf(con_diario.idmon=2 And [con_tc]![impven] Is Not Null,imptotal,imptotal/[con_tc]![impven]) AS imptotdol, con_diario.ridper,con_diario.rnumerodoc AS numdoc2,con_diario.rglosaope,con_diario.iddoc " _
                + vbCr + " FROM ((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN tes_cajadestinodet ON (con_diario.iddocpro = tes_cajadestinodet.corr) AND (con_diario.idmov = tes_cajadestinodet.idtes) " _
                + vbCr + " WHERE (((con_diario.idlib) In (6,44)) AND ((con_diario.ridlib)=42)) " & nSQLAjuste & nSQLWhere
            
            RST_Busq Rstabo, nSQL, xCon
            
            
        '--liquidacion gasto debito
        ElseIf OptLGD.Value = True Then
            If IdCliPro <> 0 Then nSQLWhere = " and con_diario.ridper = " & IdCliPro & " "
            
            nSQL = "SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, '' AS codsun, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, IIf(con_diario.idmon=1,imptotal,imptotal*[con_tc]![impven]) AS imptotsol, IIf(con_diario.idmon=2 And [con_tc]![impven] Is Not Null,imptotal,imptotal/[con_tc]![impven]) AS imptotdol, con_diario.ridper,con_diario.rnumerodoc AS numdoc2,con_diario.rglosaope,con_diario.iddoc  " _
                + vbCr + " FROM ((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN tes_cajadestinodet ON (con_diario.iddocpro = tes_cajadestinodet.corr) AND (con_diario.idmov = tes_cajadestinodet.idtes) " _
                + vbCr + " WHERE (((con_diario.idlib) In (6,44)) AND ((con_diario.ridlib)=41)) " & nSQLAjuste & nSQLWhere
            
            '--unido a referencias de liquidacion gasto credito
            If IdCliPro <> 0 Then nSQLWhere = " and vta_gastodebito.idcli = " & IdCliPro & " "
            
            nSQL = nSQL + vbCr + " Union All " _
                + vbCr + "SELECT Left([vta_gastodebito_1].[numreg],2) & [mae_libros_1].[codsun] & Right([vta_gastodebito_1].[numreg],4) AS rregistro, Mid([vta_gastodebito]![numreg],1,2) & [mae_libros]![codsun] & Mid([vta_gastodebito]![numreg],3,4) AS registro, mae_libros.descripcion AS libro, mae_libros.codsun, mae_documento.abrev, [vta_gastodebito]![numser] & '-' & [vta_gastodebito]![numdoc] AS numdoc, vta_gastodebito.fchemi, mae_moneda.simbolo, IIf([vta_gastodebito].[tc]=0,[con_tc].[impven],[vta_gastodebito].[tc]) AS tipcam, vta_gastodebitodet.impacue AS imptotal, IIf([vta_gastodebito]![idmon]=1,vta_gastodebitodet.impacue,vta_gastodebitodet.impacue*tipcam) AS imptotsol, IIf([vta_gastodebito]![idmon]=2,vta_gastodebitodet.impacue,IIf(tipcam=0,0,vta_gastodebitodet.impacue/tipcam)) AS imptotdol, vta_gastodebito.idcli AS ridper, [vta_gastodebito_1].[numser] & '-' & [vta_gastodebito_1].[numdoc] AS numdoc2, vta_gastodebito.glosa AS glosaope, vta_gastodebito_1.id AS iddoc " _
                + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((vta_gastodebito LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) ON mae_documento.id = vta_gastodebito.tipdoc) ON mae_moneda.id = vta_gastodebito.idmon) INNER JOIN (vta_gastodebitodet INNER JOIN (vta_gastodebito AS vta_gastodebito_1 LEFT JOIN mae_libros AS mae_libros_1 ON vta_gastodebito_1.idlib = mae_libros_1.id) ON vta_gastodebitodet.iddoc = vta_gastodebito_1.id) ON vta_gastodebito.id = vta_gastodebitodet.idlgd " _
                + vbCr + " WHERE (((vta_gastodebito.tipdoc)=126) AND ((vta_gastodebitodet.idmod)=11)) "

            
            
            RST_Busq Rstabo, nSQL, xCon
        
        '--compras reembolsables
        ElseIf OptReem.Value = True Then
        
        
            '--buscar de bancos, canjes de documentos , rendicion de cuenta
            If IdCliPro <> 0 Then nSQLWhere = " and con_diario.ridper = " & IdCliPro & " "
            
            If OptReem1.Value = True Then '--vista por bancos
                nSQL = "SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, '' AS codsun, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, IIf(con_diario.idmon=1,imptotal,imptotal*[con_tc]![impven]) AS imptotsol, IIf(con_diario.idmon=2 And [con_tc]![impven] Is Not Null,imptotal,imptotal/[con_tc]![impven]) AS imptotdol, con_diario.ridper,con_diario.rnumerodoc AS numdoc2,con_diario.rglosaope,con_diario.iddoc " _
                    + vbCr + " FROM ((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN tes_cajadestinodet ON (con_diario.iddocpro = tes_cajadestinodet.corr) AND (con_diario.idmov = tes_cajadestinodet.idtes) " _
                    + vbCr + " WHERE (((con_diario.idlib) In (6)) AND ((con_diario.ridlib) in (999))) " & nSQLAjuste & nSQLWhere
            
            Else '--vista por lgd
                
                nSQL = "SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, '' AS codsun, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, con_tc.impven AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, IIf(con_diario.idmon=1,imptotal,imptotal*[con_tc]![impven]) AS imptotsol, IIf(con_diario.idmon=2 And [con_tc]![impven] Is Not Null,imptotal,imptotal/[con_tc]![impven]) AS imptotdol, con_diario.ridper, con_diario.rnumerodoc AS numdoc2, con_diario.rglosaope, con_diario.iddoc " _
                    + vbCr + " FROM (((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id " _
                    + vbCr + " WHERE (((con_diario.idlib) In (41)) AND ((con_diario.ridlib) In (999))) " & nSQLAjuste & nSQLWhere

            End If
            '--unido a referencias de nota de credito
'            If IdCliPro <> 0 Then nSQLWhere = " and com_compras.idpro = " & IdCliPro & " "
'
'           nSQL = nSQL + vbCr + " Union All " _
'                + vbCr + "SELECT Left([com_compras_1].[numreg],2) & [mae_libros_1].[codsun] & Right([com_compras_1].[numreg],4) AS rregistro, Mid([com_compras]![numreg],1,2) & [mae_libros]![codsun] & Mid([com_compras]![numreg],3,4) AS registro, mae_libros.descripcion AS libro, mae_libros.codsun, mae_documento.abrev, [com_compras]![numser] & '-' & [com_compras]![numdoc] AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, con_tc.impven AS timpcam, com_compras.imptot AS imptotal, IIf([com_compras]![idmon]=1,[com_compras]![imptot],[com_compras]![imptot]*[con_tc]![impven]) AS imptotsol, IIf([com_compras]![idmon]=2,[com_compras]![imptot],[com_compras]![imptot]/[con_tc]![impven]) AS imptotdol, com_compras.idpro AS ridper,  [com_compras_1].[numser] & '-' & [com_compras_1].[numdoc] AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc " _
'                + vbCr + " FROM ((mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) INNER JOIN com_compras AS com_compras_1 ON com_compras.iddocref = com_compras_1.id) LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id " _
'                + vbCr + " WHERE (((com_compras.iddocref) Is Not Null And (com_compras.iddocref)<>0))  " & nSQLWhere
            
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
                Fg1.TextMatrix(xFila, 10) = Format(TotDebe, FORMAT_MONTO)
                Fg1.TextMatrix(xFila, 11) = Format(TotHaber, FORMAT_MONTO)
                
                '--detalle
                If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value = True Or OptPlaLetra.Value = True Then
                    Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
                
                ElseIf OptProv.Value = True Or opt4ta.Value = True Or OptReem.Value = True Then
                    Fg1.TextMatrix(xFila, 12) = Format(TotHaber - TotDebe, FORMAT_MONTO)
                
                End If
                
                '*****resumen
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotDebe, FORMAT_MONTO)
                Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotHaber, FORMAT_MONTO)
                
                If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value = True Or OptPlaLetra.Value = True Then
                    Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
                    
                ElseIf OptProv.Value = True Or opt4ta.Value = True Or OptReem.Value = True Then
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
                Fg1.TextMatrix(xFila, 1) = "Nº R.U.C. :"
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
            Fg1.TextMatrix(xFila, 3) = NulosC(rst("abrev"))
            Fg1.TextMatrix(xFila, 4) = NulosC(rst("numdoc2"))
            Fg1.TextMatrix(xFila, 5) = Format(rst("fchdoc"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 6) = Format(rst("fchven"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 7) = NulosC(rst("simbolo"))
            Fg1.TextMatrix(xFila, 8) = Format(NulosN(rst("imptotal")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 9) = Format(NulosN(rst("tipcam")), "###0.##0") & ""
            
            If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value = True Or OptPlaLetra.Value = True Then
                Fg1.TextMatrix(xFila, 10) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO)
                TotDebe = TotDebe + NulosN(rst(nCampoMuestra))
                
            ElseIf OptProv.Value = True Or opt4ta.Value = True Or OptReem.Value = True Then
                Fg1.TextMatrix(xFila, 11) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO) '--saldo
                TotHaber = TotHaber + NulosN(rst(nCampoMuestra))
                
            End If
            
            Fg1.TextMatrix(xFila, 12) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO)
            
            Fg1.TextMatrix(xFila, 13) = NulosC(rst("numdoc2"))
            Fg1.TextMatrix(xFila, 14) = NulosC(rst("glosaope"))
            
            xSaldoDoc = NulosN(rst("impsal"))
            
            
            '-------------------------------------------------------------
            '--filtrar los movimientos de las provisiones para proceder a obtener el saldo actual
            '--primero hacer que no haya nada que filtrar
            
            Rstabo.Filter = "rregistro = '-----------'"
            '--
            '--ventas
            If Optcli.Value = True Then
                Rstabo.Filter = "rregistro = '" & rst("registro") & "' and iddoc= " & rst("id")
                
            '--compras
            ElseIf OptProv.Value = True Then '--proveedor
                Rstabo.Filter = "rregistro = '" & rst("registro") & "' and iddoc= " & rst("id")
                
            '--honorarios
            ElseIf opt4ta.Value = True Then '--honorarios
                Rstabo.Filter = "rregistro = '" & rst("registro") & "' and iddoc= " & rst("id")
            
            '--letras
            ElseIf OptLetra.Value = True Then
                Rstabo.Filter = "rregistro = '" & rst("registro") & "' and iddoc= " & rst("id")
            
            '--planilla de letras
            ElseIf OptPlaLetra.Value = True Then
                Rstabo.Filter = "rregistro = '" & rst("registro") & "' and iddoc= " & rst("id")
                
            '--liquidacion gasto debito
            ElseIf OptLGD.Value = True Then
                Rstabo.Filter = "rregistro = '" & rst("registro") & "' and iddoc= " & rst("id")
                
            '--compras reembolsable
            ElseIf OptReem.Value = True Then
                Rstabo.Filter = "iddoc= " & rst("id")
                
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
                     
                     
                    '**************************************************************************************************************************
                    '--verificar si el libro es de ajuste por dif de cambio
                    If InStr(LCase(Rstabo("libro")), "ajuste") <> 0 Then
                        '--verificar si es perdida o ganancia
                        If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value = True Or OptPlaLetra.Value = True Then
                        
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
                        
                        ElseIf OptProv.Value = True Or opt4ta.Value = True Or OptReem.Value = True Then
                        
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
                         If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value = True Or OptPlaLetra.Value = True Then
                             Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                             TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
                             
                         ElseIf OptProv.Value = True Or opt4ta.Value = True Or OptReem.Value = True Then
                             Fg1.TextMatrix(xFila, 10) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                             TotDebe = TotDebe + NulosN(Rstabo(nCampoMuestra))
                             
                        End If
                        '--obteniendo utlimo saldo
                        Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                        
                    End If

                    '**************************************************************************************************************************
                    Fg1.TextMatrix(xFila, 13) = NulosC(Rstabo("numdoc2"))
                    Fg1.TextMatrix(xFila, 14) = NulosC(Rstabo("rglosaope"))
                                        
                    Rstabo.MoveNext
                Loop
                
            End If
            
            '---ACTUALIZANDO EL SALDO AL DOCUMENTO
            If (xSaldoDoc <> NulosN(Fg1.TextMatrix(xFila, 12)) And NulosN(rst("idmon")) = NulosN(TxtIdMon.Text)) Or InStr(LCase(Fg1.TextMatrix(xFila, 2)), "ajuste") <> 0 Then
            
                '--obtener el ultimo saldo del documento
                If InStr(LCase(Fg1.TextMatrix(xFila, 2)), "ajuste") <> 0 Then
                    sSaldoFinal = 0
                Else
                    sSaldoFinal = NulosN(Fg1.TextMatrix(xFila, 12))
                End If
                '--------------------------------------------------------
                
            
                If Optcli.Value = True Then     '--VENTAS
                    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & sSaldoFinal & " WHERE (((vta_ventas.id)=" & rst("id") & "))"
                
                ElseIf OptProv.Value = True Then                                '--COMPRAS
                    If LCase(NulosC(rst("libro"))) = "percepciones" Then
                        xCon.Execute "UPDATE con_percepcion SET con_percepcion.impsal = " & sSaldoFinal & " WHERE (((con_percepcion.id)=" & rst("id") & "))"
                    Else
                        xCon.Execute "UPDATE com_compras SET com_compras.impsal = " & sSaldoFinal & " WHERE (((com_compras.id)=" & rst("id") & "))"
                    End If
                    
                ElseIf opt4ta.Value Then '--honorarios
                    xCon.Execute "UPDATE com_honorarios SET com_honorarios.impsal = " & sSaldoFinal & " WHERE (((com_honorarios.id)=" & rst("id") & "))"
                    
                ElseIf OptLetra.Value = True Then '--letras
                    xCon.Execute "UPDATE let_letradet SET let_letradet.impsal = " & sSaldoFinal & " WHERE (((let_letradet.corr)=" & rst("id") & "))"
                
                ElseIf OptPlaLetra.Value = True Then '--planilla de letras
                    xCon.Execute "UPDATE let_planilla SET let_planilla.impsal = " & sSaldoFinal & " WHERE (((let_planilla.id)=" & rst("id") & "))"
                
                ElseIf OptLGD.Value = True Then     '--liquidacion gasto debito
                    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & sSaldoFinal & " WHERE (((vta_ventas.id)=" & rst("id") & "))"
                                
                ElseIf OptReem.Value = True Then '--Compras reembolsables
                    If OptReem1.Value = True Then '--vista por bancos
                        xCon.Execute "UPDATE com_reembolsables SET com_reembolsables.impsal = " & sSaldoFinal & " WHERE (((com_reembolsables.id)=" & rst("id") & "))"
                    End If
                    
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
                            If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value = True Or OptPlaLetra.Value = True Then
                                TotHaber = TotHaber - NulosN(Rstabo(nCampoMuestra))
                                
                            ElseIf OptProv.Value = True Or opt4ta.Value = True Or OptReem.Value = True Then
                                TotDebe = TotDebe - NulosN(Rstabo(nCampoMuestra))
                                
                            End If
                            Rstabo.MoveNext
                        Loop
                    End If
                    
                    If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value = True Or OptPlaLetra.Value = True Then
                        TotDebe = TotDebe - NulosN(rst(nCampoMuestra))
                        
                    ElseIf OptProv.Value = True Or opt4ta.Value = True Or OptReem.Value = True Then
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
        
        
        '---------------------------------------------------------
        TotGralHaber = TotGralHaber + TotHaber
        TotGralDebe = TotGralDebe + TotDebe
        '---------------------------------------------------------
        
        
        Fg1.Rows = Fg1.Rows + 1
        xFila = xFila + 1
        Fg1.TextMatrix(xFila, 4) = "TOTAL -->"

        Fg1.TextMatrix(xFila, 10) = Format(TotDebe, FORMAT_MONTO)
        Fg1.TextMatrix(xFila, 11) = Format(TotHaber, FORMAT_MONTO)
                
        If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value = True Or OptPlaLetra.Value = True Then
            Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
        
        ElseIf OptProv.Value = True Or opt4ta.Value = True Or OptReem.Value = True Then
            Fg1.TextMatrix(xFila, 12) = Format(TotHaber - TotDebe, FORMAT_MONTO)
            
        End If
        
        '*****resumen
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotDebe, FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotHaber, FORMAT_MONTO)
        
        If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value = True Or OptPlaLetra.Value = True Then
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
        
        ElseIf OptProv.Value = True Or opt4ta.Value = True Or OptReem.Value = True Then
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
            
            If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value = True Or OptPlaLetra.Value = True Then
                Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
                
            ElseIf OptProv.Value = True Or opt4ta.Value = True Or OptReem.Value = True Then
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
            
            If Optcli.Value = True Or OptLGD.Value = True Or OptLetra.Value = True Or OptPlaLetra.Value = True Then
                Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
                
            ElseIf OptProv.Value = True Or opt4ta.Value = True Or OptReem.Value = True Then
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
    '--------------------------------------
    Set rst = Nothing
    Set Rstabo = Nothing
    fraBarra.Visible = False
    Me.MousePointer = vbDefault
    MsgBox "La Consulta fue se realizó Correctamente", vbInformation, xTitulo
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
        If NulosC(TxtCliPro.Text) = "" Then
            If Optcli.Value = True Then
                MsgBox "No ha especificado el cliente a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
            If OptProv.Value = True Then
                MsgBox "No ha especificado el proveedor a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
            TxtCliPro.SetFocus
            Exit Sub
        End If
        CargarCli2 NulosN(LblIdCliPro.Caption)
    End If

End Sub


