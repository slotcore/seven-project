VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLetras 
   Caption         =   "Caja y Bancos - Emision de Letras"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   555
      Left            =   1800
      TabIndex        =   89
      Top             =   5040
      Visible         =   0   'False
      Width           =   7170
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   270
         TabIndex        =   90
         Top             =   210
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Left            =   5265
      TabIndex        =   88
      Top             =   360
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   6765
      TabIndex        =   84
      Top             =   360
      Width           =   5040
      Begin VB.Label LblNumRegistros 
         Alignment       =   2  'Center
         Caption         =   "LblNumRegistros"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   4110
         TabIndex        =   87
         Top             =   75
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Registros : "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   2130
         TabIndex        =   86
         Top             =   75
         Width           =   1920
      End
      Begin VB.Label LblMes 
         Alignment       =   2  'Center
         Caption         =   "LblMes"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   0
         TabIndex        =   85
         Top             =   75
         Width           =   1740
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7170
      Left            =   -15
      TabIndex        =   24
      Top             =   345
      Width           =   11820
      _cx             =   20849
      _cy             =   12647
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
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   6750
         Left            =   12465
         TabIndex        =   95
         Top             =   375
         Width           =   11730
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6750
         Left            =   45
         TabIndex        =   29
         Top             =   375
         Width           =   11730
         Begin VB.Frame Frame10 
            Caption         =   "[ Periodo ]"
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
            Left            =   9270
            TabIndex        =   97
            Top             =   690
            Width           =   2370
            Begin VB.Label LblMes1 
               Alignment       =   2  'Center
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
               Height          =   240
               Left            =   270
               TabIndex        =   98
               Top             =   270
               Width           =   1860
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "[ Datos del Girador ]"
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
            Height          =   1020
            Left            =   30
            TabIndex        =   70
            Top             =   690
            Width           =   9225
            Begin VB.CommandButton CmdBusIdDoc 
               Height          =   240
               Left            =   2040
               Picture         =   "FrmLetras.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   75
               Top             =   660
               Width           =   240
            End
            Begin VB.TextBox TxtNumDoc 
               Height          =   300
               Left            =   7590
               MaxLength       =   8
               TabIndex        =   3
               Text            =   "TxtNumDoc"
               Top             =   630
               Width           =   1590
            End
            Begin VB.TextBox TxtGirado 
               Height          =   300
               Left            =   1410
               TabIndex        =   1
               Text            =   "TxtGirado"
               Top             =   300
               Width           =   4995
            End
            Begin VB.TextBox TxtIdDocIden 
               Height          =   300
               Left            =   1410
               MaxLength       =   1
               TabIndex        =   2
               Text            =   "TxtIdDocIden"
               Top             =   630
               Width           =   900
            End
            Begin VB.Label LblDocIdentidad 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDocIdentidad"
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
               Left            =   2310
               TabIndex        =   74
               Top             =   630
               Width           =   4080
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Doc. Identidad"
               Height          =   195
               Index           =   8
               Left            =   135
               TabIndex        =   73
               Top             =   660
               Width           =   1050
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº Documento"
               Height          =   195
               Index           =   7
               Left            =   6480
               TabIndex        =   72
               Top             =   660
               Width           =   1050
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Girado A"
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   71
               Top             =   330
               Width           =   615
            End
         End
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   2730
            Picture         =   "FrmLetras.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   390
            Width           =   240
         End
         Begin VB.Frame fra_letra 
            Caption         =   "[ Datos de la Letra ]"
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
            Height          =   1710
            Left            =   30
            TabIndex        =   30
            Top             =   1725
            Width           =   11655
            Begin VB.TextBox TxtNumDocRef2 
               Height          =   300
               Left            =   6405
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   99
               Text            =   "TxtNumDocRef2"
               Top             =   1320
               Width           =   3300
            End
            Begin VB.TextBox TxtNotDeb 
               Height          =   300
               Left            =   9690
               MaxLength       =   15
               TabIndex        =   93
               Text            =   "TxtNotDeb"
               Top             =   690
               Visible         =   0   'False
               Width           =   1905
            End
            Begin VB.TextBox TxtIntervalos 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   7770
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   12
               Text            =   "TxtIntervalos"
               Top             =   1020
               Width           =   675
            End
            Begin VB.TextBox TxtPortes 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   7770
               Locked          =   -1  'True
               MaxLength       =   5
               TabIndex        =   9
               Text            =   "TxtPortes"
               Top             =   690
               Width           =   675
            End
            Begin VB.CommandButton CmdBusTipo 
               Height          =   240
               Left            =   1755
               Picture         =   "FrmLetras.frx":0264
               Style           =   1  'Graphical
               TabIndex        =   82
               Top             =   720
               Width           =   240
            End
            Begin VB.CommandButton CmdBusTipDocRef 
               Height          =   240
               Left            =   1755
               Picture         =   "FrmLetras.frx":0396
               Style           =   1  'Graphical
               TabIndex        =   81
               Top             =   1350
               Width           =   240
            End
            Begin VB.CommandButton CmdBusDocRef2 
               Height          =   240
               Left            =   10935
               Picture         =   "FrmLetras.frx":04C8
               Style           =   1  'Graphical
               TabIndex        =   80
               Top             =   1350
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.TextBox TxtDocRef 
               Height          =   300
               Left            =   10530
               Locked          =   -1  'True
               MaxLength       =   30
               TabIndex        =   15
               Text            =   "TxtDocRef"
               Top             =   1320
               Visible         =   0   'False
               Width           =   2025
            End
            Begin VB.TextBox TxtNumLet 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   9690
               Locked          =   -1  'True
               MaxLength       =   5
               TabIndex        =   13
               Text            =   "TxtNumLet"
               Top             =   1020
               Width           =   675
            End
            Begin VB.TextBox TxtImpTasa 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   5940
               Locked          =   -1  'True
               MaxLength       =   5
               TabIndex        =   8
               Text            =   "TxtImpTasa"
               Top             =   690
               Width           =   675
            End
            Begin VB.TextBox TxtDiasPlazo 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   5940
               Locked          =   -1  'True
               MaxLength       =   5
               TabIndex        =   11
               Text            =   "TxtDiasPlazo"
               Top             =   1020
               Width           =   675
            End
            Begin VB.CommandButton CmdBusMon 
               Height          =   240
               Left            =   6360
               Picture         =   "FrmLetras.frx":05FA
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   390
               Width           =   240
            End
            Begin VB.TextBox TxtTipInt 
               Height          =   300
               Left            =   1350
               Locked          =   -1  'True
               MaxLength       =   1
               TabIndex        =   7
               Text            =   "TxtTipInt"
               Top             =   690
               Width           =   675
            End
            Begin VB.TextBox TxtImpFinan 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1350
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   10
               Text            =   "TxtImpFinan"
               Top             =   1020
               Width           =   1200
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   3630
               TabIndex        =   5
               Top             =   360
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
               Locked          =   -1  'True
               Valor           =   "  /  /    "
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
               Height          =   300
               Left            =   1350
               TabIndex        =   4
               Top             =   360
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
               Locked          =   -1  'True
               Valor           =   "  /  /    "
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   5940
               Locked          =   -1  'True
               MaxLength       =   1
               TabIndex        =   6
               Text            =   "TxtIdMon"
               Top             =   360
               Width           =   675
            End
            Begin VB.TextBox TxtIdTipDocRef 
               Height          =   300
               Left            =   1350
               Locked          =   -1  'True
               MaxLength       =   1
               TabIndex        =   14
               Text            =   "Txt"
               Top             =   1320
               Width           =   675
            End
            Begin VB.Label LblNotDeb 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nota Debito"
               Height          =   195
               Left            =   8700
               TabIndex        =   94
               Top             =   720
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Intervalos"
               Height          =   195
               Index           =   2
               Left            =   6975
               TabIndex        =   92
               Top             =   1050
               Width           =   690
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Portes"
               Height          =   195
               Left            =   7215
               TabIndex        =   91
               Top             =   720
               Width           =   450
            End
            Begin VB.Label LblTipoInteres 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblTipoInteres"
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
               Left            =   2025
               TabIndex        =   83
               Top             =   690
               Width           =   2790
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tip de Doc. Ref."
               Height          =   195
               Index           =   12
               Left            =   120
               TabIndex        =   79
               Top             =   1395
               Width           =   1185
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
               Left            =   2025
               TabIndex        =   78
               Top             =   1320
               Width           =   2790
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nº Doc. Referencia"
               Height          =   195
               Index           =   13
               Left            =   4920
               TabIndex        =   77
               Top             =   1395
               Width           =   1395
            End
            Begin VB.Label LblIdDocRef 
               AutoSize        =   -1  'True
               Caption         =   "LblIdDocRef"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   10545
               TabIndex        =   76
               Top             =   1065
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº de Letras"
               Height          =   195
               Index           =   1
               Left            =   8670
               TabIndex        =   42
               Top             =   1050
               Width           =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tasa"
               Height          =   195
               Left            =   5475
               TabIndex        =   41
               Top             =   720
               Width           =   360
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Fch Inicio"
               Height          =   195
               Left            =   2850
               TabIndex        =   40
               Top             =   405
               Width           =   690
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Financiar"
               Height          =   195
               Left            =   120
               TabIndex        =   39
               Top             =   1050
               Width           =   990
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
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   9690
               TabIndex        =   38
               Top             =   360
               Width           =   1905
            End
            Begin VB.Label LblTipCam2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   9180
               TabIndex        =   37
               Top             =   405
               Width           =   375
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fch. de Emisión"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   36
               Top             =   405
               Width           =   1125
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
               Left            =   6660
               TabIndex        =   35
               Top             =   360
               Width           =   1785
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda"
               Height          =   195
               Index           =   4
               Left            =   5250
               TabIndex        =   34
               Top             =   405
               Width           =   585
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "Dias Plazo"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   9
               Left            =   5085
               TabIndex        =   33
               Top             =   1050
               Width           =   750
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Interes"
               Height          =   195
               Left            =   120
               TabIndex        =   32
               Top             =   720
               Width           =   1065
            End
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   3285
            Left            =   0
            TabIndex        =   43
            Top             =   3450
            Width           =   11625
            _cx             =   20505
            _cy             =   5794
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
            Caption         =   "    Documentos    |      Letras      "
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
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Caption         =   "5640"
               Height          =   2865
               Left            =   12270
               TabIndex        =   48
               Top             =   45
               Width           =   11535
               Begin VB.Frame Frame9 
                  Height          =   2880
                  Left            =   9960
                  TabIndex        =   59
                  Top             =   -75
                  Width           =   1560
                  Begin VB.CommandButton CmdGenLet 
                     Caption         =   "Generar Letras"
                     Enabled         =   0   'False
                     Height          =   555
                     Left            =   105
                     TabIndex        =   21
                     Top             =   540
                     Width           =   1320
                  End
                  Begin VB.CommandButton CmdAddLet 
                     Caption         =   "Agregar Letra"
                     Enabled         =   0   'False
                     Height          =   555
                     Left            =   105
                     TabIndex        =   22
                     Top             =   1185
                     Width           =   1320
                  End
                  Begin VB.CommandButton CmdDelLet 
                     Caption         =   "Eliminar Letra"
                     Enabled         =   0   'False
                     Height          =   555
                     Left            =   105
                     TabIndex        =   23
                     Top             =   1830
                     Width           =   1320
                  End
               End
               Begin VB.TextBox TxtSInteres 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Left            =   3915
                  Locked          =   -1  'True
                  TabIndex        =   58
                  Text            =   "TxtSInteres"
                  Top             =   2235
                  Width           =   1215
               End
               Begin VB.TextBox TxtDInteres 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Left            =   3915
                  Locked          =   -1  'True
                  TabIndex        =   57
                  Text            =   "TxtDInteres"
                  Top             =   2550
                  Width           =   1215
               End
               Begin VB.TextBox TxtSPortes 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Left            =   5115
                  Locked          =   -1  'True
                  TabIndex        =   56
                  Text            =   "TxtSPortes"
                  Top             =   2235
                  Width           =   1215
               End
               Begin VB.TextBox TxtDPortes 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Left            =   5115
                  Locked          =   -1  'True
                  TabIndex        =   55
                  Text            =   "TxtDPortes"
                  Top             =   2550
                  Width           =   1215
               End
               Begin VB.TextBox TxtSIGV 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Left            =   6315
                  Locked          =   -1  'True
                  TabIndex        =   54
                  Text            =   "TxtSIGV"
                  Top             =   2235
                  Width           =   1215
               End
               Begin VB.TextBox TxtDIGV 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Left            =   6315
                  Locked          =   -1  'True
                  TabIndex        =   53
                  Text            =   "TxtDIGV"
                  Top             =   2550
                  Width           =   1215
               End
               Begin VB.TextBox TxtSTotal 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Left            =   7515
                  Locked          =   -1  'True
                  TabIndex        =   52
                  Text            =   "TxtSTotal"
                  Top             =   2235
                  Width           =   1215
               End
               Begin VB.TextBox TxtDTotal 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Left            =   7515
                  Locked          =   -1  'True
                  TabIndex        =   51
                  Text            =   "TxtDTotal"
                  Top             =   2550
                  Width           =   1215
               End
               Begin VB.TextBox TxtDCapital 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Left            =   2715
                  Locked          =   -1  'True
                  TabIndex        =   50
                  Text            =   "TxtDCapital"
                  Top             =   2550
                  Width           =   1215
               End
               Begin VB.TextBox TxtSCapital 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Left            =   2715
                  Locked          =   -1  'True
                  TabIndex        =   49
                  Text            =   "TxtSCapital"
                  Top             =   2235
                  Width           =   1215
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg2 
                  Height          =   2205
                  Left            =   0
                  TabIndex        =   20
                  Top             =   15
                  Width           =   9915
                  _cx             =   17489
                  _cy             =   3889
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmLetras.frx":072C
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
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total ==>"
                  Height          =   195
                  Left            =   9015
                  TabIndex        =   62
                  Top             =   3015
                  Width           =   675
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Soles"
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
                  Left            =   1215
                  TabIndex        =   61
                  Top             =   2265
                  Width           =   975
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Dólares"
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
                  Left            =   1215
                  TabIndex        =   60
                  Top             =   2580
                  Width           =   1155
               End
            End
            Begin VB.Frame Frame5 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   2865
               Left            =   45
               TabIndex        =   44
               Top             =   45
               Width           =   11535
               Begin VB.TextBox TxtTotal3Pro 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Height          =   285
                  Left            =   8550
                  Locked          =   -1  'True
                  TabIndex        =   96
                  Text            =   "TxtTotal3Pro"
                  Top             =   2490
                  Width           =   1005
               End
               Begin VB.TextBox TxtTotal2Pro 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
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
                  Height          =   285
                  Left            =   7545
                  Locked          =   -1  'True
                  TabIndex        =   46
                  Text            =   "TxtTotal2Pro"
                  Top             =   2490
                  Width           =   1005
               End
               Begin VB.Frame Frame8 
                  Height          =   2880
                  Left            =   9960
                  TabIndex        =   45
                  Top             =   -75
                  Width           =   1560
                  Begin VB.CommandButton CmdAddDoc 
                     Caption         =   "Agregar Documentos"
                     Enabled         =   0   'False
                     Height          =   600
                     Left            =   120
                     TabIndex        =   17
                     Top             =   570
                     Width           =   1320
                  End
                  Begin VB.CommandButton CmdDelDoc 
                     Caption         =   "Eliminar Documento"
                     Enabled         =   0   'False
                     Height          =   600
                     Left            =   120
                     TabIndex        =   19
                     Top             =   1800
                     Width           =   1320
                  End
                  Begin VB.CommandButton CmdSelDoc 
                     Caption         =   "Seleccionar Documentos"
                     Enabled         =   0   'False
                     Height          =   600
                     Left            =   120
                     TabIndex        =   18
                     Top             =   1185
                     Width           =   1320
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   2445
                  Left            =   15
                  TabIndex        =   16
                  Top             =   15
                  Width           =   9900
                  _cx             =   17462
                  _cy             =   4313
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   20
                  Cols            =   13
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmLetras.frx":0863
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
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total ==>"
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
                  Left            =   6600
                  TabIndex        =   47
                  Top             =   2520
                  Width           =   825
               End
            End
         End
         Begin VB.TextBox TxtRucPro 
            Height          =   300
            Left            =   1470
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   0
            Text            =   "TxtRucPro"
            Top             =   360
            Width           =   1530
         End
         Begin VB.Frame FraAgenRet 
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   8250
            TabIndex        =   101
            Top             =   300
            Visible         =   0   'False
            Width           =   3525
            Begin VB.CheckBox ChkRetencion 
               Caption         =   "Aplicar Retención"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1740
               TabIndex        =   102
               Top             =   60
               Width           =   1665
            End
            Begin VB.Label LblAgeReten 
               Caption         =   "Agente Retención"
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
               Height          =   255
               Left            =   60
               TabIndex        =   103
               Top             =   90
               Width           =   1665
            End
         End
         Begin VB.Label LblIdCliente 
            Appearance      =   0  'Flat
            Caption         =   "LblIdCliente"
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   7560
            TabIndex        =   64
            Top             =   90
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblReg 
            Caption         =   "lblReg"
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
            Height          =   270
            Left            =   9420
            TabIndex        =   100
            Top             =   1410
            Width           =   2250
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Letra"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   60
            TabIndex        =   68
            Top             =   60
            Width           =   11610
         End
         Begin VB.Label LblProveedor 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProveedor"
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
            Left            =   3030
            TabIndex        =   67
            Top             =   360
            Width           =   5160
         End
         Begin VB.Label LblTitulo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            Height          =   195
            Left            =   135
            TabIndex        =   66
            Top             =   390
            Width           =   480
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   8010
            TabIndex        =   65
            Top             =   60
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6750
         Left            =   -12375
         TabIndex        =   25
         Top             =   375
         Width           =   11730
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6360
            Left            =   0
            TabIndex        =   26
            Top             =   330
            Width           =   11730
            _ExtentX        =   20690
            _ExtentY        =   11218
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
            Columns(1).Caption=   "Nº Reg"
            Columns(1).DataField=   "numreg2"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Emi"
            Columns(2).DataField=   "fchemi1"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Ini"
            Columns(3).DataField=   "fchini1"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Cliente "
            Columns(4).DataField=   "nombre"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "M"
            Columns(5).DataField=   "simbolo"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Tipo Int."
            Columns(6).DataField=   "tipointe"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Nº Letras"
            Columns(7).DataField=   "numlet1"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Nº Dias"
            Columns(8).DataField=   "numdias1"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Imp. Financiar"
            Columns(9).DataField=   "impcap1"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1879"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1799"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1879"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1799"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1879"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1799"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=5583"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=5503"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=794"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=714"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1720"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1640"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=1614"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1535"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=1349"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1270"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=2275"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=2196"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
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
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(27)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(28)  =   ":id=14,.fontname=MS Sans Serif"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=70,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=43,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=44,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=45,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
            _StyleDefs(78)  =   "Named:id=33:Normal"
            _StyleDefs(79)  =   ":id=33,.parent=0"
            _StyleDefs(80)  =   "Named:id=34:Heading"
            _StyleDefs(81)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(82)  =   ":id=34,.wraptext=-1"
            _StyleDefs(83)  =   "Named:id=35:Footing"
            _StyleDefs(84)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(85)  =   "Named:id=36:Selected"
            _StyleDefs(86)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(87)  =   "Named:id=37:Caption"
            _StyleDefs(88)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(89)  =   "Named:id=38:HighlightRow"
            _StyleDefs(90)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(91)  =   "Named:id=39:EvenRow"
            _StyleDefs(92)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(93)  =   "Named:id=40:OddRow"
            _StyleDefs(94)  =   ":id=40,.parent=33"
            _StyleDefs(95)  =   "Named:id=41:RecordSelector"
            _StyleDefs(96)  =   ":id=41,.parent=34"
            _StyleDefs(97)  =   "Named:id=42:FilterBar"
            _StyleDefs(98)  =   ":id=42,.parent=33"
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Emisión de Letras"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   28
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label lblperiodo 
            Caption         =   "lblperiodo(0)"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   0
            Left            =   9450
            TabIndex        =   27
            Top             =   75
            Width           =   1980
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   69
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
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
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Registro"
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7410
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":09D6
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":0F1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":12AC
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":1430
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":1884
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":199C
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":1EE0
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":2424
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":2538
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":264C
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":2AA0
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmLetras.frx":2C0C
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstLet As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean


Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim mIdRegistro& '--identificador del registro
Dim mMesActivo As Integer '--indica el mes activo
Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)

Dim xHorIni As Date


Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    CargarGrid
    TabOne1.CurrTab = 0
End Sub

Sub CargarGrid()
    Dim Rpta As Integer
    Dim xApertura As Boolean
    
'    If Mes <> 0 Then
'        xApertura = False
'    Else
'        xApertura = True
'    End If
    
        '--limpiar los filtros
    TDB_FiltroLimpiar Dg1
    Set RstLet = Nothing
    Set Dg1.DataSource = Nothing
    '----------------------
    OpcionesPeriodo
    '----------------------
'    If xApertura = False Then
        RST_Busq RstLet, "SELECT mae_cliente.nombre,mae_cliente.ageret, mae_moneda.simbolo, mae_docreferencia.descripcion AS docref, Trim(Str([let_letra]![idaduana])) & Trim(Str([let_letra]![idregimen])) & Trim(Str([let_letra]![anoorden])) & Trim([let_letra]![numorden]) AS numdocref, " _
            & " let_letratipoplazo.descripcion AS tipointe, let_letra.*, mae_cliente.numruc, Mid([let_letra]![numreg],1,2) & [mae_libros]![codsun] & Mid([let_letra]![numreg],3,4) AS numreg2, " _
            & " let_letra.fchemi & '' as fchemi1,let_letra.fchini & '' as fchini1,let_letra.numlet & '' as numlet1,let_letra.numdias & '' as numdias1,let_letra.impcap & '' as impcap1 " _
            & " FROM (((mae_moneda RIGHT JOIN (mae_docreferencia RIGHT JOIN (mae_cliente RIGHT JOIN let_letra ON mae_cliente.id = let_letra.idclipro) ON mae_docreferencia.id = let_letra.idtipdocref) ON  " _
            & " mae_moneda.id = let_letra.idmon) LEFT JOIN mae_documento ON let_letra.tipdoc = mae_documento.id) LEFT JOIN let_letratipoplazo ON let_letra.tipint = let_letratipoplazo.id) LEFT JOIN mae_libros ON let_letra.idlib = mae_libros.id " _
            & " WHERE  let_letra.idmes= " & mMesActivo & " ", xCon
'    Else
'        RST_Busq RstLet, "SELECT mae_cliente.nombre, mae_moneda.simbolo, mae_docreferencia.descripcion AS docref, " _
'            & " Trim(Str([let_letra]![idaduana])) & Trim(Str([let_letra]![idregimen])) & Trim(Str([let_letra]![anoorden])) & Trim([let_letra]![numorden]) AS numdocref, " _
'            & " let_letratipoplazo.descripcion AS tipointe, let_letra.*, mae_cliente.numruc, Mid([let_letra]![numreg],1,2) & [mae_libros]![codsun] & Mid([let_letra]![numreg],3,4) AS numreg2, " _
'            & " let_letra.fchemi & '' AS fchemi1, let_letra.fchini & '' AS fchini1, let_letra.numlet & '' AS numlet1, let_letra.numdias & '' AS numdias1, let_letra.impcap & '' AS impcap1 " _
'            & " FROM (((mae_moneda RIGHT JOIN (mae_docreferencia RIGHT JOIN (mae_cliente RIGHT JOIN let_letra ON mae_cliente.id = let_letra.idclipro) " _
'            & " ON mae_docreferencia.id = let_letra.idtipdocref) ON mae_moneda.id = let_letra.idmon) LEFT JOIN mae_documento ON let_letra.tipdoc = mae_documento.id) " _
'            & " LEFT JOIN let_letratipoplazo ON let_letra.tipint = let_letratipoplazo.id) LEFT JOIN mae_libros ON let_letra.idlib = mae_libros.id " _
'            & " WHERE (((let_letra.idmes)=0))", xCon
'    End If
    
    LblNumRegistros.Caption = RstLet.RecordCount
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    Dg1.DataSource = RstLet
    
End Sub

Private Sub CmdAddDoc_Click()
    If TxtFchEmi.Valor = "" Then
        MsgBox "No ha especificado la fecha de emision de las letras", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Sub
    End If
    
    If TxtIdMon.Text = "" Then
        MsgBox "No ha especificado la moneda en que se emitira la letra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtRucPro.Text) = "" Then
        MsgBox "No ha especificado el nombre del cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtRucPro.SetFocus
        Exit Sub
    End If
    
    Dim nSQL As String
    
    nSQL = "SELECT Mid([vta_ventas]![numreg],1,2) & [mae_libros]![codsun] & Mid([vta_ventas]![numreg],3,4) AS numreg, [vta_ventas]![numser] & '-' & [vta_ventas]![numdoc] AS numdoc, " _
        & " mae_documento.abrev , vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo AS moneda, mae_cliente.nombre AS clipro, mae_cliente.numruc, vta_ventas.imptotdoc AS importe, " _
        & " vta_ventas.impsal as saldo, vta_ventas.idmon, 2 AS idmod, vta_ventas.id AS iddoc FROM (((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_documento " _
        & " ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id  " _
        & " WHERE vta_ventas.anulado=0 and vta_ventas.tipdoc = 8 and vta_ventas.idcli=" & NulosN(LblIdCliente.Caption) & " and vta_ventas.numerodocref='" & NulosC(TxtNumDocRef2.Text) & "' " _
        & " Union " _
        & " SELECT '' AS numreg, vta_proforma.numdoc AS numdoc, mae_documento.abrev, vta_proforma.fchdoc as fchemi, mae_moneda.simbolo, mae_cliente.nombre AS clipro, mae_cliente.numruc, " _
        & " vta_proforma.imptot AS importe, vta_proforma.imptot AS saldo, vta_proforma.idmon, 12 AS idmod, vta_proforma.id AS iddoc FROM ((vta_proforma LEFT JOIN mae_documento " _
        & " ON vta_proforma.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_proforma.idmon = mae_moneda.id) LEFT JOIN mae_cliente ON vta_proforma.idcli = mae_cliente.id " _
        & " WHERE (((vta_proforma.idcli)=" & NulosN(LblIdCliente.Caption) & "))"

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(7, 4) As String
    
    xCampos(0, 0) = "Nº Documento":    xCampos(0, 1) = "numdoc":       xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Tip. Doc.":       xCampos(1, 1) = "abrev":        xCampos(1, 2) = "900":          xCampos(1, 3) = "C"
    xCampos(2, 0) = "Fch. Emi.":       xCampos(2, 1) = "fchemi":       xCampos(2, 2) = "950":          xCampos(2, 3) = "C"
    xCampos(3, 0) = "Moneda":          xCampos(3, 1) = "moneda":       xCampos(3, 2) = "800":          xCampos(3, 3) = "C"
    xCampos(4, 0) = "Cliente":         xCampos(4, 1) = "clipro":       xCampos(4, 2) = "2000":         xCampos(4, 3) = "C"
    xCampos(5, 0) = "Importe":         xCampos(5, 1) = "importe":      xCampos(5, 2) = "1000":         xCampos(5, 3) = "N"
    xCampos(6, 0) = "Saldo":           xCampos(6, 1) = "saldo":        xCampos(6, 2) = "1000":         xCampos(6, 3) = "N"
    xCampos(7, 0) = "Nº Registro":     xCampos(7, 1) = "numreg":       xCampos(7, 2) = "1000":         xCampos(7, 3) = "C"
        
    xform.SQLCad = nSQL
    
    xform.Titulo = "Buscando Documentos del Cliente"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numdoc"
    xform.CampoBusca = "numdoc"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("numreg"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRs("numdoc")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRs("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosC(xRs("fchemi")), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = xRs("moneda")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = xRs("clipro")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(xRs("importe"), FORMAT_MONTO)
            If xRs("idmon") = 1 Then
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosN(xRs("impsal")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(xRs("impsal")) / NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(xRs("saldo")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosN(xRs("saldo")) * NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            End If
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = xRs("idmod")
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = xRs("iddoc")
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = xRs("idmon")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
        
    SumarColumnaDocumentos
'    CmdGenLet_Click
    Fg1.SetFocus
End Sub

Sub SumarColumnaDocumentos()
    TxtTotal2Pro.Text = GRID_SUMAR_COL(Fg1, 8)
    TxtTotal2Pro.Text = Format(TxtTotal2Pro.Text, FORMAT_MONTO)
    
    TxtTotal3Pro.Text = GRID_SUMAR_COL(Fg1, 9)
    TxtTotal3Pro.Text = Format(TxtTotal3Pro.Text, FORMAT_MONTO)
 
End Sub

Private Sub CmdBusDocRef2_Click()
    MsgBox ""
End Sub

Private Sub CmdBusIdDoc_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "codigo":           xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_dociden.* From mae_dociden"
    
    xform.Titulo = "Buscando Moneda"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdDocIden.Text = xRs("id")
            LblDocIdentidad.Caption = xRs("descripcion")
            TxtNumDoc.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "codigo":           xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_moneda.* From mae_moneda"
    
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
            SumarColumnaDocumentos
            TxtTipInt.SetFocus
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
    
    xCampos(0, 0) = "Cliente":      xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id, mae_cliente.idven, mae_cliente.ageret From mae_cliente"
    
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
                TxtRucPro.Text = xRs("numruc")
                LblProveedor.Caption = xRs("nombre")
                LblIdCliente.Caption = xRs("id")
                '--evaluar si es agente de retencion
                If NulosN(xRs("ageret")) = -1 Then
                    FraAgenRet.Visible = True
                    ChkRetencion.value = -1
                Else
                    FraAgenRet.Visible = False
                End If
                '------------------------------------
                
                
                
                
                TxtGirado.SetFocus
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDocRef_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "codigo":           xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_docreferencia.* From mae_docreferencia "
    
    xform.Titulo = "Buscando Tipos de Documentos de Referencia"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdTipDocRef.Text = xRs("id")
            LblTipDocref.Caption = xRs("descripcion")
            TxtDocRef.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipo_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "codigo":           xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT let_letratipoplazo.* From let_letratipoplazo"
    
    xform.Titulo = "Buscando Tipo de Plazo"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipInt.Text = xRs("id")
            LblTipoInteres.Caption = xRs("descripcion")
            TxtImpTasa.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDelDoc_Click()
    If Fg1.Row < Fg1.FixedRows Then Exit Sub
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha documentos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim Rpta As Integer
    
    Rpta = MsgBox("Esta seguro de eliminar el documento", vbInformation + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Fg1.RemoveItem Fg1.Row
        
        SumarColumnaDocumentos

    End If
    Fg1.SetFocus
End Sub

Private Sub CmdGenLet_Click()
    Dim xNumeroLetra As Integer
    xNumeroLetra = InputBox("Numero de Letra Inicial ==> ", xTitulo)
    
    GenerarLetras NulosN(TxtTipInt.Text), NulosN(TxtImpTasa.Text), NulosN(TxtDiasPlazo.Text), NulosN(TxtNumLet.Text), NulosN(TxtImpFinan.Text), xNumeroLetra, NulosN(TxtPortes.Text), TxtFchIni.Valor, NulosN(TxtIntervalos.Text)
End Sub


Private Sub GenerarAsiento1()
    Dim xCampos(6, 4) As String
    
'    xCampos(0, 0) = "id":             xCampos(0, 1) = "2":              xCampos(0, 2) = "S":    xCampos(0, 3) = "N":    xCampos(0, 4) = "":
'    xCampos(1, 0) = "descripcion":    xCampos(1, 1) = "2 descripcion":  xCampos(1, 2) = "S":    xCampos(1, 3) = "C":    xCampos(1, 4) = "Ingrese la descripcion"
'    xCampos(2, 0) = "glosa":          xCampos(2, 1) = "2222222222222":  xCampos(2, 2) = "S":    xCampos(2, 3) = "M":    xCampos(2, 4) = "no ha especificado la glosa"
'    xCampos(3, 0) = "fecha":          xCampos(3, 1) = "02/02/02":       xCampos(3, 2) = "S":    xCampos(3, 3) = "F":    xCampos(2, 4) = "No ha especificado la fecha"
'    xCampos(4, 0) = "hora":           xCampos(4, 1) = "2:22":           xCampos(4, 2) = "S":    xCampos(4, 3) = "H":    xCampos(2, 4) = "No ha especificado la hora"
'    xCampos(5, 0) = "logico":         xCampos(5, 1) = "-1":             xCampos(5, 2) = "S":    xCampos(5, 3) = "L":    xCampos(2, 4) = "No ha especificado si es verdadero o falso"
'
'    'Columna    | Descripcion
'    '------------------------
'    '0          | campo
'    '1          | Valor
'    '2          | requerido
'    '3          | tipo
'    '4          | msj error
'
'    If EscribirNuevoRegistro(xCampos, "aaaa", xCon) = True Then
'        MsgBox "Se Grabo con exito"
'    End If
'    Exit Sub


    xCon.Execute "DELETE con_diario.idmes, con_diario.idlib, con_diario.* from con_diario WHERE (((con_diario.idmes)=" & mMesActivo & ") AND ((con_diario.idlib)=37))"

    Frame6.Visible = True
    ProgressBar1.Max = RstLet.RecordCount
    
    Dim A, B As Integer
    Dim RstDet As New ADODB.Recordset
    Dim RstRet As New ADODB.Recordset
    Dim xNumAsiento As String
    Dim TC As Double
    
    RstLet.MoveFirst
    Dim xCtaLetHab, CtaNotHab, xCtaLetDeb As Integer
    Dim RstDia As New ADODB.Recordset
    
    '37 libro letras
    RST_Busq RstDia, "SELECT * FROM con_diario", xCon
    
'    RST_Busq RstRet, "SELECT Consulta9.id, (([Consulta7]![total]+[Consulta7]![totalnodeb]))-[Consulta9]![implet] AS resta, Consulta7.totalnodeb, " _
'            & " ((([Consulta7]![total]+[Consulta7]![totalnodeb])-[Consulta9]![implet])/[totalnodeb]) AS porcentaje " _
'            & " FROM " _
'            & " (SELECT let_letradet.idlet, Sum(let_letradet.impcapital) AS total, Sum([impporte]+[impinteres]+[impigv]) AS totalnodeb " _
'            & " From let_letradet GROUP BY let_letradet.idlet) as  Consulta7 " _
'            & " RIGHT JOIN " _
'            & " (SELECT let_letra.id, let_letra.implet FROM let_letra ) as Consulta9 " _
'            & " ON Consulta7.idlet = Consulta9.id WHERE ((((([Consulta7]![total]+[Consulta7]![totalnodeb]))-[Consulta9]![implet])>1))", xCon

    For A = 1 To RstLet.RecordCount
    
        ProgressBar1.value = A
        Frame6.Refresh
        
        If RstLet("idmon") = 1 Then
            xCtaLetDeb = 73
            xCtaLetHab = 69
            CtaNotHab = 65
        Else
            xCtaLetDeb = 74
            xCtaLetHab = 70
            CtaNotHab = 66
        End If
        
        TC = HallaTipoCambio(RstLet("fchemi"), 2, Venta, xCon)
        '------------------------------
        'escribimos el debe del diarios
        RST_Busq RstDet, "SELECT let_letradet.* From let_letradet WHERE (((let_letradet.idlet)=" & RstLet("id") & "))", xCon
        
        xNumAsiento = NuevoNumAsiento(37, mMesActivo, xCon)
        If RstDet.RecordCount <> 0 Then
            RstDet.MoveFirst
            For B = 1 To RstDet.RecordCount
                   
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = mMesActivo
                RstDia("idlib") = 37
                RstDia("idmov") = RstLet("id")
                RstDia("numasi") = xNumAsiento
                RstDia("tc") = TC
                If mMesActivo = 0 Then
                    RstDia("fchasi") = "01/" + Format(1, "00") + "/" + Trim(Str(AnoTra))
                Else
                    RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
                End If
                RstDia("fchdoc") = RstLet("fchemi")
                RstDia("idcue") = xCtaLetDeb
                RstDia("correlativo") = RstDet("corr")
                'documento de referencia
                RstDia("rtipdoc1") = 108 'RstLet("idtipdocref")
                RstDia("rnumerodoc1") = RstLet("numdocref")
                
                RstDia("rtipdoc") = 95
                RstDia("rnumerodoc") = Trim(RstLet("ano")) & " " & Format(RstDet("numdoc"), "00000000") & " " & Format(RstDet("numser"), "00")
                RstDia("rfchope") = RstLet("fchemi")
                RstDia("rregistro") = Format(mMesActivo, "00") & "LE" & xNumAsiento
                RstDia("ridtipper") = 2
                RstDia("ridper") = RstLet("idclipro")
                
                RstDia("rtipdoc") = 95
                If RstLet("idmon") = 1 Then
                    RstDia("impdebsol") = RstDet("implet")
                    RstDia("impdebdol") = 0
                Else
                    RstDia("impdebsol") = NulosN(RstDet("implet")) * TC
                    RstDia("impdebdol") = NulosN(RstDet("implet"))
                End If
                RstDia.Update
                
                RstDet.MoveNext
                If RstDet.EOF = True Then
                    Exit For
                End If
            Next B
            
            
            
        End If

        '-------------------
        'ESCRIBIMOS EL HABER
        'escribimos la cuenta anticipos
        RST_Busq RstDet, "SELECT Sum([impcapital]) AS xTotal From let_letradet WHERE (((let_letradet.idlet)=" & RstLet("id") & "))", xCon

        RstDia.AddNew
        RstDia("año") = AnoTra
        RstDia("idmes") = mMesActivo
        RstDia("idlib") = 37
        RstDia("idmov") = RstLet("id")
        RstDia("numasi") = xNumAsiento
        RstDia("tc") = TC
        'RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
        If mMesActivo = 0 Then
            RstDia("fchasi") = "01/" + Format(1, "00") + "/" + Trim(Str(AnoTra))
        Else
            RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
        End If
        
        RstDia("fchdoc") = RstLet("fchemi")
        RstDia("idcue") = xCtaLetHab
        RstDia("rregistro") = Format(mMesActivo, "00") & "LE" & xNumAsiento
        
        RstDia("rtipdoc") = 8
        RstDia("rnumerodoc") = Mid(RstLet("notdeb"), 1, 15)
        RstDia("ridtipper") = 2
        RstDia("ridper") = RstLet("idclipro")
        
        
        'documento de referencia
        RstDia("rtipdoc1") = 108 'RstLet("idtipdocref")
        RstDia("rnumerodoc1") = RstLet("numdocref")
       
        If RstLet("idmon") = 1 Then
            RstDia("imphabsol") = RstDet("xTotal")
            RstDia("imphabdol") = 0
        Else
            RstDia("imphabsol") = NulosN(RstDet("xTotal")) * TC
            RstDia("imphabdol") = NulosN(RstDet("xTotal"))
        End If
        
        RstDia.Update
                    
        'escribimos la cuenta nota de debito
        RST_Busq RstDet, "SELECT Sum([let_letradet]![impporte]+[let_letradet]![impinteres]+[let_letradet]![impigv]) AS xtotal From let_letradet " _
            & " WHERE (((let_letradet.idlet)=" & RstLet("id") & "))", xCon

        If RstDet.RecordCount <> 0 Then
            RstDet.MoveFirst
            RstDia.AddNew
            RstDia("año") = AnoTra
            RstDia("idmes") = mMesActivo
            RstDia("idlib") = 37
            RstDia("idmov") = RstLet("id")
            RstDia("numasi") = xNumAsiento
            RstDia("tc") = TC
            'RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
            If mMesActivo = 0 Then
                RstDia("fchasi") = "01/" + Format(1, "00") + "/" + Trim(Str(AnoTra))
            Else
                RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
            End If
            
            RstDia("fchdoc") = RstLet("fchemi")
            RstDia("idcue") = CtaNotHab
            RstDia("rregistro") = Format(mMesActivo, "00") & "LE" & xNumAsiento
            
            RstDia("ridtipper") = 2
            RstDia("ridper") = RstLet("idclipro")
            
            'documento de referencia
            RstDia("rtipdoc1") = 108 'RstLet("idtipdocref")
            RstDia("rnumerodoc1") = RstLet("numdocref")
            
            If RstLet("idmon") = 1 Then
                RstDia("imphabsol") = RstDet("xTotal")
                RstDia("imphabdol") = 0
            Else
                RstDia("imphabsol") = NulosN(RstDet("xTotal")) * TC
                RstDia("imphabdol") = NulosN(RstDet("xTotal"))
            End If
            
            RstDia.Update
        End If
        'xNumAsiento = Format(mMesActivo, "00") & xNumAsiento
        xCon.Execute "UPDATE let_letra SET let_letra.numreg = '" & Format(mMesActivo, "00") & xNumAsiento & "' WHERE (((let_letra.id)=" & RstLet("id") & "))"

'        'escribimos el asiento para las notas de debito que tengan retencion
'        RstRet.Filter = adFilterNone
'        RstRet.Filter = "id = " & RstLet("id") & ""
'
'        If RstRet.RecordCount <> 0 Then
'            RstDet.MoveFirst
'            RstDia.AddNew
'            RstDia("año") = AnoTra
'            RstDia("idmes") = mMesActivo
'            RstDia("idlib") = 37
'            RstDia("idmov") = RstLet("id")
'            RstDia("numasi") = xNumAsiento
'            RstDia("tc") = TC
'            RstDia("fchasi") = "01/" + Format(mMesActivo, "00") + "/" + Trim(Str(AnoTra))
'            RstDia("fchdoc") = RstLet("fchemi")
'            RstDia("rregistro") = Format(mMesActivo, "00") & "LE" & xNumAsiento
'
''            RstDia("ridtipper") = 2
''            RstDia("ridper") = RstLet("idclipro")
'
'            'documento de referencia
'            RstDia("rtipdoc1") = 108 'RstLet("idtipdocref")
'            RstDia("rnumerodoc1") = RstLet("numdocref")
'
'            If RstLet("idmon") = 1 Then
'                RstDia("idcue") = 65
'                RstDia("impdebsol") = RstRet("resta")
'                RstDia("impdebdol") = 0
'            Else
'                RstDia("idcue") = 66
'                RstDia("impdebsol") = NulosN(RstRet("resta")) * TC
'                RstDia("impdebdol") = NulosN(RstRet("resta"))
'            End If
'
'            RstDia.Update
'
'        End If
        
        RstLet.MoveNext
        If RstLet.EOF = True Then Exit For
    Next A
    Frame6.Visible = False
End Sub



Private Sub Command1_Click()
    GenerarAsiento1
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLet
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)

    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLet.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
    
End Sub

Private Sub Dg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TabOne1.CurrTab = 1
    End If
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 19, NulosN(RstLet("id")), xCon
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        CmdAddDoc_Click
    End If

    If KeyCode = 46 Then
        CmdDelDoc_Click
    End If
End Sub


Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col = 3 Then
        Dim Factor As Double
        Factor = HallarFactor(NulosN(TxtImpTasa.Text), NulosN(Fg2.TextMatrix(Fg2.Row, 3)), NulosN(TxtTipInt.Text))
        Fg2.TextMatrix(Fg2.Row, 5) = Format(Fg2.TextMatrix(Fg2.Row, 4) * Factor, FORMAT_MONTO)
        'Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(Portes, FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Row, 7) = ((NulosN(Fg2.TextMatrix(Fg2.Row, 5)) + NulosN(Fg2.TextMatrix(Fg2.Row, 6))) * 0.19)
        Fg2.TextMatrix(Fg2.Row, 7) = Format(Fg2.TextMatrix(Fg2.Row, 7), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Row, 8) = NulosN(Fg2.TextMatrix(Fg2.Row, 4)) + NulosN(Fg2.TextMatrix(Fg2.Row, 5)) + NulosN(Fg2.TextMatrix(Fg2.Row, 6)) + NulosN(Fg2.TextMatrix(Fg2.Row, 7))
        Fg2.TextMatrix(Fg2.Row, 8) = Format(Fg2.TextMatrix(Fg2.Row, 8), FORMAT_MONTO)
        
        If NulosC(TxtFchIni.Valor) <> "" Then
            Fg2.TextMatrix(Fg2.Row, 9) = CDate(Fg2.TextMatrix(Fg2.Row, 3)) + CDate(TxtFchIni.Valor)
            Fg2.TextMatrix(Fg2.Row, 9) = Format(Fg2.TextMatrix(Fg2.Row, 9), "dd/mm/yy")
        End If
        
        SumarColumnas
    End If
End Sub

Private Sub Fg2_EnterCell()
     If Fg2.Col = 3 Then
        Fg2.SelectionMode = flexSelectionFree
        
        Fg2.Editable = flexEDKbdMouse
    Else
        Fg2.SelectionMode = flexSelectionByRow
        Fg2.Editable = flexEDNone
     End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim Rpta As Integer
        mMesActivo = xMes
        CargarGrid
        SeEjecuto = True
        If RstLet.RecordCount = 0 Then
            Rpta = MsgBox("No se han registrado letras ¿Desea agregar uno ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            End If
        End If
        
    End If
End Sub

Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivarTool
    Blanquea
    Bloquea
    Label5.Caption = "Agregando Nueva Letra"
    Fg2.Rows = 1
    xHorIni = Time

    TxtRucPro.SetFocus
End Sub

Sub Modificar()
    QueHace = 2
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivarTool
    Blanquea
    Bloquea
    Label5.Caption = "Modificando Letra"
    Fg2.Rows = 1
    MuestraSegundoTab
    xHorIni = Time
    
    TxtRucPro.SetFocus
End Sub

Sub ActivarTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = 3
    Fg2.ColWidth(10) = 0
    TabOne1.CurrTab = 0
    
    Dg1.Columns("fchemi1").NumberFormat = FORMAT_DATE
    Dg1.Columns("fchini1").NumberFormat = FORMAT_DATE
    Dg1.Columns("numdias1").NumberFormat = FORMAT_CANTIDAD
    Dg1.Columns("impcap1").NumberFormat = FORMAT_MONTO
    
    '--color de fondo igual al formulario
    Frame3.BackColor = &H8000000F
    FraAgenRet.BackColor = &H8000000F
    
    
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            If RstLet.RecordCount = 0 Then
                Cancel = 1
                Exit Sub
            End If
            
            MuestraSegundoTab
        End If
    End If
End Sub

Sub Bloquea()
    TxtRucPro.Locked = Not TxtRucPro.Locked
    TxtGirado.Locked = Not TxtGirado.Locked
    TxtIdDocIden.Locked = Not TxtIdDocIden.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    TxtFchIni.Locked = Not TxtFchIni.Locked
    TxtTipInt.Locked = Not TxtTipInt.Locked
    TxtImpTasa.Locked = Not TxtImpTasa.Locked
    TxtDiasPlazo.Locked = Not TxtDiasPlazo.Locked
    TxtNumLet.Locked = Not TxtNumLet.Locked
    'TxtImpFinan.Locked = Not TxtImpFinan.Locked
    TxtIdTipDocRef.Locked = Not TxtIdTipDocRef.Locked
    TxtPortes.Locked = Not TxtPortes.Locked
    TxtIntervalos.Locked = Not TxtIntervalos.Locked
    TxtIdDocIden.Locked = Not TxtIdDocIden.Locked
    TxtGirado.Locked = Not TxtGirado.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtNumDocRef2.Locked = Not TxtNumDocRef2.Locked
    CmdAddLet.Enabled = Not CmdAddLet.Enabled
    CmdDelLet.Enabled = Not CmdDelLet.Enabled
    CmdGenLet.Enabled = Not CmdGenLet.Enabled
        
    ChkRetencion.Enabled = Not ChkRetencion.Enabled
    
    CmdAddDoc.Enabled = Not CmdAddDoc.Enabled
    CmdSelDoc.Enabled = Not CmdSelDoc.Enabled
    CmdDelDoc.Enabled = Not CmdDelDoc.Enabled
    
End Sub

Sub MuestraSegundoTab()

    Blanquea
    If RstLet.EOF = True Or RstLet.BOF = True Or RstLet.RecordCount = 0 Then Exit Sub
    
    lblReg.Caption = "Nº Reg. " & NulosC(RstLet("numreg2"))
    
    TxtRucPro.Text = RstLet("numruc")
    LblProveedor.Caption = RstLet("nombre")
    LblIdCliente.Caption = RstLet("idclipro")
    
    '--evaluar si es agente de retencion
    If NulosN(RstLet("ageret")) = -1 Then
        FraAgenRet.Visible = True
        ChkRetencion.value = Abs(NulosN(RstLet("aplicaret")))
    Else
        FraAgenRet.Visible = False
    End If
    '------------------------------------
    
    
    
    TxtGirado.Text = NulosC(RstLet("girado"))
    TxtIdDocIden.Text = NulosN(RstLet("iddocgir"))
    LblDocIdentidad.Caption = Busca_Codigo(NulosN(RstLet("iddocgir")), "id", "descripcion", "mae_dociden", "N", xCon)
    TxtNumDoc.Text = NulosC(RstLet("numdocgir"))
    TxtFchEmi.Valor = Format(RstLet("fchemi"), "dd/mm/yyyy")
    TxtFchIni.Valor = Format(RstLet("fchini"), "dd/mm/yyyy")
    TxtIdMon.Text = RstLet("idmon")
    LblMoneda.Caption = Busca_Codigo(RstLet("idmon"), "id", "descripcion", "mae_moneda", "N", xCon)
    TxtTipInt.Text = RstLet("tipint")
    LblTipoInteres.Caption = Busca_Codigo(RstLet("tipint"), "id", "descripcion", "let_letratipoplazo", "N", xCon)
    TxtImpTasa.Text = Format(NulosN(RstLet("inttasa")), FORMAT_MONTO)
    TxtDiasPlazo.Text = NulosN(RstLet("numdias"))
    TxtNumLet.Text = NulosN(RstLet("numlet"))
    TxtImpFinan.Text = Format(NulosN(RstLet("impcap")), FORMAT_MONTO)
    TxtIdTipDocRef.Text = NulosN(RstLet("idtipdocref"))
    LblTipDocref.Caption = Busca_Codigo(RstLet("idtipdocref"), "id", "descripcion", "mae_docreferencia", "N", xCon)
    TxtPortes.Text = NulosN(RstLet("imppor"))
    TxtIntervalos.Text = NulosN(RstLet("diaint"))
    
    If NulosN(RstLet("idtipdocref")) = 0 Then
        TxtIdTipDocRef.Text = ""
        TxtDocRef.Text = ""
    Else
        TxtIdTipDocRef.Text = NulosN(RstLet("idtipdocref"))
        TxtDocRef.Text = "" ' aqui cargamos el numero del documento de referencia
        LblIdDocRef.Caption = "" ' aqui cargamos el id del documento de referencia
    End If
    
    TxtNumDocRef2.Text = NulosC(RstLet("numerodocref"))
    
    LblTipoCambio.Caption = HallaTipoCambio(TxtFchEmi.Valor, "2", Venta, xCon)
    
    'CARGAMOS LOS DOCUMENTOS
    Dim RstDoc As New ADODB.Recordset
    RST_Busq RstDoc, "SELECT '' AS numreg, vta_proforma.numdoc AS numdoc, 'Prf' AS tipdoc, vta_proforma.fchdoc, mae_moneda.simbolo AS moneda, mae_cliente.nombre, " _
        & " vta_proforma.imptot AS imptotdoc, let_letradoc.idmod, let_letradoc.iddoc, let_letradoc.impfin AS impsal, vta_proforma.idmon FROM ((let_letradoc " _
        & " RIGHT JOIN vta_proforma ON let_letradoc.iddoc = vta_proforma.id) LEFT JOIN mae_cliente ON vta_proforma.idcli = mae_cliente.id) LEFT JOIN mae_moneda " _
        & " ON vta_proforma.idmon = mae_moneda.id WHERE (((let_letradoc.idmod)=12) AND ((let_letradoc.idlet)=" & NulosN(RstLet("id")) & ")) " _
        & " Union " _
        & " SELECT Mid([vta_ventas]![numreg],1,2) & [mae_libros]![codsun] & Mid([vta_ventas]![numreg],3,4) AS numreg, [vta_ventas]![numser] & '-' & [vta_ventas]![numdoc] AS numdoc, " _
        & " mae_documento.abrev AS tipdoc, vta_ventas.fchdoc, mae_moneda.simbolo AS moneda, mae_cliente.nombre, vta_ventas.imptotdoc, let_letradoc.idmod, " _
        & " let_letradoc.iddoc, let_letradoc.impfin AS impsal, vta_ventas.idmon FROM ((let_letradoc LEFT JOIN ((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) " _
        & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON let_letradoc.iddoc = vta_ventas.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) " _
        & " LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id WHERE (((let_letradoc.idmod)=2) AND ((let_letradoc.idlet)=" & NulosN(RstLet("id")) & "))", xCon
    
    Fg1.Rows = 1
   
    Dim A As Integer
    Agregando = True
    If RstDoc.RecordCount <> 0 Then
        RstDoc.MoveFirst
        
        For A = 1 To RstDoc.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstDoc("numreg"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDoc("numdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstDoc("tipdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(RstDoc("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(RstDoc("moneda"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(RstDoc("nombre"))
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(RstDoc("imptotdoc"), FORMAT_MONTO)
            
            If RstDoc("idmon") = 1 Then
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosN(RstDoc("impsal")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(RstDoc("impsal")) / NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(RstDoc("impsal")), FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosN(RstDoc("impsal")) * NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            End If
            
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = RstDoc("idmod")
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = RstDoc("iddoc")
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(RstDoc("idmon"))
            RstDoc.MoveNext
            
            If RstDoc.EOF = True Then Exit For
        Next A
    End If
    
    SumarColumnaDocumentos
    
    'CARGAMOS LAS LETRAS
    Dim RstDet As New ADODB.Recordset
    
    RST_Busq RstDet, "SELECT let_letradet.* From let_letradet Where (((let_letradet.idlet) = " & RstLet("id") & ")) ORDER BY let_letradet.fchven", xCon

    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        Fg2.Rows = 1
        For A = 1 To RstDet.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = RstDet("corr")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = RstDet("numdoc")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = RstDet("diasplazo")
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(RstDet("impcapital")), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(NulosN(RstDet("impinteres")), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(NulosN(RstDet("impporte")), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 7) = Format(NulosN(RstDet("impigv")), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(NulosN(RstDet("implet")), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 9) = Format(NulosC(RstDet("fchven")), "dd/mm/yy")
            
            RstDet.MoveNext
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
    End If
    Agregando = False
    If NulosN(TxtIdMon.Text) = 1 Then
        TxtSCapital.Text = GRID_SUMAR_COL(Fg2, 4)
        TxtSInteres.Text = GRID_SUMAR_COL(Fg2, 5)
        TxtSPortes.Text = GRID_SUMAR_COL(Fg2, 6)
        TxtSIGV.Text = GRID_SUMAR_COL(Fg2, 7)
        TxtSTotal.Text = GRID_SUMAR_COL(Fg2, 8)
    
        TxtDCapital.Text = NulosN(TxtSCapital.Text) / NulosN(LblTipoCambio.Caption)
        TxtDInteres.Text = NulosN(TxtSInteres.Text) / NulosN(LblTipoCambio.Caption)
        TxtDPortes.Text = NulosN(TxtSPortes.Text) / NulosN(LblTipoCambio.Caption)
        TxtDIGV.Text = NulosN(TxtSIGV.Text) / NulosN(LblTipoCambio.Caption)
        TxtDTotal.Text = NulosN(TxtSTotal.Text) / NulosN(LblTipoCambio.Caption)
    Else
    
        TxtDCapital.Text = GRID_SUMAR_COL(Fg2, 4)
        TxtDInteres.Text = GRID_SUMAR_COL(Fg2, 5)
        TxtDPortes.Text = GRID_SUMAR_COL(Fg2, 6)
        TxtDIGV.Text = GRID_SUMAR_COL(Fg2, 7)
        TxtDTotal.Text = GRID_SUMAR_COL(Fg2, 8)
    
        TxtSCapital.Text = NulosN(TxtDCapital.Text) * NulosN(LblTipoCambio.Caption)
        TxtSInteres.Text = NulosN(TxtDInteres.Text) * NulosN(LblTipoCambio.Caption)
        TxtSPortes.Text = NulosN(TxtDPortes.Text) * NulosN(LblTipoCambio.Caption)
        TxtSIGV.Text = NulosN(TxtDIGV.Text) * NulosN(LblTipoCambio.Caption)
        TxtSTotal.Text = NulosN(TxtDTotal.Text) * NulosN(LblTipoCambio.Caption)
    End If
    
    TxtSCapital.Text = Format(TxtSCapital.Text, FORMAT_MONTO)
    TxtSInteres.Text = Format(TxtSInteres.Text, FORMAT_MONTO)
    TxtSPortes.Text = Format(TxtSPortes.Text, FORMAT_MONTO)
    TxtSIGV.Text = Format(TxtSIGV.Text, FORMAT_MONTO)
    TxtSTotal.Text = Format(TxtSTotal.Text, FORMAT_MONTO)
    
    TxtDCapital.Text = Format(TxtDCapital.Text, FORMAT_MONTO)
    TxtDInteres.Text = Format(TxtDInteres.Text, FORMAT_MONTO)
    TxtDPortes.Text = Format(TxtDPortes.Text, FORMAT_MONTO)
    TxtDIGV.Text = Format(TxtDIGV.Text, FORMAT_MONTO)
    TxtDTotal.Text = Format(TxtDTotal.Text, FORMAT_MONTO)
End Sub

Sub Blanquea()
    lblReg.Caption = ""
    
    TxtRucPro.Text = ""
    TxtGirado.Text = ""
    TxtNumDoc.Text = ""
    
    TxtFchEmi.Valor = ""
    TxtFchIni.Valor = ""
    TxtIdMon.Text = ""
    LblTipoCambio.Caption = ""
    TxtTipInt.Text = ""
    TxtImpTasa.Text = ""
    TxtDiasPlazo.Text = ""
    TxtNumLet.Text = ""
    TxtImpFinan.Text = ""
    TxtIdTipDocRef.Text = ""
    TxtDocRef.Text = ""
    LblIdDocRef.Caption = ""
    TxtPortes.Text = ""
    TxtIdDocIden.Text = ""
    TxtIntervalos.Text = ""
    TxtTotal2Pro.Text = ""
    
    LblProveedor.Caption = ""
    LblIdCliente.Caption = ""
    LblDocIdentidad.Caption = ""
    LblMoneda.Caption = ""
    LblTipoCambio.Caption = ""
    LblTipoInteres.Caption = ""
    LblTipDocref.Caption = ""
    
    TxtSCapital.Text = ""
    TxtSInteres.Text = ""
    TxtSPortes.Text = ""
    TxtSIGV.Text = ""
    TxtSTotal.Text = ""
    
    TxtDCapital.Text = ""
    TxtDInteres.Text = ""
    TxtDPortes.Text = ""
    TxtDIGV.Text = ""
    TxtDTotal.Text = ""
    
    TxtNumDocRef2.Text = ""
    Fg1.Rows = 1
    Fg2.Rows = 1
    
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    If RstLet.EOF = True Or RstLet.BOF = True Or RstLet.RecordCount = 0 Then
        MsgBox "No hay registro para eliminar", vbExclamation, xTitulo
        Exit Sub
    End If
    
    
    Rpta = MsgBox("Esta seguro de eliminar la letra seleccionada", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        'si no tiene movimientos la letra la eliminamos
        If BuscarMovimientosLetra(RstLet("id"), 1) = False Then
            xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & RstLet("id") & " and idlib = 37 "
            
            xCon.Execute "DELETE * FROM let_letradet WHERE idlet = " & RstLet("id") & ""
            xCon.Execute "DELETE * FROM let_letradoc WHERE idlet = " & RstLet("id") & ""
            
            xCon.Execute "DELETE * FROM let_letra WHERE id = " & RstLet("id") & ""
            RstLet.Requery
            Dg1.Refresh
        End If
    End If
End Sub

Function BuscarMovimientosLetra(idLetra As Integer, Tipo As Integer) As Boolean
    'valores para tipo
    'tipo = 1  Ingresos
    'tipo = 2  Egresos
    Dim Rst As New ADODB.Recordset
    
    BuscarMovimientosLetra = False
    
    If Tipo = 1 Then
        'BUSCAMOS MOVIMIENTOS EN PLANILLA DE LETRAS
        RST_Busq Rst, "SELECT let_planilladet.idlet From let_planilladet WHERE (((let_planilladet.idlet)=" & idLetra & "))", xCon
        If Rst.RecordCount <> 0 Then
            BuscarMovimientosLetra = True
            Set Rst = Nothing
            Exit Function
        End If
        Set Rst = Nothing
        
        'BUSCAMOS EN CAJA Y BANCOS INGRESOS
        
    End If
    
    If Tipo = 2 Then
        'BUSCAMOS EN CAJA Y BANCOS EGRESOS
    End If
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            
            Cancelar
            RstLet.Requery
            Dg1.Refresh
            If RstLet.RecordCount <> 0 Then RstLet.MoveFirst
            RstLet.Find "id = " & mIdRegistro & ""
            If RstLet.EOF = False Then RstLet.MoveFirst
            
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 6 Then
        '--limpiar los filtros
        RstLet.Filter = ""
        TDB_FiltroLimpiar Dg1
            
    End If
    
    If Button.Index = 11 Then
        CambiarMes
    End If
    If Button.Index = 15 Then
        Set RstLet = Nothing
        Unload Me
    End If
End Sub

Private Sub TxtDiasPlazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDocRef2_Click
    End If
End Sub

Private Sub TxtFchEmi_Validate(Cancel As Boolean)
    If TxtFchEmi.Valor = "" Then Exit Sub
    
    LblTipoCambio.Caption = HallaTipoCambio(TxtFchEmi.Valor, 2, Venta, xCon)
End Sub

Private Sub TxtFchIni_Validate(Cancel As Boolean)
    If NulosC(TxtFchIni.Valor) = "" Then Exit Sub
End Sub

Private Sub TxtGirado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDocIden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdDocIden_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusIdDoc_Click
    End If
End Sub

Private Sub TxtIdDocIden_Validate(Cancel As Boolean)
    If NulosN(TxtIdDocIden.Text) = 0 Then
        LblDocIdentidad.Caption = ""
        TxtNumDoc.Text = ""
        Exit Sub
    End If
    
    LblDocIdentidad.Caption = Busca_Codigo(NulosN(TxtIdDocIden.Text), "id", "descripcion", "mae_dociden", "N", xCon)
    
    If NulosC(LblDocIdentidad.Caption) = "" Then
        TxtIdDocIden.Text = ""
    End If
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
    If NulosN(TxtIdMon.Text) = 0 Then
        LblMoneda.Caption = ""
        Exit Sub
    End If
    
    LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
    If LblMoneda.Caption = "" Then
        TxtIdMon.Text = ""
    End If
    SumarColumnaDocumentos
End Sub

Private Sub TxtIdTipDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdTipDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDocRef_Click
    End If
End Sub

Private Sub TxtIdTipDocRef_Validate(Cancel As Boolean)
    If NulosN(TxtIdTipDocRef.Text) = 0 Then
        LblTipDocref.Caption = ""
        TxtNumDocRef2.Text = ""
        Exit Sub
    End If
    
    LblTipDocref.Caption = Busca_Codigo(TxtIdTipDocRef.Text, "id", "descripcion", "mae_docreferencia", "N", xCon)
    
    If NulosC(LblTipDocref.Caption) = "" Then
        TxtIdTipDocRef.Text = ""
    End If
End Sub

Private Sub TxtImpFinan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtImpFinan_Validate(Cancel As Boolean)
    If NulosN(TxtImpFinan.Text) = 0 Then Exit Sub
    TxtImpFinan.Text = Format(NulosN(TxtImpFinan), FORMAT_MONTO)
End Sub

Private Sub TxtImpTasa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtImpTasa_Validate(Cancel As Boolean)
    If NulosN(TxtImpTasa.Text) = 0 Then
        If NulosN(TxtPortes.Text) <> 0 Then Exit Sub
        LblNotDeb.Visible = False
        TxtNotDeb.Visible = False
        TxtNotDeb.Text = ""
        Exit Sub
    End If
    
    LblNotDeb.Visible = True
    TxtNotDeb.Visible = True
    TxtNotDeb.Text = "0001" & "-" & HallaNumdocVenta(8, "0001", xCon)
    
    TxtImpTasa.Text = Format(NulosN(TxtImpTasa.Text), FORMAT_MONTO)
End Sub

Private Sub TxtIntervalos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumLet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtPortes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtPortes_Validate(Cancel As Boolean)
    If NulosN(TxtPortes.Text) = 0 Then
        If NulosN(TxtImpTasa.Text) <> 0 Then Exit Sub
        LblNotDeb.Visible = False
        TxtNotDeb.Visible = False
        TxtNotDeb.Text = ""
        Exit Sub
    End If
    
    LblNotDeb.Visible = True
    TxtNotDeb.Visible = True
    TxtNotDeb.Text = "0001" & "-" & HallaNumdocVenta(8, "0001", xCon)
    
    TxtPortes.Text = Format(NulosN(TxtPortes.Text), FORMAT_MONTO)
End Sub

Private Sub TxtRucPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtRucPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub TxtRucPro_Validate(Cancel As Boolean)
    If NulosC(TxtRucPro.Text) = "" Then
        LblProveedor.Caption = ""
        LblIdCliente.Caption = ""
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    RST_Busq xRs1, "SELECT * FROM mae_cliente WHERE numruc like '" & NulosC(TxtRucPro.Text) & "%' ORDER BY numruc", xCon
    If xRs1.RecordCount <> 0 Then
        TxtRucPro.Text = NulosC(xRs1("numruc"))
        LblProveedor.Caption = NulosC(xRs1("nombre"))
        LblIdCliente.Caption = xRs1("id")
        '--evaluar si es agente de retencion
        If NulosN(xRs1("ageret")) = -1 Then
            FraAgenRet.Visible = True
            
        Else
            FraAgenRet.Visible = False
            ChkRetencion.value = 0
        End If
        '------------------------------------
        
    Else
        TxtRucPro.Text = ""
        LblProveedor.Caption = ""
        LblIdCliente.Caption = ""
    End If
    
    
    Set xRs1 = Nothing
End Sub

Private Sub TxtTipInt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtTipInt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipo_Click
    End If
End Sub

Private Sub TxtTipInt_Validate(Cancel As Boolean)
    If NulosN(TxtTipInt.Text) = 0 Then
        LblTipoInteres.Caption = ""
        Exit Sub
    End If
    
    LblTipoInteres.Caption = Busca_Codigo(NulosN(TxtTipInt.Text), "id", "descripcion", "let_letratipoplazo", "N", xCon)
    
    If NulosC(LblTipoInteres.Caption) = "" Then
        TxtTipInt.Text = ""
    End If
End Sub

Sub GenerarLetras(TipInteres As Integer, Tasa As Double, NunDias As Integer, NumLetras As Integer, Capital As Double, NumeroLetraInicial As Integer, Portes As Double, FchInicio As String, Intervalos As Integer)
    Fg2.Rows = 1
    Dim Factor As Double
    Dim ImporteLetra As Double
    Dim NumeroLetra As Integer
    Dim A As Integer
    Dim ValorRedondea As Double
    
    If TipInteres = 0 Then Exit Sub
    If Tasa = 0 Then Exit Sub
    If NunDias = 0 Then Exit Sub
    If NumLetras = 0 Then Exit Sub
    If Capital = 0 Then Exit Sub
    If NumeroLetraInicial = 0 Then Exit Sub
    If Portes = 0 Then Exit Sub
    
    Agregando = True
    TxtSCapital.Text = ""
    TxtSInteres.Text = ""
    TxtSPortes.Text = ""
    TxtSIGV.Text = ""
    TxtSTotal.Text = ""
    
    TxtDCapital.Text = ""
    TxtDInteres.Text = ""
    TxtDPortes.Text = ""
    TxtDIGV.Text = ""
    TxtDTotal.Text = ""
    
    ImporteLetra = (Capital / NumLetras)
    ImporteLetra = Int(ImporteLetra * 100) / 100
    NumeroLetra = NumeroLetraInicial
    Dim ValorRedondea2 As Integer
    ValorRedondea = Format((Capital - (ImporteLetra * NumLetras)), "0.00")
    
    If Fg2.Rows = 1 Then
        For A = 1 To NumLetras
            Fg2.Rows = Fg2.Rows + 1
            Factor = HallarFactor(Tasa, NunDias, TipInteres)
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = Format(A, "00")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = Format(NumeroLetra, "00000000")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(NunDias, "00")
            If A = 10 Then
                ImporteLetra = ImporteLetra + ValorRedondea
            End If
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(ImporteLetra, FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(ImporteLetra * Factor, FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(Portes, FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 7) = ((NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 5)) + NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 6))) * 0.19)
            Fg2.TextMatrix(Fg2.Rows - 1, 7) = Format(Fg2.TextMatrix(Fg2.Rows - 1, 7), FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 8) = NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 4)) + NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 5)) + NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 6)) + NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 7))
            Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(Fg2.TextMatrix(Fg2.Rows - 1, 8), FORMAT_MONTO)
            
            If NulosC(FchInicio) <> "" Then
                Fg2.TextMatrix(Fg2.Rows - 1, 9) = CDate(FchInicio) + NunDias
                Fg2.TextMatrix(Fg2.Rows - 1, 9) = Format(Fg2.TextMatrix(Fg2.Rows - 1, 9), "dd/mm/yy")
            End If
            
            NumeroLetra = NumeroLetra + 1
            NunDias = NunDias + Intervalos
        Next A
        
        SumarColumnas
    End If
    Agregando = False
End Sub

Public Function TRUNC(ByVal value As Double, ByVal escala As Integer) As Double
    TRUNC = Int(value * 10 ^ escala)
End Function

Sub SumarColumnas()
    Dim xTotal1, xTotal2, xTotal3, xTotal4, xTotal5 As Double
    
    xTotal1 = GRID_SUMAR_COL(Fg2, 4)
    xTotal2 = GRID_SUMAR_COL(Fg2, 5)
    xTotal3 = GRID_SUMAR_COL(Fg2, 6)
    xTotal4 = GRID_SUMAR_COL(Fg2, 7)
    xTotal5 = GRID_SUMAR_COL(Fg2, 8)
    
    If TxtIdMon.Text = 1 Then
        TxtSCapital.Text = Format(xTotal1, FORMAT_MONTO)
        TxtSInteres.Text = Format(xTotal2, FORMAT_MONTO)
        TxtSPortes.Text = Format(xTotal3, FORMAT_MONTO)
        TxtSIGV.Text = Format(xTotal4, FORMAT_MONTO)
        TxtSTotal.Text = Format(xTotal5, FORMAT_MONTO)
    
        If NulosN(LblTipoCambio.Caption) <> 0 Then
            TxtDCapital.Text = Format(xTotal1 / NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            TxtDInteres.Text = Format(xTotal2 / NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            TxtDPortes.Text = Format(xTotal3 / NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            TxtDIGV.Text = Format(xTotal4 / NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            TxtDTotal.Text = Format(xTotal5 / NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
        End If
    Else
        TxtDCapital.Text = Format(xTotal1, FORMAT_MONTO)
        TxtDInteres.Text = Format(xTotal2, FORMAT_MONTO)
        TxtDPortes.Text = Format(xTotal3, FORMAT_MONTO)
        TxtDIGV.Text = Format(xTotal4, FORMAT_MONTO)
        TxtDTotal.Text = Format(xTotal5, FORMAT_MONTO)
    
        If NulosN(LblTipoCambio.Caption) <> 0 Then
            TxtSCapital.Text = Format(xTotal1 * NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            TxtSInteres.Text = Format(xTotal2 * NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            TxtSPortes.Text = Format(xTotal3 * NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            TxtSIGV.Text = Format(xTotal4 * NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
            TxtSTotal.Text = Format(xTotal5 * NulosN(LblTipoCambio.Caption), FORMAT_MONTO)
        End If
    End If
End Sub

Function HallarFactor(Tasa As Double, NumDias As Integer, TipInteres As Integer) As Double
    If TipInteres = 2 Then
        HallarFactor = (((Tasa / 100) + 1) ^ ((NumDias / 360)) - 1)
    Else
        MsgBox "No procede mensual", xTitulo
        'Factor = ((((Tasa / 100) + 1) ^ ((numdias / 360)) - 1) * ImporteLetra)
        HallarFactor = 0
    End If
End Function

Function Grabar() As Boolean

On Error GoTo LaCague
    
    Dim xCampos(31, 4) As String
    Dim xCampos2(11, 4) As String
    Dim xCampos3(3, 4) As String
    
    Dim xId, xIdNotDebito As Double
    Dim xTipLet, A As Integer
    Dim FchReg As String
    
    Dim xNumAsiento As String
    
    
    xCon.BeginTrans
    
    'ESPECIFICAMOS EL ID DEL MOVIMIENTO
    If QueHace = 1 Then
        xId = HallaCodigoTabla("let_letra", xCon, "id")
    Else
        xId = RstLet("id")
        '--ELIMINAMOS LOS REGISTROS PARA VOLVER A ESCRIBIRLOS
        xCon.Execute "delete from con_letradoc where idlet = " & xId
        xCon.Execute "delete from con_letradet where idlet = " & xId
        xCon.Execute "delete from let_letra where id = " & xId
    End If
    
    'ESPECIFICAMOS LA FECHA DE REGISTRO
    If mMesActivo = 0 Then
        FchReg = "01/01/" & Format(AnoTra, "0000")
    Else
        FchReg = "01/" & Format(mMesActivo, "00") & "/" & Format(AnoTra, "0000")
    End If
    
    mIdRegistro = xId
        
    'ESPECIFICAMOS EL ID DE LA NOTA DE DEBITO
    xIdNotDebito = 0
    
    'ESPECIFICAMOS EL TIPO DE LETRA
    xTipLet = 1 ' Clientes
    'xTipLet = 2 ' Proveedores
    
    'Columna    | Descripcion
    '------------------------
    '0          | campo
    '1          | Valor
    '2          | requerido
    '3          | tipo
    '4          | msj error
    
    '--------------------------------
    'GRABAMOS LA CABECERA DE LA LETRA
    xCampos(0, 0) = "id":           xCampos(0, 1) = Str(xId):                  xCampos(0, 2) = "S":    xCampos(0, 3) = "N":     xCampos(0, 4) = "":
    xCampos(1, 0) = "idlib":        xCampos(1, 1) = "37":                       xCampos(1, 2) = "S":    xCampos(1, 3) = "N":     xCampos(1, 4) = ""
    xCampos(2, 0) = "tipdoc":       xCampos(2, 1) = "95":                       xCampos(2, 2) = "S":    xCampos(2, 3) = "N":     xCampos(2, 4) = ""
    xCampos(3, 0) = "ano":          xCampos(3, 1) = Str(AnoTra):                xCampos(3, 2) = "S":    xCampos(3, 3) = "N":     xCampos(3, 4) = ""
    xCampos(4, 0) = "idmes":        xCampos(4, 1) = Str(mMesActivo):            xCampos(4, 2) = "S":    xCampos(4, 3) = "N":     xCampos(4, 4) = ""
    xCampos(5, 0) = "fchemi":       xCampos(5, 1) = TxtFchEmi.Valor:            xCampos(5, 2) = "S":    xCampos(5, 3) = "F":     xCampos(5, 4) = "No ha especificado la fecha de emision para las letras"
    xCampos(6, 0) = "fchini":       xCampos(6, 1) = TxtFchIni.Valor:            xCampos(6, 2) = "S":    xCampos(6, 3) = "F":     xCampos(6, 4) = "No ha especificado la fecha de inicio para las letras"
    xCampos(7, 0) = "tiplet":       xCampos(7, 1) = Str(xTipLet):               xCampos(7, 2) = "S":    xCampos(7, 3) = "N":     xCampos(7, 4) = ""
    xCampos(8, 0) = "idclipro":     xCampos(8, 1) = Str(LblIdCliente.Caption):  xCampos(8, 2) = "S":    xCampos(8, 3) = "N":     xCampos(8, 4) = "No ha especificado el cliente o proveedor para emitir la letra"
    xCampos(9, 0) = "numlet":       xCampos(9, 1) = TxtNumLet.Text:             xCampos(9, 2) = "S":    xCampos(9, 3) = "N":     xCampos(9, 4) = "No ha especificado el numero de letras a generar"
    xCampos(10, 0) = "idmon":       xCampos(10, 1) = TxtIdMon.Text:             xCampos(10, 2) = "S":   xCampos(10, 3) = "N":    xCampos(10, 4) = "No ha especificado la moneda"
    xCampos(11, 0) = "impcap":      xCampos(11, 1) = TxtImpFinan.Text:          xCampos(11, 2) = "S":   xCampos(11, 3) = "N":    xCampos(11, 4) = "No ha especificado el capital a financiar"
    xCampos(12, 0) = "idtipdocref": xCampos(12, 1) = TxtIdTipDocRef.Text:       xCampos(12, 2) = "N":   xCampos(12, 3) = "N":    xCampos(12, 4) = "":
    xCampos(13, 0) = "iddocref2":   xCampos(13, 1) = LblIdDocRef.Caption:       xCampos(13, 2) = "N":   xCampos(13, 3) = "N":    xCampos(13, 4) = ""
    xCampos(14, 0) = "tipint":      xCampos(14, 1) = TxtTipInt.Text:            xCampos(14, 2) = "S":   xCampos(14, 3) = "N":    xCampos(14, 4) = "No ha especificado el tipo de interes que se aplicara"
    xCampos(15, 0) = "inttasa":     xCampos(15, 1) = TxtImpTasa.Text:           xCampos(15, 2) = "S":   xCampos(15, 3) = "N":    xCampos(15, 4) = "No ha especificado la tasa de interes que se aplicara"
    xCampos(16, 0) = "glosa":       xCampos(16, 1) = "":                        xCampos(16, 2) = "N":   xCampos(16, 3) = "C":    xCampos(16, 4) = ""
    xCampos(17, 0) = "fchreg":      xCampos(17, 1) = FchReg:                    xCampos(17, 2) = "S":   xCampos(17, 3) = "F":    xCampos(17, 4) = ""
    xCampos(18, 0) = "numdias":     xCampos(18, 1) = TxtDiasPlazo.Text:         xCampos(18, 2) = "S":   xCampos(18, 3) = "N":    xCampos(18, 4) = "No ha especificado a cuantos dias se financiara la letra"
    xCampos(19, 0) = "girado":      xCampos(19, 1) = TxtGirado.Text:            xCampos(19, 2) = "N":   xCampos(19, 3) = "C":    xCampos(19, 4) = "No ha especificado a nombre de quien se gira la letra"
    xCampos(20, 0) = "iddocgir":    xCampos(20, 1) = TxtIdDocIden.Text:         xCampos(20, 2) = "N":   xCampos(20, 3) = "N":    xCampos(20, 4) = "No ha especificado el tipo de documento de identidad de la persona a quien se gira la letra"
    xCampos(21, 0) = "numdocgir":   xCampos(21, 1) = TxtNumDoc.Text:            xCampos(21, 2) = "N":   xCampos(21, 3) = "C":    xCampos(21, 4) = "No ha especificado el Nº de documento de identidad de la persona a quien se gira la letra"
    xCampos(22, 0) = "dir":         xCampos(22, 1) = "":                        xCampos(22, 2) = "N":   xCampos(22, 3) = "C":    xCampos(22, 4) = ""
    xCampos(23, 0) = "idnotdeb":    xCampos(23, 1) = Str(xIdNotDebito):         xCampos(23, 2) = "N":   xCampos(23, 3) = "N":    xCampos(23, 4) = ""
    xCampos(24, 0) = "diaint":      xCampos(24, 1) = TxtIntervalos.Text:        xCampos(24, 2) = "S":   xCampos(24, 3) = "N":    xCampos(24, 4) = "No ha especificado el intervalo de dias para emitir las letras"
    xCampos(25, 0) = "imppor":      xCampos(25, 1) = TxtPortes.Text:            xCampos(25, 2) = "N":   xCampos(25, 3) = "N":    xCampos(25, 4) = ""
    xCampos(26, 0) = "aplicaret":   xCampos(26, 1) = ChkRetencion.value:        xCampos(26, 2) = "N":   xCampos(26, 3) = "N":    xCampos(26, 4) = ""
    
    xCampos(27, 0) = "numorden":    xCampos(27, 1) = Mid(TxtNumDocRef2.Text, 10, 6): xCampos(27, 2) = "N":    xCampos(27, 3) = "C":    xCampos(27, 4) = "No ha especificado el numero de la orden de despacho"
    xCampos(28, 0) = "anoorden":    xCampos(28, 1) = Mid(TxtNumDocRef2.Text, 6, 4):  xCampos(28, 2) = "N":    xCampos(28, 3) = "N":    xCampos(28, 4) = "No ha especificado el año de la orden de despacho"
    xCampos(29, 0) = "idaduana":    xCampos(29, 1) = Mid(TxtNumDocRef2.Text, 1, 3):  xCampos(29, 2) = "N":    xCampos(29, 3) = "N":    xCampos(29, 4) = "No ha especificado la aduana de la orden de despacho"
    xCampos(30, 0) = "idregimen":   xCampos(30, 1) = Mid(TxtNumDocRef2.Text, 4, 2):  xCampos(30, 2) = "N":    xCampos(30, 3) = "N":    xCampos(30, 4) = "No ha especificado el regimen de la orden de despacho"
    
    xCampos(31, 0) = "numerodocref":   xCampos(31, 1) = TxtNumDocRef2.Text:     xCampos(31, 2) = "N":   xCampos(31, 3) = "C":    xCampos(31, 4) = "No ha especificado la orden de despacho"
    
    
    If EscribirNuevoRegistro(xCampos, "let_letra", xCon) = False Then
        xCon.RollbackTrans
        Exit Function
    End If
    
    Dim SeGrabo As Boolean
    
    '---------------------------------
    'GRABAMOS EL DETALLE DE LAS LETRAS
    For A = 1 To Fg2.Rows - 1
        xCampos2(0, 0) = "idlet":       xCampos2(0, 1) = Str(xId):           xCampos2(0, 2) = "S":    xCampos2(0, 3) = "N":    xCampos2(0, 4) = "":
        xCampos2(1, 0) = "corr":        xCampos2(1, 1) = Str(A):                   xCampos2(1, 2) = "S":    xCampos2(1, 3) = "N":    xCampos2(1, 4) = ""
        xCampos2(2, 0) = "numser":      xCampos2(2, 1) = "00":                     xCampos2(2, 2) = "S":    xCampos2(2, 3) = "C":    xCampos2(2, 4) = ""
        xCampos2(3, 0) = "numdoc":      xCampos2(3, 1) = Fg2.TextMatrix(A, 2):     xCampos2(3, 2) = "S":    xCampos2(3, 3) = "C":    xCampos2(3, 4) = ""
        xCampos2(4, 0) = "fchemi":      xCampos2(4, 1) = TxtFchIni.Valor:          xCampos2(4, 2) = "S":    xCampos2(4, 3) = "F":    xCampos2(4, 4) = ""
        xCampos2(5, 0) = "fchven":      xCampos2(5, 1) = Fg2.TextMatrix(A, 9):     xCampos2(5, 2) = "S":    xCampos2(5, 3) = "F":    xCampos2(5, 4) = ""
        xCampos2(6, 0) = "impcapital":  xCampos2(6, 1) = Fg2.TextMatrix(A, 4):     xCampos2(6, 2) = "S":    xCampos2(6, 3) = "N":    xCampos2(6, 4) = ""
        xCampos2(7, 0) = "impporte":    xCampos2(7, 1) = Fg2.TextMatrix(A, 6):     xCampos2(7, 2) = "S":    xCampos2(7, 3) = "N":    xCampos2(7, 4) = ""
        xCampos2(8, 0) = "impinteres":  xCampos2(8, 1) = Fg2.TextMatrix(A, 5):     xCampos2(8, 2) = "S":    xCampos2(8, 3) = "N":    xCampos2(8, 4) = ""
        xCampos2(9, 0) = "impigv":      xCampos2(9, 1) = Fg2.TextMatrix(A, 7):     xCampos2(9, 2) = "S":    xCampos2(9, 3) = "N":    xCampos2(9, 4) = ""
        xCampos2(10, 0) = "implet":     xCampos2(10, 1) = Fg2.TextMatrix(A, 8):    xCampos2(10, 2) = "S":   xCampos2(10, 3) = "N":   xCampos2(10, 4) = ""
        xCampos2(11, 0) = "diasplazo":  xCampos2(11, 1) = Fg2.TextMatrix(A, 3):    xCampos2(11, 2) = "S":   xCampos2(11, 3) = "N":   xCampos2(11, 4) = ""
        
        
        If EscribirNuevoRegistro(xCampos2, "let_letradet", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Next A
        
    'GRABAMOS LOS DOCUMENTOS QUE ORIGINAN LA LETRA
    For A = 1 To Fg1.Rows - 1
        xCampos3(0, 0) = "idlet":      xCampos3(0, 1) = Str(xId):           xCampos3(0, 2) = "S":    xCampos3(0, 3) = "N":    xCampos3(0, 4) = "":
        xCampos3(1, 0) = "idmod":      xCampos3(1, 1) = Fg1.TextMatrix(A, 10):    xCampos3(1, 2) = "S":    xCampos3(1, 3) = "N":    xCampos3(1, 4) = ""
        xCampos3(2, 0) = "iddoc":      xCampos3(2, 1) = Fg1.TextMatrix(A, 11):    xCampos3(2, 2) = "S":    xCampos3(2, 3) = "C":    xCampos3(2, 4) = ""
        If NulosN(Fg1.TextMatrix(A, 12)) = 1 Then
            xCampos3(3, 0) = "impfin":     xCampos3(3, 1) = Fg1.TextMatrix(A, 9):     xCampos3(3, 2) = "S":    xCampos3(3, 3) = "N":    xCampos3(3, 4) = ""
        Else
            xCampos3(3, 0) = "impfin":     xCampos3(3, 1) = Fg1.TextMatrix(A, 8):     xCampos3(3, 2) = "S":    xCampos3(3, 3) = "N":    xCampos3(3, 4) = ""
        End If
        
        If EscribirNuevoRegistro(xCampos3, "let_letradoc", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
        
    Next A
    
    
    '--generamos es asiento
    xNumAsiento = GenerarAsiento(xCon, 37, CDbl(xId), AnoTra, mMesActivo, 1, 0)
    If xNumAsiento = "" Then GoTo LaCague
    
    '---------------------------------------------------------------------------
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, 19, QueHace, xHorIni, Time, Date, xCon, CDbl(xId)
    
    MsgBox "El movimiento se grabó con éxito" & vbCr & "Registro Nº: " & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    
    
    MsgBox "La letra se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Grabar = True
    
    xCon.CommitTrans
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar la letra por el siguiente motivo : " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Err.Clear
    
    Grabar = False
End Function

Sub Cancelar()
    ActivarTool
    Bloquea
    TabOne1.TabEnabled(0) = True
    Label5.Caption = "Detalle de Letra"
    TabOne1.CurrTab = 0
    QueHace = 3
End Sub



Private Sub OpcionesPeriodo()
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    LblMes1.Caption = LblMes.Caption
    
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, 19, mMesActivo, fCierrePeriodo, xCon
    '------------------------------------------------------------------------------------------
End Sub
