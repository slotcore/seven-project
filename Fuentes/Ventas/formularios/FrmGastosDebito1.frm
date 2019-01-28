VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGastosDebito1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas - Gastos Débito y Crédito"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdApertura 
      Caption         =   "Agregar Apertura"
      Height          =   315
      Left            =   10200
      TabIndex        =   95
      Top             =   390
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   2025
      Left            =   12180
      TabIndex        =   53
      Top             =   90
      Visible         =   0   'False
      Width           =   6705
      Begin VB.TextBox TxtNewSaldo2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1320
         TabIndex        =   65
         Text            =   "TxtNewSaldo2"
         Top             =   1515
         Width           =   1395
      End
      Begin VB.TextBox TxtSaldo2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1320
         TabIndex        =   63
         Text            =   "TxtSaldo2"
         Top             =   1200
         Width           =   1395
      End
      Begin VB.TextBox TxtCliente2 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1320
         TabIndex        =   61
         Text            =   "TxtCliente2"
         Top             =   780
         Width           =   5280
      End
      Begin VB.TextBox TxtNumDoc2 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1320
         TabIndex        =   57
         Text            =   "TxtNumDoc2"
         Top             =   465
         Width           =   2055
      End
      Begin VB.Frame Frame9 
         Height          =   870
         Left            =   3240
         TabIndex        =   67
         Top             =   1050
         Width           =   3375
         Begin VB.CommandButton Command2 
            Height          =   630
            Left            =   1710
            Picture         =   "FrmGastosDebito1.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   180
            Width           =   750
         End
         Begin VB.CommandButton Command1 
            Height          =   630
            Left            =   930
            Picture         =   "FrmGastosDebito1.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   68
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
         TabIndex        =   66
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   64
         Top             =   1245
         Width           =   405
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   62
         Top             =   825
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   59
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
         TabIndex        =   55
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
      Height          =   2280
      Left            =   12150
      TabIndex        =   37
      Top             =   2280
      Visible         =   0   'False
      Width           =   5565
      Begin VB.CommandButton CmdBusAlmacen2 
         Height          =   240
         Left            =   3180
         Picture         =   "FrmGastosDebito1.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   525
         Width           =   240
      End
      Begin VB.TextBox TxtAlmacen2 
         Height          =   300
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "TxtAlmacen2"
         Top             =   495
         Width           =   2025
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmiAnul 
         Height          =   300
         Left            =   1425
         TabIndex        =   52
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
         TabIndex        =   56
         Text            =   "TxtNumDocGen"
         Top             =   1755
         Width           =   1335
      End
      Begin VB.CommandButton CmdBusSerGen 
         Height          =   240
         Left            =   2490
         Picture         =   "FrmGastosDebito1.frx":0746
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1485
         Width           =   240
      End
      Begin VB.TextBox TxtNumSerGen 
         Height          =   300
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "TxtNumSerGen"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton CmdBusTipDocGen 
         Height          =   240
         Left            =   5160
         Picture         =   "FrmGastosDebito1.frx":0878
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   840
         Width           =   240
      End
      Begin VB.TextBox TxtIdDocGen 
         Height          =   300
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "TxtIdDocGen"
         Top             =   810
         Width           =   4005
      End
      Begin VB.Frame Frame7 
         Height          =   1020
         Left            =   3030
         TabIndex        =   46
         Top             =   1065
         Width           =   2400
         Begin VB.CommandButton cmdsalirseldoc 
            Height          =   600
            Left            =   1200
            Picture         =   "FrmGastosDebito1.frx":09AA
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   270
            Width           =   720
         End
         Begin VB.CommandButton cmdokseldoc 
            Height          =   600
            Left            =   450
            Picture         =   "FrmGastosDebito1.frx":0CB4
            Style           =   1  'Graphical
            TabIndex        =   58
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
         TabIndex        =   51
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
         TabIndex        =   47
         Top             =   510
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Documento"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   45
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   165
         TabIndex        =   44
         Top             =   1785
         Width           =   1050
      End
      Begin VB.Label LblIdDocumentoGen 
         AutoSize        =   -1  'True
         Caption         =   "LblIdDocumentoGen"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3765
         TabIndex        =   43
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   38
         Top             =   840
         Width           =   1185
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7470
      Left            =   -15
      TabIndex        =   13
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
         TabIndex        =   19
         Top             =   375
         Width           =   11805
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   345
            Left            =   9780
            TabIndex        =   94
            Top             =   2070
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.CheckBox ChkTC 
            Caption         =   "Check2"
            Enabled         =   0   'False
            Height          =   195
            Left            =   10425
            TabIndex        =   89
            Top             =   870
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
            Left            =   10665
            TabIndex        =   88
            Text            =   "TxtTC"
            Top             =   810
            Width           =   1065
         End
         Begin VB.Frame Frame4 
            Height          =   765
            Left            =   60
            TabIndex        =   20
            Top             =   6345
            Width           =   11700
            Begin VB.TextBox txtimpAcuenta 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080FFFF&
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
               Left            =   9150
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   91
               TabStop         =   0   'False
               Text            =   "txtimpAcue"
               Top             =   390
               Width           =   1200
            End
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "&Eliminar Item"
               Enabled         =   0   'False
               Height          =   495
               Left            =   1235
               TabIndex        =   12
               Top             =   165
               Width           =   1170
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "&Agregar Item"
               Enabled         =   0   'False
               Height          =   495
               Left            =   30
               TabIndex        =   10
               Top             =   165
               Width           =   1170
            End
            Begin VB.TextBox TxtimpMN 
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
               Left            =   6540
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   23
               TabStop         =   0   'False
               Text            =   "TxtImpMN"
               Top             =   390
               Width           =   1215
            End
            Begin VB.TextBox txtimpSaldo 
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
               Left            =   10410
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   22
               TabStop         =   0   'False
               Text            =   "txtimpSald"
               Top             =   390
               Width           =   1200
            End
            Begin VB.TextBox TxtImpME 
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
               Left            =   7860
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   21
               TabStop         =   0   'False
               Text            =   "TxtImpME"
               Top             =   390
               Width           =   1200
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imp. ACuenta"
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
               Left            =   9180
               TabIndex        =   92
               Top             =   150
               Width           =   1155
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   2460
               X2              =   2460
               Y1              =   135
               Y2              =   855
            End
            Begin VB.Line Line3 
               BorderColor     =   &H80000003&
               Index           =   0
               X1              =   6450
               X2              =   6450
               Y1              =   210
               Y2              =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Total MN"
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
               Left            =   6540
               TabIndex        =   26
               Top             =   150
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Saldo"
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
               Left            =   10440
               TabIndex        =   25
               Top             =   150
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Total ME"
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
               Left            =   7890
               TabIndex        =   24
               Top             =   150
               Width           =   1200
            End
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   5
            Text            =   "TxtNumSer"
            Top             =   1425
            Width           =   915
         End
         Begin VB.TextBox txtglosa 
            Height          =   315
            Left            =   1575
            TabIndex        =   9
            Text            =   "TxtGlosa"
            Top             =   2085
            Width           =   8085
         End
         Begin VB.CommandButton CmdBusAlm 
            Height          =   240
            Left            =   6720
            Picture         =   "FrmGastosDebito1.frx":0FBE
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   825
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.TextBox TxtIdAlm 
            Height          =   300
            Left            =   6270
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "TxtIdAlm"
            Top             =   795
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.CommandButton CmdBusDocRef2 
            Height          =   240
            Left            =   9405
            Picture         =   "FrmGastosDebito1.frx":10F0
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   1785
            Width           =   240
         End
         Begin VB.TextBox TxtNumDocRef 
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   8
            Text            =   "TxtNumDocRef"
            Top             =   1755
            Width           =   3390
         End
         Begin VB.CommandButton CmdBusIdTipDocRef 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmGastosDebito1.frx":1222
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   1785
            Width           =   240
         End
         Begin VB.Frame Frame10 
            Height          =   465
            Left            =   9630
            TabIndex        =   70
            Top             =   1155
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
               TabIndex        =   71
               Top             =   120
               Width           =   1860
            End
         End
         Begin VB.CommandButton CmdBusNumSer 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmGastosDebito1.frx":1354
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1455
            Width           =   240
         End
         Begin VB.CommandButton CmdBusCli 
            Height          =   240
            Left            =   3075
            Picture         =   "FrmGastosDebito1.frx":1486
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1140
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmGastosDebito1.frx":15B8
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   825
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2700
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "TxtNumDoc"
            Top             =   1425
            Width           =   1545
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   6720
            Picture         =   "FrmGastosDebito1.frx":16EA
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   480
            Width           =   240
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   6285
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "TxtIdMon"
            Top             =   450
            Width           =   705
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   3870
            Left            =   45
            TabIndex        =   11
            Top             =   2430
            Width           =   11700
            _cx             =   20637
            _cy             =   6826
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
            Cols            =   24
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmGastosDebito1.frx":181C
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
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   4
            Text            =   "TxtNumRuc"
            Top             =   1110
            Width           =   1770
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   2
            Text            =   "TxtTipDoc"
            Top             =   795
            Width           =   915
         End
         Begin VB.TextBox TxtIdTipDoc 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   7
            Text            =   "TxtIdTipDoc"
            Top             =   1755
            Width           =   915
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "T.C."
            Height          =   195
            Left            =   10110
            TabIndex        =   90
            Top             =   855
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Doc. Ref."
            Height          =   195
            Index           =   8
            Left            =   105
            TabIndex        =   87
            ToolTipText     =   "Tipo de Documento de Referencia"
            Top             =   1845
            Width           =   1005
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión"
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   86
            Top             =   525
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   7
            Left            =   105
            TabIndex        =   85
            Top             =   1185
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Documento"
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   84
            Top             =   855
            Width           =   1410
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   83
            Top             =   1515
            Width           =   1275
         End
         Begin VB.Label lblglosa 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Left            =   105
            TabIndex        =   82
            Top             =   2175
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   11
            Left            =   5580
            TabIndex        =   80
            Top             =   855
            Visible         =   0   'False
            Width           =   615
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
            Left            =   6990
            TabIndex        =   79
            Top             =   795
            Visible         =   0   'False
            Width           =   2655
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
            TabIndex        =   77
            Top             =   30
            Width           =   2190
         End
         Begin VB.Label LblIdDocRef2 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDocRef2"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   9735
            TabIndex        =   76
            Top             =   1800
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Ref."
            Height          =   195
            Index           =   9
            Left            =   5505
            TabIndex        =   74
            ToolTipText     =   "Documento de Referencia"
            Top             =   1845
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
            Left            =   2490
            TabIndex        =   73
            Top             =   1755
            Width           =   2655
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Gastos de Débito y Crédito"
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
            TabIndex        =   36
            Top             =   45
            Width           =   11595
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "T.C."
            Height          =   195
            Left            =   10155
            TabIndex        =   35
            Top             =   495
            Visible         =   0   'False
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
            Left            =   10530
            TabIndex        =   34
            Top             =   450
            Visible         =   0   'False
            Width           =   1215
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
            Left            =   6990
            TabIndex        =   15
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
            TabIndex        =   33
            Top             =   1110
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
            Left            =   2490
            TabIndex        =   32
            Top             =   795
            Width           =   2715
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2550
            Top             =   1530
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   5610
            TabIndex        =   14
            Top             =   525
            Width           =   585
         End
         Begin VB.Label LblIdCliente 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCliente"
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   2865
            TabIndex        =   31
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
         TabIndex        =   16
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6645
            Left            =   0
            TabIndex        =   81
            Top             =   360
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11721
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
            Columns(1).DataField=   "numreg1"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "TD"
            Columns(2).DataField=   "abrev"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nº Documento"
            Columns(3).DataField=   "numerodoc"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Emi"
            Columns(4).DataField=   "fchdoc1"
            Columns(4).NumberFormat=   "Short Date"
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
            Columns(7).Caption=   "T.C."
            Columns(7).DataField=   "impven1"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Importe"
            Columns(8).DataField=   "imptot1"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Saldo"
            Columns(9).DataField=   "impsal1"
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
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1693"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1614"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=794"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=714"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2566"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2487"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1773"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1693"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=6615"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=6535"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=979"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=900"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=1005"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=926"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=1879"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1799"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=1852"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1773"
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
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
            _StyleDefs(76)  =   "Named:id=33:Normal"
            _StyleDefs(77)  =   ":id=33,.parent=0"
            _StyleDefs(78)  =   "Named:id=34:Heading"
            _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   ":id=34,.wraptext=-1"
            _StyleDefs(81)  =   "Named:id=35:Footing"
            _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(83)  =   "Named:id=36:Selected"
            _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=37:Caption"
            _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(87)  =   "Named:id=38:HighlightRow"
            _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=39:EvenRow"
            _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(91)  =   "Named:id=40:OddRow"
            _StyleDefs(92)  =   ":id=40,.parent=33"
            _StyleDefs(93)  =   "Named:id=41:RecordSelector"
            _StyleDefs(94)  =   ":id=41,.parent=34"
            _StyleDefs(95)  =   "Named:id=42:FilterBar"
            _StyleDefs(96)  =   ":id=42,.parent=33"
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
            TabIndex        =   17
            Top             =   30
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Gastos de Débito - Crédito"
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
            TabIndex        =   18
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
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":1AD5
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":2019
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":23AB
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":252F
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":2983
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":2A9B
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":2FDF
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":3523
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":3637
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":374B
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":3B9F
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":3D0B
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":4253
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGastosDebito1.frx":456D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Documento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Restaurar Documento"
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
                  Text            =   "Anular Documento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Documento"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Emitir Documento Anulado"
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
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "FrmGastosDebito1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstVent As New ADODB.Recordset
Dim QueHace As Integer

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
Dim mMesActivo As Integer '--indica el mes activo
Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO



Sub ActivarEntorno()
    TabOne1.Enabled = Not TabOne1.Enabled
    Toolbar1.Enabled = Not Toolbar1.Enabled
End Sub


Sub VisImportesMonedas(idmon As Integer, importe As Double)


    On Error GoTo erroraviso
    
    Dim tcventa As Double
                        
        With Fg1
            '--obtenemos el tc para hacer los calculos
            If NulosN(.TextMatrix(.Row, 9)) = 0 Then
                tcventa = NulosN(TxtTC.Text)
            Else
                tcventa = NulosN(.TextMatrix(.Row, 9))
            End If
                            
            If NulosN(TxtIdMon) = idmon Then
                If idmon = 2 Then
                    .TextMatrix(.Row, 10) = Format(importe * tcventa, FORMAT_MONTO)
                    .TextMatrix(.Row, 11) = Format(importe, FORMAT_MONTO)
                Else
                    .TextMatrix(.Row, 10) = Format(importe, FORMAT_MONTO)
                    If tcventa <> 0 Then .TextMatrix(.Row, 11) = Format(importe / tcventa, FORMAT_MONTO)
                End If
            Else
                If idmon = 2 Then
                    .TextMatrix(.Row, 10) = Format(importe * tcventa, FORMAT_MONTO)
                    .TextMatrix(.Row, 11) = Format(importe, FORMAT_MONTO)
                Else
                    .TextMatrix(.Row, 10) = Format(importe, FORMAT_MONTO)
                    If tcventa <> 0 Then .TextMatrix(.Row, 11) = Format(importe / tcventa, FORMAT_MONTO)
                End If
            End If
            
            '--acuenta
            .TextMatrix(.Row, 12) = importe
            '--saldo
            .TextMatrix(.Row, 13) = 0
        End With
    
    Exit Sub
erroraviso:


    Select Case Err.Number
    Case 11
        MsgBox "No se pudo guardar el registro no hay diferencia de cambio ó o monto mal ingresado", vbInformation, Me.Caption
    Case Else
        MsgBox "No se pudo guardar el registro por el siguiente motivo :" & Chr(13) & Trim(Err.Number) & Err.Description
    End Select


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
        
        '--eliminamos el registro del analisis de cta cte
        xCon.Execute "DELETE * FROM var_analisisctacte WHERE idlib = 41 AND idope = " & RstVent("id") & ""
        
        xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & RstVent("id") & " AND idlib = 41"
        
        xCon.Execute "DELETE * FROM vta_gastodebitodet WHERE idlgd = " & RstVent("id") & ""
        xCon.Execute "DELETE * FROM vta_gastodebito WHERE id = " & RstVent("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstVent("id") & " AND idform = " & IdMenuActivo
        
        
        MsgBox RstVent("nomdoc") & " se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstVent.Requery
        Dg1.Refresh
        If RstVent.RecordCount = 0 Then
            Rpta = MsgBox("No se han registrado movimientos en el periodo especificado, ¿ Desea agregar uno ahora ?", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then Nuevo
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
    
    Rpta = MsgBox("Esta seguro de restaurar el Documento Nº " + RstVent("numser") & "-" & RstVent("numdoc"), vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption)
    If Rpta = vbYes Then
        xCon.Execute "UPDATE vta_gastodebito SET vta_gastodebito.Anulado = 0, " _
            & " vta_gastodebito.idcli = 0  " _
            & " WHERE vta_gastodebito.id =" & RstVent("id") & ""
        
        xCon.Execute "DELETE * FROM vta_gastodebitodet WHERE vta_gastodebitodet.idlgd  =" & RstVent("id") & ""
        RstVent.Requery
        Dg1.Refresh
        MsgBox "El documento se restauró con éxito" & vbCr & "Puede proceder a modificar el Documento", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
    End If
End Sub


Sub Anular()
    Dim Rpta As Integer
    Dim A As Integer
    Dim xId As Double
    Dim Rst As New ADODB.Recordset
    Dim xNumAsiento As String
    
    xHorIni = Time
    
    Rpta = MsgBox("¿Esta seguro de anular " & RstVent("nomdoc") & " Nº " & RstVent("numser") & "-" & RstVent("numdoc") + "?", vbYesNo + vbDefaultButton1 + vbQuestion, Me.Caption)
    
    If Rpta = vbYes Then
        
        xId = RstVent("id")
        
        '--inciando transaccion
        xCon.BeginTrans
        
        '--actualizar el documento
        xCon.Execute "UPDATE vta_gastodebito  SET vta_gastodebito.idcli=0,vta_gastodebito.Anulado = -1, " _
            & " vta_gastodebito.imptot = 0,  vta_gastodebito.impsal = 0  " _
            & " WHERE vta_gastodebito.id = " & xId & " "
        
        '--eliminamos el detalle del documento
        xCon.Execute "DELETE * FROM vta_gastodebitodet WHERE vta_gastodebitodet.idlgd  = " & xId & ""
        
        '--generando el asiento contable
        xNumAsiento = GenerarAsiento(xCon, 41, xId, AnoTra, mMesActivo, 1, 0)
        If xNumAsiento = "" Then GoTo LaCague
        
        'Grabamos el movimiento en la tabla var_edicion
        GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, CDbl(xId)
        
        MsgBox RstVent("nomdoc") & " se anuló con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        
        '--grabando transaccion
        xCon.CommitTrans
        
        RstVent.Requery
        Dg1.Refresh
    End If
    
    Exit Sub
    
LaCague:
    xCon.RollbackTrans
    MsgBox Err.Description & vbCr & Err.Source & vbCr & Err.Number, vbCritical, xTitulo
    Err.Clear
    
End Sub

Sub Cancelar()
    Dim X As Integer
    Bloquea
    Fg1.ColComboList(1) = ""
    Label5.Caption = "Detalle de Gasto Debito - Credito"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
     
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
       
 
    swguiafact = 0
End Sub

Sub Nuevo()
    QueHace = 1
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Agregando Gastos "
    Fg1.ColComboList(1) = "0 Seleccion|1 Manual"
    Fg1.Editable = flexEDKbdMouse
    
    Fg1.SelectionMode = flexSelectionFree

    Fg1.Rows = 1
    
    TxtFchDoc.Valor = Format(Date, "dd/mm/yyyy")
        
    xHorIni = Time
    TxtFchDoc.SetFocus
End Sub

Sub Modificar()
    If RstVent.RecordCount = 0 Then
        MsgBox "No hay Registros para Modificar", vbInformation, Me.Caption
        Exit Sub
    End If
    If NulosC(RstVent("nombre")) = "ANULADO" Then
        MsgBox "El Documento de Venta esta Anulado" & vbCr & "No se Puede Modificar", vbInformation, Me.Caption
        Exit Sub
    End If
   
    QueHace = 2
    
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    MuestraSegundoTab
    Label5.Caption = "Modificando Gastos de Debito - Credito"
    
    
    Fg1.ColComboList(1) = "0 Seleccion|1 Manual"
    
    
    
    Fg1.SelectionMode = flexSelectionFree
    xHorIni = Time
    TxtFchDoc.SetFocus
End Sub

Sub MuestraSegundoTab()
    
    Dim Rst As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    
    If RstVent.RecordCount = 0 Then Exit Sub
    Blanquea
    lblReg.Caption = "Nº Reg. " & NulosC(RstVent("numreg1"))
    
    TxtTipDoc.Text = NulosN(RstVent("tipdoc"))
    TxtNumRuc.Text = NulosC(RstVent("numruc"))
    TxtNumSer.Text = NulosC(RstVent("numser"))
    TxtNumDoc.Text = NulosC(RstVent("numdoc"))
    
   'El id de almacen debe grabarse en la tabla vta_gastodebito
'''    TxtIdAlm = NulosN(RstVent("idalm"))
'''    TxtIdAlm_Validate False
            
    If IsDate(RstVent("fchemi")) = True Then TxtFchDoc.Valor = CDate(RstVent("fchemi"))
    

    TxtIdMon.Text = NulosN(RstVent("idmon"))
    
    
    TxtNumDocRef.Text = NulosC(RstVent("numerodocref"))
    
    
    If NulosN(RstVent("idtipdocref")) <> 0 Then
        
        TxtIdTipDoc.Text = NulosC(RstVent("idtipdocref"))
        LblDescTipDocRef.Caption = Busca_Codigo(NulosC(RstVent("idtipdocref")), "id", "descripcion", "mae_docreferencia", "N", xCon)
        
        If NulosN(RstVent("idtipdocref")) = 4 Then
            RST_Busq Rst, "SELECT var_ordendespacho.id, var_ordendespacho.numerodoc AS numdoc " _
                & " From var_ordendespacho WHERE (((var_ordendespacho.id)=" & NulosN(RstVent("iddocref2")) & "))", xCon
        End If
        
        If Rst.RecordCount <> 0 Then
            TxtNumDocRef.Text = Rst("numdoc")
            LblIdDocRef2.Caption = Rst("id")
        End If
        
        Set Rst = Nothing
        
        
    End If
    
    
    If RstVent("idmon") = 1 Then
        Me.TxtimpMN.Text = Format(NulosN(RstVent("imptot1")), FORMAT_MONTO)
    Else
        Me.TxtImpME.Text = Format(NulosN(RstVent("imptot1")), FORMAT_MONTO)
    End If
    
    txtglosa = NulosC(RstVent("glosa"))
        
    LblNomDoc.Caption = NulosC(RstVent("nomdoc"))
    LblNomCli.Caption = NulosC(RstVent("nombre"))

    TxtNumRuc.Text = NulosC(RstVent("numruc"))
    LblMoneda.Caption = NulosC(RstVent("descmon"))
    LblIdCliente.Caption = NulosN(RstVent("idcli"))
    

    '--tipo de cambio
    If NulosN(RstVent("tc")) = 0 Then
        ChkTC.Value = 0
        TxtTC.Text = NulosN(RstVent("impven1"))
        TxtTC.BackColor = &H8000000F
        TxtTC.Enabled = False
    Else
        ChkTC.Value = 1
        TxtTC.Text = NulosN(RstVent("tc"))
        TxtTC.BackColor = vbWhite
        TxtTC.Enabled = True
    End If
    If QueHace = 3 Then TxtTC.BackColor = &H8000000F
        
    
    'Detalle del Documento
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer

     
    'CARGAMOS LOS ITEMS DE LA FACTURA
    Set RstDet = Nothing
    Mostrando = True
    
    
    
    RST_Busq RstDet, " SELECT vta_gastodebitodet.*, mae_moneda.simbolo AS desmoneda, mae_documento.abrev, tes_modulos.descripcion, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanombre, " _
        & " IIf(vta_gastodebitodet.idtipper In (2,5),mae_cliente.numruc,mae_prov.numruc) AS numruc, IIf(vta_gastodebitodet.idtipper In (2,5),mae_cliente.nombre,mae_prov.nombre) AS nombre " _
        & " FROM (((((vta_gastodebitodet LEFT JOIN mae_moneda ON vta_gastodebitodet.idmon = mae_moneda.id) LEFT JOIN mae_documento ON vta_gastodebitodet.tipdoc = mae_documento.id) LEFT JOIN con_planctas ON vta_gastodebitodet.idcuen = con_planctas.id) LEFT JOIN mae_cliente ON vta_gastodebitodet.idper = mae_cliente.id) LEFT JOIN mae_prov ON vta_gastodebitodet.idper = mae_prov.id) LEFT JOIN tes_modulos ON vta_gastodebitodet.idmod = tes_modulos.id " _
        & " WHERE vta_gastodebitodet.idlgd =" & RstVent("id") & "", xCon
            
    If RstDet.RecordCount <> 0 Then
        Do While Not RstDet.EOF
            
             
             Me.Fg1.AddItem ""
             Fg1.Row = Fg1.Rows - 1
             
             
             'SI ES MOVIMIENTO ES POR SELECCION
             If RstDet!tipreg = 0 Then
                '--verificar si el registros es de apertura o proviene de otros modulos
                If RstDet!esapertura = 0 Then
                    Select Case RstDet!idmod
                    
                    Case 1 'Compras
                        RST_Busq xRs, " SELECT mae_prov.numruc, mae_prov.nombre, com_compras.fchdoc, mae_documento.abrev,  IIf([com_compras].[numser]='',[com_compras].[numdoc],[com_compras].[numser]+ '-' +[com_compras].[numdoc]) AS NroDoc, mae_moneda.simbolo as desmoneda, com_compras.imptot, com_compras.impsal, " _
                            & " com_compras.id, com_compras.tipdoc, com_compras.idmon, com_compras.idpro as idclipro, com_compras.numreg, " _
                            & " IIf([com_compras].[tc]=0,[con_tc].[impven],[com_compras].[tc]) & '' AS impven,com_compras.glosa,'No' as espaertura " _
                            & " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_prov RIGHT JOIN com_compras ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
                            & " WHERE  com_compras.id = " & NulosN(RstDet!iddoc) & " ORDER BY  IIf(com_compras.numser='',com_compras.numdoc,com_compras.numser+'-'+com_compras.numdoc) ", xCon
            
                    
                    Case 2 'Ventas
                    
                        RST_Busq xRs, " SELECT mae_cliente.numruc, mae_cliente.nombre, vta_ventas.fchdoc,mae_documento.abrev, vta_ventas.numser+ '-'+ vta_ventas.numdoc AS nrodoc, mae_moneda.simbolo as DesMoneda, vta_ventas.imptotdoc as imptot,vta_ventas.impsal, " _
                            & " vta_ventas.id, vta_ventas.tipdoc, vta_ventas.idmon, vta_ventas.idcli as idclipro, vta_ventas.numreg, " _
                            & " IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc]) & '' AS impven,vta_ventas.glosa,'No' as espaertura  " _
                            & " FROM ((mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_documento.id = vta_ventas.tipdoc) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
                            & " WHERE vta_ventas.id = " & NulosN(RstDet!iddoc) & " ORDER BY vta_ventas.numser + '-'+ vta_ventas.numdoc DESC", xCon
                    
                    
                    Case 9 'Honorarios
                            
                        RST_Busq xRs, " SELECT mae_prov.numruc, mae_prov.nombre, com_honorarios.fchdoc, mae_documento.abrev, [com_honorarios].[numser]+ '-' + [com_honorarios].[numdoc] AS nrodoc, mae_moneda.simbolo as Desmoneda , com_honorarios.imptot, com_honorarios.impsal, " _
                            & " com_honorarios.id, com_honorarios.tipdoc, com_honorarios.idmon, com_honorarios.idpro as idclipro, com_honorarios.numreg, " _
                            & " IIf([com_honorarios].[tc]=0,[con_tc].[impven],[com_honorarios].[tc]) & '' AS impvencom_honorarios.glosa,'No' as espaertura " _
                            & " FROM (((mae_documento RIGHT JOIN com_honorarios ON mae_documento.id = com_honorarios.tipdoc) LEFT JOIN mae_moneda ON com_honorarios.idmon = mae_moneda.id) LEFT JOIN mae_prov ON com_honorarios.idpro = mae_prov.id) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha " _
                            & " WHERE com_honorarios.id = " & NulosN(RstDet!iddoc) & " ORDER BY com_honorarios.numser + '-' + com_honorarios.numdoc ", xCon
                    
                    Case 10 'Reembolsable
                    
                        RST_Busq xRs, " SELECT mae_prov.numruc, mae_prov.nombre, com_reembolsables.fchdoc, mae_documento.abrev, [com_reembolsables]![numser]+ '-'+[com_reembolsables]![numdoc] AS nrodoc, mae_moneda.simbolo as desmoneda, com_reembolsables.imptot, com_reembolsables.impsal, " _
                            & " com_reembolsables.id, com_reembolsables.tipdoc, com_reembolsables.idmon, com_reembolsables.idpro as idclipro, '' AS numreg , " _
                            & " IIf([com_reembolsables].[tc]=0,[con_tc].[impven],[com_reembolsables].[tc]) & '' AS impven,com_reembolsables.glosa,'No' as espaertura " _
                            & " FROM (((com_reembolsables LEFT JOIN mae_documento ON com_reembolsables.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON com_reembolsables.idmon = mae_moneda.id) INNER JOIN mae_prov ON com_reembolsables.idpro = mae_prov.id) LEFT JOIN con_tc ON com_reembolsables.fchdoc = con_tc.fecha " _
                            & " WHERE com_reembolsables.id = " & NulosN(RstDet!iddoc) & " ORDER BY [com_reembolsables]![numser]+'-'+[com_reembolsables]![numdoc] ", xCon
                    Case 11 'Liquidacion
                    
                        RST_Busq xRs, " SELECT mae_cliente.numruc, mae_cliente.nombre, vta_gastodebito.fchemi AS fchdoc,mae_documento.abrev, [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc] AS nrodoc, mae_moneda.simbolo AS desmoneda, vta_gastodebito.imptot, vta_gastodebito.impsal,  " _
                            & " vta_gastodebito.id, vta_gastodebito.tipdoc, vta_gastodebito.idmon, vta_gastodebito.idcli AS idclipro, vta_gastodebito.numreg, " _
                            & " iif(vta_gastodebito.anulado=-1,0,IIf([vta_gastodebito].[tc]=0,[con_tc].[impven],[vta_gastodebito].[tc])) & '' AS impven, vta_gastodebito.glosa,'No' as espaertura " _
                            & " FROM (mae_cliente RIGHT JOIN ((vta_gastodebito LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id) ON mae_cliente.id = vta_gastodebito.idcli) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha " _
                            & " WHERE vta_gastodebito.id =" & NulosN(RstDet!iddoc) & "  ORDER BY [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc]", xCon
                    Case Else
                        
                    End Select
                Else
                '--si es apertura
                
                    RST_Busq xRs, "SELECT IIf(con_provicionesdetdoc.idtipper=1,mae_prov.numruc,IIf(con_provicionesdetdoc.idtipper=2,mae_cliente.numruc,IIf(con_provicionesdetdoc.idtipper=3,pla_empleados.numdoc,IIf(con_provicionesdetdoc.idtipper=5,mae_bancos.numruc,'')))) AS numruc, " _
                        & " IIf(con_provicionesdetdoc.idtipper=1,mae_prov.nombre,IIf(con_provicionesdetdoc.idtipper=2,mae_cliente.nombre,IIf(con_provicionesdetdoc.idtipper=3,pla_empleados.nombre,IIf(con_provicionesdetdoc.idtipper=5,mae_bancos.descripcion,'')))) AS nombre, " _
                        & " con_provicionesdetdoc.fchemi AS fchdoc, mae_documento.abrev, [con_provicionesdetdoc].[numser] & '-' & [con_provicionesdetdoc].[numdoc] AS nrodoc, mae_moneda.simbolo AS desmoneda, con_provicionesdetdoc.impdoc AS imptot, con_provicionesdetdoc.impsal, con_provicionesdetdoc.id, con_provicionesdetdoc.tipdoc, con_provicionesdetdoc.idclipro, mae_libros.codsun, Format([con_proviciones].[idmes],'00') & Format([mae_libros].[codsun],'00') & Right([con_proviciones].[numreg],4) AS numreg, con_proviciones.tc AS impven, con_proviciones.glosa,'Si' as espaertura " _
                        & " FROM mae_libros RIGHT JOIN (((((((con_provicionesdetdoc LEFT JOIN mae_cliente ON con_provicionesdetdoc.idclipro = mae_cliente.id) LEFT JOIN mae_prov ON con_provicionesdetdoc.idclipro = mae_prov.id) LEFT JOIN mae_moneda ON con_provicionesdetdoc.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_provicionesdetdoc.tipdoc = mae_documento.id) LEFT JOIN mae_bancos ON con_provicionesdetdoc.idclipro = mae_bancos.id) LEFT JOIN pla_empleados ON con_provicionesdetdoc.idclipro = pla_empleados.id) INNER JOIN con_proviciones ON con_provicionesdetdoc.idpro = con_proviciones.id) ON mae_libros.id = con_proviciones.idlib " _
                        & " WHERE (((con_provicionesdetdoc.id)=" & NulosN(RstDet!iddoc) & "));", xCon


                
                End If
                '--colocando la informacion cuando es seleccionado
                If xRs.State = 1 Then
                    If xRs.RecordCount <> 0 Then
                        With Me.Fg1
                            .TextMatrix(.Row, 1) = IIf(RstDet("tipreg") = 0, "0 Seleccion", "1 Manual")
                            .TextMatrix(.Row, 2) = NulosC(RstDet("Descripcion"))
                            .TextMatrix(.Row, 3) = NulosC(xRs("nombre"))
                            .TextMatrix(.Row, 4) = NulosC(xRs("numreg"))
                            .TextMatrix(.Row, 5) = NulosC(xRs("abrev"))
                            .TextMatrix(.Row, 6) = NulosC(xRs("nrodoc"))
                            .TextMatrix(.Row, 7) = Format(xRs("fchdoc"), FORMAT_DATE)
                            .TextMatrix(.Row, 8) = NulosC(xRs("desmoneda"))
                            .TextMatrix(.Row, 9) = NulosN(xRs("impven"))
                            
                            .TextMatrix(.Row, 14) = RstDet("idmod")
                            .TextMatrix(.Row, 15) = xRs("id")
                            .TextMatrix(.Row, 16) = xRs("TipDoc")
                            .TextMatrix(.Row, 17) = xRs("idmon")
                            .TextMatrix(.Row, 18) = NulosN(xRs("idclipro"))
                            
                            .TextMatrix(.Row, 20) = NulosC(xRs("glosa"))
                        
                        End With
                        
                        Call VisImportesMonedas(NulosN(xRs("idmon")), NulosN(xRs("imptot")))
                        
                        Set xRs = Nothing
                    End If
                End If
             Else
                                        
                With Me.Fg1
                    .TextMatrix(.Row, 1) = IIf(RstDet!tipreg = -1, "1 Manual", "0 Seleccion")
                    .TextMatrix(.Row, 2) = NulosC(RstDet!Descripcion)
                    .TextMatrix(.Row, 3) = NulosC(RstDet!nombre)
                    .TextMatrix(.Row, 4) = ""
                    .TextMatrix(.Row, 5) = NulosC(RstDet!abrev)
                    .TextMatrix(.Row, 6) = NulosC(RstDet!numser) + "-" + NulosC(RstDet!NumDoc)
                    .TextMatrix(.Row, 7) = NulosC(RstDet!fchdoc)
                    .TextMatrix(.Row, 8) = NulosC(RstDet!desmoneda)
                    .TextMatrix(.Row, 9) = NulosC(RstDet!tc)
                    
                    
                    
                    .TextMatrix(.Row, 14) = NulosN(RstDet!idmod)
                    .TextMatrix(.Row, 15) = NulosN(RstDet!iddoc) '0
                    .TextMatrix(.Row, 16) = NulosN(RstDet!TipDoc)
                    .TextMatrix(.Row, 17) = NulosN(RstDet!idmon) '
                    .TextMatrix(.Row, 18) = NulosN(RstDet!idper)
                    
                    .TextMatrix(.Row, 20) = NulosC(RstDet!glosa)
                    
                    Call VisImportesMonedas(NulosN(RstDet!idmon), NulosN(RstDet!imptot))
                    
                End With
                
             End If
             
            Fg1.TextMatrix(Fg1.Row, 13) = NulosN(RstDet!impsal) '--impsaldo
            Fg1.TextMatrix(Fg1.Row, 12) = NulosN(RstDet!impacue) '--impacue
            
            '*****************************************************
            '--cta contable
            Fg1.TextMatrix(Fg1.Row, 19) = NulosC(RstDet!idcuen) '--id
            Fg1.TextMatrix(Fg1.Row, 21) = NulosC(RstDet!ctanum) '--num
            Fg1.TextMatrix(Fg1.Row, 22) = NulosC(RstDet!ctanombre) '--descripcion
            '*****************************************************
            
            RstDet.MoveNext
        Loop
    End If
    
    Set RstDet = Nothing
    Mostrando = False
    
    Set RstDet = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & NulosN(TxtIdMon) & " and tipope = -1", xCon)
    If RstDet.RecordCount = 1 Then
        xCuentaDoc = RstDet("idcuen")
    End If
    
    
    HallarTotal
    
    
    
    Set RstDet = Nothing
End Sub

Sub Bloquea()
    
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    
    'If QueHace = 1 Then
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    'End If
    
    
    TxtIdAlm.Locked = Not TxtIdAlm.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
    
    
    TxtIdTipDoc.Locked = Not TxtIdTipDoc.Locked
    TxtNumDocRef.Locked = Not TxtNumDocRef.Locked
    
    ChkTC.Enabled = Not ChkTC.Enabled
    TxtTC.BackColor = &H8000000F

End Sub

Sub Blanquea()
    lblReg.Caption = ""
    TxtIdAlm = ""

    TxtTipDoc.Text = ""
    TxtNumRuc.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtFchDoc.Valor = ""
    txtglosa.Text = ""
    TxtIdMon.Text = ""

    
    LblAlmacen = ""
    LblNomDoc.Caption = ""
    LblNomCli.Caption = ""
    LblAlmacen = ""
    LblMoneda.Caption = ""
    LblIdCliente.Caption = ""

    
    TxtimpMN.Text = ""
    TxtImpME.Text = ""
    txtimpSaldo.Text = ""
    txtimpAcuenta.Text = ""

    TxtIdTipDoc.Text = ""
    TxtNumDocRef.Text = ""
    
    LblDescTipDocRef.Caption = ""
    
    LblIdDocRef2.Caption = ""
    
    ChkTC.Value = 0
    TxtTC.Text = ""
    
    Fg1.Rows = 1
End Sub


Private Sub CmdAddItem_Click()
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Sub
    End If
    
    If QueHace = 3 Then Exit Sub
    
    
    
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = "" Then
        Fg1.Col = 1
        Fg1.Row = Fg1.Rows - 1
'        Fg1.SetFocus
        Fg1_CellButtonClick Fg1.Rows - 1, 1
        Fg1.SetFocus
        Exit Sub
    End If
    
    Fg1.Rows = Fg1.Rows + 1
    
    
    
    With Fg1
        .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
    End With
    
    
   ' Fg1_CellButtonClick Fg1.Rows - 1, 1
    
    Fg1.SetFocus
End Sub

Sub CargarDocumentos(tipolgd As Integer, idclipro As Long)
    
    
    Dim xCampos(6, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim tcventa As Double
    Dim nSQLIdPer As String
    Dim nSQL As String
    Dim nTitulo As String
    
    Dim xRs1 As New ADODB.Recordset
    Dim X As Integer
    
    '1 Compras
    '2 Ventas
    '10 Reembolsables
    '9 Renta de 4ta
    '11 Liquidacion
            
        xCampos(0, 0) = "Cliente":    xCampos(0, 1) = "nombre":       xCampos(0, 2) = "3500":   xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
        xCampos(1, 0) = "Fch. Doc":   xCampos(1, 1) = "fchdoc":       xCampos(1, 2) = "1000":   xCampos(1, 3) = "C":     xCampos(1, 4) = "N"
        xCampos(2, 0) = "TD":         xCampos(2, 1) = "abrev":        xCampos(2, 2) = "600":    xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
        xCampos(3, 0) = "Nro Doc":    xCampos(3, 1) = "nrodoc":       xCampos(3, 2) = "1500":   xCampos(3, 3) = "C":     xCampos(3, 4) = "S"
        xCampos(4, 0) = "M":          xCampos(4, 1) = "desmoneda":    xCampos(4, 2) = "1000":   xCampos(4, 3) = "C":     xCampos(4, 4) = "N"
        xCampos(5, 0) = "Importe":    xCampos(5, 1) = "imptot":       xCampos(5, 2) = "1200":   xCampos(5, 3) = "N":     xCampos(5, 4) = "N"
                
        Select Case tipolgd
        
        Case 1
        
            'Compras
            If idclipro <> 0 Then nSQLIdPer = "WHERE  com_compras.idpro = " & idclipro & " "
            
            nSQL = " SELECT 0 as xsel, mae_prov.numruc, mae_prov.nombre, com_compras.fchdoc, mae_documento.abrev,  IIf([com_compras].[numser]='',[com_compras].[numdoc],[com_compras].[numser]+ '-' +[com_compras].[numdoc]) AS NroDoc, mae_moneda.simbolo as desmoneda, com_compras.imptot, com_compras.impsal, " _
                & " com_compras.id, com_compras.tipdoc, com_compras.idmon, com_compras.idpro as idclipro, com_compras.numreg, " _
                & " IIf([com_compras].[tc]=0,[con_tc].[impven],[com_compras].[tc]) & '' AS impven,com_compras.glosa,'No' as espaertura " _
                & " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_prov RIGHT JOIN com_compras ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " & nSQLIdPer _
                & " ORDER BY  IIf(com_compras.numser='',com_compras.numdoc,com_compras.numser+'-'+com_compras.numdoc) "
    
        Case 2
            'Ventas
            If idclipro <> 0 Then nSQLIdPer = " WHERE vta_ventas.idcli = " & idclipro & " "
            
            nSQL = " SELECT  0 as xsel, mae_cliente.numruc, mae_cliente.nombre, vta_ventas.fchdoc,mae_documento.abrev, vta_ventas.numser+ '-'+ vta_ventas.numdoc AS nrodoc, mae_moneda.simbolo as DesMoneda, vta_ventas.imptotdoc as imptot,vta_ventas.impsal, " _
                & " vta_ventas.id, vta_ventas.tipdoc, vta_ventas.idmon, vta_ventas.idcli as idclipro, vta_ventas.numreg, " _
                & " IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc]) & '' AS impven,vta_ventas.glosa,'No' as espaertura  " _
                & " FROM ((mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_documento.id = vta_ventas.tipdoc) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " & nSQLIdPer _
                & " ORDER BY vta_ventas.numser + '-'+ vta_ventas.numdoc DESC"
        Case 10
            'Reembolsables
            If idclipro <> 0 Then nSQLIdPer = " Where com_reembolsables.idpro = " & idclipro & " "
            
            nSQL = " SELECT 0 as xsel, mae_prov.numruc, mae_prov.nombre, com_reembolsables.fchdoc, mae_documento.abrev, [com_reembolsables]![numser]+ '-'+[com_reembolsables]![numdoc] AS nrodoc, mae_moneda.simbolo as desmoneda, com_reembolsables.imptot, com_reembolsables.impsal, " _
                & " com_reembolsables.id, com_reembolsables.tipdoc, com_reembolsables.idmon, com_reembolsables.idpro as idclipro, '' AS numreg, " _
                & " IIf([com_reembolsables].[tc]=0,[con_tc].[impven],[com_reembolsables].[tc]) & '' AS impven,com_reembolsables.glosa,'No' as espaertura " _
                & " FROM (((com_reembolsables LEFT JOIN mae_documento ON com_reembolsables.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON com_reembolsables.idmon = mae_moneda.id) INNER JOIN mae_prov ON com_reembolsables.idpro = mae_prov.id) LEFT JOIN con_tc ON com_reembolsables.fchdoc = con_tc.fecha " & nSQLIdPer _
                & " ORDER BY [com_reembolsables]![numser]+'-'+[com_reembolsables]![numdoc];"

        
        Case 9
            'Honorarios
            If idclipro <> 0 Then nSQLIdPer = " Where com_honorarios.idpro = " & idclipro & " "
            
            nSQL = " SELECT 0 as xsel, mae_prov.numruc, mae_prov.nombre, com_honorarios.fchdoc, mae_documento.abrev, [com_honorarios].[numser]+ '-' + [com_honorarios].[numdoc] AS nrodoc, mae_moneda.simbolo as Desmoneda , com_honorarios.imptot, com_honorarios.impsal, " _
                & " com_honorarios.id, com_honorarios.tipdoc, com_honorarios.idmon, com_honorarios.idpro as idclipro, com_honorarios.numreg, " _
                & " IIf([com_honorarios].[tc]=0,[con_tc].[impven],[com_honorarios].[tc]) & '' AS impven, com_honorarios.glosa,'No' as espaertura " _
                & " FROM (((mae_documento RIGHT JOIN com_honorarios ON mae_documento.id = com_honorarios.tipdoc) LEFT JOIN mae_moneda ON com_honorarios.idmon = mae_moneda.id) LEFT JOIN mae_prov ON com_honorarios.idpro = mae_prov.id) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha " & nSQLIdPer _
                & " ORDER BY com_honorarios.numser + '-' + com_honorarios.numdoc "

        Case 11
            'Liquidacion de Gastos
            If idclipro <> 0 Then nSQLIdPer = " Where vta_gastodebito.idcli =" & idclipro & " "
            
            nSQL = " SELECT 0 as xsel, mae_cliente.numruc, mae_cliente.nombre, vta_gastodebito.fchemi AS fchdoc,mae_documento.abrev, [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc] AS nrodoc, mae_moneda.simbolo AS desmoneda, vta_gastodebito.imptot, vta_gastodebito.impsal, " _
                & " vta_gastodebito.id, vta_gastodebito.tipdoc, vta_gastodebito.idmon, vta_gastodebito.idcli AS idclipro, vta_gastodebito.numreg, " _
                & " iif(vta_gastodebito.anulado=-1,0,IIf([vta_gastodebito].[tc]=0,[con_tc].[impven],[vta_gastodebito].[tc])) & '' AS impven, vta_gastodebito.glosa,'No' as espaertura " _
                & " FROM (mae_cliente RIGHT JOIN ((vta_gastodebito LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id) ON mae_cliente.id = vta_gastodebito.idcli) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha " & nSQLIdPer _
                & " ORDER BY [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc]"
        Case Else
        
            Exit Sub
        End Select
        '-FALTAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA!!!!!!!!!!!!!!!!!!!!!!!!!1
        
        '--uniendo con la informacion de apertura
'''''''        xfrm.SQLCad = xfrm.SQLCad & " union all " & _
'''''''            vbCr + "SELECT IIf(con_provicionesdetdoc.idtipper=1,mae_prov.numruc,IIf(con_provicionesdetdoc.idtipper=2,mae_cliente.numruc,IIf(con_provicionesdetdoc.idtipper=3,pla_empleados.numdoc,IIf(con_provicionesdetdoc.idtipper=5,mae_bancos.numruc,'')))) AS numruc, " _
'''''''                & " IIf(con_provicionesdetdoc.idtipper=1,mae_prov.nombre,IIf(con_provicionesdetdoc.idtipper=2,mae_cliente.nombre,IIf(con_provicionesdetdoc.idtipper=3,pla_empleados.nombre,IIf(con_provicionesdetdoc.idtipper=5,mae_bancos.descripcion,'')))) AS nombre, " _
'''''''                & " con_provicionesdetdoc.fchemi AS fchdoc, mae_documento.abrev, [con_provicionesdetdoc].[numser] & '-' & [con_provicionesdetdoc].[numdoc] AS nrodoc, mae_moneda.simbolo AS desmoneda, con_provicionesdetdoc.impdoc AS imptot, con_provicionesdetdoc.impsal, con_provicionesdetdoc.id, con_provicionesdetdoc.tipdoc, con_provicionesdetdoc.idclipro, mae_libros.codsun, Format([con_proviciones].[idmes],'00') & Format([mae_libros].[codsun],'00') & Right([con_proviciones].[numreg],4) AS numreg, con_proviciones.tc AS impven, con_proviciones.glosa,'Si' as espaertura " _
'''''''                & " FROM mae_libros RIGHT JOIN (((((((con_provicionesdetdoc LEFT JOIN mae_cliente ON con_provicionesdetdoc.idclipro = mae_cliente.id) LEFT JOIN mae_prov ON con_provicionesdetdoc.idclipro = mae_prov.id) LEFT JOIN mae_moneda ON con_provicionesdetdoc.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_provicionesdetdoc.tipdoc = mae_documento.id) LEFT JOIN mae_bancos ON con_provicionesdetdoc.idclipro = mae_bancos.id) LEFT JOIN pla_empleados ON con_provicionesdetdoc.idclipro = pla_empleados.id) INNER JOIN con_proviciones ON con_provicionesdetdoc.idpro = con_proviciones.id) ON mae_libros.id = con_proviciones.idlib " _
'''''''                & " WHERE (((con_provicionesdetdoc.id)=" & NulosN(RstDet!iddoc) & "));"
'''''''
        '--------------------------------------
        
        nTitulo = "Buscando Documentos de Operación "
        
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), nTitulo
        
        If xRs.State = 1 Then
            If xRs.RecordCount = 0 Then
                Set xRs = Nothing
                Exit Sub
            End If
            
            Agregando = True
            
            xRs.MoveFirst
            Do While Not xRs.EOF
                With Me.Fg1
                                               
                    If X >= 1 Then
                        .AddItem ""
                        .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Row, 1)
                        .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Row, 2)
                        .TextMatrix(.Rows - 1, 14) = .TextMatrix(.Row, 14)
                    End If
                    .Row = .Rows - 1
                    
                    .TextMatrix(.Row, 3) = NulosC(xRs("nombre"))
                    .TextMatrix(.Row, 4) = NulosC(xRs("numreg"))
                    .TextMatrix(.Row, 5) = NulosC(xRs("abrev"))
                    .TextMatrix(.Row, 6) = NulosC(xRs("nrodoc"))
                    .TextMatrix(.Row, 7) = NulosC(xRs("fchdoc"))
                    .TextMatrix(.Row, 8) = NulosC(xRs("desmoneda"))
                    .TextMatrix(.Row, 9) = NulosN(xRs("impven"))
                    
                    '.TextMatrix(.Row, 10) = Format(xRs!imptot, FORMAT_MONTO)
                    
                    
                    Call VisImportesMonedas(xRs("idmon"), NulosN(xRs("imptot")))
                                                                    
                    'colocamos el idcuenta del detalle por documento cargado

                    .TextMatrix(.Row, 15) = xRs("id")
                    .TextMatrix(.Row, 16) = NulosN(xRs("tipdoc"))
                    .TextMatrix(.Row, 17) = NulosN(xRs("idmon"))
                    .TextMatrix(.Row, 18) = NulosN(xRs("idclipro"))
                    
                    .TextMatrix(.Row, 20) = NulosC(xRs("glosa"))
                    
                    '************************************
                    '--cuenta contable
                    pHallaCtaDetalle NulosN(xRs("tipdoc")), NulosN(xRs("idmon")), .Row
                    '************************************
                    X = X + 1
                    
                End With
                
                xRs.MoveNext
            Loop
        Agregando = False
        End If
    
    
    HallarTotal
    
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



Private Sub CmdBusDocRef2_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Nº Documento":      xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Emi.":         xCampos(1, 1) = "fchemi":      xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Fch. Ven.":         xCampos(2, 1) = "fchven":      xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    
    If NulosN(TxtIdTipDoc.Text) = 4 Then
        'Orden de Despacho
        xform.SQLCad = "SELECT var_ordendespacho.id, var_ordendespacho.numerodoc AS numdoc, " _
            & " mae_cliente.nombre, var_ordendespacho.idcli, var_ordendespacho.fchemi, var_ordendespacho.fchven FROM var_ordendespacho LEFT JOIN mae_cliente " _
            & " ON var_ordendespacho.idcli = mae_cliente.id WHERE var_ordendespacho.idcli =" & NulosN(LblIdCliente) & ""
        
        xform.Titulo = "Orden de Despacho"
    ElseIf NulosN(TxtIdTipDoc.Text) = 5 Then
        'Orden de pedido
        MsgBox "Opcion no disponible"
        xform.Titulo = "Orden de Producción"
        Exit Sub
    Else
        Exit Sub
    End If
    
    xform.FormaBusca = CualquierParte
    xform.Criterio = ""
    xform.Ordenado = "numdoc"
    xform.CampoBusca = "numdoc"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumDocRef.Text = NulosC(xRs("numdoc"))
            LblIdDocRef2.Caption = NulosN(xRs("id"))
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
    
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "idtipdoc":       xCampos(0, 2) = "1500":    xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion": xCampos(1, 2) = "2500":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Serie":          xCampos(2, 1) = "numser":      xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"

    
        xform.SQLCad = " SELECT mae_documento.*,  alm_numseries.idtipdoc, alm_numseries.numser " & _
                   " FROM mae_documento LEFT JOIN alm_numseries ON mae_documento.id = alm_numseries.idtipdoc " & _
                   " WHERE alm_numseries.idalm =" & NulosN(Me.TxtIdAlm) & " AND alm_numseries.idtipdoc = " & NulosN(TxtTipDoc) & ""

        
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
            
            Dim Rst As New ADODB.Recordset
            RST_Busq Rst, "SELECT top 1 numdoc AS numero from vta_gastodebito  WHERE numser ='" & NulosC(TxtNumSer.Text) & "' AND tipdoc =" & NulosN(TxtTipDoc) & " ORDER BY numdoc DESC ", xCon

            If Rst.RecordCount = 0 Then
                TxtNumDoc.Text = "0000000001"
            Else
                Rst.MoveFirst
                TxtNumDoc.Text = Format(NulosN(Rst("numero")) + 1, "0000000000")
            End If
            Set Rst = Nothing
        End If
        
        TxtIdTipDoc.SetFocus
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
    

    

    xform.SQLCad = " SELECT mae_documento.*, alm_numseries.numser" & _
                    " FROM mae_documento LEFT JOIN alm_numseries ON mae_documento.id = alm_numseries.idtipdoc " & _
                    " WHERE alm_numseries.idalm =" & NulosN(TxtIdAlm) & "  AND ( alm_numseries.idtipdoc = 120 OR alm_numseries.idtipdoc = 126 )"

    
    
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
            Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & NulosN(TxtIdMon) & " and tipope = -1", xCon)
            If xRs2.RecordCount > 0 Then
                xCuentaDoc = NulosN(xRs2("idcuen"))
            End If
            Set xRs2 = Nothing
            TxtNumRuc.SetFocus
        End If
    
        If TxtTipDoc.Text <> "" Then
            Dim Rst As New ADODB.Recordset
            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(TxtIdAlm) & "", xCon)
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
    
    xform.SQLCad = " SELECT mae_documento.*, alm_numseries.numser" & _
                    " FROM mae_documento LEFT JOIN alm_numseries ON mae_documento.id = alm_numseries.idtipdoc " & _
                    " WHERE alm_numseries.idalm =" & NulosN(LblidAlmacen2) & "  AND ( alm_numseries.idtipdoc = 120 OR alm_numseries.idtipdoc = 126 )"

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
            
            
            Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(LblIdDocumentoGen.Caption) & " and mae_documentocta.idmon = 1  and tipope = -1", xCon)
            If xRs2.RecordCount > 0 Then
                xCuentaDoc = NulosN(xRs2("idcuen"))
            Else
                MsgBox "No se ha encontrado cuenta contable para el documento especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
            Set xRs2 = Nothing
            

        End If
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
    Dim xRs As New ADODB.Recordset
    Dim xId As Double
    Dim xNumAsiento As String

    RST_Busq xRs, "SELECT vta_gastodebito.tipdoc, vta_gastodebito.numser, vta_gastodebito.numdoc From vta_gastodebito " _
        & " WHERE (((vta_gastodebito.tipdoc)=" & NulosN(LblIdDocumentoGen.Caption) & ") AND ((vta_gastodebito.numser)='" & TxtNumSerGen.Text & "') " _
        & " AND ((vta_gastodebito.numdoc)='" & TxtNumDocGen.Text & "'))", xCon

    If xRs.RecordCount = 1 Then
        Set xRs = Nothing
        MsgBox "El numero de documento que quiere emitir ya existe", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDocGen.SetFocus
        Exit Sub
    End If
    
    On Error GoTo LaCague
    
    xCon.BeginTrans

    'Validar si el nro de documento existe solo en modo adicionar documento
    RST_Busq RstCab, "SELECT TOP 1 * FROM vta_gastodebito", xCon
    
    xId = HallaCodigoTabla("vta_gastodebito", xCon, "id")
    
    mIdRegistro = xId
    
    RstCab.AddNew
    RstCab("id") = xId
    RstCab("idlib") = 41
    RstCab("tipdoc") = NulosN(LblIdDocumentoGen.Caption)
    RstCab("idcli") = 0
    RstCab("numser") = TxtNumSerGen.Text
    RstCab("numdoc") = TxtNumDocGen.Text
    RstCab("Fchemi") = TxtFchEmiAnul.Valor
    RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    RstCab("idmes") = mMesActivo
    RstCab("idmon") = 1
    RstCab("impbru") = 0
    RstCab("impina") = 0
    RstCab("igv") = 0
    RstCab("imptot") = 0
    RstCab("impsal") = 0
    RstCab("idmon") = 1
    'RstCab("numreg") = Format(mMesActivo, "00") + Trim(xNumAsiento)
    RstCab("anulado") = -1
    RstCab("glosa") = "ANULADO"
    RstCab.Update
    
    '--generando el asiento contable
    xNumAsiento = GenerarAsiento(xCon, 41, xId, AnoTra, mMesActivo, 1, 0)
    If xNumAsiento = "" Then GoTo LaCague
            
        
    'Grabamos el movimiento en la tabla var_edicion
    '--parametro idOperacion por defecto = 1 Nuevo Registro
    GrabarOperacion xIdUsuario, IdMenuActivo, 1, xHorIni, Time, Date, xCon, xId

    
        
    xCon.CommitTrans
        
    MsgBox "El documento anulado se genero con éxito" & vbCr & "Registro Nº: " & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
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


Private Sub cmdsalirseldoc_Click()
    ActivarEntorno
    Fraseldoc.Visible = False
End Sub


Private Sub Command1_Click()
    Dim Rpta As Integer
    
    Rpta = MsgBox("Esta seguro de modificar el saldo del documento", vbInformation + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        'actualizamos el saldo del documento en vta_gastodebito
        xCon.Execute "UPDATE vta_gastodebito SET vta_gastodebito.impsal = " & NulosN(TxtNewSaldo2.Text) & " WHERE (((vta_gastodebito.id)=" & RstVent("id") & "))"

        
    End If
End Sub

Private Sub Command2_Click()
    ActivarEntorno
    Frame8.Visible = False
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
        VerMovimientos1 IdMenuActivo, NulosN(RstVent("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    CargarDatosEnDetalle Row, Col
                    
End Sub


Private Sub CargarDatosEnDetalle(xFil As Long, xCol As Long)
    '===================================================================================================
    'Creado : 23/09/11 Por: Johan Castro
    'Propósito: Mostrar ventana de seleccion
    '
    'Entradas:  xfil= Posicion de la Fila
    '           xCol= Posicion de la Columna
    '
    '===================================================================================================

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    Dim TipDoc As Integer
    Dim idmondoc As Integer


    'Descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    If xCol = 2 Then
                
                    
        xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "4800":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Código":       xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":    xCampos(1, 3) = "N"
                                
        xform.SQLCad = "SELECT tes_modulos.descripcion, tes_modulos.id  " _
            & " FROM tes_modulos  ORDER BY tes_modulos.descripcion "
        
        xform.Titulo = "Buscando Tipos de Movimiento"
        
        xform.FormaBusca = Principio
        xform.Criterio = ""
                
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
                
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        Dim A As Integer
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 2) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 14) = xRs("id") 'Tipo de Movimiento
            End If
        End If
        
        
        If NulosN(LblIdCliente.Caption) <> 0 Then
        
            'SI EN CASO EL TIPO ES VENTAS  O GASTOS DE DEBITO
            If NulosN(Fg1.TextMatrix(Fg1.Row, 14)) = 2 Or NulosN(Fg1.TextMatrix(Fg1.Row, 14)) = 5 Then
            
                Fg1.TextMatrix(Fg1.Row, 3) = LblNomCli.Caption
                Fg1.TextMatrix(Fg1.Row, 18) = LblIdCliente.Caption
            End If
        
        End If
                
    
    'Razon Social
    ElseIf xCol = 3 Then
                                    
        
        'VALIDAMOS SI ES
        
        xCampos(0, 0) = "Cliente":      xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
        
    
        'SI ES VENTAS - LGD
        If NulosN(Me.Fg1.TextMatrix(Me.Fg1.Row, 14)) = 2 Or NulosN(Me.Fg1.TextMatrix(Me.Fg1.Row, 14)) = 11 Then
            xform.SQLCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id From mae_cliente"
            xform.Titulo = "Buscando Cliente"
        
        '--
        ElseIf NulosN(Me.Fg1.TextMatrix(Me.Fg1.Row, 14)) = 1 Or NulosN(Me.Fg1.TextMatrix(Me.Fg1.Row, 14)) = 9 Or NulosN(Me.Fg1.TextMatrix(Me.Fg1.Row, 14)) = 11 Then
            xform.SQLCad = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id From mae_prov WHERE mae_prov.tipper <>  2 "
            xform.Titulo = "Buscando Proveedor"
            
        Else
            xform.SQLCad = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id From mae_prov"
            xform.Titulo = "Buscando Proveedor"
            
        End If
        
        
        xform.FormaBusca = Principio
                      
        xform.Criterio = ""
        xform.Ordenado = "nombre"
        xform.CampoBusca = "nombre"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 3) = NulosC(xRs("nombre"))
                Fg1.TextMatrix(Fg1.Row, 18) = NulosN(xRs("id")) 'idcli
            End If
        End If
                        
    
    ElseIf xCol = 6 Then
        '--lista de documentos
        If NulosN(Fg1.TextMatrix(Me.Fg1.Row, 14)) = 0 Then
            MsgBox "Falta especificar el Origen", vbExclamation, xTitulo
            Fg1.Col = 2
            Exit Sub
        End If
        
        Call CargarDocumentos(NulosN(Fg1.TextMatrix(Me.Fg1.Row, 14)), NulosN(Fg1.TextMatrix(Me.Fg1.Row, 18)))
        
    
    'Tipo de Documentos
    ElseIf xCol = 5 Then
    
        xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
        xform.SQLCad = "SELECT mae_documento.* FROM mae_documento ORDER BY Descripcion"

        xform.Titulo = "Buscando Tipo de Documento"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
    
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 5) = NulosC(xRs("abrev"))
                Fg1.TextMatrix(Fg1.Row, 16) = NulosN(xRs("id")) 'tipdoc
                
                pHallaCtaDetalle NulosN(Fg1.TextMatrix(xFil, 16)), NulosN(Fg1.TextMatrix(xFil, 17)), xFil
                
                
            End If
        End If
    
    'Monedas
    ElseIf xCol = 8 Then
 
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
                Fg1.TextMatrix(Fg1.Row, 8) = NulosC(xRs("simbolo"))
                Fg1.TextMatrix(Fg1.Row, 17) = NulosN(xRs("id")) 'idmon
             End If
             
             pHallaCtaDetalle NulosN(Fg1.TextMatrix(xFil, 16)), NulosN(Fg1.TextMatrix(xFil, 17)), xFil
             
             '--si es soles enviar como parametro columna en soles caso contrario en moneda extranjera
             VisImportesMonedas NulosN(Fg1.TextMatrix(xFil, 17)), IIf(NulosN(xRs("id")) = 1, NulosN(Fg1.TextMatrix(xFil, 10)), NulosN(Fg1.TextMatrix(xFil, 11)))
             
             HallarTotal
             
        End If
                
    ElseIf xCol = 21 Then
        
        xCampos(0, 0) = "Cuenta":         xCampos(0, 1) = "cuenta":           xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Descripción":    xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "4000":         xCampos(1, 3) = "C"
        
        xform.SQLCad = "SELECT * FROM con_planctas ORDER BY cuenta"
        
        xform.Titulo = "Buscando Cuenta Contable"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "cuenta"
        xform.CampoBusca = "cuenta"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 18) = NulosN(xRs("id")) '
                Fg1.TextMatrix(Fg1.Row, 21) = NulosC(xRs("cuenta"))
                Fg1.TextMatrix(Fg1.Row, 22) = NulosC(xRs("descripcion")) '
             End If
        
        End If
            
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    
    Dim xRs1 As New ADODB.Recordset
    'Dim tipdoc As Integer
    'Dim idmondoc As Integer
    
    If Agregando = True Then Exit Sub
    If Mostrando = True Then Exit Sub
    If Fg1.Row < 0 Then Exit Sub
    
    
    'si se modifica documento ó moneda
    'If Col = 5 Or Col = 8 Then
    
    'tipdoc = Fg1.TextMatrix(Me.Fg1.Row, 14)
    'idmondoc = Fg1.TextMatrix(Me.Fg1.Row, 15)
    
    
   ' RST_Busq xrs1, "Select idcuen FROM mae_documentolgdcta WHERE iddoc =" & nulosn(TxtTipDoc) & " AND iddocref = " & tipdoc & " AND idmon =" & idmondoc & "", xCon
   '                             If xrs1.RecordCount <> 0 Then
   '                                 Fg1.TextMatrix(Fg1.Row, 17) = NulosN(xrs1!idcuen)
   '                             End If
    
    'End If
    If Col = 3 Then
        If NulosC(Fg1.TextMatrix(Row, Col)) = "" Then Fg1.TextMatrix(Row, 18) = 0
        
    ElseIf Col = 10 Then
        '--cambiar si la moneda es en moneda nacional
         VisImportesMonedas NulosN(Fg1.TextMatrix(Fg1.Row, 17)), NulosN(Fg1.TextMatrix(Fg1.Row, 10))
        HallarTotal
        
    ElseIf Col = 11 Then
        '--cambiar si la moneda es en moneda extran
         VisImportesMonedas NulosN(Fg1.TextMatrix(Fg1.Row, 17)), NulosN(Fg1.TextMatrix(Fg1.Row, 11))
        HallarTotal
        
    ElseIf Col = 12 Then
    
        Agregando = True
        If IsNumeric(Fg1.TextMatrix(Fg1.Row, 12)) = False Then Fg1.TextMatrix(Fg1.Row, 12) = 0
        If NulosN(Fg1.TextMatrix(Fg1.Row, 17)) = 2 Then
            Fg1.TextMatrix(Fg1.Row, 13) = NulosN(Fg1.TextMatrix(Fg1.Row, 11)) - NulosN(Fg1.TextMatrix(Fg1.Row, 12))
        Else
            Fg1.TextMatrix(Fg1.Row, 13) = NulosN(Fg1.TextMatrix(Fg1.Row, 10)) - NulosN(Fg1.TextMatrix(Fg1.Row, 12))
        End If
        Agregando = False
        
        HallarTotal
        
    ElseIf Col = 21 Then
    
        RST_Busq xRs1, "SELECT * FROM con_planctas where cuenta='" & NulosC(Fg1.TextMatrix(Row, 21)) & "'", xCon
        
        If xRs1.RecordCount <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 19) = NulosN(xRs1("id")) '
            Fg1.TextMatrix(Fg1.Row, 21) = NulosC(xRs1("cuenta"))
            Fg1.TextMatrix(Fg1.Row, 22) = NulosC(xRs1("descripcion")) '
        End If

            
    End If
    
    Set xRs1 = Nothing
End Sub

Sub HallarTotal()
    
    TxtimpMN.Text = Format(GRID_SUMAR_COL(Fg1, 10), FORMAT_MONTO)
    TxtImpME.Text = Format(GRID_SUMAR_COL(Fg1, 11), FORMAT_MONTO)
    txtimpAcuenta.Text = Format(GRID_SUMAR_COL(Fg1, 12), FORMAT_MONTO)
    txtimpSaldo.Text = Format(GRID_SUMAR_COL(Fg1, 13), FORMAT_MONTO)
    
End Sub


Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Agregando = True Then Exit Sub
    

'    If Fg1.Col = 2 Then
    'SI ES MANUAL
    
    If (NulosC(Fg1.TextMatrix(Me.Fg1.Row, 1)) = "") And Fg1.Col > 2 Then
        Fg1.Editable = flexEDNone
    Else
         If NulosN(Mid(Me.Fg1.TextMatrix(Me.Fg1.Row, 1), 1, 1)) = 1 Then
            Fg1.ColComboList(2) = "|..."
            Fg1.Editable = flexEDKbdMouse
            Fg1.ColComboList(3) = "|..."
            Fg1.Editable = flexEDKbdMouse
            
            Fg1.ColComboList(6) = ""
            Fg1.Editable = flexEDKbdMouse
            
            Fg1.ColComboList(5) = "|..."
            Fg1.Editable = flexEDKbdMouse
            Fg1.ColComboList(8) = "|..."
            Fg1.Editable = flexEDKbdMouse
            
         Else
            Fg1.ColComboList(2) = "|..."
            Fg1.Editable = flexEDKbdMouse
            
            Fg1.ColComboList(3) = "|..."
            Fg1.Editable = flexEDKbdMouse
            
            Fg1.ColComboList(6) = "|..."
            Fg1.Editable = flexEDKbdMouse
            
            Fg1.ColComboList(5) = ""
            Fg1.Editable = flexEDKbdMouse
            Fg1.ColComboList(8) = ""
            Fg1.Editable = flexEDKbdMouse
        End If
        
        Fg1.ColComboList(21) = "|..."
        Fg1.Editable = flexEDKbdMouse
        
    End If
    
    
        'Fg1.Refresh
        
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    If KeyAscii = 13 Then Exit Sub
    
    '--si es seleccionado no hacer nada
    If NulosN(Mid(Me.Fg1.TextMatrix(Me.Fg1.Row, 1), 1, 1)) = 0 Then
            Select Case Col
            Case 4, 5, 6, 7, 8, 20, 22
                KeyAscii = 0
            End Select
            
    Else
            Select Case Col
                Case 9, 10, 11, 12, 13
                    If validar_numero(KeyAscii) = False Then KeyAscii = 0
                Case 2, 3, 4, 5, 8, 11, 22
                    KeyAscii = 0
            End Select
    End If
    
    
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = vbKeyF10 Then
        CargarDatosEnDetalle Fg1.Row, Fg1.Col
        Fg1.SetFocus
    End If
     '   If KeyCode = 45 Then CmdAddItem_Click
    
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
       
End Sub


Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        mMesActivo = xMes
        pCargarGrid
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
    Agregando = False
    
    Dg1.Columns("fchdoc1").NumberFormat = FORMAT_DATE
    'Dg1.Columns("fchven1").NumberFormat = FORMAT_DATE
    
    'Dg1.Columns("impbru1").NumberFormat = FORMAT_MONTO
    'Dg1.Columns("impigv1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("imptot1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impsal1").NumberFormat = FORMAT_MONTO
        
    CaracteresNumericos = "0123456789." & Chr(8)
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    Fg1.SelectionMode = flexSelectionByRow
    
    Fg1.ColWidth(4) = 0 'numregistro
    Fg1.ColWidth(14) = 0 'idmod
    Fg1.ColWidth(15) = 0 'idcli
    Fg1.ColWidth(16) = 0 'tipdoc
    Fg1.ColWidth(17) = 0 'idmon
    Fg1.ColWidth(18) = 0 'idper
    Fg1.ColWidth(19) = 0 'idcta
    
''    Fg1.ColFormat(7) = FORMAT_DATE
''    Fg1.ColFormat(10) = FORMAT_MONTO
''    Fg1.ColFormat(11) = FORMAT_MONTO
''    Fg1.ColFormat(12) = FORMAT_MONTO
''    Fg1.ColFormat(13) = FORMAT_MONTO
    
    swguiafact = 0
    
    TxtFchDoc.Valor = Date

    
    TxtFchDoc.Valor = ""

    
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

Private Sub menu2_1_Click()
    'cmdagregardocs_Click
End Sub


Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        'Validamos si la cuadricula tiene datos
        If QueHace = 3 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No existe información para visualizar", vbInformation, Me.Caption
                Blanquea
                Exit Sub
            ElseIf NulosC(RstVent("nombre")) = "ANULADO" Then
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
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        If RstVent.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then
        If RstVent.RecordCount = 0 Then
            MsgBox "No se han registrado ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
    
        'Validamos si el documento esta anulado
        If RstVent("Anulado") = -1 Then
            MsgBox RstVent("nomdoc") & " ya fue anulado, seleccione otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Anular
    End If
        
    If Button.Index = 5 Then
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
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstVent.Filter = ""
    End If
    If Button.Index = 10 Then Buscar
    If Button.Index = 11 Then CambiarMes
    
    
    If Button.Index = 13 Then pExportar
    If Button.Index = 14 Then Imprimir
    
    If Button.Index = 16 Then
        Set RstVent = Nothing
        Unload Me
    End If
End Sub

Sub OpcionesPeriodo()
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    LblPeriodo2.Caption = LblMes.Caption
    
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    
    '--mostrar el boton para agregar apertura
    If mMesActivo = 0 Then CmdApertura.Visible = True Else CmdApertura.Visible = False
    '-------------------------

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        'MODIFICACION DE DOCUMENTOS
        If ButtonMenu.Index = 1 Then
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
        If ButtonMenu.Index = 2 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opcion", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If RstVent("anulado") = -1 Then ' SI EL DOCUMENTO ESTA ANULADO
                RestaurarFactura
            End If
        End If
    End If
  
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opcion", vbInformation, Me.Caption
                Exit Sub
            End If
            Anular
        End If
        If ButtonMenu.Index = 2 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opcion", vbInformation, Me.Caption
                Exit Sub
            End If
            
            Eliminar
        End If
        
        If ButtonMenu.Index = 3 Then EmitirAnulada
        
    End If
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

Private Sub TxtFchDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtFchDoc.Valor) <> "" Then
        If ChkTC.Value = 0 Then TxtTC.Text = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
    Else
        If ChkTC.Value = 0 Then TxtTC.Text = "0.00"
    End If
End Sub

Private Sub txtglosa_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtIdAlm_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        SendKeys vbTab
        
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdAlm_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
        CmdBusAlm_Click
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
        LblMoneda.Caption = (NulosC(xRs1("descripcion")))
    End If
    Set xRs1 = Nothing

End Sub

Private Sub TxtIdTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusIdTipDocRef_Click
    End If
End Sub

Private Sub TxtIdTipDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosN(TxtIdTipDoc.Text) = 0 Then Exit Sub
    
    Dim xRs1 As New ADODB.Recordset
    
    RST_Busq xRs1, "SELECT * FROM MAE_DocReferencia WHERE id = " & NulosN(TxtIdTipDoc.Text) & "", xCon
    
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
        If QueHace <> 1 Then nSQL = " and vta_gastodebito.id <> " & NulosN(RstVent("id"))
        
        RST_Busq Rst, "SELECT vta_gastodebito.numser, vta_gastodebito.numdoc, vta_gastodebito.fchemi, mae_cliente.nombre, Left([vta_gastodebito].[numreg],2) & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & Right([vta_gastodebito].[numreg],4) AS registro FROM (mae_cliente RIGHT JOIN vta_gastodebito ON mae_cliente.id = vta_gastodebito.idcli) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id " _
            & " WHERE (((vta_gastodebito.numser)='" & Trim(TxtNumSer.Text) & "') AND ((vta_gastodebito.numdoc)='" & TxtNumDoc.Text & "'))" & nSQL, xCon
                
        If Rst.RecordCount <> 0 Then
            '--poner el nuevo numero doc
            TxtNumSer_Validate True
            MsgBox "El número de documento de gasto de debito - credito ya existe " & vbCr & "Nº Registro: " & NulosC(Rst("registro")) & vbCr & "Fecha Doc.   " & NulosC(Rst("fchemi")) & vbCr & "Cliente:         " & NulosC(Rst("nombre")) & vbCr & "Será reemplazado por " + Trim(TxtNumSer.Text) + "-" + Trim(TxtNumDoc.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
        
    Else
        TxtNumRuc.Text = ""
        LblNomCli.Caption = ""
        LblIdCliente.Caption = ""
        

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
        
        RST_Busq Rst, "SELECT top 1 numdoc AS numero from vta_gastodebito  WHERE numser ='" & NulosC(TxtNumSer.Text) & "' AND tipdoc =" & NulosN(TxtTipDoc) & " ORDER BY numdoc DESC ", xCon

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
    
    xHorIni = Time
    
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
    
    
    
    'Cuenta Contable del documento
    If NulosN(TxtTipDoc.Text) <> 0 Then
        If xCuentaDoc = 0 Then
            MsgBox "No se ha asignado una cuenta contable al documento " + LblNomDoc.Caption & Chr(13) _
                & "Asignele una cuenta en el menu Contabilidad opcion Asignar Ctas. Contables a documentos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    End If
    
    
    'Cuenta Contable por item
    If mMesActivo <> 0 Then
        For A = 1 To Fg1.Rows - 1
            If NulosN(Fg1.TextMatrix(A, 19)) = 0 Then
                MsgBox "No se le ha asignado una cuenta contable para venta al item : " & Chr(13) _
                    & Fg1.TextMatrix(A, 3) & Chr(13) _
                    & "Asignele una cuenta  ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Function
            End If
        Next A
        'id del tipo origen ver tabla tes_modulos
        For A = 1 To Fg1.Rows - 1
            If NulosN(Fg1.TextMatrix(A, 14)) = 0 Then
                MsgBox "Seleccione el origen " & Chr(13) _
                    & Fg1.TextMatrix(A, 13), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Function
            End If
        Next A
    End If
    

    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento de venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    
    
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
    
    If mMesActivo <> 0 Then
        If CDate(TxtFchDoc.Valor) > CDate(xFchFin) Then
            MsgBox "No se puede grabar este documento en el periodo actual la fecha de emision es mayor a : " + xFchFin, vbInformation + vbOKOnly + vbDefaultButton1
            TxtFchDoc.SetFocus
            Exit Function
        End If
        
        If Fg1.Rows = 1 Then
            MsgBox "No ha especificado items para la liquidacion de gastos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Fg1.SetFocus
            Exit Function
        End If
    End If
    
    If TxtIdMon.Text = "" Then
        MsgBox "No ha especificado la moneda del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    
    
    If QueHace = 1 Then 'Validamos si existe el numero del documento en modo adicion
        Dim RstCab As New ADODB.Recordset
    
        RST_Busq RstCab, "SELECT * FROM vta_gastodebito  WHERE tipdoc =" & NulosN(TxtTipDoc.Text) & " AND numser ='" & TxtNumSer.Text & "' AND numdoc = '" & TxtNumDoc.Text & "' ", xCon
    
        If RstCab.RecordCount > 0 Then
            MsgBox "El Nro de documento ha sido registrado por otro usuario se grabara con otro numero", vbInformation, Me.Caption
            'TxtNumDoc.Text = HallaNumdocVenta(NulosN(TxtTipDoc.Text), TxtNumSer.Text, xCon)
        End If
        
        Set RstCab = Nothing
    End If
    
    'If NulosN(TxtIdTipDoc.Text) <> 0 Then
    '    If NulosN(LblIdDocRef2.Caption) = 0 Then
    '        MsgBox "No ha especificado el documento de referencia para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    '        TxtNumDocRef.SetFocus
    '        Exit Function
    '    End If
    'End If
    
    
    
    Dim RstDeta2 As New ADODB.Recordset
    Dim RstActPro As New ADODB.Recordset
    
    Dim RstDet As New ADODB.Recordset
'''    Dim RstDia As New ADODB.Recordset
    Dim xIdCuen As Integer
    Dim xTotal As Double
    Dim xSaldo As Double
    
    
    Dim xNumAsiento As String
    
    Dim xId As Double
    Dim X As Integer
    Dim P As Integer
    
    Dim nSQL As String
    
On Error GoTo LaCague
    
    swguiafact = 1
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("vta_gastodebito", xCon, "id")
        xNumAsiento = NuevoNumAsiento(41, mMesActivo, xCon)
        RST_Busq RstCab, "SELECT TOP 1 * FROM vta_gastodebito", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
        
    Else
        xId = RstVent("id")
        RST_Busq RstCab, "SELECT * FROM vta_gastodebito WHERE id = " & xId & "", xCon
        
        'Eliminamos el detalle
        xCon.Execute "DELETE * FROM vta_gastodebitodet WHERE idlgd  = " & xId & ""
        
        'Diario
        
'''         RST_Busq RstDia, "SELECT * FROM con_diario WHERE idmes = " & Format(CDate(TxtFchDoc.Valor), "mm") & " AND " _
'''            & " idlib = 41 AND idmov = " & xId & " And iddoc = " & NulosN(TxtTipDoc) & "", xCon
'''
'''         If RstDia.RecordCount <> 0 Then
'''             xNumAsiento = RstDia("numasi")
'''         Else
'''             xNumAsiento = NuevoNumAsiento(41, mMesActivo, xCon)
'''         End If
         
'''         Set RstDia = Nothing
         
'''        'Eliminamos el asiento contable
'''         xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & Format(CDate(TxtFchDoc.Valor), "mm") & " AND " _
'''             & " idlib = 41 AND idmov = " & xId & " And iddoc = " & NulosN(TxtTipDoc) & ""
            
        
        
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM vta_gastodebitodet", xCon
    'RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    
    mIdRegistro = xId
    RstCab("idlib") = 41
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("idcli") = NulosN(LblIdCliente.Caption)
    RstCab("numser") = TxtNumSer.Text
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("fchemi") = CDate(TxtFchDoc.Valor)
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("impina") = 0
    RstCab("igv") = 0
    RstCab("idmes") = mMesActivo
    
    RstCab("imptot") = NulosN(txtimpAcuenta.Text)
    RstCab("impsal") = NulosN(txtimpAcuenta.Text)
   
    RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    RstCab("idtipdocref") = NulosN(TxtIdTipDoc)
    RstCab("iddocref2") = NulosN(LblIdDocRef2)
    
    RstCab("tc") = NulosN(TxtTC.Text)
    
    RstCab("glosa") = NulosC(txtglosa)
    RstCab("anulado") = 0
    
    '--provisional
    RstCab("numerodocref") = NulosC(TxtNumDocRef.Text)
    
    'Actualizamos el saldo del documento
    'ActualizaSaldoDoc NulosN(LblIdDocRef.Caption), 2, NulosN(TxtTotal.Text)
    
    RstCab.Update
    
    'Detalle
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idlgd") = xId
        
        If NulosN(Mid(Fg1.TextMatrix(A, 1), 1, 1)) = 0 Then
            RstDet("tipreg") = 0
        Else
            RstDet("tipreg") = -1
        End If
        
        RstDet("idmod") = NulosN(Fg1.TextMatrix(A, 14))
        
        'SI ES POR SELECCION
        
        If NulosN(Mid(Fg1.TextMatrix(A, 1), 1, 1)) = 0 Then
            
            RstDet("iddoc") = NulosN(Fg1.TextMatrix(A, 15))
            RstDet("tipdoc") = NulosN(Fg1.TextMatrix(A, 16))
            RstDet("idtipper") = 0
            RstDet("idper") = 0
            RstDet("fchdoc") = Null
            RstDet("numser") = Null
            RstDet("numdoc") = Null
            RstDet("idmon") = 0
            
            If NulosN(Me.TxtIdMon) = 1 Then
                RstDet("imptot") = NulosN(Fg1.TextMatrix(A, 10))
            Else
                RstDet("imptot") = NulosN(Fg1.TextMatrix(A, 11))
            End If
            RstDet("impsal") = NulosN(Fg1.TextMatrix(A, 13))
            RstDet("impacue") = NulosN(Fg1.TextMatrix(A, 12))
            
        Else
                        
            RstDet("iddoc") = 0
            RstDet("tipdoc") = NulosN(Fg1.TextMatrix(A, 16))
            
            If NulosN(Fg1.TextMatrix(A, 14)) = 2 Or NulosN(Fg1.TextMatrix(A, 14)) = 5 Then
                RstDet("idtipper") = 2
            Else
                RstDet("idtipper") = 1
            End If
            
            RstDet("idper") = NulosN(Fg1.TextMatrix(A, 18))
            If IsDate(Fg1.TextMatrix(A, 7)) = True Then RstDet("fchdoc") = CDate(Fg1.TextMatrix(A, 7))
            
            '--obteniendo el num de documento
            RstDet("numser") = NulosC(Mid(Fg1.TextMatrix(A, 6), 1, 4))
            RstDet("numdoc") = NulosC(Mid(Fg1.TextMatrix(A, 6), 6, 10))
            
            RstDet("idmon") = NulosN(Fg1.TextMatrix(A, 17))
                        
            If NulosN(Me.TxtIdMon) = 1 Then
                RstDet("imptot") = NulosN(Fg1.TextMatrix(A, 10))
            Else
                RstDet("imptot") = NulosN(Fg1.TextMatrix(A, 11))
            End If
            RstDet("impsal") = NulosN(Fg1.TextMatrix(A, 13))
            RstDet("impacue") = NulosN(Fg1.TextMatrix(A, 12))
            
        End If
    
        RstDet("glosa") = NulosC(txtglosa)
        
        RstDet("idtipdocref") = NulosN(TxtIdTipDoc)
        RstDet("iddocref2") = NulosN(LblIdDocRef2)
        RstDet("numerodocref") = Null
        
        '--cuenta contable
        RstDet("idcuen") = NulosN(Fg1.TextMatrix(A, 19))
        
        RstDet.Update
        
    Next A
   
    '--generando el asiento contable
    xNumAsiento = GenerarAsiento(xCon, 41, xId, AnoTra, mMesActivo, 1, 0)
    If xNumAsiento = "" Then GoTo LaCague
    
    'Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
   
    xCon.CommitTrans
    
    MsgBox "La " & Trim(LblNomDoc) & " se registró con éxito" & vbCr & "Registro Nº: " & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Set RstCab = Nothing
    Set RstDet = Nothing
    'Set RstDia = Nothing
    Grabar = True
    Exit Function
    
LaCague:
'    Resume
    xCon.RollbackTrans
''    Set rstdocus = Nothing
    Set RstCab = Nothing
    Set RstDet = Nothing
''    Set RstDia = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Function HallaNumAsiento(Mes As Integer) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_diario.idmes, con_diario.idlib, con_diario.numasi From con_diario " _
        & " WHERE (((con_diario.idmes)=" & Mes & ") AND ((con_diario.idlib)=41)) ORDER BY numasi", xCon
    
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
       
        
    
    If NulosN(TxtTipDoc.Text) <> 120 And NulosN(TxtTipDoc.Text) <> 126 Then
            MsgBox "Registrar Liquidación Gasto Débito ó Liquidación Gasto Crédito ", vbInformation, Me.Caption
            TxtTipDoc.Text = ""
            LblNomDoc.Caption = ""
            TxtTipDoc.SetFocus
            Exit Sub
    End If
    
    
    
    RST_Busq xRs, "SELECT mae_documento.* FROM MAE_documento WHERE id = " & NulosN(Me.TxtTipDoc) & "", xCon
        
    If xRs.RecordCount = 0 Then
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
    Else
        TxtTipDoc.Text = xRs("id")
        LblNomDoc.Caption = xRs("descripcion")
        
        
        Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & NulosN(TxtIdMon) & " and tipope = -1", xCon)
        If xRs2.RecordCount > 0 Then
            xCuentaDoc = NulosN(xRs2("idcuen"))
        End If
        
        Set xRs2 = Nothing
               
    End If
    
    'Buscamos para hallar el numero de serie asignado al almacen
    
    If TxtTipDoc.Text <> "" Then
        Dim Rst As New ADODB.Recordset
        Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(TxtIdAlm) & "", xCon)
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
    xCampos(7, 0) = "importe":         xCampos(7, 1) = "imptot":     xCampos(7, 2) = "N":         xCampos(7, 3) = "1500"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstVent
    Set RstVent = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstVent
    Dg1.Refresh
End Sub



Sub Imprimir()
    
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
    
   If mMesActivo < 13 Then
            
        nSQL = "SELECT vta_gastodebito.*, IIf(vta_gastodebito.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, IIf(IsNull([vta_gastodebito]![numser])=1,[vta_gastodebito]![numdoc],[vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc]) AS numerodoc, " _
             & " mae_documento.abrev, mae_cliente.numruc, mae_documento.descripcion AS nomdoc, mae_moneda.descripcion AS descmon, IIf(vta_gastodebito.anulado=-1,'',mae_moneda.simbolo) AS simbolo, con_tc.impven,Mid([vta_gastodebito].[numreg],1,2)+[mae_libros].[codsun]+Mid([vta_gastodebito].[numreg],3,4) AS numreg1, " _
             & " vta_gastodebito.fchemi & '' as fchdoc1 ,vta_gastodebito.imptot & '' as imptot1, vta_gastodebito.impsal & '' as impsal1, " _
             & " iif(vta_gastodebito.anulado=-1,0,IIf([vta_gastodebito].[tc]=0,[con_tc].[impven],[vta_gastodebito].[tc])) & '' AS impven1 " _
             & " FROM (con_tc RIGHT JOIN (((vta_gastodebito LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id) ON con_tc.fecha = vta_gastodebito.fchemi) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id " _
             & " WHERE vta_gastodebito.idmes=" & mMesActivo & " " _
             & " ORDER BY vta_gastodebito!numser+'-'+ vta_gastodebito!numdoc DESC"
            
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
        Rpta = MsgBox("No se ha registrado ninguna operacion, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
        If Rpta = vbYes Then
            Nuevo

        End If
    End If

    

    
    
    
End Sub

Sub Buscar()
    TabOne1.CurrTab = 0
    Dim xRs As New ADODB.Recordset
    
    Dim nSQL As String
    Dim xCampos(7, 4) As String
    
    xCampos(0, 0) = "NumReg":        xCampos(0, 1) = "registro":   xCampos(0, 2) = "820":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":      xCampos(1, 2) = "400":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "N°. Documento": xCampos(2, 1) = "numerodoc":  xCampos(2, 2) = "1400":  xCampos(2, 3) = "C"
    xCampos(3, 0) = "FchEmi":        xCampos(3, 1) = "fchemi":     xCampos(3, 2) = "830":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "Cliente":       xCampos(4, 1) = "nombre":     xCampos(4, 2) = "2600":  xCampos(4, 3) = "C"
    xCampos(5, 0) = "M":             xCampos(5, 1) = "simbolo":    xCampos(5, 2) = "450":   xCampos(5, 3) = "C"
    xCampos(6, 0) = "Importe":       xCampos(6, 1) = "imptot":     xCampos(6, 2) = "850":   xCampos(6, 3) = "N"
    
    
    
    'nSQL = " SELECT vta_gastodebito.id,Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4) AS registro, IIf(vta_gastodebito.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, vta_gastodebito!numser+'-'+vta_gastodebito!numdoc AS numerodoc, mae_documento.abrev, IIf(vta_gastodebito.anulado=-1,'',mae_moneda.simbolo) AS simbolo, format(vta_gastodebito.fchemi,'dd/mm/yy') as fchemi,  vta_gastodebito.imptot " & _
    '       " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN vta_gastodebito ON mae_documento.id = vta_gastodebito.tipdoc) ON mae_moneda.id = vta_gastodebito.idmon) ON mae_cliente.id = vta_gastodebito.idcli) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id "
        
    nSQL = " SELECT vta_gastodebito.id, Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4) AS registro, IIf(vta_gastodebito.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, vta_gastodebito!numser+'-'+vta_gastodebito!numdoc AS numerodoc, mae_documento.abrev, IIf(vta_gastodebito.anulado=-1,'',mae_moneda.simbolo) AS simbolo, format(vta_gastodebito.fchemi,'dd/mm/yy') as fchemi,  vta_gastodebito.imptot " _
        + vbCr + " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN vta_gastodebito ON mae_documento.id = vta_gastodebito.tipdoc) ON mae_moneda.id = vta_gastodebito.idmon) ON mae_cliente.id = vta_gastodebito.idcli) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id " _
        + vbCr + " WHERE (((vta_gastodebito.numreg) Like '" & Format(mMesActivo, "00") & "%')) " _


    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Gastos", "nombre", "nombre", Principio

    If xRs.State = 1 Then
        RstVent.MoveFirst
        RstVent.Find "id = " & xRs("id") & ""
    End If
    
    Set xRs = Nothing
End Sub




Sub ActualizaSaldoDoc(idDocumento As Double, Tabla As Integer, ImporteRestar As Double)
'    '1 = compras
'    '2 = Ventas
'    '3 = honorarios
'
'    Dim Rst As New ADODB.Recordset
'    Dim Total As Double
'
'    If Tabla = 2 Then
'        RST_Busq Rst, "SELECT Sum(tes_cajadestinodet.acuenta) AS total FROM tes_caja LEFT JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes " _
'            & " GROUP BY tes_cajadestinodet.iddoc, tes_caja.tipmov HAVING (((tes_cajadestinodet.iddoc)=" & idDocumento & ") AND ((tes_caja.tipmov)=1))", xCon
'
'        Total = BuscaImporteDocumento(idDocumento, 1)
'    End If
'
'    If Rst.RecordCount <> 0 Then
'        Total = ((Total - Rst("total")) - ImporteRestar)
'    Else
'        Total = (Total - ImporteRestar)
'    End If
'
'    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & Total & " WHERE (((vta_ventas.id)=" & idDocumento & "))"
'    Set Rst = Nothing
End Sub


Function BuscaImporteDocumento(idDocumento As Integer, Tabla As Integer) As Double
    
    Dim Rst As New ADODB.Recordset
    
    
    If Tabla = 1 Then RST_Busq Rst, "SELECT * FROM vta_gastodebito WHERE id = " & idDocumento & "", xCon
    
    If Rst.RecordCount <> 0 Then
        BuscaImporteDocumento = Rst("imptot")
    Else
        BuscaImporteDocumento = 0
    End If
    
    Set Rst = Nothing
End Function

Private Sub pGridConfigurar()
    
        
        Fg1.ColWidth(2) = 1000
        Fg1.ColWidth(3) = 1700
        Fg1.ColWidth(5) = 450
        Fg1.ColWidth(6) = 1200
        Fg1.ColWidth(7) = 1050
        Fg1.ColWidth(8) = 450
        
        Fg1.ColWidth(10) = 1000
        Fg1.ColWidth(11) = 1000
        
        Fg1.ColWidth(12) = 0 'idmod
        Fg1.ColWidth(13) = 0 'idcli
        Fg1.ColWidth(14) = 0 'tipdoc
        Fg1.ColWidth(15) = 0 'idmon
        Fg1.ColWidth(16) = 0 'iddoc
        Fg1.ColWidth(17) = 0 'iddoc
    
End Sub


Private Sub ChkTC_Click()
    If QueHace = 3 Then Exit Sub
    
    If ChkTC.Value = 0 Then
        TxtTC.BackColor = &H8000000F
        TxtTC.Enabled = False
        If IsDate(TxtFchDoc.Valor) = True Then
            TxtTC.Text = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
        Else

            Exit Sub
        End If
    Else
        TxtTC.Enabled = True
        TxtTC.BackColor = vbWhite
        TxtTC.SetFocus
    End If
End Sub


Private Sub pHallaCtaDetalle(TipDoc As Integer, IdMoneda As Integer, xFila As Long)
    '************************************
    Dim xRs1 As New ADODB.Recordset
    '--cuenta contable
    RST_Busq xRs1, "SELECT mae_documentolgdcta.idcuen, con_planctas.cuenta, con_planctas.descripcion " _
    & " FROM mae_documentolgdcta INNER JOIN con_planctas ON mae_documentolgdcta.idcuen = con_planctas.id " _
    & " WHERE mae_documentolgdcta.iddoc =" & NulosN(TxtTipDoc.Text) & " AND mae_documentolgdcta.iddocref = " & TipDoc & " AND mae_documentolgdcta.idmon =" & IdMoneda & "", xCon
    If xRs1.RecordCount <> 0 Then
        Fg1.TextMatrix(xFila, 19) = NulosN(xRs1!idcuen)
        Fg1.TextMatrix(xFila, 21) = NulosC(xRs1!cuenta)
        Fg1.TextMatrix(xFila, 22) = NulosC(xRs1!Descripcion)
    End If
    Set xRs1 = Nothing
    '************************************

End Sub

Private Sub pExportar()
    TabOne1.CurrTab = 0
    

    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset

    Dim xCampos(9, 3) As String
    
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Nº Reg":       xCampos(0, 1) = "numreg1":      xCampos(0, 2) = 0:   xCampos(0, 3) = "900"
    xCampos(1, 0) = "T.D.":         xCampos(1, 1) = "abrev":        xCampos(1, 2) = 0:   xCampos(1, 3) = "350"
    xCampos(2, 0) = "Num. Doc":     xCampos(2, 1) = "numerodoc":    xCampos(2, 2) = 0:   xCampos(2, 3) = "1600"
    xCampos(3, 0) = "Fch.Emi":      xCampos(3, 1) = "fchdoc1":      xCampos(3, 2) = 1:   xCampos(3, 3) = "900"
    xCampos(4, 0) = "R.U.C.":       xCampos(4, 1) = "numruc":       xCampos(4, 2) = 0:   xCampos(4, 3) = "1200"
    xCampos(5, 0) = "Cliente":      xCampos(5, 1) = "nombre":       xCampos(5, 2) = 0:   xCampos(5, 3) = "3290"
    xCampos(6, 0) = "M":            xCampos(6, 1) = "simbolo":      xCampos(6, 2) = 1:   xCampos(6, 3) = "500"
    xCampos(7, 0) = "T.C.":         xCampos(7, 1) = "impven1":      xCampos(7, 2) = 2:   xCampos(7, 3) = "700"
    xCampos(8, 0) = "Imp Total":    xCampos(8, 1) = "imptot":       xCampos(8, 2) = 2:   xCampos(8, 3) = "900"
    xCampos(9, 0) = "Imp Saldo":    xCampos(9, 1) = "impsal":       xCampos(9, 2) = 2:   xCampos(9, 3) = "1000"

    Set RstTmp = RstVent.Clone
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "LISTADO DE LGD/LGC", "Periodo " & LblMes.Caption, "", "Listado de Lgd/Lgc - " & LblMes.Caption, RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
    
End Sub




Private Sub CmdApertura_Click()
    '===================================================================================================
    'Creado : 05/02/10 Por: Johan Castro
    'Propósito:  Seleccionar los registros del ejercicio anterior y agregarlos como apertura de ejercicio actual
    '
    'Entradas:   Ninguna
    '
    'Resultados: Registros que seleccionan seran agregados como apertura
    '
    'Nota:       1.- Buscar el documento
    '            2.- Activar con el check
    '            3.- Repetir pasos 1,2 tantas veces documentos tenga
    '            4.- Hacer clic en boton aceptar(se graba automaticamente en su tabla correspondiente)
    'Mofificado: 24/01/12 Johan Castro
    '            Enviar parametro IdMenuActivo
    '===================================================================================================
    AperturaDocumento xCon, xIdUsuario, 41, IdMenuActivo
    RstVent.Filter = ""
    TDB_FiltroLimpiar Dg1
    RstVent.Requery
    
End Sub





Private Sub Command3_Click()
    '--grabar analisis cuenta cte savar
    Dim A As Integer
    TabOne1.CurrTab = 0
    RstVent.MoveFirst
''    Dim xCodSunLib  As String
''    Dim xTc As Double
''    xCodSunLib = Busca_Codigo(41, "id", "codsun", "mae_libros", "N", xCon)
    For A = 1 To RstVent.RecordCount
        
''        If NulosN(RstVent("tc")) = 0 Then
''            xTc = HallaTipoCambio(RstVent("fchemi"), 2, Venta, xCon)
''        Else
''            xTc = RstVent("tc")
''        End If
'
''        If RstVent("tipdoc") = 120 Then   ' LIQUIDACION GASTO DEBITO
''            If NulosN(RstVent("idmon")) = 1 Then
''                GrabarOperacionCtaCteDocRef 41, RstVent("id"), NulosC(RstVent("numerodocref")), RstVent("idcli"), RstVent("tipdoc"), RstVent("numerodoc"), _
''                    RstVent("fchemi"), RstVent("idmon"), xTc, 0, RstVent("imptot"), 0, 0, Format(xCodSunLib, "00") & RstVent("numreg"), xCon
''            Else
''                GrabarOperacionCtaCteDocRef 41, RstVent("id"), NulosC(RstVent("numerodocref")), RstVent("idcli"), RstVent("tipdoc"), RstVent("numerodoc"), _
''                    RstVent("fchemi"), RstVent("idmon"), xTc, 0, 0, 0, RstVent("imptot"), Format(xCodSunLib, "00") & RstVent("numreg"), xCon
''            End If
''        End If
''
''        If RstVent("tipdoc") = 126 Then   ' LIQUIDACION GASTO CREDITO
''            If NulosN(RstVent("idmon")) = 1 Then
''                GrabarOperacionCtaCteDocRef 41, RstVent("id"), NulosC(RstVent("numerodocref")), RstVent("idcli"), RstVent("tipdoc"), RstVent("numerodoc"), _
''                    RstVent("fchemi"), RstVent("idmon"), xTc, RstVent("imptot"), 0, 0, 0, Format(xCodSunLib, "00") & RstVent("numreg"), xCon
''            Else
''                GrabarOperacionCtaCteDocRef 41, RstVent("id"), NulosC(RstVent("numerodocref")), RstVent("idcli"), RstVent("tipdoc"), RstVent("numerodoc"), _
''                    RstVent("fchemi"), RstVent("idmon"), xTc, 0, 0, RstVent("imptot"), 0, Format(xCodSunLib, "00") & RstVent("numreg"), xCon
''            End If
''        End If
        
        GrabarOperacionCtaCte 41, RstVent("id"), xCon

        
        RstVent.MoveNext
        If RstVent.EOF = True Then Exit For
    Next A
    MsgBox "se termino con exito"
End Sub


