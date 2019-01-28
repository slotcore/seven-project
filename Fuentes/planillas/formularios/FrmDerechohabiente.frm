VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmDerechohabiente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Nómina del Personal - Derechohabiente"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   405
      Index           =   12
      Left            =   75
      TabIndex        =   99
      Top             =   3810
      Width           =   5355
      Begin VB.Label lbl_persona 
         AutoSize        =   -1  'True
         Caption         =   "lbl_persona"
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
         Left            =   90
         TabIndex        =   107
         Top             =   90
         Width           =   990
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   5340
         X2              =   5340
         Y1              =   15
         Y2              =   395
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   6380
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   -15
         X2              =   6395
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
      Height          =   420
      Index           =   1
      Left            =   7290
      TabIndex        =   98
      Top             =   3795
      Width           =   1650
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Grabar"
      Height          =   420
      Index           =   0
      Left            =   5505
      TabIndex        =   97
      Top             =   3795
      Width           =   1650
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   3600
      Left            =   45
      TabIndex        =   27
      Top             =   105
      Width           =   8895
      _cx             =   15690
      _cy             =   6350
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
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   12632256
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   " Datos Personales  | Vínculo Familiar |     Domicilio     "
      Align           =   0
      CurrTab         =   2
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
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Height          =   3180
         Index           =   2
         Left            =   45
         TabIndex        =   59
         Top             =   45
         Width           =   8805
         Begin VB.Frame fra 
            Caption         =   "[ Ubigeo ]"
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
            Height          =   1290
            Index           =   11
            Left            =   105
            TabIndex        =   74
            Top             =   1845
            Width           =   6195
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   10
               Left            =   1890
               Picture         =   "FrmDerechohabiente.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   77
               ToolTipText     =   "Seleccione la Provincia"
               Top             =   580
               Width           =   210
            End
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   9
               Left            =   1890
               Picture         =   "FrmDerechohabiente.frx":0132
               Style           =   1  'Graphical
               TabIndex        =   76
               ToolTipText     =   "Seleccione el Departamento"
               Top             =   250
               Width           =   210
            End
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   11
               Left            =   1890
               Picture         =   "FrmDerechohabiente.frx":0264
               Style           =   1  'Graphical
               TabIndex        =   75
               ToolTipText     =   "Seleccione el Distrito"
               Top             =   910
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   9
               Left            =   1365
               MaxLength       =   20
               TabIndex        =   24
               Tag             =   "null"
               Text            =   "txt_cb(9)"
               Top             =   225
               Width           =   765
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   10
               Left            =   1365
               MaxLength       =   20
               TabIndex        =   25
               Tag             =   "null"
               Text            =   "txt_cb(10)"
               Top             =   555
               Width           =   765
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   11
               Left            =   1365
               MaxLength       =   20
               TabIndex        =   26
               Tag             =   "null"
               Text            =   "txt_cb(11)"
               Top             =   885
               Width           =   765
            End
            Begin VB.Label lbl_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cod(11)"
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
               Index           =   11
               Left            =   3900
               TabIndex        =   86
               Top             =   885
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cod(10)"
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
               Index           =   10
               Left            =   3900
               TabIndex        =   85
               Top             =   555
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cod(9)"
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
               Index           =   9
               Left            =   3900
               TabIndex        =   84
               Top             =   225
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(11)"
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
               Index           =   11
               Left            =   2130
               TabIndex        =   83
               Top             =   885
               Width           =   3075
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(10)"
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
               Index           =   10
               Left            =   2130
               TabIndex        =   82
               Top             =   555
               Width           =   3075
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(9)"
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
               Index           =   9
               Left            =   2130
               TabIndex        =   81
               Top             =   225
               Width           =   3075
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Distrito"
               Height          =   195
               Index           =   11
               Left            =   150
               TabIndex        =   80
               Top             =   960
               Width           =   480
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Provincia"
               Height          =   195
               Index           =   10
               Left            =   150
               TabIndex        =   79
               Top             =   645
               Width           =   660
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Departamento"
               Height          =   195
               Index           =   9
               Left            =   150
               TabIndex        =   78
               Top             =   300
               Width           =   1005
            End
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   11
            Left            =   1470
            MaxLength       =   40
            TabIndex        =   23
            Tag             =   "null"
            Text            =   "txt(11)"
            Top             =   1500
            Width           =   6555
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   10
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   22
            Tag             =   "null"
            Text            =   "txt(10)"
            Top             =   1185
            Width           =   3180
         End
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   8
            Left            =   1995
            Picture         =   "FrmDerechohabiente.frx":0396
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Seleccione el Tipo de Zona"
            Top             =   885
            Width           =   210
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   7035
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "null"
            Text            =   "txt(9)"
            Top             =   540
            Width           =   1035
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   8
            Left            =   5385
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "null"
            Text            =   "txt(8)"
            Top             =   540
            Width           =   1035
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   7
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   18
            Tag             =   "null"
            Text            =   "txt(7)"
            Top             =   540
            Width           =   3180
         End
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   7
            Left            =   1995
            Picture         =   "FrmDerechohabiente.frx":04C8
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Seleccione el Tipo de Vía"
            Top             =   240
            Width           =   210
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   7
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   17
            Tag             =   "null"
            Text            =   "txt_cb(7)"
            Top             =   210
            Width           =   765
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   8
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   21
            Tag             =   "null"
            Text            =   "txt_cb(8)"
            Top             =   855
            Width           =   765
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cod(8)"
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
            Index           =   8
            Left            =   4005
            TabIndex        =   71
            Top             =   855
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia"
            Height          =   195
            Index           =   11
            Left            =   270
            TabIndex        =   70
            Top             =   1590
            Width           =   780
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre Zona"
            Height          =   195
            Index           =   10
            Left            =   270
            TabIndex        =   69
            Top             =   1260
            Width           =   975
         End
         Begin VB.Label lbl_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Zona"
            Height          =   195
            Index           =   8
            Left            =   270
            TabIndex        =   68
            Top             =   930
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interior"
            Height          =   195
            Index           =   9
            Left            =   6480
            TabIndex        =   67
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
            Height          =   195
            Index           =   8
            Left            =   4740
            TabIndex        =   66
            Top             =   600
            Width           =   555
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre Vía"
            Height          =   195
            Index           =   7
            Left            =   270
            TabIndex        =   65
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cod(7)"
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
            Index           =   7
            Left            =   4005
            TabIndex        =   64
            Top             =   210
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(7)"
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
            Index           =   7
            Left            =   2205
            TabIndex        =   63
            Top             =   210
            Width           =   3075
         End
         Begin VB.Label lbl_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Vía"
            Height          =   195
            Index           =   7
            Left            =   270
            TabIndex        =   62
            Top             =   270
            Width           =   615
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(8)"
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
            Index           =   8
            Left            =   2235
            TabIndex        =   72
            Top             =   855
            Width           =   3075
         End
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   3180
         Index           =   1
         Left            =   -9450
         TabIndex        =   29
         Top             =   45
         Width           =   8805
         Begin VB.Frame fra 
            Caption         =   "[ Situación del Derechohabiente ]"
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
            Height          =   1860
            Index           =   8
            Left            =   150
            TabIndex        =   48
            Top             =   1125
            Width           =   4755
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   5
               Left            =   1770
               Picture         =   "FrmDerechohabiente.frx":05FA
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   1170
               Width           =   210
            End
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   4
               Left            =   1770
               Picture         =   "FrmDerechohabiente.frx":072C
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   300
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   4
               Left            =   1245
               MaxLength       =   20
               TabIndex        =   10
               Tag             =   "null"
               Text            =   "txt_cb(4)"
               Top             =   270
               Width           =   765
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   5
               Left            =   1245
               MaxLength       =   20
               TabIndex        =   12
               Tag             =   "null"
               Text            =   "txt_cb(5)"
               Top             =   1140
               Width           =   765
            End
            Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
               Height          =   300
               Index           =   1
               Left            =   1245
               TabIndex        =   11
               Top             =   585
               Width           =   1290
               _ExtentX        =   2275
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
            Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
               Height          =   300
               Index           =   2
               Left            =   1245
               TabIndex        =   13
               Top             =   1455
               Width           =   1290
               _ExtentX        =   2275
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
            Begin VB.Label lbl_fecha 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de baja"
               Height          =   195
               Index           =   2
               Left            =   105
               TabIndex        =   58
               Top             =   1560
               Width           =   1020
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Motivo de Baja"
               Height          =   195
               Index           =   5
               Left            =   105
               TabIndex        =   57
               Top             =   1215
               Width           =   1065
            End
            Begin VB.Label lbl_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cod(5)"
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
               Index           =   5
               Left            =   3585
               TabIndex        =   56
               Top             =   1140
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_fecha 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de alta"
               Height          =   195
               Index           =   1
               Left            =   105
               TabIndex        =   55
               Top             =   690
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Situación"
               Height          =   195
               Index           =   4
               Left            =   105
               TabIndex        =   54
               Top             =   345
               Width           =   660
            End
            Begin VB.Label lbl_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cod(4)"
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
               Index           =   4
               Left            =   3585
               TabIndex        =   53
               Top             =   270
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(5)"
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
               Index           =   5
               Left            =   1995
               TabIndex        =   52
               Top             =   1140
               Width           =   2640
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(4)"
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
               Index           =   4
               Left            =   1995
               TabIndex        =   51
               Top             =   270
               Width           =   2640
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Vínculo Familiar ]"
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
            Height          =   810
            Index           =   6
            Left            =   150
            TabIndex        =   38
            Top             =   150
            Width           =   8595
            Begin VB.Frame fra 
               Caption         =   "[ Documento que acredita la paternidad ]"
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
               Height          =   615
               Index           =   7
               Left            =   3060
               TabIndex        =   40
               Top             =   135
               Width           =   5490
               Begin VB.TextBox txt 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   5
                  Left            =   3795
                  MaxLength       =   20
                  TabIndex        =   9
                  Tag             =   "null"
                  Text            =   "txt(5)"
                  Top             =   210
                  Width           =   1590
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   3
                  Left            =   945
                  Picture         =   "FrmDerechohabiente.frx":085E
                  Style           =   1  'Graphical
                  TabIndex        =   41
                  Top             =   240
                  Width           =   210
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   3
                  Left            =   420
                  MaxLength       =   20
                  TabIndex        =   8
                  Tag             =   "null"
                  Text            =   "txt_cb(3)"
                  Top             =   210
                  Width           =   765
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Número"
                  Height          =   195
                  Index           =   5
                  Left            =   3180
                  TabIndex        =   106
                  Top             =   315
                  Width           =   555
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo"
                  Height          =   195
                  Index           =   3
                  Left            =   75
                  TabIndex        =   44
                  Top             =   315
                  Width           =   315
               End
               Begin VB.Label lbl_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cod(3)"
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
                  Index           =   3
                  Left            =   2175
                  TabIndex        =   43
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(3)"
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
                  Index           =   3
                  Left            =   1185
                  TabIndex        =   42
                  Top             =   210
                  Width           =   1920
               End
            End
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   2
               Left            =   1200
               Picture         =   "FrmDerechohabiente.frx":0990
               Style           =   1  'Graphical
               TabIndex        =   39
               ToolTipText     =   "Seleccione el Vínculo Familiar"
               Top             =   435
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   2
               Left            =   690
               MaxLength       =   20
               TabIndex        =   7
               Tag             =   "null"
               Text            =   "txt_cb(2)"
               Top             =   405
               Width           =   765
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Código"
               Height          =   195
               Left            =   120
               TabIndex        =   73
               Top             =   495
               Width           =   495
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vínculo Familiar"
               Height          =   195
               Index           =   2
               Left            =   1875
               TabIndex        =   47
               Top             =   195
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label lbl_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cod(2)"
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
               Index           =   2
               Left            =   1575
               TabIndex        =   45
               Top             =   660
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(2)"
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
               Index           =   2
               Left            =   1455
               TabIndex        =   46
               Top             =   405
               Width           =   1575
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Incapacidad ]"
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
            Height          =   915
            Index           =   9
            Left            =   4965
            TabIndex        =   35
            Top             =   1125
            Width           =   3780
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   6
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   15
               Tag             =   "null"
               Text            =   "txt(6)"
               ToolTipText     =   "Vínculo = Hijo y la cantidad de años entre la fecha de nacimiento y el presente es mayor a 18 años."
               Top             =   450
               Width           =   1590
            End
            Begin VB.OptionButton opt_incapacidad 
               Caption         =   "No"
               Height          =   210
               Index           =   0
               Left            =   330
               TabIndex        =   14
               Top             =   255
               Width           =   645
            End
            Begin VB.OptionButton opt_incapacidad 
               Caption         =   "Si"
               Height          =   210
               Index           =   1
               Left            =   330
               TabIndex        =   36
               Top             =   570
               Width           =   555
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Número"
               Height          =   195
               Index           =   6
               Left            =   1395
               TabIndex        =   37
               Top             =   540
               Width           =   555
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Domicilio del Derechohabiente ]"
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
            Index           =   10
            Left            =   4980
            TabIndex        =   30
            Top             =   2160
            Width           =   3780
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   6
               Left            =   675
               Picture         =   "FrmDerechohabiente.frx":0AC2
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   435
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   6
               Left            =   150
               MaxLength       =   20
               TabIndex        =   16
               Tag             =   "null"
               Text            =   "txt_cb(6)"
               Top             =   405
               Width           =   765
            End
            Begin VB.Label lbl_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cod(6)"
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
               Index           =   6
               Left            =   2370
               TabIndex        =   33
               Top             =   405
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Domicilio del Derechohabiente"
               Height          =   195
               Index           =   6
               Left            =   1260
               TabIndex        =   32
               Top             =   225
               Visible         =   0   'False
               Width           =   2160
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(6)"
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
               Index           =   6
               Left            =   900
               TabIndex        =   34
               Top             =   405
               Width           =   2700
            End
         End
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3180
         Index           =   0
         Left            =   -9750
         TabIndex        =   28
         Top             =   45
         Width           =   8805
         Begin VB.Frame fra 
            Height          =   690
            Index           =   5
            Left            =   150
            TabIndex        =   100
            Top             =   2265
            Width           =   8175
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   1
               Left            =   4965
               Picture         =   "FrmDerechohabiente.frx":0BF4
               Style           =   1  'Graphical
               TabIndex        =   101
               ToolTipText     =   "Seleccione el Sexo"
               Top             =   235
               Width           =   210
            End
            Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
               Height          =   300
               Index           =   0
               Left            =   1140
               TabIndex        =   5
               ToolTipText     =   "Seleccione la Fecha de Nacimiento"
               Top             =   210
               Width           =   1425
               _ExtentX        =   2514
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
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   1
               Left            =   4440
               MaxLength       =   20
               TabIndex        =   6
               Text            =   "txt_cb(1)"
               ToolTipText     =   "Ingrese el Sexo (1:Masculino, 2:Femenino)"
               Top             =   210
               Width           =   765
            End
            Begin VB.Label lbl_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cod(1)"
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
               Index           =   1
               Left            =   6330
               TabIndex        =   104
               Top             =   210
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sexo"
               Height          =   195
               Index           =   1
               Left            =   4020
               TabIndex        =   103
               Top             =   315
               Width           =   360
            End
            Begin VB.Label lbl_fecha 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Nac."
               Height          =   195
               Index           =   0
               Left            =   105
               TabIndex        =   102
               Top             =   315
               Width           =   840
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(1)"
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
               Index           =   1
               Left            =   5205
               TabIndex        =   105
               Top             =   210
               Width           =   2325
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Identificación del Derechohabiente]"
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
            Index           =   3
            Left            =   150
            TabIndex        =   93
            Top             =   135
            Width           =   8175
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   1
               Left            =   1155
               MaxLength       =   40
               TabIndex        =   0
               Text            =   "txt(1)"
               Top             =   255
               Width           =   2715
            End
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   2
               Left            =   5010
               MaxLength       =   40
               TabIndex        =   1
               Text            =   "txt(2)"
               Top             =   255
               Width           =   2715
            End
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   3
               Left            =   1155
               MaxLength       =   40
               TabIndex        =   2
               Text            =   "txt(3)"
               Top             =   570
               Width           =   6570
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ap. Paterno"
               Height          =   195
               Index           =   1
               Left            =   105
               TabIndex        =   96
               Top             =   345
               Width           =   840
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ap. Materno"
               Height          =   195
               Index           =   2
               Left            =   4020
               TabIndex        =   95
               Top             =   345
               Width           =   870
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nombres"
               Height          =   195
               Index           =   3
               Left            =   105
               TabIndex        =   94
               Top             =   675
               Width           =   630
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Documento de Identificación ]"
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
            Index           =   4
            Left            =   150
            TabIndex        =   87
            Top             =   1215
            Width           =   8175
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   4
               Left            =   1140
               MaxLength       =   15
               TabIndex        =   4
               Text            =   "txt(4)"
               Top             =   585
               Width           =   1485
            End
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   0
               Left            =   1665
               Picture         =   "FrmDerechohabiente.frx":0D26
               Style           =   1  'Graphical
               TabIndex        =   88
               ToolTipText     =   "Seleccione el Tipo de Documento"
               Top             =   280
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   0
               Left            =   1140
               MaxLength       =   20
               TabIndex        =   3
               Text            =   "txt_cb(0)"
               Top             =   255
               Width           =   765
            End
            Begin VB.Label lbl_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cod(0)"
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
               Index           =   0
               Left            =   3660
               TabIndex        =   92
               Top             =   255
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Doc."
               Height          =   195
               Index           =   0
               Left            =   105
               TabIndex        =   91
               Top             =   360
               Width           =   705
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Número Doc."
               Height          =   195
               Index           =   4
               Left            =   105
               TabIndex        =   90
               Top             =   690
               Width           =   945
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(0)"
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
               Index           =   0
               Left            =   1905
               TabIndex        =   89
               Top             =   255
               Width           =   3075
            End
         End
      End
   End
End
Attribute VB_Name = "FrmDerechohabiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Quehace As Integer
Dim mCorrelativo As Long
Dim mIdEmpleado As Long
Dim Agregando As Boolean

Public Sub pRecibeLink(QueHace1 As Integer, Optional mCorr As Long)
    Quehace = QueHace1
    mCorrelativo = mCorr
    With FrmNomina
        mIdEmpleado = .txt(0).Text
        lbl_persona.Caption = .lbl_persona(0).Caption
    End With
    '------
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cb
    LimpiaText lbl_cod
    LimpiaText txtfecha
    Agregando = True
    TabOne1.CurrTab = 0
    If Quehace = 2 Then pPonerDatos
End Sub


'*******************************************************************************************

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--GRABAR
            If Grabar() = False Then Exit Sub
            FrmNomina.pCargarDatosDerechoHabiente
            Unload Me
        Case 1 '--CANCELAR
            Unload Me
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CentrarFrm Me
End Sub



Private Sub opt_incapacidad_Click(Index As Integer)
    If Index = 0 Then
        txt(6).Text = ""
        txt(6).Enabled = False
    Else
        txt(6).Enabled = True
    End If
End Sub


Private Sub TabOne1_Click()
    If TabOne1.CurrTab = 0 Then
        If Agregando = False Then txt(1).SetFocus
    ElseIf TabOne1.CurrTab = 1 Then
        If Agregando = False Then txt_cb(2).SetFocus
    Else
        If Agregando = False Then txt_cb(7).SetFocus
    End If
End Sub

Private Sub cb_Click(Index As Integer)
    If Quehace = 3 Then Exit Sub
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nOrden As String
    Dim nCampoBusca As String
    Dim nSQL As String
    On Error GoTo error
    Select Case Index

        Case 0 '--DOCUMENTO DE IDENTIDAD
            nTitulo = "Documento de Identidad"
            nSQL = "SELECT mae_dociden.id, mae_dociden.descripcion as nombre, mae_dociden.id AS cod " _
                + vbCr + " From mae_dociden " _
                + vbCr + " ORDER BY mae_dociden.descripcion;"
        
        Case 1 '--SEXO
            nTitulo = "Buscando Sexo"
            nSQL = "SELECT mae_sexo.id, mae_sexo.descripcion as nombre , mae_sexo.id AS cod " _
                + vbCr + " From mae_sexo " _
                + vbCr + " ORDER BY mae_sexo.descripcion;"
        
        Case 2 '--VINCULO FAMILIAR
            nTitulo = "Buscando Vínculo Familiar"
            nSQL = "SELECT mae_vinculofam.id, mae_vinculofam.descripcion AS nombre, mae_vinculofam.id AS cod " _
                + vbCr + " From mae_vinculofam " _
                + vbCr + " ORDER BY mae_vinculofam.codsun; "
                
        Case 3 '--DOCUMENTO QUE ACREDITA LA PATERNIDAD
            nTitulo = "Buscando Documento que acredita la Paternidad"
            nSQL = "SELECT mae_docacrepat.id, mae_docacrepat.descripcion AS nombre, mae_docacrepat.id AS cod " _
                + vbCr + " From mae_docacrepat " _
                + vbCr + " ORDER BY mae_docacrepat.codsun;"
        
        Case 4 '--SITUACION DERECHO HABIENTE
            nTitulo = "Buscando Situación de Derechohabiente"
            nSQL = "SELECT mae_situacionderhab.id, mae_situacionderhab.descripcion AS nombre, mae_situacionderhab.id AS cod " _
                + vbCr + " From mae_situacionderhab " _
                + vbCr + " ORDER BY mae_situacionderhab.codsun; "
        
        Case 5 '--MOTIVO DE BAJA
            If NulosN(txt_cb(4).Text) = 0 Then
                MsgBox "Falta especificar la situación de derechohabiente", vbExclamation, xTitulo
                txt_cb(4).SetFocus
                Exit Sub
            End If
            nTitulo = "Buscando Motivo de Baja"
            nSQL = "SELECT mae_tipobaja.id, mae_tipobaja.descripcion AS nombre, mae_tipobaja.id AS cod " _
                + vbCr + " From mae_tipobaja " _
                + vbCr + " ORDER BY mae_tipobaja.codsun;"

        Case 6 '--DOMICILIADO
            nTitulo = "Buscando Si es Domiciliado"
            nSQL = "SELECT mae_indicadomderhab.id, mae_indicadomderhab.descripcion AS nombre, mae_indicadomderhab.id AS cod " _
                + vbCr + " From mae_indicadomderhab " _
                + vbCr + " ORDER BY mae_indicadomderhab.codsun; "
                
        Case 7 '--TIPO DE VIA
            nTitulo = "Buscando Tipo de Vía"
            nSQL = "SELECT mae_tipovia.id, mae_tipovia.descripcion AS nombre, mae_tipovia.id AS cod " _
                + vbCr + " From mae_tipovia " _
                + vbCr + " ORDER BY mae_tipovia.descripcion;"
        
        Case 8 '--TIPO ZONA
            nTitulo = "Buscando Tipo de Zona"
            nSQL = "SELECT mae_tipozona.id, mae_tipozona.descripcion AS nombre, mae_tipozona.id AS cod " _
                + vbCr + " From mae_tipozona " _
                + vbCr + " ORDER BY mae_tipozona.descripcion;"

        Case 9 '--DEPARTAMENTO
            nTitulo = "Buscando Departamento"
            nSQL = "SELECT mae_departamento.id, mae_departamento.descripcion as nombre, mae_departamento.id AS cod " _
                + vbCr + " From mae_departamento " _
                + vbCr + " ORDER BY mae_departamento.descripcion;"
        
        Case 10 '--PROVINCIA
            If NulosN(txt_cb(9).Text) = 0 Then
                MsgBox "Falta especificar el Departamento", vbExclamation, xTitulo
                txt_cb(9).SetFocus
                Exit Sub
            End If
            nTitulo = "Buscando Provincia"
            nSQL = "SELECT mae_provincia.id, mae_provincia.descripcion AS nombre, mae_provincia.id AS cod " _
                + vbCr + " From mae_provincia " _
                + vbCr + " Where (((mae_provincia.iddepa) = " & NulosN(txt_cb(9).Text) & " )) " _
                + vbCr + " ORDER BY mae_provincia.descripcion; "

        Case 11 '--DISTRITO
            If NulosN(txt_cb(9).Text) = 0 Then
                MsgBox "Falta especificar el Departamento", vbExclamation, xTitulo
                txt_cb(9).SetFocus
                Exit Sub
            End If
            If NulosN(txt_cb(10).Text) = 0 Then
                MsgBox "Falta especificar la Provincia", vbExclamation, xTitulo
                txt_cb(10).SetFocus
                Exit Sub
            End If
            nTitulo = "Buscando Distrito"
            nSQL = "SELECT mae_distrito.id, mae_distrito.descripcion AS nombre, mae_distrito.id AS cod " _
                + vbCr + " From mae_distrito " _
                + vbCr + " Where (((mae_distrito.idprov) = " & NulosN(txt_cb(10).Text) & ")) " _
                + vbCr + " ORDER BY mae_distrito.descripcion;"
                
    End Select

    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "450":    xCampos(1, 3) = "N"
            
    Dim xRs As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    lbl_cb(Index).ToolTipText = xRs.Fields(1) & "" '--NOMBRE
    '--SI DATO ANTERIOR Y EL ACTUAL SON DIFERENTES => LIMPIAR CAMPOS DEPENDIENTES
    If Trim(lbl_cod(Index).Tag) <> Trim(lbl_cod(Index).Caption) Then
        Select Case Index
            Case 2 '--VINCULO FAMILIAR
                txt_cb(3).Text = ""
            Case 4 '--SITUACION DERECHOHABIENTE
                If NulosN(txt_cb(4).Text) = 1 Then '--ACTIVO
                    txtfecha(2).Valor = ""
                    txt_cb(5).Text = ""
                ElseIf NulosN(txt_cb(4).Text) = 2 Then '--BAJA
                    txtfecha(1).Valor = ""
                End If
                    
            Case 9 '--DEPARTAMENTO
                txt_cb(10).Text = ""
                txt_cb(11).Text = ""
                
            Case 10 '--PROVINCIA
                txt_cb(11).Text = ""
                
        End Select
    End If
    Select Case Index

        Case 0 '--DOCUMENTO DE IDENTIDAD
            txt(4).SetFocus
        Case 1 '--SEXO
            TabOne1.CurrTab = 1
            txt_cb(2).SetFocus
        Case 2 '--VINCULO FAMILIAR
            '----------
            opt_incapacidad(0).Value = True
            habilitar opt_incapacidad, False
            '----------
            If NulosN(txt_cb(2).Text) = 4 Then         '--ES GESTANTE
                txt_cb(3).Enabled = True
                cb(3).Enabled = True
                txt(5).Enabled = True
                txt_cb(3).Tag = ""
                txt_cb(3).SetFocus
            Else
                txt_cb(3).Enabled = False
                cb(3).Enabled = False
                txt(5).Enabled = False
                txt_cb(3).Tag = "null"
                
                If NulosN(txt_cb(2).Text) = 1 Then         '--ES HIJO
                    habilitar opt_incapacidad, True
                    
                End If
                txt_cb(4).SetFocus
            End If
            
        Case 3 '--DOCUMENTO QUE ACREDITA LA PATERNIDAD
            txt(5).SetFocus
        Case 4 '--SITUACION DERECHO HABIENTE
            If NulosN(txt_cb(4).Text) = 1 Then '--ACTIVO
                txtfecha(2).Enabled = False
                txt_cb(5).Enabled = False
                cb(5).Enabled = False
                txtfecha(1).Enabled = True
                txt_cb(5).Tag = "null" '--para saltar la validación en funcion fValidar
                txtfecha(1).SetFocus
            ElseIf NulosN(txt_cb(4).Text) = 2 Then '--BAJA
                txtfecha(1).Enabled = False
                txtfecha(2).Enabled = True
                txt_cb(5).Enabled = True
                cb(5).Enabled = True
                txt_cb(5).Tag = "" '--obligar que se ingrese la información
                txt_cb(5).SetFocus
            End If
        Case 5 '--MOTIVO DE BAJA
            txtfecha(2).SetFocus
        Case 6 '--DOMICILIADO
            TabOne1.CurrTab = 2
            txt_cb(7).SetFocus
        Case 7 '--TIPO DE VIA
            txt(7).SetFocus
        Case 8 '--TIPO ZONA
            txt(10).SetFocus
        Case 9 '--DEPARTAMENTO
            txt_cb(10).SetFocus
        Case 10 '--PROVINCIA
            txt_cb(11).SetFocus
        Case 11 '--DISTRITO
            cmd(0).SetFocus
    End Select
Salir:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If Quehace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
        If Index = 3 Then '--departamento
            txt_cb(4).Text = ""
            txt_cb(5).Text = ""
        ElseIf Index = 4 Then '--provincia
            txt_cb(5).Text = ""
        End If
    End If

End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Quehace = 3 Then Exit Sub
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
   
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If Quehace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index

        Case 0 '--DOCUMENTO DE IDENTIDAD
            nSQL = "SELECT mae_dociden.id, mae_dociden.descripcion as nombre, mae_dociden.id AS cod " _
                + vbCr + " From mae_dociden " _
                + vbCr + " WHERE mae_dociden.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 1 '--SEXO
             nSQL = "SELECT mae_sexo.id, mae_sexo.descripcion as nombre , mae_sexo.id AS cod " _
                + vbCr + " From mae_sexo " _
                + vbCr + " WHERE mae_sexo.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 2 '--VINCULO FAMILIAR
             nSQL = "SELECT mae_vinculofam.id, mae_vinculofam.descripcion AS nombre, mae_vinculofam.id AS cod " _
                + vbCr + " From mae_vinculofam " _
                + vbCr + " WHERE mae_vinculofam.id = " & NulosN(txt_cb(Index).Text) & ";"
                
        Case 3 '--DOCUMENTO QUE ACREDITA LA PATERNIDAD
             nSQL = "SELECT mae_docacrepat.id, mae_docacrepat.descripcion AS nombre, mae_docacrepat.id AS cod " _
                + vbCr + " From mae_docacrepat " _
                + vbCr + " WHERE mae_docacrepat.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 4 '--SITUACION DERECHO HABIENTE
             nSQL = "SELECT mae_situacionderhab.id, mae_situacionderhab.descripcion AS nombre, mae_situacionderhab.id AS cod " _
                + vbCr + " From mae_situacionderhab " _
                + vbCr + " WHERE mae_situacionderhab.id = " & NulosN(txt_cb(Index).Text) & ";"

        Case 5 '--MOTIVO DE BAJA
            If NulosN(txt_cb(4).Text) = 0 Then
                MsgBox "Falta especificar la situación de derechohabiente", vbExclamation, xTitulo
                txt_cb(4).SetFocus
                GoTo Salir
            End If
            nSQL = "SELECT mae_tipobaja.id, mae_tipobaja.descripcion AS nombre, mae_tipobaja.id AS cod " _
                + vbCr + " From mae_tipobaja " _
                + vbCr + " WHERE mae_tipobaja.id = " & NulosN(txt_cb(Index).Text) & ";"

        Case 6 '--DOMICILIADO
             nSQL = "SELECT mae_indicadomderhab.id, mae_indicadomderhab.descripcion AS nombre, mae_indicadomderhab.id AS cod " _
                + vbCr + " From mae_indicadomderhab " _
                + vbCr + " WHERE mae_indicadomderhab.id = " & NulosN(txt_cb(Index).Text) & ";"
                
        Case 7 '--TIPO DE VIA
             nSQL = "SELECT mae_tipovia.id, mae_tipovia.descripcion AS nombre, mae_tipovia.id AS cod " _
                + vbCr + " From mae_tipovia " _
                + vbCr + " WHERE mae_tipovia.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 8 '--TIPO ZONA
             nSQL = "SELECT mae_tipozona.id, mae_tipozona.descripcion AS nombre, mae_tipozona.id AS cod " _
                + vbCr + " From mae_tipozona " _
                + vbCr + " WHERE mae_tipozona.id = " & NulosN(txt_cb(Index).Text) & ";"

        Case 9 '--DEPARTAMENTO
             nSQL = "SELECT mae_departamento.id, mae_departamento.descripcion as nombre, mae_departamento.id AS cod " _
                + vbCr + " From mae_departamento " _
                + vbCr + " WHERE mae_departamento.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 10 '--PROVINCIA
            If NulosN(txt_cb(9).Text) = 0 Then
                MsgBox "Falta especificar el Departamento", vbExclamation, xTitulo
                txt_cb(9).SetFocus
                GoTo Salir
            End If

             nSQL = "SELECT mae_provincia.id, mae_provincia.descripcion AS nombre, mae_provincia.id AS cod " _
                + vbCr + " From mae_provincia " _
                + vbCr + " Where (((mae_provincia.iddepa) = " & NulosN(txt_cb(9).Text) & " )) " _
                + vbCr + " ORDER BY mae_provincia.descripcion; "

        Case 11 '--DISTRITO
            If NulosN(txt_cb(9).Text) = 0 Then
                MsgBox "Falta especificar el Departamento", vbExclamation, xTitulo
                txt_cb(9).SetFocus
                GoTo Salir
            End If
            If NulosN(txt_cb(10).Text) = 0 Then
                MsgBox "Falta especificar la Provincia", vbExclamation, xTitulo
                txt_cb(10).SetFocus
                GoTo Salir
            End If
             nSQL = "SELECT mae_distrito.id, mae_distrito.descripcion AS nombre, mae_distrito.id AS cod " _
                + vbCr + " From mae_distrito " _
                + vbCr + " Where (((mae_distrito.idprov) = " & NulosN(txt_cb(10).Text) & ")) " _
                + vbCr + " ORDER BY mae_distrito.descripcion;"
                
    End Select

    If xCon.State = 0 Then GoTo Salir
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo Salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
        lbl_cb(Index).ToolTipText = RstTmp.Fields(1) & "" '--NOMBRE
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    '--------------
    '--SI DATO ANTERIOR Y EL ACTUAL SON DIFERENTES => LIMPIAR CAMPOS DEPENDIENTES
    If Trim(lbl_cod(Index).Tag) <> Trim(lbl_cod(Index).Caption) Then
        Select Case Index
            Case 2 '--VINCULO FAMILIAR
                txt_cb(3).Text = ""
            Case 4 '--SITUACION DERECHOHABIENTE
                If NulosN(txt_cb(4).Text) = 1 Then '--ACTIVO
                    txtfecha(2).Valor = ""
                    txt_cb(5).Text = ""
                ElseIf NulosN(txt_cb(4).Text) = 2 Then '--BAJA
                    txtfecha(1).Valor = ""
                End If
                    
            Case 9 '--DEPARTAMENTO
                txt_cb(10).Text = ""
                txt_cb(11).Text = ""
                
            Case 10 '--PROVINCIA
                txt_cb(11).Text = ""
                
        End Select
    End If
    Select Case Index

        Case 1 '--SEXO
            TabOne1.CurrTab = 1
            If Agregando = False Then txt_cb(2).SetFocus
        Case 2 '--VINCULO FAMILIAR
            If NulosN(txt_cb(2).Text) = 4 Then         '--ES GESTANTE
                txt_cb(3).Enabled = True
                cb(3).Enabled = True
                txt(5).Enabled = True
                txt_cb(3).Tag = ""
            Else
                txt_cb(3).Enabled = False
                cb(3).Enabled = False
                txt(5).Enabled = False
                txt_cb(3).Tag = "null"
            End If
            
        Case 3 '--DOCUMENTO QUE ACREDITA LA PATERNIDAD
            If Agregando = False Then txt(5).SetFocus
        Case 4 '--SITUACION DERECHO HABIENTE
            If NulosN(txt_cb(4).Text) = 1 Then '--ACTIVO
                txtfecha(2).Enabled = False
                txt_cb(5).Enabled = False
                cb(5).Enabled = False
                txtfecha(1).Enabled = True
                txt_cb(5).Tag = "null" '--para saltar la validación en funcion fValidar
            ElseIf NulosN(txt_cb(4).Text) = 2 Then '--BAJA
                txtfecha(1).Enabled = False
                txtfecha(2).Enabled = True
                txt_cb(5).Enabled = True
                cb(5).Enabled = True
                txt_cb(5).Tag = "" '--obligar que se ingrese la información
            End If

        Case 6 '--DOMICILIADO
            TabOne1.CurrTab = 2
            If Agregando = False Then txt_cb(7).SetFocus
        Case 11 '--DISTRITO
            If Agregando = False Then cmd(0).SetFocus
    End Select
    '--------------
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
Salir:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub


'****************************************************************************************

Private Sub pPonerDatos()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
'    On Error GoTo error
    nSQL = "SELECT pla_derechohab.*, mae_sexo.abrev AS sexo, mae_vinculofam.descripcion AS vinculo, mae_dociden.abrev AS docabrev, [pla_derechohab].[apepat] & ' ' & [pla_derechohab].[apemat] & ' ' & [pla_derechohab].[nombre] AS nombres " _
        + vbCr + " FROM mae_vinculofam RIGHT JOIN (mae_dociden RIGHT JOIN (mae_sexo RIGHT JOIN pla_derechohab ON mae_sexo.id = pla_derechohab.idsex) ON mae_dociden.id = pla_derechohab.idtipdoc) ON mae_vinculofam.id = pla_derechohab.idvinfam " _
        + vbCr + " WHERE (((pla_derechohab.idemp)=" & mIdEmpleado & " ) AND ((pla_derechohab.corr)= " & mCorrelativo & " )) " _
        + vbCr + " ORDER BY [pla_derechohab].[apepat] & ' ' & [pla_derechohab].[apemat] & ' ' & [pla_derechohab].[nombre]; "

    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.RecordCount = 0 Then
        MsgBox "No Existe el Registro", vbExclamation, xTitulo
        Exit Sub
    End If
    Agregando = True
    '--IDENTIFICACION DEL DERECHOHABIENTE
    txt(1).Text = NulosC(RstTmp("apepat"))
    txt(2).Text = NulosC(RstTmp("apemat"))
    txt(3).Text = NulosC(RstTmp("nombre"))
    '--DOCUMENTO DE IDENTIFACION
    If NulosN(RstTmp("idtipdoc")) <> 0 Then
        txt_cb(0).Text = NulosN(RstTmp("idtipdoc"))
        txt_cb_Validate 0, False
    End If
    txt(4).Text = NulosC(RstTmp("numdoc"))
    '--
    If IsDate(RstTmp("fchnac")) = True Then
        txtfecha(0).Valor = CDate(RstTmp("fchnac"))
    End If
    If NulosN(RstTmp("idsex")) <> 0 Then
        txt_cb(1).Text = NulosN(RstTmp("idsex"))
        txt_cb_Validate 1, False
    End If
    '--VINCULO FAMILIAR
    If NulosN(RstTmp("idvinfam")) <> 0 Then
        txt_cb(2).Text = NulosN(RstTmp("idvinfam"))
        txt_cb_Validate 2, False
        If NulosN(RstTmp("idtipdocpat")) <> 0 Then
            txt_cb(3).Text = NulosN(RstTmp("idtipdocpat"))
            txt_cb_Validate 3, False
            txt(5).Text = NulosC(RstTmp("numdocpat"))
        End If
    End If
    '--SITUACION DERECHOHABIENTE
    If NulosN(RstTmp("idsitderhab")) <> 0 Then
        txt_cb(4).Text = NulosN(RstTmp("idsitderhab"))
        txt_cb_Validate 4, False
        
        If NulosN(RstTmp("idsitderhab")) = 1 Then '--ACTIVO
            If IsDate(RstTmp("fchalt")) = True Then txtfecha(1).Valor = CDate(RstTmp("fchalt"))
        End If
        If NulosN(RstTmp("idsitderhab")) = 2 Then '--BAJA
            txt_cb(5).Text = NulosN(RstTmp("idtipbaj"))
            If IsDate(RstTmp("fchbaj")) = True Then txtfecha(2).Valor = CDate(RstTmp("fchbaj"))
        End If
    End If
    
    '--INCAPACIDAD
    If Trim(NulosC(RstTmp("numresinc"))) = "" Then
        opt_incapacidad(0).Value = True
    Else
        opt_incapacidad(1).Value = True
        txt(6).Text = NulosC(RstTmp("numresinc"))
    End If
    
    '--DOMICILIO DE DERECHOHABIENTE
    If NulosN(RstTmp("idinddom")) <> 0 Then
        txt_cb(6).Text = NulosN(RstTmp("idinddom"))
        txt_cb_Validate 6, False
    End If
    
    '--DE LA VIA
    If NulosN(RstTmp("idtipvia")) <> 0 Then
        txt_cb(7).Text = NulosN(RstTmp("idtipvia"))
        txt_cb_Validate 7, False
    End If
    txt(7).Text = NulosC(RstTmp("nomvia"))
    txt(8).Text = NulosC(RstTmp("numvia"))
    txt(9).Text = NulosC(RstTmp("intvia"))
    
    '--DE LA ZONA
    If NulosN(RstTmp("idtipzon")) <> 0 Then
        txt_cb(8).Text = NulosN(RstTmp("idtipzon"))
        txt_cb_Validate 8, False
    End If
    txt(10).Text = NulosC(RstTmp("nomzon"))
    txt(11).Text = NulosC(RstTmp("refdom"))
    
    '-DEL UBIGEO
    If NulosN(RstTmp("iddep")) <> 0 Then
        txt_cb(9).Text = NulosN(RstTmp("iddep"))
        txt_cb_Validate 9, False
        If NulosN(RstTmp("idpro")) <> 0 Then
            txt_cb(10).Text = NulosN(RstTmp("idpro"))
            txt_cb_Validate 10, False
            If NulosN(RstTmp("iddis")) <> 0 Then
                txt_cb(11).Text = NulosN(RstTmp("iddis"))
                txt_cb_Validate 11, False
            End If
        End If
    End If
    Set RstTmp = Nothing
    TabOne1.CurrTab = 0
    Agregando = False
    Exit Sub
error:
    Agregando = False
    Set RstTmp = Nothing
    cmd(0).Enabled = False
    SHOW_ERROR Me.Name, "pPonerDatos"
End Sub

Function Grabar() As Boolean

    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(Quehace = 1, "Grabar", "Modificar") + " al Derechohabiente ", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim xId As Integer

'    On Error GoTo LaCague

    xCon.BeginTrans

    If Quehace = 1 Then

        '*****************************************************
        Dim RstTmp As New ADODB.Recordset
        RST_Busq RstTmp, "SELECT corr From pla_derechohab Where idemp = " & mIdEmpleado & " ORDER BY corr DESC;", xCon
        If RstTmp.RecordCount <> 0 Then
            mCorrelativo = NulosN(RstTmp.Fields(0)) + 1
        Else
            mCorrelativo = 1
        End If
        Set RstTmp = Nothing
        '*****************************************************
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_derechohab", xCon

        RstCab.AddNew
    Else
        RST_Busq RstCab, "SELECT * FROM pla_derechohab WHERE idemp =  " & mIdEmpleado & " and corr = " & mCorrelativo & "", xCon
    End If
    
    RstCab("idemp") = mIdEmpleado
    RstCab("corr") = mCorrelativo
    '--IDENTIFICACION DEL DERECHOHABIENTE
    RstCab("apepat") = Trim(txt(1).Text)
    RstCab("apemat") = Trim(txt(2).Text)
    RstCab("nombre") = Trim(txt(3).Text)
    '--DOCUMENTO DE IDENTIFACION
    RstCab("idtipdoc") = NulosN(txt_cb(0).Text)
    RstCab("numdoc") = Trim(txt(4).Text)
    '--
    RstCab("fchnac") = CDate(txtfecha(0).Valor)
    RstCab("idsex") = NulosN(txt_cb(1).Text)
    '--VINCULO FAMILIAR
    RstCab("idvinfam") = NulosN(txt_cb(2).Text)
    RstCab("idtipdocpat") = NulosN(txt_cb(3).Text)
    RstCab("numdocpat") = Trim(txt(5).Text)
    '--SITUACION DERECHOHABIENTE
    RstCab("idsitderhab") = NulosN(txt_cb(4).Text)
    If NulosN(txt_cb(4).Text) = 1 Then '--ACTIVO
        RstCab("fchalt") = CDate(txtfecha(1).Valor)
    End If
    If NulosN(txt_cb(5).Text) = 2 Then '--BAJA
        RstCab("idtipbaj") = NulosN(txt_cb(5).Text)
        RstCab("fchbaj") = CDate(txtfecha(2).Valor)
    End If
    '--INCAPACIDAD
    RstCab("numresinc") = Trim(txt(6).Text)

    RstCab("idinddom") = NulosN(txt_cb(6).Text)
    
    '--DE LA VIA
    RstCab("idtipvia") = NulosN(txt_cb(7).Text)
    RstCab("nomvia") = Trim(txt(7).Text)
    RstCab("numvia") = Trim(txt(8).Text)
    RstCab("intvia") = Trim(txt(9).Text)
    '--DE LA ZONA
    RstCab("idtipzon") = NulosN(txt_cb(8).Text)
    RstCab("nomzon") = Trim(txt(10).Text)
    RstCab("refdom") = Trim(txt(11).Text)
    '-DEL UBIGEO
    RstCab("iddep") = NulosN(txt_cb(9).Text)
    RstCab("idpro") = NulosN(txt_cb(10).Text)
    RstCab("iddis") = NulosN(txt_cb(11).Text)
    '--
    RstCab.Update

    MsgBox "El Derechohabiente " + IIf(Quehace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo

    xCon.CommitTrans
    Set RstCab = Nothing
    Grabar = True
    Exit Function

LaCague:
    Set RstCab = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar al derechohabiente por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Private Function fValidarDatos() As Boolean
    
    Dim band As Integer
    band = Validar(txt)
    If band <> -1 Then
        MsgBox "Llene el Campo de " & lbl(band).Caption, vbInformation, xTitulo
        If band >= 1 And band <= 4 Then '--TAB 0
            TabOne1.CurrTab = 0
        ElseIf band >= 5 And band <= 6 Then '--TAB 1
            TabOne1.CurrTab = 1
        Else
            TabOne1.CurrTab = 2
        End If
        txt(band).SetFocus
        Exit Function
    End If
    
    band = Validar(txt_cb)
    If band <> -1 Then
        MsgBox "Llene el Campo de " & lbl_capt(band).Caption, vbInformation, xTitulo
        If band >= 0 And band <= 1 Then '--TAB 0
            TabOne1.CurrTab = 0
        ElseIf band >= 2 And band <= 6 Then '--TAB 1
            TabOne1.CurrTab = 1
        Else
            TabOne1.CurrTab = 2
        End If
       txt_cb(band).SetFocus
       Exit Function
    End If
    If (Len(Trim(txt(4).Text)) < 8 And NulosN(txt_cb(0).Text) = 1) Or (Len(Trim(txt(4).Text)) < 11 And NulosN(txt_cb(0).Text) = 5) Then
        MsgBox "Falta completar el Número de Documento", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        txt(4).SetFocus
        Exit Function
    End If

    If IsDate(txtfecha(0).Valor) = False Then
        MsgBox "Falta especificar la fecha de nacimiento", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        txtfecha(0).SetFocus
        Exit Function
    End If
    
    '--SITUACION DEL DERECHOHABIENTE
    If NulosN(txt_cb(4).Text) = 1 Then '--ALTA
        If IsDate(txtfecha(1).Valor) = False Then
            MsgBox "Falta especificar la fecha de alta", vbExclamation, xTitulo
            TabOne1.CurrTab = 1
            txtfecha(1).SetFocus
            Exit Function
        End If
    ElseIf NulosN(txt_cb(5).Text) = 1 Then  '--BAJA
        If IsDate(txtfecha(2).Valor) = False Then
            MsgBox "Falta especificar la fecha de baja", vbExclamation, xTitulo
            TabOne1.CurrTab = 1
            txtfecha(2).SetFocus
            Exit Function
        End If
    End If
    
    If opt_incapacidad(1).Value = True And Trim(txt(6).Text) = "" Then
        TabOne1.CurrTab = 1
        MsgBox "Falta especificar la Resolución Directoral por Incapacidad", vbExclamation, xTitulo
        txt(6).SetFocus
        Exit Function
    End If

    If NulosN(txt_cb(2).Text) = 1 Then '--ES HIJO
        If fObtenerEdad(txtfecha(0).Valor) >= 18 And opt_incapacidad(0).Value = True Then
            TabOne1.CurrTab = 1
            MsgBox "La persona que intenta registrar tiene " + CStr(fObtenerEdad(txtfecha(0).Valor)) + " Años" + vbCr + _
            "Es necesario ingresar el Número de la Resolución Directoral de incapacidad", vbExclamation, xTitulo
            opt_incapacidad(1).Value = True
            txt(6).SetFocus
            Exit Function
        End If
    End If
    
    
    fValidarDatos = True
    
End Function
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case Index
        Case 4 '--
            Select Case NulosN(txt_cb(0).Text)
                Case 1, 5 '--DNI,RUC
                    If validar_numero(KeyAscii) = False Then KeyAscii = 0
                Case Else
                    
            End Select

    End Select
End Sub

Private Function fObtenerEdad(fNacimiento As Date) As Integer
    '--ESTA FUNCION CALCULARA LA EDAD ACTUAL EN FUNCION DE LA FECHA DE NACIMIENTO, Y LA FECHA ACTUAL
    '--NAC: 24/03/1981
    '--ACT: 04/05/2008  EDAD: 27
    '--ACT: 20/02/2008  EDAD: 26
    Dim mAnnoFin As Integer
    Dim mAnnoInicio As Integer
    Dim fActual As Date
    fActual = Date
    mAnnoInicio = Year(fNacimiento)
    mAnnoFin = Year(fActual)
    
    If Month(fNacimiento) > Month(fActual) Then
        mAnnoFin = mAnnoFin - 1
    ElseIf Month(fNacimiento) = Month(fActual) Then
        If Day(fNacimiento) > Day(fActual) Then
            mAnnoFin = mAnnoFin - 1
        End If
    End If
    
    fObtenerEdad = mAnnoFin - mAnnoInicio
    
End Function
