VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmCatTrabajador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Nómina del Personal - Trabajador"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Index           =   0
      Left            =   915
      MaxLength       =   40
      TabIndex        =   40
      Text            =   "txt(0)"
      Top             =   5190
      Width           =   1140
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   345
      Index           =   13
      Left            =   90
      TabIndex        =   38
      Top             =   75
      Width           =   5355
      Begin VB.Line lin 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   3
         X1              =   -15
         X2              =   6395
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   -30
         X2              =   6380
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   5340
         X2              =   5340
         Y1              =   15
         Y2              =   395
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Categoría: Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   90
         TabIndex        =   39
         Top             =   45
         Width           =   2325
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   405
      Index           =   12
      Left            =   135
      TabIndex        =   36
      Top             =   4725
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
         TabIndex        =   37
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
      Left            =   7605
      TabIndex        =   35
      Top             =   4710
      Width           =   1755
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Grabar"
      Height          =   420
      Index           =   0
      Left            =   5625
      TabIndex        =   34
      Top             =   4710
      Width           =   1755
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   4155
      Left            =   90
      TabIndex        =   25
      Top             =   465
      Width           =   9300
      _cx             =   16404
      _cy             =   7329
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
      Caption         =   "    Datos Principales    |  Datos Complementarios"
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
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Height          =   3735
         Index           =   2
         Left            =   10245
         TabIndex        =   28
         Top             =   45
         Width           =   9210
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   3735
         Index           =   1
         Left            =   9945
         TabIndex        =   27
         Top             =   45
         Width           =   9210
         Begin VB.Frame fra 
            Caption         =   "[ Periodicidad del Ingreso ]"
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
            Height          =   645
            Index           =   17
            Left            =   0
            TabIndex        =   96
            Top             =   2955
            Width           =   4965
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   11
               Left            =   540
               Picture         =   "FrmCatTrabajador.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   97
               ToolTipText     =   "Seleccione el Periodo de Ingreso"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   11
               Left            =   135
               MaxLength       =   20
               TabIndex        =   20
               Tag             =   "null"
               Text            =   "txt_cb(11)"
               Top             =   285
               Width           =   645
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
               Left            =   3795
               TabIndex        =   99
               Top             =   300
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   " Periodicidad del Ingreso"
               Height          =   195
               Index           =   11
               Left            =   3030
               TabIndex        =   98
               Top             =   75
               Visible         =   0   'False
               Width           =   1740
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
               Left            =   795
               TabIndex        =   100
               Top             =   285
               Width           =   4065
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Situalión Especial ]"
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
            Height          =   645
            Index           =   21
            Left            =   4965
            TabIndex        =   108
            Top             =   1605
            Width           =   4245
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   10
               Left            =   540
               Picture         =   "FrmCatTrabajador.frx":0132
               Style           =   1  'Graphical
               TabIndex        =   109
               ToolTipText     =   "Seleccione la Situación Especial"
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   10
               Left            =   135
               MaxLength       =   20
               TabIndex        =   17
               Tag             =   "null"
               Text            =   "txt_cb(10)"
               Top             =   240
               Width           =   645
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
               Left            =   2655
               TabIndex        =   111
               Top             =   240
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Situalión Especial"
               Height          =   195
               Index           =   10
               Left            =   2655
               TabIndex        =   110
               Top             =   75
               Visible         =   0   'False
               Width           =   1245
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
               Left            =   795
               TabIndex        =   112
               Top             =   240
               Width           =   3350
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Tipo de Pago ]"
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
            Height          =   645
            Index           =   20
            Left            =   4965
            TabIndex        =   103
            Top             =   2955
            Width           =   4245
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   12
               Left            =   540
               Picture         =   "FrmCatTrabajador.frx":0264
               Style           =   1  'Graphical
               TabIndex        =   104
               ToolTipText     =   "Seleccione el Tipo de Pago"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   12
               Left            =   135
               MaxLength       =   20
               TabIndex        =   21
               Tag             =   "null"
               Text            =   "txt_cb(12)"
               Top             =   285
               Width           =   645
            End
            Begin VB.Label lbl_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cod(12)"
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
               Index           =   12
               Left            =   2655
               TabIndex        =   107
               Top             =   285
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   " Tipo de Pago"
               Height          =   195
               Index           =   12
               Left            =   2655
               TabIndex        =   106
               Top             =   90
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(12)"
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
               Index           =   12
               Left            =   795
               TabIndex        =   105
               Top             =   285
               Width           =   3350
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ ¿Sidicalizado? ]"
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
            Index           =   19
            Left            =   5820
            TabIndex        =   102
            Top             =   2310
            Width           =   2535
            Begin VB.OptionButton opt_sindato 
               Caption         =   "Si"
               Height          =   330
               Index           =   1
               Left            =   1575
               TabIndex        =   24
               Top             =   225
               Width           =   585
            End
            Begin VB.OptionButton opt_sindato 
               Caption         =   "No"
               Height          =   330
               Index           =   0
               Left            =   480
               TabIndex        =   19
               Top             =   225
               Value           =   -1  'True
               Width           =   585
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ ¿Tiene rentas de quinta categoría exoneradas o inafectas ? ]"
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
            Index           =   18
            Left            =   30
            TabIndex        =   101
            Top             =   2310
            Width           =   5670
            Begin VB.OptionButton opt_Renta5ta 
               Caption         =   "No"
               Height          =   330
               Index           =   0
               Left            =   1290
               TabIndex        =   18
               Top             =   240
               Value           =   -1  'True
               Width           =   585
            End
            Begin VB.OptionButton opt_Renta5ta 
               Caption         =   "Si"
               Height          =   330
               Index           =   1
               Left            =   2955
               TabIndex        =   23
               Top             =   240
               Width           =   585
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Situación del Trabajador ]"
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
            Index           =   16
            Left            =   4965
            TabIndex        =   91
            Top             =   982
            Width           =   4245
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   9
               Left            =   540
               Picture         =   "FrmCatTrabajador.frx":0396
               Style           =   1  'Graphical
               TabIndex        =   92
               ToolTipText     =   "Seleccione la Situación del Trabajador"
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   9
               Left            =   135
               MaxLength       =   20
               TabIndex        =   16
               Tag             =   "null"
               Text            =   "txt_cb(9)"
               Top             =   240
               Width           =   645
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Situación del Trabajador"
               Height          =   195
               Index           =   9
               Left            =   2430
               TabIndex        =   94
               Top             =   75
               Visible         =   0   'False
               Width           =   1725
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
               Left            =   2655
               TabIndex        =   93
               Top             =   240
               Visible         =   0   'False
               Width           =   975
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
               Left            =   795
               TabIndex        =   95
               Top             =   240
               Width           =   3350
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Prestaciones de Salud ]"
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
            Index           =   15
            Left            =   4965
            TabIndex        =   85
            Top             =   0
            Width           =   4245
            Begin VB.OptionButton opt_eps 
               Caption         =   "No"
               Height          =   330
               Index           =   0
               Left            =   2910
               TabIndex        =   14
               Top             =   195
               Value           =   -1  'True
               Width           =   540
            End
            Begin VB.OptionButton opt_eps 
               Caption         =   "Si"
               Height          =   330
               Index           =   1
               Left            =   3570
               TabIndex        =   22
               Top             =   195
               Width           =   585
            End
            Begin VB.CommandButton cb 
               Enabled         =   0   'False
               Height          =   225
               Index           =   8
               Left            =   540
               Picture         =   "FrmCatTrabajador.frx":04C8
               Style           =   1  'Graphical
               TabIndex        =   86
               Top             =   585
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Enabled         =   0   'False
               Height          =   300
               Index           =   8
               Left            =   135
               MaxLength       =   20
               TabIndex        =   15
               Tag             =   "null"
               Text            =   "txt_cb(8)"
               Top             =   555
               Width           =   645
            End
            Begin VB.Label Label3 
               Caption         =   "¿Afiliado a EPS/Servicios Propios?"
               Height          =   225
               Left            =   105
               TabIndex        =   90
               Top             =   300
               Width           =   2640
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
               Left            =   2655
               TabIndex        =   89
               Top             =   555
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre EPS / Serv. Propios."
               Height          =   195
               Index           =   8
               Left            =   1860
               TabIndex        =   88
               Top             =   825
               Visible         =   0   'False
               Width           =   2070
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
               Left            =   795
               TabIndex        =   87
               Top             =   555
               Width           =   3350
            End
         End
         Begin VB.Frame fra 
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
            Index           =   14
            Left            =   3345
            TabIndex        =   83
            Top             =   765
            Width           =   1605
            Begin VB.OptionButton opt_Ingreso5ta 
               Caption         =   "No"
               Height          =   330
               Index           =   0
               Left            =   135
               TabIndex        =   12
               Top             =   1020
               Value           =   -1  'True
               Width           =   585
            End
            Begin VB.OptionButton opt_Ingreso5ta 
               Caption         =   "Si"
               Height          =   330
               Index           =   1
               Left            =   885
               TabIndex        =   13
               Top             =   1020
               Width           =   585
            End
            Begin VB.Label Label1 
               Caption         =   "[ El trabajador informó otros ingresos de 5ta Categoría]"
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
               Left            =   120
               TabIndex        =   84
               Top             =   -15
               Width           =   1365
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ El Trabajador está sujeto a ] "
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
            Index           =   11
            Left            =   30
            TabIndex        =   80
            Top             =   765
            Width           =   3285
            Begin VB.CheckBox chk 
               Caption         =   "¿Trabajo en horario nocturno?"
               Height          =   225
               Index           =   2
               Left            =   90
               TabIndex        =   82
               Top             =   1170
               Width           =   2760
            End
            Begin VB.CheckBox chk 
               Caption         =   "¿Jornada de trabajo máxima?"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   81
               Top             =   870
               Width           =   2760
            End
            Begin VB.CheckBox chk 
               Caption         =   "Trabajador sujeto a régimen alternativo, acumulativo o atípico de jornada de trabajo y descanso"
               Height          =   540
               Index           =   0
               Left            =   90
               TabIndex        =   11
               Top             =   270
               Width           =   3105
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Tipo de Contrato de Trabajo ]"
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
            Height          =   690
            Index           =   10
            Left            =   0
            TabIndex        =   75
            Top             =   0
            Width           =   4965
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   7
               Left            =   540
               Picture         =   "FrmCatTrabajador.frx":05FA
               Style           =   1  'Graphical
               TabIndex        =   76
               ToolTipText     =   "Seleccione el Tipo de Contrato de Trabajo"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   7
               Left            =   135
               MaxLength       =   20
               TabIndex        =   10
               Tag             =   "null"
               Text            =   "txt_cb(7)"
               Top             =   285
               Width           =   645
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Contrato de Trabajo"
               Height          =   195
               Index           =   7
               Left            =   2835
               TabIndex        =   78
               Top             =   90
               Visible         =   0   'False
               Width           =   1995
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
               Left            =   3675
               TabIndex        =   77
               Top             =   270
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
               Left            =   795
               TabIndex        =   79
               Top             =   285
               Width           =   4065
            End
         End
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3735
         Index           =   0
         Left            =   45
         TabIndex        =   26
         Top             =   45
         Width           =   9210
         Begin VB.Frame fra 
            Caption         =   "[ Seguro Complementario de Trabajo de Riesgo (SCTR) ]"
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
            Height          =   690
            Index           =   9
            Left            =   45
            TabIndex        =   66
            Top             =   2985
            Width           =   9075
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   6
               Left            =   5640
               Picture         =   "FrmCatTrabajador.frx":072C
               Style           =   1  'Graphical
               TabIndex        =   71
               ToolTipText     =   "Seleccione el SCTR Pensión"
               Top             =   270
               Width           =   210
            End
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   5
               Left            =   990
               Picture         =   "FrmCatTrabajador.frx":085E
               Style           =   1  'Graphical
               TabIndex        =   67
               ToolTipText     =   "Seleccione el SCTR Salud"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   5
               Left            =   600
               MaxLength       =   20
               TabIndex        =   8
               Tag             =   "null"
               Text            =   "txt_cb(5)"
               Top             =   285
               Width           =   645
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   6
               Left            =   5235
               MaxLength       =   20
               TabIndex        =   9
               Tag             =   "null"
               Text            =   "txt_cb(6)"
               Top             =   240
               Width           =   645
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
               Left            =   7305
               TabIndex        =   74
               Top             =   240
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pensión"
               Height          =   195
               Index           =   6
               Left            =   4455
               TabIndex        =   73
               Top             =   330
               Width           =   570
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
               Left            =   5880
               TabIndex        =   72
               Top             =   240
               Width           =   3045
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
               Left            =   2655
               TabIndex        =   70
               Top             =   285
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Salud"
               Height          =   195
               Index           =   5
               Left            =   105
               TabIndex        =   69
               Top             =   375
               Width           =   405
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
               Left            =   1245
               TabIndex        =   68
               Top             =   285
               Width           =   2940
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Régimen Pensionario ]"
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
            Height          =   690
            Index           =   8
            Left            =   45
            TabIndex        =   59
            Top             =   2244
            Width           =   9075
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   1
               Left            =   7440
               MaxLength       =   12
               TabIndex        =   7
               Tag             =   "null"
               Text            =   "txt(1)"
               ToolTipText     =   "Código Unico de Identificación del Sistema Privado de Pensiones"
               Top             =   270
               Width           =   1455
            End
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   4
               Left            =   540
               Picture         =   "FrmCatTrabajador.frx":0990
               Style           =   1  'Graphical
               TabIndex        =   60
               ToolTipText     =   "Seleccione el Régimen Pensionario"
               Top             =   315
               Width           =   210
            End
            Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
               Height          =   300
               Index           =   0
               Left            =   5445
               TabIndex        =   6
               ToolTipText     =   "Ingrese la Fecha de Inscripción al Régimen Pensionario"
               Top             =   270
               Width           =   1350
               _ExtentX        =   2381
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
               Index           =   4
               Left            =   135
               MaxLength       =   20
               TabIndex        =   5
               Tag             =   "null"
               Text            =   "txt_cb(4)"
               Top             =   285
               Width           =   645
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CUSPP"
               Height          =   195
               Index           =   1
               Left            =   6870
               TabIndex        =   65
               Top             =   360
               Width           =   540
            End
            Begin VB.Label lbl_fecha 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "F. Inscripción"
               Height          =   195
               Index           =   0
               Left            =   4455
               TabIndex        =   64
               Top             =   360
               Width           =   945
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Régimen Pensionario"
               Height          =   195
               Index           =   4
               Left            =   2490
               TabIndex        =   62
               Top             =   90
               Visible         =   0   'False
               Width           =   1500
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
               Left            =   2655
               TabIndex        =   61
               Top             =   285
               Visible         =   0   'False
               Width           =   975
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
               Left            =   780
               TabIndex        =   63
               Top             =   285
               Width           =   3390
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ ¿Es Discapacitado? ]"
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
            Height          =   690
            Index           =   7
            Left            =   6105
            TabIndex        =   57
            Top             =   768
            Width           =   2985
            Begin VB.OptionButton opt_discapacidad 
               Caption         =   "No"
               Height          =   210
               Index           =   0
               Left            =   630
               TabIndex        =   3
               Top             =   315
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.OptionButton opt_discapacidad 
               Caption         =   "Si"
               Height          =   210
               Index           =   1
               Left            =   1680
               TabIndex        =   58
               Top             =   315
               Width           =   675
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Ocupación ]"
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
            Height          =   690
            Index           =   6
            Left            =   45
            TabIndex        =   51
            Top             =   1506
            Width           =   9075
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   3
               Left            =   825
               Picture         =   "FrmCatTrabajador.frx":0AC2
               Style           =   1  'Graphical
               TabIndex        =   52
               ToolTipText     =   "Seleccione la Ocupación"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   3
               Left            =   135
               MaxLength       =   20
               TabIndex        =   4
               Tag             =   "null"
               Text            =   "txt_cb(3)"
               Top             =   285
               Width           =   930
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
               Left            =   2655
               TabIndex        =   54
               Top             =   285
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ocupación "
               Height          =   195
               Index           =   3
               Left            =   3675
               TabIndex        =   53
               Top             =   105
               Visible         =   0   'False
               Width           =   825
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
               Left            =   1065
               TabIndex        =   55
               Top             =   285
               Width           =   7710
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Nivel Educativo]"
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
            Height          =   690
            Index           =   5
            Left            =   60
            TabIndex        =   46
            Top             =   768
            Width           =   5925
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   2
               Left            =   540
               Picture         =   "FrmCatTrabajador.frx":0BF4
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Seleccione el Nivel Educativo"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   2
               Left            =   135
               MaxLength       =   20
               TabIndex        =   2
               Tag             =   "null"
               Text            =   "txt_cb(2)"
               Top             =   285
               Width           =   645
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
               Left            =   2655
               TabIndex        =   49
               Top             =   285
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nivel Educativo"
               Height          =   195
               Index           =   2
               Left            =   3675
               TabIndex        =   48
               Top             =   90
               Visible         =   0   'False
               Width           =   1125
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
               Left            =   780
               TabIndex        =   50
               Top             =   285
               Width           =   4980
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Régimen Laboral ]"
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
            Height          =   690
            Index           =   3
            Left            =   6105
            TabIndex        =   42
            Top             =   30
            Width           =   2985
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   1
               Left            =   540
               Picture         =   "FrmCatTrabajador.frx":0D26
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Seleccione el Régimen Laboral"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   1
               Left            =   135
               MaxLength       =   20
               TabIndex        =   1
               Tag             =   "null"
               Text            =   "txt_cb(1)"
               Top             =   285
               Width           =   645
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Régimen Laboral"
               Height          =   195
               Index           =   1
               Left            =   1650
               TabIndex        =   56
               Top             =   90
               Visible         =   0   'False
               Width           =   1200
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
               Left            =   1785
               TabIndex        =   44
               Top             =   285
               Visible         =   0   'False
               Width           =   975
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
               Left            =   780
               TabIndex        =   45
               Top             =   285
               Width           =   1845
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Tipo de Trabajador]"
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
            Height          =   690
            Index           =   4
            Left            =   60
            TabIndex        =   29
            Top             =   30
            Width           =   5925
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   0
               Left            =   540
               Picture         =   "FrmCatTrabajador.frx":0E58
               Style           =   1  'Graphical
               TabIndex        =   30
               ToolTipText     =   "Seleccione el Tipo de Trabajador"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   0
               Left            =   135
               MaxLength       =   20
               TabIndex        =   0
               Text            =   "txt_cb(0)"
               Top             =   285
               Width           =   645
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
               Left            =   2655
               TabIndex        =   33
               Top             =   285
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Trabajador"
               Height          =   195
               Index           =   0
               Left            =   3675
               TabIndex        =   32
               Top             =   120
               Visible         =   0   'False
               Width           =   1350
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
               Left            =   780
               TabIndex        =   31
               Top             =   285
               Width           =   5010
            End
         End
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   41
      Top             =   5280
      Width           =   495
   End
End
Attribute VB_Name = "FrmCatTrabajador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Quehace As Integer
Dim mCorrelativo As Long
Dim mIdEmpleado As Long
Dim Agregando As Boolean

Public Sub pRecibeLink(QueHace1 As Integer)
    Quehace = QueHace1
    mCorrelativo = mCorr
    With FrmNomina2
        mIdEmpleado = NulosN(.txt(0).Text)
        lbl_persona.Caption = .lbl_persona(0).Caption
    End With
    '------
    Agregando = True
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cb
    LimpiaText lbl_cod
    LimpiaText txtfecha
    TabOne1.CurrTab = 0
    pPonerDatos
    Agregando = False
End Sub


'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    If Quehace = 3 Then Exit Sub
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error
    Select Case Index
        Case 0 '--TIPO DE TRABAJADOR
            nTitulo = "Buscando Tipo de Trabajador"
            nSQL = "SELECT mae_tipotrabajador.id, mae_tipotrabajador.descripcion AS nombre, mae_tipotrabajador.id AS cod, mae_tipotrabajador.codsun " _
                + vbCr + " FROM mae_tipotrabajador INNER JOIN mae_tipotrabajadorcat ON mae_tipotrabajador.id = mae_tipotrabajadorcat.id " _
                + vbCr + " WHERE mae_tipotrabajadorcat.idcat = 1 " _
                + vbCr + " ORDER BY mae_tipotrabajador.codsun;"
        
        Case 1 '--REGIMEN LABORAL
            nTitulo = "Buscando Régimen Laboral"
            nSQL = "SELECT mae_regimenlab.id, mae_regimenlab.descripcion AS nombre, mae_regimenlab.id AS cod " _
                + vbCr + " From mae_regimenlab " _
                + vbCr + " ORDER BY mae_regimenlab.codsun;"
        
        Case 2 '--NIVEL EDUCATIVO
            nTitulo = "Buscando Nivel Educativo"
            nSQL = "SELECT mae_niveleducativo.id, mae_niveleducativo.descripcion AS nombre, mae_niveleducativo.id AS cod " _
                + vbCr + " From mae_niveleducativo " _
                + vbCr + " ORDER BY mae_niveleducativo.codsun;"
                 
        Case 3 '--OCUPACION
            nTitulo = "Buscando Ocupación"
            nSQL = "SELECT mae_ocupacion.id, mae_ocupacion.descripcion AS nombre, mae_ocupacion.id AS cod " _
                + vbCr + " From mae_ocupacion " _
                + vbCr + " ORDER BY mae_ocupacion.codsun;"

        Case 4 '--REGIMEN PENSIONARIO
            nTitulo = "Buscando Situación de Derechohabiente"
            nSQL = "SELECT mae_regimenpen.id, mae_regimenpen.descripcion AS nombre, mae_regimenpen.id AS cod, mae_regimenpen.cuspp " _
                + vbCr + " From mae_regimenpen " _
                + vbCr + " ORDER BY mae_regimenpen.codsun;"
        
        Case 5 '--SCTR - SALUD
            nTitulo = "Buscando SCTR - Salud"
            nSQL = "SELECT mae_sctrsalud.id, mae_sctrsalud.descripcion AS nombre, mae_sctrsalud.id AS cod " _
                + vbCr + " From mae_sctrsalud " _
                + vbCr + " ORDER BY mae_sctrsalud.codsun;"

        Case 6 '--SCTR - PENSION
            nTitulo = "Buscando SCTR - Pensión"
            nSQL = "SELECT mae_sctrpension.id, mae_sctrpension.descripcion AS nombre, mae_sctrpension.id AS cod " _
                + vbCr + " From mae_sctrpension " _
                + vbCr + " ORDER BY mae_sctrpension.codsun;"
                
        Case 7 '--CONTRATO DE TRABAJO
            nTitulo = "Buscando Tipo de Contrato"
            nSQL = "SELECT mae_tipocontrato.id, mae_tipocontrato.descripcion AS nombre, mae_tipocontrato.id AS cod " _
                + vbCr + " From mae_tipocontrato " _
                + vbCr + " ORDER BY mae_tipocontrato.codsun;"
        
        Case 8 '--PRESTACION DE SALUD
            nTitulo = "Buscando Entidad Prestadora de Salud"
            nSQL = "SELECT mae_eps.id, mae_eps.descripcion AS nombre, mae_eps.id AS cod, mae_eps.numruc " _
                + vbCr + " From mae_eps " _
                + vbCr + " ORDER BY mae_eps.codsun;"

        Case 9 '--SITUACION DE TRABAJO
            If opt_eps(0).Value = False And opt_eps(1).Value = False Then
                MsgBox "Especifique si el Trabajador está afiliado a una EPS/Servicios Propios", vbExclamation, xTitulo
                opt_eps(0).SetFocus
                Exit Sub
            End If
            nTitulo = "Buscando Situación del Trajabador"
            nSQL = "SELECT mae_situacion.id, mae_situacion.descripcion AS nombre, mae_situacion.id AS cod " _
                + vbCr + " From mae_situacion " _
                + vbCr + " Where (((mae_situacion.afiliado) = " & IIf(opt_eps(0).Value = True, "0", "-1") & " )) " _
                + vbCr + " ORDER BY mae_situacion.codsun;"
        
        Case 10 '--SITUACION ESPECIAL
            nTitulo = "Buscando Situación Especial"
            nSQL = "SELECT mae_situatraba.id, mae_situatraba.descripcion AS nombre, mae_situatraba.id AS cod " _
                + vbCr + " From mae_situatraba " _
                + vbCr + " ORDER BY mae_situatraba.codsun;"
    
        Case 11 '--PERIODICIDAD DE INGRESO
            nTitulo = "Buscando Periodicidad del Ingreso"
            nSQL = "SELECT mae_periocidad.id, mae_periocidad.descripcion AS nombre, mae_periocidad.id AS cod " _
                + vbCr + " From mae_periocidad " _
                + vbCr + " ORDER BY mae_periocidad.codsun;"

        Case 12 '--TIPO DE PAGO
            nTitulo = "Buscando Tipo de Pago"
            nSQL = "SELECT mae_tipopago.id, mae_tipopago.descripcion AS nombre, mae_tipopago.id AS cod " _
                + vbCr + " From mae_tipopago " _
                + vbCr + " ORDER BY mae_tipopago.codsun;"
    
    End Select
    
    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
            
    Dim xRs As New ADODB.Recordset
    If Index = 3 Or Index = 2 Then
    '--SOLO OCUPACION
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", CualquierParte
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
    End If

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    lbl_cb(Index).ToolTipText = xRs.Fields(1) & "" '--NOMBRE
   
   
    Select Case Index
        Case 0 '--TIPO DE TRABAJADOR
            txt_cb(1).SetFocus
        Case 1 '--NIVEL EDUCATIVO
            txt_cb(2).SetFocus
        Case 2 '--REGIMEN LABORAL
            opt_discapacidad(0).SetFocus
        Case 3 '--OCUPACION
            txt_cb(4).SetFocus
        Case 4 '--REGIMEN PENSIONARIO
'            If NulosN(xRs.Fields("cuspp")) = -1 Then
'                txt(1).Visible = True
'                lbl(1).Visible = True
'            Else
'                txt(1).Text = ""
'                txt(1).Visible = False
'                lbl(1).Visible = False
'            End If
            txtfecha(0).SetFocus
        Case 5 '--SCTR - SALUD
            txt_cb(6).SetFocus
        Case 6 '--SCTR - PENSION
            TabOne1.CurrTab = 1
            txt_cb(7).SetFocus
        Case 7 '--CONTRATO DE TRABAJO
            chk(0).SetFocus '--TRABAJO SUJETO A REGIMEN ALTERNATIVO,ACUMULAT....
        Case 8 '--PRESTACION DE SALUD
            txt_cb(9).SetFocus
        Case 9 '--SITUACION DE TRABAJO
            txt_cb(10).SetFocus
        Case 10 '--PERIODICIDAD DE INGRESO
            opt_Renta5ta(0).SetFocus '--TIENE RENTAS DE QUINTA CATEG...
        Case 11 '--TIPO DE PAGO
            txt_cb(12).SetFocus
        Case 12 '--SITUACION ESPECIAL
            cmd(0).SetFocus
    End Select
Salir:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        opt_Ingreso5ta(0).SetFocus
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--GRABAR
            If Grabar() = False Then Exit Sub
            FrmNomina2.pCargarDatosPeriodoLaboral
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

Private Sub opt_discapacidad_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_cb(3).SetFocus
End Sub

Private Sub opt_eps_Click(Index As Integer)
    If Index = 0 Then
        txt_cb(8).Text = ""
        txt_cb(8).Enabled = False
        cb(8).Enabled = False
        
        txt_cb(9).Text = "" '--SITUACION DE TRABAJADOR
    Else
        txt_cb(8).Enabled = True
        cb(8).Enabled = True
    End If
    If Agregando = False Then opt_eps(Index).SetFocus
End Sub

Private Sub opt_eps_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        txt_cb(9).SetFocus
    Else
        txt_cb(8).SetFocus
    End If
End Sub

Private Sub opt_Ingreso5ta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then opt_eps(0).SetFocus
End Sub

Private Sub opt_Renta5ta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then opt_sindato(0).SetFocus
End Sub

Private Sub opt_sindato_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_cb(11).SetFocus
End Sub

Private Sub TabOne1_Click()
    If TabOne1.CurrTab = 0 Then
        If Agregando = False Then txt_cb(0).SetFocus
    Else
        If Agregando = False Then txt_cb(7).SetFocus
    End If
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If Quehace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
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
        Case 0 '--TIPO DE TRABAJADOR
            nSQL = "SELECT mae_tipotrabajador.id, mae_tipotrabajador.descripcion AS nombre, mae_tipotrabajador.id AS cod, mae_tipotrabajador.codsun " _
                + vbCr + " FROM mae_tipotrabajador INNER JOIN mae_tipotrabajadorcat ON mae_tipotrabajador.id = mae_tipotrabajadorcat.id  " _
                + vbCr + " WHERE mae_tipotrabajador.id = " & NulosN(txt_cb(Index).Text) & " AND mae_tipotrabajadorcat.idcat = 1 ;"
               
        Case 1 '--REGIMEN LABORAL
            nSQL = "SELECT mae_regimenlab.id, mae_regimenlab.descripcion AS nombre, mae_regimenlab.id AS cod " _
                + vbCr + " From mae_regimenlab " _
                + vbCr + " WHERE mae_regimenlab.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 2 '--NIVEL EDUCATIVO
            nSQL = "SELECT mae_niveleducativo.id, mae_niveleducativo.descripcion AS nombre, mae_niveleducativo.id AS cod " _
                + vbCr + " From mae_niveleducativo " _
                + vbCr + " WHERE mae_niveleducativo.id = " & NulosN(txt_cb(Index).Text) & ";"
                 
        Case 3 '--OCUPACION
            nSQL = "SELECT mae_ocupacion.id, mae_ocupacion.descripcion AS nombre, mae_ocupacion.id AS cod " _
                + vbCr + " From mae_ocupacion " _
                + vbCr + " WHERE mae_ocupacion.id = " & NulosN(txt_cb(Index).Text) & ";"

        Case 4 '--REGIMEN PENSIONARIO
            nSQL = "SELECT mae_regimenpen.id, mae_regimenpen.descripcion AS nombre, mae_regimenpen.id AS cod, mae_regimenpen.cuspp " _
                + vbCr + " From mae_regimenpen " _
                + vbCr + " WHERE mae_regimenpen.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 5 '--SCTR - SALUD
            nSQL = "SELECT mae_sctrsalud.id, mae_sctrsalud.descripcion AS nombre, mae_sctrsalud.id AS cod " _
                + vbCr + " From mae_sctrsalud " _
                + vbCr + " WHERE mae_sctrsalud.id = " & NulosN(txt_cb(Index).Text) & ";"

        Case 6 '--SCTR - PENSION
            nSQL = "SELECT mae_sctrpension.id, mae_sctrpension.descripcion AS nombre, mae_sctrpension.id AS cod " _
                + vbCr + " From mae_sctrpension " _
                + vbCr + " WHERE mae_sctrpension.id = " & NulosN(txt_cb(Index).Text) & ";"
                
        Case 7 '--CONTRATO DE TRABAJO
            nSQL = "SELECT mae_tipocontrato.id, mae_tipocontrato.descripcion AS nombre, mae_tipocontrato.id AS cod " _
                + vbCr + " From mae_tipocontrato " _
                + vbCr + " WHERE mae_tipocontrato.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 8 '--PRESTACION DE SALUD
            nSQL = "SELECT mae_eps.id, mae_eps.descripcion AS nombre, mae_eps.id AS cod, mae_eps.numruc " _
                + vbCr + " From mae_eps " _
                + vbCr + " WHERE mae_eps.id = " & NulosN(txt_cb(Index).Text) & ";"

        Case 9 '--SITUACION DE TRABAJO
            If opt_eps(0).Value = False And opt_eps(1).Value = False And Agregando = False Then
                
                MsgBox "Especifique si el Trabajador está afiliado a una EPS/Servicios Propios", vbExclamation, xTitulo
                opt_eps(0).SetFocus
                GoTo Salir
            End If
            nSQL = "SELECT mae_situacion.id, mae_situacion.descripcion AS nombre, mae_situacion.id AS cod " _
                + vbCr + " From mae_situacion " _
                + vbCr + " Where (((mae_situacion.afiliado) = " & IIf(opt_eps(0).Value = True, "0", "-1") & " )) " _
                + vbCr + " AND mae_situacion.id = " & NulosN(txt_cb(Index).Text) & ";"
    
        Case 10 '--SITUACION ESPECIAL
            nSQL = "SELECT mae_situatraba.id, mae_situatraba.descripcion AS nombre, mae_situatraba.id AS cod " _
                + vbCr + " From mae_situatraba " _
                + vbCr + " WHERE mae_situatraba.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 11 '--PERIODICIDAD DE INGRESO
            nSQL = "SELECT mae_periocidad.id, mae_periocidad.descripcion AS nombre, mae_periocidad.id AS cod " _
                + vbCr + " From mae_periocidad " _
                + vbCr + " WHERE mae_periocidad.id = " & NulosN(txt_cb(Index).Text) & ";"

        Case 12 '--TIPO DE PAGO
            nSQL = "SELECT mae_tipopago.id, mae_tipopago.descripcion AS nombre, mae_tipopago.id AS cod " _
                + vbCr + " From mae_tipopago " _
                + vbCr + " WHERE mae_tipopago.id = " & NulosN(txt_cb(Index).Text) & ";"
    
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
    '--------------If Agregando = False Then
    Select Case Index
        Case 2  '--NIVEL EDUCATIVO
            If Agregando = False Then opt_discapacidad(0).SetFocus
        Case 4 '--REGIMEN PENSIONARIO
'            If NulosN(RstTmp.Fields("cuspp")) = -1 Then
'                txt(1).Visible = True
'                lbl(1).Visible = True
'            Else
'                txt(1).Text = ""
'                txt(1).Visible = False
'                lbl(1).Visible = False
'            End If
            If Agregando = False Then txtfecha(0).SetFocus
        Case 6 '--SCTR - PENSION
            TabOne1.CurrTab = 1
            If Agregando = False Then txt_cb(7).SetFocus
        Case 7 '--CONTRATO DE TRABAJO
            If Agregando = False Then chk(0).SetFocus '--TRABAJO SUJETO A REGIMEN ALTERNATIVO,ACUMULAT....
        Case 10 '--PERIODICIDAD DE INGRESO
            If Agregando = False Then opt_Renta5ta(0).SetFocus '--TIENE RENTAS DE QUINTA CATEG...
        Case 12 '--SITUACION ESPECIAL
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
    nSQL = "SELECT pla_categoria1.* From pla_categoria1 WHERE (((pla_categoria1.idemp)=" & mIdEmpleado & "));"

    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.RecordCount = 0 Then
        Quehace = 1
        txt_cb(1).Text = "1"
        txt_cb_Validate 1, False
        Exit Sub
    End If
    Quehace = 2
    Agregando = True

    '************************************************************TAB 0
    '--TIPO TRABAJADOR
    If NulosN(RstTmp("idtiptra")) <> 0 Then
        txt_cb(0).Text = NulosN(RstTmp("idtiptra"))
        txt_cb_Validate 0, False
    End If
    '--REGIMEN LABORAL
    If NulosN(RstTmp("idreglab")) <> 0 Then
        txt_cb(1).Text = NulosN(RstTmp("idreglab"))
        txt_cb_Validate 1, False
    End If
    '--NIVEL EDUCATIVO
    If NulosN(RstTmp("idnivedu")) <> 0 Then
        txt_cb(2).Text = NulosN(RstTmp("idnivedu"))
        txt_cb_Validate 2, False
    End If
    '--DISCAPACIDAD
    If NulosN(RstTmp("discapacidad")) = 0 Then opt_discapacidad(0).Value = True
    If NulosN(RstTmp("discapacidad")) = 1 Then opt_discapacidad(1).Value = True
    '--OCUPACION
    If NulosN(RstTmp("idocu")) <> 0 Then
        txt_cb(3).Text = NulosN(RstTmp("idocu"))
        txt_cb_Validate 3, False
    End If
    '--REGIMEN PENSIONARIO
    If NulosN(RstTmp("idregpen")) <> 0 Then
        txt_cb(4).Text = NulosN(RstTmp("idregpen"))
        txt_cb_Validate 4, False
    End If
    If IsDate(RstTmp("fchinsregpen")) = True Then
        txtfecha(0).Valor = CDate(RstTmp("fchinsregpen"))
    End If
    txt(1).Text = NulosC(RstTmp("cuspp"))
    '--SCTR SALUD
    If NulosN(RstTmp("sctrsalud")) <> 0 Then
        txt_cb(5).Text = NulosN(RstTmp("sctrsalud"))
        txt_cb_Validate 5, False
    End If
    '--SCTR PENSION
    If NulosN(RstTmp("sctrpension")) <> 0 Then
        txt_cb(6).Text = NulosN(RstTmp("sctrpension"))
        txt_cb_Validate 6, False
    End If
    '************************************************************TAB 1
    '--TIPO CONTRATO
    If NulosN(RstTmp("idtipcon")) <> 0 Then
        txt_cb(7).Text = NulosN(RstTmp("idtipcon"))
        txt_cb_Validate 7, False
    End If
    '--TRABAJADOR SUJETO A
    chk(0).Value = Abs(NulosN(RstTmp("opc1"))) '--TRAJADOR SUJETO A REGIMEN ALTERNATIVO, ACUMULAT...
    chk(1).Value = Abs(NulosN(RstTmp("opc2"))) '--JORNADA DE TRABAJO MAXIMA
    chk(2).Value = Abs(NulosN(RstTmp("opc3"))) '--TRABAJO EN HORARIO NOCTURNO
    '--INFORMO RENTAS DE 5TA CATEGORIA
    If NulosN(RstTmp("opc4")) = 0 Then
        opt_Ingreso5ta(0).Value = True
    Else
        opt_Ingreso5ta(1).Value = True
    End If
    '--ES SINDICALIZADO
    If NulosN(RstTmp("opc5")) = 0 Then
        opt_sindato(0).Value = True
    Else
        opt_sindato(1).Value = True
    End If
    '--TIENE RENTAS DE 5TA CAT.. EXONERADAS O INAFECTAS
    If NulosN(RstTmp("opc6")) = 0 Then
        opt_Renta5ta(0).Value = True
    Else
        opt_Renta5ta(1).Value = True
    End If
    
    '--PRESTACIONES DE SALUD
        '--AFILIADO A E´PS/SERVICIO PROPIOS
    If NulosN(RstTmp("opc7")) = 0 Then
        opt_eps(0).Value = True
    Else
        opt_eps(1).Value = True
    End If
    
    If NulosN(RstTmp("ideps")) <> 0 Then
        txt_cb(8).Text = NulosN(RstTmp("ideps"))
        txt_cb_Validate 8, False
    End If
    '--SITUACION DE TRABAJADOR
    If NulosN(RstTmp("idsituacion")) <> 0 Then
        txt_cb(9).Text = NulosN(RstTmp("idsituacion"))
        txt_cb_Validate 9, False
    End If
    '--SITUACION ESPECIAL
    If NulosN(RstTmp("idsittra")) <> 0 Then
        txt_cb(10).Text = NulosN(RstTmp("idsittra"))
        txt_cb_Validate 10, False
    End If
    '--PERIODICIDAD
    If NulosN(RstTmp("idperiocidad")) <> 0 Then
        txt_cb(11).Text = NulosN(RstTmp("idperiocidad"))
        txt_cb_Validate 11, False
    End If
    '--TIPO DE PAGO
    If NulosN(RstTmp("idtippag")) <> 0 Then
        txt_cb(12).Text = NulosN(RstTmp("idtippag"))
        txt_cb_Validate 12, False
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
    
    If MsgBox("Seguro desea " + IIf(Quehace = 1, "Grabar", "Modificar") + " los datos del Trabajador", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim xId As Integer

    On Error GoTo LaCague

    xCon.BeginTrans

    '*****************************************************
    If Quehace = 1 Then
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_categoria1 ; ", xCon
        RstCab.AddNew
    Else
        RST_Busq RstCab, "SELECT * FROM pla_categoria1 WHERE idemp =  " & mIdEmpleado & " ;", xCon
    End If
    
    RstCab("idemp") = mIdEmpleado
    '************************************************************TAB 0
    '--TIPO TRABAJADOR
    RstCab("idtiptra") = NulosN(txt_cb(0).Text)
    '--REGIMEN LABORAL
    RstCab("idreglab") = NulosN(txt_cb(1).Text)
    '--NIVEL EDUCATIVO
    RstCab("idnivedu") = NulosN(txt_cb(2).Text)
    '--DISCAPACIDAD
    If opt_discapacidad(0).Value = True Then RstCab("discapacidad") = 0
    If opt_discapacidad(1).Value = True Then RstCab("discapacidad") = 1
    '--OCUPACION
    RstCab("idocu") = NulosN(txt_cb(3).Text)
    '--REGIMEN PENSIONARIO
    RstCab("idregpen") = NulosN(txt_cb(4).Text)
    If IsDate(txtfecha(0).Valor) = True Then
        RstCab("fchinsregpen") = CDate(txtfecha(0).Valor)
    End If
    RstCab("cuspp") = Trim(txt(1).Text)
    '--SCTR SALUD
    RstCab("sctrsalud") = NulosN(txt_cb(5).Text)
    '--SCTR PENSION
    RstCab("sctrpension") = NulosN(txt_cb(6).Text)
    '************************************************************TAB 1
    '--TIPO CONTRATO
    RstCab("idtipcon") = NulosN(txt_cb(7).Text)
    '--TRABAJADOR SUJETO A
    RstCab("opc1") = IIf(chk(0).Value = 0, 0, 1) '--TRAJADOR SUJETO A REGIMEN ALTERNATIVO, ACUMULAT...
    RstCab("opc2") = IIf(chk(1).Value = 0, 0, 1) '--JORNADA DE TRABAJO MAXIMA
    RstCab("opc3") = IIf(chk(2).Value = 0, 0, 1) '--TRABAJO EN HORARIO NOCTURNO
    RstCab("opc4") = IIf(opt_Ingreso5ta(0).Value = True, 0, 1) '--INFORMO RENTAS DE 5TA CATEGORIA
    RstCab("opc5") = IIf(opt_sindato(0).Value = True, 0, 1) '--ES SINDICALIZADO
    RstCab("opc6") = IIf(opt_Renta5ta(0).Value = True, 0, 1) '--TIENE RENTAS DE 5TA CAT.. EXONERADAS O INAFECTAS
    '--PRESTACIONES DE SALUD
    RstCab("opc7") = IIf(opt_eps(0).Value = True, 0, 1) '--AFILIADO A E´PS/SERVICIO PROPIOS
    
    RstCab("ideps") = NulosN(txt_cb(8).Text)
    '--SITUACION DE TRABAJADOR
    RstCab("idsituacion") = NulosN(txt_cb(9).Text)
    '--SITUACION ESPECIAL
    RstCab("idsittra") = NulosN(txt_cb(10).Text)
    '--PERIODICIDAD
    RstCab("idperiocidad") = NulosN(txt_cb(11).Text)
    '--TIPO DE PAGO
    RstCab("idtippag") = NulosN(txt_cb(12).Text)
    
    '--
    RstCab.Update

    MsgBox "Los datos del Trabajador " + IIf(Quehace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

    xCon.CommitTrans
    Set RstCab = Nothing
    Grabar = True
    Exit Function

LaCague:
    Set RstCab = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar los datos del empleado por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Private Function fValidarDatos() As Boolean
    
    Dim band As Integer
    
    band = Validar(txt_cb)
    If band <> -1 Then
        MsgBox "Llene el Campo de " & lbl_capt(band).Caption, vbInformation, xTitulo
        If band >= 0 And band <= 6 Then '--TAB 0
            TabOne1.CurrTab = 0
        Else            '--TAB 1
            TabOne1.CurrTab = 1
        End If
        txt_cb(band).SetFocus
        Exit Function
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

Private Sub txtfecha_Validate(Index As Integer, Cancel As Boolean)
    If IsDate(txtfecha(Index)) = True Then
        If txt(1).Visible = True Then
            txt(1).SetFocus
        Else
            txt_cb(5).SetFocus
        End If
    End If
End Sub
