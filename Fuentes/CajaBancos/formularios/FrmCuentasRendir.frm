VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmCuentasRendir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja y Bancos - Cuentas por Rendir"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7245
      Left            =   15
      TabIndex        =   20
      Top             =   375
      Width           =   11880
      _cx             =   20955
      _cy             =   12779
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
         BorderStyle     =   0  'None
         Height          =   6825
         Left            =   12525
         TabIndex        =   23
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdModificar 
            Caption         =   "&Cancelar"
            Height          =   375
            Index           =   1
            Left            =   9945
            TabIndex        =   68
            Top             =   5130
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.CommandButton CmdModificar 
            Caption         =   "&Modificar"
            Height          =   375
            Index           =   0
            Left            =   8490
            TabIndex        =   67
            Top             =   5130
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Frame Frame3 
            Caption         =   "( Periodo )"
            Height          =   720
            Left            =   9480
            TabIndex        =   65
            Top             =   225
            Width           =   2010
            Begin VB.Label lblperiodo 
               Alignment       =   2  'Center
               Caption         =   "lblperiodo(1)"
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
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   66
               Top             =   330
               Width           =   1740
            End
         End
         Begin VB.Frame fra_estado 
            Height          =   1080
            Left            =   8415
            TabIndex        =   53
            Top             =   3975
            Width           =   2955
            Begin VB.CommandButton cmd_estado 
               Caption         =   "&Rechazar"
               Height          =   300
               Index           =   1
               Left            =   1515
               TabIndex        =   59
               Top             =   690
               Width           =   1365
            End
            Begin VB.CommandButton cmd_estado 
               Caption         =   "&Aprobar"
               Height          =   300
               Index           =   0
               Left            =   120
               TabIndex        =   58
               Top             =   690
               Width           =   1365
            End
            Begin VB.Label LblEstado 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   465
               Index           =   1
               Left            =   2385
               TabIndex        =   57
               Top             =   180
               Visible         =   0   'False
               Width           =   510
            End
            Begin VB.Label LblEstado 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Pendiente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   450
               Index           =   0
               Left            =   120
               TabIndex        =   54
               Top             =   195
               Width           =   2775
            End
         End
         Begin VB.TextBox txt 
            Height          =   1095
            Index           =   3
            Left            =   480
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Tag             =   "null"
            Text            =   "FrmCuentasRendir.frx":0000
            Top             =   5670
            Width           =   10920
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   1995
            TabIndex        =   11
            Text            =   "txt(2)"
            Top             =   5115
            Width           =   1605
         End
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   4
            Left            =   2505
            Picture         =   "FrmCuentasRendir.frx":0009
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   4500
            Width           =   195
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   4
            Left            =   1995
            MaxLength       =   4
            TabIndex        =   9
            Text            =   "txt_cb(4)"
            Top             =   4470
            Width           =   735
         End
         Begin VB.TextBox txt 
            Height          =   315
            Index           =   1
            Left            =   1995
            TabIndex        =   10
            Text            =   "txt(1)"
            Top             =   4785
            Width           =   2070
         End
         Begin VB.Frame fr 
            Caption         =   "[ Del Beneficiario ]"
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
            Height          =   675
            Index           =   2
            Left            =   480
            TabIndex        =   37
            Top             =   3255
            Width           =   10965
            Begin VB.OptionButton opt_per 
               Caption         =   "Proveedor"
               Height          =   195
               Index           =   1
               Left            =   1200
               TabIndex        =   17
               Top             =   330
               Width           =   1050
            End
            Begin VB.OptionButton opt_per 
               Caption         =   "Persona"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   330
               Value           =   -1  'True
               Width           =   900
            End
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   3
               Left            =   5040
               Picture         =   "FrmCuentasRendir.frx":013B
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   3
               Left            =   3270
               MaxLength       =   12
               TabIndex        =   6
               Text            =   "txt_cb(3)"
               ToolTipText     =   "Ingrese DNI de la persona"
               Top             =   240
               Width           =   2010
            End
            Begin VB.Label lbl_cb_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb_cod(3)"
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
               Index           =   3
               Left            =   9465
               TabIndex        =   49
               Top             =   240
               Visible         =   0   'False
               Width           =   1395
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
               Height          =   300
               Index           =   3
               Left            =   5280
               TabIndex        =   39
               Top             =   240
               Width           =   5580
            End
            Begin VB.Label lbl_cb_capt 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Persona"
               Height          =   195
               Index           =   3
               Left            =   2655
               TabIndex        =   38
               Top             =   330
               Width           =   585
            End
         End
         Begin VB.Frame fr 
            Caption         =   "[ Del Destino ]"
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
            Height          =   675
            Index           =   1
            Left            =   480
            TabIndex        =   34
            Top             =   2520
            Width           =   10965
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   2
               Left            =   1365
               Picture         =   "FrmCuentasRendir.frx":026D
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   285
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   2
               Left            =   930
               MaxLength       =   4
               TabIndex        =   4
               Text            =   "txt_cb(2)"
               Top             =   255
               Width           =   675
            End
            Begin VB.Label lblCtaDestino 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblCtaDestino"
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
               Left            =   9525
               TabIndex        =   72
               Top             =   255
               Visible         =   0   'False
               Width           =   1290
            End
            Begin VB.Label lbl_cb_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb_cod(2)"
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
               Index           =   2
               Left            =   7740
               TabIndex        =   48
               Top             =   270
               Visible         =   0   'False
               Width           =   1425
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
               Height          =   300
               Index           =   2
               Left            =   1605
               TabIndex        =   36
               Top             =   255
               Width           =   7470
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Destino"
               Height          =   195
               Index           =   2
               Left            =   210
               TabIndex        =   35
               Top             =   345
               Width           =   540
            End
         End
         Begin VB.Frame fr 
            Caption         =   "[ Del Origen ]"
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
            Height          =   675
            Index           =   0
            Left            =   480
            TabIndex        =   31
            Top             =   1800
            Width           =   10965
            Begin VB.OptionButton opt_mov 
               Caption         =   "Banco"
               Height          =   195
               Index           =   1
               Left            =   1200
               TabIndex        =   14
               Top             =   360
               Width           =   840
            End
            Begin VB.OptionButton opt_mov 
               Caption         =   "Caja"
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   2
               Top             =   360
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.CommandButton cb 
               Height          =   240
               Index           =   0
               Left            =   5025
               Picture         =   "FrmCuentasRendir.frx":039F
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   285
               Width           =   225
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   0
               Left            =   3270
               MaxLength       =   20
               TabIndex        =   3
               Text            =   "txt_cb(0)"
               Top             =   255
               Width           =   2010
            End
            Begin VB.Label lblCtaOrigen 
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblCtaOrigen"
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
               Left            =   9495
               TabIndex        =   71
               Top             =   255
               Visible         =   0   'False
               Width           =   1290
            End
            Begin VB.Label lbl_cb_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb_cod(0)"
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
               Index           =   0
               Left            =   7470
               TabIndex        =   45
               Top             =   285
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Label lbl_cb_capt 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Origen"
               Height          =   195
               Index           =   0
               Left            =   2700
               TabIndex        =   32
               Top             =   360
               Width           =   465
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
               Height          =   300
               Index           =   0
               Left            =   5280
               TabIndex        =   33
               Top             =   255
               Width           =   5580
            End
         End
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   1
            Left            =   2325
            Picture         =   "FrmCuentasRendir.frx":04D1
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1485
            Width           =   210
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   1
            Left            =   1890
            MaxLength       =   4
            TabIndex        =   1
            Text            =   "txt_cb(1)"
            Top             =   1455
            Width           =   675
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0080FF80&
            Height          =   315
            Index           =   0
            Left            =   8325
            TabIndex        =   27
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   675
            Visible         =   0   'False
            Width           =   1170
         End
         Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
            Height          =   300
            Index           =   0
            Left            =   1890
            TabIndex        =   0
            Top             =   1140
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
         End
         Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
            Height          =   300
            Index           =   1
            Left            =   1995
            TabIndex        =   7
            Tag             =   "b"
            Top             =   4155
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
         End
         Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
            Height          =   300
            Index           =   2
            Left            =   5415
            TabIndex        =   8
            Tag             =   "b"
            Top             =   4155
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
         End
         Begin VB.Label LblTipCam2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Left            =   8835
            TabIndex        =   70
            Top             =   1560
            Width           =   1110
         End
         Begin VB.Label LblTipoCambio 
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
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   10020
            TabIndex        =   69
            Top             =   1455
            Width           =   1350
         End
         Begin VB.Label lblfch 
            AutoSize        =   -1  'True
            Caption         =   "Rendir Al"
            Height          =   195
            Index           =   2
            Left            =   4650
            TabIndex        =   63
            Top             =   4230
            Width           =   645
         End
         Begin VB.Label lbl_aut 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Autorizador:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   62
            Top             =   810
            Width           =   840
         End
         Begin VB.Label lbl_prog 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Programador:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   61
            Top             =   495
            Width           =   945
         End
         Begin VB.Label lbl_prog 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_prog(1)"
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
            Index           =   1
            Left            =   1890
            TabIndex        =   56
            Top             =   450
            Width           =   4665
         End
         Begin VB.Label lbl_aut 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_aut(1)"
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
            Index           =   1
            Left            =   1890
            TabIndex        =   55
            Top             =   765
            Width           =   4665
         End
         Begin VB.Label lbl_aut 
            BackColor       =   &H0000FFFF&
            Caption         =   "lbl_aut(0)"
            Height          =   270
            Index           =   0
            Left            =   6600
            TabIndex        =   52
            Top             =   765
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl_prog 
            BackColor       =   &H0000FFFF&
            Caption         =   "lbl_prog(0)"
            Height          =   270
            Index           =   0
            Left            =   6615
            TabIndex        =   51
            Top             =   450
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl_cb_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb_cod(4)"
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
            Index           =   4
            Left            =   5505
            TabIndex        =   50
            Top             =   4470
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label lbl_cb_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb_cod(1)"
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
            Index           =   1
            Left            =   4155
            TabIndex        =   47
            Top             =   1455
            Visible         =   0   'False
            Width           =   1230
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
            Height          =   300
            Index           =   1
            Left            =   2550
            TabIndex        =   46
            Top             =   1455
            Width           =   1830
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Observación"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   44
            Top             =   5445
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   43
            Top             =   5175
            Width           =   525
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
            Height          =   300
            Index           =   4
            Left            =   2730
            TabIndex        =   42
            Top             =   4470
            Width           =   3945
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Documento"
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   41
            Top             =   4560
            Width           =   1410
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "N° de Documento"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   40
            Top             =   4875
            Width           =   1275
         End
         Begin VB.Label lblfch 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Pago"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   30
            Top             =   4230
            Width           =   870
         End
         Begin VB.Label lblfch 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Emisión"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   29
            Top             =   1185
            Width           =   1035
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   28
            Top             =   1500
            Width           =   585
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "codigo"
            Height          =   195
            Index           =   0
            Left            =   7725
            TabIndex        =   26
            Top             =   795
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle del la Cuenta por Rendir"
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
            Left            =   0
            TabIndex        =   24
            Top             =   30
            Width           =   11550
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6825
         Left            =   45
         TabIndex        =   21
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   6465
            Left            =   15
            TabIndex        =   25
            Top             =   345
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11404
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Num.Reg"
            Columns(0).DataField=   "numreg"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tipo Mov."
            Columns(1).DataField=   "tipmov"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Emi."
            Columns(2).DataField=   "fchemi"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "T.Doc"
            Columns(3).DataField=   "abrev"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "N° Doc"
            Columns(4).DataField=   "numdoc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "M"
            Columns(5).DataField=   "simbolo"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Importe"
            Columns(6).DataField=   "imp"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Saldo"
            Columns(7).DataField=   "saldo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Emitido por"
            Columns(8).DataField=   "prog"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "T. Persona"
            Columns(9).DataField=   "tipper"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Entreg. A"
            Columns(10).DataField=   "benef"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Autoriz. por"
            Columns(11).DataField=   "aut"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Estado"
            Columns(12).DataField=   "estdesc"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   13
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=13"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1561"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1482"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1667"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1588"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1561"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1482"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1085"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1005"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1746"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1667"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=794"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=714"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=1455"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1376"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=1588"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1508"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1826"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1746"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(55)=   "Column(9).Width=1799"
            Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1720"
            Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=516"
            Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(61)=   "Column(10).Width=1561"
            Splits(0)._ColumnProps(62)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(63)=   "Column(10)._WidthInPix=1482"
            Splits(0)._ColumnProps(64)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(65)=   "Column(10)._ColStyle=516"
            Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(67)=   "Column(11).Width=1879"
            Splits(0)._ColumnProps(68)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(69)=   "Column(11)._WidthInPix=1799"
            Splits(0)._ColumnProps(70)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(71)=   "Column(11)._ColStyle=516"
            Splits(0)._ColumnProps(72)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(73)=   "Column(12).Width=1693"
            Splits(0)._ColumnProps(74)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(75)=   "Column(12)._WidthInPix=1614"
            Splits(0)._ColumnProps(76)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(77)=   "Column(12)._ColStyle=516"
            Splits(0)._ColumnProps(78)=   "Column(12).Order=13"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80&"
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=90,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=87,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=88,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=89,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=32,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=78,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=82,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
            _StyleDefs(88)  =   "Named:id=33:Normal"
            _StyleDefs(89)  =   ":id=33,.parent=0"
            _StyleDefs(90)  =   "Named:id=34:Heading"
            _StyleDefs(91)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(92)  =   ":id=34,.wraptext=-1"
            _StyleDefs(93)  =   "Named:id=35:Footing"
            _StyleDefs(94)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(95)  =   "Named:id=36:Selected"
            _StyleDefs(96)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=37:Caption"
            _StyleDefs(98)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(99)  =   "Named:id=38:HighlightRow"
            _StyleDefs(100) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(101) =   "Named:id=39:EvenRow"
            _StyleDefs(102) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(103) =   "Named:id=40:OddRow"
            _StyleDefs(104) =   ":id=40,.parent=33"
            _StyleDefs(105) =   "Named:id=41:RecordSelector"
            _StyleDefs(106) =   ":id=41,.parent=34"
            _StyleDefs(107) =   "Named:id=42:FilterBar"
            _StyleDefs(108) =   ":id=42,.parent=33"
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
            TabIndex        =   64
            Top             =   75
            Width           =   1980
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Cuentas por Rendir"
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
            Left            =   15
            TabIndex        =   22
            Top             =   30
            Width           =   11550
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1005
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
         Left            =   5535
         Top             =   60
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
               Picture         =   "FrmCuentasRendir.frx":0603
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCuentasRendir.frx":0B47
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCuentasRendir.frx":0ED9
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCuentasRendir.frx":105D
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCuentasRendir.frx":14B1
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCuentasRendir.frx":15C9
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCuentasRendir.frx":1B0D
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCuentasRendir.frx":2051
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCuentasRendir.frx":2165
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCuentasRendir.frx":2279
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCuentasRendir.frx":26CD
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCuentasRendir.frx":2839
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar producto           "
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar producto"
      End
   End
End
Attribute VB_Name = "FrmCuentasRendir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstFrm As New ADODB.Recordset
Dim Agregando As Boolean

Dim fOcultarToolbar As Boolean  '--FALSE::SE OCULTA TRUE::MOSTRAR

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta

'--de los estados
'LblEstado(1).Caption = "2"
'LblEstado(1).Caption = "4"
'

Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nOrden As String
    Dim nCampoBusca As String
    Dim nSQL As String
    
    If QueHace = 3 Then Exit Sub
    '----------
'    On Error GoTo error
    Select Case Index
    
        Case 0 '--ORIGEN
            If fValidarMoneda() = False Then Exit Sub
            
            If opt_mov(0).Value = True Then
                ReDim xCampos(4, 3) As String
                xCampos(0, 0) = "Nombre":       xCampos(0, 1) = "nombre":    xCampos(0, 2) = "3500":   xCampos(0, 3) = "C"
                xCampos(1, 0) = "Cuenta":       xCampos(1, 1) = "cuenta":    xCampos(1, 2) = "3000":   xCampos(1, 3) = "C"
                xCampos(2, 0) = "N° Cuenta":    xCampos(2, 1) = "numcta":    xCampos(2, 2) = "1200":   xCampos(2, 3) = "C"
                xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":        xCampos(3, 2) = "450":   xCampos(3, 3) = "N"
                nTitulo = "Buscando Origen"
                nSQL = "SELECT con_destino.id, con_destino.descripcion AS nombre, con_destino.id AS cod, con_planctas.descripcion AS cuenta, con_planctas.cuenta AS numcta, con_destino.idcuen as idcta " _
                    + vbCr + " FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuen " _
                    + vbCr + " WHERE con_destino.idmon=" + CStr(lbl_cb_cod(1).Caption)
                nCampoBusca = "nombre"
                
            Else
                ReDim xCampos(4, 3) As String
                xCampos(0, 0) = "N° de Cuenta":     xCampos(0, 1) = "numcue":    xCampos(0, 2) = "1500":   xCampos(0, 3) = "C"
                xCampos(1, 0) = "Banco":            xCampos(1, 1) = "nombre":    xCampos(1, 2) = "2000":   xCampos(1, 3) = "C"
                xCampos(2, 0) = "Cuenta":           xCampos(2, 1) = "cuenta":    xCampos(2, 2) = "3500":   xCampos(2, 3) = "C"
                xCampos(3, 0) = "N° Cuenta":        xCampos(3, 1) = "numcta":    xCampos(3, 2) = "1200":   xCampos(3, 3) = "C"
                
                nTitulo = "Buscando N° de Cuenta"
                nSQL = "SELECT con_bancocuenta.numcue, mae_bancos.descripcion AS nombre, con_bancocuenta.id AS cod, con_planctas.descripcion AS cuenta, con_planctas.cuenta AS numcta, con_bancocuenta.idcuen as idcta " _
                    + vbCr + " FROM con_planctas RIGHT JOIN (mae_bancos RIGHT JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban) ON con_planctas.id = con_bancocuenta.idcuen " _
                    + vbCr + " WHERE con_bancocuenta.idmon=" + CStr(lbl_cb_cod(1).Caption)
                    
                nCampoBusca = "numcue"
            End If
           
        Case 1 '--MONEDA
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Moneda":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "3500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Símbolo":   xCampos(1, 1) = "simbolo":    xCampos(1, 2) = "1000":   xCampos(1, 3) = "C"
            nTitulo = "Buscando Moneda"
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion as nombre,mae_moneda.id as cod,mae_moneda.simbolo  " _
                + vbCr + " From mae_moneda "
            nCampoBusca = "nombre"
            
        Case 2 '--DESTINO
        
            If fValidarMoneda() = False Then Exit Sub
        
            ReDim xCampos(4, 3) As String
            xCampos(0, 0) = "Destino":      xCampos(0, 1) = "nombre":    xCampos(0, 2) = "3200":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Cuenta":       xCampos(1, 1) = "cuenta":    xCampos(1, 2) = "3000":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "N° Cuenta":    xCampos(2, 1) = "numcta":    xCampos(2, 2) = "1200":   xCampos(2, 3) = "C"
            xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":        xCampos(3, 2) = "450":   xCampos(3, 3) = "N"
            
            nTitulo = "Buscando Destinos"

             nSQL = "SELECT con_destino.id, con_destino.descripcion AS nombre, con_destino.id AS cod, con_planctas.descripcion AS cuenta, con_planctas.cuenta AS numcta,con_destino.idcuen as idcta " _
                + vbCr + " FROM con_planctas RIGHT JOIN con_destino ON con_planctas.id = con_destino.idcuen" _
                + vbCr + "WHERE (((con_destino.idmon)=" + lbl_cb_cod(1).Caption + ") AND ((con_destino.tipmov)=2)) and con_destino.rendir = -1 ;"
            
            nCampoBusca = "nombre"
            
        Case 3 '--BENEFICIARIO
            ReDim xCampos(2, 3) As String
                        
            If Me.opt_per(0).Value = True Then '--PERSONA
                xCampos(0, 0) = "Empleado":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5500":   xCampos(0, 3) = "C"
                xCampos(1, 0) = "dni":        xCampos(1, 1) = "numdoc":    xCampos(1, 2) = "1500":   xCampos(1, 3) = "C"
                nTitulo = "Buscando Personas"
                nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nombre, con_emptes.id AS cod " _
                    & " FROM pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp " _
                    & " WHERE (((con_emptes.id)<>" & lbl_prog(0).Caption & "))"
            
            Else '--PROVEEDOR
                xCampos(0, 0) = "Proveedor":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5500":   xCampos(0, 3) = "C"
                xCampos(1, 0) = "RUC":         xCampos(1, 1) = "numruc":    xCampos(1, 2) = "1500":   xCampos(1, 3) = "C"
                nTitulo = "Buscando Proveedores"
                nSQL = "SELECT  mae_prov.numruc, mae_prov.nombre ,mae_prov.id " _
                + vbCr + " From mae_prov "
            
            End If
            nCampoBusca = "nombre"
            
        Case 4 '--TIPO DE DOCUMENTO
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Tipo Documento":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "3500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Abrev":   xCampos(1, 1) = "abrev":    xCampos(1, 2) = "1000":   xCampos(1, 3) = "C"
       
            nTitulo = "Buscando Tipo de Documento"
            
            nSQL = "SELECT mae_doccajaban.id, mae_doccajaban.descripcion as nombre, mae_doccajaban.id AS cod, mae_doccajaban.abrev " _
                + vbCr + " From mae_doccajaban " _
                + vbCr + " WHERE (((mae_doccajaban.tipo)=" + IIf(opt_mov(0).Value = True, "1", "2") + "));"
            
            nCampoBusca = "nombre"
            
    End Select
    nOrden = "nombre"
    
    Dim xRs As New ADODB.Recordset

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, nOrden, nCampoBusca, Principio
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    
    Select Case Index
        Case 0 '--origen
            lblCtaOrigen.Caption = NulosN(xRs.Fields("idcta"))
            txt_cb(2).SetFocus
        Case 1 '--moneda
            opt_mov(0).SetFocus
        Case 2 '--destino
            opt_per(0).SetFocus
            lblCtaDestino.Caption = NulosN(xRs.Fields("idcta"))
        Case 3 '--beneficiario
            TxtFecha(1).SetFocus
        Case 4 '--tipo doc
            If txt(1).Enabled = True Then
                txt(1).SetFocus
            Else
                txt(2).SetFocus
            End If
    End Select
Salir:
    Set xRs = Nothing
Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR
End Sub

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub


Private Sub cmd_estado_Click(Index As Integer)
    Dim mIdEstado As Integer
    If QueHace = 1 Then
        MsgBox "Primero guarde el registro" + vbCr + "Luego proceda a " + cmd_estado(Index).Caption, vbInformation, xTitulo
        Exit Sub
    End If
    If RstFrm.EOF = True Or RstFrm.BOF = True Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    If IsNull(RstFrm.Fields("idaut")) = False Then
        If NulosN(RstFrm.Fields("idaut")) <> 0 And NulosN(RstFrm.Fields("idaut")) <> NulosN(lbl_aut(0).Caption) Then
            If MsgBox("Este registro ha sido autorizado por otra persona " + vbCr + "Autorizador: " + RstFrm.Fields("aut") & "" + vbCr + "Desea continuar", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
        End If
    End If
    
    mIdEstado = NulosN(LblEstado(1).Caption)
    
    If Index = 1 And LblEstado(1).Caption = "2" Then
        '--VALIDAR QUE NO
        Dim xRs As New ADODB.Recordset
        Dim nEstado As Integer '-- -1::NO ESTA EN NINGUNA TABLA, 0::ESTA EN CAJABANCO, 1::ESTA EN DEVOLUCIONES, 2 ESTA EN AMBOS
        'On Error GoTo error
        '--VALIDAR SI ESTA EN DEVOLUCIONES
        RST_Busq xRs, "SELECT con_devoluciones.idren From con_devoluciones WHERE (((con_devoluciones.idren)=" & NulosN(RstFrm("id")) & "));", xCon
        
        If xRs.RecordCount > 0 Then nEstado = 1
        
        Set xRs = Nothing
        If nEstado <> 0 Then
            MsgBox "No puede Continuar" + vbCr + "Pues el registro esta asociado a: Devoluciones", vbExclamation, xTitulo
            Exit Sub
        End If
    End If
    '************
    If Index = 0 Then
        PONER_COLOR_ESTADO xCon, LblEstado(0), 2 '--APROBADO
        LblEstado(1).Caption = "2"
    ElseIf Index = 1 Then
        PONER_COLOR_ESTADO xCon, LblEstado(0), 4 '--RECHAZADO
        LblEstado(1).Caption = "4"
    End If
    
    If fValidarDatos(True) = False Then GoTo Salir
    If MsgBox("Seguro desea modificar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then GoTo Salir
    
    If Grabar(True) = True Then
        Dim mCod As Variant
        mCod = RstFrm.Fields("id")
        RstFrm.Requery
        If RstFrm.RecordCount <> 0 Then
            RstFrm.MoveFirst
            RstFrm.Find "id=" & mCod
        End If
        TabOne1.CurrTab = 0
    End If
    '************
    Exit Sub
Salir:
    PONER_COLOR_ESTADO xCon, LblEstado(0), mIdEstado   '--RECHAZADO
    LblEstado(1).Caption = mIdEstado
    Exit Sub
error:
    Set xRs = Nothing
    PONER_COLOR_ESTADO xCon, LblEstado(0), mIdEstado  '--RECHAZADO
    LblEstado(1).Caption = mIdEstado
    SHOW_ERROR
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    Dim Rpta As Integer
    Blanquea
    SeEjecuto = False
    pCargarGrid
    SeEjecuto = True
    If RstFrm.RecordCount = 0 Then
        If fOcultarToolbar = False Then Exit Sub
        If MsgBox("No se ha registrado ninguna cuenta por rendir, ¿Desea agergar uno ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbYes Then
            Nuevo
        End If
    End If
End Sub

Private Sub pCargarGrid()
    Dim nSQL  As String
    
    LblPeriodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
    LblPeriodo(1).Caption = LblPeriodo(0).Caption
    

    nSQL = "SELECT con_ctasrendir.numreg, con_ctasrendir.id, con_ctasrendir.fchemi, con_ctasrendir.fchpag, con_ctasrendir.fchren, con_ctasrendir.numdoc, " _
        + vbCr + " IIf(con_ctasrendir.tipmov=1,'Caja','Banco') AS tipmov, con_ctasrendir.idmon, mae_moneda.descripcion AS monnom, mae_moneda.simbolo, " _
        + vbCr + " con_ctasrendir.idori, " _
        + vbCr + " IIf(con_ctasrendir.tipmov=1,(con_ctasrendir.idori),(SELECT  [con_bancocuenta].[numcue]  AS origen FROM mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban WHERE (((con_bancocuenta.id)=con_ctasrendir.idori)) )) AS origen_num, " _
        + vbCr + " IIf(con_ctasrendir.tipmov=1,(SELECT destino.descripcion FROM con_destino AS destino WHERE (((destino.id)=con_ctasrendir.idori)) ),(SELECT [mae_bancos].[descripcion] AS origen FROM mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban WHERE (((con_bancocuenta.id)=con_ctasrendir.idori)) )) AS origen, " _
        + vbCr + " con_ctasrendir.iddes, con_destino.descripcion AS destino, con_ctasrendir.tipdoc, mae_doccajaban.descripcion AS tipdocnom, mae_doccajaban.abrev, IIf(con_ctasrendir.tipper=1,'Persona','Proveedor') AS tipper, con_ctasrendir.idper, " _
        + vbCr + " IIf(con_ctasrendir.tipper=1,(SELECT [pla_empleados].[nom] & ' ' & [pla_empleados].[ape] AS nombre FROM pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp WHERE (((con_emptes.id)=con_ctasrendir.idper))),(SELECT mae_prov.nombre FROM mae_prov WHERE (((mae_prov.id)=con_ctasrendir.idper)) )) AS benef, " _
        + vbCr + " IIf(con_ctasrendir.tipper=1,(SELECT pla_empleados.numdoc  FROM pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp WHERE (((con_emptes.id)=con_ctasrendir.idper))),(SELECT mae_prov.numruc FROM mae_prov WHERE (((mae_prov.id)=con_ctasrendir.idper)) )) AS docper, " _
        + vbCr + " con_ctasrendir.[imp],con_ctasrendir.[saldo], con_ctasrendir.obs, con_ctasrendir.idprog, pla_empleados.nom & ' ' & pla_empleados.ape AS prog, con_ctasrendir.idaut, pla_empleados_1.ape & ' ' & pla_empleados_1.nom AS aut, con_ctasrendir.idest, mae_estados.descripcion AS estdesc, " _
        + vbCr + " IIF(con_ctasrendir.numreg IS NULL OR con_ctasrendir.numreg='','', FORMAT (con_ctasrendir.idmes,'00') & IIF(mae_libros.codsun IS NULL OR mae_libros.codsun ='','FF',mae_libros.codsun) & MID(con_ctasrendir.numreg,3)) AS numreg " _
        + vbCr + " FROM (pla_empleados RIGHT JOIN (mae_moneda RIGHT JOIN (mae_doccajaban RIGHT JOIN (con_emptes RIGHT JOIN (mae_estados RIGHT JOIN (con_destino RIGHT JOIN (con_ctasrendir LEFT JOIN (con_emptes AS con_emptes_1 LEFT JOIN pla_empleados AS pla_empleados_1 ON con_emptes_1.idemp = pla_empleados_1.id) ON con_ctasrendir.idaut = con_emptes_1.id) ON con_destino.id = con_ctasrendir.iddes) ON mae_estados.id = con_ctasrendir.idest) ON con_emptes.id = con_ctasrendir.idprog) ON mae_doccajaban.id = con_ctasrendir.tipdoc) ON mae_moneda.id = con_ctasrendir.idmon) ON pla_empleados.id = con_emptes.idemp) LEFT JOIN mae_libros ON con_ctasrendir.idlib = mae_libros.id " _
        + vbCr + " WHERE con_ctasrendir.ano = " & AnoTra & " And con_ctasrendir.idmes IN (-1," & xMes & ")" _
        + vbCr + " ORDER BY con_ctasrendir.fchemi ASC "
    
    TabOne1.CurrTab = 0
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    Agregando = False
    QueHace = 3
    Dg3.BatchUpdates = False
    Dg3.Columns("fchemi").NumberFormat = FORMAT_DATE
    Dg3.Columns("imp").NumberFormat = FORMAT_MONTO
    Dg3.Columns("saldo").NumberFormat = FORMAT_MONTO
   
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    '--SI NO ES PROGRAMADOR
    If fVerificarProgAut(True, lbl_prog(0), lbl_prog(1)) = False Then    '--PROGAMADOR
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        fOcultarToolbar = False
    Else
        fOcultarToolbar = True
    End If
    '--
    Habilitar_Obj False
    '--AUTORIZADOR
    If fVerificarProgAut(False, lbl_aut(0), lbl_aut(1)) = True Then
        habilitar cmd_estado, True
        Ocultar CmdModificar, True
        CmdModificar(1).Enabled = False
    Else
        Ocultar CmdModificar, False
        habilitar cmd_estado, False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set Dg3.DataSource = Nothing
End Sub




Private Sub opt_mov_Click(Index As Integer)
    On Error GoTo error
    If QueHace = 3 Then Exit Sub
    If opt_mov(0).Value = True Then '--CAJA
        lbl(1).Caption = "N° de Documento"
        txt(1).Enabled = False
        If QueHace = 1 Then txt(1).Text = Format(HallaCodigoTabla("con_ctasrendir", xCon, "id"), "0000")
        If QueHace = 2 Then txt(1).Text = Format(RstFrm.Fields("id"), "0000")
    Else '--BANCO
        txt(1).Enabled = True
        txt(1).Text = ""
        lbl(1).Caption = "N° Cheque"
    End If
    txt_cb(0).Text = "":    lbl_cb(0).Caption = "":    lbl_cb_cod(0).Caption = ""
    txt_cb(4).Text = "":    lbl_cb(4).Caption = "":    lbl_cb_cod(4).Caption = ""
    Exit Sub
error:
    SHOW_ERROR
End Sub

Private Sub opt_mov_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        txt_cb(0).SetFocus
    End If
End Sub

Private Sub opt_per_Click(Index As Integer)
    lbl_cb_capt(3).Caption = opt_per(Index).Caption
    If Index = 0 Then txt_cb(3).ToolTipText = "Ingrese DNI de la persona"
    If Index = 1 Then txt_cb(3).ToolTipText = "Ingrese RUC del proveedor"
    txt_cb(3).Text = "":    lbl_cb(3).Caption = "":    lbl_cb_cod(3).Caption = ""
End Sub

Private Sub opt_per_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        txt_cb(3).SetFocus
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
    ElseIf OldTab = 1 Then
        QueHace = 3
    End If
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar(False) = True Then
            Cancelar
            If RstFrm.State = 0 Then Exit Sub
            RstFrm.Requery
            Dg3.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then
        If RstFrm.State = 0 Then Exit Sub
        RstFrm.Filter = ""
    End If
    If Button.Index = 10 Then Buscar
    If Button.Index = 11 Then CambiarMes
    
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Eliminar()
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.BOF = True Or RstFrm.EOF = True Or RstFrm.RecordCount = 0 Then Exit Sub

    If RstFrm.Fields("idest") <> "1" Then
        MsgBox "No puede Eliminar" + vbCr + "Ya fue " + LblEstado(0).Caption, vbInformation, xTitulo
        Exit Sub
    End If
    
    Dim Rpta As Integer
    
    Rpta = MsgBox("¿Esta seguro de eliminar El registro seleccionado?", vbQuestion + vbYesNo, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETe * FROM con_ctasrendir WHERE id = " & RstFrm("id") & ""
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
        RstFrm.Requery
        TabOne1.CurrTab = 0
        Dg3.Refresh
        If RstFrm.RecordCount = 0 Then
            Rpta = MsgBox("No hay ningún registro, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstFrm = Nothing
                Unload Me
                Exit Sub
            End If
        End If
    End If
End Sub

Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    Habilitar_Obj False
    Label1.Caption = "Detalle del la Cuenta por Rendir"
    TabOne1.CurrTab = 0
    '-----
    fra_estado.Visible = True
    Ocultar CmdModificar, True
    '------
    Dg3.SetFocus
End Sub

Private Sub Modificar()

    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    If NulosN(RstFrm.Fields("idest")) = 2 Then
        MsgBox "No puede Modificar" + vbCr + "Ya fue Aprobado" + vbCr + "Solo puede modificar algunos campos", vbInformation, xTitulo
        If TabOne1.CurrTab = 1 Then CmdModificar(0).SetFocus
        Exit Sub
    End If
    '--VER SI EL PROGRAMADOR ES EL MISMO AL QUE CREO EL  REGISTRO
    fVerificarProgAut True, lbl_prog(0), lbl_prog(1)
    If RstFrm.Fields("idprog") & "" <> lbl_prog(0).Caption Then
        MsgBox "Ust. no ha Programado este registro" + vbCr + "Sólo puede modificarlo quien lo programó", vbInformation, xTitulo
        Exit Sub
    End If
    '------
    
    QueHace = 2
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool

    Habilitar_Obj True
    MuestraSegundoTab
    
    Label1.Caption = "Modificando Cuentas por Rendir"
    
    fVerificarProgAut True, lbl_prog(0), lbl_prog(1)   '--PROGRAMADOR
    lbl_aut(0).Caption = ""
    lbl_aut(1).Caption = ""
    '-----
    fra_estado.Visible = False
    Ocultar CmdModificar, False
    '------
    TxtFecha(0).SetFocus
End Sub

Sub MuestraSegundoTab()
    On Error GoTo error
    Dim QueHaceTmp As Integer
    With RstFrm
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Then Exit Sub
        txt(0).Text = .Fields("id") & "" '--CODIGO
        TxtFecha(0).Valor = .Fields("fchemi")  '--FECHA DE EMISION
        txtfecha_Validate 0, True
        '--DEL ORIGEN
        QueHaceTmp = QueHace
        QueHace = -1
        '--DE LA MONEDA
        If NulosN(.Fields("idmon")) <> 0 Then
            Me.txt_cb(1).Text = NulosN(.Fields("idmon"))
            txt_cb_Validate 1, False
        End If
        If LCase(.Fields("tipmov")) = "caja" Then '--TIPO DE MOVIENTO
            Me.opt_mov(0).Value = True '--ES CAJA
        Else
            Me.opt_mov(1).Value = True '--ES BANCO
        End If
        If NulosC(.Fields("origen_num")) <> "" Then
            txt_cb(0).Text = NulosC(.Fields("origen_num"))
            txt_cb_Validate 0, False
        End If
        '--DEL DESTINO
        If NulosN(.Fields("iddes")) <> 0 Then
            Me.txt_cb(2).Text = NulosN(.Fields("iddes"))
            txt_cb_Validate 2, False
        End If
        
        '--DEL BENEFICIARIO
        If LCase(.Fields("tipper")) = "persona" Then '--TIPO DE PERSONA
            Me.opt_per(0).Value = True  '--ES PERSONA
        Else
            Me.opt_per(1).Value = True '--ES PROVEEDOR
        End If
        Me.txt_cb(3).Text = NulosC(.Fields("docper")) '--IDENTIDAD PERSONA::DNI, PROVEEDOR::RUC
        Me.lbl_cb(3).Caption = NulosC(.Fields("benef")) '--NOMBRE DE LA PERSONA,NOMBRE DEL PROVEEDOR
        Me.lbl_cb_cod(3).Caption = NulosN(.Fields("idper")) '--CODIGO
        If IsDate(.Fields("fchpag")) = True Then
            TxtFecha(1).Valor = CDate(.Fields("fchpag")) '--FECHA DE PAGO
        End If
        If IsDate(.Fields("fchren")) = True Then
            TxtFecha(2).Valor = CDate(.Fields("fchren")) '--FECHA DE PAGO
        End If
        '--DEL TIPO DE DOCUMENTO
        If NulosN(.Fields("tipdoc")) <> 0 Then
            Me.txt_cb(4).Text = NulosN(.Fields("tipdoc"))
            txt_cb_Validate 4, False
        End If
        
        txt(1).Text = NulosC(.Fields("numdoc"))
        txt(2).Text = Format(NulosN(.Fields("imp")), FORMAT_MONTO)
        txt(3).Text = NulosC(.Fields("obs"))
        '----
        LblEstado(1).Caption = NulosN(.Fields("idest"))
        If NulosN(.Fields("idest")) <> 0 Then
            PONER_COLOR_ESTADO xCon, LblEstado(0), CInt(.Fields("idest"))
            fra_estado.Visible = True
            habilitar cmd_estado, True
        
        Else
            LblEstado(0).Caption = ""
            LblEstado(0).ForeColor = vbBlack
        End If
        '---DEL PROGRAMADOR
        lbl_prog(0).Caption = NulosC(.Fields("idprog"))
        lbl_prog(1).Caption = NulosC(.Fields("prog"))
    End With
    
    
    
    If cmd_estado(0).Enabled = True Then
        If fVerificarProgAut(False, lbl_aut(0), lbl_aut(1)) = False Then
            fra_estado.Visible = False
        End If
    End If
    If CmdModificar(0).Caption = "&Grabar" Then
        CmdModificar(0).Caption = "&Modificar"
        CmdModificar(0).Enabled = True
    End If
    CmdModificar(1).Enabled = False
    QueHace = QueHaceTmp
    Exit Sub
error:
    QueHace = QueHaceTmp
    SHOW_ERROR
End Sub

Private Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked TxtFecha, Not band
    habilitar_Locked txt, Not band
    habilitar_Locked txt_cb, Not band
    habilitar Me.opt_mov, band
    habilitar Me.opt_per, band
End Sub

Private Sub Blanquea()

    LblTipoCambio.Caption = ""
    LimpiaText Me.TxtFecha
    LimpiaText txt
    LimpiaText lbl_cb
    LimpiaText lbl_cb
    LimpiaText txt_cb
    
    lblCtaDestino.Caption = ""
    lblCtaOrigen.Caption = ""

End Sub

Private Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    '---
    TxtFecha(0).Valor = Date
    TxtFecha(1).Valor = Date
    TxtFecha(2).Valor = Date
    '---
    Blanquea
    Habilitar_Obj True
    Label1.Caption = "Programando Cuentas por Rendir"
    '-----
    PONER_COLOR_ESTADO xCon, LblEstado(0), 1 'PENDIENTE
    LblEstado(1).Caption = "1"
    '--
    opt_mov(0).Value = True
    opt_mov_Click 0
    '--CARGAR PROGRAMADOR
    fVerificarProgAut True, lbl_prog(0), lbl_prog(1)   '--PROGRAMADOR
    lbl_aut(0).Caption = ""
    lbl_aut(1).Caption = ""
    '-----
    fra_estado.Visible = False
    Ocultar CmdModificar, False
    '------
    
    TxtFecha(0).SetFocus
End Sub

Function Grabar(fConDiario As Boolean) As Boolean

    If fConDiario = False Then
        If fValidarDatos(False) = False Then Exit Function
        If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo Salir
    End If
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim xCod&
    Dim xCol&, xFil&
    Dim xNumAsiento As String
    
    On Error GoTo LaCague

    xCon.BeginTrans
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM con_ctasrendir ", xCon
        xCod = HallaCodigoTabla("con_ctasrendir", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xCod
        txt(0).Text = xCod
        RstCab("saldo") = NulosN(Trim(txt(2).Text))           '--SALDO
    Else
        xCod = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM con_ctasrendir WHERE id =" & xCod & "", xCon
        
        If fConDiario = True And NulosN(LblEstado(1).Caption) <> 1 Then
            xNumAsiento = DevuelveNumAsiento(38, NulosN(RstFrm("id")), xMes, xCon)
            If xNumAsiento = "" Then xNumAsiento = NuevoNumAsiento(38, xMes, xCon)
            'ELIMINAMOS EL ASIENTO REGISTRADO EN EL DIARIO
            xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & xMes & " and idlib = 38 AND idmov = " & xCod & " ;"
            RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
        End If
    End If
    
    RstCab("ano") = AnoTra
    RstCab("idlib") = 38
    If fConDiario = False Then
        RstCab("idmes") = -1 '--nos indica que es registro puede ser visto desde cualquier periodo
    Else
        If NulosN(LblEstado(1).Caption) = 2 Then '--aprobado
            RstCab("idmes") = xMes
            RstCab("numreg") = Format(xMes, "00") + xNumAsiento
            If xMes <> 0 And xMes <> 13 Then
                RstCab("fchreg") = CDate("01/" + Format(xMes, "00") + "/" + AnoTra)
            End If
            RstCab("idaut") = NulosN(lbl_aut(0).Caption)
            RstCab("idest") = NulosN(LblEstado(1).Caption)
        Else '--rechazado or pendiente
            RstCab("idmes") = -1
            RstCab("numreg") = Null
            RstCab("fchreg") = Null
            RstCab("idaut") = 0
        End If
        
    End If
    RstCab("fchemi") = CDate(TxtFecha(0).Valor)
    RstCab("fchpag") = CDate(TxtFecha(1).Valor)
    RstCab("fchren") = CDate(TxtFecha(2).Valor)
    RstCab("tipmov") = IIf(opt_mov(0).Value = True, "1", "2") '--CAJA O BANCO
    RstCab("idori") = NulosN(lbl_cb_cod(0).Caption) '--ORIGEN
    RstCab("iddes") = NulosN(lbl_cb_cod(2).Caption) '--DESTINO
    
    RstCab("tipper") = IIf(opt_per(0).Value = True, "1", "2") '--TIPO DE PERSONA
    RstCab("idper") = NulosN(lbl_cb_cod(3).Caption)      '--IDPERSONA
    RstCab("tipdoc") = NulosN(lbl_cb_cod(4).Caption)    '--TIPO DOC
    RstCab("numdoc") = Trim(txt(1).Text)
    RstCab("idmon") = NulosN(lbl_cb_cod(1).Caption)     '--MONEDA
    RstCab("imp") = NulosN(Trim(txt(2).Text))           '--IMPORTE
    
    RstCab("idprog") = NulosN(lbl_prog(0).Caption)      '--PROGRAMADOR
    RstCab("idest") = NulosN(LblEstado(1).Caption)      '--ESTADO
    RstCab("obs") = Trim(txt(3).Text)

    RstCab.Update
    If fConDiario = True And NulosN(LblEstado(1).Caption) = 2 Then
        '---del debe
        pGenerarAsiento RstDia, AnoTra, xMes, 38, xCod, 0, 0, xNumAsiento, NulosN(LblTipoCambio.Caption), TxtFecha(0).Valor, NulosN(lblCtaDestino.Caption), NulosN(lbl_cb_cod(1).Caption), NulosN(txt(2).Text), True
        '--del haber
        pGenerarAsiento RstDia, AnoTra, xMes, 38, xCod, 0, 0, xNumAsiento, NulosN(LblTipoCambio.Caption), TxtFecha(0).Valor, NulosN(lblCtaOrigen.Caption), NulosN(lbl_cb_cod(1).Caption), NulosN(txt(2).Text), False
    End If
    xCon.CommitTrans
    MsgBox "La Cuenta por Rendir se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + IIf(xNumAsiento = "", "", vbCr + "Num.Reg. " + Format(xMes, "00") & xNumAsiento), vbInformation, xTitulo
    Grabar = True
Salir:
    Set RstCab = Nothing
    Set RstDia = Nothing
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDia = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description), vbCritical, xTitulo
    Grabar = False
    Exit Function
End Function

Private Function fValidarDatos(fConDiario As Boolean) As Boolean

    If IsDate(TxtFecha(0).Valor) = False Then
        MsgBox "No ha especificado la fecha de emisión", vbInformation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If
    
    If IsDate(TxtFecha(1).Valor) = False Then
        MsgBox "No ha especificado la fecha de pago", vbInformation, xTitulo
        TxtFecha(1).SetFocus
        Exit Function
    End If
    If IsDate(TxtFecha(2).Valor) = False Then
        MsgBox "No ha especificado la fecha de rendición de cuentas", vbInformation, xTitulo
        TxtFecha(2).SetFocus
        Exit Function
    End If
    If CDate(TxtFecha(0).Valor) > CDate(TxtFecha(1).Valor) Then
        MsgBox "La fecha de Pago es inferior a la fecha de emisión" + vbCr + "Modifique la fecha de Pago", vbInformation, xTitulo
        TxtFecha(1).SetFocus
        Exit Function
    End If
    If CDate(TxtFecha(1).Valor) > CDate(TxtFecha(2).Valor) Then
        MsgBox "La fecha de Rendición de cuentas es inferior a la fecha de Pago" + vbCr + "Modifique la fecha Rendición de Cuentas", vbInformation, xTitulo
        TxtFecha(2).SetFocus
        Exit Function
    End If
    
    If lbl_prog(1).Caption = "0" And QueHace = 1 Then
        MsgBox "Ust. No no puede programar Cuentas por Rendir", vbInformation, xTitulo
        Exit Function
    End If
    
    Dim band As Integer
    band = Validar(txt_cb)
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl_cb_capt(band).Caption, vbInformation, xTitulo
       txt_cb(band).SetFocus
       Exit Function
    End If
    band = Validar(txt)
    If band > 0 Then
       MsgBox "Llene el Campo de " & lbl(band).Caption, vbInformation, xTitulo
       txt(band).SetFocus
       Exit Function
    End If

    
    If IsNumeric(Trim(txt(2).Text)) = False Then
        MsgBox "El importe no es correcto", vbExclamation, xTitulo
        txt(2).Text = ""
        Exit Function
    End If
    If fConDiario = True Then
        If NulosN(LblEstado(1).Caption) = 2 Then '--solo aprobado
            If NulosN(lblCtaOrigen.Caption) = 0 Then
                MsgBox "Falta especificar el N° de Cuenta Contable del Origen:" + vbCr + _
                    "Moneda: " + lbl_cb(1).Caption + vbCr + _
                    "Origen: " + IIf(opt_mov(0).Value = True, "Caja", "Banco") + vbCr + _
                    "Descripción: " + lbl_cb(0).Caption, vbExclamation, xTitulo
                Exit Function
            End If
            If NulosN(lblCtaDestino.Caption) = 0 Then
                MsgBox "Falta especificar el N° de Cuenta Contable del Destino: " + vbCr + _
                    "Moneda: " + lbl_cb(1).Caption + vbCr + _
                    "Descripción: " + lbl_cb(0).Caption, vbExclamation, xTitulo
                Exit Function
            End If
        End If
    End If
            
    fValidarDatos = True
End Function
 

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then Imprimir True

    If ButtonMenu.Index = 2 Then Imprimir

End Sub

Private Sub txt_cb_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cb_cod(Index).Caption = ""
    End If
    If Index = 0 Then
        lblCtaOrigen.Caption = ""
    ElseIf Index = 2 Then
        lblCtaDestino.Caption = ""
    ElseIf Index = 1 Then
        Me.txt_cb(0).Text = ""
        Me.txt_cb(2).Text = ""
    End If
    
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
    If KeyCode = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
'--restringir cuando deseen modificar independientemente
    '--habilitar cuando sea tipo doc
    If CmdModificar(0).Visible = True Then
        If Index <> 4 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    '----------

    Select Case Index
        Case 1, 2, 3, 4: If validar_numero(KeyAscii) = False Then KeyAscii = 0
        
    End Select
    
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    
    If txt_cb(Index).Text = "" Then Exit Sub
    If QueHace = 3 Then Exit Sub
    '----------
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
    
        Case 0 '--ORIGEN
            If fValidarMoneda() = False Then Exit Sub
            If opt_mov(0).Value = True Then
                nSQL = "SELECT con_destino.id, con_destino.descripcion,con_destino.id as cod ,con_destino.idcuen as idcta " _
                    + vbCr + " From con_destino " _
                    + vbCr + " WHERE (((con_destino.id)=" + CStr(Trim(txt_cb(Index).Text)) + ")) AND con_destino.idmon=" + CStr(lbl_cb_cod(1).Caption)
            Else
                nSQL = "SELECT [con_bancocuenta].[numcue], [mae_bancos].[descripcion]  AS nombre,con_bancocuenta.id as cod , con_bancocuenta.idcuen as idcta " _
                    + vbCr + " FROM mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban " _
                    + vbCr + " WHERE ((([con_bancocuenta].[numcue])='" + CStr(Trim(txt_cb(Index).Text)) + "')) AND con_bancocuenta.idmon=" + CStr(lbl_cb_cod(1).Caption)
            End If
        Case 1 '--MONEDA
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion,mae_moneda.id as cod " _
                + vbCr + " From mae_moneda " _
                + vbCr + " WHERE (((mae_moneda.id)=" + CStr(Trim(txt_cb(Index).Text)) + "));"
        
        Case 2 '--DESTINO
            If fValidarMoneda() = False Then Exit Sub
             
             nSQL = "SELECT con_destino.id, con_destino.descripcion AS nombre, con_destino.id AS cod ,con_destino.idcuen as idcta " _
                    + vbCr + " FROM con_destino " _
                    + vbCr + "WHERE (((con_destino.id)=" + CStr(Trim(txt_cb(Index).Text)) + ")) AND (((con_destino.idmon)=" + lbl_cb_cod(1).Caption + ") AND ((con_destino.tipmov)=2)) and con_destino.rendir = -1;"
    
        
        Case 3 '--BENEFICIARIO
            If Me.opt_per(0).Value = True Then '--PERSONA
                '--no muestra al usuario logueado
                nSQL = "SELECT pla_empleados.numdoc , [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nombre,con_emptes.id " _
                    + vbCr + " FROM pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp " _
                    + vbCr + " WHERE (((pla_empleados.numdoc)='" + CStr(Trim(txt_cb(Index).Text)) + "')) " _
                    + vbCr + " AND con_emptes.id NOT IN (SELECT emptes.id " _
                                + vbCr + " FROM (pla_empleados AS emp INNER JOIN mae_usuarios AS usuario ON emp.id = usuario.id) INNER JOIN con_emptes AS emptes ON emp.id = emptes.idemp " _
                                + vbCr + " WHERE (((usuario.id)=" + CStr(xIdUsuario) + "));)"
            Else '--PROVEEDOR
                nSQL = "SELECT  mae_prov.numruc, mae_prov.nombre,mae_prov.id " _
                    + vbCr + " From mae_prov " _
                    + vbCr + " WHERE (((mae_prov.numruc)='" + CStr(Trim(txt_cb(Index).Text)) + "'));"
            
            End If
            
        Case 4 '--TIPO DE DOCUMENTO
            nSQL = "SELECT mae_doccajaban.id, mae_doccajaban.descripcion as nombre, mae_doccajaban.id AS cod, mae_doccajaban.abrev " _
                + vbCr + " From mae_doccajaban " _
                + vbCr + " WHERE mae_doccajaban.tipo=" + IIf(opt_mov(0).Value = True, "1", "2") + " AND mae_doccajaban.id=" + CStr(Trim(txt_cb(Index).Text))
    
        
    End Select
    If xCon.State = 0 Then Exit Sub
    RST_Busq xRs, nSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount > 0 Then
        txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
        If Index = 0 Then '--origen
            lblCtaOrigen.Caption = NulosN(xRs.Fields("idcta"))
        ElseIf Index = 2 Then '--destino
            lblCtaDestino.Caption = NulosN(xRs.Fields("idcta"))
        End If
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cb_cod(Index).Caption = ""
        If Index = 0 Then '--origen
            lblCtaOrigen.Caption = ""
        ElseIf Index = 2 Then '--destino
            lblCtaDestino.Caption = ""
        End If
    End If
    
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    Select Case Index
        Case 2:
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub


Private Function fVerificarProgAut(BUSCAPROGRAMADOR As Boolean, OBJ_ID As Label, OBJ_NOMBRE As Label) As Boolean
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQL_PROG As String
    If BUSCAPROGRAMADOR = True Then
        nSQL_PROG = " AND con_emptes.prog=-1; "
    Else
        nSQL_PROG = " AND con_emptes.aut=-1; "
    End If
    
    nSQL = "SELECT  con_emptes.id, [pla_empleados].[nom] & ' ' & [pla_empleados].[ape] AS nombre " _
        + vbCr + " FROM (pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp) INNER JOIN mae_usuarios ON pla_empleados.id = mae_usuarios.idemp " _
        + vbCr + " WHERE mae_usuarios.id= " + CStr(xIdUsuario) + nSQL_PROG
    
    
    RST_Busq xRs, nSQL, xCon
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Then
        OBJ_ID.Caption = "--"
        OBJ_NOMBRE.Caption = "--"
    Else
        OBJ_ID.Caption = xRs.Fields(0) & ""
        OBJ_NOMBRE.Caption = xRs.Fields(1) & ""
        fVerificarProgAut = True
    End If
Salir:
    Set xRs = Nothing
End Function

Private Function fValidarMoneda() As Boolean
    '--FUNCTION QUE VALIDAR SI SELECCIONO LA MONEDA
    If NulosN(lbl_cb_cod(1).Caption) = 0 Then
        MsgBox "Seleccione primero la Moneda", vbInformation, xTitulo
        cb_Click 1
        Exit Function
    End If
    fValidarMoneda = True
End Function

'------DEL CAMBIO DE PERIODO
Private Sub CambiarMes()
    
    xMes = SeleccionaMes(xCon)
    pCargarGrid
End Sub

Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(8, 4) As String
    
    xCampos(0, 0) = "Tipo Mov.":        xCampos(0, 1) = "tipmov":   xCampos(0, 2) = "1000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Emi":         xCampos(1, 1) = "fchemi":   xCampos(1, 2) = "850":     xCampos(1, 3) = "F"
    xCampos(2, 0) = "T.Doc":            xCampos(2, 1) = "abrev":   xCampos(2, 2) = "500":     xCampos(2, 3) = "C"
    xCampos(3, 0) = "M":                xCampos(3, 1) = "simbolo":  xCampos(3, 2) = "450":     xCampos(3, 3) = "C"
    xCampos(4, 0) = "Importe":          xCampos(4, 1) = "imp":      xCampos(4, 2) = "700":     xCampos(4, 3) = "N"
    xCampos(5, 0) = "Emitido Por:":     xCampos(5, 1) = "prog":     xCampos(5, 2) = "2000":    xCampos(5, 3) = "C"
    xCampos(6, 0) = "T. Persona":       xCampos(6, 1) = "tipper":   xCampos(6, 2) = "1000":    xCampos(6, 3) = "C"
    xCampos(7, 0) = "Entregado A":      xCampos(7, 1) = "benef":    xCampos(7, 2) = "1500":    xCampos(7, 3) = "C"
        
    nSQL = "SELECT con_ctasrendir.id, format(con_ctasrendir.fchemi,'dd/mm/yy') as fchemi, con_ctasrendir.fchpag, con_ctasrendir.fchren, con_ctasrendir.numdoc, " _
        + vbCr + " IIf(con_ctasrendir.tipmov=1,'Caja','Banco') AS tipmov, con_ctasrendir.idmon, mae_moneda.descripcion AS monnom, mae_moneda.simbolo, " _
        + vbCr + " con_ctasrendir.idori, " _
        + vbCr + " IIf(con_ctasrendir.tipmov=1,(con_ctasrendir.idori),(SELECT  [con_bancocuenta].[numcue]  AS origen FROM mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban WHERE (((con_bancocuenta.id)=con_ctasrendir.idori)) )) AS origen_num, " _
        + vbCr + " IIf(con_ctasrendir.tipmov=1,(SELECT destino.descripcion FROM con_destino AS destino WHERE (((destino.id)=con_ctasrendir.idori)) ),(SELECT [mae_bancos].[descripcion] AS origen FROM mae_bancos INNER JOIN con_bancocuenta ON mae_bancos.id = con_bancocuenta.idban WHERE (((con_bancocuenta.id)=con_ctasrendir.idori)) )) AS origen, " _
        + vbCr + " con_ctasrendir.iddes, con_destino.descripcion AS destino, con_ctasrendir.tipdoc, mae_doccajaban.descripcion AS tipdocnom, mae_doccajaban.abrev, IIf(con_ctasrendir.tipper=1,'Persona','Proveedor') AS tipper, con_ctasrendir.idper, " _
        + vbCr + " IIf(con_ctasrendir.tipper=1,(SELECT [pla_empleados].[nom] & ' ' & [pla_empleados].[ape] AS nombre FROM pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp WHERE (((con_emptes.id)=con_ctasrendir.idper))),(SELECT mae_prov.nombre FROM mae_prov WHERE (((mae_prov.id)=con_ctasrendir.idper)) )) AS benef, " _
        + vbCr + " IIf(con_ctasrendir.tipper=1,(SELECT pla_empleados.numdoc  FROM pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp WHERE (((con_emptes.id)=con_ctasrendir.idper))),(SELECT mae_prov.numruc FROM mae_prov WHERE (((mae_prov.id)=con_ctasrendir.idper)) )) AS docper, " _
        + vbCr + " con_ctasrendir.[imp],con_ctasrendir.[saldo], con_ctasrendir.obs, con_ctasrendir.idprog, pla_empleados.nom & ' ' & pla_empleados.ape AS prog, con_ctasrendir.idaut, pla_empleados_1.ape & ' ' & pla_empleados_1.nom AS aut, con_ctasrendir.idest, mae_estados.descripcion AS estdesc, " _
        + vbCr + " IIF(con_ctasrendir.numreg IS NULL OR con_ctasrendir.numreg='','', FORMAT (con_ctasrendir.idmes,'00') & IIF(mae_libros.codsun IS NULL OR mae_libros.codsun ='','FF',mae_libros.codsun) & MID(con_ctasrendir.numreg,3)) AS numreg " _
        + vbCr + " FROM (pla_empleados RIGHT JOIN (mae_moneda RIGHT JOIN (mae_doccajaban RIGHT JOIN (con_emptes RIGHT JOIN (mae_estados RIGHT JOIN (con_destino RIGHT JOIN (con_ctasrendir LEFT JOIN (con_emptes AS con_emptes_1 LEFT JOIN pla_empleados AS pla_empleados_1 ON con_emptes_1.idemp = pla_empleados_1.id) ON con_ctasrendir.idaut = con_emptes_1.id) ON con_destino.id = con_ctasrendir.iddes) ON mae_estados.id = con_ctasrendir.idest) ON con_emptes.id = con_ctasrendir.idprog) ON mae_doccajaban.id = con_ctasrendir.tipdoc) ON mae_moneda.id = con_ctasrendir.idmon) ON pla_empleados.id = con_emptes.idemp) LEFT JOIN mae_libros ON con_ctasrendir.idlib = mae_libros.id " _
        + vbCr + " WHERE con_ctasrendir.ano = " & AnoTra & " And con_ctasrendir.idmes IN (-1," & xMes & ")" _
        + vbCr + " ORDER BY con_ctasrendir.fchemi ASC "
    
    
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Cuentas x Rendir", "fchemi", "prog", Principio
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub




Private Sub Filtrar()
    
    Dim xCampos(8, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    xCampos(0, 0) = "Num.Reg.":         xCampos(0, 1) = "numreg":   xCampos(0, 2) = "C":     xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Tipo Mov.":        xCampos(1, 1) = "tipmov":   xCampos(1, 2) = "C":     xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Fch. Emi":         xCampos(2, 1) = "fchemi":   xCampos(2, 2) = "F":     xCampos(2, 3) = "850"
    xCampos(3, 0) = "N°.Documento":     xCampos(3, 1) = "numdoc":   xCampos(3, 2) = "C":     xCampos(3, 3) = "850"
    xCampos(4, 0) = "M":                xCampos(4, 1) = "simbolo":  xCampos(4, 2) = "C":     xCampos(4, 3) = "450"
    xCampos(5, 0) = "Importe":          xCampos(5, 1) = "imp":      xCampos(5, 2) = "N":     xCampos(5, 3) = "800"
    xCampos(6, 0) = "Emitido Por:":     xCampos(6, 1) = "prog":     xCampos(6, 2) = "C":     xCampos(6, 3) = "2500"
    xCampos(7, 0) = "T. Persona":       xCampos(7, 1) = "tipper":   xCampos(7, 2) = "C":     xCampos(7, 3) = "1000"
    xCampos(8, 0) = "Entregado A":      xCampos(8, 1) = "benef":    xCampos(8, 2) = "C":     xCampos(8, 3) = "2500"
        
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3

    TabOne1.CurrTab = 0
End Sub


Private Sub Imprimir(Optional IMP_LISTADO As Boolean = False)

    On Error GoTo error

    Me.MousePointer = vbHourglass
    If IMP_LISTADO = False Then
        If Me.TabOne1.CurrTab = 0 Then
        
        Else
''            MsgBox "Primero muestre el detalle del Registro" + vbCr + _
''                   "Luego inténtelo otra vez", vbExclamation, xTitulo
        End If
    Else
    
        TDB_IMPRIMIR Dg3, "IMPRESIÓN DE CUENTAS POR RENDIR", "LISTADO DE CUENTAS POR RENDIR-  Periodo: " + MonthName(xMes, False)
   
    End If

    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "IMPRIMIR"

End Sub


Private Sub CmdModificar_Click(Index As Integer)
    Select Case Index
        Case 0 '--modificar
            If CmdModificar(0).Caption = "&Modificar" Then
                CmdModificar(0).Caption = "&Grabar"
                pBloqueaModificar False
                QueHace = 2
                TxtFecha(1).SetFocus
            Else
                '---
                If fValidarDatos(True) = False Then Exit Sub
                If MsgBox("Seguro desea modificar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
                If Grabar(True) = True Then
                    Dim mCod As Variant
                    mCod = RstFrm.Fields("id")
                    RstFrm.Requery
                    CmdModificar_Click 1
                    If RstFrm.RecordCount <> 0 Then
                        RstFrm.MoveFirst
                        RstFrm.Find "id=" & mCod
                    End If
                    TabOne1.CurrTab = 0
                End If
                '---
            End If
        Case 1 '--cancelar
            CmdModificar(0).Caption = "&Modificar"
            pBloqueaModificar True
            QueHace = 3
    End Select
End Sub

Private Sub txtfecha_Validate(Index As Integer, Cancel As Boolean)
    If Index <> 0 Then Exit Sub
    If IsDate(TxtFecha(0).Valor) = True Then
        LblTipoCambio.Caption = HallaTipoCambio(TxtFecha(0).Valor, 2, Venta, xCon)
    Else
        LblTipoCambio.Caption = ""
    End If
End Sub


Private Sub pGenerarAsiento(RstDiario As ADODB.Recordset, nAnoTrabajo, mMesActivo, IDLibro, IDMov, mIdDocPro, mCorr, nAsiento, mTipoCambio, FchDoc, IDcuenta, IDMoneda, mImporte, Optional EsDEBE As Boolean)
    '--mCorr por le general es igual a 0
    RstDiario.AddNew
    RstDiario("año") = nAnoTrabajo
    RstDiario("idmes") = mMesActivo  'CODIGO DEL MES
    RstDiario("idlib") = IDLibro     'CODIGO DEL LIBRO
    RstDiario("idmov") = IDMov       'CODIGO DEL MOVIMIENTO
    RstDiario("iddocpro") = mIdDocPro
    RstDiario("correlativo") = mCorr
    RstDiario("numasi") = nAsiento
    RstDiario("tc") = mTipoCambio
    If mMesActivo = 0 Then
        RstDiario("fchasi") = CDate("01/01/" + nAnoTrabajo)
    ElseIf mMesActivo = 13 Then
        RstDiario("fchasi") = CDate("31/12/" + nAnoTrabajo)
    Else
        RstDiario("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + nAnoTrabajo)
    End If
    RstDiario("fchdoc") = FchDoc
    RstDiario("idcue") = IDcuenta
    If EsDEBE = False Then
        If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("imphabsol") = mImporte
            RstDiario("imphabdol") = 0
        Else
            RstDiario("imphabsol") = mImporte * mTipoCambio
            RstDiario("imphabdol") = mImporte
        End If
    Else
        If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("impdebsol") = mImporte
            RstDiario("impdebdol") = 0
        Else
            RstDiario("impdebsol") = mImporte * mTipoCambio
            RstDiario("impdebdol") = mImporte
        End If
    End If

    RstDiario.Update
End Sub

Function fGrabarMofificar() As Boolean
    If IsDate(TxtFecha(1).Valor) = False Then
        MsgBox "Falta ingresar la Fecha de Pago ", vbExclamation, xTitulo
        TxtFecha(1).SetFocus
        Exit Function
    End If
    If IsDate(TxtFecha(2).Valor) = False Then
        MsgBox "Falta ingresar la Fecha a Rendir", vbExclamation, xTitulo
        TxtFecha(2).SetFocus
        Exit Function
    End If
    If NulosN(txt_cb(4).Text) = 0 Then
        MsgBox "Falta ingresar el Tipo de Documento", vbExclamation, xTitulo
        txt_cb(4).SetFocus
        Exit Function
    End If
    If Trim(txt(2).Text) = "" Then
        MsgBox "Falta ingresar el Importe", vbExclamation, xTitulo
        txt(2).SetFocus
        Exit Function
    End If
    
    If MsgBox("Seguro desea modificar el registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo Salir
    
    
    Dim RstCab As New ADODB.Recordset
    Dim xCod As Integer
        
    On Error GoTo LaCague

    xCon.BeginTrans
    
    
    RST_Busq RstCab, "SELECT * FROM con_ctasrendir WHERE id =" & RstFrm("id") & "", xCon
   
    RstCab("fchpag") = CDate(TxtFecha(1).Valor)
    RstCab("fchren") = CDate(TxtFecha(2).Valor)
    RstCab("tipdoc") = lbl_cb_cod(4).Caption '--TIPO DOC
    RstCab("numdoc") = Trim(txt(1).Text)
    RstCab("imp") = NulosN(Trim(txt(2).Text)) '--IMPORTE
    RstCab("obs") = Trim(txt(3).Text)

    RstCab.Update
    
    MsgBox "La Cuenta por Rendir se modificó con éxito", vbInformation, xTitulo
    
    
    xCon.CommitTrans
    fGrabarMofificar = True
Salir:
    Set RstCab = Nothing
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    MsgBox "No se pudo modificar el registro por el siguiente motivo :" + Trim(Err.Description), vbCritical, xTitulo
    Exit Function
End Function


Private Sub pBloqueaModificar(band As Boolean)

    TxtFecha(1).Locked = band
    TxtFecha(2).Locked = band
    txt_cb(4).Locked = band
    txt(1).Locked = band '--numdoc
    If txt(1).Enabled = False Then txt(1).Enabled = True
    txt(2).Locked = band '--impote
    txt(3).Locked = band '--obs
    habilitar cmd_estado, band
    habilitar opt_mov, Not band
    habilitar opt_per, Not band
    
    CmdModificar(1).Enabled = Not band
    
End Sub

