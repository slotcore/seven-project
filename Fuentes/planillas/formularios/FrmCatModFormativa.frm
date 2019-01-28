VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Begin VB.Form FrmCatModFormativa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Nómina del Personal - Modalidad Formativa"
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
      TabIndex        =   23
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
      TabIndex        =   21
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
         Caption         =   "Categoría: Modalidad Formativa"
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
         TabIndex        =   22
         Top             =   45
         Width           =   3360
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   405
      Index           =   12
      Left            =   135
      TabIndex        =   19
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
         TabIndex        =   20
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
      TabIndex        =   9
      Top             =   4710
      Width           =   1755
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Grabar"
      Height          =   420
      Index           =   0
      Left            =   5625
      TabIndex        =   8
      Top             =   4710
      Width           =   1755
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   4155
      Left            =   90
      TabIndex        =   10
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
      Caption         =   "    Datos Principales    "
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
         TabIndex        =   13
         Top             =   45
         Width           =   9210
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   3735
         Index           =   1
         Left            =   9945
         TabIndex        =   12
         Top             =   45
         Width           =   9210
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3735
         Index           =   0
         Left            =   45
         TabIndex        =   11
         Top             =   45
         Width           =   9210
         Begin VB.Frame fra 
            Caption         =   "[ ¿Es madre con responsabilidad familiar? ]"
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
            Index           =   18
            Left            =   4665
            TabIndex        =   49
            Top             =   750
            Width           =   4440
            Begin VB.OptionButton opt_Madre 
               Caption         =   "Si"
               Height          =   330
               Index           =   1
               Left            =   2580
               TabIndex        =   50
               Top             =   240
               Width           =   585
            End
            Begin VB.OptionButton opt_Madre 
               Caption         =   "No"
               Height          =   330
               Index           =   0
               Left            =   1290
               TabIndex        =   2
               Top             =   240
               Value           =   -1  'True
               Width           =   585
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ ¿Sujeto a horario nocturno? ]"
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
            Index           =   19
            Left            =   6120
            TabIndex        =   47
            Top             =   1500
            Width           =   2985
            Begin VB.OptionButton opt_HoraNocturno 
               Caption         =   "No"
               Height          =   330
               Index           =   0
               Left            =   720
               TabIndex        =   4
               Top             =   225
               Value           =   -1  'True
               Width           =   585
            End
            Begin VB.OptionButton opt_HoraNocturno 
               Caption         =   "Si"
               Height          =   330
               Index           =   1
               Left            =   1770
               TabIndex        =   48
               Top             =   225
               Width           =   585
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Centro de Formación Profesional ]"
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
            Left            =   30
            TabIndex        =   42
            Top             =   2985
            Width           =   5925
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   4
               Left            =   540
               Picture         =   "FrmCatModFormativa.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Seleccione el Centro de Formación Profesional"
               Top             =   285
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   4
               Left            =   135
               MaxLength       =   20
               TabIndex        =   6
               Tag             =   "null"
               Text            =   "txt_cb(4)"
               Top             =   255
               Width           =   645
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
               TabIndex        =   46
               Top             =   255
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Centro de Formación Profesional"
               Height          =   195
               Index           =   4
               Left            =   3495
               TabIndex        =   45
               Top             =   75
               Visible         =   0   'False
               Width           =   2295
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
               TabIndex        =   44
               Top             =   255
               Width           =   3915
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ Seguro Médico ]"
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
            Left            =   30
            TabIndex        =   37
            Top             =   720
            Width           =   4500
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   1
               Left            =   540
               Picture         =   "FrmCatModFormativa.frx":0132
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "Seleccione el Seguro Médico"
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
               Caption         =   "Seguro Médico"
               Height          =   195
               Index           =   1
               Left            =   2490
               TabIndex        =   40
               Top             =   90
               Visible         =   0   'False
               Width           =   1080
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
               Left            =   2655
               TabIndex        =   39
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
               TabIndex        =   41
               Top             =   285
               Width           =   3630
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
            Height          =   660
            Index           =   7
            Left            =   6120
            TabIndex        =   35
            Top             =   3015
            Width           =   2985
            Begin VB.OptionButton opt_discapacidad 
               Caption         =   "No"
               Height          =   210
               Index           =   0
               Left            =   720
               TabIndex        =   7
               Top             =   315
               Value           =   -1  'True
               Width           =   675
            End
            Begin VB.OptionButton opt_discapacidad 
               Caption         =   "Si"
               Height          =   210
               Index           =   1
               Left            =   1770
               TabIndex        =   36
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
            Left            =   30
            TabIndex        =   30
            Top             =   2205
            Width           =   9075
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   3
               Left            =   540
               Picture         =   "FrmCatModFormativa.frx":0264
               Style           =   1  'Graphical
               TabIndex        =   31
               ToolTipText     =   "Seleccione la Ocupación"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   3
               Left            =   135
               MaxLength       =   20
               TabIndex        =   5
               Tag             =   "null"
               Text            =   "txt_cb(3)"
               Top             =   285
               Width           =   645
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
               TabIndex        =   33
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
               TabIndex        =   32
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
               Left            =   780
               TabIndex        =   34
               Top             =   285
               Width           =   8040
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
            Left            =   30
            TabIndex        =   25
            Top             =   1470
            Width           =   5925
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   2
               Left            =   540
               Picture         =   "FrmCatModFormativa.frx":0396
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "Seleccione el Nivel Educativo"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   2
               Left            =   135
               MaxLength       =   20
               TabIndex        =   3
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
               TabIndex        =   28
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
               TabIndex        =   27
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
               TabIndex        =   29
               Top             =   285
               Width           =   4980
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
            Left            =   30
            TabIndex        =   14
            Top             =   30
            Width           =   5925
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   0
               Left            =   540
               Picture         =   "FrmCatModFormativa.frx":04C8
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Seleccione el Tipo de Trabajador"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               BackColor       =   &H8000000F&
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
               TabIndex        =   18
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
               TabIndex        =   17
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
               TabIndex        =   16
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
      TabIndex        =   24
      Top             =   5280
      Width           =   495
   End
End
Attribute VB_Name = "FrmCatModFormativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QueHace As Integer
Dim mCorrelativo As Long
Dim mIdEmpleado As Long
Dim Agregando As Boolean
Dim SeEjecuto As Boolean

Public Sub pRecibeLink(QueHace1 As Integer)
    QueHace = QueHace1
    mCorrelativo = mCorr
    With FrmNomina
        mIdEmpleado = .txt(0).Text
        lbl_persona.Caption = .lbl_persona(0).Caption
    End With
    '------
    Agregando = True
    LimpiaText txt_cb
    TabOne1.CurrTab = 0
    pPonerDatos
    Agregando = False
End Sub


'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
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
                + vbCr + " WHERE mae_tipotrabajadorcat.idcat = 3 " _
                + vbCr + " ORDER BY mae_tipotrabajador.codsun;"
        
        Case 1 '--SEGURO MEDICO
            nTitulo = "Buscando Seguro Médico"
            nSQL = "SELECT mae_seguromedico.id, mae_seguromedico.descripcion AS nombre, mae_seguromedico.id AS cod " _
                + vbCr + " FROM mae_seguromedico" _
                + vbCr + " ORDER BY mae_seguromedico.codsun; "
        
        Case 2 '--NIVEL EDUCATIVO
            nTitulo = "Buscando Nivel Educativo"
            nSQL = "SELECT mae_niveleducativo.id, mae_niveleducativo.descripcion AS nombre, mae_niveleducativo.id AS cod " _
                + vbCr + " From mae_niveleducativo " _
                + vbCr + " ORDER BY mae_niveleducativo.codsun;"
                 
        Case 3 '--OCUPACION
            nTitulo = "Buscando Ocupación"
            nSQL = "SELECT mae_ocupacion.id, mae_ocupacion.descripcion AS nombre, mae_ocupacion.id AS cod " _
                + vbCr + " From mae_ocupacion " _
                + vbCr + " ORDER BY mae_ocupacion.codsun; "

        Case 4 '--FORMACION PROFESIONAL
            nTitulo = "Buscando Centro Formación Profesional"
            nSQL = "SELECT mae_centroformacion.id, mae_centroformacion.descripcion AS nombre, mae_centroformacion.id AS cod " _
                + vbCr + " FROM mae_centroformacion " _
                + vbCr + " ORDER BY mae_centroformacion.codsun; "
        
    End Select
    
    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
            
    Dim xRs As New ADODB.Recordset
    If Index <> 3 And Index <> 2 Then
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
    Else
        '--SOLO OCUPACION
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", CualquierParte
    End If

    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    lbl_cb(Index).ToolTipText = xRs.Fields(1) & "" '--NOMBRE
   
   
    Select Case Index
        Case 0 '--TIPO DE TRABAJADOR
            txt_cb(1).SetFocus
        Case 1 '--SEGURO MEDICO
            If opt_Madre(0).Value = True Then opt_Madre(0).SetFocus
            If opt_Madre(1).Value = True Then opt_Madre(1).SetFocus
        Case 2 '--NIVEL EDUCATIVO
            If opt_HoraNocturno(0).Value = True Then opt_HoraNocturno(0).SetFocus
            If opt_HoraNocturno(1).Value = True Then opt_HoraNocturno(1).SetFocus
        Case 3 '--OCUPACION
            txt_cb(4).SetFocus
        Case 4 '--CENTRO FORMACION PROFESIONAL
                If opt_discapacidad(0).Value = True Then opt_discapacidad(0).SetFocus
                If opt_discapacidad(1).Value = True Then opt_discapacidad(1).SetFocus
    End Select
salir:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--GRABAR
            If Grabar() = False Then Exit Sub
            FrmNomina.pCargarDatosPeriodoLaboral
            Unload Me
        Case 1 '--CANCELAR
            Unload Me
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = True
    txt_cb(1).SetFocus
End Sub
Private Sub Form_Load()
    SeEjecuto = False
    CentrarFrm Me
End Sub

Private Sub opt_discapacidad_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmd(0).SetFocus
End Sub


Private Sub opt_HoraNocturno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_cb(3).SetFocus
End Sub

Private Sub opt_Madre_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txt_cb(2).SetFocus
End Sub

Private Sub TabOne1_Click()
    If TabOne1.CurrTab = 0 Then
        If Agregando = False Then txt_cb(0).SetFocus
    Else
        If Agregando = False Then txt_cb(7).SetFocus
    End If
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
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
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--TIPO DE TRABAJADOR
            nSQL = "SELECT mae_tipotrabajador.id, mae_tipotrabajador.descripcion AS nombre, mae_tipotrabajador.id AS cod, mae_tipotrabajador.codsun " _
                + vbCr + " FROM mae_tipotrabajador INNER JOIN mae_tipotrabajadorcat ON mae_tipotrabajador.id = mae_tipotrabajadorcat.id  " _
                + vbCr + " WHERE mae_tipotrabajador.id = " & NulosN(txt_cb(Index).Text) & " AND mae_tipotrabajadorcat.idcat = 3 ;"
               
        Case 1 '--SEGURO MEDICO
            nSQL = "SELECT mae_seguromedico.id, mae_seguromedico.descripcion AS nombre, mae_seguromedico.id AS cod " _
                + vbCr + " FROM mae_seguromedico" _
                + vbCr + " WHERE mae_seguromedico.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 2 '--NIVEL EDUCATIVO
            nSQL = "SELECT mae_niveleducativo.id, mae_niveleducativo.descripcion AS nombre, mae_niveleducativo.id AS cod " _
                + vbCr + " From mae_niveleducativo " _
                + vbCr + " WHERE mae_niveleducativo.id = " & NulosN(txt_cb(Index).Text) & ";"
                 
        Case 3 '--OCUPACION
            nSQL = "SELECT mae_ocupacion.id, mae_ocupacion.descripcion AS nombre, mae_ocupacion.id AS cod " _
                + vbCr + " From mae_ocupacion " _
                + vbCr + " WHERE mae_ocupacion.id = " & NulosN(txt_cb(Index).Text) & ";"

        Case 4 '--REGIMEN PENSIONARIO
             nSQL = "SELECT mae_centroformacion.id, mae_centroformacion.descripcion AS nombre, mae_centroformacion.id AS cod " _
                + vbCr + " FROM mae_centroformacion " _
                + vbCr + " WHERE mae_centroformacion.id = " & NulosN(txt_cb(Index).Text) & ";"
         
    End Select

    If xCon.State = 0 Then GoTo salir
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
        lbl_cb(Index).ToolTipText = RstTmp.Fields(1) & "" '--NOMBRE
    Else
        txt_cb(Index).Text = ""
    End If
    '--------------If Agregando = False Then
   Select Case Index
        Case 0 '--TIPO DE TRABAJADOR
            If Agregando = False Then txt_cb(1).SetFocus
        Case 1 '--SEGURO MEDICO
            If Agregando = False Then
                If opt_Madre(0).Value = True Then opt_Madre(0).SetFocus
                If opt_Madre(1).Value = True Then opt_Madre(1).SetFocus
            End If
        Case 2 '--NIVEL EDUCATIVO
            If Agregando = False Then
                If opt_HoraNocturno(0).Value = True Then opt_HoraNocturno(0).SetFocus
                If opt_HoraNocturno(1).Value = True Then opt_HoraNocturno(1).SetFocus
            End If
        Case 3 '--OCUPACION
            If Agregando = False Then txt_cb(4).SetFocus
        Case 4 '--CENTRO FORMACION PROFESIONAL
            If Agregando = False Then
                If opt_discapacidad(0).Value = True Then opt_discapacidad(0).SetFocus
                If opt_discapacidad(1).Value = True Then opt_discapacidad(1).SetFocus
            End If
    End Select
    '--------------
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
salir:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub



'****************************************************************************************

Private Sub pPonerDatos()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    nSQL = "SELECT pla_categoria4.* From pla_categoria4 WHERE (((pla_categoria4.idemp)=" & mIdEmpleado & "));"

    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.RecordCount = 0 Then
        QueHace = 1
        txt_cb(0).Text = 23
        txt_cb_Validate 0, False
        
        Exit Sub
    End If
    QueHace = 2
    Agregando = True

    '************************************************************TAB 0
    '--TIPO TRABAJADOR
    txt_cb(0).Text = 23
    txt_cb_Validate 0, False
    '--REGIMEN LABORAL
    If NulosN(RstTmp("idsegmed")) <> 0 Then
        txt_cb(1).Text = NulosN(RstTmp("idsegmed"))
        txt_cb_Validate 1, False
    End If
    '--NIVEL EDUCATIVO
    If NulosN(RstTmp("idnivedu")) <> 0 Then
        txt_cb(2).Text = NulosN(RstTmp("idnivedu"))
        txt_cb_Validate 2, False
    End If
    '--OCUPACION
    If NulosN(RstTmp("idocu")) <> 0 Then
        txt_cb(3).Text = NulosN(RstTmp("idocu"))
        txt_cb_Validate 3, False
    End If
    '--SUJETO A HORARIOS NOCTURNOS
    If NulosN(RstTmp("indica2")) = 0 Then
        opt_HoraNocturno(0).Value = True
    Else
        opt_HoraNocturno(1).Value = True
    End If
    '--CENTRO DE FORMACION PROFECIONAL
    If NulosN(RstTmp("idcenfor")) <> 0 Then
        txt_cb(4).Text = NulosN(RstTmp("idcenfor"))
        txt_cb_Validate 4, False
    End If
    '--ES MADRE CON RESPONSABILIDAD FAMILIAR
    If NulosN(RstTmp("indica1")) = 0 Then
        opt_Madre(0).Value = True
    Else
        opt_Madre(1).Value = True
    End If
    '--ES DISCAPACITADO
    If NulosN(RstTmp("indica3")) = 0 Then
        opt_discapacidad(0).Value = True
    Else
        opt_discapacidad(1).Value = True
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
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " los datos del  Prestador de Servicios - Modalidad Formativa ", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim xId As Integer

'    On Error GoTo LaCague

    xCon.BeginTrans

    '*****************************************************
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_categoria4 ; ", xCon
        RstCab.AddNew
    Else
        RST_Busq RstCab, "SELECT * FROM pla_categoria4 WHERE idemp =  " & mIdEmpleado & " ;", xCon
    End If
    
    RstCab("idemp") = mIdEmpleado
    '************************************************************TAB 0
''    '--TIPO PRESTADOR SERVICIO - MODALIDAD FORMATIVA DEFAULT
''    RstCab("idtiptra") = NulosN(txt_cb(0).Text) '23
    '--SEGURO MEDICO
    RstCab("idsegmed") = NulosN(txt_cb(1).Text)
    '--NIVEL EDUCATIVO
    RstCab("idnivedu") = NulosN(txt_cb(2).Text)
    '--OCUPACION
    RstCab("idocu") = NulosN(txt_cb(3).Text)
    '--CENTRO DE FORMACION PROFECIONAL
    RstCab("idcenfor") = NulosN(txt_cb(4).Text)
    '--ES MADRE CON RESPONSABILIDAD FAMILIAR
    If opt_Madre(0).Value = True Then RstCab("indica1") = 0
    If opt_Madre(1).Value = True Then RstCab("indica1") = 1
    '--SUJETO A HORARIO NOCTURNO
    If opt_HoraNocturno(0).Value = True Then RstCab("indica2") = 0
    If opt_HoraNocturno(1).Value = True Then RstCab("indica2") = 1
    '--ES MADRE CON RESPONSABILIDAD FAMILIAR
    If opt_discapacidad(0).Value = True Then RstCab("indica3") = 0
    If opt_discapacidad(1).Value = True Then RstCab("indica3") = 1
    '--
    RstCab.Update

    MsgBox "Los datos del Prestador de Servicios - Modalidad Formativa se " + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

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


