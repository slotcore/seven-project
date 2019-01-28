VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "aspatextboxfecha.ocx"
Begin VB.Form FrmManVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punto de Venta"
   ClientHeight    =   7035
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1740
      Index           =   12
      Left            =   11865
      TabIndex        =   115
      Top             =   1500
      Visible         =   0   'False
      Width           =   5100
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   8
         Left            =   2385
         Picture         =   "FrmManVenta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Seleccioene el Documento"
         Top             =   810
         Width           =   240
      End
      Begin VB.CommandButton CmdDocNum 
         Caption         =   "Aceptar"
         Height          =   465
         Index           =   0
         Left            =   975
         TabIndex        =   118
         Top             =   1170
         Width           =   1665
      End
      Begin VB.CommandButton CmdDocNum 
         Caption         =   "Cancelar"
         Height          =   465
         Index           =   1
         Left            =   2685
         TabIndex        =   117
         Top             =   1170
         Width           =   1665
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   7
         Left            =   2385
         Picture         =   "FrmManVenta.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Seleccione el Almacén"
         Top             =   420
         Width           =   240
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   6
         Left            =   1365
         MaxLength       =   10
         TabIndex        =   121
         Text            =   "txt_cb(6)"
         ToolTipText     =   "Ingrese el Nº. de Serie"
         Top             =   390
         Width           =   1290
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   7
         Left            =   1365
         MaxLength       =   150
         TabIndex        =   120
         Text            =   "txt_cb(7)"
         ToolTipText     =   "Ingrese el Nº. del Documento"
         Top             =   765
         Width           =   1290
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         Caption         =   "Nª Documento"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   128
         Top             =   870
         Width           =   1050
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
         Height          =   300
         Index           =   7
         Left            =   2655
         TabIndex        =   127
         Top             =   765
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.Label lbl_cb_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb_cod(7)"
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
         Index           =   7
         Left            =   3720
         TabIndex        =   126
         Top             =   750
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   15
         X1              =   0
         X2              =   6375
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line ln 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   14
         X1              =   -30
         X2              =   5655
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Line ln 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   13
         X1              =   5085
         X2              =   5085
         Y1              =   -270
         Y2              =   3500
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   12
         X1              =   0
         X2              =   15
         Y1              =   0
         Y2              =   3195
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00800000&
         X1              =   105
         X2              =   4905
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label lbl_cb_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb_cod(6)"
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
         Index           =   6
         Left            =   3705
         TabIndex        =   125
         Top             =   390
         Visible         =   0   'False
         Width           =   1230
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
         Height          =   300
         Index           =   6
         Left            =   2655
         TabIndex        =   124
         Top             =   390
         Width           =   2310
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serie Nº."
         Height          =   195
         Index           =   6
         Left            =   105
         TabIndex        =   123
         Top             =   495
         Width           =   630
      End
      Begin VB.Label lbl_titulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cambiar Nº de Documento"
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
         Index           =   2
         Left            =   75
         TabIndex        =   122
         Top             =   60
         Width           =   2250
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   45
         Top             =   15
         Width           =   5535
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   1260
      Index           =   11
      Left            =   135
      TabIndex        =   103
      Top             =   7830
      Visible         =   0   'False
      Width           =   7935
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   9
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   108
         Text            =   "txt_importe(9)"
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   1653
         Locked          =   -1  'True
         TabIndex        =   107
         Text            =   "txt_importe(8)"
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   3201
         Locked          =   -1  'True
         TabIndex        =   106
         Text            =   "txt_importe(7)"
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   4749
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "txt_importe(6)"
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   104
         Text            =   "txt_importe(5)"
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label lbl_titulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_titulo(1)"
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
         Index           =   1
         Left            =   90
         TabIndex        =   114
         Top             =   60
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   45
         Top             =   45
         Width           =   7830
      End
      Begin VB.Line ln 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   11
         X1              =   7905
         X2              =   7920
         Y1              =   -45
         Y2              =   3150
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   10
         X1              =   0
         X2              =   15
         Y1              =   0
         Y2              =   3195
      End
      Begin VB.Line ln 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   9
         X1              =   0
         X2              =   8030
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   8
         X1              =   -30
         X2              =   8000
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line2 
         X1              =   75
         X2              =   7695
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label lbl_importe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imp Afecto"
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
         Index           =   9
         Left            =   105
         TabIndex        =   113
         Top             =   615
         Width           =   930
      End
      Begin VB.Label lbl_importe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imp Inafecto"
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
         Index           =   8
         Left            =   1653
         TabIndex        =   112
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label lbl_importe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.G.V."
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
         Index           =   7
         Left            =   3201
         TabIndex        =   111
         Top             =   630
         Width           =   510
      End
      Begin VB.Label lbl_importe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.S.C."
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
         Index           =   6
         Left            =   4749
         TabIndex        =   110
         Top             =   630
         Width           =   495
      End
      Begin VB.Label lbl_importe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   5
         Left            =   6300
         TabIndex        =   109
         Top             =   615
         Width           =   450
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2730
      Index           =   10
      Left            =   11805
      TabIndex        =   70
      Top             =   3330
      Visible         =   0   'False
      Width           =   5100
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   5
         Left            =   2385
         Picture         =   "FrmManVenta.frx":0264
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Seleccione el Nº. de Serie"
         Top             =   1455
         Width           =   240
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   4
         Left            =   2385
         Picture         =   "FrmManVenta.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Seleccione el Almacén"
         Top             =   420
         Width           =   240
      End
      Begin VB.CommandButton CmdDoc 
         Caption         =   "Cancelar"
         Height          =   465
         Index           =   1
         Left            =   2685
         TabIndex        =   92
         Top             =   2205
         Width           =   1665
      End
      Begin VB.CommandButton CmdDoc 
         Caption         =   "Aceptar"
         Height          =   465
         Index           =   0
         Left            =   975
         TabIndex        =   91
         Top             =   2205
         Width           =   1665
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   2
         Left            =   2385
         Picture         =   "FrmManVenta.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Seleccioene el Documento"
         Top             =   1815
         Width           =   240
      End
      Begin VB.ComboBox cb_doc 
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   85
         ToolTipText     =   "Seleccione el Documento"
         Top             =   720
         Width           =   3555
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   2
         Left            =   1365
         MaxLength       =   150
         TabIndex        =   89
         Text            =   "txt_cb(2)"
         ToolTipText     =   "Ingrese el Nº. del Documento"
         Top             =   1770
         Width           =   1290
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   4
         Left            =   1365
         MaxLength       =   10
         TabIndex        =   83
         Text            =   "txt_cb(4)"
         ToolTipText     =   "Ingrese código de Almacén"
         Top             =   390
         Width           =   1290
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   5
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   87
         Text            =   "txt_cb(5)"
         ToolTipText     =   "Ingrese el Nº. de Serie"
         Top             =   1410
         Width           =   1290
      End
      Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
         Height          =   315
         Index           =   1
         Left            =   1365
         TabIndex        =   86
         ToolTipText     =   "Ingrese la Fecha de Emisión del Documento"
         Top             =   1065
         Width           =   1290
         _ExtentX        =   2275
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
         Valor           =   "  /  /    "
      End
      Begin VB.Label lbl_titulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_titulo(0)"
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
         Index           =   0
         Left            =   75
         TabIndex        =   74
         Top             =   60
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F&echa Registro:"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   102
         Top             =   1185
         Width           =   1125
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         Caption         =   "Serie Nº."
         Height          =   195
         Index           =   5
         Left            =   105
         TabIndex        =   101
         Top             =   1515
         Width           =   630
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
         Height          =   300
         Index           =   5
         Left            =   2655
         TabIndex        =   100
         Top             =   1410
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.Label lbl_cb_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb_cod(5)"
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
         Index           =   5
         Left            =   3735
         TabIndex        =   99
         Top             =   1395
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Almacén"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   98
         Top             =   495
         Width           =   615
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
         Left            =   2655
         TabIndex        =   97
         Top             =   390
         Width           =   2310
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
         Left            =   3705
         TabIndex        =   96
         Top             =   390
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   105
         X2              =   4905
         Y1              =   2145
         Y2              =   2145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Documento"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   75
         Top             =   840
         Width           =   1185
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   15
         Y1              =   0
         Y2              =   3195
      End
      Begin VB.Line ln 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   2
         X1              =   5085
         X2              =   5085
         Y1              =   -270
         Y2              =   3500
      End
      Begin VB.Line ln 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   -75
         X2              =   5610
         Y1              =   2715
         Y2              =   2715
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   6375
         Y1              =   15
         Y2              =   15
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
         Left            =   3720
         TabIndex        =   73
         Top             =   1755
         Visible         =   0   'False
         Width           =   1230
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
         Left            =   2655
         TabIndex        =   72
         Top             =   1770
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         Caption         =   "Nª Documento"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   71
         Top             =   1875
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   45
         Top             =   15
         Width           =   5535
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   2925
      Index           =   9
      Left            =   11805
      TabIndex        =   76
      Top             =   45
      Visible         =   0   'False
      Width           =   1920
      Begin VB.CommandButton CmdOtros 
         Caption         =   "&Imprimir "
         Height          =   465
         Index           =   4
         Left            =   60
         TabIndex        =   81
         Top             =   1845
         Width           =   1845
      End
      Begin VB.CommandButton CmdOtros 
         Caption         =   "Emitir &Documentos Anulados"
         Height          =   465
         Index           =   3
         Left            =   60
         TabIndex        =   80
         Top             =   1395
         Width           =   1845
      End
      Begin VB.CommandButton CmdOtros 
         Caption         =   "&Anular  Documento"
         Height          =   465
         Index           =   2
         Left            =   60
         TabIndex        =   79
         Top             =   945
         Width           =   1845
      End
      Begin VB.CommandButton CmdOtros 
         Caption         =   "&Eliminar Documento"
         Height          =   465
         Index           =   1
         Left            =   60
         TabIndex        =   78
         Top             =   495
         Width           =   1845
      End
      Begin VB.CommandButton CmdOtros 
         Caption         =   "&Cancelar"
         Height          =   465
         Index           =   5
         Left            =   60
         TabIndex        =   82
         Top             =   2385
         Width           =   1845
      End
      Begin VB.CommandButton CmdOtros 
         Caption         =   "&Modificar Documento"
         Height          =   465
         Index           =   0
         Left            =   60
         TabIndex        =   77
         Top             =   45
         Width           =   1845
      End
      Begin VB.Line ln 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   7
         X1              =   30
         X2              =   5715
         Y1              =   2910
         Y2              =   2910
      End
      Begin VB.Line ln 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   6
         X1              =   1905
         X2              =   1905
         Y1              =   15
         Y2              =   3500
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   5
         X1              =   -90
         X2              =   2820
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   4
         X1              =   15
         X2              =   30
         Y1              =   -435
         Y2              =   2760
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[Tipo de Cambio]"
      ForeColor       =   &H00400000&
      Height          =   690
      Left            =   5565
      TabIndex        =   55
      Top             =   150
      Width           =   1725
      Begin VB.Label lbl_TipoCambio 
         Alignment       =   2  'Center
         Caption         =   "lbl_TipoCambio"
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
         Height          =   255
         Left            =   60
         TabIndex        =   56
         Top             =   315
         Width           =   1605
      End
   End
   Begin VB.Frame fra 
      Height          =   810
      Index           =   5
      Left            =   5280
      TabIndex        =   32
      Top             =   5130
      Width           =   6435
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "txt_importe(4)"
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "txt_importe(3)"
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "txt_importe(2)"
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "txt_importe(1)"
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txt_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "txt_importe(0)"
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label lbl_importe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   4
         Left            =   5070
         TabIndex        =   42
         Top             =   135
         Width           =   450
      End
      Begin VB.Label lbl_importe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.S.C."
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
         Index           =   3
         Left            =   3780
         TabIndex        =   41
         Top             =   150
         Width           =   495
      End
      Begin VB.Label lbl_importe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.G.V."
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
         Index           =   2
         Left            =   2550
         TabIndex        =   40
         Top             =   150
         Width           =   510
      End
      Begin VB.Label lbl_importe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imp Inafecto"
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
         Index           =   1
         Left            =   1305
         TabIndex        =   39
         Top             =   150
         Width           =   1080
      End
      Begin VB.Label lbl_importe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imp Afecto"
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
         Index           =   0
         Left            =   60
         TabIndex        =   33
         Top             =   135
         Width           =   930
      End
   End
   Begin VB.Frame fra 
      ForeColor       =   &H00400000&
      Height          =   810
      Index           =   3
      Left            =   15
      TabIndex        =   31
      Top             =   5130
      Width           =   5130
      Begin VB.CommandButton cmd_item 
         Caption         =   "&Agregar"
         Height          =   450
         Index           =   0
         Left            =   75
         TabIndex        =   16
         ToolTipText     =   "Agregar Producto"
         Top             =   225
         Width           =   1515
      End
      Begin VB.CommandButton cmd_item 
         Caption         =   "&Eliminar"
         Height          =   450
         Index           =   1
         Left            =   1710
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Eliminar Producto Seleccionado"
         Top             =   225
         Width           =   1515
      End
   End
   Begin VB.Frame fra 
      ForeColor       =   &H00400000&
      Height          =   510
      Index           =   1
      Left            =   2865
      TabIndex        =   29
      Top             =   -30
      Width           =   2670
      Begin VB.TextBox txt_cotiza 
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   2775
         TabIndex        =   63
         Text            =   "txt_cotiza(0)"
         Top             =   150
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txt_cotiza 
         Height          =   300
         Index           =   1
         Left            =   1335
         TabIndex        =   6
         Text            =   "txt_cotiza(1)"
         ToolTipText     =   "Ingrese el Nº de Cotización"
         Top             =   150
         Width           =   1185
      End
      Begin VB.Label lbl_cotiza 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cotización"
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
         Index           =   0
         Left            =   105
         TabIndex        =   30
         Top             =   255
         Width           =   1170
      End
   End
   Begin VB.Frame fra 
      ForeColor       =   &H00400000&
      Height          =   1350
      Index           =   0
      Left            =   15
      TabIndex        =   23
      Top             =   -30
      Width           =   11700
      Begin VB.CommandButton cb 
         Enabled         =   0   'False
         Height          =   240
         Index           =   3
         Left            =   2010
         Picture         =   "FrmManVenta.frx":05FA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Seleccione el Almacén"
         Top             =   585
         Width           =   240
      End
      Begin VB.CommandButton Cmd_Cliente 
         Height          =   270
         Left            =   2310
         Picture         =   "FrmManVenta.frx":072C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Nuevo Cliente"
         Top             =   930
         Width           =   270
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   0
         Left            =   2025
         Picture         =   "FrmManVenta.frx":082E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Seleccione al Cliente"
         Top             =   945
         Width           =   225
      End
      Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
         Height          =   315
         Index           =   0
         Left            =   945
         TabIndex        =   7
         ToolTipText     =   "Ingrese fecha del Emisión"
         Top             =   225
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
         Valor           =   "  /  /    "
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   945
         MaxLength       =   11
         TabIndex        =   10
         Text            =   "txt_cb(0)"
         ToolTipText     =   "Ingrese wel Nº. R.U.C. del Cliente"
         Top             =   915
         Width           =   1350
      End
      Begin VB.Frame fra 
         Height          =   1230
         Index           =   2
         Left            =   7365
         TabIndex        =   26
         Top             =   120
         Width           =   4290
         Begin VB.Label lbl_codigo 
            BackColor       =   &H008080FF&
            Caption         =   "lbl_codigo"
            Height          =   315
            Left            =   0
            TabIndex        =   69
            Top             =   0
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label lbl_num 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_num(1)"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Index           =   1
            Left            =   2055
            TabIndex        =   58
            Top             =   960
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label lbl_num 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_num(0)"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Index           =   0
            Left            =   405
            TabIndex        =   57
            Top             =   960
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label lbl_encabezado 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_encabezado(1)"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   28
            Top             =   645
            Width           =   4005
         End
         Begin VB.Label lbl_encabezado 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "lbl_encabezado(0)"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   390
            Index           =   0
            Left            =   90
            TabIndex        =   27
            Top             =   210
            Width           =   4005
         End
      End
      Begin VB.TextBox txt_cb 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   945
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "txt_cb(3)"
         ToolTipText     =   "Ingrese código de Almacén"
         Top             =   555
         Width           =   1350
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
         Height          =   315
         Index           =   3
         Left            =   4200
         TabIndex        =   95
         Top             =   570
         Visible         =   0   'False
         Width           =   1230
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
         Height          =   315
         Index           =   3
         Left            =   2325
         TabIndex        =   94
         Top             =   555
         Width           =   3195
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Almacén"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   93
         Top             =   660
         Width           =   615
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
         Left            =   5745
         TabIndex        =   64
         Top             =   945
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nª R.U.C."
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   25
         Top             =   1020
         Width           =   705
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F&echa:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   24
         Top             =   345
         Width           =   495
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
         Left            =   2640
         TabIndex        =   66
         Top             =   930
         Width           =   4590
      End
   End
   Begin VB.Frame fra_MsgAcceso 
      Caption         =   "[Control de Accesos]"
      ForeColor       =   &H000000C0&
      Height          =   840
      Left            =   11805
      TabIndex        =   67
      Top             =   6135
      Visible         =   0   'False
      Width           =   7590
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6750
         Picture         =   "FrmManVenta.frx":0960
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lbl_MsgAcceso 
         Caption         =   "No puede registrar Ventas: Consulte con el Adiministrador para que le otorge permisos"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   135
         TabIndex        =   68
         Top             =   225
         Width           =   6495
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   3150
      Left            =   15
      TabIndex        =   15
      Top             =   1905
      Width           =   11700
      _cx             =   20637
      _cy             =   5556
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
      Rows            =   2
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmManVenta.frx":0DA2
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
   Begin VB.Frame fra 
      Height          =   555
      Index           =   6
      Left            =   15
      TabIndex        =   43
      Top             =   6000
      Width           =   11700
      Begin VB.CommandButton CmdTipo 
         Caption         =   "&Otros"
         Height          =   465
         Index           =   5
         Left            =   7818
         TabIndex        =   4
         Top             =   60
         Width           =   1920
      End
      Begin VB.CommandButton CmdTipo 
         Caption         =   "&Salir"
         Height          =   465
         Index           =   4
         Left            =   9765
         TabIndex        =   5
         Top             =   60
         Width           =   1920
      End
      Begin VB.CommandButton CmdTipo 
         Caption         =   "&Grabar"
         Height          =   465
         Index           =   3
         Left            =   5871
         TabIndex        =   3
         Top             =   60
         Width           =   1920
      End
      Begin VB.CommandButton CmdTipo 
         Caption         =   "&Boleta"
         Height          =   465
         Index           =   2
         Left            =   3924
         TabIndex        =   2
         Top             =   60
         Width           =   1920
      End
      Begin VB.CommandButton CmdTipo 
         Caption         =   "&Factura"
         Height          =   465
         Index           =   1
         Left            =   1977
         TabIndex        =   1
         Top             =   60
         Width           =   1920
      End
      Begin VB.CommandButton CmdTipo 
         Caption         =   "&Cotización"
         Height          =   465
         Index           =   0
         Left            =   30
         TabIndex        =   0
         Top             =   60
         Width           =   1920
      End
      Begin VB.PictureBox picTipo 
         BackColor       =   &H000000FF&
         Height          =   465
         Index           =   0
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   1860
         TabIndex        =   44
         Top             =   60
         Width           =   1920
         Begin VB.Label lblTipo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Cotización"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   540
            TabIndex        =   45
            Top             =   90
            Width           =   1110
         End
      End
      Begin VB.PictureBox picTipo 
         BackColor       =   &H000000FF&
         Height          =   465
         Index           =   1
         Left            =   1977
         ScaleHeight     =   405
         ScaleWidth      =   1860
         TabIndex        =   46
         Top             =   60
         Visible         =   0   'False
         Width           =   1920
         Begin VB.Label lblTipo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Factura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   690
            TabIndex        =   47
            Top             =   90
            Width           =   810
         End
      End
      Begin VB.PictureBox picTipo 
         BackColor       =   &H000000FF&
         Height          =   465
         Index           =   2
         Left            =   3924
         ScaleHeight     =   405
         ScaleWidth      =   1860
         TabIndex        =   48
         Top             =   60
         Visible         =   0   'False
         Width           =   1920
         Begin VB.Label lblTipo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Boleta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   750
            TabIndex        =   49
            Top             =   90
            Width           =   705
         End
      End
   End
   Begin VB.Frame fra 
      Height          =   480
      Index           =   7
      Left            =   15
      TabIndex        =   59
      Top             =   6495
      Width           =   11700
      Begin VB.Label lbl_usuario 
         Alignment       =   1  'Right Justify
         Caption         =   "lbl_usuario(2)"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   2
         Left            =   9120
         TabIndex        =   62
         Top             =   165
         Width           =   2490
      End
      Begin VB.Label lbl_usuario 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "lbl_usuario(0)"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   4845
         TabIndex        =   61
         Top             =   165
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label lbl_usuario 
         Caption         =   "lbl_usuario(1)"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   60
         Top             =   165
         Width           =   9330
      End
   End
   Begin VB.Frame fra 
      Height          =   555
      Index           =   8
      Left            =   0
      TabIndex        =   51
      Top             =   1305
      Width           =   4290
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   1
         Left            =   2025
         Picture         =   "FrmManVenta.frx":0F64
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Seleccione la Moneda"
         Top             =   225
         Width           =   240
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   1
         Left            =   975
         MaxLength       =   4
         TabIndex        =   13
         Text            =   "txt_cb(1)"
         ToolTipText     =   "Ingrese el Código de la Moneda"
         Top             =   195
         Width           =   1350
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         Caption         =   "&Moneda"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   54
         Top             =   255
         Width           =   585
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
         Left            =   2325
         TabIndex        =   53
         Top             =   195
         Width           =   1890
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
         Left            =   2955
         TabIndex        =   52
         Top             =   210
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      Height          =   555
      Left            =   8880
      TabIndex        =   65
      Top             =   1305
      Width           =   2790
      Begin VB.OptionButton opt_descuento1 
         Caption         =   "Valor"
         Height          =   195
         Index           =   1
         Left            =   1905
         TabIndex        =   22
         Top             =   270
         Width           =   690
      End
      Begin VB.OptionButton opt_descuento1 
         Caption         =   "Porcentaje"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   21
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.Frame fra 
      Caption         =   "[Selecc. Descuento ] "
      ForeColor       =   &H00400000&
      Height          =   555
      Index           =   4
      Left            =   4350
      TabIndex        =   50
      ToolTipText     =   "Presione F7 para seleccionar un descuento"
      Top             =   1305
      Width           =   7365
      Begin VB.OptionButton opt_descuento 
         Caption         =   "Ninguno"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   18
         Top             =   270
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton opt_descuento 
         Caption         =   "General"
         Height          =   195
         Index           =   1
         Left            =   1605
         TabIndex        =   19
         Top             =   270
         Width           =   900
      End
      Begin VB.OptionButton opt_descuento 
         Caption         =   "Corporativo"
         Height          =   195
         Index           =   2
         Left            =   3000
         TabIndex        =   20
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar Producto"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "Seleccionar Producto"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Eliminar Producto"
      End
      Begin VB.Menu Menu1_5 
         Caption         =   "Eliminar Todo"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "&Anular"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_4 
         Caption         =   "&Emitir Documentos Anulados"
      End
      Begin VB.Menu Menu2_5 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_6 
         Caption         =   "&Imprimir Ticket"
      End
   End
End
Attribute VB_Name = "FrmManVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean

Dim Agregando As Boolean
'-------------
Dim mImpuestoValor  As Double
Dim nImpuestoNombre As String
Dim nTipoCambio As Double
Dim nDocumento As String        '--NOMBRE DEL DOCUMENTO
Dim mIdDocumento As Integer     '--1::COTIZACION, 2::FACTURA, 3::BOLETA DE VENTA
Dim mMesActivo As Integer       '--ESPECIFICA EL MES ACTIVO

Dim mIdCuentaDoc As Integer
Dim mIdCuentaImpuesto As Integer

Dim mNivelAcceso As Integer     '--1::Vendedor, 2::Cajero,3::Supervisor
Dim mIdEventoOtrosBotones As Integer '--IDENTIFICA EL TIPO DE ACCION
    '-1::NO SELECCIONA ALGUN BOTON  1:: MODIFICAR, 2::ELIMINAR  3::ANULAR  4::EMITIR DOCUMENTOS ANULADOS  5::IMPRIMIR TICKET

Private Sub cb_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim nSQL As String
    Dim nTitulo As String
    
    Dim mIdDoc As Integer
    Dim nSQLEvento As String
    Dim mIdAlm As Integer '--COD ALMACEN
    'On Error GoTo error

    Select Case Index
        Case 0 '--CLIENTE
            nSQL = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id " _
            + vbCr + " FROM mae_cliente " _
            + vbCr + " WHERE (((mae_cliente.activo) = -1)) " _
            + vbCr + " ORDER BY mae_cliente.nombre;"
        
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":       xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "R.U.C.":       xCampos(1, 1) = "numruc":    xCampos(1, 2) = "1500":   xCampos(1, 3) = "C"
            nTitulo = "Buscando cliente"
        Case 1 '--MONEDA
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Moneda":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "3500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Símbolo":   xCampos(1, 1) = "simbolo":    xCampos(1, 2) = "1000":   xCampos(1, 3) = "C"
            
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion as nombre,mae_moneda.id as cod,mae_moneda.simbolo  " _
            + vbCr + " From mae_moneda "
            nTitulo = "Buscando Moneda"
        
        Case 3 '--ALMACEN
            ReDim xCampos(1, 3) As String
            xCampos(0, 0) = "Almacén":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            
            nSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion AS nombre, alm_almacenes.id AS cod " _
            + vbCr + " FROM alm_almacenes ORDER BY alm_almacenes.descripcion;"
            nTitulo = "Buscando Almacén"
        '--------------------------------------------
        Case 5, 6 '--MUM SERIE
            If Index = 5 Then
                If NulosN(lbl_cb_cod(4).Caption) = 0 Then
                    MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
                    txt_cb(4).SetFocus
                    Exit Sub
                End If
                If cb_doc.ListIndex = -1 Then
                    MsgBox "Seleccioe un Documento" + vbCr + "Luego Proceda", vbExclamation, xTitulo
                    cb_doc.SetFocus
                    Exit Sub
                End If
                If UCase(cb_doc.Text) = "FACTURA" Then mIdDoc = 1 '--FACTURA
                If UCase(cb_doc.Text) = "BOLETA DE VENTA" Then mIdDoc = 3 '--BOLETA DE VENTA
                mIdAlm = NulosN(lbl_cb_cod(4).Caption)
            Else
                If NulosN(lbl_cb_cod(3).Caption) = 0 Then
                    MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
                    Exit Sub
                End If
                mIdAlm = NulosN(lbl_cb_cod(3).Caption)
            End If
            
            ReDim xCampos(1, 3) As String
            xCampos(0, 0) = "Número":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            
            nSQL = "SELECT Format([alm_numseries].[numser],'0000') AS nombre,alm_numseries.id , Format([alm_numseries].[numser],'0000') AS cod " _
                + vbCr + " FROM alm_numseries " _
                + vbCr + " WHERE alm_numseries.idtipdoc=" + CStr(mIdDoc) + " AND alm_numseries.idalm=" + CStr(mIdAlm) + ";"
            
            nTitulo = "Buscando Series"
            
        Case 2 '--NUMERO DOCUMENTO
        
            If NulosN(lbl_cb_cod(4).Caption) = 0 Then
                MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
                txt_cb(4).SetFocus
                Exit Sub
            End If
            If cb_doc.ListIndex = -1 Then
                MsgBox "Seleccioe un Documento" + vbCr + "Luego Proceda", vbExclamation, xTitulo
                cb_doc.SetFocus
                Exit Sub
            End If
        
            ReDim xCampos(5, 3) As String
            
            xCampos(0, 0) = "NºDocumento":  xCampos(0, 1) = "nombre":       xCampos(0, 2) = "1300":     xCampos(0, 3) = "C":
            xCampos(1, 0) = "Fch. Doc":     xCampos(1, 1) = "fchdoc":       xCampos(1, 2) = "950":     xCampos(1, 3) = "C":
            xCampos(2, 0) = "Cliente":      xCampos(2, 1) = "clidesc":      xCampos(2, 2) = "4000":    xCampos(2, 3) = "C":
            xCampos(3, 0) = "M":            xCampos(3, 1) = "simbolo":      xCampos(3, 2) = "500":     xCampos(3, 3) = "C":
            xCampos(4, 0) = "Importe":      xCampos(4, 1) = "imptotdoc":    xCampos(4, 2) = "1200":    xCampos(4, 3) = "N":

            If UCase(cb_doc.Text) <> "COTIZACIÓN" Then
                If UCase(cb_doc.Text) = "FACTURA" Then mIdDoc = 1 '--FACTURA
                If UCase(cb_doc.Text) = "BOLETA DE VENTA" Then mIdDoc = 3 '--BOLETA DE VENTA
                If mIdEventoOtrosBotones <= 3 Then '--SOLO MODIFICAR(1) , ELIMINAR(2), ANULAR(3)
                    nSQLEvento = " AND vta_ventas.evento=" + CStr(mIdEventoOtrosBotones) + " "
                End If
                nSQL = "SELECT  vta_ventas.numdoc as nombre,vta_ventas.id,vta_ventas.id as cod, format(vta_ventas.fchdoc,'mm/dd/yy') as fchdoc, mae_moneda.simbolo, vta_ventas.imptotdoc, mae_cliente.nombre AS clidesc " _
                    + vbCr + " FROM (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id " _
                    + vbCr + " WHERE vta_ventas.numser ='" + lbl_cb_cod(5).Caption + "'  AND vta_ventas.idalm=" + CStr(NulosN(lbl_cb_cod(4).Caption)) + " AND vta_ventas.tipdoc=" + CStr(mIdDoc) + " AND vta_ventas.anulado=0 " _
                           + nSQLEvento + _
                    vbCr + " ORDER BY vta_ventas.numdoc;"
            Else
                mIdDoc = 3 '--COTIZACION
                nSQL = "SELECT pvt_cotizacion.numdoc as nombre ,pvt_cotizacion.id, pvt_cotizacion.id as cod,  format(pvt_cotizacion.fchdoc,'dd/mm/yy') as fchdoc, mae_moneda.simbolo, pvt_cotizacion.imptotdoc, mae_cliente.nombre AS clidesc " _
                    + vbCr + " FROM (mae_cliente RIGHT JOIN pvt_cotizacion ON mae_cliente.id = pvt_cotizacion.idcli) LEFT JOIN mae_moneda ON pvt_cotizacion.idmon = mae_moneda.id " _
                    + vbCr + " WHERE (pvt_cotizacion.iddocven Is Null Or pvt_cotizacion.iddocven = 0) And pvt_cotizacion.anulado = 0"
            End If

            
    End Select
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    '--SI ESTA EN LA OPCION DE SELECIONAR NUM. DOC, ACTUALIZAR LA FECHA DE REGISTRO
    If Index = 2 Then '--
        txtfecha(1).Valor = CDate(xRs.Fields("fchdoc"))
        CmdDoc(0).Enabled = True
    End If
    '--SI SELECCIONA UNA MONEDA CARGAR LAS CUENTAS CONTABLES
    If Index = 1 And txt_cb(Index) <> "" Then pCargarCuentaContable mIdCuentaDoc, mIdCuentaImpuesto, mIdDocumento, NulosN(lbl_cb_cod(3).Caption), NulosN(CStr(lbl_cb_cod(1).Caption))
    '---------------------------------------------------------------------
Salir:
    SendKeys vbTab
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub


Private Sub REGISTRO_ADD(Optional F_SELECCION_VARIOS As Boolean = False)

    '--GENERAR EL WHERE DE LOS ID'S RECETA PARA QUE NO SE REPITAN
    Dim SQL_ITEM As String
    SQL_ITEM = GENERAR_SQL_ID(Fg1, 1, "alm_inventario.id", "NOT IN")
    If SQL_ITEM <> "" Then SQL_ITEM = " AND " + SQL_ITEM
    '----
    On Error GoTo error
    Dim xRs  As New ADODB.Recordset
    Dim nSQL As String
    ReDim xCampos(5, 4) As String
    
    xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "5000":     xCampos(0, 3) = "C":
    xCampos(1, 0) = "M":                xCampos(1, 1) = "simbolo":      xCampos(1, 2) = "600":      xCampos(1, 3) = "C":
    xCampos(2, 0) = "U.M.":             xCampos(2, 1) = "abrev":        xCampos(2, 2) = "600":      xCampos(2, 3) = "C":
    xCampos(3, 0) = "Tipo Producto":    xCampos(3, 1) = "tipprodesc":   xCampos(3, 2) = "1100":     xCampos(3, 3) = "C":
    xCampos(4, 0) = "Stock":            xCampos(4, 1) = "stckact":      xCampos(4, 2) = "1000":     xCampos(4, 3) = "N":


    nSQL = "SELECT alm_inventario.id AS iditem, alm_inventario.idunimed, alm_inventario.idmon, alm_inventario.descripcion, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.codpro, mae_moneda.simbolo, mae_unidades.abrev, pvt_items.precio, alm_inventario.stckact, alm_inventario.stckmin, alm_inventario.idtipven, alm_inventario.idcuentaven " _
        + vbCr + " FROM mae_tipoproducto RIGHT JOIN ((mae_unidades RIGHT JOIN (mae_moneda RIGHT JOIN alm_inventario ON mae_moneda.id = alm_inventario.idmon) ON mae_unidades.id = alm_inventario.idunimed) INNER JOIN pvt_items ON alm_inventario.id = pvt_items.iditem) ON mae_tipoproducto.id = alm_inventario.tippro " _
        + vbCr + " WHERE (((pvt_items.activo) = -1)) " + SQL_ITEM _
        + vbCr + " ORDER BY alm_inventario.descripcion;"

    If F_SELECCION_VARIOS = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Item"
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Item", "descripcion", "descripcion", Principio
    End If

    Agregando = True
    Dim A As Integer
    Dim xFila As Integer
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    If F_SELECCION_VARIOS = True Then xRs.MoveFirst
    Do While Not xRs.EOF
        With Fg1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = xRs.Fields("iditem") & ""
            .TextMatrix(.Rows - 1, 2) = xRs.Fields("idunimed") & ""
            .TextMatrix(.Rows - 1, 3) = xRs.Fields("descripcion") & ""
            .TextMatrix(.Rows - 1, 4) = xRs.Fields("abrev") & ""
            .TextMatrix(.Rows - 1, 6) = xRs.Fields("precio") & ""
            .TextMatrix(.Rows - 1, 9) = xRs.Fields("idtipven") & ""
            .TextMatrix(.Rows - 1, 12) = xRs.Fields("idcuentaven") & ""
            If F_SELECCION_VARIOS = False Then Exit Do
            '---
        End With
        If F_SELECCION_VARIOS = False Then Exit Do
        xRs.MoveNext
    Loop
Salir:
    Agregando = False
    If Fg1.Rows >= 2 Then Fg1.Row = Fg1.Rows - 1: Fg1.Col = 5:   'Fg1_RowColChange
    Set xRs = Nothing
    Fg1.SetFocus
    Exit Sub
error:
    Agregando = False
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "REGISTRO_ADD"
End Sub


Private Sub REGISTRO_DEL()
    If Fg1.Row < 0 Then Exit Sub
    If Fg1.Row = 0 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una correcta", vbExclamation
        Exit Sub
    End If
    If MsgBox("Seguro desea eliminar el item", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub


    '--ELIMINAR EL PRODUCTO
    Fg1.RemoveItem (Fg1.Row)
    If Fg1.Rows > 1 Then Fg1.Row = 1
    '---------------------------
    pCalculosTotales
End Sub

Private Sub cb_doc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    If txtfecha(1).Enabled = True Then
        txtfecha(1).SetFocus
    ElseIf txt_cb(5).Enabled = True Then
        txt_cb(5).SetFocus
    Else
        txt_cb(2).SetFocus
    End If
End Sub

Private Sub Cmd_Cliente_Click()
    On Error GoTo error
    '------------
    Dim ID_CLIENTE As Long
    txt_cb(0).Text = ""
    lbl_cb_cod(0).Caption = ""
    lbl_cb(0).Caption = ""
    MsgBox "En Contrucción", vbExclamation
    Exit Sub
    ID_CLIENTE = 10
    If ID_CLIENTE = -1 Then Exit Sub
    '------------
    
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    nSQL = "SELECT  mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id " _
        + vbCr + " FROM mae_cliente " _
        + vbCr + " WHERE (((mae_cliente.id)=" + CStr(ID_CLIENTE) + "));"


    If xCon.State = 0 Then Exit Sub
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.State = 0 Then Exit Sub
    If RstTmp.RecordCount > 0 Then
        txt_cb(0) = RstTmp.Fields(0) & "" '--TEXTO A MOSTRAR
        lbl_cb(0).Caption = RstTmp.Fields(1) & ""  '--NOMBRE
        lbl_cb_cod(0).Caption = RstTmp.Fields(2) & "" '--CODIGO
    Else
        txt_cb(0).Text = "":   lbl_cb(0).Caption = "":    lbl_cb_cod(0).Caption = ""
    End If
    cmd_item(0).SetFocus
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "cmd_cliente"
End Sub

Private Sub cmd_item_Click(Index As Integer)
    Select Case Index
        Case 0 '--AGREGAR
            REGISTRO_ADD
        Case 1 '--ELIMINAR REGISTROS AGREGADOS
            REGISTRO_DEL
    End Select
End Sub




Private Sub CmdDoc_Click(Index As Integer)
    '------EMPEZAMOS A VALIDAR QUE LOS DATOS INGRESADOS SEAN CORRECTO
    If Index = 1 Then '--CANCELAR
        pActivarObjetos True
        Exit Sub
    End If
    If NulosN(lbl_cb_cod(2).Caption) = 0 And mIdEventoOtrosBotones <> 4 Then
        MsgBox "Seleccione Otra vez el Nº de Documento", vbExclamation, xTitulo
        Exit Sub
    End If
    
    If NulosN(lbl_cb_cod(4).Caption) = 0 Then
        MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
        txt_cb(4).SetFocus
        Exit Sub
    End If
    If cb_doc.ListIndex = -1 Then
        MsgBox "Seleccioe un Documento" + vbCr + "Luego Proceda", vbExclamation, xTitulo
        cb_doc.SetFocus
        Exit Sub
    End If
    
    If IsDate(txtfecha(1).Valor) = False And mIdEventoOtrosBotones = 4 Then '--SOLO CUANDO SE EMITEN DOCUMENTOS ANULADOS
        MsgBox "Ingrese la Fecha de Registro", vbExclamation, xTitulo
        txtfecha(1).SetFocus
        Exit Sub
    End If
    
    If mNivelAcceso <> 1 Then '--SI ES DIFERENTE DE VENDEDOR
        If NulosN(lbl_cb_cod(5).Caption) = 0 Then
            MsgBox "Seleccione en Nº de Serie", vbExclamation, xTitulo
            txt_cb(5).SetFocus
            Exit Sub
        End If
        
        
        If NulosN(lbl_cb_cod(5).Caption) = 0 Then
            MsgBox "Seleccione en Nº del Documento", vbExclamation, xTitulo
            txt_cb(2).SetFocus
            Exit Sub
        End If
    End If
    
    '---------- SI TODO ESTA CORRECTO COMENZAMOS A EJECUTAR LOS PROCESOS ----------
    '--COLOCANDO EL CODIGO DEL DOCUMENTO
    If UCase(cb_doc.Text) = "COTIZACIÓN" Then
        lbl_codigo.Caption = NulosN(txt_cb(2).Text)
    Else
        lbl_codigo.Caption = NulosN(lbl_cb_cod(2).Caption)
    End If
    '------------------------------------------------------------------------------
    Select Case mIdEventoOtrosBotones '-- SE CARGAR EN CmdOtros_Click()
        Case 1 '--MODIFICAR
            If Modificar() = False Then Exit Sub
        Case 2 '--ELIMINAR
            If Eliminar() = False Then Exit Sub
            
        Case 3 '--ANULAR
            If Anular() = False Then Exit Sub

        Case 4 '--EMITIR DOCUMENTOS ANULADOS
            If EmitirAnulada() = False Then Exit Sub
            
        Case 5 '--IMPRIMIR TICKET
            Select Case UCase(cb_doc.Text)
                Case "COTIZACIÓN"
                    Imprimir 0
                Case "FACTURA"
                    Imprimir 1
                Case "BOLETA DE VENTA"
                    Imprimir 3
                Case Else '--CANCELAR
                    MsgBox "No se puede Imprimir, proque se necesita un Documento", vbExclamation, xTitulo
                    Exit Sub
            End Select
    End Select
    
    pActivarObjetos True
    '--------------------
    
End Sub

Private Sub CmdOtros_Click(Index As Integer)
    Select Case Index
        Case 0 '--MODIFICAR
            mIdEventoOtrosBotones = 1
            pCargarSegundaVentana Index
        Case 1 '--ELIMINAR
            mIdEventoOtrosBotones = 2
        Case 2 '--ANULAR
            mIdEventoOtrosBotones = 3
        Case 3 '--EMITIR DOCUMENTOS ANULADOS
            mIdEventoOtrosBotones = 4
        Case 4 '--IMPRIMIR TICKET
            mIdEventoOtrosBotones = 5
        Case 5 '--CANCELAR
            mIdEventoOtrosBotones = -1
            pActivarObjetos True
    End Select
    If Index <> 5 Then
        fra(9).Visible = False '--OCULTAR BOTONES OTROS
        pCargarSegundaVentana Index
        CmdDoc(0).Enabled = False
    End If
End Sub


Private Sub CmdTipo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> 6 Then Exit Sub
    If Button = 2 Then PopupMenu Menu2
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim fSinCalculos As Boolean  '--INDICA SI SE APLICA LOS CALCULOS CUANDO CUANDO SE APLICAN LOS DESCUENTOS
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    fSinCalculos = False
    Select Case Col
        Case 5, 6
            If IsNumeric(Fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
            Else
                '--REVIZAR EL STOCK
                If Col = 5 Then
                    Dim mSaldo As Double
                    Dim RstTmp As New ADODB.Recordset
                    
                    RST_Busq RstTmp, "SELECT stckact,stckmin From alm_inventario WHERE id = " + CStr(Fg1.TextMatrix(Row, 1)) + ";", xCon
                    mSaldo = NulosN(RstTmp.Fields("stckact")) - NulosN(Fg1.TextMatrix(Row, 5))
                    If mSaldo < NulosN(RstTmp.Fields("stckmin")) Then
                        MsgBox "No hay suficiente stock " & Chr(13) & "Producto : " + Fg1.TextMatrix(Row, 3) & Chr(13) _
                            & "Cantidad Solicitada : " + Trim(Fg1.TextMatrix(Row, 5)) + Chr(13) _
                            & "Stock Actual  : " + Trim(Format(NulosN(RstTmp.Fields("stckact")), "0.00")) + Chr(13) _
                            & "Faltante        : " + Trim(Str(Format(mSaldo - NulosN(RstTmp.Fields("stckmin")), "0.00"))) + Chr(13), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                        Fg1.TextMatrix(Row, 5) = 0
                        Fg1.TextMatrix(Row, 7) = 0
                    End If
                    Set RstTmp = Nothing
                End If
                '--DEL DESCUENTO
                If opt_descuento(0).Value = True Then
                    Fg1.TextMatrix(Row, 7) = 0
                    Fg1.TextMatrix(Row, 13) = 0
                Else
                    pAplicarDescuento NulosN(Fg1.TextMatrix(Row, 1))
gEfecturaCalculos:
                    Fg1.TextMatrix(Row, 8) = NulosN(Fg1.TextMatrix(Row, 5)) * NulosN(Fg1.TextMatrix(Row, 6)) 'COSTO=CANT * P.U.
                    '--OBTENER EL DESCUENTO EN IMPORTE
                    If opt_descuento1(0).Value = True Then '--PORCENTAJE
                        Fg1.TextMatrix(Row, 13) = NulosN(Fg1.TextMatrix(Row, 8)) * NulosN(Fg1.TextMatrix(Row, 7))
                    Else '--VALOR
                        Fg1.TextMatrix(Row, 13) = NulosN(Fg1.TextMatrix(Row, 7))
                    End If
                End If

                '--CANTIDAD * PRECIO UNITARIO - DESCUENTO EN IMPORTE
                Fg1.TextMatrix(Row, 8) = (NulosN(Fg1.TextMatrix(Row, 5)) * NulosN(Fg1.TextMatrix(Row, 6))) - NulosN(Fg1.TextMatrix(Row, 13))
                '--ES GRAVADA O NO
                If NulosN(Fg1.TextMatrix(Row, 9)) <> 0 And NulosN(Fg1.TextMatrix(Row, 9)) <> 3 Then
                    Fg1.TextMatrix(Row, 10) = NulosN(Fg1.TextMatrix(Row, 8))
                    Fg1.TextMatrix(Row, 11) = 0
                Else
                    Fg1.TextMatrix(Row, 10) = 0
                    Fg1.TextMatrix(Row, 11) = NulosN(Fg1.TextMatrix(Row, 8))
                End If
                '--NUEVO PRECIO UNITARIO
                If NulosN(Fg1.TextMatrix(Row, 8)) = 0 Or NulosN(Fg1.TextMatrix(Row, 5)) = 0 Then
                    Fg1.TextMatrix(Row, 14) = 0
                Else
                    Fg1.TextMatrix(Row, 14) = NulosN(Fg1.TextMatrix(Row, 8)) / NulosN(Fg1.TextMatrix(Row, 5))
                End If
                '--HACE LOS CALCULOS
                If fSinCalculos = False Then pCalculosTotales
                '----------------------
            End If
        Case 7 '--DESCUENTO
            fSinCalculos = True
            GoTo gEfecturaCalculos:
    End Select
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Fg1_CellChanged"
End Sub


Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub
    If Col <> 3 Then Exit Sub
    If Row < 1 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    ReDim xCampos(5, 4) As String
    '--GENERAR EL WHERE DE LOS ID'S RECETA PARA QUE NO SE REPITAN
    Dim SQL_ITEM As String
    SQL_ITEM = GENERAR_SQL_ID(Fg1, 1, "alm_inventario.id", "NOT IN")
    If SQL_ITEM <> "" Then SQL_ITEM = " AND " + SQL_ITEM
    '----

    xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "5000":     xCampos(0, 3) = "C":
    xCampos(1, 0) = "M":                xCampos(1, 1) = "simbolo":      xCampos(1, 2) = "600":      xCampos(1, 3) = "C":
    xCampos(2, 0) = "U.M.":             xCampos(2, 1) = "abrev":        xCampos(2, 2) = "600":      xCampos(2, 3) = "C":
    xCampos(3, 0) = "Tipo Producto":    xCampos(3, 1) = "tipprodesc":   xCampos(3, 2) = "1100":     xCampos(3, 3) = "C":
    xCampos(4, 0) = "Stock":            xCampos(4, 1) = "stckact":      xCampos(4, 2) = "1000":     xCampos(4, 3) = "N":

    nSQL = "SELECT alm_inventario.id AS iditem, alm_inventario.idunimed, alm_inventario.idmon, alm_inventario.descripcion, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.codpro, mae_moneda.simbolo, mae_unidades.abrev, pvt_items.precio, alm_inventario.stckact, alm_inventario.stckmin, alm_inventario.idtipven, alm_inventario.idcuentaven " _
        + vbCr + " FROM mae_tipoproducto RIGHT JOIN ((mae_unidades RIGHT JOIN (mae_moneda RIGHT JOIN alm_inventario ON mae_moneda.id = alm_inventario.idmon) ON mae_unidades.id = alm_inventario.idunimed) INNER JOIN pvt_items ON alm_inventario.id = pvt_items.iditem) ON mae_tipoproducto.id = alm_inventario.tippro " _
        + vbCr + " WHERE (((pvt_items.activo) = -1)) " + SQL_ITEM _
        + vbCr + " ORDER BY alm_inventario.descripcion;"
        
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Item's", "descripcion", "descripcion", Principio, ""

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    Agregando = True
    With Fg1
        .TextMatrix(Row, 1) = xRs.Fields("iditem") & ""
        .TextMatrix(Row, 2) = xRs.Fields("idunimed") & ""
        .TextMatrix(Row, 3) = xRs.Fields("descripcion") & ""
        .TextMatrix(Row, 4) = xRs.Fields("abrev") & ""
        .TextMatrix(Row, 6) = xRs.Fields("precio") & ""
        .TextMatrix(Row, 9) = xRs.Fields("idtipven") & ""
        .TextMatrix(Row, 12) = xRs.Fields("idcuentaven") & ""
        If .Rows >= 1 Then .Row = Row: .Col = 5:
    End With
    '--EFECTUAR EL CALCULO SI CANTIDAD ES DIF<>0
    Agregando = False
    If NulosN(Fg1.TextMatrix(Row, 5)) <> 0 Then Fg1_CellChanged Row, 5
    
    Set xRs = Nothing
    Exit Sub
Salir:
    Set xRs = Nothing
    Agregando = False
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Fg1.Col = 4 Or Fg1.Col = 7 Then
        Fg1.Editable = flexEDNone
    Else
        If Fg1.Col = Fg1.Cols - 1 Then
            Fg1.Editable = flexEDNone
        Else
            Fg1.Editable = flexEDKbdMouse
        End If
    End If
End Sub


Private Sub Fg1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
        cmd_item(0).SetFocus
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    Select Case Col
        Case 5, 6, 7
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
        REGISTRO_ADD
    End If
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        REGISTRO_DEL   'F4 = Eliminar Item
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
'        If QueHace = 3 Then
'            PopupMenu Menu4
'        Else
'            PopupMenu menu1
'        End If
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = True
    txtfecha(0).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
    Select Case KeyCode
'        Case vbKeyF '--FECHA
'            txtfecha(0).SetFocus
'        Case vbKeyM '--MONEDA
'            txt_cb(1).SetFocus
'        Case vbKey1 '--CLIENTE
'            txt_cb(0).SetFocus
'
'        Case vbKeyC '--COTIZACION
'            If CmdTipo(0).Enabled = True And CmdTipo(0).Value = True Then CmdTipo(0).SetFocus
'        Case vbKeyF '--FACTURA
'            If CmdTipo(1).Enabled = True And CmdTipo(1).Value = True Then CmdTipo(1).SetFocus
'        Case vbKeyB '--BOLETA
'            If CmdTipo(2).Enabled = True And CmdTipo(2).Value = True Then CmdTipo(2).SetFocus
'        Case vbKeyG '--GRABAR
'            If cmdtipo(3).Enabled = True Then cmdtipo(3).SetFocus
'        Case vbKeyS '--SALIR
'            If cmdtipo(4).Enabled = True Then cmdtipo(4).SetFocus
'        '--ITEM
'        Case vbKeyA '--AGREGAR ITEM
'           If cmd_item(0).Enabled = True Then cmd_item(0).SetFocus
'        Case vbKeyE '--ELIMINAR ITEM
'            If cmd_item(1).Enabled = True Then cmd_item(1).SetFocus
        Case vbKeyEscape
            '--SI ESTA ACTIVO LOS BOTONES DE OTROS
            If fra(9).Visible = True Then pActivarObjetos True '--BOTOMES OTROS
            If fra(10).Visible = True Then pActivarObjetos True '--SEGUNDA VENTANA
            If fra(12).Visible = True Then pActivarObjetos True '--MODIFICAR Nº DOC.
        Case vbKeyF7
            If opt_descuento(0).Enabled = False Then Exit Sub
            If opt_descuento(0).Value = True Then opt_descuento(0).SetFocus
            If opt_descuento(1).Value = True Then opt_descuento(1).SetFocus
            If opt_descuento(2).Value = True Then opt_descuento(2).SetFocus
        Case vbKeyF8
            If Fg1.Enabled = False Then Exit Sub
            If Fg1.Rows <= 1 Then Exit Sub
            Fg1.Row = 1:        Fg1.Col = 5:        Fg1.SetFocus
        Case vbKeyF9
            If cmd_item(0).Enabled = False Then Exit Sub
            cmd_item(0).SetFocus
    End Select
End Sub

Private Sub Form_Load()
    Me.Width = 11835
    Me.Height = 7485
    CentrarFrm Me
    '-------------------------------
    mIdEventoOtrosBotones = -1
    '-------------------------------
    SeEjecuto = False
    Agregando = False
    QueHace = 3
    '---------------------
    Fg1.FrozenCols = 4
    Fg1.Tag = Fg1.FormatString

    GRID_COMBOLIST Fg1, 3
    
    txtfecha(0).Valor = Date
    
    '--------------
    Fg1.ColFormat(5) = "##.00000"
    Fg1.ColFormat(6) = "###,###.00000"
    Fg1.ColFormat(7) = "###,###.00000"
    Fg1.ColFormat(8) = "###,###.00000"
    OCULTAR_COL Fg1, 1, 2
    OCULTAR_COL Fg1, 9, 14
    Me.Tag = "Punto de Venta"
    '-------
    fCargarImpuesto
    '-------
    pCargarUsuario
    '-------
    mMesActivo = Month(Date)
    '-------
End Sub


Sub Blanquea()
    LimpiaText txt_importe
    LimpiaText txt_cb, True
    Fg1.Rows = 1
    LimpiaText lbl_num, True
    LimpiaText txtfecha
    lbl_TipoCambio.Caption = ""
    pCargarAlmacen

End Sub

Function Grabar() As Boolean
    

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDiario As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim xCod As Integer
    Dim xCodDet As Integer '--al detalle
    Dim xCol, xFil As Integer
    Dim xCorr As Integer
    Dim nNumeroRegistro As String '--INDICA EL NUMERO DE REGISTRO
    
    Dim xNumAsiento As String

Dim xId  As Integer
'On Error GoTo LaCague

    xCon.BeginTrans
    If QueHace = 1 Then '--NUEVO REGISTRO
        If mIdDocumento = 0 Then '--COTIZACION
            RST_Busq RstCab, "SELECT top 1 * FROM pvt_cotizacion ", xCon
            RST_Busq RstDet, "SELECT top 1 * FROM pvt_cotizaciondet", xCon
            xCod = HallaCodigoTabla("pvt_cotizacion", xCon, "id")
        Else '--FACTURA, BOLETA DE VENTA
            RST_Busq RstCab, "SELECT top 1 * FROM vta_ventas ", xCon
            RST_Busq RstDet, "SELECT top 1 * FROM vta_ventasdet", xCon
            RST_Busq RstDiario, "SELECT top 1 * FROM con_diario", xCon
            xCod = HallaCodigoTabla("vta_ventas", xCon, "id")
            xNumAsiento = NuevoNumAsiento(2, mMesActivo, xCon)
        End If
        RstCab.AddNew
        RstCab("id") = xCod
        
    Else '--MODIFICAR REGISTRO
    
        xCod = CStr(lbl_codigo.Caption)
        If mIdDocumento = 0 Then '--COTIZACION
            'Eliminamos el detalle de la cotizacion
            xCon.Execute "DELETE * FROM pvt_cotizaciondet WHERE idcot = " & xCod & ""
                
            RST_Busq RstCab, "SELECT * FROM pvt_cotizacion WHERE id = " & xCod & "", xCon
            RST_Busq RstDet, "SELECT TOP 1 * FROM pvt_cotizaciondet", xCon
       
       Else '--FACTURA, BOLETA DE VENTA
       
            'Eliminamos el stock agregado con la venta
            Dim RstDeta2 As New ADODB.Recordset
            RST_Busq RstDeta2, "SELECT iditem, canpro  From vta_ventasdet WHERE idvta = " & xCod & " ;", xCon
            If RstDeta2.RecordCount <> 0 Then RstDeta2.MoveFirst
            Do While Not RstDeta2.EOF
                xCon.Execute "UPDATE alm_inventario SET stckact = stckact + " + CStr(NulosN(RstDeta2("canpro"))) + " WHERE id =" + CStr(RstDeta2("iditem")) + "; "
                RstDeta2.MoveNext
            Loop
            Set RstDeta2 = Nothing
            '------------DEL NUMERO DE REGISTRO
            RST_Busq RstDiario, "SELECT numasi FROM con_diario WHERE idmes = " & Format(CDate(txtfecha(0).Valor), "mm") & " AND " _
                & " idlib = 2 AND idmov = " & xCod & " And iddoc = " & mIdDocumento & "", xCon
            If RstDiario.RecordCount <> 0 Then
                xNumAsiento = RstDiario("numasi") & ""
            Else
                If QueHace = 1 Then
                    xNumAsiento = NuevoNumAsiento(2, mMesActivo, xCon)
                Else
                    xNumAsiento = DevuelveNumAsiento(2, NulosN(lbl_codigo.Caption), mMesActivo, xCon)
                End If
            End If
            Set RstDiario = Nothing
            
            'Eliminamos el detalle de la venta
            xCon.Execute "DELETE * FROM vta_ventasdet WHERE idvta = " & xCod & ""
            'Eliminamos el asiento contable
            xCon.Execute "DELETE * FROM con_diario WHERE idmes = " & Format(CDate(txtfecha(0).Valor), "mm") & " AND " _
                & " idlib = 2 AND idmov = " & xCod & " And iddoc = " & mIdDocumento & ""
                
            RST_Busq RstCab, "SELECT * FROM vta_ventas WHERE id = " & xCod & "", xCon
            RST_Busq RstDet, "SELECT TOP 1 * FROM vta_ventasdet", xCon
            RST_Busq RstDiario, "SELECT TOP 1 * FROM con_diario", xCon
            
        End If
    End If
    RstCab("idcli") = NulosN(lbl_cb_cod(0).Caption)
    
    RstCab("numser") = lbl_num(0).Caption
    RstCab("numdoc") = lbl_num(1).Caption
    RstCab("fchdoc") = CDate(txtfecha(0).Valor)
    RstCab("fchven") = CDate(txtfecha(0).Valor)
    RstCab("idmon") = NulosN(lbl_cb_cod(1).Caption)
    
    RstCab("idconpag") = "1" '--EFECTIVO
    
    RstCab("impbru") = NulosN(txt_importe(0).Text)      '--importe bruto del documento (Base imponible para operacion gravada)
    RstCab("impinaf") = NulosN(txt_importe(1).Text)     '--importe de operacion inafecta
    RstCab("impigv") = NulosN(txt_importe(2).Text)      '--importe igv del documento
    RstCab("impisc") = NulosN(txt_importe(3).Text)      '--importe del Impuesto Selectivo al Consumo
    RstCab("impotr") = 0                                '--importe otros tributos
    RstCab("imptotdoc") = NulosN(txt_importe(4).Text)   '--importe total del documento
    RstCab("idven") = NulosN(lbl_usuario(0))            '--Id del vendedor (ver tabla de vendedores)
    '--especifica el modo de descuento que se le esta efectuando (1 = general;  2 = corporativo)
    If opt_descuento(1).Value = True Then RstCab("moddes") = 1
    If opt_descuento(2).Value = True Then RstCab("moddes") = 2
    '--especifica el tipo de descuento que se le esta efectuando (1 = Porcentaje;  2 = valor)
    If opt_descuento1(0).Value = True Then RstCab("tipdes") = 1
    If opt_descuento1(1).Value = True Then RstCab("tipdes") = 2
    '-------------------------
    
    
    If mIdDocumento <> 0 Then '--SOLO FACTURA Y BOLETA
        nNumeroRegistro = Trim(Format(Str(mMesActivo), "00")) + xNumAsiento
        'RstCab ("idpunvencli")="" '--codigo del punto de venta del cliente (no se usa actualmente, pero esta para cuando se le facture de frente a los puntos de venta de los clientes)
        RstCab("tipdoc") = mIdDocumento '--tipo de documento (ver tabla mae_documento)
        RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
        RstCab("percepcion") = "0"
        RstCab("idnumper") = "0"
        If Contabilizar = True Then RstCab("numreg") = nNumeroRegistro
        RstCab("impven") = NulosN(lbl_TipoCambio.Caption)   '--Tipo de Cambio de Venta
        RstCab("idtipven") = "1"                           '--Tipo de venta (ver tabla mae_tipoventa) ==>>1::Ventas Grabadas
        
        If Trim(txt_cotiza(1).Text) <> "" Then
            '--especifica el origen del item (1 = directo; 2 = Guia de Remision;  3 = Cotizacion)
            RstCab("oriitem") = "3"
            '--ACTUALIZAR LA COTIZACION CON EL DOCUMENTO DE VENTA
            xCon.Execute "UPDATE pvt_cotizacion SET iddocven = " + CStr(xCod) + " WHERE id=" + CStr(NulosN(txt_cotiza(0).Text))
        Else
            RstCab("oriitem") = "1"
        End If
        RstCab("idalm") = NulosN(lbl_cb_cod(3).Caption)  '--almacen
        RstCab("tipgen") = "2" '--indica quien genera la venta1=oficina, 2 = punto de venta
        
    End If
    
    RstCab.Update
    
    For xFil = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(xFil, 1)) > 0 And NulosN(Fg1.TextMatrix(xFil, 2)) > 0 Then
            RstDet.AddNew
            '--LLAVE
            If mIdDocumento = 0 Then RstDet("idcot") = xCod
            If mIdDocumento <> 0 Then RstDet("idvta") = xCod
            '--FIN LLAVE
            RstDet("descripusu") = ""
            RstDet("iditem") = NulosN(Fg1.TextMatrix(xFil, 1))
            RstDet("idunimed") = NulosN(Fg1.TextMatrix(xFil, 2))
            RstDet("canpro") = NulosN(Fg1.TextMatrix(xFil, 5))
            RstDet("preuni") = NulosN(Fg1.TextMatrix(xFil, 14))
            RstDet("preunibru") = NulosN(Fg1.TextMatrix(xFil, 6))
            RstDet("imptot") = NulosN(Fg1.TextMatrix(xFil, 8))
            RstDet("revisado") = 0
            RstDet("valdes") = NulosN(Fg1.TextMatrix(xFil, 7))
            
            'ACTUALIZAMOS EL STOCK
            If mIdDocumento <> 0 Then
                xCon.Execute "UPDATE alm_inventario " _
                    + vbCr + " SET stckact = ( stckact - " & NulosN(Fg1.TextMatrix(xFil, 5)) & ") " _
                    + vbCr + " WHERE id =" & Val(Fg1.TextMatrix(xFil, 1)) & "; "
            End If
            '---------------------
            RstDet.Update
        End If
    Next xFil
    '-------------------------------------------------------------------------------------------------------------------------------
    '--DE LOS ASIENTOS
    If Contabilizar = True And mIdDocumento <> 0 Then
        '---------------------------------------
        'Grabamos el libro diario del movimiento
        '--ASIENTO DEL IMPORTE TOTAL
        pGenerarAsiento RstDiario, 2, xCod, xNumAsiento, NulosN(lbl_TipoCambio.Caption), CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra), CDate(txtfecha(0).Valor), mIdCuentaDoc, NulosN(lbl_cb_cod(1).Caption), NulosN(txt_importe(4).Text), True
        '-----------------------------------------------------
        'grabamos el impuesto si la operacion esta afecta a el
        If NulosN(txt_importe(2).Text) > 0 And mIdDocumento <> 0 Then
            '--ASIENTO DEL IMPUESTO
            pGenerarAsiento RstDiario, 2, xCod, xNumAsiento, NulosN(lbl_TipoCambio.Caption), CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra), CDate(txtfecha(0).Valor), mIdCuentaImpuesto, NulosN(lbl_cb_cod(1).Caption), NulosN(txt_importe(2).Text), False
        End If
       '********Rutina para que extraer la base imponible sea afecta o inafecta
        Dim xIdCuen As Double
        Dim xTotal As Double
        Dim rstdocus As New ADODB.Recordset
        
        RST_Busq RstTmp, "SELECT top 1 con_diario.idcue as cuenta, con_diario.impdebsol as importe FROM con_diario;", xCon
        DEFINIR_RST_TMP rstdocus, RstTmp
        Set RstTmp = Nothing
        For xFil = 1 To Fg1.Rows - 1
            xIdCuen = Trim(Fg1.TextMatrix(xFil, 12))
            xTotal = NulosN(Fg1.TextMatrix(xFil, 8))
            If rstdocus.RecordCount <> 0 Then rstdocus.MoveFirst
            rstdocus.Find ("cuenta ='" & xIdCuen & "'")
            If rstdocus.EOF = True Then
                rstdocus.AddNew
                rstdocus("cuenta") = xIdCuen
                rstdocus("importe") = xTotal
            Else
                rstdocus("importe") = rstdocus("importe") + xTotal
            End If
            rstdocus.Update
        Next xFil
        '------------------
        'Grabamos el diario
        If rstdocus.RecordCount > 0 Then rstdocus.MoveFirst
        
        Do While Not rstdocus.EOF
            pGenerarAsiento RstDiario, 2, xCod, xNumAsiento, NulosN(lbl_TipoCambio.Caption), CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra), CDate(txtfecha(0).Valor), NulosN(rstdocus("cuenta")), NulosN(lbl_cb_cod(1).Caption), NulosN(rstdocus("importe")), False
            rstdocus.MoveNext
        Loop
        
                
    End If
    Dim nMSG As String
    
    ''''''''
    If QueHace = 2 And mIdDocumento <> 0 Then
        xCon.Execute "UPDATE vta_ventas SET evento = 0 WHERE id = " + CStr(xCod) + ";"
    End If
    
    nMSG = "La " + nDocumento + " se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + IIf(mIdDocumento = 0, "", vbCr + "Núm. Registro: " + nNumeroRegistro)
    MsgBox nMSG, vbInformation, xTitulo
    
    xCon.CommitTrans
    Grabar = True
Salir:
    Set RstCab = Nothing:    Set RstDet = Nothing:
    Exit Function
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing:    Set RstDet = Nothing:
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function

Private Sub pGenerarAsiento(RstDiario As ADODB.Recordset, IDLibro, IDMov, NAsiento, mTipoCambio, FchAsiento, FchDoc, IDcuenta, IDMoneda, Importe, Optional EsDEBE As Boolean)
    RstDiario.AddNew
    RstDiario("año") = AnoTra
    RstDiario("idmes") = mMesActivo  'LLAVE - CODIGO DEL MES
    RstDiario("idlib") = IDLibro     'LLAVE - CODIGO DEL LIBRO
    RstDiario("idmov") = IDMov       'LLAVE - CODIGO DEL MOVIMIENTO
    RstDiario("numasi") = NAsiento   'LLAVE - NUMERO DE ASIENTO
    RstDiario("tc") = mTipoCambio
    RstDiario("fchasi") = FchAsiento
    RstDiario("fchdoc") = FchDoc
    RstDiario("idcue") = IDcuenta
    If EsDEBE = False Then
        If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("imphabsol") = Importe
            RstDiario("imphabdol") = 0
        Else
            RstDiario("imphabsol") = Importe * mTipoCambio
            RstDiario("imphabdol") = Importe
        End If
    Else
        If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("impdebsol") = Importe
            RstDiario("impdebdol") = 0
        Else
            RstDiario("impdebsol") = Importe * mTipoCambio
            RstDiario("impdebdol") = Importe
        End If
    End If

    RstDiario.Update
End Sub



Private Function fValidarDatos() As Boolean
    If txtfecha(0).Valor = "" Or IsDate(txtfecha(0).Valor) = False Then
        MsgBox "No ha especificado la Fecha", vbExclamation, xTitulo
        txtfecha(0).SetFocus
        Exit Function
    End If
    If Year(txtfecha(0).Valor) <> AnoTra Then
        MsgBox "La fecha ingresada es diferente al Año de trabajo" + vbCr + "Año de Trabajo: " + AnoTra + vbCr + "Cambie la fecha...", vbExclamation, xTitulo
        txtfecha(0).Valor = ""
        txtfecha(0).SetFocus
        Exit Function
    End If

    If mIdDocumento = 1 Then '--SOLO FACTURA
        If NulosN(Trim(lbl_cb_cod(0).Caption)) = 0 Then
           MsgBox "Seleccione el Cliente", vbExclamation, xTitulo
           txt_cb(0).SetFocus
           Exit Function
        End If
    End If
    If Trim(lbl_cb_cod(1).Caption) = "" Then
        MsgBox "Seleccione la Moneda", vbExclamation, xTitulo
        txt_cb(1).SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha ingresado los productos ", vbInformation, xTitulo
        cmd_item(0).SetFocus
        Exit Function
    End If

    '---------------------------------------------------------------------------
    '--VALIDAR EL INGRESO DE LOS DATOS
    Dim Q_ROW  As Long
    Dim Q_COL As Long '--COLUMNA A POSICIONAR SI FALTAN DATOS
    Q_COL = -1
    For Q_ROW = 1 To Fg1.Rows - 1
        If IsNumeric(Fg1.TextMatrix(Q_ROW, 5)) = False Or Fg1.TextMatrix(Q_ROW, 5) = "0" Then
            MsgBox "Ingrese la Cantidad:" + vbCr + _
            "Item:  " + Fg1.TextMatrix(Q_ROW, 3) & "", vbExclamation, xTitulo

            Q_COL = 5:       Exit For
        ElseIf IsNumeric(Fg1.TextMatrix(Q_ROW, 6)) = False Or Fg1.TextMatrix(Q_ROW, 6) = "0" Then
            MsgBox "Ingrese el Precio Unitario:" + vbCr + _
            "Producto:  " + Fg1.TextMatrix(Q_ROW, 3) & "", vbExclamation, xTitulo

            Q_COL = 6:       Exit For
        ElseIf NulosN(Fg1.TextMatrix(Q_ROW, 12)) = 0 Then
            MsgBox "No se le ha asignado una cuenta contable para venta al item : " + Fg1.TextMatrix(Q_ROW, 3) + vbCr _
                + "Asignele una cuenta en el menú Almacén opción Mantenimiento Items de Compra y Venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo

            Q_COL = 3:       Exit For
        
        End If
    Next Q_ROW
    If Q_COL <> -1 Then
        Agregando = True:  Fg1.Row = Q_ROW: Fg1.Col = Q_COL: Agregando = False
        Fg1.SetFocus
        Exit Function
    End If
    '---------------------------------------------------------------------------
    Dim RstTmp As New ADODB.Recordset
    If QueHace = 1 Then
        If mIdDocumento <> 0 Then
            RST_Busq RstTmp, "SELECT * FROM vta_ventas WHERE tipdoc =" + CStr(mIdDocumento) + " AND numser ='" + lbl_num(0).Caption + "' AND numdoc = '" + lbl_num(1).Caption + "' ", xCon
            If RstTmp.RecordCount > 0 Then
                MsgBox "El Nro de documento ha sido registrado por otro usuario se grabará con otro número", vbInformation, xTitulo
                pCargarNumeroDoc
            End If
        Else
            RST_Busq RstTmp, "SELECT * FROM pvt_cotizacion WHERE numdoc = '" + lbl_num(1).Caption + "' ", xCon
            If RstTmp.RecordCount > 0 Or NulosN(lbl_num(1).Caption) = 0 Then
                MsgBox "El Nro de la Cotización ha sido registrado por otro usuario se grabará con otro número", vbInformation, xTitulo
                pCargarNumeroDoc
            End If
        End If
    End If
    
    '--VALIDAR QUE UNA COTIZACION SOLO SE REFERENCIE A UN DOCUMENTO
    If mIdDocumento <> 0 And NulosN(txt_cotiza(0).Text) <> 0 Then '--
        RST_Busq RstTmp, "SELECT id FROM pvt_cotizacion WHERE (iddocven IS NULL AND iddocven<>0 ) and id = " + CStr(NulosN(txt_cotiza(0).Text)), xCon
        If RstTmp.RecordCount > 0 Then
            MsgBox "La Cotización ya fue usada por una venta" + vbCr + "Ingrese otra cotización", vbInformation, xTitulo
            Blanquea
            LimpiaText txt_cotiza, True
            txt_cotiza(1).SetFocus
            Set RstTmp = Nothing
            Exit Function
        End If
    End If
    Set RstTmp = Nothing
    ''---------
    If mIdDocumento <> 0 Then
        If mIdCuentaDoc = 0 Then
            MsgBox "No se ha asignado una cuenta contable al documento " + nDocumento & Chr(13) _
                & "Asignele una cuenta en el menu Contabilidad opcion Asignar Ctas. Contables a documentos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
        If mIdCuentaImpuesto = 0 Then
            MsgBox "El impuesto asignado al documento " + nDocumento & Chr(13) & " no tiene cuenta contable" & Chr(13) _
                & "Asignele una cuenta en el menu Contabilidad opcion Maestro de Impuestos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    End If
    
    fValidarDatos = True
End Function

Private Sub fra_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Index <> 2 Then Exit Sub
    '--CARGAR LA VENTANA PARA CAMBIAR DE NUMERO DE DOCUMENTO
    lbl_encabezado_DblClick 0
End Sub

Private Sub lbl_encabezado_DblClick(Index As Integer)
    
    fra(12).Visible = True
    fra(12).Left = 1740
    fra(12).Top = 1305
    pActivarObjetos False
    CmdDocNum(0).SetFocus
    
End Sub

Private Sub lbl_num_DblClick(Index As Integer)
    '--CARGAR LA VENTANA PARA CAMBIAR DE NUMERO DE DOCUMENTO
    lbl_encabezado_DblClick 0
End Sub

Private Sub opt_descuento_Click(Index As Integer)
    Select Case Index
        Case 0: '--NINGUNO
            opt_descuento1(0).Value = False
            opt_descuento1(1).Value = False
            pAplicarDescuento
        Case 1 '--GENERAL
            If opt_descuento1(0).Value = False And opt_descuento1(1).Value = False Then
                opt_descuento1(0).Value = True
            Else
                pAplicarDescuento
            End If
        Case 2 '--CORPORATIVO
            If opt_descuento1(0).Value = False And opt_descuento1(1).Value = False Then
                opt_descuento1(0).Value = True
            Else
                pAplicarDescuento
            End If
    End Select
End Sub


Private Sub opt_descuento_KeyPress(Index As Integer, KeyAscii As Integer)
    If opt_descuento1(0).Value = True Then opt_descuento1(0).SetFocus
    If opt_descuento1(1).Value = True Then opt_descuento1(1).SetFocus
End Sub

Private Sub opt_descuento1_Click(Index As Integer)
    pAplicarDescuento
End Sub

Private Sub opt_descuento1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    If Fg1.Rows > 1 Then
        Fg1.Row = 1
        Fg1.Col = 5
        Fg1.SetFocus
    Else
        cmd_item(0).SetFocus
    End If
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        lbl_cb(Index).Caption = ""
        lbl_cb_cod(Index).Caption = ""
        If Index = 0 Then '--CLIENTE
            opt_descuento(0).Value = True
            opt_descuento(2).Enabled = False
        End If
        If Index = 4 Then txt_cb(5).Text = "" '--ALMACEN
        If Index = 5 Then txt_cb(2).Text = "" '--SERIE DEL DOCUMENTO
        If Index = 2 Then CmdDoc(0).Enabled = False
    Else
        If Index = 0 Then '--CLIENTE
            opt_descuento(0).Value = True
            opt_descuento(2).Enabled = True
        End If
        If mIdEventoOtrosBotones = 4 And Index = 2 Then '--SI SELECCIONA EMITIR DOC. ANULADOS => SALIR
            CmdDoc(0).Enabled = True
        End If
    End If

End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'On Error GoTo error
    Dim mIdDoc As Integer

    If mIdEventoOtrosBotones = 4 And Index = 2 And KeyCode = 13 Then
        txt_cb(2).Text = Format(txt_cb(2).Text, "0000000000")
        Exit Sub '--SI SELECCIONA EMITIR DOC. ANULADOS => SALIR
    End If
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
    If KeyCode <> 13 Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub
    
    
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    Select Case Index
        Case 0 '--CLIENTE
            nSQL = "SELECT  mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id " _
                + vbCr + " FROM mae_cliente " _
                + vbCr + " WHERE (((mae_cliente.numruc)='" + Trim(txt_cb(Index).Text) + "'));"
        Case 1 '--MONEDA
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion,mae_moneda.id as cod " _
            + vbCr + " From mae_moneda " _
            + vbCr + " WHERE (((mae_moneda.id)=" + CStr(Trim(txt_cb(Index).Text)) + "));"
        
                '--------------------------------------------
        Case 5 '--MUM SERIE
            If NulosN(lbl_cb_cod(4).Caption) = 0 Then
                MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
                txt_cb(4).SetFocus
                Exit Sub
            End If
            If cb_doc.ListIndex = -1 Then
                MsgBox "Seleccioe un Documento" + vbCr + "Luego Proceda", vbExclamation, xTitulo
                cb_doc.SetFocus
                Exit Sub
            End If

            If UCase(cb_doc.Text) = "FACTURA" Then mIdDoc = 1 '--FACTURA
            If UCase(cb_doc.Text) = "BOLETA DE VENTA" Then mIdDoc = 3 '--BOLETA DE VENTA
            nSQL = "SELECT Format([alm_numseries].[numser],'0000') AS nombre, alm_numseries.id, Format([alm_numseries].[numser],'0000') AS cod  " _
            + vbCr + " FROM alm_numseries " _
            + vbCr + " WHERE Format(alm_numseries.numser,'0000') ='" + Format(NulosN(txt_cb(Index).Text), "0000") + "' AND alm_numseries.idtipdoc=" + CStr(mIdDoc) + " AND alm_numseries.idalm=" + CStr(NulosN(lbl_cb_cod(4).Caption)) + " ;"
            
        Case 2 '--NUMERO DOCUMENTO
        
            If NulosN(lbl_cb_cod(4).Caption) = 0 Then
                MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
                txt_cb(4).SetFocus
                Exit Sub
            End If
            If cb_doc.ListIndex = -1 Then
                MsgBox "Seleccioe un Documento" + vbCr + "Luego Proceda", vbExclamation, xTitulo
                cb_doc.SetFocus
                Exit Sub
            End If
            
            Dim nSQLEvento As String
            If UCase(cb_doc.Text) <> "COTIZACIÓN" Then
                If mIdEventoOtrosBotones <> 4 Then
                    If UCase(cb_doc.Text) = "FACTURA" Then mIdDoc = 1 '--FACTURA
                    If UCase(cb_doc.Text) = "BOLETA DE VENTA" Then mIdDoc = 3 '--BOLETA DE VENTA
                    If mIdEventoOtrosBotones <= 3 Then '--SOLO MODIFICAR(1) , ELIMINAR(2), ANULAR(3)
                        nSQLEvento = " AND vta_ventas.evento=" + CStr(mIdEventoOtrosBotones) + " "
                    End If
                    nSQL = "SELECT  vta_ventas.numdoc as nombre,vta_ventas.id,vta_ventas.id as cod, vta_ventas.fchdoc " _
                        + vbCr + " FROM vta_ventas " _
                        + vbCr + " WHERE FORMAT(vta_ventas.numdoc,'0000000000') = '" + Format(NulosN(txt_cb(Index).Text), "0000000000") + "' AND  vta_ventas.numser ='" + lbl_cb_cod(5).Caption + "'  AND vta_ventas.idalm=" + CStr(NulosN(lbl_cb_cod(4).Caption)) + " AND vta_ventas.tipdoc=" + CStr(mIdDoc) + " AND vta_ventas.anulado=0 " _
                               + nSQLEvento
                Else
                   CmdDoc(0).Enabled = True
                    Exit Sub
                End If
            Else
                mIdDoc = 3 '--COTIZACION
                nSQL = "SELECT pvt_cotizacion.numdoc as nombre ,pvt_cotizacion.id, pvt_cotizacion.id as cod,  pvt_cotizacion.fchdoc  " _
                    + vbCr + " FROM pvt_cotizacion " _
                    + vbCr + " WHERE FORMAT(pvt_cotizacion.numdoc,'0000000000')='" + Format(NulosN(txt_cb(Index).Text), "0000000000") + "' AND  (pvt_cotizacion.iddocven Is Null Or pvt_cotizacion.iddocven = 0) And pvt_cotizacion.anulado = 0 "
            End If
    End Select

    If xCon.State = 0 Then Exit Sub
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then Exit Sub
    If RstTmp.RecordCount > 0 Then
        txt_cb(Index) = RstTmp.Fields(0) & "" '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
        '--SI ESTA EN LA OPCION DE SELECIONAR NUM. DOC, ACTUALIZAR LA FECHA DE REGISTRO
        If Index = 2 Then '--
            txtfecha(1).Valor = CDate(RstTmp.Fields("fchdoc"))
            CmdDoc(0).Enabled = True
        End If
    Else
        txt_cb(Index) = ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = ""  '--NOMBRE
        lbl_cb_cod(Index).Caption = ""  '--CODIGO
        '--SI ESTA EN LA OPCION DE SELECIONAR NUM. DOC, ACTUALIZAR LA FECHA DE REGISTRO
        If Index = 2 Then '--
            CmdDoc(0).Enabled = False
        End If
    End If
    Set RstTmp = Nothing
    '--SI SELECCIONA UNA MONEDA CARGAR LAS CUENTAS CONTABLES
    If Index = 1 And txt_cb(Index) <> "" Then pCargarCuentaContable mIdCuentaDoc, mIdCuentaImpuesto, mIdDocumento, NulosN(lbl_cb_cod(3).Caption), NulosN(CStr(lbl_cb_cod(1).Caption))
    '---------------------------------------------------------------------
'    If lbl_cb_cod(index).Caption <> "" Then SendKeys vbTab
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt_cb(Index).Enabled = False Then Exit Sub
        Select Case Index
            Case 0 '--CLIENTE
                If Trim(txt_cb(Index).Text) <> "" Then txt_cb(1).SetFocus
            Case 1 '--MONEDA
                If Trim(txt_cb(Index).Text) <> "" Then cmd_item(0).SetFocus
            Case 2 '--Nº DOCUMENTO 2DA VENTANA
                If Trim(txt_cb(Index).Text) <> "" And CmdDoc(0).Enabled = True Then CmdDoc(0).SetFocus
            Case 3 '--ALMACEN
                If Trim(txt_cb(Index).Text) <> "" Then txt_cb(0).SetFocus
            Case 4 '--ALMACEN 2DA VENTANA
                If Trim(txt_cb(Index).Text) <> "" Then cb_doc.SetFocus
            Case 5 '--Nº SERIE 2DA VENTANA
                If Trim(txt_cb(Index).Text) <> "" Then txt_cb(2).SetFocus
        End Select
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
        If KeyAscii = 46 Then KeyAscii = 0
    End If
End Sub

Private Function HallaValor(conn As ADODB.Connection, tabla As String, campo As String) As Long
Dim xRs As New ADODB.Recordset
On Error GoTo error
RST_Busq xRs, "SELECT top 1 CLng([" + campo + "]) AS num FROM " + tabla + " ORDER BY CLng([" + campo + "]) DESC;", conn
If xRs.State = 1 Then
    If xRs.EOF = False And xRs.BOF = False And xRs.RecordCount <> 0 Then
        HallaValor = NulosN(xRs.Fields(0)) + 1
    End If
Else
    HallaValor = -1
End If
Set xRs = Nothing
Exit Function
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "HallarValor"
End Function
'
'
'----------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
'
'
Private Sub CmdTipo_Click(Index As Integer)
    If Index = 4 Then '--SALIR
        Unload Me
        Exit Sub
    End If
    Select Case Index
        Case 0 '--COTIZACION
            nDocumento = "Cotización"
        Case 1 '--FACTURA
            nDocumento = "Factura"
        Case 2 '--BOLETA
            nDocumento = "Boleta de Venta"
    End Select

    '--LIMPANDO EK FORMULARIO
    If Index = 0 Or Index = 1 Or Index = 2 Then
        If SeEjecuto = False Or QueHace = 2 Then GoTo ir_limpia '--SOLO SE EJECITA AL INICIAR EL FORMULARIO, PARA CARGAR POR DEFECTO EL DOCUMENTO COTIZACION
'        If MsgBox("Desea crear nueva " + nDocumento, vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbYes Then
        Dim vRpt As Integer
        If Fg1.Rows - 1 > 1 Or IsDate(txtfecha(0).Valor) = True Or NulosN(lbl_cb_cod(0).Caption) <> 0 Then
            vRpt = MsgBox("Desea Conservar los valores ingresados " + nDocumento, vbYesNoCancel + vbDefaultButton1 + vbQuestion, xTitulo)
        Else
            vRpt = 6
        End If
        If vRpt = 6 Then '--si

        ElseIf vRpt = 7 Then '--no
ir_limpia:
            Blanquea
        ElseIf vRpt = 2 Then '--cancelar
            Exit Sub
        End If
        Ocultar picTipo, True
        Ocultar CmdTipo, True
        LimpiaText txt_cotiza
        fra(1).Visible = False '--numero de cotizacion
    End If
    
    Select Case Index
        Case 0 '--COTIZACION
            QueHace = 1
            picTipo(0).Visible = True
            CmdTipo(0).Visible = False
            mIdDocumento = 0
        Case 1 '--FACTURA
            picTipo(1).Visible = True
            CmdTipo(1).Visible = False
            fra(1).Visible = True   '--numero de cotizacion
            mIdDocumento = 1
            QueHace = 1
        Case 2 '--BOLETA
            picTipo(2).Visible = True
            CmdTipo(2).Visible = False
            fra(1).Visible = True   '--numero de cotizacion
            mIdDocumento = 3
            QueHace = 1

        Case 3 '--GRABAR
            If fValidarDatos() = False Then Exit Sub
            If MsgBox("¿Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " la " + nDocumento + "?", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
        
            If Grabar() = True Then
                If MsgBox("¿Desea Imprimir la " + nDocumento + "?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbYes Then
                    Imprimir mIdDocumento
                End If
                Blanquea
                LimpiaText txt_cotiza, True
                pCargarNumeroDoc
            Else
                Exit Sub
            End If
            If QueHace = 2 Then CmdTipo_Click 6 '--CANCELAR
            
        Case 4 '--SALIR
            
        Case 5 '--OTROS
            If CmdTipo(5).Caption = "&Otros" Then
                fra(9).Visible = True
                fra(9).Left = 8370
                fra(9).Top = 3570
                pActivarObjetos False
                CmdOtros(0).SetFocus
                
                Exit Sub
            Else
                CmdTipo(5).Caption = "&Otros"
                pCargarUsuario
            End If
    End Select

    '--CARGAR NUMERO DE DOCUMENTO, VERIFICAR SI EL USUARIO PUEDE IMPRIMIR DOCUMENTOS
    If Index = 0 Or Index = 1 Or Index = 2 Then pCargarNumeroDoc
    '-----
    lbl_encabezado(0).Caption = nDocumento
    Me.Caption = Me.Tag + " - " + nDocumento
    Select Case Index
        Case 0 '--COTIZACION
            If SeEjecuto = True Then txtfecha(0).SetFocus
        Case 1 '--FACTURA
            If SeEjecuto = True Then txt_cotiza(1).SetFocus
        Case 2 '--BOLETA
            If SeEjecuto = True Then txt_cotiza(1).SetFocus
    End Select
    
End Sub


Private Sub fCargarImpuesto()
    On Error GoTo error
    Dim RstTmp As New ADODB.Recordset
    Dim nImpuestoValor As String
    RST_Busq RstTmp, "SELECT mae_impuestos.id, mae_impuestos.tasa, mae_impuestos.Abrev From mae_impuestos WHERE (((mae_impuestos.id)=1)); ", xCon
    If RstTmp.BOF = False Or RstTmp.BOF = False Or RstTmp.RecordCount <> 0 Then
        mImpuestoValor = IIf(NulosN(RstTmp.Fields("tasa")) > 1, (NulosN(RstTmp.Fields("tasa"))) / 100, NulosN(RstTmp.Fields("tasa")))
        nImpuestoValor = FormatPercent(mImpuestoValor, 0)
        lbl_importe(2).Caption = RstTmp.Fields("abrev") & " (" + nImpuestoValor + ")"
    Else
        MsgBox "No hay impuesto" + vbCr + "Verifique el impuesto I.G.V.", vbInformation, xTitulo
    End If
    Set RstTmp = Nothing
    '--------
    
    '--------
    Exit Sub
error:
    SHOW_ERROR Me.Name, "fCargarImpuesto"
End Sub

Private Sub pCalculosTotales()
    Dim mTotal, mAfecto, mInafecto, mImpuesto, mImpuestoISC As Double
    '------------------------------------------------
    LimpiaText txt_importe '--LIMPIAR OBJETO
    mAfecto = GRID_SUMAR_COL(Fg1, 10)
    mInafecto = GRID_SUMAR_COL(Fg1, 11)
    mImpuesto = mAfecto * mImpuestoValor
    mTotal = mAfecto + mInafecto + mImpuesto
    mImpuestoISC = 0
    txt_importe(0).Text = Format(mAfecto, FORMAT_MONTO)
    txt_importe(1).Text = Format(mInafecto, FORMAT_MONTO)
    txt_importe(2).Text = Format(mImpuesto, FORMAT_MONTO)
    txt_importe(3).Text = Format(mImpuestoISC, FORMAT_MONTO)
    txt_importe(4).Text = Format(mTotal, FORMAT_MONTO)
End Sub

Private Function fHallarNumeroDoc(tabla As String, campo As String) As Long
'    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    RST_Busq xRs, "SELECT top 1 CLng([" + campo + "]) AS num FROM " + tabla + " ORDER BY CLng(TRIM([" + campo + "])) DESC;", xCon
    If xRs.State = 1 Then
        If xRs.EOF = False And xRs.BOF = False And xRs.RecordCount <> 0 Then
            fHallarNumeroDoc = NulosN(xRs.Fields(0)) + 1
        Else
            fHallarNumeroDoc = 1
        End If
    Else
        fHallarNumeroDoc = -1
    End If
    Set xRs = Nothing
    Exit Function
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "fHallarNumeroDoc"
End Function

Private Sub txt_cotiza_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If Index = 0 Then Exit Sub
    If Trim(txt_cotiza(1).Text) = "" Then
        LimpiaText txt_cotiza, True
        Blanquea
    End If
End Sub

Private Sub txt_cotiza_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    If Index <> 1 Then Exit Sub
    
    If KeyCode = vbKeyF5 Then
        '------CARGAR LAS COTIZACIONES
        ReDim xCampos(5, 4) As String
        
        xCampos(0, 0) = "Número":       xCampos(0, 1) = "numero":       xCampos(0, 2) = "600":      xCampos(0, 3) = "N":
        xCampos(1, 0) = "Fecha":        xCampos(1, 1) = "fecha":        xCampos(1, 2) = "1000":     xCampos(1, 3) = "F":
        xCampos(2, 0) = "Cliente":      xCampos(2, 1) = "clidesc":      xCampos(2, 2) = "4500":     xCampos(2, 3) = "C":
        xCampos(3, 0) = "M":            xCampos(3, 1) = "simbolo":      xCampos(3, 2) = "450":      xCampos(3, 3) = "C":
        xCampos(4, 0) = "Importe":      xCampos(4, 1) = "imptotdoc":    xCampos(4, 2) = "1000":     xCampos(4, 3) = "N":
    
        nSQL = "SELECT pvt_cotizacion.id, pvt_cotizacion.numdoc, CDbl([numdoc]) AS numero, format(pvt_cotizacion.fchdoc,'dd/mm/yy') as fecha, mae_cliente.nombre AS clidesc, mae_moneda.simbolo, pvt_cotizacion.imptotdoc " _
            + vbCr + " FROM (pvt_cotizacion LEFT JOIN mae_moneda ON pvt_cotizacion.idmon = mae_moneda.id) LEFT JOIN mae_cliente ON pvt_cotizacion.idcli = mae_cliente.id " _
            + vbCr + " WHERE (pvt_cotizacion.iddocven Is Null Or pvt_cotizacion.iddocven = 0) ;"
        
        CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), "Buscando Cotización", "numero", "numero", Principio
        LimpiaText txt_cotiza, True
        If RstTmp.State = 0 Then GoTo Salir
        If RstTmp.EOF = True And RstTmp.BOF = True And RstTmp.RecordCount = 0 Then
            LimpiaText txt_cotiza, True
            GoTo Salir
        Else
            txt_cotiza(0).Text = RstTmp.Fields("id") & ""
            txt_cotiza(1).Text = RstTmp.Fields("numdoc") & ""
            GoTo Continuar: '--LLENAR LOS DEMAS DATOS
        End If
    End If
    
    If KeyCode <> 13 Then Exit Sub
    If Trim(txt_cotiza(1).Text) = "" Then Exit Sub

Continuar:
    '--LIMPIAR LOS DATOS
    Blanquea
    Set RstTmp = Nothing
    nSQL = "SELECT pvt_cotizacion.id, pvt_cotizacion.numdoc, pvt_cotizacion.fchdoc, pvt_cotizacion.idcli AS cliid, mae_cliente.numruc AS clinum, mae_cliente.nombre AS clidesc, pvt_cotizacion.idmon AS monid, mae_moneda.descripcion AS mondesc,pvt_cotizacion.moddes,pvt_cotizacion.tipdes " _
        + vbCr + " FROM (pvt_cotizacion LEFT JOIN mae_moneda ON pvt_cotizacion.idmon = mae_moneda.id) LEFT JOIN mae_cliente ON pvt_cotizacion.idcli = mae_cliente.id " _
        + vbCr + " WHERE  CDbl(pvt_cotizacion.numdoc)=" + CStr(NulosN(txt_cotiza(1).Text)) + " AND (pvt_cotizacion.iddocven Is Null Or pvt_cotizacion.iddocven = 0); "
        
    If xCon.State = 0 Then GoTo Salir
    
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then
        LimpiaText txt_cotiza, True
        GoTo Salir
    End If
    If RstTmp.BOF = True Or RstTmp.EOF = True Or RstTmp.RecordCount = 0 Then
        LimpiaText txt_cotiza, True
        GoTo Salir
    End If
    
    pPonerDatosEncabezado RstTmp '--CARGAR LOS DATOS EN EL ENCABEZADO
    
    Set RstTmp = Nothing
    '------DEL DETALLE
    nSQL = "SELECT pvt_cotizaciondet.iditem, pvt_cotizaciondet.idunimed, alm_inventario.descripcion, mae_unidades.abrev, pvt_cotizaciondet.canpro, pvt_cotizaciondet.preuni, pvt_cotizaciondet.valdes, pvt_cotizaciondet.imptot, pvt_cotizaciondet.preunibru, alm_inventario.idtipven, alm_inventario.idcuentaven " _
        + vbCr + " FROM mae_unidades RIGHT JOIN (pvt_cotizaciondet LEFT JOIN alm_inventario ON pvt_cotizaciondet.iditem = alm_inventario.id) ON mae_unidades.id = alm_inventario.idunimed " _
        + vbCr + " WHERE (((pvt_cotizaciondet.idcot)=" + CStr(txt_cotiza(0).Text) + "));"

    RST_Busq RstTmp, nSQL, xCon
    
    pPonerDatosDetalle RstTmp '--CARGAR LOS DATOS EN EL DETALLE
    
   
Salir:
    Set RstTmp = Nothing
    '------------------------------------------------
    If NulosN(txt_cotiza(0).Text) = 0 Then
        txt_cotiza(1).SetFocus
    End If
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown(" + CStr(Index) + ")"
End Sub


Private Sub pPonerDatosEncabezado(RstTmp As ADODB.Recordset, Optional ES_COTIZACION As Boolean = True)
    If RstTmp.State = 0 Then Exit Sub
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then Exit Sub
    txtfecha(0).Valor = RstTmp.Fields("fchdoc")
    '--DEL CLIENTE
    txt_cb(0).Text = RstTmp.Fields("clinum") & ""
    lbl_cb(0).Caption = RstTmp.Fields("clidesc") & ""
    lbl_cb_cod(0).Caption = RstTmp.Fields("cliid") & ""
    '--DE LA MONEDA
    txt_cb(1).Text = RstTmp.Fields("monid") & ""
    lbl_cb(1).Caption = RstTmp.Fields("mondesc") & ""
    lbl_cb_cod(1).Caption = RstTmp.Fields("monid") & ""
    '--MOSTRAR EL TIPO DE CAMBIO EN FUNCION DE LA FECHA DEL DOC.
    txtfecha_LostFocus 0
    '----
    If ES_COTIZACION = True Then
        txt_cotiza(0).Text = RstTmp.Fields("id") & ""
        txt_cotiza(1).Text = RstTmp.Fields("numdoc") & ""
    Else
        '--COLOCANDO LA SERIE Y EL NUMERO
        lbl_num(0).Caption = Format(RstTmp("numser") & "", "0000")
        lbl_num(1).Caption = RstTmp.Fields("numdoc")
        lbl_encabezado(1).Caption = lbl_num(0).Caption + " " + lbl_num(1).Caption
    End If
    If NulosN(RstTmp.Fields("moddes")) <> 0 Then
        opt_descuento(NulosN(RstTmp.Fields("moddes"))).Value = True
    End If
    If NulosN(RstTmp.Fields("tipdes")) <> 0 Then
        opt_descuento1(NulosN(RstTmp.Fields("tipdes")) - 1).Value = True
    End If

    
    
End Sub

Private Sub pPonerDatosDetalle(RstTmp As ADODB.Recordset)
    If RstTmp.State = 0 Then GoTo Salir
    If RstTmp.BOF = True Or RstTmp.EOF = True Or RstTmp.RecordCount = 0 Then GoTo Salir
    Agregando = True
    Dim A&, xFila&
    With Fg1
        .Rows = 1
        RstTmp.MoveFirst
        For A = 1 To RstTmp.RecordCount
            xFila = .Rows
            .Rows = .Rows + 1
            .TextMatrix(xFila, 1) = RstTmp.Fields("iditem") & ""
            .TextMatrix(xFila, 2) = RstTmp.Fields("idunimed") & ""
            .TextMatrix(xFila, 3) = RstTmp.Fields("descripcion") & ""
            .TextMatrix(xFila, 4) = RstTmp.Fields("abrev") & ""
            .TextMatrix(xFila, 5) = RstTmp.Fields("canpro") & ""
            .TextMatrix(xFila, 6) = NulosN(RstTmp.Fields("preuni"))
            
            .TextMatrix(xFila, 7) = NulosN(RstTmp.Fields("valdes"))
            .TextMatrix(xFila, 8) = NulosN(RstTmp.Fields("imptot"))

            .TextMatrix(xFila, 9) = RstTmp.Fields("idtipven") & ""
            '--ES GRAVADA O NO
            If NulosN(.TextMatrix(xFila, 9)) <> 0 And NulosN(.TextMatrix(xFila, 9)) <> 3 Then
                .TextMatrix(xFila, 10) = NulosN(.TextMatrix(xFila, 8))
                .TextMatrix(xFila, 11) = 0
            Else
                .TextMatrix(xFila, 10) = 0
                .TextMatrix(xFila, 11) = NulosN(.TextMatrix(xFila, 8))
            End If
            '--NUEVO PRECIO UNITARIO
            If NulosN(Fg1.TextMatrix(xFila, 8)) = 0 Or NulosN(.TextMatrix(xFila, 5)) = 0 Then
                .TextMatrix(xFila, 14) = 0
            Else
                .TextMatrix(xFila, 14) = NulosN(.TextMatrix(xFila, 8)) / NulosN(.TextMatrix(xFila, 5))
            End If
            .TextMatrix(xFila, 12) = RstTmp.Fields("idcuentaven") & ""
            '----------
            RstTmp.MoveNext
            If RstTmp.EOF = True Then Exit For
        Next A
    End With
    '--
    pCargarCuentaContable mIdCuentaDoc, mIdCuentaImpuesto, mIdDocumento, mIdAlmacen, NulosN(CStr(lbl_cb_cod(1).Caption))
    '--calcular los totales
    pCalculosTotales
   
Salir:
    Set RstTmp = Nothing
    Agregando = False
End Sub



Private Sub txt_cotiza_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txt_cotiza(1).Text) <> "" Then
            If Fg1.FixedRows <> Fg1.Rows - 1 Then
                Fg1.Row = 1
                Fg1.Col = 5
                Fg1.SetFocus
            End If
        Else
            SendKeys vbTab
        End If
        Exit Sub
    End If
    Select Case Index
        Case 1
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)

    If IsDate(txtfecha(0).Valor) = True Then
        Dim RstTmp As New ADODB.Recordset
        RST_Busq RstTmp, "SELECT con_tc.fecha, con_tc.impcom, con_tc.impven From con_tc WHERE (((con_tc.fecha)=CDate('" + txtfecha(0).Valor + "')) AND ((con_tc.idmon)=2));", xCon
        If RstTmp.BOF = False Or RstTmp.BOF = False Or RstTmp.RecordCount <> 0 Then
            nTipoCambio = NulosN(RstTmp.Fields("impven"))
        Else
            nTipoCambio = 0#
        End If
        lbl_TipoCambio.Caption = nTipoCambio
        Set RstTmp = Nothing
    End If
End Sub

Private Sub pCargarUsuario()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nCargo As String
    nSQL = "SELECT pvt_emp.id,  [pla_empleados].[nom] & ' ' & [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] AS nombre, pvt_emp.codigo ,pvt_emp.ven, pvt_emp.caj, pvt_emp.sup " _
        + vbCr + " FROM (pla_empleados INNER JOIN pvt_emp ON pla_empleados.id = pvt_emp.idemp) INNER JOIN mae_usuarios ON pla_empleados.id = mae_usuarios.idemp " _
        + vbCr + " WHERE pvt_emp.id= " + CStr(mIdEmpleado)
        
    nSQL = "SELECT pvt_emp.id, [pla_empleados].[nom] & ' ' & [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] AS nombre, " _
        & " pvt_emp.codigo, pvt_emp.ven, pvt_emp.caj, pvt_emp.sup" _
        & " FROM pla_empleados RIGHT JOIN pvt_emp ON pla_empleados.id = pvt_emp.idemp " _
        & " WHERE (((pvt_emp.id)=1))"

    
    LimpiaText lbl_usuario, True
    '--
    CmdTipo(0).Enabled = False '--COTIZACION
    CmdTipo(1).Enabled = False '--FACTURA
    CmdTipo(2).Enabled = False '--BOLETA
    '--
    RST_Busq RstTmp, nSQL, xCon
    With RstTmp
        If .State = 0 Then GoTo Salir
        If Abs(Val(NulosN(.Fields("ven")))) = "1" Then
            CmdTipo(0).Enabled = True
            nCargo = "Vendedor"
        End If
        If Abs(Val(NulosN(.Fields("caj")))) = "1" Then
            CmdTipo(1).Enabled = True
            CmdTipo(2).Enabled = True
            nCargo = "Cajero"
        End If
        If Abs(Val(NulosN(.Fields("sup")))) = "1" Then
            CmdTipo(1).Enabled = True
            CmdTipo(2).Enabled = True
            nCargo = "Supervisor"
        End If
        
        If Abs(Val(NulosN(.Fields("ven")))) = "1" Or Abs(Val(NulosN(.Fields("caj")))) = "1" Or Abs(Val(NulosN(.Fields("sup")))) = "1" Then
            lbl_usuario(0).Caption = .Fields(0) & "" '--ID PVT_EMP
            lbl_usuario(1).Caption = nCargo + " :  " + StrConv(.Fields(1) & "", 3) '--NOMBRE EMPLEADO
            lbl_usuario(2).Caption = "Código:  " + .Fields(2) & "" '--CODIGO
            lbl_usuario(1).ForeColor = &H400000
            '--ESTABLECIENDO EL NIVEL DE ACCESO
            '--HABILITAR LOS BOTONES POR DEFECTO
            Select Case UCase(nCargo)
                Case "VENDEDOR"
                    CmdTipo_Click 0
                    mNivelAcceso = 1
                    CmdOtros(2).Enabled = False '--ANULAR DOCUMENTOS
                    CmdOtros(3).Enabled = False '--EMITIR DOC ANULADOS
                    
                Case "CAJERO"
                    CmdTipo_Click 1
                    mNivelAcceso = 2
                                        
                Case "SUPERVISOR"
                    CmdTipo_Click 1
                    mNivelAcceso = 3
                                        
            End Select
        Else
            '--RESTRINGIR ACCESO (MUESTRA MENSAJE)
            fra_MsgAcceso.Top = 2610
            fra_MsgAcceso.Left = 2400
            fra_MsgAcceso.Visible = True
            '--BLOQUEAR OBJETOS SI NO ES USUARIO
            habilitar CmdTipo, False
            habilitar cb, False
            CmdTipo(4).Enabled = True
            Cmd_Cliente.Enabled = False
            txtfecha(0).Enabled = False
            habilitar_Locked txt_cb, True
            habilitar_Locked txt_cotiza, True
            fra(4).Enabled = False '--DESCUENTO
            habilitar cmd_item, False
        End If
    End With
Salir:
    Set RstTmp = Nothing
End Sub

Private Sub pCargarCuentaContable(IDCuentaDoc As Integer, IDCuentaImpuesto As Integer, IDDocumento As Integer, IDAlmacen As Integer, IDMoneda As Integer)
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    nSQL = "SELECT mae_documento.id, mae_documento.descripcion, mae_documento.abrev, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, con_planctas.descripcion AS ctadesc, mae_impuestos.idcuenvta AS cuentaimp, alm_numseries.numser, mae_documentocta.idcuen AS cuentadoc " _
        + vbCr + " FROM (alm_numseries RIGHT JOIN (mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) ON mae_documento.idimp = mae_impuestos.id) ON alm_numseries.idtipdoc = mae_documento.id) INNER JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc " _
        + vbCr + " WHERE (((mae_documento.id)=" + CStr(IDDocumento) + ") AND ((alm_numseries.idalm)=" + CStr(IDAlmacen) + ") AND ((mae_documentocta.idmon)= " + CStr(IDMoneda) + ") AND ((mae_documentocta.tipope)= -1 ));"
    
    RST_Busq RstTmp, nSQL, xCon
    
    IDCuentaDoc = 0
    IDCuentaImpuesto = 0
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        IDCuentaDoc = NulosN(RstTmp.Fields("cuentadoc"))
        IDCuentaImpuesto = NulosN(RstTmp.Fields("cuentaimp"))
    End If
    
End Sub


Private Sub pActivarObjetos(band As Boolean)
    '--SE ACTIVARA CUANDO SE EJECUTE EL BOTON OTROS
    fra(0).Enabled = band '--ENCABEZADO DE FORMULARIO
    fra(1).Enabled = band '--COTIZACION
    fra(4).Enabled = band '--DESCUENTO
    fra(3).Enabled = band '--BOTONES DE AGREGAR ITEM
    fra(5).Enabled = band '--MONTOS ACUMULADOS
    fra(6).Enabled = band '--BOTONES
    fra(8).Enabled = band '--MONEDA
    '--fra(9) '--BOTONES OTROS
    If band = True Then
        fra(9).Visible = False '--BOTONES OTROS
        fra(10).Visible = False '-- 2 FORM PARA SELECCIONAR EL DOC.
        fra(11).Visible = False '--IMPORTE TOTAL CONVERTIDO
        fra(12).Enabled = False '--MODIFICAR Nº DOCUMENTO
    End If
    
End Sub



Private Function Eliminar() As Boolean
    Dim Rpta As Integer
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    If NulosN(lbl_codigo.Caption) = 0 Then
        MsgBox "No hay documento para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If
    Dim mIdDoc As Integer
    Select Case UCase(cb_doc.Text)
        Case "COTIZACIÓN"
            mIdDoc = 0
            nSQL = "SELECT [vta_ventas].[numser] & ' ' & [vta_ventas].[numdoc] AS numerodoc, vta_ventas.fchdoc, mae_cliente.nombre AS cliente, vta_ventas.numreg " _
                + vbCr + " FROM pvt_cotizacion INNER JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON pvt_cotizacion.iddocven = vta_ventas.id " _
                + vbCr + " WHERE (((pvt_cotizacion.id)=" + CStr(NulosN(lbl_codigo.Caption)) + "));"
            RST_Busq RstTmp, nSQL, xCon
            If RstTmp.RecordCount <> 0 Then
                MsgBox "La Cotización que desea eliminar esta relacionado con una venta" + vbCr + "Nº Documento: " + RstTmp.Fields("numerodoc") & "" + vbCr + "Fecha Registro: " + RstTmp.Fields("fchdoc") & "" + vbCr + "Cliente: " + RstTmp.Fields("cliente") & "" + vbCr + "Num.Reg:" + RstTmp.Fields("numreg") & "", vbExclamation, xTitulo
                Set RstTmp = Nothing
                Exit Function
            End If
            
        Case "FACTURA"
            mIdDoc = 1
        Case "BOLETA DE VENTA"
            mIdDoc = 3
        Case Else '--CANCELAR
            MsgBox "Seleccione Otra vez el Documento", vbExclamation, xTitulo
            Exit Function
    End Select
    If MsgBox("¿ Esta seguro de eliminar el documento seleccionado ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then Exit Function
    xCon.BeginTrans
    If mIdDoc <> 0 Then
        'ACTUALIZAR EL CAMPO DE LA COTIZACION PAR QUE SE PUEDA GENERAR LA VENTA
        xCon.Execute "UPDATE pvt_cotizacion SET iddocven = 0 WHERE iddocven=" + CStr(NulosN(lbl_codigo.Caption)) + ";"
        'ELIMINAR EL STOCK
        Set RstTmp = Nothing
        RST_Busq RstTmp, "SELECT iditem, canpro  FROM vta_ventasdet WHERE idvta = " + CStr(NulosN(lbl_codigo.Caption)) + ";", xCon
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            xCon.Execute "UPDATE alm_inventario SET stckact = stckact + " + CStr(NulosN(RstTmp("canpro"))) + " WHERE id =" + CStr(RstTmp("iditem")) + "; "
            RstTmp.MoveNext
        Loop
        Set RstTmp = Nothing
        '----------------
        xCon.Execute "DELETE * FROM vta_ventasdet WHERE idvta = " + CStr(NulosN(lbl_codigo.Caption)) + ";"
        xCon.Execute "DELETE * FROM vta_ventas WHERE id = " + CStr(NulosN(lbl_codigo.Caption)) + ";"
        xCon.Execute "DELETE * FROM con_diario WHERE idlib = 2 AND idmov = " + CStr(NulosN(lbl_codigo.Caption)) + " AND Iddoc = " + CStr(mIdDoc) + ";"
    Else
    
        xCon.Execute "DELETE * FROM pvt_cotizaciondet WHERE idcot = " + CStr(NulosN(lbl_codigo.Caption)) + ";"
        xCon.Execute "DELETE * FROM pvt_cotizacion WHERE id = " + CStr(NulosN(lbl_codigo.Caption)) + ";"
        
    End If
    xCon.CommitTrans
    MsgBox "La " + cb_doc.Text + " se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Eliminar = True
    Exit Function
error:
    xCon.RollbackTrans
    SHOW_ERROR Me.Name, "Eliminar", True
End Function


Private Function Anular() As Boolean
    On Error GoTo error
    Dim RstTmp As New ADODB.Recordset
    Dim nDocumento As String
    If UCase(cb_doc.Text) = "COTIZACIÓN" Then
        nDocumento = txt_cb(2).Text
    Else
        nDocumento = txt_cb(5).Text + "-" + txt_cb(2).Text
    End If
    
    If MsgBox("¿Esta seguro de anular la" + cb_doc.Text + vbCr + "Nº: " + nDocumento + "?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Function
    
    xCon.BeginTrans
    xCon.Execute "UPDATE vta_ventas SET vta_ventas.Anulado = -1, " _
        & " vta_ventas.impbru = 0, vta_ventas.impinaf = 0, vta_ventas.impigv = 0,  vta_ventas.impisc = 0,  " _
        & " vta_ventas.impotr = 0, vta_ventas.imptotdoc = 0,  vta_ventas.impsal = 0  " _
        & " WHERE vta_ventas.id = " + CStr(NulosN(lbl_codigo.Caption)) + ";"
    
    'ACTUALIZAR EL CAMPO DE LA COTIZACION PAR QUE SE PUEDA GENERAR LA VENTA
    xCon.Execute "UPDATE pvt_cotizacion SET iddocven = 0 WHERE iddocven=" + CStr(NulosN(lbl_codigo.Caption)) + ";"
    'ELIMINAR EL STOCK
    Set RstTmp = Nothing
    RST_Busq RstTmp, "SELECT iditem, canpro  FROM vta_ventasdet WHERE idvta = " + CStr(NulosN(lbl_codigo.Caption)) + ";", xCon
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        xCon.Execute "UPDATE alm_inventario SET stckact = stckact + " + CStr(NulosN(RstTmp("canpro"))) + " WHERE id =" + CStr(RstTmp("iditem")) + "; "
        RstTmp.MoveNext
    Loop
    Set RstTmp = Nothing
    
    xCon.Execute "DELETE * FROM vta_ventasdet WHERE vta_ventasdet.idvta = " + CStr(NulosN(lbl_codigo.Caption)) + ";"
            
    'ponemos el diario a valor 0
    RST_Busq RstTmp, "SELECT * FROM con_diario WHERE idlib = 2 AND idmov = " + CStr(NulosN(lbl_codigo.Caption)) + ";", xCon
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
         RstTmp("impdebsol") = 0
         RstTmp("imphabsol") = 0
         RstTmp("impdebdol") = 0
         RstTmp("imphabdol") = 0
         RstTmp.MoveNext
     Loop
    Set RstTmp = Nothing
    
    MsgBox "la" + cb_doc.Text + " con Nº: " + nDocumento + " se anuló con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Anular = True
    xCon.CommitTrans
    Exit Function
error:
    Set RstTmp = Nothing
    xCon.RollbackTrans
    SHOW_ERROR Me.Name, "Anular"
End Function


Private Function EmitirAnulada() As Boolean
    Dim RstCab As New ADODB.Recordset
    Dim RstDiario As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim xId As Integer
    Dim xNumAsiento As String

    Dim IDCuentaDoc As Integer
    Dim IDCuentaImp As Integer
    Dim mIdDoc As Integer
    Dim mTipoCambio As Double
    Dim nSQL As String
    
    Select Case UCase(cb_doc.Text)
        Case "FACTURA"
            mIdDoc = 1
        Case "BOLETA DE VENTA"
            mIdDoc = 3
        Case Else '--CANCELAR
            MsgBox "Seleccione Otra vez el Documento", vbExclamation, xTitulo
            Exit Function
    End Select

    nSQL = "SELECT vta_ventas.tipdoc, vta_ventas.numser, vta_ventas.numdoc " _
            + vbCr + " FROM  vta_ventas " _
            + vbCr + " WHERE vta_ventas.tipdoc=" + CStr(mIdDoc) + " AND vta_ventas.numser='" + txt_cb(5).Text + "' AND vta_ventas.numdoc='" + txt_cb(2).Text + "';"
            
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.RecordCount = 1 Then
        MsgBox "El número de documento que quiere emitir ya existe", vbInformation, xTitulo
        Set RstTmp = Nothing
        txt_cb(2).SetFocus
        Exit Function
    End If
    Set RstTmp = Nothing
    '--GENERAR EL NUMERO DE ASIENTO
    xNumAsiento = NuevoNumAsiento(2, mMesActivo, xCon)
    '--CARGAR LOS ID'S DE LAS CUENTAS CONTABLES DEL DOCUMENTO Y DEL IMPUESTO
    pCargarCuentaContable IDCuentaDoc, IDCuentaImp, mIdDoc, NulosN(lbl_cb_cod(4).Caption), 1
    '--CARGAR EL TIPO DE CAMBIO SEGUN FECHA DE REGISTRO
    RST_Busq RstTmp, "SELECT con_tc.fecha, con_tc.impcom, con_tc.impven From con_tc WHERE (((con_tc.fecha)=CDate('" + txtfecha(1).Valor + "')) AND ((con_tc.idmon)=2));", xCon
    If RstTmp.BOF = False Or RstTmp.BOF = False Or RstTmp.RecordCount <> 0 Then
        mTipoCambio = NulosN(RstTmp.Fields("impven"))
    Else
        mTipoCambio = 0#
    End If
    Set RstTmp = Nothing
        
On Error GoTo LaCague
    xCon.BeginTrans

    RST_Busq RstCab, "SELECT TOP 1 * FROM vta_ventas", xCon
    RST_Busq RstDiario, "SELECT TOP 1 * FROM con_diario", xCon
    
    xId = HallaCodigoTabla("vta_ventas", xCon, "id")
    
    RstCab.AddNew
    RstCab("id") = xId
    RstCab("tipdoc") = mIdDoc
    RstCab("idcli") = 0
    RstCab("numser") = txt_cb(5).Text
    RstCab("numdoc") = txt_cb(2).Text
    RstCab("Fchdoc") = CDate(txtfecha(1).Valor)
    RstCab("Fchven") = CDate(txtfecha(1).Valor)
    RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    RstCab("idconpag") = 0
    RstCab("idmon") = 1
    RstCab("impbru") = 0
    RstCab("impinaf") = 0
    RstCab("impigv") = 0
    RstCab("impisc") = 0
    RstCab("impotr") = 0
    RstCab("imptotdoc") = 0
    RstCab("impsal") = 0
    RstCab("numreg") = Format(mMesActivo, "00") + Trim(xNumAsiento)
    RstCab("anulado") = -1
    'Determinamos si es una exportacion
    RstCab("idtipven") = 0 'en el cual puede ser venta afecta o inafecta para el registro de de ventas
                           'se valida por programa ver tabla mae_tipoventa
    RstCab.Update

    '--ASIENTO DEL TOTAL
    pGenerarAsiento RstDiario, 2, NulosN(lbl_codigo.Caption), xNumAsiento, mTipoCambio, CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra), CDate(txtfecha(1).Valor), IDCuentaDoc, 1, 0, True

    '--ASIENTO DEL IMPUESTO
    pGenerarAsiento RstDiario, 2, NulosN(lbl_codigo.Caption), xNumAsiento, mTipoCambio, CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra), CDate(txtfecha(1).Valor), IDCuentaImp, 1, 0, False
    
    xCon.CommitTrans
    MsgBox "El documento anulado se generó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDiario = Nothing
    EmitirAnulada = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstTmp = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)

End Function


Private Sub pCargarAlmacen()
    '--cargarrá el almacen seleccionado al momento de logearse el usuario
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    If mIdAlmacen = 0 Then
        MsgBox "Inicie otra vez su Sesión" + vbCr + "Seleccione Otra vez el Almacén" + vbCr + "No puede continuar", vbExclamation, xTitulo
        GoTo Salir
    End If
    nSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion AS nombre, alm_almacenes.id AS cod " _
                + vbCr + " FROM alm_almacenes WHERE alm_almacenes.id = " & mIdAlmacen & " ;"
            
    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.State = 0 Then Exit Sub
    If RstTmp.RecordCount > 0 Then
        '---VENTANA PRINCIPAL
        txt_cb(3).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(3).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(3).Caption = RstTmp.Fields(2) & "" '--CODIGO
        '---VENTANA 2VENTANA OTROS_BOTONES
        txt_cb(4).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(4).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(4).Caption = RstTmp.Fields(2) & "" '--CODIGO
    Else
        '---VENTANA PRINCIPAL
        txt_cb(3).Text = "":    lbl_cb(3).Caption = "":    lbl_cb_cod(3).Caption = ""
        '---VENTANA 2VENTANA OTROS_BOTONES
        txt_cb(4).Text = "":    lbl_cb(4).Caption = "":    lbl_cb_cod(4).Caption = ""
        
    End If
    '-----BLOQUEANDO LOS OBJETOS RELACIONADOS AL ALMACEN
    '---VENTANA PRINCIPAL
    txt_cb(3).Locked = True:    cb(3).Enabled = False
    '---VENTANA 2VENTANA OTROS_BOTONES
    txt_cb(4).Locked = True:    cb(4).Enabled = False
Salir:
    Set RstTmp = Nothing
End Sub

Private Sub pCargarSegundaVentana(Index As Integer)
    '--cargara las opciones para modificar, anular, eliminar, imprimir ticket
    '--Index es el indice de objeto CmdOtros
    '--     0::Modificar 1::Eliminar 2:Anular  3::Emitir Doc. Anulados  4::Imprimir Ticket

    '--DAR POSICION AL FORM_2
    fra(10).Visible = True
    fra(10).Left = 3270
    fra(10).Top = 1725
    cb_doc.Clear
    
    cb(2).Visible = True '--SI ANTERIORMENTE SE ELIGIO EMITIR DOC. ANULADOS
    txt_cb(5).Text = "" '--LIMPIO EL NUM SERIE, POR DEFECTO ESTE LIMPIA NUM DOC EN txt_cb_Change(5)
    If Trim(txt_cb(2).Text) <> "" Then txt_cb(2).Text = ""
    '--CARGANDO LOS TIPOS DE DOCUMENTOS
    Select Case mNivelAcceso 'SE CARGA EN pCargarUsuario()
        Case 1 '--VENDEDOR
            cb_doc.AddItem "Cotización"
            cb_doc.ListIndex = 0
        Case 2 '--CAJERO
            cb_doc.AddItem "Factura"
            cb_doc.AddItem "Boleta de Venta"
            cb_doc.ListIndex = IIf(mIdDocumento = 3, 1, 0) '--COLOCANDO EL SETFOCUS SEGUN DOCUMENTO QUE INVOQUE LA OPCION OTROS
        Case 3 '--SUPERVISOR
            cb_doc.AddItem "Factura"
            cb_doc.AddItem "Boleta de Venta"
            cb_doc.ListIndex = IIf(mIdDocumento = 3, 1, 0)
    End Select
    '--COLOCANDO TITULO
    Select Case Index
        Case 0 '--mofificar
            lbl_titulo(0).Caption = "Modificar Documento"
        Case 1 '--eliminar
            lbl_titulo(0).Caption = "Eliminar Documento"
        Case 2 '--anular
            lbl_titulo(0).Caption = "Anular Documento"
        Case 3 '--emitir documentos anulados
            lbl_titulo(0).Caption = "Emitir Documentos Anulados"
        Case 4 '--imprimir documentos
            lbl_titulo(0).Caption = "Imprimir Documento"
    End Select
    '--BLOQUEANDO BOTONES
    Select Case Index
        Case 0 '--mofificar
            txtfecha(1).Valor = "":     txtfecha(1).Enabled = False
        Case 1 '--eliminar
            txtfecha(1).Valor = "":     txtfecha(1).Enabled = False
        Case 2 '--anular
            txtfecha(1).Valor = "":     txtfecha(1).Enabled = False
        Case 3 '--emitir documentos anulados
            txtfecha(1).Valor = "":     txtfecha(1).Enabled = True
            cb(2).Visible = False
        Case 4 '--imprimir documentos
            txtfecha(1).Valor = "":     txtfecha(1).Enabled = False
    End Select
    If mNivelAcceso = 1 Then '--SI ES VENDEDOR
        txt_cb(5).Enabled = False
        cb(5).Enabled = False
    End If
    '--COLOCANDO EL ENFOQUE
    Select Case mIdEventoOtrosBotones
        Case 1, 2, 3, 5
            cb_doc.SetFocus
        Case 4
            txtfecha(1).SetFocus
    End Select
    
End Sub


Private Function Modificar() As Boolean
    On Error GoTo error
    '--ACTIVANDO LOS BOTOENS DE FORM. PRINCIPAL SEGUN TIPO DE DOCUMENTO SELECCIONADO
    SeEjecuto = False '--ES FALSE PARA QUE NO APAREZCA LA PREGUNTA SI DESEA CONSERVAR LOS DATOS ANTERIOERES
    Select Case UCase(cb_doc.Text)
        Case "COTIZACIÓN"
            CmdTipo_Click 0
            CmdTipo(1).Enabled = False
            CmdTipo(2).Enabled = False
        Case "FACTURA"
            CmdTipo_Click 1
            CmdTipo(0).Enabled = False
            CmdTipo(2).Enabled = False
        Case "BOLETA DE VENTA"
            CmdTipo_Click 2
            CmdTipo(0).Enabled = False
            CmdTipo(1).Enabled = False
        Case Else '--CANCELAR
            CmdDoc_Click 1
            Exit Function
    End Select
    SeEjecuto = True
    '--------------------
    If UCase(cb_doc.Text) = "COTIZACIÓN" Then
        txt_cotiza(1).Text = NulosN(lbl_codigo.Caption)
        txt_cotiza_KeyDown 1, 13, 0  '--HAY UNA RUTINA QUE CARGA LOS DATOS SEGUN NUMERO DE CORIZACION
        
        '--COLOCANDO LA SERIE Y EL NUMERO
        lbl_num(0).Caption = ""
        lbl_num(1).Caption = Format(NulosN(lbl_codigo.Caption), "0000000000")
        lbl_encabezado(1).Caption = lbl_num(1).Caption

        
    Else
        Dim RstTmp As New ADODB.Recordset
        Dim nSQL As String
        nSQL = "SELECT vta_ventas.id, vta_ventas.numser, vta_ventas.numdoc, vta_ventas.fchdoc, vta_ventas.idalm, alm_almacenes.descripcion AS almdesc, vta_ventas.idalm AS almcod, vta_ventas.idcli as cliid, mae_cliente.nombre AS clidesc, mae_cliente.numruc AS clinum, vta_ventas.idmon AS monid, mae_moneda.descripcion AS mondesc,vta_ventas.moddes ,vta_ventas.tipdes " _
            + vbCr + " FROM ((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN alm_almacenes ON vta_ventas.idalm = alm_almacenes.id " _
            + vbCr + " WHERE (((vta_ventas.id)= " + CStr(NulosN(lbl_codigo.Caption)) + "));"
        RST_Busq RstTmp, nSQL, xCon
        pPonerDatosEncabezado RstTmp, False
        Set RstTmp = Nothing
            
        nSQL = "SELECT vta_ventasdet.iditem, vta_ventasdet.idunimed, alm_inventario.descripcion, mae_unidades.abrev, vta_ventasdet.canpro, vta_ventasdet.preuni, vta_ventasdet.valdes, vta_ventasdet.imptot, alm_inventario.idtipven, alm_inventario.idcuentaven " _
            + vbCr + " FROM vta_ventasdet LEFT JOIN (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) ON vta_ventasdet.iditem = alm_inventario.id " _
            + vbCr + " WHERE (((vta_ventasdet.idvta) = " + CStr(NulosN(lbl_codigo.Caption)) + ")) " _
            + vbCr + " ORDER BY alm_inventario.descripcion ; "
        RST_Busq RstTmp, nSQL, xCon
        pPonerDatosDetalle RstTmp
    End If
    
    QueHace = 2
    MsgBox "La " + cb_doc.Text + " esta listo para ser modificado...", vbInformation, xTitulo
    
    CmdTipo(5).Caption = "Cancelar"
    Modificar = True
    Exit Function
error:
    Set RstTmp = Nothing
End Function

Private Sub pCargarNumeroDoc()
    Dim mPosicionArray As Integer
    Select Case mIdDocumento
        Case 0: mPosicionArray = 0 '--TICKET
        Case 1: mPosicionArray = 1 '--FACTURA
        Case 3: mPosicionArray = 2 '--BOLETA
    End Select
    LimpiaText lbl_num
    '--DEL Nº SERIE
    If NulosN(ArrDocumento(mPosicionArray, 1)) <> 0 Then
        lbl_num(0).Caption = ArrDocumento(mPosicionArray, 1)
    Else
        Dim Rst As New ADODB.Recordset
        RST_Busq Rst, "SELECT * FROM alm_numseries WHERE idtipdoc = " + CStr(mIdDocumento) + " AND idalm = " + CStr(mIdAlmacen), xCon
        If Rst.RecordCount <> 0 Then
            lbl_num(0).Caption = Rst("numser") & ""
        End If
        Set Rst = Nothing
    End If
    If mIdDocumento <> 0 Then
        lbl_num(1).Caption = HallaNumdocVenta(mIdDocumento, lbl_num(0).Caption, xCon)
        '--ESTE NUMERO PUEDE CAMBIAR SI AL MOMENTO DE INGRESAR LOS DATOS, OTRO USUARIO INGRESA UNA VENTA ANTES, PARA QUE NO SE DUPLIQUE LOS NUMEROS
        '--AL GRABAR VALIDARA SI EXISTE EL NUMERO, SI EXISTE SE VUELVE A GENERAR EL NUMERO
    Else
        lbl_num(1).Caption = Format(fHallarNumeroDoc("pvt_cotizacion", "numdoc"), "0000000000")
    End If
    lbl_encabezado(1).Caption = lbl_num(0).Caption + " " + lbl_num(1).Caption

End Sub

Private Sub pAplicarDescuento(Optional IdItem As Integer = -1)
    Dim mRow As Integer
    Dim RstTmp As ADODB.Recordset
    Dim RstTemp As New ADODB.Recordset
    Dim nSQL As String
    '-------------------------
    Dim mCodItem As Integer
    Dim mCant As Integer
    Dim mCliente As Integer
    '--ESTE PROCEDIMIENTO APLICARA EL DESCUENTO SEGUN LAS OPCIONES SELECCIONADAS
    If Fg1.Rows = 1 Then Exit Sub
    '-------------------------
    Agregando = False
    '------
    Me.MousePointer = vbHourglass
    If opt_descuento(0).Value = True Or (opt_descuento1(0).Value = False And opt_descuento1(1).Value = False) Then
        With Fg1
            For mRow = 1 To .Rows - 1
                mCodItem = NulosN(.TextMatrix(mRow, 1))
                If IdItem <> -1 And mCodItem <> IdItem Then GoTo Avance: '--SOLO CUANDO SE MODIFICA EL la CANTIDAD
                .TextMatrix(mRow, 7) = 0
                '----------------------
                If IdItem <> -1 Then Exit For
                '----------------------
'                Fg1_CellChanged mRow, 7
                '----------------------
Avance:
            Next mRow
        End With
    Else
        mCliente = NulosN(lbl_cb_cod(0).Caption)
        With Fg1
            For mRow = 1 To .Rows - 1
                mCodItem = NulosN(.TextMatrix(mRow, 1))
                mCant = NulosN(.TextMatrix(mRow, 5))
                If IdItem <> -1 And mCodItem <> IdItem Then GoTo Avance1: '--SOLO CUANDO SE MODIFICA EL la CANTIDAD
                If opt_descuento(1).Value = True Then '--GENERAL
                    nSQL = "SELECT pvt_descgeneral.porcentaje, pvt_descgeneral.valor " _
                        + vbCr + " FROM pvt_descgeneral " _
                        + vbCr + " WHERE (((pvt_descgeneral.iditem)= " & mCodItem & " ) AND ((pvt_descgeneral.inicio)<= " & mCant & " ) AND ((pvt_descgeneral.fin)>= " & mCant & " )); "

                Else '--CORPORATIVO
                    nSQL = "SELECT pvt_desccorporativo.porcentaje, pvt_desccorporativo.valor " _
                        + vbCr + " FROM pvt_desccorporativo " _
                        + vbCr + " WHERE (((pvt_desccorporativo.idcli)= " & mCliente & " ) AND ((pvt_desccorporativo.iditem)= " & mCodItem & " ) AND ((pvt_desccorporativo.inicio)<= " & mCant & " ) AND ((pvt_desccorporativo.fin)>= " & mCant & " )); "
                End If
                Set RstTemp = Nothing
                RST_Busq RstTemp, nSQL, xCon
                If RstTemp.EOF = False Or RstTemp.BOF = False Or RstTemp.RecordCount <> 0 Then
                    If opt_descuento1(0).Value = True Then '--PORCENTAJE
                        .TextMatrix(mRow, 7) = NulosN(RstTemp.Fields("porcentaje"))
                    Else '--VALOR
                        .TextMatrix(mRow, 7) = NulosN(RstTemp.Fields("valor"))
                    End If
                    
                Else
                    .TextMatrix(mRow, 7) = 0
                End If
                '----------------------
                If IdItem <> -1 Then Exit For
                '----------------------
'                Fg1_CellChanged mRow, 7
Avance1:
            Next mRow
        End With
    End If
    Set RstTmp = Nothing
    Agregando = False
    If IdItem = -1 Then pCalculosTotales
    Me.MousePointer = vbDefault
End Sub

Private Sub Imprimir(mTipoDoc As Integer)
    Dim RsPDoc As New ADODB.Recordset
    Dim RsPCab As New ADODB.Recordset
    Dim RsPDet As New ADODB.Recordset
    
    Dim xRsDoc As New ADODB.Recordset
    Dim xRsDet As New ADODB.Recordset
       
    Dim mPosicionArray As Integer
    Dim mIdImpresion As Integer '--INDICA LA PLANTILLA DE IMPRESION


    Select Case mIdDocumento
        Case 0: mPosicionArray = 0 '--TICKET
        Case 1: mPosicionArray = 1 '--FACTURA
        Case 3: mPosicionArray = 2 '--BOLETA
    End Select
    If NulosN(ArrDocumento(mPosicionArray, 2)) = 0 Then
        MsgBox "No se le ha definido la plantilla de impresión para este tipo de documento" + vbCr + "Solicite al Supervisor le Configure la Plantilla de Impresión", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    mIdImpresion = NulosN(ArrDocumento(mPosicionArray, 2))
    
    RST_Busq RsPDoc, "SELECT * FROM var_plantilladoc WHERE id = " & mIdImpresion & " ", xCon
    '-----------------
    Dim nSQLCab As String
    Dim nSQLDet As String

    If mTipoDoc = 0 Then '--COTIZACIÓN
        nSQLCab = "SELECT pvt_cotizacion.id, pvt_cotizacion.numdoc, pvt_cotizacion.fchdoc, pvt_cotizacion.idcli AS cliid, mae_cliente.numruc AS clinum, mae_cliente.nombre AS clidesc, pvt_cotizacion.idmon AS monid, mae_moneda.descripcion AS mondesc,pvt_cotizacion.moddes,pvt_cotizacion.tipdes " _
            + vbCr + " FROM (pvt_cotizacion LEFT JOIN mae_moneda ON pvt_cotizacion.idmon = mae_moneda.id) LEFT JOIN mae_cliente ON pvt_cotizacion.idcli = mae_cliente.id " _
            + vbCr + " WHERE  pvt_cotizacion.id=" + CStr(NulosN(lbl_codigo.Caption)) + ";"
        '------DEL DETALLE
        nSQLDet = "SELECT pvt_cotizaciondet.iditem, pvt_cotizaciondet.idunimed, alm_inventario.descripcion, mae_unidades.abrev, pvt_cotizaciondet.canpro, pvt_cotizaciondet.preuni, pvt_cotizaciondet.valdes, pvt_cotizaciondet.imptot, pvt_cotizaciondet.preunibru, alm_inventario.idtipven, alm_inventario.idcuentaven " _
            + vbCr + " FROM mae_unidades RIGHT JOIN (pvt_cotizaciondet LEFT JOIN alm_inventario ON pvt_cotizaciondet.iditem = alm_inventario.id) ON mae_unidades.id = alm_inventario.idunimed " _
            + vbCr + " WHERE (((pvt_cotizaciondet.idcot)=" + CStr(NulosN(lbl_codigo.Caption)) + "));"
    Else
        nSQLCab = "SELECT vta_ventas.id, vta_ventas.*, alm_almacenes.descripcion AS almacen, mae_cliente.nombre AS nombre, mae_cliente.numruc, mae_cliente.dir " _
            + vbCr + " FROM ((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN alm_almacenes ON vta_ventas.idalm = alm_almacenes.id " _
            + vbCr + " WHERE (((vta_ventas.id)= " + CStr(NulosN(lbl_codigo.Caption)) + "));"
        
        nSQLDet = "SELECT vta_ventasdet.iditem, vta_ventasdet.idunimed, alm_inventario.descripcion, mae_unidades.abrev, vta_ventasdet.canpro, vta_ventasdet.preuni, vta_ventasdet.valdes, vta_ventasdet.imptot, alm_inventario.idtipven, alm_inventario.idcuentaven " _
            + vbCr + " FROM vta_ventasdet LEFT JOIN (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) ON vta_ventasdet.iditem = alm_inventario.id " _
            + vbCr + " WHERE (((vta_ventasdet.idvta) = " + CStr(NulosN(lbl_codigo.Caption)) + ")) " _
            + vbCr + " ORDER BY alm_inventario.descripcion ; "
    End If
    
    RST_Busq xRsDoc, nSQLCab, xCon
    RST_Busq xRsDet, nSQLDet, xCon
    '-----------------
    RST_Busq RsPCab, "SELECT * FROM var_plantilladoc WHERE id = " + CStr(mIdImpresion) + " ", xCon
    RST_Busq RsPCab, "SELECT * FROM var_plantillacab WHERE idplan = " + CStr(mIdImpresion) + " ORDER BY item", xCon
    RST_Busq RsPDet, "SELECT * FROM var_plantilladet WHERE idplan = " + CStr(mIdImpresion) + " ORDER BY item", xCon

'    Exit Sub
    
    Printer.Font = "Super Draft 15cpi"

    Printer.FontBold = True
    Printer.FontSize = 9
    Printer.ScaleMode = 6
    
    Dim xCampo, xFormato As String
    Dim nCampo As String
    Dim nValor As String
    Dim mCampo As Integer
    
    'imprime cabezera
    Do While Not RsPCab.EOF And xRsDoc.EOF = False
        xCampo = UCase(RsPCab.Fields("campo") & "")
        For mCampo = 0 To xRsDoc.Fields.Count - 1
            nCampo = UCase(xRsDoc.Fields(mCampo).Name)
            If nCampo = xCampo Then
                xFormato = NulosC(RsPCab.Fields("formato"))
                Printer.CurrentX = NulosN(RsPCab.Fields("posx"))
                Printer.CurrentY = NulosN(RsPCab.Fields("posy"))
                
                If LCase(nCampo) <> "x-numeletra" Then
                    If xFormato = "" Then
                        nValor = NulosC(xRsDoc.Fields(nCampo))
                        Printer.Print nValor
                    Else
                        'RSet nValor = Format((NulosC(xRsDoc.Fields(nCampo))), xFormato) '--ALINEAR A LA DERECHA
                        nValor = Format((NulosC(xRsDoc.Fields(nCampo))), xFormato) '--ALINEAR A LA DERECHA
                        Printer.Print nValor
                    End If
                Else
                    Printer.Print "Son : "; NumeroLetra(NulosN(xRsDoc.Fields("imptotdoc")), xRsDoc.Fields("idmon"))
                End If
            End If
        Next
        RsPCab.MoveNext
    Loop
    
    'imprime detalle
    '-------------------------------------------
    Dim Fila As Integer
    If xRsDet.RecordCount <> 0 Then
        xRsDet.MoveFirst
        Fila = NulosN(RsPDet.Fields("posy"))
    End If
    Do While Not xRsDet.EOF
        If RsPDet.RecordCount <> 0 Then RsPDet.MoveFirst
        Do While Not RsPDet.EOF
            xCampo = UCase(RsPDet.Fields("campo") & "")
            
            For mCampo = 0 To xRsDet.Fields.Count - 1
                nCampo = UCase(xRsDet.Fields(mCampo).Name)
                If nCampo = xCampo Then
                    xFormato = NulosC(RsPDet.Fields("formato"))
                    Printer.CurrentX = NulosN(RsPDet.Fields("posx"))
                    Printer.CurrentY = Fila
                    If xFormato = "" Then
                        nValor = NulosC(xRsDet.Fields(nCampo))
                        Printer.Print nValor
                    Else
                        'RSet nValor = Format((NulosC(xRsDet.Fields(nCampo))), xFormato) '--ALINEAR A LA DERECHA
                        nValor = Format((NulosC(xRsDet.Fields(nCampo))), xFormato) '--ALINEAR A LA DERECHA
                        Printer.Print nValor
                    End If
                    
                End If
            Next
            
            RsPDet.MoveNext
        Loop
        Fila = Fila + 4
        xRsDet.MoveNext
    Loop
    
    Printer.EndDoc
    
    Set xRsDoc = Nothing
    Set xRsDet = Nothing
    Set RsPDoc = Nothing
    Set RsPDet = Nothing
End Sub

Private Sub pCargarNumeroDocumento()
    Dim nNumero As String
    Dim nNumeroAnterior As String  '--ALMACENAR EL NUMERO DE DOCUMENTO ANTES DE CAMBIAR
    If mIdDocumento = 0 Then Exit Sub
    nNumeroAnterior = lbl_num(1).Caption
    nNumero = InputBox("Ingrese el número de la " + nDocumento + vbCr + "Número a Reemplazar: " + nNumeroAnterior, "Ingrese nuevo número de la " + nDocumento)

    If Trim(nNumero) = "" Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    nNumero = Format(NulosN(nNumero), "0000000000")
    
    RST_Busq RstTmp, "SELECT vta_ventas.tipdoc, vta_ventas.numser, vta_ventas.numdoc From vta_ventas " _
        & " WHERE (((vta_ventas.tipdoc)=" + CStr(mIdDocumento) + ") AND ((vta_ventas.numser)='" + Trim(lbl_num(0).Caption) + "') " _
        & " AND ((vta_ventas.numdoc)='" + nNumero + "'))", xCon

    If RstTmp.RecordCount = 1 Then
        Set RstTmp = Nothing
        MsgBox "El numero de documento que quiere registrar ya existe", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    Set RstTmp = Nothing
    lbl_num(1).Caption = nNumero
    lbl_encabezado(1).Caption = lbl_num(0).Caption + " " + lbl_num(1).Caption
    
End Sub
