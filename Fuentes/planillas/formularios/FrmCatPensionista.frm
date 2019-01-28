VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmCatPensionista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Nómina del Personal - Pensionista"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Index           =   0
      Left            =   960
      MaxLength       =   40
      TabIndex        =   21
      Text            =   "txt(0)"
      Top             =   4530
      Width           =   1140
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   345
      Index           =   13
      Left            =   90
      TabIndex        =   19
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
         Caption         =   "Categoría: Pensionista"
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
         TabIndex        =   20
         Top             =   45
         Width           =   2370
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   405
      Index           =   12
      Left            =   75
      TabIndex        =   17
      Top             =   3885
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
         TabIndex        =   18
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
      Left            =   7545
      TabIndex        =   7
      Top             =   3870
      Width           =   1755
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Grabar"
      Height          =   420
      Index           =   0
      Left            =   5565
      TabIndex        =   6
      Top             =   3870
      Width           =   1755
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   3300
      Left            =   90
      TabIndex        =   8
      Top             =   465
      Width           =   9300
      _cx             =   16404
      _cy             =   5821
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
         Height          =   2880
         Index           =   2
         Left            =   10245
         TabIndex        =   11
         Top             =   45
         Width           =   9210
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   2880
         Index           =   1
         Left            =   9945
         TabIndex        =   10
         Top             =   45
         Width           =   9210
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2880
         Index           =   0
         Left            =   45
         TabIndex        =   9
         Top             =   45
         Width           =   9210
         Begin VB.Frame fra 
            Caption         =   "[ Situación del Pensionista ]"
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
            Left            =   60
            TabIndex        =   35
            Top             =   1545
            Width           =   4965
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   2
               Left            =   540
               Picture         =   "FrmCatPensionista.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Seleccione la Situación del Pensionista"
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   2
               Left            =   135
               MaxLength       =   20
               TabIndex        =   4
               Tag             =   "null"
               Text            =   "txt_cb(2)"
               Top             =   240
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
               TabIndex        =   38
               Top             =   240
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Situación del Trabajador"
               Height          =   195
               Index           =   2
               Left            =   3120
               TabIndex        =   37
               Top             =   75
               Visible         =   0   'False
               Width           =   1725
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
               Left            =   795
               TabIndex        =   39
               Top             =   240
               Width           =   4065
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
            Index           =   17
            Left            =   60
            TabIndex        =   30
            Top             =   2190
            Width           =   4965
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   3
               Left            =   540
               Picture         =   "FrmCatPensionista.frx":0132
               Style           =   1  'Graphical
               TabIndex        =   31
               ToolTipText     =   "Seleccione el Tipo de Pago"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   3
               Left            =   135
               MaxLength       =   20
               TabIndex        =   5
               Text            =   "txt_cb(3)"
               Top             =   285
               Width           =   645
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Pago"
               Height          =   195
               Index           =   3
               Left            =   3030
               TabIndex        =   33
               Top             =   75
               Visible         =   0   'False
               Width           =   960
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
               Left            =   3795
               TabIndex        =   32
               Top             =   300
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
               Left            =   795
               TabIndex        =   34
               Top             =   285
               Width           =   4065
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
            Left            =   60
            TabIndex        =   23
            Top             =   780
            Width           =   9075
            Begin VB.TextBox txt 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   1
               Left            =   7440
               MaxLength       =   12
               TabIndex        =   3
               Tag             =   "null"
               Text            =   "txt(1)"
               ToolTipText     =   "Código Unico de Identificación del Sistema Privado de Pensiones"
               Top             =   270
               Width           =   1455
            End
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   1
               Left            =   540
               Picture         =   "FrmCatPensionista.frx":0264
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Seleccione el Régimen Pensionario"
               Top             =   315
               Width           =   210
            End
            Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
               Height          =   300
               Index           =   0
               Left            =   5445
               TabIndex        =   2
               ToolTipText     =   "Ingrese la Fecha de Inscripción"
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
               Index           =   1
               Left            =   135
               MaxLength       =   20
               TabIndex        =   1
               Tag             =   "null"
               Text            =   "txt_cb(1)"
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
               TabIndex        =   29
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
               TabIndex        =   28
               Top             =   360
               Width           =   945
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Régimen Pensionario"
               Height          =   195
               Index           =   1
               Left            =   2490
               TabIndex        =   26
               Top             =   90
               Visible         =   0   'False
               Width           =   1500
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
               TabIndex        =   25
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
               TabIndex        =   27
               Top             =   285
               Width           =   3390
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
            TabIndex        =   12
            Top             =   30
            Width           =   9075
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   0
               Left            =   540
               Picture         =   "FrmCatPensionista.frx":0396
               Style           =   1  'Graphical
               TabIndex        =   13
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
               TabIndex        =   16
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
               TabIndex        =   15
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
               TabIndex        =   14
               Top             =   285
               Width           =   8160
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
      Left            =   270
      TabIndex        =   22
      Top             =   4485
      Width           =   495
   End
End
Attribute VB_Name = "FrmCatPensionista"
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
    TabOne1.CurrTab = 0
    pPonerDatos
End Sub


'*******************************************************************************************

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

Private Sub Form_Load()
    CentrarFrm Me
End Sub

Private Sub TabOne1_Click()
    If TabOne1.CurrTab = 0 Then
        If Agregando = False Then txt_cb(0).SetFocus
    Else
        If Agregando = False Then txt_cb(7).SetFocus
    End If
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
                + vbCr + " WHERE mae_tipotrabajadorcat.idcat = 2 " _
                + vbCr + " ORDER BY mae_tipotrabajador.codsun;"

        Case 1 '--REGIMEN PENSIONARIO
            nTitulo = "Buscando Situación de Derechohabiente"
            nSQL = "SELECT mae_regimenpen.id, mae_regimenpen.descripcion AS nombre, mae_regimenpen.id AS cod, mae_regimenpen.cuspp " _
                + vbCr + " From mae_regimenpen " _
                + vbCr + " ORDER BY mae_regimenpen.codsun;"
        
        Case 2 '--SITUACION DE TRABAJO
            nTitulo = "Buscando Situación del Trajabador"
            nSQL = "SELECT mae_situacion.id, mae_situacion.descripcion AS nombre, mae_situacion.id AS cod " _
                + vbCr + " From mae_situacion " _
                + vbCr + " Where (((mae_situacion.afiliado) = 0 )) " _
                + vbCr + " ORDER BY mae_situacion.codsun;"

        Case 3 '--TIPO DE PAGO
            nTitulo = "Buscando Tipo de Pago"
            nSQL = "SELECT mae_tipopago.id, mae_tipopago.descripcion AS nombre, mae_tipopago.id AS cod " _
                + vbCr + " From mae_tipopago " _
                + vbCr + " ORDER BY mae_tipopago.codsun;"
    
    End Select
    
    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
            
    Dim xRs As New ADODB.Recordset
    If Index <> 3 Then
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
    Else
        '--SOLO OCUPACION
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", CualquierParte
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
        Case 1 '--REGIMEN PENSIONARIO
            If NulosN(xRs.Fields("cuspp")) = -1 Then
                txt(1).Visible = True
                lbl(1).Visible = True
            Else
                txt(1).Text = ""
                txt(1).Visible = False
                lbl(1).Visible = False
            End If
            txtfecha(0).SetFocus
        Case 2 '--SITUACION DE PENSIONISTA
            txt_cb(3).SetFocus
        Case 3 '--TIPO DE PAGO
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
                + vbCr + " WHERE mae_tipotrabajador.id = " & NulosN(txt_cb(Index).Text) & " AND mae_tipotrabajadorcat.idcat = 2 ;"

        Case 1 '--REGIMEN PENSIONARIO
            nSQL = "SELECT mae_regimenpen.id, mae_regimenpen.descripcion AS nombre, mae_regimenpen.id AS cod, mae_regimenpen.cuspp " _
                + vbCr + " From mae_regimenpen " _
                + vbCr + " WHERE mae_regimenpen.id = " & NulosN(txt_cb(Index).Text) & ";"
        

        Case 2 '--SITUACION DE PENSIONISTA

            nSQL = "SELECT mae_situacion.id, mae_situacion.descripcion AS nombre, mae_situacion.id AS cod " _
                + vbCr + " From mae_situacion " _
                + vbCr + " Where (((mae_situacion.afiliado) = 0 )) " _
                + vbCr + " AND mae_situacion.id = " & NulosN(txt_cb(Index).Text) & ";"
    
        Case 3 '--TIPO DE PAGO
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
        Case 1 '--REGIMEN PENSIONARIO
            If NulosN(RstTmp.Fields("cuspp")) = -1 Then
                txt(1).Visible = True
                lbl(1).Visible = True
            Else
                txt(1).Text = ""
                txt(1).Visible = False
                lbl(1).Visible = False
            End If
            If Agregando = False Then txtfecha(0).SetFocus
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
    nSQL = "SELECT pla_categoria2.* From pla_categoria2 WHERE (((pla_categoria2.idemp)=" & mIdEmpleado & "));"

    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.RecordCount = 0 Then
        Quehace = 1
        Exit Sub
    End If
    Quehace = 2
    Agregando = True

    '************************************************************TAB 0
    '--TIPO TRABAJADOR
    If NulosN(RstTmp("idtippen")) <> 0 Then
        txt_cb(0).Text = NulosN(RstTmp("idtippen"))
        txt_cb_Validate 0, False
    End If
    '--REGIMEN PENSIONARIO
    If NulosN(RstTmp("idregpen")) <> 0 Then
        txt_cb(1).Text = NulosN(RstTmp("idregpen"))
        txt_cb_Validate 1, False
    End If
    If IsDate(RstTmp("fchins")) = True Then
        txtfecha(0).Valor = CDate(RstTmp("fchins"))
    End If
    txt(1).Text = NulosC(RstTmp("cuspp"))
    '--SITUACION DE TRABAJADOR
    If NulosN(RstTmp("idsituacion")) <> 0 Then
        txt_cb(2).Text = NulosN(RstTmp("idsituacion"))
        txt_cb_Validate 2, False
    End If
    '--TIPO DE PAGO
    If NulosN(RstTmp("idtippag")) <> 0 Then
        txt_cb(3).Text = NulosN(RstTmp("idtippag"))
        txt_cb_Validate 3, False
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

'    On Error GoTo LaCague

    xCon.BeginTrans

    '*****************************************************
    If Quehace = 1 Then
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_categoria2 ; ", xCon
        RstCab.AddNew
    Else
        RST_Busq RstCab, "SELECT * FROM pla_categoria2 WHERE idemp =  " & mIdEmpleado & " ;", xCon
    End If
    
    RstCab("idemp") = mIdEmpleado
    '************************************************************TAB 0
    '--TIPO PENSIONISTA
    RstCab("idtippen") = NulosN(txt_cb(0).Text)
    '--REGIMEN PENSIONARIO
    RstCab("idregpen") = NulosN(txt_cb(1).Text)
    If IsDate(txtfecha(0).Valor) = True Then
        RstCab("fchins") = CDate(txtfecha(0).Valor)
    End If
    RstCab("cuspp") = Trim(txt(1).Text)
    '--SITUACION DE TRABAJADOR
    RstCab("idsituacion") = NulosN(txt_cb(2).Text)
    '--TIPO DE PAGO
    RstCab("idtippag") = NulosN(txt_cb(3).Text)
    
    '--
    RstCab.Update

    MsgBox "Los datos del Pensionista " + IIf(Quehace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

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
    
    
'    If IsDate(txtfecha(0).Valor) = False Then
'        MsgBox "Ingrese la fecha de incripción al régimen pensionario", vbExclamation, xTitulo
'        txtfecha(0).SetFocus
'        Exit Function
'    End If
    
    If NulosN(txt(1).Text) <> 0 And txt(1).Visible = True Then
        MsgBox "Ingrese el Código Unico de Identificación del Sistema Privado de Pensiones", vbExclamation, xTitulo
        txt(1).SetFocus
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
