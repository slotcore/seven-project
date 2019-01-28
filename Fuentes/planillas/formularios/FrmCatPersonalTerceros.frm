VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Begin VB.Form FrmCatPersonalTerceros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Nómina del Personal - Personal de Terceros"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Index           =   0
      Left            =   900
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "txt(0)"
      Top             =   3375
      Width           =   1140
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   345
      Index           =   13
      Left            =   90
      TabIndex        =   11
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
         Caption         =   "Categoría: Personal de Terceros"
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
         TabIndex        =   12
         Top             =   45
         Width           =   3420
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   405
      Index           =   12
      Left            =   120
      TabIndex        =   9
      Top             =   2910
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
         TabIndex        =   10
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
      Left            =   7590
      TabIndex        =   8
      Top             =   2895
      Width           =   1755
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Grabar"
      Height          =   420
      Index           =   0
      Left            =   5610
      TabIndex        =   7
      Top             =   2895
      Width           =   1755
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   2325
      Left            =   90
      TabIndex        =   3
      Top             =   465
      Width           =   9300
      _cx             =   16404
      _cy             =   4101
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
         Height          =   1905
         Index           =   2
         Left            =   10245
         TabIndex        =   6
         Top             =   45
         Width           =   9210
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   1905
         Index           =   1
         Left            =   9945
         TabIndex        =   5
         Top             =   45
         Width           =   9210
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1905
         Index           =   0
         Left            =   45
         TabIndex        =   4
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
            Left            =   60
            TabIndex        =   20
            Top             =   1020
            Width           =   9075
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   2
               Left            =   5640
               Picture         =   "FrmCatPersonalTerceros.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   25
               ToolTipText     =   "Seleccione el SCTR Pensión"
               Top             =   270
               Width           =   210
            End
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   1
               Left            =   990
               Picture         =   "FrmCatPersonalTerceros.frx":0132
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Seleccione el SCTR Salud"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   1
               Left            =   600
               MaxLength       =   20
               TabIndex        =   1
               Tag             =   "null"
               Text            =   "txt_cb(1)"
               Top             =   285
               Width           =   645
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   2
               Left            =   5235
               MaxLength       =   20
               TabIndex        =   2
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
               Left            =   7305
               TabIndex        =   28
               Top             =   240
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pensión"
               Height          =   195
               Index           =   2
               Left            =   4455
               TabIndex        =   27
               Top             =   330
               Width           =   570
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
               Left            =   5880
               TabIndex        =   26
               Top             =   240
               Width           =   3045
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
               TabIndex        =   24
               Top             =   285
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lbl_capt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Salud"
               Height          =   195
               Index           =   1
               Left            =   105
               TabIndex        =   23
               Top             =   375
               Width           =   405
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
               Left            =   1245
               TabIndex        =   22
               Top             =   285
               Width           =   2940
            End
         End
         Begin VB.Frame fra 
            Caption         =   "[ RUC de Empresa de destaca o desplaza ]"
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
            Left            =   75
            TabIndex        =   15
            Top             =   180
            Width           =   9075
            Begin VB.CommandButton cb 
               Height          =   225
               Index           =   0
               Left            =   1305
               Picture         =   "FrmCatPersonalTerceros.frx":0264
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Seleccione la Empresa que destaca o desplaza"
               Top             =   315
               Width           =   210
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   0
               Left            =   135
               MaxLength       =   11
               TabIndex        =   0
               Text            =   "txt_cb(0)"
               Top             =   285
               Width           =   1395
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
               Caption         =   "RUC de Empresa de destaca o desplaza"
               Height          =   195
               Index           =   0
               Left            =   5850
               TabIndex        =   17
               Top             =   90
               Visible         =   0   'False
               Width           =   2880
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
               Left            =   1620
               TabIndex        =   19
               Top             =   270
               Width           =   7200
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
      Left            =   210
      TabIndex        =   14
      Top             =   3465
      Width           =   495
   End
End
Attribute VB_Name = "FrmCatPersonalTerceros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QueHace As Integer
Dim mCorrelativo As Long
Dim mIdEmpleado As Long
Dim Agregando As Boolean

Public Sub pRecibeLink(QueHace1 As Integer)
    QueHace = QueHace1
    mCorrelativo = mCorr
    With FrmNomina
        mIdEmpleado = .txt(0).Text
        lbl_persona.Caption = .lbl_persona(0).Caption
    End With
    '------
    Agregando = True
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cb
    LimpiaText lbl_cod
    TabOne1.CurrTab = 0
    pPonerDatos
    Agregando = False
End Sub


'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQl As String
    Dim nTitulo As String
    On Error GoTo error
    Select Case Index
        Case 0 '--EMPRESA QUE DESTACA
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "RUC":       xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
        
            nTitulo = "Buscando Empresas que destaca o desplaza"
            
            nSQl = "SELECT mae_empresadestaca.numruc, mae_empresadestaca.descripcion AS nombre, mae_empresadestaca.id AS cod " _
                + vbCr + " FROM mae_empresadestaca;"

        Case 1 '--SCTR - SALUD
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
            
            nTitulo = "Buscando SCTR - Salud"
            nSQl = "SELECT mae_sctrsalud.id, mae_sctrsalud.descripcion AS nombre, mae_sctrsalud.id AS cod " _
                + vbCr + " From mae_sctrsalud " _
                + vbCr + " WHERE mae_sctrsalud.id <> 1 " _
                + vbCr + " ORDER BY mae_sctrsalud.codsun;"

        Case 2 '--SCTR - PENSION
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
        
            nTitulo = "Buscando SCTR - Pensión"
            nSQl = "SELECT mae_sctrpension.id, mae_sctrpension.descripcion AS nombre, mae_sctrpension.id AS cod " _
                + vbCr + " From mae_sctrpension " _
                + vbCr + " WHERE mae_sctrpension.id <> 1 " _
                + vbCr + " ORDER BY mae_sctrpension.codsun;"
    
    End Select
    
            
    Dim xRs As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQl, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    lbl_cb(Index).ToolTipText = xRs.Fields(1) & "" '--NOMBRE
      
    Select Case Index
        Case 0 '--EMPRESA QUE DESTACA
            txt_cb(1).SetFocus
        Case 1 '--SCTR - SALUD
            txt_cb(2).SetFocus
        Case 2 '--SCTR - PENSION
            cmd(0).SetFocus
    End Select
Salir:
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
    Dim nSQl As String
    On Error GoTo error
    Select Case Index
        Case 0 '--TIPO DE TRABAJADOR
            
            nSQl = "SELECT mae_empresadestaca.numruc, mae_empresadestaca.descripcion AS nombre, mae_empresadestaca.id AS cod " _
                + vbCr + " FROM mae_empresadestaca " _
                + vbCr + " WHERE mae_empresadestaca.numruc  = '" & NulosN(txt_cb(Index).Text) & "';"
                              
        Case 1 '--SCTR - SALUD
            nSQl = "SELECT mae_sctrsalud.id, mae_sctrsalud.descripcion AS nombre, mae_sctrsalud.id AS cod " _
                + vbCr + " FROM mae_sctrsalud " _
                + vbCr + " WHERE mae_sctrsalud.id = " & NulosN(txt_cb(Index).Text) & " AND mae_sctrsalud.id <> 1 ;"

        Case 2 '--SCTR - PENSION
            nSQl = "SELECT mae_sctrpension.id, mae_sctrpension.descripcion AS nombre, mae_sctrpension.id AS cod " _
                + vbCr + " FROM mae_sctrpension " _
                + vbCr + " WHERE mae_sctrpension.id = " & NulosN(txt_cb(Index).Text) & " AND  mae_sctrpension.id <> 1 ;"
    End Select

    If xCon.State = 0 Then GoTo Salir
    RST_Busq RstTmp, nSQl, xCon

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
    Dim nSQl As String
'    On Error GoTo error
    nSQl = "SELECT pla_categoria5.*, mae_empresadestaca.numruc " _
        + vbCr + " FROM mae_empresadestaca INNER JOIN pla_categoria5 ON mae_empresadestaca.id = pla_categoria5.iddestaca " _
        + vbCr + " WHERE (((pla_categoria5.idemp)= " & mIdEmpleado & "));"

    RST_Busq RstTmp, nSQl, xCon
    
    If RstTmp.RecordCount = 0 Then
        QueHace = 1
        Exit Sub
    End If
    QueHace = 2
    Agregando = True

    '--EMPRESA QUE DETACA
    If NulosC(RstTmp("numruc")) <> "" Then
        txt_cb(0).Text = NulosN(RstTmp("numruc"))
        txt_cb_Validate 0, False
    End If
    '--SCTR SALUD
    If NulosN(RstTmp("sctrsalud")) <> 0 Then
        txt_cb(1).Text = NulosN(RstTmp("sctrsalud"))
        txt_cb_Validate 1, False
    End If
    '--SCTR PENSION
    If NulosN(RstTmp("sctrpension")) <> 0 Then
        txt_cb(2).Text = NulosN(RstTmp("sctrpension"))
        txt_cb_Validate 2, False
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
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " los datos del Personal de Terceros", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim xId As Integer

'    On Error GoTo LaCague

    xCon.BeginTrans

    '*****************************************************
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_categoria5 ; ", xCon
        RstCab.AddNew
    Else
        RST_Busq RstCab, "SELECT * FROM pla_categoria5 WHERE idemp =  " & mIdEmpleado & " ;", xCon
    End If
    
    RstCab("idemp") = mIdEmpleado
    '************************************************************TAB 0
    '--EMPRESA QUE DESTACA
    RstCab("iddestaca") = NulosN(lbl_cod(0).Caption)
    '--SCTR SALUD
    RstCab("sctrsalud") = NulosN(lbl_cod(1).Caption)
    '--SCTR PENSION
    RstCab("sctrpension") = NulosN(lbl_cod(2).Caption)
    
    '--
    RstCab.Update

    MsgBox "Los datos del Personal de Terceros " + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

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
