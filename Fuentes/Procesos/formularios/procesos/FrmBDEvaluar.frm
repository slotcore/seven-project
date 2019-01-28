VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBDEvaluar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herramientas - Evaluar Base de Datos"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2310
      TabIndex        =   16
      Top             =   3165
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   17
         Top             =   345
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   3
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   -60
         X2              =   6360
         Y1              =   675
         Y2              =   690
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5745
         X2              =   5745
         Y1              =   -90
         Y2              =   4800
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Index           =   2
         Left            =   4140
         TabIndex        =   20
         Top             =   75
         Width           =   1530
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando:"
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
         Index           =   0
         Left            =   90
         TabIndex        =   19
         Top             =   75
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base de Datos"
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
         Index           =   1
         Left            =   1185
         TabIndex        =   18
         Top             =   75
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4860
         Top             =   90
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
               Picture         =   "FrmBDEvaluar.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBDEvaluar.frx":2A98
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fra 
      Caption         =   "[ Destino ]"
      Height          =   630
      Index           =   1
      Left            =   30
      TabIndex        =   7
      Top             =   945
      Width           =   8025
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   1
         Left            =   1950
         Picture         =   "FrmBDEvaluar.frx":2E2A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   255
         Width           =   225
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   1
         Left            =   735
         MaxLength       =   12
         TabIndex        =   9
         Text            =   "txt_cb(1)"
         Top             =   225
         Width           =   1470
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
         Height          =   300
         Index           =   1
         Left            =   5055
         TabIndex        =   13
         Top             =   270
         Visible         =   0   'False
         Width           =   1185
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
         Left            =   2190
         TabIndex        =   12
         Top             =   225
         Width           =   4440
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   330
         Width           =   615
      End
      Begin VB.Label lbl_cb1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb1(1)"
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
         Left            =   6660
         TabIndex        =   10
         Top             =   225
         Width           =   1305
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   4335
      Left            =   30
      TabIndex        =   14
      Top             =   1605
      Width           =   9780
      _cx             =   17251
      _cy             =   7646
      _ConvInfo       =   1
      Appearance      =   1
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
      Rows            =   2
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmBDEvaluar.frx":2F5C
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
      Caption         =   "[ Origen ]"
      Height          =   630
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   8025
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   0
         Left            =   1950
         Picture         =   "FrmBDEvaluar.frx":2F85
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   255
         Width           =   225
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   735
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "txt_cb(0)"
         Top             =   225
         Width           =   1470
      End
      Begin VB.Label lbl_cb1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb1(0)"
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
         Left            =   6645
         TabIndex        =   6
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   330
         Width           =   615
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
         Height          =   300
         Index           =   0
         Left            =   5040
         TabIndex        =   4
         Top             =   225
         Visible         =   0   'False
         Width           =   1185
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
         Left            =   2190
         TabIndex        =   3
         Top             =   225
         Width           =   4440
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Seleccionar ]"
      Height          =   1260
      Left            =   8085
      TabIndex        =   21
      Top             =   330
      Width           =   1725
      Begin VB.OptionButton OptTipo 
         Caption         =   "Pendiente"
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   23
         Top             =   750
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Toda la Lista"
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   22
         Top             =   390
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmBDEvaluar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConexOri As New ADODB.Connection
Dim ConexDest As New ADODB.Connection

Dim BAND_INTERRUMPIR As Boolean

Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nTitulo As String
    On Error GoTo error
    '--generar la conexion a la base de datos principal
    Dim CnnTmp As New ADODB.Connection '--Conexion Temporal
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    '-----
    '--si la base de datos principal existe
    If ArchivoExiste(AP_RUTABD + "data.mdb") = False Then
        MsgBox "No existe la ruta a la Base de Datos Principal", vbCritical, "Mensaje..."
        Exit Sub
    End If
    
    '--ABRIENDO LA CONEXION PARA OBTENER EL LISTADO DE LOS AÑOS
    OPEN_CONEX_TMP CnnTmp, AP_RUTABD + "data.mdb"
    If CnnTmp.State = 0 Then GoTo SALIR
    '--definir la consulta
    nSQL = "SELECT mae_empresa.numruc, mae_empresa.nomemp as descripcion ,mae_empresa.id, mae_empresa.anotra, mae_empresa.ruta, 'Periodo: ' & [mae_empresa].[anotra] AS periodo " _
        + vbCr + " FROM mae_empresa " _
        + vbCr + " WHERE (((mae_empresa.activo)=-1)) "
    
    '-------------------------------------------------------------------------------------------
    Select Case Index
        Case 0 '--origen
            nTitulo = "Buscando Base de Datos Origen"
        Case 1 '--destino
            If NulosN(lbl_cod(0).Caption) = 0 Then
                MsgBox "Primero seleccione la Base de Datos Origen", vbExclamation, xTitulo
                txt_cb(0).SetFocus
                Exit Sub
            End If
            nSQL = nSQL & " and  mae_empresa.id <> " & NulosN(lbl_cod(0).Caption)
            nTitulo = "Buscando Base de Datos Destino"
    End Select
    
    ReDim xCampos(3, 3) As String
    xCampos(0, 0) = "Nº Ruc":       xCampos(0, 1) = "numruc":       xCampos(0, 2) = "1400":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "4500":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Año":          xCampos(2, 1) = "anotra":       xCampos(2, 2) = "500":    xCampos(2, 3) = "N"
    
    
    CARGAR_DLL_EPSBUSCAR CnnTmp, RstTmp, nSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio

    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index).Text = NulosC(RstTmp.Fields("numruc"))   '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = NulosC(RstTmp.Fields("descripcion")) '--NOMBRE
    lbl_cod(Index).Caption = NulosN(RstTmp.Fields("id")) '--CODIGO
    lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields("descripcion"))  '--NOMBRE
    
    lbl_cb1(Index).Caption = NulosC(RstTmp.Fields("periodo"))  '--NOMBRE
    
    '--si la base de datos principal existe
    If ArchivoExiste(AP_RUTABD & NulosC(RstTmp.Fields("ruta"))) = False Then
        MsgBox "No existe la ruta a la Base de Datos " & fra(Index).Caption, vbCritical, xTitulo
        txt_cb(Index).Text = ""
        txt_cb(Index).SetFocus
        Exit Sub
    End If
            
    '--cargar las conexiones
    Select Case Index
        Case 0 '--origen
                        
            OPEN_CONEX_TMP ConexOri, AP_RUTABD & NulosC(RstTmp.Fields("ruta"))
            If ConexOri.State = 0 Then
                MsgBox "Error en la Conexión", vbCritical, xTitulo
                txt_cb(0).Text = ""
                Exit Sub
            End If
        
            txt_cb(1).SetFocus
        Case 1 '--destino
            OPEN_CONEX_TMP ConexDest, AP_RUTABD & NulosC(RstTmp.Fields("ruta"))
            If ConexDest.State = 0 Then
                MsgBox "Error en la Conexión", vbCritical, xTitulo
                txt_cb(0).Text = ""
                Exit Sub
            End If
            
    End Select
SALIR:
    CnnTmp.Close
    Set RstTmp = Nothing
Exit Sub
error:
'    CnnTmp.Close
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    End If
End Sub

Private Sub Form_Load()

    CentrarFrm Me
    LimpiaText txt_cb
    LimpiaText lbl_cod
    LimpiaText lbl_cb1
    
    pConfigurarGrilla
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BAND_INTERRUMPIR = True '--interrumpir
    
    If ConexOri.State = 1 Then ConexOri.Close
    If ConexDest.State = 1 Then ConexDest.Close
    
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
        lbl_cb1(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index <> 1 Then
            SendKeys vbTab
        Else
            If Fg1.Rows >= 2 Then
                Fg1.Row = 1: Fg1.Col = 1
            Else
                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 1
            End If
            Fg1.SetFocus
        End If
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 5 Then pExportarExcel
    If Button.Index = 6 Then pImprimir
    If Button.Index = 8 Then
        BAND_INTERRUMPIR = True
        Unload Me
    End If
End Sub


'------------

Private Sub pConsultar()
    On Error GoTo error
    
    Dim rstOri As New ADODB.Recordset
    Dim rstDest As New ADODB.Recordset
    
    Dim rstOriCampo As New ADODB.Recordset
    Dim rstDestCampo As New ADODB.Recordset
    
    Dim mCampo&
    
    Dim nSQL As String
    
    Fg1.Rows = Fg1.FixedRows
    
    '*********************************************************************
    If NulosN(lbl_cod(0).Caption) = 0 Then
        MsgBox "Falta especificar la Base de Datos Origen", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Sub
    ElseIf NulosN(lbl_cod(1).Caption) = 0 Then
        MsgBox "Falta especificar la Base de Datos Origen", vbExclamation, xTitulo
        txt_cb(1).SetFocus
        Exit Sub
    End If
    '*********************************************************************
    
    '--establecer la consulta
    nSQL = "SELECT MSysObjects.Name as descripcion FROM MSysObjects WHERE (((MSysObjects.Type)=1) AND ((MSysObjects.Flags)=0)) OR (((MSysObjects.Database) Is Not Null)) ORDER BY MSysObjects.Name;"
    
    RST_Busq rstOri, nSQL, ConexOri '--lista de tablas origen
    RST_Busq rstDest, nSQL, ConexDest '--lista de tablas destino
    
    BAND_INTERRUMPIR = False
    If rstOri.RecordCount = 0 Then
        MsgBox "No hay registros en la Base Origen", vbInformation, xTitulo
        Exit Sub
    End If
    PosicionarProgBar
    PgBar.Min = 0
    PgBar.Max = rstOri.RecordCount
    
    '*********************************************************************
    Do While Not rstOri.EOF
        DoEvents
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        
        PgBar.Value = CLng(rstOri.Bookmark)
        
        Fg1.Rows = Fg1.Rows + 1
        
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = rstOri("descripcion")
        
        rstDest.MoveFirst
        rstDest.Find "descripcion='" & rstOri("descripcion") & "'"
        
        If rstDest.EOF = False And rstDest.BOF = False Then
            
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = rstDest("descripcion")
            
            nSQL = "select TOP 1 * FROM " & rstOri("descripcion")
            RST_Busq rstOriCampo, nSQL, ConexOri  '--lista de campos origen
            RST_Busq rstDestCampo, nSQL, ConexDest  '--lista de campos destino
            
            For mCampo = 0 To rstOriCampo.Fields.Count - 1
            
                DoEvents
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = rstOriCampo.Fields(mCampo).Name
                
                If RstRegistroBuscaCampo(rstDestCampo, rstOriCampo.Fields(mCampo).Name) = True Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 5) = rstOriCampo.Fields(mCampo).Name

                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = "!!!!!!OK"
                
                    If rstOriCampo(mCampo).Type <> rstDestCampo(rstOriCampo.Fields(mCampo).Name).Type Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 6) = "TIPO DATO!!!!!!"
                    ElseIf rstOriCampo(mCampo).DefinedSize <> rstDestCampo(rstOriCampo.Fields(mCampo).Name).DefinedSize Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 6) = "LONGITUD!!!!!!"
                    Else
                        '--eliminando las filas cuando tipo sea pendientes
                        If OptTipo(1).Value = True Then
                            Fg1.Rows = Fg1.Rows - 1
                        End If
                    End If
                    
                Else
                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = "FALTA!!!!!!"
                End If
                
            Next mCampo
            
        Else
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = "FALTA!!!!!!"
        
        End If
        '--eliminando las filas cuando tipo sea pendientes
        If Fg1.TextMatrix(Fg1.Rows - 1, 6) = "" And OptTipo(1).Value = True Then
            Fg1.Rows = Fg1.Rows - 1
        End If
        
        rstOri.MoveNext
    Loop
    '*********************************************************************
   '
SALIR:
    FraProgreso.Visible = False
    Set rstOri = Nothing:       Set rstOriCampo = Nothing
    Set rstDest = Nothing:      Set rstDestCampo = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
'Resume
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    
    Set rstOri = Nothing:       Set rstOriCampo = Nothing
    Set rstDest = Nothing:      Set rstDestCampo = Nothing
    
    SHOW_ERROR Me.Name, "pConsultar"
    
End Sub

Private Sub pExportarExcel()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "EVALUAR BASE DE DATOS", "Origen: " & lbl_cb(0).Caption & " " & lbl_cb1(0).Caption, "Destino: " & lbl_cb(1).Caption & " " & lbl_cb1(1).Caption, "Evaluar Base de Datos"
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub


Private Sub pImprimir()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, "EVALUAR BASE DE DATOS", "Origen: " & lbl_cb(0).Caption & " " & lbl_cb1(0).Caption, "Destino: " & lbl_cb(1).Caption & " " & lbl_cb1(1).Caption, True, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"

End Sub

Private Sub pConfigurarGrilla()
    '===================================================================================================
    'Propósito: Establecer los encabezados del grid
    '
    'Entradas:  Ninguna
    '
    'Resultados: Grilla con Encabezado
    '===================================================================================================
    
    With Fg1
        '-----
        .Cols = 7
                 
        '.FrozenCols = Q_POS_MES_INICIO - 1
        .ColWidth(0) = 200
        .FrozenCols = 0
        .Rows = 2
        .FixedRows = 2
        
        UNIR_CELDAS Fg1, 0, 1, 0, 2, "Origen", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 4, 0, 5, "Destino", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 6, 0, 6, "Observación", flexAlignCenterCenter
        
        .TextMatrix(1, 1) = "Tabla":    .ColWidth(1) = 2500:   .ColAlignment(1) = flexAlignLeftBottom:        .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 2) = "Campo":    .ColWidth(2) = 1500:   .ColAlignment(2) = flexAlignLeftBottom:        .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(1, 3) = " ":        .ColWidth(3) = 150:
        
        .TextMatrix(1, 4) = "Tabla":    .ColWidth(4) = 2500:   .ColAlignment(4) = flexAlignLeftBottom:        .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(1, 5) = "Campo":    .ColWidth(5) = 1500:   .ColAlignment(5) = flexAlignLeftBottom:        .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftBottom
        
        .TextMatrix(1, 6) = " ":      .ColWidth(6) = 1000:   .ColAlignment(6) = flexAlignLeftBottom:        .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftBottom
                
    End With
    DoEvents
End Sub

Private Sub PosicionarProgBar()
    '--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
    '    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    Me.PgBar.Value = 1
    FraProgreso.Visible = True
End Sub
