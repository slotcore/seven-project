VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmAjusteDifCambioOrdDpcho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herramientas - Ajuste po Diferencia de Cambio  Orden de Despacho"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Seleccionar "
      Height          =   600
      Left            =   4500
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   7065
      Begin VB.TextBox TxtAjuste 
         Height          =   300
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "TxtAjuste"
         Top             =   240
         Width           =   5535
      End
      Begin VB.CommandButton CmdBusProv 
         Height          =   230
         Left            =   5340
         Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   270
         Width           =   210
      End
      Begin VB.Label LblIdMon 
         AutoSize        =   -1  'True
         Caption         =   "LblIdMon"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5580
         TabIndex        =   26
         Top             =   210
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label LblIdAjuste 
         AutoSize        =   -1  'True
         Caption         =   "LblIdAjuste"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2070
         TabIndex        =   25
         Top             =   90
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resumen"
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
      Height          =   705
      Left            =   7500
      TabIndex        =   16
      Top             =   345
      Width           =   4365
      Begin VB.Label lblPer 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblPer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2940
         TabIndex        =   20
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lblGan 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblGan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   840
         TabIndex        =   19
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pérdida:"
         Height          =   195
         Left            =   2310
         TabIndex        =   18
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ganancia:"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   420
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Hasta el dia ]"
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
      Height          =   705
      Left            =   5460
      TabIndex        =   12
      Top             =   345
      Width           =   2025
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   315
         Left            =   630
         TabIndex        =   13
         Top             =   270
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
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   315
         Width           =   450
      End
   End
   Begin VB.Frame fraBarra 
      BorderStyle     =   0  'None
      Caption         =   "FrmConsultaDiario"
      Height          =   780
      Left            =   2760
      TabIndex        =   8
      Top             =   3525
      Visible         =   0   'False
      Width           =   6285
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   150
         TabIndex        =   9
         Top             =   315
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   -15
         Y2              =   900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   6270
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   1
         X1              =   -75
         X2              =   6500
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   6270
         X2              =   6270
         Y1              =   -30
         Y2              =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Ordenes"
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
         Left            =   195
         TabIndex        =   11
         Top             =   75
         Width           =   1785
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   4605
         TabIndex        =   10
         Top             =   75
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Cliente"
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
      Height          =   705
      Left            =   30
      TabIndex        =   1
      Top             =   345
      Width           =   5415
      Begin VB.CheckBox chk_SelCliente 
         Height          =   195
         Left            =   1830
         TabIndex        =   15
         Top             =   30
         Width           =   195
      End
      Begin VB.CommandButton CmdBusCliPro 
         Enabled         =   0   'False
         Height          =   240
         Left            =   5130
         Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   330
         Width           =   210
      End
      Begin VB.TextBox TxtCliPro 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   645
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "TxtCliPro"
         Top             =   300
         Width           =   4725
      End
      Begin VB.Label LblIdCliPro 
         Caption         =   "LblIdCliPro"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4605
         TabIndex        =   4
         Top             =   135
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   390
         Width           =   480
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6465
      Left            =   30
      TabIndex        =   5
      Top             =   1080
      Width           =   11850
      _cx             =   20902
      _cy             =   11404
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
      FrontTabColor   =   14215660
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "      Detalle     "
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
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6045
         Left            =   45
         TabIndex        =   6
         Top             =   45
         Width           =   11760
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5925
            Left            =   30
            TabIndex        =   7
            Top             =   90
            Width           =   11685
            _cx             =   20611
            _cy             =   10451
            _ConvInfo       =   1
            Appearance      =   0
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   128
            ForeColorSel    =   16777215
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16777215
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
            Cols            =   13
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmAjusteDifCambioOrdDpcho.frx":0264
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
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":03EB
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":092F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":0CC1
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":0E45
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":1299
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":13B1
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":18F5
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":1E39
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":1F4D
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":2061
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":24B5
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":2621
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambioOrdDpcho.frx":2B69
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmAjusteDifCambioOrdDpcho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstCta As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE


Private Sub chk_SelCliente_Click()
    If chk_SelCliente.Value = 1 Then
        CmdBusCliPro.Enabled = True
    
    Else
        CmdBusCliPro.Enabled = False
        TxtCliPro.Text = ""
        LblIdCliPro.Caption = ""
    End If
End Sub

Private Sub CmdBusCliPro_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xform.Titulo = "Buscando Clientes"
    xform.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente ORDER BY mae_cliente.nombre"
    xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"

    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"

    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCliPro.Text = xRs("nombre")
        LblIdCliPro.Caption = xRs("id")
        TxtFecha.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        lblPer.Caption = "0.00"
        lblGan.Caption = "0.00"
        pConfigurarGrilla4
        LblIdMon.Caption = "2"
        SeEjecuto = True
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    ElseIf KeyCode = vbKeyF3 And Shift = 0 Then
        BuscarVSFlexGrid
    End If

End Sub

Private Sub Form_Load()
    TxtCliPro.Text = ""
    TxtFecha.Valor = ""
    TxtFecha.Valor = Date
    SeEjecuto = False

End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        If chk_SelCliente.Value = 1 And NulosN(LblIdCliPro.Caption) = 0 Then
            MsgBox "Seleccione el Cliente", vbExclamation, xTitulo
            Exit Sub
        End If
        CargarCli4 NulosN(LblIdCliPro.Caption)
    End If
    
    If Button.Index = 2 Then Grabar
    If Button.Index = 4 Then pExportar
    If Button.Index = 6 Then
        Set RstCta = Nothing
        Unload Me
    End If
End Sub

Sub pExportar()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    Dim nPeriodo As String
    Dim nTitulo1 As String
    
    nPeriodo = "Al  " + CStr(TxtFecha.Valor)

    nTitulo1 = "" '"(Expresado en " & LblMoneda.Caption & ")"
    
        
        GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "Cuenta Corriente - Orden de Despacho", " ", "Cuenta Corriente Análisis"


    
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportar"
End Sub

Private Sub TxtCliPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCliPro_Click
    End If
End Sub


Private Sub BuscarVSFlexGrid()
    On Error GoTo error
    
    Dim oExport As New SGI2_funciones.formularios
    Dim xCampos(4, 3) As String
    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
    xCampos(0, 0) = "Nº.Registro":      xCampos(0, 1) = "1":    xCampos(0, 2) = "C":    xCampos(0, 3) = "-1"
    xCampos(1, 0) = "Origen":           xCampos(1, 1) = "2":    xCampos(1, 2) = "C":    xCampos(1, 3) = "0"
    xCampos(2, 0) = "Nº Documento":     xCampos(2, 1) = "4":    xCampos(2, 2) = "C":    xCampos(2, 3) = "0"
    xCampos(3, 0) = Label1(0):          xCampos(3, 1) = "4":    xCampos(3, 2) = "C":    xCampos(3, 3) = "0"
    xCampos(4, 0) = "Fch.Emi.":         xCampos(4, 1) = "5":    xCampos(4, 2) = "F":    xCampos(4, 3) = "0"
    
    oExport.VSFlexGrid_Buscar Me.hWnd, Fg1, xCampos(), Fg1.Row
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "BuscarVSFlexGrid"
End Sub


Private Sub pImprimir()

    On Error GoTo error
    
        Dim oPrint As New SGI2_funciones.formularios
        Dim nPeriodo As String
        Dim nTitulo As String
        Dim nTitulo1 As String
        Dim nTipo As String
        nPeriodo = "Al  " + CStr(TxtFecha.Valor)
        nTitulo1 = " "
        nTipo = "Orden de Despacho"
        
    Me.MousePointer = vbHourglass
    
'        FrmPrinCtaCtaCli.Show
'        FrmPrinCtaCtaCli.SetFocus
        nTitulo = "Detalle de Cuenta Corriente - " + nTipo
        oPrint.Imprimir_x_VSFlexGrid Fg1, nTitulo, nTitulo1, nPeriodo, True, True

    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub


Private Function EvaluarCta(rst As ADODB.Recordset) As Boolean

    Select Case Mid(NulosC(rst("ctanum")), 1, 6)
        Case "121010" '--facturas por cobrar
            '--salir si es contrapartida de las letras
            If rst("idlib") = 37 Then Exit Function
            
            EvaluarCta = True
            
        Case "124030" '--lgd/lgc
            EvaluarCta = True
        Case "123010" '--letras
            EvaluarCta = True
        Case Else
            EvaluarCta = False
    End Select
End Function


Sub CargarCli4(IdCliPro)
    Dim rst As New ADODB.Recordset
    
    Dim rstVta As New ADODB.Recordset '--ventas
    Dim rstLgd As New ADODB.Recordset '--liquidacion de gasto debito
    Dim rstLet As New ADODB.Recordset '--para letras

    
    Dim A, B, xFila As Long
    Dim TotDebe, TotHaber As Double
    Dim TotGralDebe, TotGralHaber As Double
    
    Dim xNomPro As String
    Dim xNumOrden  As String
    
    Dim Cambio As Boolean
    Dim nSQL As String
    
    'On Error GoTo error

    BAND_INTERRUMPIR = False
    
    pConfigurarGrilla4
    '--------------------------
    fraBarra.Left = 2798
    fraBarra.Top = 2925
    
    ProgressBar1.Max = 1
    ProgressBar1.Value = 0
    fraBarra.Visible = True
    fraBarra.Refresh
    
    
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    
    Dim nSQLWhere As String
    Dim nCampoMuestra As String '--indica el campo que se mostrara esta en funcion de la moneda seleccionada
    nSQLWhere = ""
    
    '***************************************************
    If IdCliPro <> 0 Then
        nSQLWhere = " and vta_ventas.idcli = " & IdCliPro & " "
        nSQL = "SELECT vta_ventas.idcli, [mae_cliente]![numruc], [mae_cliente]![nombre], vta_ventas.numerodocref as numorden " _
            + vbCr + " FROM (vta_ventas LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) INNER JOIN con_diario ON vta_ventas.numerodocref = con_diario.rnumerodoc1 " _
            + vbCr + " WHERE (((vta_ventas.anulado)=0) AND ((con_diario.idlib)=37)) " & nSQLWhere _
            + vbCr + " GROUP BY vta_ventas.idcli, [mae_cliente]![numruc],[mae_cliente]![nombre], vta_ventas.numerodocref " _
            + vbCr + " HAVING (((vta_ventas.numerodocref)<>'')) " _
            + vbCr + " ORDER BY [mae_cliente]![nombre], vta_ventas.numerodocref; "
    Else
        nSQL = "SELECT vta_ventas.numerodocref AS numorden  FROM vta_ventas INNER JOIN con_diario ON vta_ventas.numerodocref = con_diario.rnumerodoc1 " _
            + vbCr + " WHERE (((vta_ventas.anulado)=0) AND ((con_diario.idlib)=37))  GROUP BY vta_ventas.numerodocref  " _
            + vbCr + " HAVING (((vta_ventas.numerodocref)<>'')) ORDER BY vta_ventas.numerodocref; "
    End If
    
    RST_Busq rst, nSQL, xCon
    '***************************************************
    '--filtrar lo que se va mostrar
    
    If rst.RecordCount = 0 Then
        MsgBox "No hay Ordenes de Despacho del cliente seleccionado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fraBarra.Visible = False
        Set rst = Nothing
        Exit Sub
    End If
    
    '***************************************************
    '--mostrar el detalle del diario
''    nSQL = "SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, con_diario.rfchope, mae_documento.abrev, con_diario.rnumerodoc, mae_moneda.simbolo, IIf([con_diario].[idlib] In (3,6,44),[con_diario].[tc],IIf([con_tc].[impven] Is Null,0,[con_tc].[impven])) AS tipcam, IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS impreal, IIf(con_diario.idmon=1,impreal,impreal*tipcam) AS imptotsol, IIf(con_diario.idmon=2,impreal,impreal/tipcam) AS imptotdol, con_diario.rnumerodoc AS numdoc2, con_diario.rglosaope, con_diario.rnumerodoc1 as numorden, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, con_diario.idlib, con_diario.idcue " _
''        + vbCr + " FROM ((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN con_planctas ON con_diario.idcue = con_planctas.id " _
''        + vbCr + " WHERE (((con_diario.rnumerodoc1) Is Not Null And (con_diario.rnumerodoc1)<>'') AND ((con_diario.idlib) In (2,6,37,41,42))); "


    '--ventas

    nSQL = "SELECT con_planctas.tipsal, Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, mae_libros.descripcion AS libdesc, con_diario.ridtipper, " _
         + vbCr + " IIf([con_diario].[ridtipper]=1,[mae_prov].[numruc],IIf([con_diario].[ridtipper]=2,[mae_cliente].[numruc],'')) AS numruc, IIf([con_diario].[ridtipper]=1,[mae_prov].[nombre],IIf([con_diario].[ridtipper]=2,[mae_cliente].[nombre],'')) AS apenom, " _
        + vbCr + " con_diario.rregistro AS registroref,mae_documento.abrev, con_diario.rfchope AS fchdoc, con_diario.rnumerodoc AS numdoc, con_diario.rglosa AS glosa, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
        + vbCr + " IIf(con_diario.idlib In (3,6,44),con_diario.tc,IIf(con_tc.impven Is Null,0,con_tc.impven)) AS tipcam, " _
        + vbCr + " IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS impreal, " _
        + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol, " _
        + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol, " _
        + vbCr + " IIf(con_planctas.tipsal='D' Or con_planctas.tipsal='',impdebesol-imphabersol,imphabersol-impdebesol) AS impsalsol, " _
        + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
        + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) AS imphaberdol, " _
        + vbCr + " IIf(con_planctas.tipsal='D' Or con_planctas.tipsal='',impdebedol-imphaberdol,imphaberdol-impdebedol) AS impsaldol, " _
        + vbCr + " con_diario.rnumerodoc1 AS numorden,con_diario.idlib, con_diario.ridlib, con_diario.idcue, con_diario.ridper, con_diario.idmov, con_diario.idmes, con_diario.rtipdoc " _
        + vbCr + " FROM mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper " _
        + vbCr + " WHERE  con_diario.rtipdoc<>7 and con_diario.idmon=2 and con_planctas.cuenta Like '1210%' and con_diario.idlib in (2) and  (((con_diario.rnumerodoc1) Is Not Null And (con_diario.rnumerodoc1)<>'') AND ((IIf([con_diario].[idmon]=1,([con_diario].[impdebsol]+[con_diario].[imphabsol]),([con_diario].[impdebdol]+[con_diario].[imphabdol])))<>0)) " _
        + vbCr + "  "
        
        'ORDER BY con_planctas.cuenta;
'    RST_Busq rstVta, nSQL, xCon

    
    '--lgd
    
    nSQL = nSQL & vbCr & "union" & vbCr & "SELECT con_planctas.tipsal, Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, mae_libros.descripcion AS libdesc, con_diario.ridtipper, " _
         + vbCr + " IIf([con_diario].[ridtipper]=1,[mae_prov].[numruc],IIf([con_diario].[ridtipper]=2,[mae_cliente].[numruc],'')) AS numruc, IIf([con_diario].[ridtipper]=1,[mae_prov].[nombre],IIf([con_diario].[ridtipper]=2,[mae_cliente].[nombre],'')) AS apenom, " _
        + vbCr + " con_diario.rregistro AS registroref,mae_documento.abrev, con_diario.fchdoc, con_diario.rnumerodoc AS numdoc, con_diario.rglosa AS glosa, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
        + vbCr + " IIf(con_diario.idlib In (3,6,44),con_diario.tc,IIf(con_tc.impven Is Null,0,con_tc.impven)) AS tipcam, " _
        + vbCr + " IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS impreal, " _
        + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol, " _
        + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol, " _
        + vbCr + " IIf(con_planctas.tipsal='D' Or con_planctas.tipsal='',impdebesol-imphabersol,imphabersol-impdebesol) AS impsalsol, " _
        + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
        + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) AS imphaberdol, " _
        + vbCr + " IIf(con_planctas.tipsal='D' Or con_planctas.tipsal='',impdebedol-imphaberdol,imphaberdol-impdebedol) AS impsaldol, " _
        + vbCr + " con_diario.rnumerodoc1 AS numorden,con_diario.idlib, con_diario.ridlib, con_diario.idcue, con_diario.ridper, con_diario.idmov, con_diario.idmes, con_diario.rtipdoc " _
        + vbCr + " FROM mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper " _
        + vbCr + " WHERE con_diario.idmon=2 and con_planctas.cuenta Like '124030%' and con_diario.idlib in (41) and  (((con_diario.rnumerodoc1) Is Not Null And (con_diario.rnumerodoc1)<>'') AND ((IIf([con_diario].[idmon]=1,([con_diario].[impdebsol]+[con_diario].[imphabsol]),([con_diario].[impdebdol]+[con_diario].[imphabdol])))<>0)) " _
        + vbCr + "  "
        'ORDER BY con_planctas.cuenta;
    RST_Busq rstVta, nSQL, xCon

    '--letras
    
    nSQL = "SELECT con_planctas.tipsal, Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, mae_libros.descripcion AS libdesc, con_diario.ridtipper, " _
         + vbCr + " IIf([con_diario].[ridtipper]=1,[mae_prov].[numruc],IIf([con_diario].[ridtipper]=2,[mae_cliente].[numruc],'')) AS numruc, IIf([con_diario].[ridtipper]=1,[mae_prov].[nombre],IIf([con_diario].[ridtipper]=2,[mae_cliente].[nombre],'')) AS apenom, " _
        + vbCr + " con_diario.rregistro AS registroref,mae_documento.abrev, con_diario.rfchope AS fchdoc, con_diario.rnumerodoc AS numdoc, con_diario.rglosa AS glosa, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
        + vbCr + " IIf(con_diario.idlib In (3,6,44),con_diario.tc,IIf(con_tc.impven Is Null,0,con_tc.impven)) AS tipcam, " _
        + vbCr + " IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS impreal, " _
        + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol, " _
        + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol, " _
        + vbCr + " IIf(con_planctas.tipsal='D' Or con_planctas.tipsal='',impdebesol-imphabersol,imphabersol-impdebesol) AS impsalsol, " _
        + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
        + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) AS imphaberdol, " _
        + vbCr + " IIf(con_planctas.tipsal='D' Or con_planctas.tipsal='',impdebedol-imphaberdol,imphaberdol-impdebedol) AS impsaldol, " _
        + vbCr + " con_diario.rnumerodoc1 AS numorden,con_diario.idlib, con_diario.ridlib, con_diario.idcue, con_diario.ridper, con_diario.idmov, con_diario.idmes, con_diario.rtipdoc " _
        + vbCr + " FROM mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper " _
        + vbCr + " WHERE con_diario.idmon=2 and con_planctas.cuenta Like '1230%'  and con_diario.idlib in (37) and  (con_diario.rnumerodoc1 Is Not Null And con_diario.rnumerodoc1<>'')  " _
        + vbCr + " ORDER BY con_planctas.cuenta; "


    RST_Busq rstLet, nSQL, xCon
    
    '***************************************************
    
    ProgressBar1.Max = rst.RecordCount
    
    Dim xSaldoDoc As Double
    Dim xFilaIni&
    Dim xColor&
    Dim ArrTotales(5) As Double
    Dim ImpVta As Double
    Dim ImpSaldoLetra As Double
    Dim numletra As String
    Dim ImpLetEvaluar  As Double
    
    Me.MousePointer = vbHourglass
     
    xColor = 0
    
    '--agregando 1ra fila
    rst.MoveFirst

    xNumOrden = ""
    
    Do While Not rst.EOF
        DoEvents
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        ProgressBar1.Value = rst.Bookmark
        

        rstVta.Filter = "numorden= '" & NulosC(rst("numorden")) & "'"
        rstLgd.Filter = "numorden= '" & NulosC(rst("numorden")) & "'"
        rstLet.Filter = "numorden= '" & NulosC(rst("numorden")) & "'"
        
        
        Erase ArrTotales
        
        If rstVta.RecordCount <> 0 Then
            If Fg1.Rows = Fg1.FixedRows Then
                Fg1.Rows = Fg1.Rows + 1
            Else
                Fg1.Rows = Fg1.Rows + 2
            End If
            
            xFila = Fg1.Rows - 1
            'Fg1.TextMatrix(xFila, 1) = "DESCRIPCIÓN DEL NÚMERO DE ORDEN:  " & NulosC(rstVta("numorden"))
            
            ''''GRID_COMBINAR Fg1, xFila, 1, xFila, 7, "DESCRIPCIÓN DEL NÚMERO DE ORDEN:  " & NulosC(rstVta("numorden")), flexAlignCenterCenter, True, , , , True
            
            GRID_COMBINAR Fg1, xFila, 1, xFila, 9, "DESCRIPCIÓN DEL NÚMERO DE ORDEN:  " & NulosC(rstVta("numorden")), flexAlignLeftCenter, , , &H800000, vbWhite
            xFilaIni = xFila + 1
            GRID_COLOR_FONDO Fg1, xFila, 1, xFila, Fg1.Cols - 1, &H80000005
            'GRID_COLOR_FONDO Fg1, xFila, 1, xFila, Fg1.Cols - 1, &HE0DCDA
                    
            rstVta.MoveFirst
            rstVta.Sort = "fchdoc"
            
            '***************************************************************
            '--obteniendo el primer saldo de la letra que sera igual al importe de la letra
            ImpSaldoLetra = 0
            numletra = ""
            Do While Not rstVta.EOF
                '--agregando los detalles
                Fg1.Rows = Fg1.Rows + 1
                xFila = Fg1.Rows - 1
                '----
                '--agregando las ventas
                Fg1.TextMatrix(xFila, 1) = NulosC(rstVta("registro"))
                Fg1.TextMatrix(xFila, 2) = NulosC(rstVta("numruc"))
                Fg1.TextMatrix(xFila, 3) = NulosC(rstVta("apenom"))
                Fg1.TextMatrix(xFila, 4) = NulosC(rstVta("abrev"))
                Fg1.TextMatrix(xFila, 5) = NulosC(rstVta("numdoc"))
                Fg1.TextMatrix(xFila, 6) = Format(rstVta("fchdoc"), FORMAT_DATE)
                Fg1.TextMatrix(xFila, 7) = NulosC(rstVta("simbolo"))
                Fg1.TextMatrix(xFila, 8) = Format(NulosN(rstVta("tipcam")), "###0.##0") & ""
                Fg1.TextMatrix(xFila, 9) = Format(NulosN(rstVta("impreal")), FORMAT_MONTO)
                
                ImpVta = NulosN(rstVta("impreal"))
                '--agregando la letra
                
                If rstLet.RecordCount <> 0 Then
                    Do While Not rstLet.EOF
                        Fg1.TextMatrix(xFila, 10) = NulosC(rstLet("registro"))
                        Fg1.TextMatrix(xFila, 11) = NulosC(rstLet("numruc"))
                        Fg1.TextMatrix(xFila, 12) = NulosC(rstLet("apenom"))
                        Fg1.TextMatrix(xFila, 13) = NulosC(rstLet("abrev"))
                        Fg1.TextMatrix(xFila, 14) = NulosC(rstLet("numdoc"))
                        Fg1.TextMatrix(xFila, 15) = Format(rstLet("fchdoc"), FORMAT_DATE)
                        Fg1.TextMatrix(xFila, 16) = NulosC(rstLet("simbolo"))
                        Fg1.TextMatrix(xFila, 17) = Format(NulosN(rstLet("tipcam")), "###0.##0") & ""
                        Fg1.TextMatrix(xFila, 18) = Format(NulosN(rstLet("impreal")), FORMAT_MONTO)
                        
                        If numletra <> NulosC(rstLet("numdoc")) Then
                            ImpLetEvaluar = NulosN(rstLet("impreal"))
                        Else
                            ImpLetEvaluar = 0
                        End If
                        
                        If ImpVta < ImpSaldoLetra + ImpLetEvaluar Then
                                                    
                            Fg1.TextMatrix(xFila, 19) = Format(ImpVta, FORMAT_MONTO)
                            
                            If numletra <> NulosC(rstLet("numdoc")) Then
                                ImpSaldoLetra = ImpSaldoLetra + NulosN(rstLet("impreal")) - ImpVta
                                numletra = NulosC(rstLet("numdoc"))
                            Else
                                ImpSaldoLetra = ImpSaldoLetra - ImpVta
                            End If
                            
                            
                            Fg1.TextMatrix(xFila, 20) = Format(ImpSaldoLetra, FORMAT_MONTO)
                            
                            '---empezar a evaluar los demas datos !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!11
                            
                            '--cta debe
                            Fg1.TextMatrix(xFila, 21) = NulosC(rstVta("ctanum"))
                            Fg1.TextMatrix(xFila, 22) = NulosC(rstVta("ctadesc"))
                            Fg1.TextMatrix(xFila, 23) = Format(NulosN(rstVta("impreal")) * NulosN(rstVta("tipcam")), FORMAT_MONTO)
                            
                            '--cta haber
                            Fg1.TextMatrix(xFila, 24) = NulosC(rstLet("ctanum"))
                            Fg1.TextMatrix(xFila, 25) = NulosC(rstLet("ctadesc"))
                            Fg1.TextMatrix(xFila, 26) = Format(NulosN(Fg1.TextMatrix(xFila, 19)) * NulosN(rstLet("tipcam")), FORMAT_MONTO)
                            
                            '--resumen (haber - debe)
                             
                            If CDate(Fg1.TextMatrix(xFila, 15)) > CDate(Fg1.TextMatrix(xFila, 6)) Then
                                Fg1.TextMatrix(xFila, 27) = Format(Fg1.TextMatrix(xFila, 26) - Fg1.TextMatrix(xFila, 23), FORMAT_MONTO)
                            Else
                                Fg1.TextMatrix(xFila, 27) = Format(Fg1.TextMatrix(xFila, 23) - Fg1.TextMatrix(xFila, 26), FORMAT_MONTO)
                            End If
                                                        
                            
                            If NulosN(Fg1.TextMatrix(xFila, 27)) > 0 Then '--
                                Fg1.TextMatrix(xFila, 28) = "Si"
                                Fg1.TextMatrix(xFila, 29) = ""
                                
                                Fg1.TextMatrix(xFila, 30) = NulosN(Fg1.TextMatrix(xFila, 27))
                                'Fg1.TextMatrix(xFila, 31) = "0.00"
                            ElseIf NulosN(Fg1.TextMatrix(xFila, 27)) < 0 Then '--
                                Fg1.TextMatrix(xFila, 28) = ""
                                Fg1.TextMatrix(xFila, 29) = "Si"
                                
                                'Fg1.TextMatrix(xFila, 30) = "0.00"
                                Fg1.TextMatrix(xFila, 31) = Abs(NulosN(Fg1.TextMatrix(xFila, 27)))
                            End If
                            
                            If Abs(NulosN(Fg1.TextMatrix(xFila, 27))) >= 10000 Then
                                '--cambiar de color a la fila cuando el saldo supere a 10.000
                                '--considerara pagado en su totalidad
                                GRID_COLOR_FONDO Fg1, xFila, 1, xFila, Fg1.Cols - 1, &H9BFF79

                            End If
                            
                            '--codigo
                            Fg1.TextMatrix(xFila, 32) = NulosN(rstVta("idmov"))
                            Fg1.TextMatrix(xFila, 33) = NulosN(rstLet("idmov"))
                            Fg1.TextMatrix(xFila, 34) = NulosN(rstLet("idmes"))
                            
                            Fg1.TextMatrix(xFila, 35) = NulosN(rstVta("idcue"))
                            Fg1.TextMatrix(xFila, 36) = NulosN(rstLet("idcue"))
                            
                            Fg1.TextMatrix(xFila, 37) = NulosN(rstVta("ridper"))
                            Fg1.TextMatrix(xFila, 38) = NulosN(rstLet("ridper"))
                            
                            Fg1.TextMatrix(xFila, 39) = NulosN(rstVta("rtipdoc"))
                            Fg1.TextMatrix(xFila, 40) = NulosN(rstLet("rtipdoc"))
                            
                            Fg1.TextMatrix(xFila, 41) = NulosC(rstLet("numorden"))
                            
                            Exit Do
                        Else
                        
                            ImpSaldoLetra = ImpSaldoLetra + NulosN(rstLet("impreal"))
                            
                        End If
                        
                        
                        rstLet.MoveNext
                    Loop
                End If
                    
                    
                    
                
                rstVta.MoveNext
            Loop
        End If
        rst.MoveNext
    Loop
    
    Set rst = Nothing
    Set rstVta = Nothing
    DoEvents
    lblGan.Caption = Format(GRID_SUMAR_COL(Fg1, 30), FORMAT_MONTO)
    lblPer.Caption = Format(GRID_SUMAR_COL(Fg1, 31), FORMAT_MONTO)
    DoEvents
    
    fraBarra.Visible = False
    Me.MousePointer = vbDefault
    MsgBox "La Consulta fue se realizó Correctamente", vbInformation, xTitulo
    Exit Sub
SALIR:
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set rst = Nothing
    Set rstVta = Nothing
    Set rstLgd = Nothing
    Set rstLet = Nothing
    MsgBox "La Consulta fue Interrumpida", vbInformation, xTitulo
    Exit Sub
error:
    
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set rst = Nothing
    Set rstVta = Nothing
    Set rstLgd = Nothing
    Set rstLet = Nothing
    SHOW_ERROR Me.Name, "CargarCli2"
    
End Sub







Private Sub pConfigurarGrilla4()
    Dim A As Integer
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    With Fg1
        '-----
        .Rows = 2
        .Cols = 42
        .FixedRows = 2
        .FrozenCols = 0
        .RowHeight(0) = 250
        .ColWidth(0) = 200
        UNIR_CELDAS Fg1, 0, 1, 0, 9, "DATOS DEL DOCUMENTO", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 10, 0, 20, "DATOS DE LA LETRA", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 21, 0, 23, "CTA DEBE", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 24, 0, 26, "CTA HABER", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 27, 0, 31, "RESUMEN", flexAlignCenterCenter
        
        UNIR_CELDAS Fg1, 0, 32, 0, 40, "CODIGOS", flexAlignCenterCenter
        
        
        FORMATO_CELDA Fg1, 0, 1, vbBlack, True, &HD8E9EC
        FORMATO_CELDA Fg1, 0, 10, vbBlack, True, &HD8E9EC
        FORMATO_CELDA Fg1, 0, 21, vbBlack, True, &HD8E9EC
        FORMATO_CELDA Fg1, 0, 24, vbBlack, True, &HD8E9EC
        FORMATO_CELDA Fg1, 0, 27, vbBlack, True, &HD8E9EC
                
        '--ventas
        .TextMatrix(1, 1) = "N° Reg.":  .ColWidth(1) = 820:    .ColAlignment(1) = flexAlignLeftCenter:     .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 2) = "Num Ruc":      .ColWidth(2) = 1100:   .ColAlignment(2) = flexAlignLeftCenter:     .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "Razón Social": .ColWidth(3) = 1500:   .ColAlignment(3) = flexAlignLeftCenter:     .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 4) = "T.D.":         .ColWidth(4) = 400:    .ColAlignment(4) = flexAlignLeftCenter:     .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 5) = "N°.Documento": .ColWidth(5) = 1400:   .ColAlignment(5) = flexAlignLeftCenter:     .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 6) = "Fch.Doc":      .ColWidth(6) = 800:    .ColAlignment(6) = flexAlignCenterBottom:   .Row = 1: .Col = 6: .CellAlignment = flexAlignCenterBottom
        .TextMatrix(1, 7) = "M":            .ColWidth(7) = 450:    .ColAlignment(7) = flexAlignLeftCenter:     .Row = 1: .Col = 7: .CellAlignment = flexAlignCenterBottom
        .TextMatrix(1, 8) = "T.C.":         .ColWidth(8) = 500:    .ColAlignment(8) = flexAlignRightCenter:    .Row = 1: .Col = 8: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 9) = "Imp":          .ColWidth(9) = 900:    .ColAlignment(9) = flexAlignRightCenter: .Row = 1: .Col = 9: .CellAlignment = flexAlignRightCenter
        
        '--letras
        .TextMatrix(1, 10) = "N° Reg.":  .ColWidth(10) = 820:    .ColAlignment(10) = flexAlignLeftCenter:     .Row = 1: .Col = 10: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 11) = "Num Ruc":      .ColWidth(11) = 1100:   .ColAlignment(11) = flexAlignLeftCenter:     .Row = 1: .Col = 11: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 12) = "Razón Social": .ColWidth(12) = 1500:   .ColAlignment(12) = flexAlignLeftCenter:     .Row = 1: .Col = 12: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 13) = "T.D.":         .ColWidth(13) = 400:    .ColAlignment(13) = flexAlignLeftCenter:     .Row = 1: .Col = 13: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 14) = "N°.Documento": .ColWidth(14) = 1400:   .ColAlignment(14) = flexAlignLeftCenter:     .Row = 1: .Col = 14: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 15) = "Fch.Doc":      .ColWidth(15) = 800:    .ColAlignment(15) = flexAlignCenterBottom:   .Row = 1: .Col = 15: .CellAlignment = flexAlignCenterBottom
        .TextMatrix(1, 16) = "M":            .ColWidth(16) = 450:    .ColAlignment(16) = flexAlignLeftCenter:     .Row = 1: .Col = 16: .CellAlignment = flexAlignCenterBottom
        .TextMatrix(1, 17) = "T.C.":         .ColWidth(17) = 500:    .ColAlignment(17) = flexAlignRightCenter:    .Row = 1: .Col = 17: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 18) = "Imp":          .ColWidth(18) = 900:    .ColAlignment(18) = flexAlignRightCenter: .Row = 1: .Col = 18: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 19) = "ImpUtil":          .ColWidth(19) = 900:    .ColAlignment(19) = flexAlignRightCenter: .Row = 1: .Col = 19: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 20) = "Saldo":          .ColWidth(20) = 900:    .ColAlignment(20) = flexAlignRightCenter: .Row = 1: .Col = 20: .CellAlignment = flexAlignRightCenter
        
        '--cta debe
        .TextMatrix(1, 21) = "Num Cta":      .ColWidth(21) = 1000:  .ColAlignment(21) = flexAlignLeftCenter:   .Row = 1: .Col = 21: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 22) = "Nombre Cta":   .ColWidth(22) = 1700:  .ColAlignment(22) = flexAlignLeftCenter:   .Row = 1: .Col = 22: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 23) = "Debe":         .ColWidth(23) = 900:  .ColAlignment(23) = flexAlignRightCenter:   .Row = 1: .Col = 23: .CellAlignment = flexAlignRightCenter
        
        '--cta haber
        .TextMatrix(1, 24) = "Num Cta":      .ColWidth(24) = 1000:  .ColAlignment(24) = flexAlignLeftCenter:   .Row = 1: .Col = 24: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 25) = "Nombre Cta":   .ColWidth(25) = 1700:  .ColAlignment(25) = flexAlignLeftCenter:   .Row = 1: .Col = 25: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 26) = "Haber":        .ColWidth(26) = 900:  .ColAlignment(26) = flexAlignRightCenter:   .Row = 1: .Col = 26: .CellAlignment = flexAlignRightCenter
        
        '--resumen
        .TextMatrix(1, 27) = "Saldo":        .ColWidth(27) = 900:  .ColAlignment(27) = flexAlignRightCenter:   .Row = 1: .Col = 27: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 28) = "Gan":          .ColWidth(28) = 400:  .ColAlignment(28) = flexAlignCenterCenter:  .Row = 1: .Col = 28: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 29) = "Per":          .ColWidth(29) = 400:  .ColAlignment(29) = flexAlignCenterCenter:  .Row = 1: .Col = 29: .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(1, 30) = "Imp Gan":          .ColWidth(30) = 1100:  .ColAlignment(30) = flexAlignRightCenter:  .Row = 1: .Col = 30: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 31) = "Imp Per":          .ColWidth(31) = 1100:  .ColAlignment(31) = flexAlignRightCenter:  .Row = 1: .Col = 31: .CellAlignment = flexAlignCenterCenter
        
        
        '--codigos
        .TextMatrix(1, 32) = "iddoc":        .ColWidth(32) = 400:  .ColAlignment(32) = flexAlignRightCenter:   .Row = 1: .Col = 32: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 33) = "idlet":        .ColWidth(33) = 400:  .ColAlignment(33) = flexAlignRightCenter:   .Row = 1: .Col = 33: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 34) = "idmes":        .ColWidth(34) = 400:  .ColAlignment(34) = flexAlignRightCenter:   .Row = 1: .Col = 34: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 35) = "idcuedeb":     .ColWidth(35) = 400:  .ColAlignment(35) = flexAlignRightCenter:   .Row = 1: .Col = 35: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 36) = "idcuehab":     .ColWidth(36) = 400:  .ColAlignment(36) = flexAlignRightCenter:   .Row = 1: .Col = 36: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 37) = "idpervta":     .ColWidth(37) = 400:  .ColAlignment(37) = flexAlignRightCenter:   .Row = 1: .Col = 37: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 38) = "idperlet":     .ColWidth(37) = 400:  .ColAlignment(37) = flexAlignRightCenter:   .Row = 1: .Col = 38: .CellAlignment = flexAlignRightCenter
                       
        .TextMatrix(1, 39) = "idtdocvta":     .ColWidth(39) = 400:  .ColAlignment(39) = flexAlignRightCenter:   .Row = 1: .Col = 39: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 40) = "idtdoclet":     .ColWidth(40) = 400:  .ColAlignment(40) = flexAlignRightCenter:   .Row = 1: .Col = 40: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 41) = "Orden Despacho":     .ColWidth(41) = 1000:  .ColAlignment(41) = flexAlignRightCenter:   .Row = 1: .Col = 41: .CellAlignment = flexAlignRightCenter
                       
                       
        .SelectionMode = flexSelectionByRow
    End With
    
    
    TabOne1.CurrTab = 0
    DoEvents
End Sub


Function Grabar() As Boolean
    Dim A, B, Rpta As Integer
    Dim RstDia As New ADODB.Recordset
    Dim nSQL As String
    
    Dim xIdCuen, xId As Integer
    Dim xTotal As Double
    Dim xNumAsiento As String
    Dim xSaldo As Double '--indica el saldo actual del documento
    Dim mRow&
    Dim mIdMes As Integer
    
    Dim IdCtaGanancia As Long
    Dim IdCtaPerdida As Long
    Dim TipCam As Double
    
    Dim IdCtaDestDeb As Long
    Dim IdCtaDestHab As Long
    
    Dim sSaldo As Double
    
    Dim mIdMov As Double '--codigo de movimiento
    
    Dim rst As New ADODB.Recordset
    Err.Clear
    
    If Fg1.Rows = Fg1.FixedRows Then
        MsgBox "No hay datos para grabar", vbInformation, xTitulo
        Exit Function
    End If
    
    
    
    On Error GoTo LaCague
    
    RST_Busq rst, "SELECT mae_ajuste.* FROM mae_ajuste WHERE mae_ajuste.idmon=2 and mae_ajuste.idlib = 2 ;", xCon

    If rst.RecordCount = 0 Then
        MsgBox "Falta Configurar la Cuentas para el ajuste Ventas", vbInformation, xTitulo
        Set rst = Nothing
        Exit Function
    End If
    IdCtaGanancia = NulosN(rst("idcuengan"))
    IdCtaPerdida = NulosN(rst("idcuenper"))
    Set rst = Nothing
    
    
    
    nSQL = "SELECT con_planctas.ctadesdeb, con_planctas.ctadeshab FROM con_planctas WHERE (((con_planctas.id)=" & IdCtaPerdida & "));"
    RST_Busq rst, nSQL, xCon
    If rst.RecordCount <> 0 Then
        IdCtaDestDeb = NulosN(rst("ctadesdeb"))
        IdCtaDestHab = NulosN(rst("ctadeshab"))
    Else
        IdCtaDestDeb = 0
        IdCtaDestHab = 0
    End If
    Set rst = Nothing
    
    '--obteniendo el ultimo id de movimiento
    nSQL = "SELECT con_diario.idlib, Last(con_diario.idmov) AS idmov1 FROM con_diario GROUP BY con_diario.idlib HAVING (((con_diario.idlib)=44));"
    RST_Busq rst, nSQL, xCon
    If rst.RecordCount <> 0 Then
        mIdMov = NulosN(rst("idmov1"))
    Else
        mIdMov = 0
    End If
    Set rst = Nothing
    
    '-------------
    RST_Busq RstDia, "select top 1 * from con_diario", xCon
    
    xCon.BeginTrans
    
    xCon.Execute "delete from con_diario where idlib = 44 and ridlib in (2,37) "
    
    fraBarra.Visible = True
    ProgressBar1.Max = Fg1.Rows - 1
    ProgressBar1.Min = 0
    
    mIdMov = 1
    
    For mRow = 2 To Fg1.Rows - 1
        ProgressBar1.Value = mRow
        DoEvents
        
        '**************************************************************************************
        'If Fg1.TextMatrix(mRow, 28) <> "" Or Fg1.TextMatrix(mRow, 29) <> "" Then
        If LCase(Fg1.TextMatrix(mRow, 28)) = "si" Then
            mIdMov = mIdMov + 1
            
            mIdMes = Month(Fg1.TextMatrix(mRow, 34))
            sSaldo = NulosN(Fg1.TextMatrix(mRow, 27))
            xNumAsiento = NuevoNumAsiento(44, mIdMes, xCon)
            TipCam = NulosN(Fg1.TextMatrix(mRow, 17))
        
            
            '**************************************************************************************
            RstDia.AddNew
            RstDia("año") = AnoTra
            RstDia("idmes") = mIdMes
            RstDia("idlib") = 44
            RstDia("idmov") = mIdMov
            RstDia("numasi") = xNumAsiento
            RstDia("tc") = TipCam
           
            If LCase(Fg1.TextMatrix(mRow, 28)) = "si" Then
            
                 RstDia("idcue") = IdCtaGanancia
                 RstDia("impdebsol") = 0
                 RstDia("impdebdol") = 0
                 RstDia("imphabsol") = Abs(sSaldo)
                 RstDia("imphabdol") = Abs(sSaldo / TipCam)
                 
            Else
                 RstDia("idcue") = IdCtaPerdida
                 RstDia("impdebsol") = Abs(sSaldo)
                 RstDia("impdebdol") = Abs(sSaldo / TipCam)
                 RstDia("imphabsol") = 0
                 RstDia("imphabdol") = 0
            End If
             
            RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
            RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 15))
            RstDia("idmon") = 2
            RstDia("ridlib") = 2
            RstDia("iddocpro") = NulosN(Fg1.TextMatrix(mRow, 32))
            RstDia("ridtipper") = 2
            RstDia("ridper") = Fg1.TextMatrix(mRow, 37)
            RstDia("rtipdoc") = NulosN(Fg1.TextMatrix(mRow, 39))
            If IsDate(Fg1.TextMatrix(mRow, 6)) = True Then
                RstDia("rfchope") = CDate(Fg1.TextMatrix(mRow, 6))
            End If
            RstDia("rnumerodoc") = Fg1.TextMatrix(mRow, 5)
            RstDia("rregistro") = Fg1.TextMatrix(mRow, 1)
            RstDia("rglosa") = ""
            RstDia("ridmon") = 2
            RstDia("rtipdoc1") = 108
            RstDia("rnumerodoc1") = NulosC(Fg1.TextMatrix(mRow, 41))
            RstDia("ridtipper2") = 2
            RstDia("ridper2") = NulosN(Fg1.TextMatrix(mRow, 38))
            RstDia("rtipdoc2") = CDate(Fg1.TextMatrix(mRow, 40))
            RstDia("rfchope2") = Fg1.TextMatrix(mRow, 15)
            RstDia("rnumerodoc2") = Fg1.TextMatrix(mRow, 14)
            
            RstDia("ajuste") = 1
            RstDia.Update
            
        
            '**************************************************************************************
            RstDia.AddNew
            RstDia("año") = AnoTra
            RstDia("idmes") = mIdMes
            RstDia("idlib") = 44
            RstDia("idmov") = mIdMov
            RstDia("numasi") = xNumAsiento
            RstDia("tc") = TipCam
            
            RstDia("idcue") = NulosN(Fg1.TextMatrix(mRow, 35))
            
            RstDia("impdebsol") = Abs(sSaldo)
            RstDia("impdebdol") = Abs(sSaldo / TipCam)
            RstDia("imphabsol") = 0
            RstDia("imphabdol") = 0
             
            If LCase(Fg1.TextMatrix(mRow, 28)) = "si" Then
                 RstDia("impdebsol") = Abs(sSaldo)
                 RstDia("impdebdol") = Abs(sSaldo / TipCam)
                 RstDia("imphabsol") = 0
                 RstDia("imphabdol") = 0
                 
            Else
                 RstDia("impdebsol") = 0
                 RstDia("impdebdol") = 0
                 RstDia("imphabsol") = Abs(sSaldo)
                 RstDia("imphabdol") = Abs(sSaldo / TipCam)
            End If
             
            RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
            RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 15))
            RstDia("idmon") = 2
            RstDia("ridlib") = 2
            RstDia("iddocpro") = NulosN(Fg1.TextMatrix(mRow, 32))
            RstDia("ridtipper") = 2
            RstDia("ridper") = Fg1.TextMatrix(mRow, 37)
            RstDia("rtipdoc") = NulosN(Fg1.TextMatrix(mRow, 39))
            If IsDate(Fg1.TextMatrix(mRow, 6)) = True Then
                RstDia("rfchope") = CDate(Fg1.TextMatrix(mRow, 6))
            End If
            RstDia("rnumerodoc") = Fg1.TextMatrix(mRow, 5)
            RstDia("rregistro") = Fg1.TextMatrix(mRow, 1)
            RstDia("rglosa") = ""
            RstDia("ridmon") = 2
            RstDia("rtipdoc1") = 108
            RstDia("rnumerodoc1") = NulosC(Fg1.TextMatrix(mRow, 41))
            RstDia("ridtipper2") = 2
            RstDia("ridper2") = NulosN(Fg1.TextMatrix(mRow, 38))
            RstDia("rtipdoc2") = CDate(Fg1.TextMatrix(mRow, 40))
            RstDia("rfchope2") = Fg1.TextMatrix(mRow, 15)
            RstDia("rnumerodoc2") = Fg1.TextMatrix(mRow, 14)
            
            RstDia("ajuste") = 1
            
             RstDia.Update
            
            '**************************************************************************************
            
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            '--destinos de perdida
           
            If IdCtaDestDeb <> 0 And IdCtaDestHab <> 0 And LCase(Fg1.TextMatrix(mRow, 29)) = "si" Then
                '************************************************************************************************
                '--destinos automatico cta debe
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = mIdMes
                RstDia("idlib") = 44
                RstDia("idmov") = mIdMov
                RstDia("numasi") = xNumAsiento
                RstDia("tc") = TipCam
                
                RstDia("impdebsol") = 0
                RstDia("impdebdol") = 0
                RstDia("imphabsol") = 0
                RstDia("imphabdol") = 0
                RstDia("idcue") = IdCtaDestDeb
                If NulosN(LblIdMon.Caption) = 2 Then
                    RstDia("impdebdol") = Abs(sSaldo) / TipCam
                Else
                    RstDia("impdebsol") = Abs(sSaldo) * TipCam
                End If
                
                RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
                RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 15))
                RstDia("idmon") = 2
                RstDia("ridlib") = 2
                RstDia("iddocpro") = NulosN(Fg1.TextMatrix(mRow, 32))
                RstDia("ridtipper") = 2
                RstDia("ridper") = Fg1.TextMatrix(mRow, 37)
                RstDia("rtipdoc") = NulosN(Fg1.TextMatrix(mRow, 39))
                If IsDate(Fg1.TextMatrix(mRow, 6)) = True Then
                    RstDia("rfchope") = CDate(Fg1.TextMatrix(mRow, 6))
                End If
                RstDia("rnumerodoc") = Fg1.TextMatrix(mRow, 5)
                RstDia("rregistro") = Fg1.TextMatrix(mRow, 1)
                RstDia("rglosa") = ""
                RstDia("ridmon") = 2
                RstDia("rtipdoc1") = 108
                RstDia("rnumerodoc1") = NulosC(Fg1.TextMatrix(mRow, 41))
                RstDia("ridtipper2") = 2
                RstDia("ridper2") = NulosN(Fg1.TextMatrix(mRow, 38))
                RstDia("rtipdoc2") = CDate(Fg1.TextMatrix(mRow, 40))
                RstDia("rfchope2") = Fg1.TextMatrix(mRow, 15)
                RstDia("rnumerodoc2") = Fg1.TextMatrix(mRow, 14)
                
                RstDia("ajuste") = 1
                
                RstDia.Update
                
                '************************************************************************************************
                '--destinos automatico cta debe
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = mIdMes
                RstDia("idlib") = 44
                RstDia("idmov") = mIdMov
                RstDia("numasi") = xNumAsiento
                RstDia("tc") = TipCam
                
                RstDia("impdebsol") = 0
                RstDia("impdebdol") = 0
                RstDia("imphabsol") = 0
                RstDia("imphabdol") = 0
                RstDia("idcue") = IdCtaDestHab
                
                If NulosN(LblIdMon.Caption) = 2 Then
                    RstDia("imphabdol") = Abs(sSaldo) / TipCam
                Else
                    RstDia("imphabsol") = Abs(sSaldo) * TipCam
                End If
                
                RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
                RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 15))
                RstDia("idmon") = 2
                RstDia("ridlib") = 2
                RstDia("iddocpro") = NulosN(Fg1.TextMatrix(mRow, 32))
                RstDia("ridtipper") = 2
                RstDia("ridper") = Fg1.TextMatrix(mRow, 37)
                RstDia("rtipdoc") = NulosN(Fg1.TextMatrix(mRow, 39))
                If IsDate(Fg1.TextMatrix(mRow, 6)) = True Then
                    RstDia("rfchope") = CDate(Fg1.TextMatrix(mRow, 6))
                End If
                RstDia("rnumerodoc") = Fg1.TextMatrix(mRow, 5)
                RstDia("rregistro") = Fg1.TextMatrix(mRow, 1)
                RstDia("rglosa") = ""
                RstDia("ridmon") = 2
                RstDia("rtipdoc1") = 108
                RstDia("rnumerodoc1") = NulosC(Fg1.TextMatrix(mRow, 41))
                RstDia("ridtipper2") = 2
                RstDia("ridper2") = NulosN(Fg1.TextMatrix(mRow, 38))
                RstDia("rtipdoc2") = CDate(Fg1.TextMatrix(mRow, 40))
                RstDia("rfchope2") = Fg1.TextMatrix(mRow, 15)
                RstDia("rnumerodoc2") = Fg1.TextMatrix(mRow, 14)
                
                RstDia("ajuste") = 1

                RstDia.Update
                '************************************************************************************************
                
            End If
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        End If
        
    Next

    xCon.CommitTrans
    
    MsgBox "El proceso termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo

    Set RstDia = Nothing
    
    Grabar = True
    
    fraBarra.Visible = False
    
    Exit Function
    
LaCague:
'    Resume
    xCon.RollbackTrans
    Set RstDia = Nothing
    fraBarra.Visible = False
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" & vbCr & Trim(Err.Description)
End Function
