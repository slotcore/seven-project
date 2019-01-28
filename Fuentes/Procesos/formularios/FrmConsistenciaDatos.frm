VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsistenciaDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herramientas - Consistencia de Datos"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "[ Seleccionar ]"
      Height          =   585
      Left            =   -15
      TabIndex        =   6
      Top             =   390
      Width           =   7995
      Begin VB.CommandButton Command2 
         Caption         =   "Cargar Todos"
         Height          =   285
         Left            =   3570
         TabIndex        =   12
         Top             =   210
         Width           =   1155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Procesar"
         Height          =   285
         Left            =   2475
         TabIndex        =   9
         Top             =   210
         Width           =   1005
      End
      Begin VB.TextBox TxtNum 
         Height          =   285
         Index           =   1
         Left            =   1380
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   210
         Width           =   1020
      End
      Begin VB.TextBox TxtNum 
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   210
         Width           =   1020
      End
      Begin VB.Label lblTotal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   1
         Left            =   6660
         TabIndex        =   11
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Registros: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   5205
         TabIndex        =   10
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   960
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   2
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   75
         Width           =   1035
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   3
         Top             =   75
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8025
      _ExtentX        =   14155
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
               Picture         =   "FrmConsistenciaDatos.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsistenciaDatos.frx":2A98
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   4875
      Left            =   15
      TabIndex        =   13
      Top             =   1005
      Width           =   7965
      _cx             =   14049
      _cy             =   8599
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
      FrontTabForeColor=   -2147483630
      Caption         =   "Condicion de Busqueda|Detalle de la Busqueda"
      Align           =   0
      CurrTab         =   1
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4455
         Left            =   -8520
         TabIndex        =   16
         Top             =   45
         Width           =   7875
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   4380
            Left            =   0
            TabIndex        =   17
            Top             =   45
            Width           =   7725
            _cx             =   13626
            _cy             =   7726
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
            FormatString    =   $"FrmConsistenciaDatos.frx":2E2A
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
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4455
         Left            =   45
         TabIndex        =   14
         Top             =   45
         Width           =   7875
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   4380
            Left            =   30
            TabIndex        =   15
            Top             =   30
            Width           =   7725
            _cx             =   13626
            _cy             =   7726
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
            BackColor       =   14745342
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   128
            ForeColorSel    =   16777215
            BackColorBkg    =   -2147483636
            BackColorAlternate=   14745342
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
            Rows            =   1
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConsistenciaDatos.frx":2E53
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
End
Attribute VB_Name = "FrmConsistenciaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BAND_INTERRUMPIR As Boolean


Private Sub Command1_Click()

    pProcesar 0
End Sub

Private Sub Command2_Click()
    pProcesar 1
End Sub

Private Sub Fg1_EnterCell()
    Fg1.SelectionMode = flexSelectionFree
    Fg1.Editable = flexEDKbdMouse
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

Private Sub pExportarExcel()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Numeración Faltante", "De:" & TxtNum(0).Text & " A: " & TxtNum(1).Text, "Consistencia de Datos"
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
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, "Consistencia de Datos", "Numeración Faltante", "De:" & TxtNum(0).Text & " A: " & TxtNum(1).Text, True, True
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
        .Cols = 6
                 
        '.FrozenCols = Q_POS_MES_INICIO - 1
        .ColWidth(0) = 200
        .FrozenCols = 0
        .Rows = 1
        .FixedRows = 1
        .TextMatrix(0, 1) = "Id":    .ColWidth(1) = 300:   .ColAlignment(1) = flexAlignLeftBottom:        .Row = 0: .Col = 1: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(0, 2) = "Sel":    .ColWidth(2) = 450:   .ColAlignment(2) = flexAlignLeftBottom:        .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(0, 3) = "Nombre":    .ColWidth(3) = 2500:   .ColAlignment(3) = flexAlignLeftBottom:   .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(0, 4) = "Inicio":  .ColWidth(4) = 1200:   .ColAlignment(4) = flexAlignCenterCenter:        .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
        .TextMatrix(0, 5) = "Final":   .ColWidth(5) = 1200:   .ColAlignment(5) = flexAlignCenterCenter:        .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftBottom
        .ColDataType(2) = flexDTBoolean
        
    End With
    
    
    With Fg2
        '-----
        .Cols = 2
                 
        '.FrozenCols = Q_POS_MES_INICIO - 1
        .ColWidth(0) = 200
        .FrozenCols = 0
        .Rows = 1
        .FixedRows = 1
        .TextMatrix(0, 1) = "Número":    .ColWidth(1) = 2500:   .ColAlignment(1) = flexAlignLeftBottom:        .Row = 0: .Col = 1: .CellAlignment = flexAlignLeftBottom
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




Private Sub pProcesar(TIPO As Integer)
    '--tipo =0 procesar
    '--tipo =1 cargar todos
    Dim nSQL As String
    Dim rst As New ADODB.Recordset
    Dim nFormatNumero As String
    Dim mRow As Integer
    
    Fg2.Rows = Fg2.FixedRows
    DoEvents
    
    For mRow = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 2)) = -1 Then
        '--ventas
            nSQL = "SELECT vta_ventas.tipdoc, vta_ventas.numser, vta_ventas.numdoc as numero, vta_ventas.numreg, vta_ventas.fchreg " _
                + vbCr + " From vta_ventas " _
                + vbCr + " WHERE (((vta_ventas.tipdoc)=" & NulosN(Fg1.TextMatrix(mRow, 1)) & ") AND ((vta_ventas.numser)='0001') AND ((vta_ventas.numreg)<>'0001'));"
        '
            nFormatNumero = "0000000000"
        
        '--lgd, lgc
'            nSQL = "SELECT vta_gastodebito.tipdoc, vta_gastodebito.numser, vta_gastodebito.numdoc as numero " _
'                + vbCr + " From vta_gastodebito " _
'                + vbCr + " WHERE vta_gastodebito.tipdoc=" & NulosN(Fg1.TextMatrix(mRow, 1)) & " and  vta_gastodebito.numser = '0001' " _
'                + vbCr + " ORDER BY vta_gastodebito.numdoc asc; "
'
'            nFormatNumero = "0000000"
            
            '--cheques
            Dim nBanco As String
            Dim nNumCta As String
            
        '    nSQL = "SELECT zzz_lista_cheques.obanco, zzz_lista_cheques.onumcta, zzz_lista_cheques.ochequenumero as numero " _
        '        + vbCr + " From zzz_lista_cheques " _
        '        + vbCr + " WHERE (((zzz_lista_cheques.obanco & '-' & zzz_lista_cheques.onumcta)='" & NulosN(Fg1.TextMatrix(mRow, 1)) & "') ) " _
        '        + vbCr + " ORDER BY zzz_lista_cheques.ochequenumero;"
        '     nFormatNumero = "00000000"
            
            
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = Fg1.TextMatrix(mRow, 3)
            Set rst = Nothing
            RST_Busq rst, nSQL, xCon
            
            Dim mNumeroIni&, mNumeroFin&, mNumeroBusca&
            
            mNumeroIni = NulosN(Fg1.TextMatrix(mRow, 4))
            mNumeroFin = NulosN(Fg1.TextMatrix(mRow, 5))
            
            DoEvents
            If rst.RecordCount <> 0 Then
                PgBar.Min = 0
                FraProgreso.Visible = True
                If TIPO = 0 Then
                    PgBar.Max = mNumeroFin - mNumeroIni + 1
                Else
                    PgBar.Max = rst.RecordCount
                End If
                PgBar.Value = 0
                Me.MousePointer = vbHourglass
                rst.MoveFirst
            
                DoEvents
                If TIPO = 0 Then
                
                    For mNumeroBusca = mNumeroIni To mNumeroFin
                        PgBar.Value = PgBar.Value + 1
                        DoEvents
                        rst.Filter = "numero='" & Format(mNumeroBusca, nFormatNumero) & "'"
                        If rst.RecordCount = 0 Then
                            Fg2.Rows = Fg2.Rows + 1
                            Fg2.TextMatrix(Fg2.Rows - 1, 1) = Format(mNumeroBusca, nFormatNumero)
                            lblTotal(1).Caption = Fg2.Rows - 1
                            
                        End If
                    Next
                
                Else
                    Do While Not rst.EOF
                        PgBar.Value = PgBar.Value + 1
                        Fg2.Rows = Fg2.Rows + 1
                        Fg2.TextMatrix(Fg2.Rows - 1, 1) = Format(rst("numero"), nFormatNumero)
                        lblTotal(1).Caption = Fg2.Rows - 1
                        rst.MoveNext
                    Loop
                    
                End If
            End If
        End If
    Next mRow
    
    FraProgreso.Visible = False
    Set rst = Nothing
    
    Me.MousePointer = vbDefault
End Sub


Private Sub pConsultar()
    Dim nSQL As String
    Dim rst As New ADODB.Recordset
    '--ventas
    nSQL = "SELECT vta_ventas.tipdoc AS id, mae_documento.descripcion as nombre, Min(vta_ventas.numdoc) AS nummin, Max(vta_ventas.numdoc) AS nummax " _
    + vbCr + " FROM vta_ventas INNER JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id where vta_ventas.numreg<>'0001'" _
    + vbCr + " GROUP BY vta_ventas.tipdoc, mae_documento.descripcion; "
    
    '--lgd, lgc
'    nSQL = "SELECT vta_gastodebito.tipdoc as id, mae_documento.descripcion AS nombre, Min(vta_gastodebito.numdoc) AS nummin, Max(vta_gastodebito.numdoc) AS nummax " _
'    + vbCr + " FROM vta_gastodebito LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id " _
'    + vbCr + " WHERE (((vta_gastodebito.numser)='0001')) " _
'    + vbCr + " GROUP BY vta_gastodebito.tipdoc, mae_documento.descripcion; "
    
    '--cheques
    Dim nBanco As String
    Dim nNumCta As String

''
'    nSQL = "SELECT zzz_lista_cheques.obanco & '-' & zzz_lista_cheques.onumcta AS id, zzz_lista_cheques.obanco & '   Nº Cta. ' & zzz_lista_cheques.onumcta AS nombre, zzz_lista_cheques.omonedacheque, Min(zzz_lista_cheques.ochequenumero) AS nummin, Max(zzz_lista_cheques.ochequenumero) AS nummax, Count(zzz_lista_cheques.onumcta) AS cant " _
'    + vbCr + " From zzz_lista_cheques " _
'    + vbCr + " GROUP BY zzz_lista_cheques.omonedacheque, zzz_lista_cheques.onumcta, zzz_lista_cheques.obanco, zzz_lista_cheques.onumcta; "

    
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    RST_Busq rst, nSQL, xCon
    LimpiaText TxtNum
    If rst.RecordCount <> 0 Then
        TxtNum(0).Text = NulosC(rst("nummin"))
        TxtNum(1).Text = NulosC(rst("nummax"))
    End If
    
    Do While Not rst.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = rst("id")
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = -1
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = rst("nombre")
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = rst("nummin")
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = rst("nummax")
        rst.MoveNext
    Loop

    Set rst = Nothing
    
End Sub



