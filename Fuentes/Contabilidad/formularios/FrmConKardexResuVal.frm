VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConKardexResuVal 
   Caption         =   "Contabilidad - Consulta de Kardex Resumen Valorizado"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "[ Filtro de Almacen ]"
      Height          =   1335
      Left            =   3960
      TabIndex        =   8
      Top             =   360
      Width           =   6885
      Begin VB.Frame Frame4 
         Height          =   1085
         Left            =   5280
         TabIndex        =   12
         Top             =   150
         Width           =   1450
         Begin VB.CommandButton cmd 
            Caption         =   "&Agregar"
            Height          =   330
            Index           =   0
            Left            =   50
            TabIndex        =   14
            Top             =   180
            Width           =   1305
         End
         Begin VB.CommandButton cmd 
            Caption         =   "&Eliminar"
            Height          =   330
            Index           =   1
            Left            =   50
            TabIndex        =   13
            Top             =   600
            Width           =   1305
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg 
         Height          =   945
         Index           =   1
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Buscar Linea"
         Top             =   270
         Width           =   5130
         _cx             =   9049
         _cy             =   1667
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
         ForeColorSel    =   -2147483634
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConKardexResuVal.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   -1  'True
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
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   3390
      TabIndex        =   3
      Top             =   4410
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   4
         Top             =   420
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Shape Shape1 
         Height          =   885
         Index           =   0
         Left            =   60
         Top             =   90
         Width           =   5805
      End
      Begin VB.Label LblProg 
         AutoSize        =   -1  'True
         Caption         =   "LblProg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1350
         TabIndex        =   7
         Top             =   180
         Width           =   525
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Left            =   150
         TabIndex        =   6
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cancelar = ESC"
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
         Left            =   4470
         TabIndex        =   5
         Top             =   720
         Width           =   1260
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Filtro General ]"
      Height          =   1335
      Left            =   30
      TabIndex        =   2
      Top             =   360
      Width           =   3900
      Begin VB.ComboBox cbMes 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   2865
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   11
         Top             =   630
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7365
      Top             =   45
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
            Picture         =   "FrmConKardexResuVal.frx":0092
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":05D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":0968
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":0AC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":0E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":0FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":142C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":1544
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":1A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":1FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":20E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":21F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":2648
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConKardexResuVal.frx":27B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg 
      Height          =   5775
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   1800
      Width           =   12330
      _cx             =   21749
      _cy             =   10186
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
      ForeColorSel    =   -2147483634
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConKardexResuVal.frx":2CFC
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
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu00 
         Caption         =   "Insertar Ítem"
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu menu01 
         Caption         =   "Eliminar Ítem"
      End
   End
End
Attribute VB_Name = "FrmConKardexResuVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FrmVerKardex.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : MUESTRA EL VINCAR DEL ITEM SELECCIONADO, ADEMAS PERMITE COSTEAS LAS SALIDAS
'*                    MEDIANTE EL METODO PROMEDIO PONDERADO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 23/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim SeEjecuto As Boolean                  ' VARIABLE QUE CONTROLARA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim StockIni As Double                    ' ALMACENA EL STOCK INICIAL DEL ITEM
Dim xPrecioIni As Double                  ' ALMACENA EL PRECIO INICIAL DEL ITEM
Dim MuestraRpt As Integer
Dim cSQL As String
Dim INDICE_ As Integer
Dim F As New SistemaLogica.Funciones

Private Sub pIniciarCampos()
    Blanquea
    Llenar_Mes cbMes
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TextBox PARA EL INGRESO DE DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    fg(0).Rows = fg(0).FixedRows
    fg(1).Rows = fg(1).FixedRows
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    Dim nTitulo As String
    
    Select Case Index
        
        Case 0 ' Agregar Almacen
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
            
            nTitulo = "Buscando Almacenes"
            
            nSQLId = nSQLId & GENERAR_SQL_ID(fg(1), fg(1).ColIndex("_idalm"), " WHERE alm_almacenes.id", "NOT IN", True)
            cSQL = "SELECT alm_almacenes.id, alm_almacenes.codigo, alm_almacenes.descripcion FROM alm_almacenes " & nSQLId
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            fg(1).Rows = fg(1).Rows + 1
            fg(1).TextMatrix(fg(1).Rows - 1, fg(1).ColIndex("_idalm")) = F.NuloNumeric(xRs("id"))
            fg(1).TextMatrix(fg(1).Rows - 1, fg(1).ColIndex("_codalm")) = F.NuloString(xRs("codigo"))
            fg(1).TextMatrix(fg(1).Rows - 1, fg(1).ColIndex("_desalm")) = F.NuloString(xRs("descripcion"))
            fg(1).TopRow = fg(1).Rows - 1
            fg(1).Row = fg(1).Rows - 1
            Set xRs = Nothing
            
        Case 1 ' Eliminar
            If fg(1).Row < 1 Then
                MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            If fg(1).Rows = fg(1).FixedRows Then
                MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                fg(1).SetFocus
                Exit Sub
            End If
            fg(1).RemoveItem fg(1).Row
            
    End Select
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    pIniciarCampos
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    fg(0).Top = 1800
    fg(0).Width = Me.Width - 315
    fg(0).Height = Me.Height - 2415
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Buscar
    
    If Button.Index = 2 Then ExportarExcel fg(0)
        
    If Button.Index = 5 Then
        Unload Me
    End If
End Sub

Private Sub Buscar()
    If fValidarDatos() = False Then Exit Sub
    pCargarResumido
End Sub

Private Sub pCargarResumido()
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim FchInicio As Date
    Dim FchFinal As Date
    Dim mMesActual As Integer
    Dim mTotal As Double
        
    Me.MousePointer = vbHourglass
    mMesActual = cbMes.ListIndex + 1
    FchInicio = F.RetornarPrimerDiaMes(CDate("01/" & mMesActual & "/" & AnoTra & ""))
    FchFinal = F.RetornarUltimoDiaMes(FchInicio)
    fg(0).Rows = fg(0).FixedRows
    CentrarFrm FraProgreso
    FraProgreso.Visible = True
    PgBar.Min = 0
    PgBar.Max = fg(1).Rows - 1
    PgBar.Value = 0
    With fg(1)
        For A = .FixedRows To .Rows - 1
            PgBar.Value = PgBar.Value + 1
            LblProg.Caption = "Almacen: " & NulosC(fg(1).TextMatrix(A, fg(1).ColIndex("_desalm")))
            CargarValorizadoAlmacen "ALMACEN " & NulosC(fg(1).TextMatrix(A, fg(1).ColIndex("_desalm"))), NulosN(fg(1).TextMatrix(A, fg(1).ColIndex("_idalm"))), FchInicio, FchFinal
            mTotal = mTotal + fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("_costo"))
        Next
    End With
    ' SE AGREGAN LOS TOTALES
    fg(0).Rows = fg(0).Rows + 1
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("_tipdoc")) = "TOTAL STOCK " & UCase(cbMes.Text)
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("_costo")) = Format(mTotal, FORMAT_MONTO)
    fg(0).Select fg(0).Rows - 1, fg(0).ColIndex("_tipdoc"), fg(0).Rows - 1, fg(0).ColIndex("_costo")
    fg(0).FillStyle = flexFillRepeat
    fg(0).CellFontBold = True
    fg(0).Select 1, 1
    fg(0).TopRow = fg(0).Rows - 1
    
    Set xRs = Nothing
    FraProgreso.Visible = False
    Me.MousePointer = vbDefault
End Sub

Sub CargarValorizadoAlmacen(CadenaDescripcion As String, IdAlmacen As Long, FchInicio As Date, FchFinal As Date)
    Dim cSQL As String
    Dim mTotal As Double
    Dim xRs As New ADODB.Recordset
    
    With fg(0)
        '********************
        ' Saldo Inicial
        '********************
        cSQL = "SELECT SUM(CONKARDEXTOTINI.canent) AS canent, SUM(CONKARDEXTOTINI.cansal) As cansal, SUM(CONKARDEXTOTINI.canini) As canini, SUM(CONKARDEXTOTINI.costoent) As costoent, SUM(CONKARDEXTOTINI.costosal) As costosal, SUM(CONKARDEXTOTINI.costoini) As costoini " _
        + vbCr + "FROM ( " _
        + vbCr + F.SQL_MovHistoricoTotalizado(IdAlmacen, FchInicio - 1, , xCon, True) _
        + vbCr + ") As CONKARDEXTOTINI "
        Set xRs = Nothing
        Set xRs = F.GeneraRstSQL(cSQL, xCon)
        If xRs.RecordCount > 0 Then
            xRs.MoveFirst
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("_tipmov")) = "I"
            .TextMatrix(.Rows - 1, .ColIndex("_tipdoc")) = "Inventario Inicial"
            .TextMatrix(.Rows - 1, .ColIndex("_docabrev")) = "II"
            .TextMatrix(.Rows - 1, .ColIndex("_costo")) = Format(NulosN(xRs("costoini")) + NulosN(xRs("costoent")) - NulosN(xRs("costosal")), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, .ColIndex("_porc")) = Format("", FORMAT_CANTIDAD)
            mTotal = mTotal + NulosN(xRs("costoini")) + NulosN(xRs("costoent")) - NulosN(xRs("costosal"))
        End If
        
        '**********************
        ' Movimientos Actuales
        '**********************
        cSQL = "SELECT CONKARDEXTOT.tipmovcad As tipmov, CONKARDEXTOT.numser, CONKARDEXTOT.tipdocref As tipdoc, CONKARDEXTOT.doc As docabrev, CONKARDEXTOT.numserref, '' As porc, SUM(CONKARDEXTOT.costo) As costo " _
            + vbCr + "FROM ( " _
            + vbCr + F.SQL_MovDetallado(, IdAlmacen, FchInicio, FchFinal, xCon, , False) _
            + vbCr + ") AS CONKARDEXTOT " _
            + vbCr + "GROUP BY CONKARDEXTOT.tipmovcad, CONKARDEXTOT.numser, CONKARDEXTOT.tipdocref, CONKARDEXTOT.doc, CONKARDEXTOT.numserref, '' "
            
        Set xRs = Nothing
        Set xRs = F.GeneraRstSQL(cSQL, xCon)
        If xRs.RecordCount = 0 Then Exit Sub
        xRs.MoveFirst
    
        .MergeCells = flexMergeFixedOnly
        While Not xRs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("_tipmov")) = NulosC(xRs("tipmov"))
            .TextMatrix(.Rows - 1, .ColIndex("_numser")) = NulosC(xRs("numser"))
            .TextMatrix(.Rows - 1, .ColIndex("_tipdoc")) = NulosC(xRs("tipdoc"))
            .TextMatrix(.Rows - 1, .ColIndex("_docabrev")) = NulosC(xRs("docabrev"))
            .TextMatrix(.Rows - 1, .ColIndex("_numserref")) = NulosC(xRs("numserref"))
            .TextMatrix(.Rows - 1, .ColIndex("_costo")) = Format(F.NuloNumeric(xRs("costo")), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, .ColIndex("_porc")) = Format(F.NuloNumeric(xRs("porc")), FORMAT_CANTIDAD)
            
            ' Totales
            If NulosC(xRs("tipmov")) = "I" Then
                mTotal = mTotal + F.NuloNumeric(xRs("costo"))
            Else
                mTotal = mTotal - F.NuloNumeric(xRs("costo"))
            End If
            xRs.MoveNext
        Wend
        
        ' SE AGREGAN LOS TOTALES
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("_tipdoc")) = "TOTAL " & CadenaDescripcion
        .TextMatrix(.Rows - 1, .ColIndex("_costo")) = Format(mTotal, FORMAT_MONTO)
        .Select .Rows - 1, .ColIndex("_tipdoc"), .Rows - 1, .ColIndex("_costo")
        .FillStyle = flexFillRepeat
        .CellFontBold = True
        .Select 1, 1
        .TopRow = .Rows - 1
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : ExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ExportarExcel(ByRef GRID_ As VSFlexGrid)
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
    
    TITULO_ = "REPORTE RESUMEN DE KARDEX VALORIZADO"

    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, GRID_, TITULO_, cbMes.Text, ""
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
End Sub

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VERIFICA QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    ' si esta la fecha correcta
    If fg(1).Rows = fg(1).FixedRows Then
        MsgBox "Debe de escoger al menos un almacen de proceso", vbExclamation, xTitulo
        fg(1).SetFocus
        fValidarDatos = False
        Exit Function
    End If
    
    fValidarDatos = True
End Function
