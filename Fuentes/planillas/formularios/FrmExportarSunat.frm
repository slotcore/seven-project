VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmExportarSunat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Exportar a Sunat"
   ClientHeight    =   7560
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11775
   Begin VB.Frame FraPath 
      BorderStyle     =   0  'None
      Height          =   3450
      Left            =   4005
      TabIndex        =   10
      Top             =   2220
      Visible         =   0   'False
      Width           =   3390
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3120
         Picture         =   "FrmExportarSunat.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   18
         ToolTipText     =   "Cerrar"
         Top             =   75
         Width           =   195
      End
      Begin VB.TextBox txtpath 
         BackColor       =   &H8000000F&
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
         Index           =   1
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "txtpath(1)"
         Top             =   2370
         Width           =   3240
      End
      Begin VB.TextBox txtpath 
         BackColor       =   &H8000000F&
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
         Index           =   0
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "txtpath(0)"
         Top             =   2025
         Width           =   3240
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   15
         Top             =   390
         Width           =   3240
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   75
         TabIndex        =   14
         Top             =   735
         Width           =   3240
      End
      Begin VB.CommandButton CmdPath 
         Caption         =   "&Aceptar"
         Height          =   420
         Index           =   0
         Left            =   765
         TabIndex        =   12
         Top             =   2880
         Width           =   1020
      End
      Begin VB.CommandButton CmdPath 
         Caption         =   "&Cancelar"
         Height          =   420
         Index           =   1
         Left            =   1860
         TabIndex        =   11
         Top             =   2880
         Width           =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         Index           =   2
         X1              =   0
         X2              =   3500
         Y1              =   2745
         Y2              =   2745
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   0
         Y1              =   15
         Y2              =   3450
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   -30
         X2              =   4015
         Y1              =   3435
         Y2              =   3435
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -330
         X2              =   3715
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   3375
         X2              =   3375
         Y1              =   15
         Y2              =   3465
      End
      Begin VB.Label LblTituloFrame 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Guardar en..."
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
         Left            =   90
         TabIndex        =   13
         Top             =   90
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   300
         Index           =   1
         Left            =   30
         Top             =   45
         Width           =   3315
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   30
      TabIndex        =   5
      Top             =   390
      Width           =   11745
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   0
         Left            =   1665
         Picture         =   "FrmExportarSunat.frx":02EC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Seleccione el Personal"
         Top             =   600
         Width           =   210
      End
      Begin VB.ComboBox cb_sel 
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   6330
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   1155
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "txt_cb(0)"
         Top             =   570
         Width           =   750
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   675
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
         Height          =   285
         Index           =   0
         Left            =   6450
         TabIndex        =   8
         Top             =   570
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Seleccionar"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   840
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
         Left            =   1905
         TabIndex        =   7
         Top             =   570
         Width           =   4590
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6045
      Left            =   30
      TabIndex        =   4
      Top             =   1470
      Width           =   11745
      _cx             =   20717
      _cy             =   10663
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
      BackColorSel    =   -2147483635
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmExportarSunat.frx":041E
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
      DataMode        =   1
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar Bloc de Notas"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cerrar"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6585
         Top             =   15
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExportarSunat.frx":045A
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExportarSunat.frx":08AE
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExportarSunat.frx":0A1A
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExportarSunat.frx":0F62
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExportarSunat.frx":12FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExportarSunat.frx":140C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExportarSunat.frx":151E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExportarSunat.frx":16B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmExportarSunat.frx":1B0A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmExportarSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

'--DE LA IMPRESION
Dim T_RPT_PERIODO As String '--PERIODO DEL REPORTE
Dim T_RPT_TITULO As String  '--TITULO DE REPORTE
'------------

Dim Q_COL_FILA As Integer   '--INDICA LA CANTIDAD DE COLUMNAS ANTES DE LOS MESES
                            '--EJ. IDCLI,CLIENTE => Q_COL_FILA=2
                            '--    IDCLI,ID_PTO_VTA,CLIENTE,PTO_VENTA => Q_COL_FILA=4
                            
Dim Q_COL_FILA_OCULTA As Integer '--INDICA LAS COLUMNAS QUE CONTENDRAN LOS ID'S, ESTOS SE OCULTARAN
                                '-- -1 NO SE OCULTA, <> -1 SE PROCEDE A ACULTAR

Dim Q_POSICION_TOTAL  As Integer '--INDICA LA POCISION DE LA COLUMNA DONDE SE COLOCARA EL NOMBRE DEL TOTAL Y TOTAL_GRL
                                 '--OBTENDRA VALOR EN fGenerarConsulta()

'------------
'-------
'------------
Dim ARR_FORMATOS() As String  '--INDICA LA LISTA DE FORMATOS
Dim Quehace As Integer

Private Sub cb_sel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_cb(0).SetFocus
    End If
End Sub

Private Sub Fg1_DblClick()
If Fg1.Row <= 1 Then Exit Sub
Fg1.TextMatrix(Fg1.Row, 1) = Not Fg1.TextMatrix(Fg1.Row, 1)
End Sub

Private Sub Fg1_EnterCell()
    If Quehace = 3 Then Exit Sub
    If Fg1.Col = 1 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pHabilitarBotonPath False '--ocultar guardar en...
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo error
    CentrarFrm Me
    Quehace = 3
    ReDim ARR_FORMATOS(8, 2)
    'ARR_FORMATOS(?,0) extension del archivo
    'ARR_FORMATOS(?,1) indica el tipo de dato del identificador de registro N::Numerico  c::Caracter
    'ARR_FORMATOS(?,2) Nombre del formato a exportar
    
    ARR_FORMATOS(0, 0) = "t00":  ARR_FORMATOS(0, 1) = "N":    ARR_FORMATOS(0, 2) = "Datos Principales"
    ARR_FORMATOS(1, 0) = "t01":  ARR_FORMATOS(1, 1) = "N":    ARR_FORMATOS(1, 2) = "Datos del Trabajador"
    ARR_FORMATOS(2, 0) = "t02":  ARR_FORMATOS(2, 1) = "N":    ARR_FORMATOS(2, 2) = "Datos del Pensionista"
    ARR_FORMATOS(3, 0) = "t03":  ARR_FORMATOS(3, 1) = "N":    ARR_FORMATOS(3, 2) = "Datos del Prestador de Servicio - Cuarta Categoría"
    ARR_FORMATOS(4, 0) = "s00":  ARR_FORMATOS(4, 1) = "N":    ARR_FORMATOS(4, 2) = "Datos de Suspención de Cuarta Categoría"
    ARR_FORMATOS(5, 0) = "t04":  ARR_FORMATOS(5, 1) = "N":    ARR_FORMATOS(5, 2) = "Datos del Prestador de Servicios - Modalidad Formativa"
    ARR_FORMATOS(6, 0) = "t05":  ARR_FORMATOS(6, 1) = "N":    ARR_FORMATOS(6, 2) = "Datos del Personal de Terceros"
    ARR_FORMATOS(7, 0) = "p00":  ARR_FORMATOS(7, 1) = "C":    ARR_FORMATOS(7, 2) = "Datos de Periodos Laborales"
    ARR_FORMATOS(8, 0) = "der":  ARR_FORMATOS(8, 1) = "C":    ARR_FORMATOS(8, 2) = "Datos de derechohabientes"
    
    Dim K&
    cb_sel.Clear
    LimpiaText txt_cb, True
    For K = 0 To UBound(ARR_FORMATOS())
        cb_sel.AddItem ARR_FORMATOS(K, 2)
    Next K
    cb_sel.ListIndex = 0
    '----------------
    
    fGenerarConsulta
    pConfigurarGrilla
    
    Exit Sub
error:
    SHOW_ERROR
End Sub

Private Sub pConsultar()
    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    '--
    Dim nSQL As String '--RECIBIR LA CONSULTA
    Quehace = 3
    If fValidarConsulta() = False Then Exit Sub
    '--configurar la presentacion de la grilla
    LimpiarGrid Me.Fg1, False, 1
    '--entrar solo una vez
    nSQL = fGenerarConsulta()
    pConfigurarGrilla
    '----
    Me.MousePointer = vbHourglass
    DoEvents
    '------------------------------------------------
    If nSQL = "" Then GoTo Salir
    DoEvents
    '--cargar el rst
    RST_Busq rst_select, nSQL, xCon
   '--------------------------------------
    pCargarDatosGrilla rst_select
    
   '--------------------------------------
   If rst_select.RecordCount <> 0 Then pColorCamposSunat rst_select
   '--------------------------------------
   Quehace = 2 '--modificar
Salir:
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Quehace = 3
    Me.MousePointer = vbDefault
    Set rst_select = Nothing
    SHOW_ERROR Me.Name, "pConsultar"
    
End Sub

Private Function pCargarDatosGrilla(RST_ORIGEN As ADODB.Recordset)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim BAND_ADD_REG As Boolean
    
    
    BAND_ADD_REG = True
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    
    While Not RST_ORIGEN.EOF
        DoEvents
        '---------------------------------------------------------
        ADD_REG Fg1
        '--CARGAR A LA GRILLA
        pCargarDatosGrillaArrayTmp RST_ORIGEN, Fg1.Rows - 1
        '---------------------------------------------------------
        RST_ORIGEN.MoveNext
    Wend

End Function

Private Sub pColorCamposSunat(RST_ORIGEN As ADODB.Recordset)
Dim mCampo&
For mCampo = 0 To RST_ORIGEN.Fields.Count - 1
    If InStr(UCase(RST_ORIGEN.Fields(mCampo).Name), "E_") <> 0 Then
        GRID_COLOR_FONDO Fg1, Fg1.FixedRows, mCampo + 1, Fg1.Rows - 1, mCampo + 1, &HE0FEFE
    End If
Next mCampo
End Sub

Private Function pCargarDatosGrillaArrayTmp(RST_ORIGEN As ADODB.Recordset, _
                                         mRow As Integer)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim mCampo As Integer
    Dim nCampo As String

    '-----------
    DoEvents
    
    For mCampo = 0 To RST_ORIGEN.Fields.Count - 1
        nCampo = RST_ORIGEN.Fields(mCampo).Name
        Fg1.TextMatrix(mRow, mCampo + 1) = RST_ORIGEN.Fields(nCampo) & ""
    Next
End Function


Private Sub Imprimir()

    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, cb_sel.Text, "", "", False, True
    
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Imprimir"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase ARR_FORMATOS()
End Sub



'------
Private Function fGenerarConsulta(Optional nLista As String = "") As String
    '--FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    '--
    '--fExportar  = true cuando se decide exportar al block de notas
    '--fExportar =false cuando se visualiza en la pantalla
    Dim nSQL As String
    
    Dim vStrFiltro As String
    
    Dim RstTmp As New ADODB.Recordset
   

    '--generar la consulta segun tipo de formato
    Dim N_VALOR As String
    Dim nCampos As String
    Dim nWhere As String
    Dim nFrom As String
    Dim nGroupBy As String
    Dim nOrderBy As String
    
    '----------------------------------
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        nWhere = " pla_empleados.id = " & NulosN(lbl_cod(0).Caption) & " "
    End If
    
    
    '----------------------------------
   
    Select Case cb_sel.ListIndex
        Case 0 '-- Datos Principales
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 29:        Q_POSICION_TOTAL = -1:
            
            nCampos = " pla_empleados.id, mae_dociden.codsun AS e_tipdoc, mae_dociden.descripcion AS tipodoc, pla_empleados.numdoc AS e_numdoc, pla_empleados.apepat AS e_apepat, pla_empleados.apemat AS e_apemat, pla_empleados.nom AS e_nom, Format([fchnac],'dd/mm/yyyy') AS e_nacimiento, mae_sexo.id AS e_sexo, mae_sexo.descripcion AS sexo, mae_nacionalidad.codsun AS e_nacionalidad, mae_nacionalidad.descripcion AS nacionalidad, pla_empleados.numtel AS e_numtel, pla_empleados.email AS e_email, mae_indicaesalud.codsun AS e_essalud, mae_indicaesalud.descripcion AS essalud,pla_empleados.numessalud, mae_indicadom.codsun AS e_domiciliado, mae_indicadom.descripcion AS domiciliado, " _
                    + vbCr + " mae_tipovia.codsun AS e_tipovia, mae_tipovia.descripcion AS tipovia, pla_empleados.nomvia AS e_nomvia, pla_empleados.numvia AS e_numvia, pla_empleados.intvia AS e_intvia, mae_tipozona.codsun AS e_tipozona, mae_tipozona.descripcion AS tipozona, pla_empleados.nomzon AS e_nomzon, pla_empleados.refdom AS e_refdom, mae_distrito.codsun AS e_ubigeo, [mae_departamento].[descripcion] & ' - ' & [mae_provincia].[descripcion] & ' - ' & [mae_distrito].[descripcion] AS ubigeo "
            nFrom = " mae_tipozona RIGHT JOIN (mae_tipovia RIGHT JOIN (mae_sexo RIGHT JOIN ((mae_departamento RIGHT JOIN mae_provincia ON mae_departamento.id = mae_provincia.iddepa) RIGHT JOIN (mae_nacionalidad RIGHT JOIN (mae_indicaesalud RIGHT JOIN (mae_indicadom RIGHT JOIN (mae_dociden RIGHT JOIN (mae_distrito RIGHT JOIN pla_empleados ON mae_distrito.id = pla_empleados.iddis) ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_indicadom.id = pla_empleados.inddomi) ON mae_indicaesalud.id = pla_empleados.indessalud) ON mae_nacionalidad.id = pla_empleados.idnac) ON mae_provincia.id = mae_distrito.idprov) ON mae_sexo.id = pla_empleados.idsex) ON mae_tipovia.id = pla_empleados.idtipvia) ON mae_tipozona.id = pla_empleados.idtipzon "
            nGroupBy = ""
            nOrderBy = " pla_empleados.apepat;"
            
            If nLista <> "" Then
                nWhere = nWhere + IIf(nWhere = "", "", " AND ") & Replace(nLista, "CAMPO_REEMPLAZA", "pla_empleados.id")
            Else
                nWhere = IIf(nWhere = "", "", nWhere & " AND ") & " pla_empleados.idcat <> 6 "
            End If
            
        Case 1 '--Datos del Trabajador
            '-----------------------------------
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 48:        Q_POSICION_TOTAL = -1:
    
            nCampos = "  pla_empleados.id, mae_dociden.codsun AS e_tipodoc, mae_dociden.descripcion AS tipodoc, pla_empleados.numdoc AS e_numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_tipotrabajador.codsun AS e_tipotrabajador, mae_tipotrabajador.descripcion AS tipotrabajador, mae_regimenlab.codsun AS e_regimenlaboral, mae_regimenlab.descripcion AS regimenlaboral, mae_niveleducativo.codsun AS e_niveleducativo, mae_niveleducativo.descripcion AS niveleducativo, mae_ocupacion.codsun AS e_ocupacion, " _
                    + vbCr + " mae_ocupacion.descripcion AS ocupacion,  Abs([pla_categoria1].[discapacidad]) AS e_discapacidad, IIf([pla_categoria1].[discapacidad]=0,'No','Si') AS discapacidad, mae_regimenpen.codsun AS e_regimenpension, mae_regimenpen.descripcion AS regimenpension, Format([fchinsregpen],'dd/mm/yyyy') AS e_fchinsregpen, pla_categoria1.cuspp AS e_cuspp, mae_sctrsalud.codsun AS e_sctrsalud, mae_sctrsalud.descripcion AS sctrsalud, mae_sctrpension.codsun AS e_sctrpension, mae_sctrpension.descripcion AS sctrpension, mae_tipocontrato.codsun AS e_tipocontratro, mae_tipocontrato.descripcion AS tipocontratro, Abs([pla_categoria1.opc1]) AS e_opc1, IIf([pla_categoria1].[opc1]=0,'No','Si') AS opc1, Abs([pla_categoria1].[opc2]) AS e_opc2, IIf([pla_categoria1].[opc2]=0,'No','Si') AS opc2, Abs([pla_categoria1].[opc3]) AS e_opc3, " _
                    + vbCr + " IIf([pla_categoria1].[opc3]=0,'No','Si') AS opc3, Abs([pla_categoria1].[opc4]) AS e_opc4, IIf([pla_categoria1].[opc4]=0,'No','Si') AS opc4, Abs([pla_categoria1].[opc5]) AS e_opc5, IIf([pla_categoria1].[opc5]=0,'No','Si') AS opc5, mae_periocidad.codsun AS e_periodopagp, mae_periocidad.descripcion AS periodopagp, Abs([pla_categoria1].[opc7]) AS e_opc7, IIf([pla_categoria1].[opc7]=0,'No','Si') AS opc7, mae_eps.codsun AS e_eps, mae_eps.descripcion AS eps, mae_situacion.codsun AS e_situacion, mae_situacion.descripcion AS situacion, Abs([pla_categoria1].[opc6]) AS e_opc6, " _
                    + vbCr + " IIf([pla_categoria1].[opc6]=0,'No','Si') AS opc6, mae_situatraba.codsun AS e_sitaciontrabajador, mae_situatraba.descripcion AS sitaciontrabajador, mae_tipopago.codsun AS e_tipopago, mae_tipopago.descripcion AS tipopago"
            nFrom = "   (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) INNER JOIN (mae_tipotrabajador RIGHT JOIN (mae_tipopago RIGHT JOIN (mae_tipocontrato RIGHT JOIN (mae_situatraba RIGHT JOIN (mae_situacion RIGHT JOIN (mae_sctrsalud RIGHT JOIN (mae_sctrpension RIGHT JOIN (mae_regimenpen RIGHT JOIN (mae_regimenlab RIGHT JOIN (mae_periocidad RIGHT JOIN (mae_ocupacion RIGHT JOIN (mae_niveleducativo RIGHT JOIN (mae_eps RIGHT JOIN pla_categoria1 ON mae_eps.id = pla_categoria1.ideps) ON mae_niveleducativo.id = pla_categoria1.idnivedu) " _
                    + vbCr + " ON mae_ocupacion.id = pla_categoria1.idocu) ON mae_periocidad.id = pla_categoria1.idperiocidad) ON mae_regimenlab.id = pla_categoria1.idreglab) ON mae_regimenpen.id = pla_categoria1.idregpen) ON mae_sctrpension.id = pla_categoria1.sctrpension) ON mae_sctrsalud.id = pla_categoria1.sctrsalud) ON mae_situacion.id = pla_categoria1.idsituacion) ON mae_situatraba.id = pla_categoria1.idsittra) ON mae_tipocontrato.id = pla_categoria1.idtipcon) ON mae_tipopago.id = pla_categoria1.idtippag) ON mae_tipotrabajador.id = pla_categoria1.idtiptra) ON pla_empleados.id = pla_categoria1.idemp "

            If nLista <> "" Then
                nWhere = nWhere + IIf(nWhere = "", "", " AND ") & Replace(nLista, "CAMPO_REEMPLAZA", "pla_empleados.id")
            Else
                nWhere = IIf(nWhere = "", "", nWhere & " AND ") & " pla_empleados.idcat <> 6 "
            End If
            nGroupBy = ""
            nOrderBy = " pla_empleados.apepat;"
            
        Case 2 '--Datos del Pensionista
            '-----------------------------------
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 14:        Q_POSICION_TOTAL = -1:
    
            nCampos = "  pla_empleados.id, mae_dociden.codsun AS e_tipodoc, mae_dociden.descripcion AS tipodoc, pla_empleados.numdoc AS e_numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_tipotrabajador.codsun AS e_tipotrabajador, mae_tipotrabajador.descripcion AS tipotrabajador, mae_regimenpen.codsun AS e_regimenpension, mae_regimenpen.descripcion AS regimenpension, Format([pla_categoria2].fchins,'dd/mm/yyyy') AS e_idregpen, pla_categoria2.cuspp AS e_cuspp, mae_situacion.codsun AS e_situacion, mae_situacion.descripcion AS situacion, mae_tipopago.codsun AS e_tipopago, mae_tipopago.descripcion AS tipopago "
            nFrom = "  (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) RIGHT JOIN (mae_tipotrabajador RIGHT JOIN (mae_tipopago RIGHT JOIN (mae_situacion RIGHT JOIN (mae_regimenpen RIGHT JOIN pla_categoria2 ON mae_regimenpen.id = pla_categoria2.idregpen) ON mae_situacion.id = pla_categoria2.idsituacion) ON mae_tipopago.id = pla_categoria2.idtippag) ON mae_tipotrabajador.id = pla_categoria2.idtippen) ON pla_empleados.id = pla_categoria2.idemp "
            If nLista <> "" Then
                nWhere = nWhere + IIf(nWhere = "", "", " AND ") & Replace(nLista, "CAMPO_REEMPLAZA", "pla_empleados.id")
            End If
            nGroupBy = ""
            nOrderBy = " pla_empleados.apepat;"
            
        Case 3 '--prestador de servicio 4ta categoria
            '-----------------------------------
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 5:        Q_POSICION_TOTAL = -1:
    
            nCampos = "  pla_empleados.id, mae_dociden.codsun AS e_tipodoc, mae_dociden.descripcion AS tipodoc, pla_empleados.numdoc AS e_numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pla_categoria3.numruc AS e_numruc "
            nFrom = " (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) INNER JOIN pla_categoria3 ON pla_empleados.id = pla_categoria3.idemp "
            If nLista <> "" Then
                nWhere = nWhere + IIf(nWhere = "", "", " AND ") & Replace(nLista, "CAMPO_REEMPLAZA", "pla_empleados.id")
            Else
                nWhere = IIf(nWhere = "", "", nWhere & " AND ") & " pla_empleados.idcat <> 6 "
            End If
            nGroupBy = ""
            nOrderBy = " pla_empleados.apepat;"
        
        Case 4 '--Datos de Suspención de Cuarta Categoría
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 10:        Q_POSICION_TOTAL = -1:
    
            nCampos = "  pla_empleados.id, mae_dociden.codsun AS e_tipodoc, mae_dociden.descripcion AS tipodoc, pla_empleados.numdoc AS e_numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pla_categoria3.[numruc] AS e_numruc, pla_categoria3susp.numope AS e_numope, Format([fchpre],'dd/mm/yyyy') AS e_fchpre, pla_categoria3susp.ejercicio AS e_ejercicio, pla_categoria3susp.medio AS e_medio, IIF([pla_categoria3susp].[medio] IS NULL OR [pla_categoria3susp].[medio]= 0 ,'', IIf([pla_categoria3susp].[medio]=1,'Internet','Dependencia SUNAT')) AS medio "
            nFrom = " (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) INNER JOIN (pla_categoria3 LEFT JOIN pla_categoria3susp ON pla_categoria3.idemp = pla_categoria3susp.idemp) ON pla_empleados.id = pla_categoria3.idemp "
            If nLista <> "" Then
                nWhere = nWhere + IIf(nWhere = "", "", " AND ") & Replace(nLista, "CAMPO_REEMPLAZA", "pla_empleados.id")
            Else
                nWhere = IIf(nWhere = "", "", nWhere & " AND ") & " pla_empleados.idcat <> 6 "
            End If
            nGroupBy = ""
            nOrderBy = " pla_empleados.apepat;"
            
        Case 5 '--Datos del Prestador de Servicios - Modalidad Formativa
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 18:        Q_POSICION_TOTAL = -1:
    
            nCampos = "  pla_empleados.id, mae_dociden.codsun AS e_tipodoc, mae_dociden.descripcion AS tipodoc, pla_empleados.numdoc AS e_numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_seguromedico.codsun AS e_seguromedico, mae_seguromedico.descripcion AS seguromedico, mae_niveleducativo.codsun AS e_niveleducativo, mae_niveleducativo.descripcion AS niveleducativo, " _
            + vbCr + "  mae_ocupacion.codsun AS e_ocupacion, mae_ocupacion.descripcion AS ocupacion, Abs([pla_categoria4.indica1]) AS e_indica1, IIf([pla_categoria4].[indica1]=0,'No','Si') AS indica1, Abs([pla_categoria4.indica2]) AS e_indica2, IIf([pla_categoria4].[indica2]=0,'No','Si') AS indica2, mae_centroformacion.codsun AS e_centroformacion, mae_centroformacion.descripcion AS centroformacion, Abs([pla_categoria4.indica3]) AS e_indica3, IIf([pla_categoria4].[indica3]=0,'No','Si') AS indica3 "
            nFrom = " (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) INNER JOIN (mae_seguromedico RIGHT JOIN (mae_ocupacion RIGHT JOIN (mae_niveleducativo RIGHT JOIN (mae_centroformacion RIGHT JOIN pla_categoria4 ON mae_centroformacion.id = pla_categoria4.idcenfor) ON mae_niveleducativo.id = pla_categoria4.idnivedu) ON mae_ocupacion.id = pla_categoria4.idocu) ON mae_seguromedico.id = pla_categoria4.idsegmed) ON pla_empleados.id = pla_categoria4.idemp "
            If nLista <> "" Then
                nWhere = nWhere + IIf(nWhere = "", "", " AND ") & Replace(nLista, "CAMPO_REEMPLAZA", "pla_empleados.id")
            Else
                nWhere = IIf(nWhere = "", "", nWhere & " AND ") & " pla_empleados.idcat <> 6 "
            End If
            nGroupBy = ""
            nOrderBy = " pla_empleados.apepat;"
        
        Case 6 '--Datos del Personal de Terceros
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 10:        Q_POSICION_TOTAL = -1:
    
            nCampos = "  pla_empleados.id, mae_dociden.codsun AS e_tipodoc, mae_dociden.descripcion AS tipodoc, pla_empleados.numdoc AS e_numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_empresadestaca.numruc AS e_empresadestaca, mae_empresadestaca.descripcion AS empresadestaca, mae_sctrsalud.codsun AS e_sctrsalud, mae_sctrsalud.descripcion AS sctrsalud, mae_sctrpension.codsun AS e_sctrpension, mae_sctrpension.descripcion AS sctrpension "
            nFrom = "  (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) INNER JOIN (mae_sctrsalud RIGHT JOIN (mae_sctrpension RIGHT JOIN (mae_empresadestaca RIGHT JOIN pla_categoria5 ON mae_empresadestaca.id = pla_categoria5.iddestaca) ON mae_sctrpension.id = pla_categoria5.sctrpension) ON mae_sctrsalud.id = pla_categoria5.sctrsalud) ON pla_empleados.id = pla_categoria5.idemp "
            If nLista <> "" Then
                nWhere = nWhere + IIf(nWhere = "", "", " AND ") & Replace(nLista, "CAMPO_REEMPLAZA", "pla_empleados.id")
            Else
                nWhere = IIf(nWhere = "", "", nWhere & " AND ") & " pla_empleados.idcat <> 6 "
            End If
            nGroupBy = ""
            nOrderBy = " pla_empleados.apepat;"
        Case 7 '--periodos laborales
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 12:        Q_POSICION_TOTAL = -1:
    
            nCampos = " [pla_periodolaboral].[idemp] & '-' & [pla_periodolaboral].[corr] AS id, mae_dociden.codsun AS e_tipodoc, mae_dociden.descripcion AS tipodoc, pla_empleados.numdoc AS e_numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_categoria.codsun AS e_categoria, mae_categoria.descripcion AS categoria, Format([pla_periodolaboral].[fchini],'dd/mm/yyyy') AS e_fchini, Format([pla_periodolaboral].[fchfin],'dd/mm/yyyy') AS e_fchfin, mae_finperiodo.codsun AS e_finperiodo, mae_finperiodo.descripcion AS finperiodo, mae_tipomodformativa.codsun AS e_tipomodformativa, mae_tipomodformativa.descripcion AS tipomodformativa "
            nFrom = " mae_categoria RIGHT JOIN (mae_tipomodformativa RIGHT JOIN (mae_finperiodo RIGHT JOIN ((mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) INNER JOIN pla_periodolaboral ON pla_empleados.id = pla_periodolaboral.idemp) ON mae_finperiodo.id = pla_periodolaboral.idfinper) ON mae_tipomodformativa.id = pla_periodolaboral.idmodfor) ON mae_categoria.id = pla_periodolaboral.idcat "
            If nLista <> "" Then
                nWhere = nWhere + IIf(nWhere = "", "", " AND ") & Replace(nLista, "CAMPO_REEMPLAZA", "[pla_periodolaboral].[idemp] & '-' & [pla_periodolaboral].[corr]")
            Else
                nWhere = IIf(nWhere = "", "", nWhere & " AND ") & " pla_empleados.idcat <> 6 "
            End If
            nGroupBy = ""
            nOrderBy = "  [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom], pla_periodolaboral.fchini ; "
        
        Case 8 '--Datos de derechohabientes
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 38:        Q_POSICION_TOTAL = -1:
    
            nCampos = "  pla_empleados.id & '-' & pla_derechohab.corr AS id, mae_dociden.codsun AS e_tipodoc, mae_dociden.descripcion AS tipodoc, pla_empleados.numdoc AS e_numdoc, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS nombres, mae_dociden_1.codsun AS e_tipodoc1, mae_dociden_1.descripcion AS tipodoc1, pla_derechohab.numdoc AS e_numdoc1, pla_derechohab.apepat AS e_apepat, pla_derechohab.apemat AS e_apemat, pla_derechohab.nombre AS e_nombre, Format(pla_derechohab.fchnac,'dd/mm/yyyy') AS e_fchnac, mae_sexo.id AS e_sexo, mae_sexo.descripcion AS sexo, " _
                + vbCr + " mae_vinculofam.codsun AS e_vinculofamiliar, mae_vinculofam.descripcion AS vinculofamiliar, mae_docacrepat.codsun AS e_tipodocpeternidad, mae_docacrepat.descripcion AS tipodocpeternidad, pla_derechohab.numdocpat AS e_numdocpat, mae_situacionderhab.codsun AS e_situacion, mae_situacionderhab.descripcion AS situacion, Format([fchalt],'dd/mm/yyyy') AS e_fchalt, mae_tipobaja.codsun AS e_tipobaja, mae_tipobaja.descripcion AS tipobaja, Format([fchbaj],'dd/mm/yyyy') AS e_fchbaj, pla_derechohab.numresinc AS e_numresolucion, " _
                + vbCr + " mae_indicadomderhab.codsun AS e_indicadordomicilio, mae_indicadomderhab.descripcion AS indicadordomicilio, mae_tipovia.codsun AS e_tipovia, mae_tipovia.descripcion AS tipovia, pla_derechohab.nomvia AS e_nomvia, pla_derechohab.numvia AS e_numvia, pla_derechohab.intvia AS e_intvia, mae_tipozona.codsun AS e_tipozona, mae_tipozona.descripcion AS tipozona, pla_derechohab.nomzon AS e_nomzon, pla_derechohab.refdom AS e_refdom, mae_distrito.codsun AS e_ubigeo, [mae_departamento].[descripcion] & ' - ' & [mae_provincia].[descripcion] & ' - ' & [mae_distrito].[descripcion] AS ubigeo "
            nFrom = "   (mae_dociden INNER JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) INNER JOIN (mae_vinculofam RIGHT JOIN (mae_tipozona RIGHT JOIN (mae_tipovia RIGHT JOIN (mae_situacionderhab RIGHT JOIN (mae_sexo RIGHT JOIN ((mae_departamento RIGHT JOIN mae_provincia ON mae_departamento.id = mae_provincia.iddepa) RIGHT JOIN (mae_indicadomderhab RIGHT JOIN (mae_distrito RIGHT JOIN (mae_tipobaja RIGHT JOIN (mae_docacrepat RIGHT JOIN (pla_derechohab LEFT JOIN mae_dociden AS mae_dociden_1 ON pla_derechohab.idtipdoc = mae_dociden_1.id)  " _
                + vbCr + " ON mae_docacrepat.id = pla_derechohab.idtipdocpat) ON mae_tipobaja.id = pla_derechohab.idtipbaj) ON mae_distrito.id = pla_derechohab.iddis) ON mae_indicadomderhab.id = pla_derechohab.idinddom) ON mae_provincia.id = mae_distrito.idprov) ON mae_sexo.id = pla_derechohab.idsex) ON mae_situacionderhab.id = pla_derechohab.idsitderhab) ON mae_tipovia.id = pla_derechohab.idtipvia) ON mae_tipozona.id = pla_derechohab.idtipzon) ON mae_vinculofam.id = pla_derechohab.idvinfam) ON pla_empleados.id = pla_derechohab.idemp "
            If nLista <> "" Then
                nWhere = nWhere + IIf(nWhere = "", "", " AND ") & Replace(nLista, "CAMPO_REEMPLAZA", "pla_empleados.id & '-' & pla_derechohab.corr")
            Else
                nWhere = IIf(nWhere = "", "", nWhere & " AND ") & " pla_empleados.idcat <> 6 "
            End If
            nGroupBy = ""
            nOrderBy = " pla_empleados.apepat;"
        
        Case Else
            
    End Select
    
    Q_COL_FILA = Q_COL_FILA + 1 '(1= primera col para seleccionar)
    '--DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    
    '------------------------------------------

    '--GENERANDO LA CONSULTA
    nSQL = "SELECT -1 as sel, " + nCampos + _
        vbCr + " FROM " + nFrom + _
        IIf(nWhere <> "", vbCr + " WHERE ", "") + nWhere + _
        IIf(nGroupBy <> "", vbCr + " GROUP BY ", "") + nGroupBy + _
        vbCr + " ORDER BY " + nOrderBy

    '------------------------------------------------------------------------------------
    fGenerarConsulta = nSQL
    
    Set RstTmp = Nothing
End Function

Private Sub pConfigurarGrilla()
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
    Dim M_ANCHO_COL As Integer '--DEPENDERA DEL TIPO DE CONSULTA
                                   
    Dim K, J As Integer
    Dim T_CONSULTA As Integer
    
    Fg1.Clear
    
    Fg1.FrozenCols = 0
    
    M_ANCHO_COL = 0

    With Fg1
        '-----
        Fg1.Cols = Q_COL_FILA_OCULTA + Q_COL_FILA
                 
        .ColWidth(0) = 200
        '--DATOS DE FILA
        Select Case cb_sel.ListIndex
            Case 0 '-- datos principales
                .Rows = 2
                .FixedRows = 2
                .RowHeight(0) = 350
                UNIR_CELDAS Fg1, 0, 3, 0, 5, "Tipo de Documento", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 6, 0, 8, "Nombres Completos", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 9, 1, 9, "Fecha de " + vbCr + "Nacimiento", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 0, 10, 0, 11, "Sexo", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 12, 0, 13, "Nacionalidad", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 14, 0, 14, "Teléfono", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 0, 15, 0, 15, "E-Mail", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 0, 16, 0, 18, "Indicador ESSALUD + Vida", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 19, 0, 20, "Indicador de domiciliado", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 21, 0, 25, "Tipo de Via", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 26, 0, 29, "Tipo de Zona", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 30, 0, 31, "Ubicación Geográfica", flexAlignCenterCenter
                .RowHeight(1) = 250
                '--tipo documento
                .TextMatrix(1, 3) = "Cod":                  .ColWidth(3) = 450:     .ColAlignment(3) = flexAlignCenterCenter:       .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 4) = "Descripción":          .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftCenter:         .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 5) = "N°.Doc.":              .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignLeftCenter:        .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 6) = "Ap. Paterno":          .ColWidth(6) = 1000:    .ColAlignment(6) = flexAlignLeftCenter:         .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 7) = "Ap. Materno":          .ColWidth(7) = 1200:    .ColAlignment(7) = flexAlignLeftCenter:         .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 8) = "Nombres":              .ColWidth(8) = 1800:    .ColAlignment(8) = flexAlignLeftCenter:         .Row = 1: .Col = 7: .CellAlignment = flexAlignLeftCenter
                '--fecha nac
                .ColWidth(9) = 1100:     .ColAlignment(9) = flexAlignCenterBottom:       .Row = 1: .Col = 9: .CellAlignment = flexAlignCenterBottom
                '--sexo
                .TextMatrix(1, 10) = "Cod":                 .ColWidth(10) = 450:       .ColAlignment(10) = flexAlignCenterCenter:    .Row = 1: .Col = 10: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 11) = "Descripción":         .ColWidth(11) = 1000:      .ColAlignment(11) = flexAlignLeftCenter:      .Row = 1: .Col = 11: .CellAlignment = flexAlignLeftCenter
                '--nacionalidad
                .TextMatrix(1, 12) = "Cod":                 .ColWidth(12) = 550:       .ColAlignment(12) = flexAlignCenterCenter:    .Row = 1: .Col = 12: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 13) = "Descripción":         .ColWidth(13) = 1000:      .ColAlignment(13) = flexAlignLeftCenter:      .Row = 1: .Col = 13: .CellAlignment = flexAlignLeftCenter
                '--telefono
                .ColWidth(14) = 1300:   .ColAlignment(14) = flexAlignLeftCenter:          .Row = 1: .Col = 14: .CellAlignment = flexAlignCenterCenter
                '--email
                .ColWidth(15) = 1200:   .ColAlignment(15) = flexAlignLeftCenter:          .Row = 1: .Col = 15: .CellAlignment = flexAlignCenterCenter
                '--Indicador ESSALUD + Vida
                .TextMatrix(1, 16) = "Cod":                     .ColWidth(16) = 450:    .ColAlignment(16) = flexAlignCenterCenter:    .Row = 1: .Col = 16: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 17) = "Descripción":             .ColWidth(17) = 2100:   .ColAlignment(17) = flexAlignLeftCenter:      .Row = 1: .Col = 17: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 18) = "Número":                  .ColWidth(18) = 1500:   .ColAlignment(18) = flexAlignLeftCenter:      .Row = 1: .Col = 18: .CellAlignment = flexAlignLeftCenter
                '--Indicador de domiciliado
                .TextMatrix(1, 19) = "Cod":                     .ColWidth(19) = 450:    .ColAlignment(19) = flexAlignCenterCenter:    .Row = 1: .Col = 19: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 20) = "Descripción":             .ColWidth(20) = 1400:   .ColAlignment(20) = flexAlignLeftCenter:      .Row = 1: .Col = 20: .CellAlignment = flexAlignLeftCenter
                '--tipo de via
                .TextMatrix(1, 21) = "Cod":                     .ColWidth(21) = 450:    .ColAlignment(21) = flexAlignCenterCenter:    .Row = 1: .Col = 21: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 22) = "Descripción":             .ColWidth(22) = 1200:   .ColAlignment(22) = flexAlignLeftCenter:      .Row = 1: .Col = 22: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 23) = "Nombre":                  .ColWidth(23) = 1500:   .ColAlignment(23) = flexAlignLeftCenter:      .Row = 1: .Col = 23: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 24) = "Núm":                     .ColWidth(24) = 600:    .ColAlignment(24) = flexAlignRightCenter:     .Row = 1: .Col = 24: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 25) = "Int":                     .ColWidth(25) = 450:    .ColAlignment(25) = flexAlignRightCenter:     .Row = 1: .Col = 25: .CellAlignment = flexAlignRightCenter
                '--tipo de zona
                .TextMatrix(1, 26) = "Cod":                     .ColWidth(26) = 450:    .ColAlignment(26) = flexAlignCenterCenter:    .Row = 1: .Col = 26: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 27) = "Descripción":             .ColWidth(27) = 1600:   .ColAlignment(27) = flexAlignLeftCenter:      .Row = 1: .Col = 27: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 28) = "Nombre":                  .ColWidth(28) = 1400:   .ColAlignment(28) = flexAlignLeftCenter:      .Row = 1: .Col = 28: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 29) = "Referencia":              .ColWidth(29) = 1800:   .ColAlignment(29) = flexAlignLeftCenter:      .Row = 1: .Col = 29: .CellAlignment = flexAlignLeftCenter
                '--ubigeo
                .TextMatrix(1, 30) = "Cod":                     .ColWidth(30) = 650:    .ColAlignment(30) = flexAlignCenterCenter:    .Row = 1: .Col = 30: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 31) = "Descripción":             .ColWidth(31) = 3800:   .ColAlignment(31) = flexAlignLeftCenter:      .Row = 1: .Col = 31: .CellAlignment = flexAlignLeftCenter
                
                .FrozenCols = 8
            Case 1 '--Datos Trabajador
                .Rows = 2
                .FixedRows = 2
                .RowHeight(0) = 700
                UNIR_CELDAS Fg1, 0, 3, 0, 6, "Información del Personal", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 7, 0, 8, "Tipo de Trabajador", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 9, 0, 10, "Régimen Laboral", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 11, 0, 12, "Nivel Educativo", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 13, 0, 14, "Ocupación", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 15, 0, 16, "¿Es Discapacitado?", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 17, 0, 20, "Régimen Pensionario", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 21, 0, 22, "SCTR Salud", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 23, 0, 24, "SCTR Pensión", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 25, 0, 26, "Tipo Contrato", flexAlignCenterCenter
                'Trabajador sujeto a régimen alternativo, acumulativo o atípico de jornada de trabajo y descanso (1 = si; 0=  No)
                UNIR_CELDAS Fg1, 0, 27, 0, 28, "Rég. Aleternativo acumulativo " + vbCr + "o atipico de jornada de" + vbCr + "trabajo y descanso", flexAlignCenterCenter
                'Trabajador sujeto a jornada de trabajo máxima (1 = si; 0=  No)'--opcion 2
                UNIR_CELDAS Fg1, 0, 29, 0, 30, "Jornada de Trabajo" + vbCr + "Máxima", flexAlignCenterCenter
                'Trabajador sujeto a horario nocturno (1 = si; 0=  No) 'opcion 3
                UNIR_CELDAS Fg1, 0, 31, 0, 32, "Horario Nocturno", flexAlignCenterCenter
                'Tiene otros ingresos de Quinta Categoría (1 = si; 0=  No) '--Opción 4
                UNIR_CELDAS Fg1, 0, 33, 0, 34, "Tiene Ingresos" + vbCr + "5ta. Categoría", flexAlignCenterCenter
                'Es sindicalizado (1 = si; 0=  No) '--Opción 5
                UNIR_CELDAS Fg1, 0, 35, 0, 36, "Es" + vbCr + "Sindicalizado", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 37, 0, 38, "Periodo de Pago", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 39, 0, 44, "Prestaciones de Salud", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 43, 0, 44, "Situación del Trabajador", flexAlignCenterCenter
                'Indicador de rentas de quinta categoría exoneradas o inafectas (1 = si; 0=  No) '--opcion 7
                'Afiliado a EPS/Servicios Propios (1 = si; 0=  No)
                UNIR_CELDAS Fg1, 0, 45, 0, 46, "Afiliado a" + vbCr + "EPS/Servicios" + vbCr + "Propios", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 47, 0, 48, "Situación Especial", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 49, 0, 50, "Tipo de Pago", flexAlignCenterCenter
                .RowHeight(1) = 250
                '--datos del personal
                .TextMatrix(1, 3) = "Cod":                  .ColWidth(3) = 450:     .ColAlignment(3) = flexAlignCenterCenter: .Row = 1: .Col = 3:  .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 4) = "Descripción":          .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftCenter:   .Row = 1: .Col = 4:  .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 5) = "N°.Doc.":              .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignLeftCenter:   .Row = 1: .Col = 5:  .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 6) = "Nombres":              .ColWidth(6) = 3000:    .ColAlignment(6) = flexAlignLeftCenter:   .Row = 1: .Col = 6:  .CellAlignment = flexAlignLeftCenter
                '--tipo de trabajdor
                .TextMatrix(1, 7) = "Cod":                  .ColWidth(7) = 450:     .ColAlignment(7) = flexAlignCenterCenter: .Row = 1: .Col = 7:  .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 8) = "Descripción":          .ColWidth(8) = 2800:    .ColAlignment(8) = flexAlignLeftCenter:   .Row = 1: .Col = 8:  .CellAlignment = flexAlignLeftCenter
                '--regimen laboral
                .TextMatrix(1, 9) = "Cod":                  .ColWidth(9) = 450:     .ColAlignment(9) = flexAlignCenterCenter: .Row = 1: .Col = 9:  .CellAlignment = flexAlignCenterBottom
                .TextMatrix(1, 10) = "Descripción":         .ColWidth(10) = 1000:   .ColAlignment(10) = flexAlignLeftCenter:  .Row = 1: .Col = 10: .CellAlignment = flexAlignLeftCenter
                '--nivel educativo
                .TextMatrix(1, 11) = "Cod":                 .ColWidth(11) = 450:    .ColAlignment(11) = flexAlignCenterCenter: .Row = 1: .Col = 11: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 12) = "Descripción":         .ColWidth(12) = 2500:   .ColAlignment(12) = flexAlignLeftCenter:   .Row = 1: .Col = 12: .CellAlignment = flexAlignLeftCenter
                '--Ocupacion
                .TextMatrix(1, 13) = "Cod":                 .ColWidth(13) = 700:    .ColAlignment(13) = flexAlignCenterCenter: .Row = 1: .Col = 13: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 14) = "Descripción":         .ColWidth(14) = 2800:   .ColAlignment(14) = flexAlignLeftCenter:   .Row = 1: .Col = 14: .CellAlignment = flexAlignCenterCenter
                '--es discapacitado
                .TextMatrix(1, 15) = "Cod":                 .ColWidth(15) = 450:    .ColAlignment(15) = flexAlignCenterCenter: .Row = 1: .Col = 15: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 16) = "Estado":              .ColWidth(16) = 900:    .ColAlignment(16) = flexAlignLeftCenter:   .Row = 1: .Col = 16: .CellAlignment = flexAlignCenterCenter
                '--regimen pensionario
                .TextMatrix(1, 17) = "Cod":                 .ColWidth(17) = 450:    .ColAlignment(17) = flexAlignCenterCenter: .Row = 1: .Col = 17: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 18) = "Descripción":         .ColWidth(18) = 1500:   .ColAlignment(18) = flexAlignLeftCenter:   .Row = 1: .Col = 18: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 19) = "Inscripción":         .ColWidth(19) = 1000:   .ColAlignment(19) = flexAlignCenterCenter: .Row = 1: .Col = 19: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 20) = "CUSPP":               .ColWidth(20) = 1300:   .ColAlignment(20) = flexAlignLeftCenter:   .Row = 1: .Col = 20: .CellAlignment = flexAlignCenterCenter
                '--SCTR Salud
                .TextMatrix(1, 21) = "Cod":                 .ColWidth(21) = 450:    .ColAlignment(21) = flexAlignCenterCenter: .Row = 1: .Col = 21: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 22) = "Descripción":         .ColWidth(22) = 1000:   .ColAlignment(22) = flexAlignLeftCenter:   .Row = 1: .Col = 22: .CellAlignment = flexAlignLeftCenter
                '--SCTR Pension
                .TextMatrix(1, 23) = "cod":                 .ColWidth(23) = 450:    .ColAlignment(23) = flexAlignCenterCenter: .Row = 1: .Col = 23: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 24) = "Descripción":         .ColWidth(24) = 1500:   .ColAlignment(24) = flexAlignLeftCenter:   .Row = 1: .Col = 24: .CellAlignment = flexAlignLeftCenter
                '--tipo Contrato
                .TextMatrix(1, 25) = "Cod":                 .ColWidth(25) = 450:    .ColAlignment(25) = flexAlignCenterCenter: .Row = 1: .Col = 25: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 26) = "Descripción":         .ColWidth(26) = 1400:   .ColAlignment(26) = flexAlignLeftCenter:   .Row = 1: .Col = 26: .CellAlignment = flexAlignLeftCenter
                '--opcion 1'--sujeto a régimen alternativo, acumulativo o atípico de jornada de trabajo y descanso
                .TextMatrix(1, 27) = "Cod":                 .ColWidth(27) = 450:    .ColAlignment(27) = flexAlignCenterCenter: .Row = 1: .Col = 27: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 28) = "Estado":              .ColWidth(28) = 1800:   .ColAlignment(28) = flexAlignCenterCenter: .Row = 1: .Col = 28: .CellAlignment = flexAlignCenterCenter
                '--opcion 2 '--jornada de trabajo maxima
                .TextMatrix(1, 29) = "Cod":                 .ColWidth(29) = 450:   .ColAlignment(29) = flexAlignCenterCenter:  .Row = 1: .Col = 29: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 30) = "Estado":              .ColWidth(30) = 1100:  .ColAlignment(30) = flexAlignCenterCenter:  .Row = 1: .Col = 30: .CellAlignment = flexAlignCenterCenter
                '--opcion 3 'Horario nocturno
                .TextMatrix(1, 31) = "Cod":                 .ColWidth(31) = 450:   .ColAlignment(31) = flexAlignCenterCenter:  .Row = 1: .Col = 31: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 32) = "Estado":              .ColWidth(32) = 900:   .ColAlignment(32) = flexAlignCenterCenter:  .Row = 1: .Col = 32: .CellAlignment = flexAlignLeftCenter
                '--opcion 4 ' Tiene Ingresos 5ta. Categoría
                .TextMatrix(1, 33) = "Cod":                 .ColWidth(33) = 450:   .ColAlignment(33) = flexAlignCenterCenter:  .Row = 1: .Col = 33: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 34) = "Estado":              .ColWidth(34) = 900:   .ColAlignment(34) = flexAlignCenterCenter:  .Row = 1: .Col = 34: .CellAlignment = flexAlignLeftCenter
                '--opcion 5 '--es Sindicalizado
                .TextMatrix(1, 35) = "Cod":                 .ColWidth(35) = 450:   .ColAlignment(35) = flexAlignCenterCenter:  .Row = 1: .Col = 35: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 36) = "Esado":               .ColWidth(36) = 800:   .ColAlignment(36) = flexAlignCenterCenter:  .Row = 1: .Col = 36: .CellAlignment = flexAlignLeftCenter
                '--periodo de pago
                .TextMatrix(1, 37) = "Cod":                 .ColWidth(37) = 450:   .ColAlignment(37) = flexAlignLeftCenter:    .Row = 1: .Col = 37: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 38) = "Descripción":         .ColWidth(38) = 1600:  .ColAlignment(38) = flexAlignLeftCenter:    .Row = 1: .Col = 38: .CellAlignment = flexAlignLeftCenter
                '--Prestaciones de Salud (opcion 7)
                .TextMatrix(1, 39) = "Cod":                 .ColWidth(39) = 450:   .ColAlignment(39) = flexAlignCenterCenter:  .Row = 1: .Col = 39: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 40) = "Estado":              .ColWidth(40) = 700:   .ColAlignment(40) = flexAlignCenterCenter:  .Row = 1: .Col = 40: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 41) = "Cod":                 .ColWidth(41) = 450:   .ColAlignment(41) = flexAlignCenterCenter:  .Row = 1: .Col = 41: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 42) = "Descripción":         .ColWidth(42) = 2500:  .ColAlignment(42) = flexAlignLeftCenter:    .Row = 1: .Col = 42: .CellAlignment = flexAlignLeftCenter
                '--Situacion del Trabajador
                .TextMatrix(1, 43) = "Cod":                 .ColWidth(43) = 450:   .ColAlignment(43) = flexAlignCenterCenter:  .Row = 1: .Col = 43: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 44) = "Descripción":         .ColWidth(44) = 2200:  .ColAlignment(44) = flexAlignLeftCenter:    .Row = 1: .Col = 44: .CellAlignment = flexAlignCenterCenter
                '--Opcion 6
                .TextMatrix(1, 45) = "Cod":                 .ColWidth(45) = 450:   .ColAlignment(45) = flexAlignCenterCenter:  .Row = 1: .Col = 45: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 46) = "Estado":              .ColWidth(46) = 800:   .ColAlignment(46) = flexAlignCenterCenter:  .Row = 1: .Col = 46: .CellAlignment = flexAlignCenterCenter
                '--situacion del trabajo
                .TextMatrix(1, 47) = "Cod":                 .ColWidth(47) = 450:   .ColAlignment(47) = flexAlignCenterCenter:  .Row = 1: .Col = 47: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 48) = "Descripción":         .ColWidth(48) = 2200:  .ColAlignment(48) = flexAlignLeftCenter:    .Row = 1: .Col = 48: .CellAlignment = flexAlignCenterCenter
                '--tipo de pago
                .TextMatrix(1, 49) = "Cod":                 .ColWidth(49) = 450:   .ColAlignment(49) = flexAlignCenterCenter:  .Row = 1: .Col = 49: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 50) = "Descripción":         .ColWidth(50) = 2000:  .ColAlignment(50) = flexAlignLeftCenter:    .Row = 1: .Col = 50: .CellAlignment = flexAlignCenterCenter
                .FrozenCols = 6
        
            Case 2 '--Datos pensionista
                .Rows = 2
                .FixedRows = 2
                .RowHeight(0) = 350
                UNIR_CELDAS Fg1, 0, 3, 0, 6, "Información del Personal", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 7, 0, 8, "Tipo de Pensionista", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 9, 0, 12, "Régimen Pensionario", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 13, 0, 14, "Situación del Pensionista", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 15, 0, 16, "Tipo de Pago", flexAlignCenterCenter
                                
                .RowHeight(1) = 250
                '--datos del personal
                .TextMatrix(1, 3) = "Cod":                  .ColWidth(3) = 450:     .ColAlignment(3) = flexAlignCenterCenter:  .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 4) = "Descripción":          .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftCenter:    .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 5) = "N°.Doc.":              .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignLeftCenter:    .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 6) = "Nombres":              .ColWidth(6) = 3000:    .ColAlignment(6) = flexAlignLeftCenter:    .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
                '--tipo de pensionista
                .TextMatrix(1, 7) = "Cod":                  .ColWidth(7) = 450:     .ColAlignment(7) = flexAlignCenterCenter:  .Row = 1: .Col = 7: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 8) = "Descripción":          .ColWidth(8) = 2800:    .ColAlignment(8) = flexAlignLeftCenter:    .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftCenter
                '--regimen pensionario
                .TextMatrix(1, 9) = "Cod":                  .ColWidth(9) = 450:     .ColAlignment(9) = flexAlignCenterCenter:  .Row = 1: .Col = 9: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 10) = "Descripción":         .ColWidth(10) = 1500:   .ColAlignment(10) = flexAlignLeftCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 11) = "Inscripción":         .ColWidth(11) = 1000:   .ColAlignment(11) = flexAlignCenterCenter: .Row = 1: .Col = 11: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 12) = "CUSPP":               .ColWidth(12) = 1300:   .ColAlignment(12) = flexAlignLeftCenter:   .Row = 1: .Col = 12: .CellAlignment = flexAlignCenterCenter
                '--Situacion del Trabajador
                .TextMatrix(1, 13) = "Cod":                 .ColWidth(13) = 450:    .ColAlignment(13) = flexAlignCenterCenter: .Row = 1: .Col = 13: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 14) = "Descripción":         .ColWidth(14) = 2200:   .ColAlignment(14) = flexAlignLeftCenter:   .Row = 1: .Col = 14: .CellAlignment = flexAlignCenterCenter
                '--tipo de pago
                .TextMatrix(1, 15) = "Cod":                 .ColWidth(15) = 450:    .ColAlignment(15) = flexAlignCenterCenter: .Row = 1: .Col = 15: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 16) = "Descripción":         .ColWidth(16) = 2000:   .ColAlignment(16) = flexAlignLeftCenter:   .Row = 1: .Col = 16: .CellAlignment = flexAlignCenterCenter
                
                .FrozenCols = 6
            Case 3 '--prestador de servicio 4ta categoria
                .Rows = 2
                .FixedRows = 2
                .RowHeight(0) = 350
                .RowHeight(1) = 250
                UNIR_CELDAS Fg1, 0, 3, 0, 6, "Información del Personal", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 7, 1, 7, "N°. RUC", flexAlignCenterCenter, False
                '--datos del personal
                .TextMatrix(1, 3) = "Cod":              .ColWidth(3) = 450:     .ColAlignment(3) = flexAlignCenterCenter:       .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 4) = "Descripción":      .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftCenter:         .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 5) = "N°.Doc.":          .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignLeftCenter:         .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 6) = "Nombres":          .ColWidth(6) = 3500:    .ColAlignment(6) = flexAlignLeftCenter:         .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
                '--num ruc
                .ColWidth(7) = 1200:    .ColAlignment(7) = flexAlignCenterCenter:         .Row = 1: .Col = 7: .CellAlignment = flexAlignLeftCenter
                
                .FrozenCols = 7
            Case 4 '--Datos de Suspención de Cuarta Categoría
                .Rows = 3
                .FixedRows = 3
                .RowHeight(0) = 350
                .RowHeight(1) = 250
                .RowHeight(2) = 250
                UNIR_CELDAS Fg1, 0, 3, 0, 6, "Información del Personal", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 7, 2, 7, "N°. RUC", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 0, 8, 0, 12, "Solicitudes de Suspensión", flexAlignCenterCenter
                
                '--datos del personal
                UNIR_CELDAS Fg1, 1, 3, 2, 3, "Cod", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 4, 2, 4, "Descripción", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 5, 2, 5, "N°. Doc.", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 6, 2, 6, "Nombres", flexAlignCenterCenter, False
                
                .ColWidth(3) = 450:     .ColAlignment(3) = flexAlignCenterCenter:       .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
                .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftCenter:         .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
                .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignLeftCenter:         .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
                .ColWidth(6) = 3000:    .ColAlignment(6) = flexAlignLeftCenter:         .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
                '--num ruc
                .ColWidth(7) = 1200:    .ColAlignment(7) = flexAlignCenterCenter:       .Row = 1: .Col = 7: .CellAlignment = flexAlignLeftCenter
                '--solicitud de suspencion
                UNIR_CELDAS Fg1, 1, 8, 2, 8, "Número de" + vbCr + "Operación", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 9, 2, 9, "Fecha de" + vbCr + "Presentación", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 10, 2, 10, "Ejercicio", flexAlignCenterCenter, False
                
                .ColWidth(8) = 1500:    .ColAlignment(8) = flexAlignLeftCenter:         .Row = 2: .Col = 8: .CellAlignment = flexAlignLeftCenter
                .ColWidth(9) = 1000:    .ColAlignment(9) = flexAlignCenterCenter:       .Row = 2: .Col = 9: .CellAlignment = flexAlignLeftCenter
                .ColWidth(10) = 800:    .ColAlignment(10) = flexAlignCenterCenter:      .Row = 2: .Col = 10: .CellAlignment = flexAlignCenterCenter
                
                UNIR_CELDAS Fg1, 1, 11, 1, 12, "Medio de Presentación", flexAlignCenterCenter
                
                .TextMatrix(2, 11) = "Cod":          .ColWidth(11) = 450:   .ColAlignment(11) = flexAlignCenterCenter:      .Row = 2: .Col = 11:    .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 12) = "Descripción":  .ColWidth(12) = 2000:  .ColAlignment(12) = flexAlignLeftCenter:        .Row = 2: .Col = 12:    .CellAlignment = flexAlignCenterCenter
                  
            Case 5 '--Datos del Prestador de Servicios - Modalidad Formativa
                .Rows = 2
                .FixedRows = 2
                .RowHeight(0) = 500
                UNIR_CELDAS Fg1, 0, 3, 0, 6, "Información del Personal", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 7, 0, 8, "Seguro Médico", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 9, 0, 10, "Nivel Educativo", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 11, 0, 12, "Ocupación", flexAlignCenterCenter
                '--Indicador de madre con responsabilidad Familiar  (1 = si; 0=  No)
                UNIR_CELDAS Fg1, 0, 13, 0, 14, "Madre con Reponsa-" + vbCr + "bilidad Familiar", flexAlignCenterCenter
                '--Indicador si está sujeto a trabajo en horario nocturno  (1 = si; 0=  No)
                UNIR_CELDAS Fg1, 0, 15, 0, 16, "Sujeto a Horario" + vbCr + "Nocturno", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 17, 0, 18, "Centro de Formación", flexAlignCenterCenter
                '--Indicador de discapacidad  (1 = si; 0=  No)
                UNIR_CELDAS Fg1, 0, 19, 0, 20, "Tiene" + vbCr + "Discapacidad", flexAlignCenterCenter
                                
                .RowHeight(1) = 250
                '--datos del personal
                .TextMatrix(1, 3) = "Cod":                  .ColWidth(3) = 450:     .ColAlignment(3) = flexAlignCenterCenter:  .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 4) = "Descripción":          .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftCenter:    .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 5) = "N°.Doc.":              .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignLeftCenter:    .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 6) = "Nombres":              .ColWidth(6) = 3000:    .ColAlignment(6) = flexAlignLeftCenter:    .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
                '--Seguro medico
                .TextMatrix(1, 7) = "Cod":                  .ColWidth(7) = 450:     .ColAlignment(7) = flexAlignCenterCenter:  .Row = 1: .Col = 7: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 8) = "Descripción":          .ColWidth(8) = 1000:    .ColAlignment(8) = flexAlignLeftCenter:    .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftCenter
                '--nivel educativo
                .TextMatrix(1, 9) = "Cod":                  .ColWidth(9) = 450:     .ColAlignment(9) = flexAlignCenterCenter:  .Row = 1: .Col = 9: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 10) = "Descripción":         .ColWidth(10) = 2500:   .ColAlignment(10) = flexAlignLeftCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignLeftCenter
                '--ocupacion
                .TextMatrix(1, 11) = "Cod":                 .ColWidth(11) = 450:   .ColAlignment(11) = flexAlignCenterCenter:  .Row = 1: .Col = 11: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 12) = "Descripción":         .ColWidth(12) = 2500:  .ColAlignment(12) = flexAlignLeftCenter:    .Row = 1: .Col = 12: .CellAlignment = flexAlignLeftCenter
                '--madre con responsabilidad Familiar
                .TextMatrix(1, 13) = "Cod":                 .ColWidth(13) = 450:   .ColAlignment(13) = flexAlignCenterCenter:  .Row = 1: .Col = 13: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 14) = "Esado":               .ColWidth(14) = 1200:  .ColAlignment(14) = flexAlignCenterCenter:  .Row = 1: .Col = 14: .CellAlignment = flexAlignCenterCenter
                '--sujeto a trabajo en horario nocturno
                .TextMatrix(1, 15) = "Cod":                 .ColWidth(15) = 450:   .ColAlignment(15) = flexAlignCenterCenter:  .Row = 1: .Col = 15: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 16) = "Estado":              .ColWidth(16) = 1000:  .ColAlignment(16) = flexAlignCenterCenter:  .Row = 1: .Col = 16: .CellAlignment = flexAlignCenterCenter
                '--centro de formacion
                .TextMatrix(1, 17) = "Cod":                 .ColWidth(17) = 450:   .ColAlignment(17) = flexAlignCenterCenter:  .Row = 1: .Col = 17: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 18) = "Descripción":         .ColWidth(18) = 1600:  .ColAlignment(18) = flexAlignLeftCenter:    .Row = 1: .Col = 18: .CellAlignment = flexAlignLeftCenter
                '--Indicador de discapacidad
                .TextMatrix(1, 19) = "Cod":                 .ColWidth(19) = 450:   .ColAlignment(19) = flexAlignCenterCenter:  .Row = 1: .Col = 19: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 20) = "Esado":               .ColWidth(20) = 1000:  .ColAlignment(20) = flexAlignCenterCenter:  .Row = 1: .Col = 20: .CellAlignment = flexAlignCenterCenter
                
                .FrozenCols = 7
            Case 6 '--Datos del Personal de Terceros
                .Rows = 2
                .FixedRows = 2
                .RowHeight(0) = 350
                UNIR_CELDAS Fg1, 0, 3, 0, 6, "Información del Personal", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 7, 0, 8, "Empresa que Destaca o Desplaza", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 9, 0, 10, "SCTR Salud", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 11, 0, 12, "SCTR Pensión", flexAlignCenterCenter
                                
                .RowHeight(1) = 250
                '--datos del personal
                .TextMatrix(1, 3) = "Cod":                  .ColWidth(3) = 450:     .ColAlignment(3) = flexAlignCenterCenter:       .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 4) = "Descripción":          .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftCenter:         .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 5) = "N°.Doc.":              .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignLeftCenter:         .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 6) = "Nombres":              .ColWidth(6) = 3000:    .ColAlignment(6) = flexAlignLeftCenter:         .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
                '--empresa que destaca o desplaza
                .TextMatrix(1, 7) = "RUC":                  .ColWidth(7) = 1200:    .ColAlignment(7) = flexAlignCenterCenter:       .Row = 1: .Col = 7: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 8) = "Razón Social":         .ColWidth(8) = 4500:    .ColAlignment(8) = flexAlignLeftCenter:         .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftCenter
                '--sctr salud
                .TextMatrix(1, 9) = "Cod":                  .ColWidth(9) = 450:     .ColAlignment(9) = flexAlignCenterCenter:       .Row = 1: .Col = 9: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 10) = "Descripción":         .ColWidth(10) = 1500:   .ColAlignment(10) = flexAlignLeftCenter:        .Row = 1: .Col = 10: .CellAlignment = flexAlignLeftCenter
                '--sctr pension
                .TextMatrix(1, 11) = "Cod":                 .ColWidth(11) = 450:    .ColAlignment(11) = flexAlignCenterCenter:      .Row = 1: .Col = 11: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 12) = "Descripción":         .ColWidth(12) = 1500:   .ColAlignment(12) = flexAlignLeftCenter:        .Row = 1: .Col = 12: .CellAlignment = flexAlignLeftCenter
                
                .FrozenCols = 6
                
            Case 7 '--Datos de periodos laborales
                .Rows = 2
                .FixedRows = 2
                .RowHeight(0) = 500
                UNIR_CELDAS Fg1, 0, 3, 0, 6, "Información del Personal", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 7, 0, 8, "Categoría", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 9, 1, 9, "Fch.Inicio" + vbCr + "o Reinicio", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 0, 10, 1, 10, "Fch.Fin, Cese" + vbCr + "/ Suspensión", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 0, 11, 0, 12, "Tipo de Extinción del Contrato" + vbCr + "(No Considerar Modalidad Formativa)", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 13, 0, 14, "Tipo Convenio" + vbCr + "(Solo Modalidad Formativa)", flexAlignCenterCenter
                
                .RowHeight(1) = 250
                '--datos del personal
                .TextMatrix(1, 3) = "Cod":                  .ColWidth(3) = 450:     .ColAlignment(3) = flexAlignCenterCenter:  .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 4) = "Descripción":          .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftCenter:    .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 5) = "N°.Doc.":              .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignLeftCenter:    .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 6) = "Nombres":              .ColWidth(6) = 3000:    .ColAlignment(6) = flexAlignLeftCenter:    .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
                '--categoría
                .TextMatrix(1, 7) = "Cod":                  .ColWidth(7) = 450:     .ColAlignment(7) = flexAlignCenterCenter:  .Row = 1: .Col = 7: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 8) = "Descripción":          .ColWidth(8) = 3200:    .ColAlignment(8) = flexAlignLeftCenter:    .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftCenter
                '--fecha inicio
                .ColWidth(9) = 1200:     .ColAlignment(9) = flexAlignCenterCenter:  .Row = 1: .Col = 9: .CellAlignment = flexAlignCenterCenter
                '--fecha fin
                .ColWidth(10) = 1200:   .ColAlignment(10) = flexAlignCenterCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignCenterCenter
                '--tipo de extincion de contrato
                .TextMatrix(1, 11) = "Cod":                 .ColWidth(11) = 450:   .ColAlignment(11) = flexAlignCenterCenter:  .Row = 1: .Col = 11: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 12) = "Descripción":         .ColWidth(12) = 2500:  .ColAlignment(12) = flexAlignLeftCenter:    .Row = 1: .Col = 12: .CellAlignment = flexAlignLeftCenter
                '--tipo convenio - modalida formativa
                .TextMatrix(1, 13) = "Cod":                 .ColWidth(13) = 450:   .ColAlignment(13) = flexAlignCenterCenter:  .Row = 1: .Col = 13: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(1, 14) = "Esado":               .ColWidth(14) = 2500:  .ColAlignment(14) = flexAlignCenterCenter:  .Row = 1: .Col = 14: .CellAlignment = flexAlignCenterCenter
                
                .FrozenCols = 6
                
                
            Case 8 '--Datos de derechohabientes
                .Rows = 3
                .FixedRows = 3
                .RowHeight(0) = 300
                .RowHeight(1) = 250
                .RowHeight(2) = 250
                
                UNIR_CELDAS Fg1, 0, 3, 0, 6, "Datos del Personal", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 7, 0, 40, "Datos del Derechohabiente", flexAlignCenterCenter
                '--datos del personal
                UNIR_CELDAS Fg1, 1, 3, 2, 3, "Cod", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 4, 2, 4, "Descripción", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 5, 2, 5, "N°.Doc.", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 6, 2, 6, "Nombres", flexAlignCenterCenter, False
                'datos del derechohabiente
                UNIR_CELDAS Fg1, 1, 7, 1, 9, "Tipo Documento", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 1, 10, 1, 12, "Apellidos y Nombres", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 1, 13, 2, 13, "Fecha de" + vbCr + "Nacimiento", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 14, 1, 15, "Sexo", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 1, 16, 1, 17, "Vínculo Familiar", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 1, 18, 1, 20, "Doc. que acredita la Paternidad", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 1, 21, 1, 23, "Situación Alta", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 1, 24, 1, 26, "Situación Baja", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 1, 27, 2, 27, "Resolución por" + vbCr + "Incapacidad", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 28, 1, 29, "Indicador de Domicilio", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 1, 30, 1, 34, "Tipo Via", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 1, 35, 1, 38, "Tipo Zona", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 1, 39, 1, 40, "Ubicación Geográfica", flexAlignCenterCenter
                                
                .RowHeight(1) = 250
                '************--datos del personal
                .ColWidth(3) = 450:     .ColAlignment(3) = flexAlignCenterCenter:  .Row = 2: .Col = 3: .CellAlignment = flexAlignLeftCenter
                .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftCenter:    .Row = 2: .Col = 4: .CellAlignment = flexAlignLeftCenter
                .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignLeftCenter:    .Row = 2: .Col = 5: .CellAlignment = flexAlignLeftCenter
                .ColWidth(6) = 2500:    .ColAlignment(6) = flexAlignLeftCenter:    .Row = 2: .Col = 6: .CellAlignment = flexAlignLeftCenter
                
                '************--datos del derechohabiente
                '--tipo de documento
                .TextMatrix(2, 7) = "Cod":                  .ColWidth(7) = 450:     .ColAlignment(7) = flexAlignCenterCenter:   .Row = 2: .Col = 7: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 8) = "Descripción":          .ColWidth(8) = 2800:    .ColAlignment(8) = flexAlignLeftCenter:     .Row = 2: .Col = 8: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 9) = "N°.Doc.":              .ColWidth(9) = 1200:    .ColAlignment(9) = flexAlignLeftCenter:     .Row = 2: .Col = 9: .CellAlignment = flexAlignLeftCenter
                '--nombres
                .TextMatrix(2, 10) = "Ap. Paterno":         .ColWidth(10) = 1000:   .ColAlignment(10) = flexAlignLeftCenter:    .Row = 2: .Col = 10: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 11) = "Ap. Materno":         .ColWidth(11) = 1200:   .ColAlignment(11) = flexAlignLeftCenter:    .Row = 2: .Col = 11: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 12) = "Nombres":             .ColWidth(12) = 1400:   .ColAlignment(12) = flexAlignLeftCenter:    .Row = 2: .Col = 12: .CellAlignment = flexAlignLeftCenter
                '--fecha nacimiento
                .ColWidth(13) = 1000:    .ColAlignment(13) = flexAlignCenterCenter: .Row = 2: .Col = 13: .CellAlignment = flexAlignLeftCenter
                '--sexo
                .TextMatrix(2, 14) = "Cod":                 .ColWidth(14) = 450:    .ColAlignment(14) = flexAlignCenterCenter:  .Row = 2: .Col = 14: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(2, 15) = "Descripción":         .ColWidth(15) = 1000:   .ColAlignment(15) = flexAlignLeftCenter:    .Row = 2: .Col = 15: .CellAlignment = flexAlignLeftCenter
                '--Vinculo Familiar
                .TextMatrix(2, 16) = "Cod":                 .ColWidth(16) = 450:    .ColAlignment(16) = flexAlignCenterCenter:  .Row = 2: .Col = 16: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(2, 17) = "Descripción":         .ColWidth(17) = 1200:   .ColAlignment(17) = flexAlignLeftCenter:    .Row = 2: .Col = 17: .CellAlignment = flexAlignLeftCenter
                '--Documento que acredita la Paternidad
                .TextMatrix(2, 18) = "Cod":                 .ColWidth(18) = 450:    .ColAlignment(18) = flexAlignCenterCenter:  .Row = 2: .Col = 18: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(2, 19) = "Descripción":         .ColWidth(19) = 1000:   .ColAlignment(19) = flexAlignLeftCenter:    .Row = 2: .Col = 19: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 20) = "N°.Doc.":             .ColWidth(20) = 1300:   .ColAlignment(20) = flexAlignLeftCenter:    .Row = 2: .Col = 20: .CellAlignment = flexAlignLeftCenter
                '--situacion alta
                .TextMatrix(2, 21) = "Cod":                 .ColWidth(21) = 450:    .ColAlignment(21) = flexAlignCenterCenter:  .Row = 2: .Col = 21: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(2, 22) = "Descripción":         .ColWidth(22) = 1000:   .ColAlignment(22) = flexAlignLeftCenter:    .Row = 2: .Col = 22: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 23) = "Fecha":               .ColWidth(23) = 1100:   .ColAlignment(23) = flexAlignCenterCenter:  .Row = 2: .Col = 23: .CellAlignment = flexAlignCenterCenter
                '--situacion baja
                .TextMatrix(2, 24) = "Cod":                 .ColWidth(24) = 450:    .ColAlignment(24) = flexAlignCenterCenter:  .Row = 2: .Col = 24: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(2, 25) = "Descripción":         .ColWidth(25) = 1000:   .ColAlignment(25) = flexAlignLeftCenter:    .Row = 2: .Col = 25: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 26) = "Fecha":               .ColWidth(26) = 1100:   .ColAlignment(26) = flexAlignCenterCenter:  .Row = 2: .Col = 26: .CellAlignment = flexAlignCenterCenter
                '--Resolución por Incapacidad
                .ColWidth(27) = 1500:    .ColAlignment(27) = flexAlignCenterCenter:  .Row = 2: .Col = 27: .CellAlignment = flexAlignCenterCenter
                '--indicador de domicilio
                .TextMatrix(2, 28) = "Cod":                 .ColWidth(28) = 450:     .ColAlignment(28) = flexAlignCenterCenter: .Row = 2: .Col = 28: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(2, 29) = "Estado":              .ColWidth(29) = 1600:    .ColAlignment(29) = flexAlignLeftCenter:   .Row = 2: .Col = 29: .CellAlignment = flexAlignLeftCenter
                '--tipo de via
                .TextMatrix(2, 30) = "Cod":                 .ColWidth(30) = 450:    .ColAlignment(30) = flexAlignCenterCenter:  .Row = 2: .Col = 30: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(2, 31) = "Descripción":         .ColWidth(31) = 1200:   .ColAlignment(31) = flexAlignLeftCenter:    .Row = 2: .Col = 31: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 32) = "Nombre":              .ColWidth(32) = 1000:   .ColAlignment(32) = flexAlignLeftCenter:    .Row = 2: .Col = 32: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 33) = "Núm":                 .ColWidth(33) = 600:    .ColAlignment(33) = flexAlignRightCenter:   .Row = 2: .Col = 33: .CellAlignment = flexAlignRightCenter
                .TextMatrix(2, 34) = "Int":                 .ColWidth(34) = 450:    .ColAlignment(34) = flexAlignRightCenter:   .Row = 2: .Col = 34: .CellAlignment = flexAlignRightCenter
                '--tipo de zona
                .TextMatrix(2, 35) = "Cod":                 .ColWidth(35) = 450:    .ColAlignment(35) = flexAlignCenterCenter:  .Row = 2: .Col = 35: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(2, 36) = "Descripción":         .ColWidth(36) = 1600:   .ColAlignment(36) = flexAlignLeftCenter:    .Row = 2: .Col = 36: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 37) = "Nombre":              .ColWidth(37) = 1400:   .ColAlignment(37) = flexAlignLeftCenter:    .Row = 2: .Col = 37: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 38) = "Referencia":          .ColWidth(38) = 1800:   .ColAlignment(38) = flexAlignLeftCenter:    .Row = 2: .Col = 38: .CellAlignment = flexAlignLeftCenter
                '--ubigeo
                .TextMatrix(2, 39) = "Cod":                 .ColWidth(39) = 650:    .ColAlignment(39) = flexAlignCenterCenter:  .Row = 2: .Col = 39: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(2, 40) = "Descripción":         .ColWidth(40) = 3800:   .ColAlignment(40) = flexAlignLeftCenter:    .Row = 2: .Col = 40: .CellAlignment = flexAlignCenterCenter
                
                .FrozenCols = 6
        End Select
        .ColDataType(1) = flexDTBoolean
        '--seleccion
        UNIR_CELDAS Fg1, 0, 1, .FixedRows - 1, 1, "Sel", flexAlignCenterCenter, False
        .ColWidth(1) = 400:    .ColAlignment(1) = flexAlignCenterCenter:  .Row = 0: .Col = 1: .CellAlignment = flexAlignCenterCenter
        ''''''''''''
        .MergeCells = flexMergeFixedOnly
        M_ANCHO_COL = 0

        '--DE LOS ID'S
        For K = 2 To Q_COL_FILA_OCULTA + 1
            .TextMatrix(0, K) = "ID" + CStr(K):         .ColWidth(K) = 500
        Next
        If Q_COL_FILA_OCULTA <> -1 Then OCULTAR_COL Fg1, 2, Q_COL_FILA_OCULTA + 1
   
    End With
    DoEvents
End Sub

Private Function fValidarConsulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    '--DE LA FECHA ES NULL
    
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Function
    End If
    
    If cb_sel.ListIndex = -1 Then
        MsgBox "Seleccione que tipo de información quiere mostrar", vbExclamation, xTitulo
        cb_sel.SetFocus
        SendKeys "{F4}"
        Exit Function
    End If
    
    '''''
    fValidarConsulta = True
End Function


Private Sub pExportarMSExcel()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, cb_sel.Text, "", , cb_sel.Text
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportar"
End Sub

Private Sub FraPath_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then pExportarMSExcel
    If Button.Index = 4 Then
        If Fg1.Rows = Fg1.FixedRows Then
            MsgBox "No hay registros para exportar", vbInformation, xTitulo
            Exit Sub
        End If
            
        If fValidarConsulta() = False Then Exit Sub
        If Trim(ARR_FORMATOS(cb_sel.ListIndex, 0)) = "" Then
            MsgBox "Falta especificar la extención del archivo", vbInformation, xTitulo
            Exit Sub
        End If
        pHabilitarBotonPath True
    End If
    If Button.Index = 6 Then Imprimir
    If Button.Index = 8 Then
        Unload Me
        Exit Sub
    End If
End Sub

'****************************************************************************************
Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error
    
    Select Case Index
        Case 0 '--pesonal
            ReDim xCampos(7, 3) As String
            xCampos(0, 0) = "Tipo Doc":             xCampos(0, 1) = "tipodoc":     xCampos(0, 2) = "900":     xCampos(0, 3) = "C"
            xCampos(1, 0) = "N°.Doc.":              xCampos(1, 1) = "numdoc":      xCampos(1, 2) = "900":     xCampos(1, 3) = "C"
            xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombre":      xCampos(2, 2) = "3500":    xCampos(2, 3) = "C"
            xCampos(3, 0) = "Fch.Nac.":             xCampos(3, 1) = "fchnac":      xCampos(3, 2) = "1100":    xCampos(3, 3) = "F"
            xCampos(4, 0) = "Sexo":                 xCampos(4, 1) = "sexo":        xCampos(4, 2) = "500":     xCampos(4, 3) = "C"
            xCampos(5, 0) = "Teléfono":             xCampos(5, 1) = "numtel":      xCampos(5, 2) = "1000":    xCampos(5, 3) = "C"
            xCampos(6, 0) = "E-mail":               xCampos(6, 1) = "email":       xCampos(6, 2) = "1300":    xCampos(6, 3) = "C"
        
            nTitulo = "Buscando Personal"
            
            nSQL = "SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.id as cod,pla_empleados.numdoc, mae_dociden.abrev AS tipodoc, mae_sexo.abrev AS sexo, Format([pla_empleados].[fchnac],'dd/mm/yyyy') AS fchnac, pla_empleados.numtel, pla_empleados.email " _
                + vbCr + " FROM mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex " _
                + vbCr + " WHERE pla_empleados.idcat <> 6 "

    End Select
                
    Dim xRs As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    lbl_cb(Index).ToolTipText = xRs.Fields(1) & "" '--NOMBRE
   
Salir:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
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
            nSQL = "SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.id as cod,pla_empleados.numdoc, mae_dociden.abrev AS tipodoc, mae_sexo.abrev AS sexo, Format([pla_empleados].[fchnac],'dd/mm/yyyy') AS fchnac, pla_empleados.numtel, pla_empleados.email " _
                + vbCr + " FROM mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex " _
                + vbCr + " WHERE  pla_empleados.id  = " & NulosN(txt_cb(Index).Text) & "; "


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

Private Sub pExportarNotePad()
'    On Error GoTo Error
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    If fValidarConsulta() = False Then Exit Sub

    Dim mRow&
    Dim nListaIDS As String
    Dim nNombreArchivo As String
    Dim nPath As String
    nListaIDS = ""
    For mRow = Fg1.FixedRows To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 1)) = "-1" Then
            If ARR_FORMATOS(cb_sel.ListIndex, 1) = "N" Then
                nListaIDS = nListaIDS & NulosC(Fg1.TextMatrix(mRow, 2)) & ","
            Else
                nListaIDS = nListaIDS & "'" & NulosC(Fg1.TextMatrix(mRow, 2)) & "',"
            End If
        End If
    Next
    If nListaIDS <> "" Then nListaIDS = " CAMPO_REEMPLAZA  IN (" + Left(nListaIDS, Len(nListaIDS) - 1) + ") "
    '****************************************************************
    nNombreArchivo = NumRuc & "." & ARR_FORMATOS(cb_sel.ListIndex, 0)
    nPath = txtpath(0).Text
    '****************************************************************
    
    nSQL = fGenerarConsulta(nListaIDS)
    
    RST_Busq RstTmp, nSQL, xCon
    
    pExportarTxt RstTmp, nNombreArchivo, nPath
    MsgBox "El Archivo se exportó correctamente", vbInformation, xTitulo
    
    Exit Sub
error:
    SHOW_ERROR
End Sub


'**************************************************

Private Sub Drive1_Change()
    Dir1.Path = Drive1
End Sub

Private Sub Dir1_Change()
    txtpath(0).Text = Dir1.Path
End Sub

'**************************************************

Private Sub pHabilitarBotonPath(band As Boolean)
    '--TRUE= MUESTRA LA OPCION PARA SELECCIONAR LA RUTA
    Dim K&
    If band = True Then
        FraPath.Top = 1755
        FraPath.Left = 4170
        Drive1_Change
    End If
    FraPath.Visible = band
    For K = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(K).Enabled = Not band
        habilitar cb, Not band
        habilitar txt_cb, Not band
        cb_sel.Enabled = Not band
        
    Next K
    
    txtpath(1).Text = "Archivo: " & NumRuc & "." & ARR_FORMATOS(cb_sel.ListIndex, 0)
    
End Sub


Private Sub CmdPath_Click(Index As Integer)
    Select Case Index
        Case 0 '--aceptar
            If Trim(txtpath(0).Text) = "" Then
                MsgBox "Seleccione la carpeta donde guardará el archivo", vbExclamation, xTitulo
                Exit Sub
            End If
            pHabilitarBotonPath False
            pExportarNotePad
        Case 1 '--cancelar
            pHabilitarBotonPath False
            cb_sel.SetFocus
    End Select
End Sub


Private Sub pic_Click()
    CmdPath_Click 1
End Sub


'*******************************************************************************************************************************
Private Sub pExportarTxt(RstTmp As ADODB.Recordset, _
                        nNombreArchivo As String, _
                        nPath As String, _
                        Optional fIncluirTotalRegistros As Boolean = False)
                        
    '--nNombreArchivo Numero de ruc.extencion
    '--nPath ruta donde se almacenara el archivo
    '--fIncluirTotalRegistros = false no se agrega en primera fila total de registros encontrados
    '--fIncluirTotalRegistros = true  se agrega en primera fila total de registros encontrados
    
    Dim oArchivo As Variant
    Dim tDatos As String
    Dim nRuta As String
    
    Const nSeparador As String = "|"
    Err.Clear
    On Error Resume Next
    If nNombreArchivo = "" Then
        MsgBox "Falta espefificar el nombre del Archivo", vbExclamation, xTitulo
        Exit Sub
    End If
    
    If RstTmp.RecordCount = 0 Then
        MsgBox "No hay registros para exportar", vbExclamation, xTitulo
        Exit Sub
    End If
    
    If nPath <> "" Then
        nRuta = nPath
    Else
        nPath = App.Path
    End If
    
    Dim mCampo&
    Do While Not RstTmp.EOF
        For mCampo = 0 To RstTmp.Fields.Count - 1
            If InStr(UCase(RstTmp.Fields(mCampo).Name), "E_") <> 0 Then
                '--la extencion "E_" se debe a la consulta
                tDatos = tDatos & NulosC(RstTmp.Fields(mCampo)) & nSeparador
            End If
        Next
        tDatos = tDatos & vbCrLf
        RstTmp.MoveNext
    Loop
    '--Eliminar el archivo si existe
    If Mid(nRuta, Len(nRuta)) <> "\" Then
        nRuta = nRuta & "\"
    End If
    If ArchivoExiste(nRuta & nNombreArchivo) = True Then
        Kill nRuta & nNombreArchivo
    End If
    '*************************************
    '--si desea agregar la cantidad de registros en primera fila del archivo
    If fIncluirTotalRegistros = True Then
        tDatos = RstTmp.RecordCount & nSeparador & vbCrLf & tDatos
    End If
    '*************************************
    Set oArchivo = CreateObject("Scripting.FileSystemObject")
    
    oArchivo.OpenTextFile(nRuta & "\" & nNombreArchivo, 8, True, 0).Write tDatos
    
    Set oArchivo = Nothing
    Err.Clear
End Sub


'*******************************************************************************************************************************






