VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmFormatos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7260
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   11910
   Begin VB.Frame fr 
      Caption         =   "Seleccionar"
      Height          =   675
      Index           =   0
      Left            =   4905
      TabIndex        =   16
      Top             =   30
      Width           =   1755
      Begin VB.OptionButton opt_mon 
         Caption         =   "Todo en S/."
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   18
         Top             =   195
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton opt_mon 
         Caption         =   "Todo en $."
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   17
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   6675
      TabIndex        =   12
      Top             =   30
      Width           =   5175
      Begin VB.CommandButton cmd 
         Caption         =   "&Exportar"
         Height          =   420
         Index           =   1
         Left            =   2630
         TabIndex        =   15
         Top             =   165
         Width           =   1125
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Imprimir"
         Height          =   420
         Index           =   2
         Left            =   1390
         TabIndex        =   14
         Top             =   165
         Width           =   1125
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Consultar"
         Height          =   420
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   165
         Width           =   1125
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Salir"
         Height          =   420
         Index           =   3
         Left            =   3870
         TabIndex        =   13
         Top             =   165
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   75
      TabIndex        =   7
      Top             =   30
      Width           =   4785
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec1 
         Height          =   300
         Left            =   735
         TabIndex        =   8
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
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
         Valor           =   "25/09/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec2 
         Height          =   300
         Left            =   3015
         TabIndex        =   9
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
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
         Valor           =   "25/09/2007"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2445
         TabIndex        =   11
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   300
         Width           =   510
      End
   End
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   2115
      TabIndex        =   1
      Top             =   2850
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   2
         Top             =   420
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
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
         Left            =   4275
         TabIndex        =   6
         Top             =   150
         Width           =   1530
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
         Left            =   225
         TabIndex        =   4
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Consulta"
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
         Left            =   1320
         TabIndex        =   3
         Top             =   150
         Width           =   735
      End
      Begin VB.Shape Shape1 
         Height          =   750
         Left            =   90
         Top             =   60
         Width           =   5805
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6435
      Left            =   60
      TabIndex        =   5
      Top             =   780
      Width           =   11805
      _cx             =   20823
      _cy             =   11351
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
      FormatString    =   $"FrmFormatos.frx":0000
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
End
Attribute VB_Name = "FrmFormatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------------
'--  FUNCION QUE RECIBE LOS PARAMETROS RECIBE_LINK_FRM ( TIPO_FORMATO (STRING))
' 3.3  >> Libro de inventarios y Balances - Detalle del saldo de la cuenta 10 - Caja y bancos
' 3.4  >> Libro de inventarios y balances - Detalle del saldo de la cuenta 12 - Clientes
' 3.8  >> Libro de inventarios y balances - Detalle del saldo de la cuenta 20 - Mercaderías y la cuenta 21 - Productos terminados
' 3.13 >> Libro de inventarios y balances - Detalle del saldo de la cuenta 42 - Proveedores
'-------------------------------------------------------------------------------------

Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
'--DE LA IMPRESION
Dim T_RPT_PERIODO As String '--PERIODO DEL REPORTE
Dim T_RPT_TITULO As String  '--TITULO DE REPORTE
'------------

Dim ARR_TMP(1, 1) As Double '--0::PROGRAMADO=>> 0::TOTAL,1::TOTAL GEN
                            '--1::TEORICO=>> 0::TOTAL,1::TOTAL GEN
                            '--2::REAL=>> 0::TOTAL,1::TOTAL GEN
                            '--3::DIF=>> 0::TOTAL,1::TOTAL GEN

Dim Q_COL_FILA As Integer   '--INDICA LA CANTIDAD DE COLUMNAS ANTES DE LOS MESES
                            '--EJ. IDCLI,CLIENTE => Q_COL_FILA=2
                            '--    IDCLI,ID_PTO_VTA,CLIENTE,PTO_VENTA => Q_COL_FILA=4
                            
                            
Dim Q_POS_MES As Integer    '--INDICA LA POSICION DEL MES, ESTO CAMBIA
                            '--UTIL PARA COLOCAR LOS DATOS EN EL GRID

Dim Q_COL_FILA_OCULTA As Integer '--INDICA LAS COLUMNAS QUE CONTENDRAN LOS ID'S, ESTOS SE OCULTARAN
                                '-- -1 NO SE OCULTA, <> -1 SE PROCEDE A ACULTAR
                                'EJ. CLIENTE  vta_ventas.idcli,
                                    'PUNTO DE VENTA vta_guia.idpunven
                                    'PRODUCTO   alm_inventario.tippro
                                    'ITEM       alm_inventario.id
                                    'EMPLEADO   vta_ventas.idven

Dim Q_POSICION_TOTAL  As Integer '--INDICA LA POCISION DE LA COLUMNA DONDE SE COLOCARA EL NOMBRE DEL TOTAL Y TOTAL_GRL
                                 '--OBTENDRA VALOR EN GENERAR_CONSULTA()

Dim Q_COL_COMPARAR_GRUPO As Integer '--INDICA LA COLUMNA PARA COMPARAR EL GRUPO
                                    '--OBTENDRA VALOR EN GENERAR_CONSULTA()

'------------
'-------
'------------
Dim TOTAL_REGISTROS As Long '--INDICA LA CANTIDAD DE REGISTROS DESCUADRADOS

Dim TIPO_FORMATO As String '--INDICA EL FORMATO DE LA CONSULTA
Dim ARR_FORMATOS(10) As String '--INDICA LA LISTA DE FORMATOS


Public Sub RECIBE_LINK_FRM(Optional formato As String = "3.3")
'--FrmFormatos.RECIBE_LINK_FRM
    TIPO_FORMATO = formato
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
    End If
    Select Case TIPO_FORMATO
        Case "3.3"
            Me.Caption = "Libro de inventarios y Balances - Detalle del saldo de la cuenta 10 - Caja y bancos"
        Case "3.4"
            Me.Caption = "Libro de inventarios y balances - Detalle del saldo de la cuenta 12 - Clientes"
        Case "3.8"
            TxtFec1.Enabled = False
            Me.Caption = "Libro de inventarios y balances - Detalle del saldo de la cuenta 20 - Mercaderías y la cuenta 21 - Productos terminados        "
            opt_mon(0).Value = True
            habilitar opt_mon, False
        Case "3.13"
            Me.Caption = "Libro de inventarios y balances - Detalle del saldo de la cuenta 42 - Proveedores"
        Case "--"
            Me.Caption = "Falta Título"
    End Select
    T_RPT_TITULO = Me.Caption

End Sub

Private Sub cmd_click(Index As Integer)
    Select Case Index
        Case 0 '--CONSULTAR
            CONSULTAR
        Case 1 '--EXPORTAR
            EXPORTAR
        Case 2 '--IMPRIMIR
            IMPRIMIR
        Case 3 '--SALIR
            BAND_INTERRUMPIR = True
            Unload Me
    End Select
End Sub


Private Sub CONSULTAR()
    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    '--
    Dim vStrSelect As String '--RECIBIR LA CONSULTA
    '--CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
   
    If Validar_Consulta() = False Then Exit Sub

    BAND_INTERRUMPIR = False
    '--CONFIGURAR LA PRESENTACION DE LA CONSULTA
    LimpiarGrid Me.Fg1, False, 1
    '--ENTRAR SOLO UNA VEZ
    vStrSelect = GENERAR_CONSULTA()
    Configurar_Grilla
        
    '--LIMPIAR ARRAY
    Limpiar_ARRAY_TOTAL True
    '----
    Me.MousePointer = vbHourglass
    DoEvents
    '------------------------------------------------
    If vStrSelect = "" Then GoTo SALIR
    PosicionarProgBar
    DoEvents
    '--CARGADO EL RST
    RST_Busq rst_select, vStrSelect, xCon
   '--------------------------------------
    
    CARGAR_DATOS_GRILLA rst_select
   '--------------------------------------
   '
SALIR:
    FraProgreso.Visible = False
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Set rst_select = Nothing
    SHOW_ERROR Me.Name, "Consultar"
    
End Sub

Private Function CARGAR_DATOS_GRILLA(RST_ORIGEN As ADODB.Recordset)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim BAND_ADD_REG As Boolean
    
    
    BAND_ADD_REG = True
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    PgBar.Min = 0
    PgBar.Max = RST_ORIGEN.RecordCount
    
    While Not RST_ORIGEN.EOF
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then Exit Function
        '---------------------------------------------------------
        Comparar_Grupo RST_ORIGEN, BAND_ADD_REG
        
        If RST_ORIGEN.Bookmark > Fg1.FixedRows Then ADD_REG Fg1
        '--ACUMULAR EN EL ARRAY_MES
        CARGAR_DATOS_ARRAY RST_ORIGEN
        '--CARGAR A LA GRILLA
        CARGAR_DATOS_GRILLA_ARRAY_TMP RST_ORIGEN, Fg1.Rows - 1
        '---------------------------------------------------------
        RST_ORIGEN.MoveNext
'        --PONER TOTALES AL FINAL DE LA GRILLA
        
        If RST_ORIGEN.EOF Then
            Select Case TIPO_FORMATO
                Case "3.3"
                    CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Totales:"
                Case "3.4", "3.13"
                    CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Saldo final total"
                Case "3.8"
                    CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Costo tot. gen."
                
            End Select
            
        Else
            PgBar.Value = CLng(RST_ORIGEN.Bookmark)
        End If
    Wend
    
    '------

End Function



Private Sub Comparar_Grupo(RST_ORIGEN As ADODB.Recordset, _
                            BAND_ADD_REG As Boolean, _
                            Optional Q_COL_COMPARAR As Integer = -1)
                            
    '--FUNCION QUE NOS PERMITE ARMAR LOS GRUPOS
    '--COMPARA CUANDO CAMBIAR DE GRUPO
    Dim RST_TEPM_1 As New ADODB.Recordset
    Dim N_GRUPO_ADD As String
    Dim Q_POS As Integer
    
    '---------------------------------------------------------
    If Q_COL_COMPARAR_GRUPO = -1 Then
        If RST_ORIGEN.Bookmark <= Fg1.FixedRows Then ADD_REG Fg1, Fila_Ninguno
        GoTo SALIR
    End If
    '---------------------------------------------------------
    If Q_COL_COMPARAR = -1 Then Q_COL_COMPARAR = Q_COL_COMPARAR_GRUPO
    
    If RST_ORIGEN.Bookmark = 1 Then
        '--SE CARGA EN GENERAR_CONSULTA() Q_COL_COMPARAR_GRUPO
        ADD_REG Fg1, Fila_Ninguno
        UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, " ", flexAlignLeftCenter
            
    Else
    
        Set RST_TEPM_1 = RST_ORIGEN.Clone
        RST_TEPM_1.Bookmark = RST_ORIGEN.Bookmark
        RST_TEPM_1.MovePrevious
        
        If RST_TEPM_1.Fields(Q_COL_COMPARAR) <> RST_ORIGEN.Fields(Q_COL_COMPARAR) Then
            CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:"
            
            ADD_REG Fg1, Fila_en_Blanco
            UNIR_CELDAS Fg1, Fg1.Rows - 1, IIf(Q_COL_FILA_OCULTA = -1, 1, Q_COL_FILA_OCULTA + 1), Fg1.Rows - 1, Fg1.Cols - 1, " ", flexAlignLeftCenter
            
            Limpiar_ARRAY_TOTAL
            
            
        End If
    End If

    
    
SALIR:
    Set RST_TEPM_1 = Nothing
End Sub

Private Sub CARGAR_DATOS_ARRAY(RST_ORIGEN As ADODB.Recordset)
    '--FUNCION QUE ACUMULARA EN EL ARRAY_TEMP
    Dim vStrCampo As String
    Dim Q_CAMPO As Integer
    Dim Q_POS As Integer
    Q_POS = 0
    '--ASIGNAR LOS DATOS AL RECORDSET TEMPORAL
    For Q_CAMPO = 0 To RST_ORIGEN.Fields.Count - 1
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then Exit Sub
        vStrCampo = RST_ORIGEN.Fields(Q_CAMPO).Name
        '--OBS: SE VA LLENAR EL ARRAY "TOTAL"
        
        Select Case LCase(vStrCampo)
            Case "deudorsol", "deudordol", "totsol", "totdol", "costotot"
                ARR_TMP(0, 0) = ARR_TMP(0, 0) + NulosN(Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO))
    
            Case "acreedorsol", "acreedordol", "saldosol", "saldodol"
                ARR_TMP(0, 1) = ARR_TMP(0, 1) + NulosN(Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO))
        
        End Select
        
    Next Q_CAMPO
    
End Sub

Private Function CARGAR_DATOS_GRILLA_ARRAY_TMP(RST_ORIGEN As ADODB.Recordset, _
                                         Q_ROW As Integer)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim Q_POS As Integer
    Dim Q_CAMPO As Integer
    Dim vStrCampo As String

    '-----------
    DoEvents
    
    For Q_CAMPO = 0 To RST_ORIGEN.Fields.Count - 1
        If BAND_INTERRUMPIR = True Then Exit Function
        vStrCampo = RST_ORIGEN.Fields(Q_CAMPO).Name
        
        Select Case LCase(vStrCampo)
            Case "saldebesol", "salhabersol", "saldebedol", "salhaberdol", "maydebesol", "mayhabersol", "maydebedol", "mayhaberdol", "deudorsol", "acreedorsol", "deudordol", "acreedordol", "debesol", "habersol", "debedol", "haberdol"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
            Case "totsol", "saldosol", "totdol", "saldodol"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
            
            Case "costotot"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                
            Case "preuni"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PU)
                
            Case "stckact"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_CANTIDAD)
                
            Case "fchdoc"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(RST_ORIGEN.Fields(vStrCampo) & "", FORMAT_DATE)
                
            Case Else
                '--AGREGAR LOS DEMAS DATOS
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
        End Select
    Next
End Function


Private Sub IMPRIMIR()

    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    
    If MsgBox("Desea conservar el formato de la consulta", vbQuestion + vbYesNo, "Imprimir...") = vbNo Then Configurar_Grilla False, True
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO + " ", "", T_RPT_PERIODO, False, True
    
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Imprimir"

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo error
    CentrarFrm Me
    
    ARR_FORMATOS(0) = "3.3"
    ARR_FORMATOS(1) = "3.4"
    ARR_FORMATOS(2) = "3.8"
    ARR_FORMATOS(3) = "3.13"
    
    TxtFec1.Valor = CDate("01/01/" + AnoTra)
    TxtFec2.Valor = CDate("31/12/" + AnoTra)
    GENERAR_CONSULTA
    Configurar_Grilla
    
    Exit Sub
error:
    SHOW_ERROR
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    BAND_INTERRUMPIR = True
    Erase ARR_TMP()
    Erase ARR_FORMATOS()
End Sub



'------
Private Function GENERAR_CONSULTA(Optional SOLO_CONFIG_GRID As Boolean = False) As String
    '--FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    '--
    Dim vStrSelect As String            '--CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    
    Dim vStrFiltro As String
    Dim vStrFiltro_1 As String      '--ESTE FILTRO SERVIRA PARA CONSULTAR EN EL SUB_SELECT
    
    Dim RstTmp As New ADODB.Recordset
   
    '--DE LA FECHA

    If CDate(TxtFec1.Valor) < CDate(TxtFec2.Valor) Then
        vStrFiltro = " ( con_diario.fchasi >=CDATE ('" + TxtFec1.Valor + "') AND con_diario.fchasi <= CDATE('" + TxtFec2.Valor + "') ) "
        T_RPT_PERIODO = " Del: " + CStr(TxtFec1.Valor) + " Al: " + CStr(TxtFec2.Valor)
    Else
        vStrFiltro = " con_diario.fchasi = CDATE('" + TxtFec1.Valor + "') "
         T_RPT_PERIODO = "Al: " + CStr(TxtFec2.Valor)
   End If

    '----------------------------------
    '----------------------------------
    'BUSCANDO LOS REGISTRO QUE TIENEN INCONSISTENCIAS
    '--DE LOS LIBROS
    '----------------------------------
    '----------------------------------
    
    '--GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim N_VALOR As String
    Dim N_CAMPOS As String
    Dim N_WHERE As String
    Dim N_FROM As String
    Dim N_GROUP_BY As String
    Dim N_ORDER_BY As String
    
    Dim monsunalt, mondesc As String
    Dim TIPO_CAMBIO As Double   '--INDICA EL ULTIMO TIPO DE CAMBIO EN FUNCION DE LA FECHA FINAL
    Select Case TIPO_FORMATO
    Case "3.3"
        Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 15:        Q_POSICION_TOTAL = 8:        Q_COL_COMPARAR_GRUPO = -1
        If opt_mon(0).Value = True Then '--TODO EN SOLES
            monsunalt = "1":            mondesc = "S/."
        Else
            monsunalt = "2":            mondesc = "$"
        End If
        
        
        
        N_CAMPOS = " con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion, IIf(InStr([con_planctas]![cuenta],'10-4')=0,'',[mae_bancos]![codsun]) AS bansunat, IIf(InStr([con_planctas]![cuenta],'10-4')=0,'',[mae_bancos]![descripcion]) AS banco, IIf(InStr([con_planctas]![cuenta],'10-4')=0,'',[con_bancocuenta]![numcue]) AS numcuenta, '" + monsunalt + "' AS monsunalt, '" + mondesc + "' AS mondesc, "
        If opt_mon(0).Value = True Then '--TODO EN SOLES
            N_CAMPOS = N_CAMPOS + " (SELECT Sum(IIf(con_diario1.impdebdol=0,con_diario1.impdebsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.impdebdol)))) AS debesol " _
                    + vbCr + " FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue " _
                    + vbCr + " WHERE (((con_planctas1.cuenta) Like '10%' And (con_planctas1.cuenta)<>'10') AND ((con_diario1.fchasi)<CDate('01/01/" + AnoTra + "') Or (con_diario1.fchasi) Is Null)) " _
                    + vbCr + " GROUP BY con_diario1.idcue " _
                    + vbCr + " HAVING (((con_diario1.idcue)=con_diario.idcue)) ) as saldebesol, " _
                    + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol=0,con_diario1.imphabsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.imphabdol)))) AS habersol " _
                    + vbCr + " FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue " _
                    + vbCr + " WHERE (((con_planctas1.cuenta) Like '10%' And (con_planctas1.cuenta)<>'10') AND ((con_diario1.fchasi)<CDate('01/01/" + AnoTra + "') Or (con_diario1.fchasi) Is Null)) " _
                    + vbCr + " GROUP BY con_diario1.idcue " _
                    + vbCr + " HAVING (((con_diario1.idcue)=con_diario.idcue))) as salhabersol," _
                    + vbCr + " Sum(IIf([con_diario].[impdebdol]=0,[con_diario].[impdebsol],IIf([con_tc].[impven] Is Null,0,([con_tc].[impven]*[con_diario].[impdebdol])))) AS debesol, " _
                    + vbCr + " Sum(IIf([con_diario].[imphabdol]=0,[con_diario].[imphabsol],IIf([con_tc].[impven] Is Null,0,([con_tc].[impven]*[con_diario].[imphabdol])))) AS habersol, " _
                    + vbCr + " iif ( debesol is null, 0 + iif(saldebesol is null,0,saldebesol),  debesol +  iif(saldebesol is null,0,saldebesol)    )  as maydebesol, " _
                    + vbCr + " iif ( habersol is null, 0 + iif(salhabersol is null,0,salhabersol),  habersol +  iif(salhabersol is null,0,salhabersol)    )  as mayhabersol, " _
                    + vbCr + " IIf(maydebesol>mayhabersol,(maydebesol-mayhabersol),0) AS deudorsol, " _
                    + vbCr + " IIf(mayhabersol>maydebesol,(mayhabersol-maydebesol),0) AS acreedorsol"
        Else '--TODO EN DOLARES
            N_CAMPOS = N_CAMPOS + " (SELECT Sum(IIf(con_diario1.impdebdol<>0,con_diario1.impdebdol,IIf(con_tc1.impven Is Null Or con_diario1.impdebsol=0,0,(con_diario1.impdebsol/con_tc1.impven)))) AS debedol " _
                    + vbCr + " FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue " _
                    + vbCr + " WHERE (((con_planctas1.cuenta) Like '10%' And (con_planctas1.cuenta)<>'10') AND ((con_diario1.fchasi)<CDate('01/01/" + AnoTra + "') Or (con_diario1.fchasi) Is Null)) " _
                    + vbCr + " GROUP BY con_diario1.idcue " _
                    + vbCr + " HAVING (((con_diario1.idcue)=con_diario.idcue))) as saldebedol, " _
                    + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol<>0,con_diario1.imphabdol,IIf(con_tc1.impven Is Null Or con_diario1.imphabsol=0,0,(con_diario1.imphabsol/con_tc1.impven)))) AS haberdol " _
                    + vbCr + " FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue " _
                    + vbCr + " WHERE (((con_planctas1.cuenta) Like '10%' And (con_planctas1.cuenta)<>'10') AND ((con_diario1.fchasi)<CDate('01/01/" + AnoTra + "') Or (con_diario1.fchasi) Is Null)) " _
                    + vbCr + " GROUP BY con_diario1.idcue " _
                    + vbCr + " HAVING (((con_diario1.idcue)=con_diario.idcue))) as salhaberdol, " _
                    + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_diario].[impdebsol]=0,0,([con_diario].[impdebsol]/[con_tc].[impven])))) AS debedol, Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol],IIf([con_tc].[impven] Is Null Or [con_diario].[imphabsol]=0,0,([con_diario].[imphabsol]/[con_tc].[impven])))) AS haberdol, " _
                    + vbCr + " iif ( debedol is null, 0 + iif(saldebedol is null,0,saldebedol),  debedol +  iif(saldebedol is null,0,saldebedol)    )  as maydebedol, " _
                    + vbCr + " iif ( haberdol is null, 0 + iif(salhaberdol is null,0,salhaberdol),  haberdol +  iif(salhaberdol is null,0,salhaberdol)    )  as mayhaberdol, " _
                    + vbCr + " IIf(maydebedol>mayhaberdol,(maydebedol-mayhaberdol),0) AS deudordol, " _
                    + vbCr + " IIf(mayhaberdol > maydebedol, (mayhaberdol - maydebedol), 0) As acreedordol "
        End If
        N_FROM = " con_planctas RIGHT JOIN (mae_bancos RIGHT JOIN (((con_diario LEFT JOIN con_cajabanco ON con_diario.idmov = con_cajabanco.id) LEFT JOIN con_bancocuenta ON con_cajabanco.idcueban = con_bancocuenta.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_bancos.id = con_bancocuenta.idban) ON con_planctas.id = con_diario.idcue "
        N_GROUP_BY = " con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion, IIf(InStr([con_planctas]![cuenta],'10-4')=0,'',[mae_bancos]![codsun]), IIf(InStr([con_planctas]![cuenta],'10-4')=0,'',[mae_bancos]![descripcion]), IIf(InStr([con_planctas]![cuenta],'10-4')=0,'',[con_bancocuenta]![numcue]), '', '' "
        N_ORDER_BY = " con_planctas.cuenta, con_planctas.descripcion; "
    
        N_WHERE = " (((con_planctas.cuenta) Like '10%' And (con_planctas.cuenta)<>'10') AND ((con_diario.fchasi)>=CDate('01/01/" + AnoTra + "') And (con_diario.fchasi)<=CDate('31/12/" + AnoTra + "'))) "
        N_WHERE = N_WHERE + " AND " + vStrFiltro
    Case "3.4", "3.13"
        Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 10:        Q_POSICION_TOTAL = 8:        Q_COL_COMPARAR_GRUPO = -1
        
        N_CAMPOS = " vta_ventas.id, '01' AS codsun, mae_cliente.numruc, mae_cliente.nombre, vta_ventas.numreg, vta_ventas.fchdoc, mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numerodoc,mae_moneda.simbolo, "
        If opt_mon(0).Value = True Then
            N_CAMPOS = N_CAMPOS + " IIf([vta_ventas]![imptotdoc] Is Null,0,IIf([vta_ventas]![idmon]=1,[vta_ventas]![imptotdoc],[vta_ventas]![imptotdoc]*[con_tc].[impven])) AS totsol, IIf([vta_ventas]![impsal] Is Null,0,IIf([vta_ventas]![idmon]=1,[vta_ventas]![impsal],[vta_ventas]![impsal]*[con_tc].[impven])) AS saldosol "
        Else
            N_CAMPOS = N_CAMPOS + " IIf([vta_ventas]![idmon]=2,[vta_ventas]![imptotdoc],IIf([vta_ventas]![imptotdoc] Is Null Or [con_tc].[impven] Is Null,0,[vta_ventas]![imptotdoc]/[con_tc].[impven])) AS totdol, IIf([vta_ventas]![idmon]=2,[vta_ventas]![impsal],IIf([vta_ventas]![impsal] Is Null Or [con_tc].[impven] Is Null,0,[vta_ventas]![impsal]/[con_tc].[impven])) AS saldodol "
        End If
        N_FROM = " (((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id "
        
        vStrFiltro = Replace(vStrFiltro, "con_diario.fchasi", "vta_ventas.fchdoc")
        
        N_WHERE = " (((vta_ventas.impsal)<>0) AND ((vta_ventas.fchdoc)>=CDate('01/01/" + AnoTra + "') And (vta_ventas.fchdoc)<=CDate('31/12/" + AnoTra + "'))) "
        N_WHERE = N_WHERE + " AND " + vStrFiltro
        
        N_GROUP_BY = ""
        N_ORDER_BY = " vta_ventas.fchdoc, mae_cliente.nombre; "
        
        If TIPO_FORMATO = "3.13" Then
            N_CAMPOS = Replace(N_CAMPOS, "vta_ventas", "com_compras")
            N_CAMPOS = Replace(N_CAMPOS, "mae_cliente", "mae_prov")
            N_CAMPOS = Replace(N_CAMPOS, ".idcli", ".idpro")
            N_CAMPOS = Replace(N_CAMPOS, "imptotdoc", "imptot")
            
            N_FROM = Replace(N_FROM, "vta_ventas", "com_compras")
            N_FROM = Replace(N_FROM, "mae_cliente", "mae_prov")
            N_FROM = Replace(N_FROM, ".idcli", ".idpro")
            
            N_WHERE = Replace(N_WHERE, "vta_ventas", "com_compras")
            N_WHERE = Replace(N_WHERE, "mae_cliente", "mae_prov")
            N_WHERE = Replace(N_WHERE, ".idcli", ".idpro")
            N_WHERE = Replace(N_WHERE, "imptotdoc", "imptot")
            
            N_ORDER_BY = Replace(N_ORDER_BY, "vta_ventas", "com_compras")
            N_ORDER_BY = Replace(N_ORDER_BY, "mae_cliente", "mae_prov")
            N_ORDER_BY = Replace(N_ORDER_BY, ".idcli", ".idpro")
            N_ORDER_BY = Replace(N_ORDER_BY, "imptotdoc", "imptot")
            
        End If
        
        
    Case "3.8"
        '--DEL ULTIMO TIPO DE CAMBIO
        TIPO_CAMBIO = 0
        RST_Busq RstTmp, "SELECT TOP 1 con_tc.impven FROM con_tc WHERE (((con_tc.fecha)<=CDate('" + Format(TxtFec2.Valor, "dd/mm/yy") + "')) AND ((con_tc.idmon)=2)) ORDER BY con_tc.fecha DESC;", xCon
        If RstTmp.State = 1 Then
            TIPO_CAMBIO = NulosN(RstTmp.Fields(0))
        End If
        Set RstTmp = Nothing
        '-----------------------------------
        Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 10:        Q_POSICION_TOTAL = 10:        Q_COL_COMPARAR_GRUPO = -1
        
        N_CAMPOS = "  alm_inventario.id AS prodid, alm_inventario.codpro AS prodcod, mae_tipoproducto.codsun AS tipprodcodsun, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.descripcion AS proddec, mae_unidades.codsun, mae_unidades.abrev, alm_inventario.stckact, mae_moneda.simbolo, alm_inventario.preuni, " _
                + vbCr + "IIf([alm_inventario]![preuni] Is Null Or [alm_inventario]![preuni]=0 Or [alm_inventario]![stckact] Is Null Or [alm_inventario]![stckact]=0,0,IIf([alm_inventario]![idmon]=1,[alm_inventario]![stckact]*[alm_inventario]![preuni]," + CStr(TIPO_CAMBIO) + " * [alm_inventario]![preuni]*[alm_inventario]![stckact])) AS CostoTot "
        N_FROM = " mae_moneda RIGHT JOIN (mae_unidades RIGHT JOIN (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) ON mae_moneda.id = alm_inventario.idmon "
        N_GROUP_BY = ""
        N_WHERE = "  (((alm_inventario.activo)=-1) AND ((alm_inventario.contable)=-1)) "
        N_ORDER_BY = "  mae_tipoproducto.codsun, alm_inventario.descripcion; "
        
        
        
    Case "--"
        
    
    End Select
    
    '--DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    
    '------------------------------------------

    '--GENERANDO LA CONSULTA
    vStrSelect = "SELECT " + N_CAMPOS + _
        vbCr + " FROM " + N_FROM + _
        vbCr + " WHERE " + N_WHERE + _
        vbCr + IIf(N_GROUP_BY <> "", " GROUP BY ", "") + N_GROUP_BY + _
        vbCr + " ORDER BY " + N_ORDER_BY

    '------------------------------------------------------------------------------------
    GENERAR_CONSULTA = vStrSelect
    
    Set RstTmp = Nothing
End Function



Private Sub Limpiar_ARRAY_TOTAL(Optional F_LIMPIA_TOT_GRL As Boolean = False)
    
    ARR_TMP(0, 0) = 0:      ARR_TMP(0, 1) = 0
    
    If F_LIMPIA_TOT_GRL = True Then
        ARR_TMP(1, 0) = 0:      ARR_TMP(1, 1) = 0
    End If

End Sub
'''
Private Sub CARGAR_DATOS_GRILLA_ADD_TOTALES(BAND_ADD_TOTAL As Boolean, _
                                            Nombre_total As String, _
                                            Optional Band_Total_gral As Boolean = False)
                
    Dim Q_MES As Integer
    '--AGREGA LOS TOTALES POR CADA GRUPO Y EL TOTAL GENERAL
    '--ACUMULA LOS TOTALES EN EL TOTAL GENERAL
    Dim X_ROW As Long
    'On Error Resume Next
    X_ROW = Fg1.Rows
    If BAND_ADD_TOTAL = True Then
        '--AGREAGNDO NUEVA FILA
        ADD_REG Fg1, IIf(Band_Total_gral = False, Fila_Total, Fila_Total_grl)

        'PONIENDO LOS NOMBRES DE LOS TOTALES  Q_POSICION_TOTAL SE OBTIENE DE GENERAR_CONSULTA()
        Fg1.TextMatrix(X_ROW, Q_POSICION_TOTAL) = Nombre_total
        FORMATO_CELDA Fg1, X_ROW, Q_POSICION_TOTAL
    End If
    
    
    '--ACUMULANDO LOS TOTALES GRLES
    If Band_Total_gral = False Then
        ARR_TMP(1, 0) = NulosN(ARR_TMP(1, 0)) + NulosN(ARR_TMP(0, 0)) '--DEBE
        ARR_TMP(1, 1) = NulosN(ARR_TMP(1, 1)) + NulosN(ARR_TMP(0, 1)) '--HABER
    End If



    Select Case TIPO_FORMATO
        Case "3.3"
            '--HABER (DEUDOR)
            Fg1.TextMatrix(X_ROW, Fg1.Cols - 2) = Format(IIf(Band_Total_gral = False, ARR_TMP(0, 0), ARR_TMP(1, 0)), FORMAT_MONTO)
            FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 2
            '--DEBE (ACREEDOR)
            Fg1.TextMatrix(X_ROW, Fg1.Cols - 1) = Format(IIf(Band_Total_gral = False, ARR_TMP(0, 1), ARR_TMP(1, 1)), FORMAT_MONTO)
            FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 1
        Case "3.4", "3.13"
            '--IMPORTE
            Fg1.TextMatrix(X_ROW, Fg1.Cols - 2) = Format(IIf(Band_Total_gral = False, ARR_TMP(0, 0), ARR_TMP(1, 0)), FORMAT_MONTO)
            FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 2
            '--SALDO
            Fg1.TextMatrix(X_ROW, Fg1.Cols - 1) = Format(IIf(Band_Total_gral = False, ARR_TMP(0, 1), ARR_TMP(1, 1)), FORMAT_MONTO)
            FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 1
        Case "3.8" '--COSTO TOTAL
            Fg1.TextMatrix(X_ROW, Fg1.Cols - 1) = Format(IIf(Band_Total_gral = False, ARR_TMP(0, 0), ARR_TMP(1, 0)), FORMAT_MONTO)
            FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 1
    Case "--"
            Me.Caption = "Falta Título"
    End Select
    
    Err.Clear
End Sub

Private Sub Configurar_Grilla(Optional F_CONSERVAR_FORMATO As Boolean = False, Optional F_IMPRIMIR As Boolean = False)
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
    Dim M_ANCHO_COL As Integer '--DEPENDERA DEL TIPO DE CONSULTA
                                   
    Dim k, j As Integer
    Dim T_CONSULTA As Integer
    
    If F_CONSERVAR_FORMATO = True Then Fg1.Clear
    
    Fg1.FrozenCols = 0
    
    M_ANCHO_COL = 0

    With Fg1
        '-----
        Fg1.Cols = Q_COL_FILA_OCULTA + Q_COL_FILA
                 
        .ColWidth(0) = 200
        '--DATOS DE FILA
        Select Case TIPO_FORMATO
            Case "3.3"
                If F_IMPRIMIR = False Then .Rows = 2
                .FixedRows = 2
                .RowHeight(0) = 350
                UNIR_CELDAS Fg1, 0, 2, 0, 3, "Cuenta Contable Divisionaria", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 4, 0, 8, "Referencia de la Cuenta", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 0, 9, 0, 16, "Saldo Contable Final", flexAlignCenterCenter
                .RowHeight(1) = 500
                .TextMatrix(1, 2) = "Código":                           .ColWidth(2) = 1100:    .ColAlignment(2) = flexAlignLeftCenter:         .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 3) = "Denominación":                     .ColWidth(3) = 3000:    .ColAlignment(3) = flexAlignLeftCenter:         .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 4) = "Código" + vbCr + "Sunat":          .ColWidth(4) = 600:     .ColAlignment(4) = flexAlignCenterCenter:       .Row = 1: .Col = 4: .CellAlignment = flexAlignCenterBottom
                .TextMatrix(1, 5) = "Entidad Financiera":               .ColWidth(5) = 1500:    .ColAlignment(5) = flexAlignLeftCenter:         .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 6) = "Número de" + vbCr + "la Cuenta":   .ColWidth(6) = 1600:    .ColAlignment(6) = flexAlignLeftCenter:         .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(1, 7) = "Código" + vbCr + "Sunat":          .ColWidth(7) = 700:     .ColAlignment(7) = flexAlignCenterBottom:       .Row = 1: .Col = 7: .CellAlignment = flexAlignCenterBottom
                .TextMatrix(1, 8) = "Tipo de" + vbCr + "Moneda":        .ColWidth(8) = 700:     .ColAlignment(8) = flexAlignCenterBottom:       .Row = 1: .Col = 8: .CellAlignment = flexAlignCenterBottom
                '----------------------
                .TextMatrix(1, 9) = "Saldo Debe":                       .ColWidth(9) = 0:       .ColAlignment(9) = flexAlignRightCenter:        .Row = 1: .Col = 9: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 10) = "Saldo Haber":                     .ColWidth(10) = 0:      .ColAlignment(10) = flexAlignRightCenter:       .Row = 1: .Col = 10: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 11) = "Debe":                            .ColWidth(11) = 0:      .ColAlignment(11) = flexAlignRightCenter:       .Row = 1: .Col = 11: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 12) = "Haber":                           .ColWidth(12) = 0:      .ColAlignment(12) = flexAlignRightCenter:       .Row = 1: .Col = 12: .CellAlignment = flexAlignRightCenter
                
                .TextMatrix(1, 13) = "Mayor Debe":                      .ColWidth(13) = 0:   .ColAlignment(13) = flexAlignRightCenter:          .Row = 1: .Col = 13: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 14) = "Mayor Haber":                     .ColWidth(14) = 0:   .ColAlignment(14) = flexAlignRightCenter:          .Row = 1: .Col = 14: .CellAlignment = flexAlignRightCenter
                
                .TextMatrix(1, 15) = "Deudor":                          .ColWidth(15) = 1050:   .ColAlignment(15) = flexAlignRightCenter:       .Row = 1: .Col = 15: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 16) = "Acreedor":                        .ColWidth(16) = 1050:   .ColAlignment(16) = flexAlignRightCenter:       .Row = 1: .Col = 16: .CellAlignment = flexAlignRightCenter
                

                
                .FrozenCols = 3
                

            Case "3.4", "3.13"
                If F_IMPRIMIR = False Then .Rows = 3
                .FixedRows = 3
                .RowHeight(0) = 300:        .RowHeight(1) = 300:       .RowHeight(2) = 300
                If TIPO_FORMATO = "3.4" Then
                    UNIR_CELDAS Fg1, 0, 2, 0, 4, "Información del Cliente", flexAlignCenterCenter
                Else
                    UNIR_CELDAS Fg1, 0, 2, 0, 4, "Información del Proveedor", flexAlignCenterCenter
                End If
                UNIR_CELDAS Fg1, 1, 2, 1, 3, "Documento de Identidad", flexAlignCenterCenter
                UNIR_CELDAS Fg1, 1, 4, 2, 4, "Apellidos y Nombres," + vbCr + "Denominación o Razón Social", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 0, 5, 0, 11, "Datos del Documento", flexAlignCenterCenter

                .TextMatrix(2, 2) = "Tipo":                       .ColWidth(2) = 500:     .ColAlignment(2) = flexAlignLeftCenter:   .Row = 2: .Col = 2: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(2, 3) = "Número":                     .ColWidth(3) = 1400:    .ColAlignment(3) = flexAlignLeftCenter:   .Row = 2: .Col = 3: .CellAlignment = flexAlignLeftCenter
                
                UNIR_CELDAS Fg1, 1, 5, 2, 5, "N°.Reg.", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 6, 2, 6, "Fec.Doc.", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 7, 2, 7, "T.D.", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 8, 2, 8, "Num. Documento", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 9, 2, 9, "M", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 10, 2, 10, "Imp.", flexAlignCenterCenter, False
                UNIR_CELDAS Fg1, 1, 11, 2, 11, "Saldo", flexAlignCenterCenter, False
                
                .ColWidth(4) = 3500:    .ColAlignment(4) = flexAlignLeftCenter
                
                .ColWidth(5) = 800:     .ColAlignment(5) = flexAlignCenterCenter
                .ColWidth(6) = 850:     .ColAlignment(6) = flexAlignCenterCenter
                .ColWidth(7) = 450:     .ColAlignment(7) = flexAlignCenterCenter
                .ColWidth(8) = 1600:    .ColAlignment(8) = flexAlignCenterCenter
                .ColWidth(9) = 0:     .ColAlignment(9) = flexAlignRightCenter
                .ColWidth(10) = 1000:   .ColAlignment(10) = flexAlignRightCenter
                .ColWidth(11) = 1000:   .ColAlignment(11) = flexAlignRightCenter
                
                .FrozenCols = 4
                .MergeCells = flexMergeFixedOnly

            Case "3.8"
                If F_IMPRIMIR = False Then .Rows = 1
                .FixedRows = 1
                .RowHeight(0) = 750

                .TextMatrix(0, 2) = "Código de la" + vbCr + "Existencia":               .ColWidth(2) = 1900:    .ColAlignment(2) = flexAlignLeftCenter:           .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(0, 3) = "Tipo de" + vbCr + "Existencia":                    .ColWidth(3) = 800:     .ColAlignment(3) = flexAlignCenterCenter:         .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(0, 4) = "Descripción de Tipo " + vbCr + "de Existencia":    .ColWidth(4) = 0:       .ColAlignment(4) = flexAlignLeftCenter:           .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(0, 5) = "Descripción":                                      .ColWidth(5) = 3585:    .ColAlignment(5) = flexAlignLeftCenter:           .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(0, 6) = "Código de" + vbCr + "Unidad de" + vbCr + "Medida": .ColWidth(6) = 850:     .ColAlignment(6) = flexAlignCenterCenter:         .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(0, 7) = "Descripción" + vbCr + "de la Unidad" + vbCr + "de Medida": .ColWidth(7) = 950: .ColAlignment(7) = flexAlignLeftCenter:       .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftCenter
                .TextMatrix(0, 8) = "Cantidad":                                         .ColWidth(8) = 800:     .ColAlignment(8) = flexAlignRightCenter:          .Row = 0: .Col = 8: .CellAlignment = flexAlignRightCenter
                .TextMatrix(0, 9) = "M":                                                .ColWidth(9) = 0:       .ColAlignment(9) = flexAlignCenterCenter:         .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(0, 10) = "Costo" + vbCr + "Unitario":                       .ColWidth(10) = 1150:   .ColAlignment(10) = flexAlignRightCenter:         .Row = 0: .Col = 10: .CellAlignment = flexAlignRightCenter
                .TextMatrix(0, 11) = "Costo Total":                                     .ColWidth(11) = 1000:   .ColAlignment(11) = flexAlignRightCenter:         .Row = 0: .Col = 11: .CellAlignment = flexAlignRightCenter

                .FrozenCols = 5
            
            Case "--"
            
            
        End Select
        
        M_ANCHO_COL = 0

        '--DE LOS ID'S
        For k = 1 To Q_COL_FILA_OCULTA
            .TextMatrix(0, k) = "ID" + CStr(k):         .ColWidth(k) = 500
        Next
        If Q_COL_FILA_OCULTA <> -1 Then OCULTAR_COL Fg1, 1, Q_COL_FILA_OCULTA
   
    End With
    DoEvents
End Sub

Sub PosicionarProgBar()
    '--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
    '    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    Me.PgBar.Value = 1
    FraProgreso.Visible = True
End Sub

'----

Private Function Validar_Consulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    '--DE LA FECHA ES NULL
    
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Function
    End If
    
    If TxtFec1.Valor = "" Or TxtFec2.Valor = "" Then
        MsgBox "Ingrese una fecha", vbExclamation, xTitulo
        If TxtFec1.Valor = "" Then TxtFec1.SetFocus Else TxtFec2.SetFocus
        Exit Function
    End If
    If CDate(TxtFec1.Valor) > CDate(TxtFec2.Valor) Then
        MsgBox "La fecha inicial es superior al Final", vbExclamation, xTitulo
        If TxtFec1.Enabled = True Then TxtFec1.SetFocus
        Exit Function
    End If
    
    If (Year(TxtFec1.Valor) <> Year(TxtFec2.Valor)) Then
        MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
        TxtFec1.SetFocus
        Exit Function
    ElseIf Year(TxtFec1.Valor) <> CStr(AnoTra) Then
        MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
        TxtFec1.SetFocus
        Exit Function
    End If
    
    '''''VALIDAR SI EXISTE EL FORMATO
    Dim k As Integer
    Dim ES_FORMATO As Boolean
    For k = 0 To UBound(ARR_FORMATOS)
        If ARR_FORMATOS(k) = TIPO_FORMATO Then
            ES_FORMATO = True
            Exit For
        End If
    Next k
    
    If ES_FORMATO = False Then
        MsgBox "No existe el formato que intenta Consultar" + vbCr + "Inténtelo en Otro Momento", vbExclamation, xTitulo
        Exit Function
    End If
    
    '''''
    
    Validar_Consulta = True
End Function


Private Sub EXPORTAR()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios

    If MsgBox("Desea conservar el formato de la consulta", vbQuestion + vbYesNo, "Exportar...") = vbNo Then Configurar_Grilla False
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO + " ", "", T_RPT_PERIODO, "Formato: " + TIPO_FORMATO
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub
