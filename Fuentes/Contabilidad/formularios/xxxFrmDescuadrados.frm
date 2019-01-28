VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDescuadrados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Diario - Asientos Descuadrados"
   ClientHeight    =   7260
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   10410
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10410
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   2115
      TabIndex        =   0
      Top             =   2850
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   1
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
         TabIndex        =   5
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
         TabIndex        =   3
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
         TabIndex        =   2
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
      Height          =   6840
      Left            =   60
      TabIndex        =   4
      Top             =   375
      Width           =   10305
      _cx             =   18177
      _cy             =   12065
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmDescuadrados.frx":0000
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDescuadrados.frx":0211
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDescuadrados.frx":0665
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDescuadrados.frx":07D1
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDescuadrados.frx":0D19
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmDescuadrados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---
'--  FNUCION QUE RECIBE LOS PARAMETROS RECIBE_LINK_FRM ( FECHA_INICIO::DATE, FECHA_FIN::DATE , ID_LIBRO (OPCIONAL)::INTEGER)
'--

Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
'--DE LA IMPRESION
Dim T_RPT_PERIODO As String '--PERIODO DEL REPORTE
Dim T_RPT_PERIODO1 As String
Dim T_RPT_TITULO As String  '--TITULO DE REPORTE
'------------

Dim ARR_TMP() As String '--0::PROGRAMADO=>> 0::TOTAL,1::TOTAL GEN
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
'------------
Dim TOTAL_REGISTROS As Long '--INDICA LA CANTIDAD DE REGISTROS DESCUADRADOS
Dim ID_LIBRO As Integer
Dim D_FECHA_INI As Date
Dim D_FECHA_FIN As Date

Dim fTipoConsulta As Boolean
Dim mMesIni As Integer
Dim mMesFin As Integer

Public Sub RECIBE_LINK_FRM(FCH_INI As Date, FCH_FIN As Date, IdMesIni As Integer, IdMesFin As Integer, Optional xFecha As Boolean = True, Optional IDLIBRO As Integer = 0)
    '--FrmDescuadrados.RECIBE_LINK_FRM CDate("01/01/07"), CDate("01/10/07"), 0

    
    D_FECHA_INI = CDate(FCH_INI)
    D_FECHA_FIN = CDate(FCH_FIN)
    mMesIni = IdMesIni
    mMesFin = IdMesFin
    fTipoConsulta = xFecha
    
    If mMesIni = mMesFin Then
        T_RPT_PERIODO = "Periodo: " + Busca_Codigo(mMesIni, "id", "descripcion", "con_meses", "N", xCon)
    Else
        T_RPT_PERIODO = "Periodo: De " + Busca_Codigo(mMesIni, "id", "descripcion", "con_meses", "N", xCon) & " A " + Busca_Codigo(mMesFin, "id", "descripcion", "con_meses", "N", xCon)
    End If
    
    If fTipoConsulta = True Then
        If CDate(FCH_INI) < CDate(FCH_INI) Then
            MsgBox "La fecha de inicio es superior al final", vbExclamation, xTitulo
            Unload Me
            Exit Sub
        End If
    End If
    ID_LIBRO = IDLIBRO
    CONSULTAR
    Me.MousePointer = vbDefault

End Sub


Private Sub CONSULTAR()
'    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    '--
    Dim vStrSelect As String '--RECIBIR LA CONSULTA
   
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
    If TOTAL_REGISTROS = 0 Then GoTo SALIR
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
        
        If RST_ORIGEN.Bookmark <> 1 Then ADD_REG Fg1
        '--ACUMULAR EN EL ARRAY_MES
        CARGAR_DATOS_ARRAY RST_ORIGEN
        '--CARGAR A LA GRILLA
        CARGAR_DATOS_GRILLA_ARRAY_TMP RST_ORIGEN, Fg1.Rows - 1
            
        '---------------------------------------------------------
        '---------------------------------------------------------
        RST_ORIGEN.MoveNext
'        --PONER TOTALES AL FINAL DE LA GRILLA
        
        If RST_ORIGEN.EOF Then
            CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:"
            
            ADD_REG Fg1, Fila_Total
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, "Total Registros Encontrados: " + CStr(TOTAL_REGISTROS), flexAlignLeftBottom
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 1
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
        If RST_ORIGEN.Bookmark = 1 Then ADD_REG Fg1, Fila_Ninguno
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
        
        If LCase(vStrCampo) = "debe" Then
            ARR_TMP(0, 0) = ARR_TMP(0, 0) + NulosN(Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO))

        ElseIf LCase(vStrCampo) = "haber" Then
            ARR_TMP(0, 1) = ARR_TMP(0, 1) + NulosN(Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO))

        End If
        
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
            Case "debe", "haber"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
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
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO, T_RPT_PERIODO1, T_RPT_PERIODO, False, True
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
    
    Exit Sub
error:
    SHOW_ERROR
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    BAND_INTERRUMPIR = True
    Erase ARR_TMP
End Sub



'------
Private Function GENERAR_CONSULTA(Optional SOLO_CONFIG_GRID As Boolean = False) As String
    '--FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    '--
    Dim vStrSelect As String            '--CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    
    Dim vStrFiltro As String
    Dim vStrFiltro_1 As String      '--ESTE FILTRO SERVIRA PARA CONSULTAR EN EL SUB_SELECT

    Dim k As Integer
    
    Dim SQL_LIBRO As String
    Dim SQL_INSUMO As String
    Dim T_CONSULTA As Integer '--DEL TIPO DE CONSULTA, SE FORMARA EL ENCABEZADO DEL GRID
    
        '--DE LA FECHA
    If fTipoConsulta = True Then
        If CDate(D_FECHA_INI) < CDate(D_FECHA_FIN) Then
            vStrFiltro = " ( con_diario.fchasi >=CDATE ('" & D_FECHA_INI & "') AND con_diario.fchasi <= CDATE('" & D_FECHA_FIN & "') ) "
            T_RPT_PERIODO = " Del: " + CStr(D_FECHA_INI) + " Al: " + CStr(D_FECHA_FIN)
        Else
            vStrFiltro = " con_diario.fchasi = CDATE('" & D_FECHA_INI & "') "
            T_RPT_PERIODO = "Al: " + CStr(D_FECHA_INI)
        End If
    Else
            vStrFiltro = " ( con_diario.idmes >= " & mMesIni & " and con_diario.idmes <= " & mMesFin & " ) "
    End If
   
   
   '--DE LOS SALDOS IDMES=0 (OPCIONAL)
   'vStrFiltro = "( " + vStrFiltro + " OR con_diario.fchasi IS NULL ) "
   
    '----------------------------------
    '----------------------------------
    'BUSCANDO LOS REGISTRO QUE TIENEN INCONSISTENCIAS
    '--DE LOS LIBROS
    TOTAL_REGISTROS = 0
    If ID_LIBRO > 0 Then SQL_LIBRO = " AND con_diario.idlib = " + CStr(ID_LIBRO)
    DoEvents
   If SOLO_CONFIG_GRID = False Then
        
        vStrFiltro = vStrFiltro + SQL_LIBRO
        
        vStrFiltro_1 = Replace(vStrFiltro, "con_diario.", "con_diario1.")
        vStrFiltro_1 = Replace(vStrFiltro_1, "con_tc.", "con_tc1.")
       
        vStrFiltro_1 = " SELECT con_diario1.idlib & con_diario1.idmov  as ID " _
            + vbCr + " FROM con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha " _
            + vbCr + " WHERE " + vStrFiltro_1 _
            + vbCr + " GROUP BY con_diario1.idlib & con_diario1.idmov " _
            + vbCr + " HAVING  (((Format(Sum(format(IIf([con_diario1].[impdebdol]=0,[con_diario1].[impdebsol],IIf([con_tc1].[impven] Is Null,0,([con_tc1].[impven]*[con_diario1].[impdebdol]))),'00000.00')),'00000.00')) <>  Format(Sum(format(IIf([con_diario1].[imphabdol]=0,[con_diario1].[imphabsol],IIf([con_tc1].[impven] Is Null,0,([con_tc1].[impven]*[con_diario1].[imphabdol]))),'00000.00')),'00000.00'))) "
    
        Dim xRs As New ADODB.Recordset
        RST_Busq xRs, vStrFiltro_1, xCon
        Dim SQL_ID As String
        
        If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then xRs.MoveFirst
        Do While Not xRs.EOF
            SQL_ID = SQL_ID + "'" + CStr(xRs.Fields("ID")) + "',"
            xRs.MoveNext
        Loop
        If SQL_ID <> "" Then SQL_ID = " AND con_diario.idlib & con_diario.idmov  IN (" + Left(SQL_ID, Len(SQL_ID) - 1) + ") "
        TOTAL_REGISTROS = xRs.RecordCount
        Set xRs = Nothing
    End If
    '----------------------------------
    '----------------------------------
    
    '--GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim N_VALOR As String
    Dim N_CAMPOS As String
    Dim N_WHERE As String
    Dim N_FROM As String
    Dim N_GROUP_BY As String
    Dim N_ORDER_BY As String
    
    N_WHERE = vStrFiltro + SQL_ID
   
    Q_COL_FILA_OCULTA = 3:         Q_COL_FILA = 7:        Q_POSICION_TOTAL = 8:        Q_COL_COMPARAR_GRUPO = 0
    
    T_RPT_TITULO = "REPORTE DE INCONSISTENCIAS DEL DIARIO"
    T_RPT_PERIODO1 = "Asientos Descuadrados"
    
    
    N_CAMPOS = " con_diario.idlib & con_diario.idmov AS IDCONSULTA, con_diario.idlib, con_diario.idmov, Format([con_diario]![idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',Format([mae_libros].[codsun],'00')) & Trim([con_diario]![numasi]) AS numreg, IIf([con_diario].[idlib]<>3,[mae_libros].[descripcion],[mae_librossub].[descripcion]) AS libdesc,  con_planctas.cuenta, con_planctas.descripcion AS ctadesc, con_tc.impven," _
            + vbCr + " IIf(con_diario.impdebdol=0,con_diario.impdebsol,IIf(con_tc.impven Is Null,0,(con_tc.impven*con_diario.impdebdol))) AS debe, IIf(con_diario.imphabdol=0,con_diario.imphabsol,IIf(con_tc.impven Is Null,0,(con_tc.impven*con_diario.imphabdol))) AS haber "
            
    N_FROM = " (mae_libros RIGHT JOIN (con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue) ON mae_libros.id = con_diario.idlib) LEFT JOIN (con_proviciones LEFT JOIN mae_librossub ON (con_proviciones.idlib = mae_librossub.idlib) AND (con_proviciones.idsublib = mae_librossub.id)) ON con_diario.idmov = con_proviciones.id "
    
    N_ORDER_BY = "  con_diario.idmes,con_diario.idlib, con_diario.idmov, con_diario.fchasi,con_planctas.cuenta; "
    
    '--DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    '------------------------------------------
    '--GENERANDO LA CONSULTA
    vStrSelect = "SELECT " + N_CAMPOS + _
    vbCr + " FROM " + N_FROM + _
    vbCr + " WHERE " + N_WHERE + _
    vbCr + " ORDER BY " + N_ORDER_BY

    '------------------------------------------------------------------------------------
    GENERAR_CONSULTA = vStrSelect

End Function

Private Sub Limpiar_ARRAY_TOTAL(Optional F_LIMPIA_TOT_GRL As Boolean = False)
    Erase ARR_TMP()
    ReDim ARR_TMP(0, 1)
    ARR_TMP(0, 0) = 0
    ARR_TMP(0, 1) = 0
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
        FORMATO_CELDA Fg1, X_ROW, Q_POSICION_TOTAL, vbRed
    End If

    '--HABER
    Fg1.TextMatrix(X_ROW, Fg1.Cols - 2) = Format(ARR_TMP(0, 0), FORMAT_MONTO)
    FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 2, vbRed
    '--DEBE
    Fg1.TextMatrix(X_ROW, Fg1.Cols - 1) = Format(ARR_TMP(0, 1), FORMAT_MONTO)
    FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 1, vbRed
    
    Err.Clear
End Sub

Private Sub Configurar_Grilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
    Dim M_ANCHO_COL As Integer '--DEPENDERA DEL TIPO DE CONSULTA
                                   
    Dim k, j As Integer
    Dim T_CONSULTA As Integer
    
    If F_CONSERVAR_FORMATO = True Then Fg1.Clear
    
    M_ANCHO_COL = 0

    With Fg1
        .FrozenCols = 0
        '-----
        .Cols = Q_COL_FILA_OCULTA + Q_COL_FILA
                 
        .ColWidth(0) = 200
        '--DATOS DE FILA
        

        
        .TextMatrix(0, 4) = "Num.Reg.":             .ColWidth(4) = 900:   .ColAlignment(4) = flexAlignLeftCenter:   .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 5) = "Libro":                .ColWidth(5) = 1000:  .ColAlignment(5) = flexAlignLeftCenter:   .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 6) = "Nº.Cuenta":            .ColWidth(6) = 1230:  .ColAlignment(6) = flexAlignLeftCenter:   .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 7) = "Nombre de la Cuenta":  .ColWidth(7) = 3000:  .ColAlignment(7) = flexAlignLeftCenter:   .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 8) = "T.C.":                 .ColWidth(8) = 600:   .ColAlignment(8) = flexAlignRightCenter:  .Row = 0: .Col = 8: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 9) = "Debe":                 .ColWidth(9) = 1100:  .ColAlignment(9) = flexAlignRightCenter:   .Row = 0: .Col = 9: .CellAlignment = flexAlignRightCenter
        .TextMatrix(0, 10) = "Haber":               .ColWidth(10) = 1100: .ColAlignment(10) = flexAlignRightCenter: .Row = 0: .Col = 10: .CellAlignment = flexAlignRightCenter
        
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then CONSULTAR
    If Button.Index = 2 Then IMPRIMIR
    If Button.Index = 3 Then Exportar
    If Button.Index = 5 Then
        Unload Me
        Exit Sub
    End If
End Sub


Private Sub Exportar()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO, T_RPT_PERIODO, T_RPT_PERIODO1, "Asientos Descuadrados"
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub
