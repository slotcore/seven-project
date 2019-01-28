VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmAnalizaPrecio_Item_det 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analizar Precio de Compra al detalle"
   ClientHeight    =   5700
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   9000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9000
   Begin VB.Frame fr 
      Height          =   1170
      Index           =   5
      Left            =   30
      TabIndex        =   3
      Top             =   -75
      Width           =   8925
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   6
         Left            =   4245
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   825
         Width           =   1395
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   5
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   870
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   4
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   515
         Width           =   2250
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   810
         Width           =   1065
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   515
         Width           =   4680
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   210
         Width           =   2040
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   210
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Pre. Unit."
         Height          =   195
         Index           =   5
         Left            =   3465
         TabIndex        =   19
         Top             =   885
         Width           =   660
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Item:"
         Height          =   195
         Index           =   4
         Left            =   5745
         TabIndex        =   9
         Top             =   615
         Width           =   705
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Unid. Med."
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   885
         Width           =   780
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Decripción:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   615
         Width           =   810
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   495
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   3180
         TabIndex        =   5
         Top             =   300
         Visible         =   0   'False
         Width           =   165
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   4095
      Left            =   0
      TabIndex        =   4
      Top             =   1590
      Width           =   7650
      _cx             =   13494
      _cy             =   7223
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
      SelectionMode   =   0
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
      FormatString    =   $"FrmAnalizaPrecio_Item_det.frx":0000
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
   Begin VB.Frame fm 
      Height          =   450
      Left            =   30
      TabIndex        =   15
      Top             =   1110
      Width           =   8925
      Begin VB.Label LblNombre 
         BackStyle       =   0  'Transparent
         Caption         =   "LblNombre(1)"
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
         Left            =   7620
         TabIndex        =   23
         Top             =   180
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label LblNombre 
         BackStyle       =   0  'Transparent
         Caption         =   "LblNombre(0)"
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
         Left            =   90
         TabIndex        =   22
         Top             =   180
         Width           =   8745
      End
      Begin VB.Label x_lbl 
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   6
         Left            =   60
         TabIndex        =   16
         Top             =   165
         Visible         =   0   'False
         Width           =   4200
      End
   End
   Begin VB.Frame fr 
      Height          =   4185
      Index           =   0
      Left            =   7695
      TabIndex        =   17
      Top             =   1470
      Width           =   1260
      Begin VB.CommandButton cmd 
         Caption         =   "&Exportar"
         Height          =   465
         Index           =   3
         Left            =   105
         TabIndex        =   21
         Top             =   1225
         Width           =   1080
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Salir"
         Height          =   465
         Index           =   2
         Left            =   105
         TabIndex        =   2
         Top             =   1740
         Width           =   1080
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Actualizar"
         Height          =   465
         Index           =   0
         Left            =   105
         TabIndex        =   0
         Top             =   195
         Width           =   1080
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Imprimir"
         Height          =   465
         Index           =   1
         Left            =   105
         TabIndex        =   1
         Top             =   710
         Width           =   1080
      End
   End
End
Attribute VB_Name = "FrmAnalizaPrecio_Item_det"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- ALMACENAR LOS TOTALES DE TODA LA CONSULTA
Dim Arr_Totales_cols() As Double '--ALMACENAR TOTALES POR TODAS LAS FILAS
Dim Arr_Totales_col() As Double     '--ALMACENAR TOTALES POR COLUMNA, SE LIMPIA DESPUES DE CAMBIO DE GRUPO

'--DE LA IMPRESION
Dim T_RPT_PERIODO As String '--PERIODO DEL REPORTE
Dim T_RPT_TITULO As String  '--TITULO DE REPORTE
'------------
Dim ARR_XX() As String      '--SE CARGARA CUANDO SE CARGA EL FORMULARIO Y CUANDO SE CAMBIE EL ESTILO(MES, TRIMESTRE,SEMESTRE)
Dim ARR_TMP() As String     '--DEPENDERA DEL ESTILO SOLO CARGARA LO QUE SELECCIONA


                            '--SE USA PARA DAR FORMATO DE LA GRILLA, SEGUN SELECCIONE EL USUARIO
Dim Q_TOTAL_ANYO As Integer '--INDICA LA CANTIDAD DE AÑOS DE BUSQUEDA,
                            '--EJ. 2004,2005 => Q_TOTAL_ANYO = 2
                            '--EJ. 2004,2005,2006 => Q_TOTAL_ANYO = 3
                            
Dim Q_COL_FILA As Integer   '--INDICA LA CANTIDAD DE COLUMNAS ANTES DE LOS MESES
                            '--EJ. IDCLI,CLIENTE => Q_COL_FILA=2
                            '--    IDCLI,ID_PTO_VTA,CLIENTE,PTO_VENTA => Q_COL_FILA=4
                            
Dim Q_POS_MES_INICIO As Integer '--INDICA LA POSICION INICIAL DE LA COLUMNA DEL PRIMER MES, NO CAMBIA
                            '--EJ. Q_POS_MES_INICIO = Q_COL_FILA +1

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
                                 '--OBTENDRA VALOR EN pGenerarConsulta()

Dim Q_COL_COMPARAR_GRUPO As Integer '--INDICA LA COLUMNA PARA COMPARAR EL GRUPO
                                    '--OBTENDRA VALOR EN pGenerarConsulta()

Dim Q_COL_ARR_TOTAL As Integer  '--NOS INDICA EL TOTAL DE COLUMNAS QUE VA A CONTENER LOS ACUMULADOS
                                '--OBTENDRA VALOR EN VALIDAR_CONSULTA()
                                '--SI SEL MES: ENE, FEB, MAR => Q_COL_ARR_TOTAL= 2
                                '--SI SEL TRI: ENE-MAR, ABR-JUN => Q_COL_ARR_TOTAL= 1 OBS: SE INICIA DESDE POS=0
                                
Dim mAnyo As String '--INDICA EL AÑO QUE SELECCIONA
Dim S_PREUNIT As Double '--PRECIO UNITARIO

Dim SGIFlex As New SGI2_funciones.JC_VSFlexGrid
Dim SGIVarios As New SGI2_funciones.JC_Varios


Dim F_ES_COMPRA As Boolean '--INDICA SI ES COMPRA O VENTA
                            '--TRUE::ES COMPRA, FALSE::ES VENTA




Public Sub RECIBE_ID_ITEM(mAnyo_ACTIVO As String, _
                            S_PRECIO_UNITARIO As String, _
                            X_ARRAY_TMP() As String, _
                            Optional F_VENTANA_COMPRA As Boolean = True)
                            
    '--N_COL_NOMBRE :: Ene, Feb,Mar, Ene-Mar, Abr-Jun
    '--DEPENDERA DEL ARRAY: ARR_TMP()
    On Error GoTo ERROR
    Dim Q_ROW       As Integer
    '--DEL AÑO
    mAnyo = mAnyo_ACTIVO
    '--txt(0).Text '--ID_ITEM (IDENTIFICADOR DE REGISTRO)
    For Q_ROW = 0 To FrmAnalizaPrecio_Item.txt.Count - 1
        txt(Q_ROW).Text = FrmAnalizaPrecio_Item.txt(Q_ROW).Text
    Next Q_ROW
    txt(6).Text = S_PRECIO_UNITARIO
    '---------
    '--DEL ARRAY TMP
    
    Limpiar_ARRAY_TOTAL True
    ReDim ARR_TMP(0, 2)
    Dim POS_ARR As Integer
    POS_ARR = 0
    For Q_ROW = 0 To UBound(X_ARRAY_TMP())
        ARR_TMP(POS_ARR, 0) = X_ARRAY_TMP(Q_ROW, 0)
        ARR_TMP(POS_ARR, 1) = X_ARRAY_TMP(Q_ROW, 1)
        ARR_TMP(POS_ARR, 2) = X_ARRAY_TMP(Q_ROW, 2)
        POS_ARR = POS_ARR + 1
    Next
    Q_COL_ARR_TOTAL = 0
    '------
    x_lbl(6).Caption = FrmAnalizaPrecio_Item.x_lbl(6).Caption
    '--------
    '--si selecciona un producto
    SGIVarios.LimpiaText LblNombre, True
    If NulosN(FrmAnalizaPrecio.lbl_cod(0).Caption) <> 0 Then
        LblNombre(0).Caption = FrmAnalizaPrecio.lbl_cb(0).Caption
        LblNombre(1).Caption = FrmAnalizaPrecio.lbl_cod(0).Caption
    End If
    '-----------

    F_ES_COMPRA = F_VENTANA_COMPRA
    If F_ES_COMPRA = False Then Me.Caption = "Analizar Precio de Venta al detalle"
    '------
    pConsultar
    Me.MousePointer = vbDefault
    Exit Sub
salir:
    SGIVarios.habilitar cmd, False
    Exit Sub
ERROR:
    Me.MousePointer = vbDefault
    SGIVarios.SHOW_ERROR
End Sub


Private Sub cmd_Click(index As Integer)
    Select Case index
    Case 0 '--CONSULTAR
        SGIFlex.LimpiarGrid Fg1, , 1
        pConfigurarGrilla
        pConsultar
    Case 1 '--IMPRIMIR
        pImprimir
    Case 2  '--SALIR
        Unload Me
    Case 3 '--EXPORTAR
        pExportar
    End Select
End Sub



Private Sub pConsultar()
    'On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    '--
    Dim CN_TMP As New ADODB.Connection '--CONEX TEMPORAL
    Dim Rst_RUTA As New ADODB.Recordset '--CARGA RUTAS DE BD'S
    
    Dim vStrSelect As String '--RECIBIR LA CONSULTA
    '--CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    Dim SQL_ANYO As String
    Dim k As Integer
    
    If Validar_Consulta() = False Then Exit Sub
    
    '--INVOCAR A ESTA FUNCION PARA OBTENER LOS VALORES DE
        '--Q_POS_MES , Q_POS_MES_INICIO
    pConfigurarGrilla
    '--CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    SQL_ANYO = " AND anotra = " + mAnyo + " "
    '--SI LA BASE DE BATOS PRINCIPAL EXISTE
    If SGIVarios.ArchivoExiste(AP_RUTABD + "data.mdb") = False Then
        MsgBox "No existe la ruta a la Base de Datos Principal", vbCritical, xTitulo
        Exit Sub
    End If
    '--ABRIENDO LA CONEXION PARA OBTENER EL LISTADO DE RUTAS A LAS BASES DE DATOS
    SGIVarios.OPEN_CONEX_TMP CN_TMP, AP_RUTABD + "data.mdb"
    If CN_TMP.State = 0 Then Exit Sub
    RST_Busq rst_select, "SELECT ruta,anotra FROM mae_empresa WHERE numruc = '" + NumRUC + "' " + SQL_ANYO + " ORDER BY anotra ASC ", CN_TMP
    '--CARGAR RST TEMPORAL
    SGIVarios.DEFINIR_RST_TMP Rst_RUTA, rst_select
    SGIVarios.CARGAR_RST_TMP Rst_RUTA, rst_select
    If Rst_RUTA.RecordCount = 0 Then
        MsgBox "No hay Base de Datos", vbInformation
        Exit Sub
    End If
    Rst_RUTA.MoveFirst
    Set rst_select = Nothing
    CN_TMP.Close
    '--LIMPIAR ARRAY
    Limpiar_ARRAY_TOTAL True
    '----
    Me.MousePointer = vbHourglass
    DoEvents
    '------------------------------------------------
    '--ENTRAR SOLO UNA VEZ
    vStrSelect = pGenerarConsulta()
    '------------------------------------------------
    If vStrSelect = "" Then GoTo salir
    '--SI EL ARCHIVO EXISTE
    If SGIVarios.ArchivoExiste(AP_RUTABD + Rst_RUTA.Fields(0) & "") = False Then
        MsgBox "No existe la ruta a la Base de Datos Año: " + CStr(Rst_RUTA.Fields(1)), vbCritical, xTitulo
        GoTo salir
    End If
    '--ABRIENDO LA CONEXION A LA BASE DE DATOS
    SGIVarios.OPEN_CONEX_TMP CN_TMP, AP_RUTABD + Rst_RUTA.Fields(0) & ""
    If CN_TMP.State = 0 Then Exit Sub
    '--CARGADO EL RST
    RST_Busq rst_select, vStrSelect, CN_TMP

    '--------------------------------------
    '--CARGA LOS DATOS DEL PRIMER AÑO
    CARGAR_DATOS_GRILLA rst_select, CStr(Rst_RUTA.Fields(1))
    CN_TMP.Close
    '--------------------------------------
salir:
    Set Rst_RUTA = Nothing
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
ERROR:
    Me.MousePointer = vbDefault
    Set rst_select = Nothing
    CN_TMP.Close
    SGIVarios.SHOW_ERROR
    
End Sub


Private Function CARGAR_DATOS_GRILLA(RST_ORIGEN As ADODB.Recordset, _
                                    mAnyo As String)
                                         
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
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        '---------------------------------------------------------
        '---------------------------------------------------------
        pCompararGrupo RST_ORIGEN, BAND_ADD_REG, mAnyo
        
        If RST_ORIGEN.Bookmark <> 1 Then SGIFlex.ADD_REG Fg1
        '---------------------------------------------------------
        '--ACUMULAR EN EL ARRAY_MES
        CARGAR_DATOS_ARRAY RST_ORIGEN
        '--CARGAR A LA GRILLA
        CARGAR_DATOS_GRILLA_ARRAY_TMP RST_ORIGEN, Fg1.Rows - 1
        '---------------------------------------------------------
        RST_ORIGEN.MoveNext
        
        If RST_ORIGEN.EOF Then
            CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:"
            CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Tot Gen:", True, True
        End If
        
        
    Wend
    
    '------

End Function

Private Sub pCompararGrupo(RST_ORIGEN As ADODB.Recordset, _
                            BAND_ADD_REG As Boolean, _
                            mAnyo As String, _
                            Optional Q_COL_COMPARAR As Integer = -1)
                            
    '--FUNCION QUE NOS PERMITE ARMAR LOS GRUPOS
    '--COMPARA CUANDO CAMBIAR DE GRUPO
    Dim RST_TEPM_1 As New ADODB.Recordset
    
    '---------------------------------------------------------
    If Q_COL_COMPARAR_GRUPO = -1 Then GoTo salir
    '---------------------------------------------------------
    If Q_COL_COMPARAR = -1 Then Q_COL_COMPARAR = Q_COL_COMPARAR_GRUPO
    
    
    If RST_ORIGEN.Bookmark = 1 Then
        '--SE CARGA EN pGenerarConsulta() Q_COL_COMPARAR_GRUPO
        SGIFlex.ADD_REG Fg1, Fila_grupo
        SGIFlex.UNIR_CELDAS Fg1, Fg1.Rows - 1, 3, Fg1.Rows - 1, 6, RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter:
        SGIFlex.FORMATO_CELDA Fg1, Fg1.Rows - 1, 3
        SGIFlex.ADD_REG Fg1, Fila_Ninguno
    Else
    
        Set RST_TEPM_1 = RST_ORIGEN.Clone
        RST_TEPM_1.Bookmark = RST_ORIGEN.Bookmark
        RST_TEPM_1.MovePrevious
        
        If RST_TEPM_1.Fields(Q_COL_COMPARAR) <> RST_ORIGEN.Fields(Q_COL_COMPARAR) Then
            CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:", , , mAnyo
            SGIFlex.ADD_REG Fg1, Fila_en_Blanco
            Limpiar_ARRAY_TOTAL

            SGIFlex.ADD_REG Fg1, Fila_grupo
            SGIFlex.UNIR_CELDAS Fg1, Fg1.Rows - 1, 3, Fg1.Rows - 1, 6, RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter
            SGIFlex.FORMATO_CELDA Fg1, Fg1.Rows - 1, 3
        End If
        
    End If
salir:
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
        vStrCampo = RST_ORIGEN.Fields(Q_CAMPO).Name
        '--OBS: SE VA LLENAR EL ARRAY "MONTOS DEL TOTAL" O "MONTOS DEL RESUMEN"
        Select Case LCase(vStrCampo)
            Case "canpro":          Arr_Totales_col(0, 0) = Arr_Totales_col(0, 0) + NulosN(RST_ORIGEN.Fields("canpro"))
        End Select
    Next Q_CAMPO
    
End Sub


Private Function CARGAR_DATOS_GRILLA_ARRAY_TMP(RST_ORIGEN As ADODB.Recordset, _
                                         Q_ROW As Integer)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim Q_CAMPO As Integer
    Dim vStrCampo As String
       
    DoEvents
    For Q_CAMPO = 0 To RST_ORIGEN.Fields.Count - 1
        vStrCampo = RST_ORIGEN.Fields(Q_CAMPO).Name
        '--COLOCANDO LOS VALORES EN LA GRILLA
            Select Case LCase(vStrCampo)
                Case "canpro"
                    Fg1.TextMatrix(Fg1.Rows - 1, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), SGIFlex.FORMAT_CANTIDAD)
                Case "impven"
                    Fg1.TextMatrix(Fg1.Rows - 1, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), SGIFlex.FORMAT_IMPUESTO)
                Case "fchdoc", "fchven"
                    Fg1.TextMatrix(Fg1.Rows - 1, Q_CAMPO + 1) = Format(RST_ORIGEN.Fields(vStrCampo), SGIFlex.FORMAT_DATE)
                Case Else
                    Fg1.TextMatrix(Fg1.Rows - 1, Q_CAMPO + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
            End Select
        '------------
    Next
End Function


Private Sub pImprimir()

    On Error GoTo ERROR
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    
    If F_ES_COMPRA = False Then T_RPT_TITULO = Replace(T_RPT_TITULO, "COMPRA", "VENTA")
    
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO + " ", "ITEM: " + txt(2).Text, x_lbl(6).Caption, False, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
ERROR:
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    SGIVarios.SHOW_ERROR
'    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"

End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then Unload Me
End Sub

Private Sub Form_Load()
    Dim k As Integer
    SGIVarios.CentrarFrm Me
    SGIFlex.LimpiarGrid Me.Fg1, , 1
    pConfigurarGrilla
End Sub

'------
Private Function Validar_Consulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA

    If mAnyo = "" Then
        MsgBox "No hay año activo ", vbCritical, xTitulo
    End If
    Q_TOTAL_ANYO = 1
    '-----------
    Validar_Consulta = True

End Function

Private Function pGenerarConsulta() As String
    '--FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    '--
    Dim vStrSelect As String            '--CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    Dim vStrFiltro_ITEM As String       '--SOLO ITEM
    
    Dim vStrFiltro As String

    Dim k As Integer
    '--DEL AÑO
    vStrFiltro = " Year(com_compras.fchdoc)= " + mAnyo + " "
    '--DEL ITEM
    vStrFiltro_ITEM = " AND com_comprasdet.iditem= " + Trim(txt(0).Text) + " "
    '--SOLO s/.
    If FrmAnalizaPrecio.opt_mon(0).Value = True Then vStrFiltro = vStrFiltro + " AND com_compras.idmon= 1 " '--SOLES
    '--SOLO $
    If FrmAnalizaPrecio.opt_mon(1).Value = True Then vStrFiltro = vStrFiltro + " AND com_compras.idmon= 2 " '--DOLARES
    '--si se selecciono un proveedor o cliente
    If NulosN(LblNombre(1).Caption) <> 0 Then vStrFiltro = vStrFiltro + " AND com_compras.idpro= " & NulosN(LblNombre(1).Caption)
    
    '--GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim nSQLValor As String
    Dim nSQLCampos As String
    Dim nSQLWhere As String
    Dim nSQLFrom As String
    Dim nSQLGroupBy As String
    Dim nSQLOrderBy As String
    Dim nSQLPivot As String
    Dim nSQLPivotSalida As String '--ORDENA LOS VALORE MES A MES(ENE,FEB,MAR,ETC.)
    nSQLWhere = vStrFiltro + vStrFiltro_ITEM
    '--DEL PRECIO UNITARIO
    nSQLWhere = nSQLWhere + " AND format(com_comprasdet.preuni, '" + SGIFlex.FORMAT_PU + "')= '" + Format(Trim(txt(6).Text), SGIFlex.FORMAT_PU) + "' "
    'nSQLWhere = nSQLWhere + " AND round(com_comprasdet.preuni, 6)= " + Format(Trim(txt(6).Text), FORMAT_PU) + " "
   
    Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 0:       Q_POSICION_TOTAL = 6:           Q_COL_COMPARAR_GRUPO = 1
    
    '--DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    Q_POS_MES_INICIO = Q_COL_FILA
    '------------------------------------------
    If FrmAnalizaPrecio.opt_estilo(0).Value = True Then  '--MES
        T_RPT_TITULO = "LISTADO DE COMPRAS MENSUAL"
        nSQLPivot = "FORMAT(com_compras.fchdoc,'m') "
    ElseIf FrmAnalizaPrecio.opt_estilo(1).Value = True Then '--TRIMESTRE
        T_RPT_TITULO = "LISTADO DE COMPRAS TRIMESTRAL"
        nSQLPivot = "FORMAT(com_compras.fchdoc,'q') "
    ElseIf FrmAnalizaPrecio.opt_estilo(2).Value = True Then '--SEMESTRE
        T_RPT_TITULO = "LISTADO DE COMPRAS SEMESTRAL"
        nSQLPivot = "FORMAT(com_compras.fchdoc,'s') "
    End If
    '--DEL PIVOT
    For k = 0 To UBound(ARR_TMP)
        nSQLPivotSalida = nSQLPivotSalida + ARR_TMP(k, 2) + ","
    Next k
    nSQLPivotSalida = " IN (" + Left(nSQLPivotSalida, Len(nSQLPivotSalida) - 1) + ") "
    nSQLWhere = nSQLWhere + " AND " + nSQLPivot + nSQLPivotSalida
    
    nSQLCampos = " mae_prov.id, mae_prov.nombre AS nombre, IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='',[com_compras].[numreg],Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4)) AS registro, mae_documento.abrev AS tdocabrev, com_compras!numser+'-'+com_compras!numdoc AS numerodoc, com_compras.fchdoc, mae_condpago.abrev AS conpagabre, mae_moneda.simbolo, mae_unidades.abrev AS prodabrev, com_comprasdet.canpro "
    nSQLFrom = "  mae_libros RIGHT JOIN (mae_condpago RIGHT JOIN (mae_prov RIGHT JOIN (mae_unidades RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_moneda.id = com_compras.idmon) ON mae_unidades.id = com_comprasdet.idunimed) ON mae_prov.id = com_compras.idpro) ON mae_condpago.id = com_compras.idconpag) ON mae_libros.id = com_compras.idlib "
 
    nSQLOrderBy = " mae_prov.nombre, com_compras.fchdoc "
        
    '--GENERANDO LA CONSULTA
    vStrSelect = "SELECT " + nSQLCampos + _
        vbCr + " FROM " + nSQLFrom + _
        vbCr + " WHERE " + nSQLWhere + _
        vbCr + " ORDER BY " + nSQLOrderBy
        
    '--SI ES POR VENTA
    If F_ES_COMPRA = False Then
        vStrSelect = Replace(vStrSelect, "com_comprasdet.idcom", "vta_ventasdet.idvta")
        vStrSelect = Replace(vStrSelect, ".idpro", ".idcli")
        vStrSelect = Replace(vStrSelect, "com_comprasdet", "vta_ventasdet")
        vStrSelect = Replace(vStrSelect, "com_compras", "vta_ventas")
        vStrSelect = Replace(vStrSelect, "mae_prov", "mae_cliente")
        vStrSelect = Replace(vStrSelect, "WHERE ", "WHERE vta_ventas.anulado=0 AND ")
    End If
    '------------------------------------------------------------------------------------
    pGenerarConsulta = vStrSelect
    
    
End Function


'--011007

Private Sub Limpiar_ARRAY_TOTAL(Optional F_LIMPIA_TOT_GRL As Boolean = False)
    Erase Arr_Totales_col
    ReDim Arr_Totales_col(2, 0) As Double
    If F_LIMPIA_TOT_GRL = True Then
        Erase Arr_Totales_cols
        ReDim Arr_Totales_cols(2, 0)
    End If
End Sub
'''
Private Sub CARGAR_DATOS_GRILLA_ADD_TOTALES(BAND_ADD_TOTAL As Boolean, _
                                            Nombre_total As String, _
                                            Optional fTotalGral As Boolean = False, _
                                            Optional fForzarSuma As Boolean = False, _
                                            Optional mAnyo As String)
                
    Dim Q_MES As Integer
    '--AGREGA LOS TOTALES POR CADA GRUPO Y EL TOTAL GENERAL
    '--ACUMULA LOS TOTALES EN EL TOTAL GENERAL
    Dim X_ROW  As Long, k As Integer
    Dim Q_INCREMENTO_X_COL As Integer   '--SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer     '--POSICIONA A LA COLUMNA DEL TOTAL X AÑO

    'On Error Resume Next
    X_ROW = Fg1.Rows
    If BAND_ADD_TOTAL = True Then
       '--AGREAGNDO NUEVA FILA
        SGIFlex.ADD_REG Fg1, IIf(fTotalGral = False, Fila_Total, Fila_Total_grl)

        'PONIENDO LOS NOMBRES DE LOS TOTALES  Q_POSICION_TOTAL SE OBTIENE DE pGenerarConsulta()
        Fg1.TextMatrix(X_ROW, Q_POSICION_TOTAL) = Nombre_total
        SGIFlex.FORMATO_CELDA Fg1, X_ROW, Q_POSICION_TOTAL
    End If
 
    If fTotalGral = False Then     '--DETALLE
        For k = 0 To UBound(Arr_Totales_col())
            Arr_Totales_cols(k, 0) = Arr_Totales_cols(k, 0) + Arr_Totales_col(k, 0)
        Next k
    End If
    '
            
    Fg1.TextMatrix(X_ROW, 10) = PONER_FORMATO(IIf(fTotalGral = False, Arr_Totales_col(0, 0), Arr_Totales_cols(0, 0)), fTotalGral): SGIFlex.FORMATO_CELDA Fg1, X_ROW, 10   '"cantidad"
    '-----------
    SGIFlex.FORMATO_CELDA Fg1, X_ROW, 10
        

    Err.Clear
End Sub

Private Sub pConfigurarGrilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA

    If F_CONSERVAR_FORMATO = True Then Fg1.Clear
    
    With Fg1
        Fg1.Cols = 11
        .FrozenCols = 6
        .TextMatrix(0, 1) = "ID":               .ColWidth(1) = 700

        .TextMatrix(0, 2) = "Proveedor":        .ColWidth(2) = 1500
        If F_ES_COMPRA = False Then .TextMatrix(0, 2) = "Cliente"
        
        .TextMatrix(0, 3) = "N°.Reg.":          .ColWidth(3) = 900: .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(0, 4) = "T.D.":             .ColWidth(4) = 450: .ColAlignment(3) = flexAlignCenterCenter
        
        .TextMatrix(0, 5) = "Num. Documento":   .ColWidth(5) = 1500: .ColAlignment(5) = flexAlignCenterCenter
        .TextMatrix(0, 6) = "Fch.Doc.":         .ColWidth(6) = 850: .ColAlignment(6) = flexAlignCenterCenter
        .TextMatrix(0, 7) = "Cond.Pago":        .ColWidth(7) = 950: .ColAlignment(7) = flexAlignCenterCenter
        .TextMatrix(0, 8) = "M":                .ColWidth(8) = 500: .ColAlignment(8) = flexAlignCenterCenter
        
        .TextMatrix(0, 9) = "U.M.":             .ColWidth(9) = 500: .ColAlignment(9) = flexAlignLeftCenter
        .TextMatrix(0, 10) = "Cant.":           .ColWidth(10) = 1000: .ColAlignment(10) = flexAlignRightCenter
        
        SGIFlex.OCULTAR_COL Fg1, 1, 2
    End With
    DoEvents
End Sub

Private Function PONER_FORMATO(S_MONTO As Double, _
                        Optional fTotalGral As Boolean = False) As String
                        
    '--ESTA FUNCION CONVERTIRA AL FORMATO
    If S_MONTO = 0 Then
            PONER_FORMATO = "0.00"
        Exit Function
    End If
    PONER_FORMATO = Format(S_MONTO, SGIFlex.FORMAT_CANTIDAD)
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Erase ARR_TMP
    Erase Arr_Totales_cols
    Erase Arr_Totales_col
    
    Set SGIFlex = Nothing
    Set SGIVarios = Nothing

End Sub



Private Sub txt_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If index <> 5 Then Exit Sub
    If SGIVarios.validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub pExportar()
    On Error GoTo ERROR
    Dim X_EXPORT As New SGI2_funciones.formularios
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, Me.Caption, "Item: " + txt(2).Text, "Tipo: " + txt(4).Text, Me.Caption
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
ERROR:
    Me.MousePointer = vbDefault
    SGIVarios.SHOW_ERROR Me.Name, "pExportar"
End Sub
