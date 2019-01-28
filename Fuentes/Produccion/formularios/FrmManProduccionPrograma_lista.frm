VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManProduccionPrograma_lista 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6105
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   11700
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Align           =   2  'Align Bottom
      Height          =   5265
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   11700
      _cx             =   20637
      _cy             =   9287
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
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmManProduccionPrograma_lista.frx":0000
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
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
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4830
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
               Picture         =   "FrmManProduccionPrograma_lista.frx":003C
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":0580
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":0912
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":0A96
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":0EEA
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":1002
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":1546
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":1A8A
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":1B9E
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":1CB2
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":2106
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":2272
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":27BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccionPrograma_lista.frx":2AD4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fr 
      Height          =   570
      Index           =   5
      Left            =   30
      TabIndex        =   0
      Top             =   270
      Width           =   11685
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl(2)"
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
         Index           =   2
         Left            =   2580
         TabIndex        =   7
         Top             =   165
         Width           =   3345
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl(0)"
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
         Left            =   600
         TabIndex        =   6
         Top             =   165
         Width           =   1170
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl(1)"
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
         Left            =   7005
         TabIndex        =   5
         Top             =   165
         Width           =   4530
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         Height          =   195
         Index           =   4
         Left            =   1890
         TabIndex        =   4
         Top             =   270
         Width           =   540
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Responsable"
         Height          =   195
         Index           =   2
         Left            =   6015
         TabIndex        =   3
         Top             =   270
         Width           =   930
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   2
         Top             =   270
         Width           =   450
      End
   End
End
Attribute VB_Name = "FrmManProduccionPrograma_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANPRODUCCIONPROGRAMA_LISTA.FRM
'* Tipo             : FORMULARIO
'* Descripcion      :
'*
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 06/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit
'-- ALMACENAR LOS TOTALES DE TODA LA CONSULTA
'--ARR_TMP(?,4)= Arr_Totales_cols() As Double '--ALMACENAR TOTALES POR TODAS LAS FILAS
'--ARR_TMP(?,3)= Arr_Totales_col() As Double     '--ALMACENAR TOTALES POR COLUMNA, SE LIMPIA DESPUES DE CAMBIO DE GRUPO


Dim BAND_INTERRUMPIR As Boolean           ' SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA     TRUE SE INTERRUMPE
' DE LA IMPRESION
Dim T_RPT_PERIODO As String               ' PERIODO DEL REPORTE
Dim T_RPT_TITULO As String                ' TITULO DE REPORTE
Dim ARR_ANYO() As String                  ' ARRAY DE AÑOS SELECCIONADOS
Dim ARR_XX() As String                    ' SE CARGARA CUANDO SE CARGA EL FORMULARIO Y CUANDO SE CAMBIE EL ESTILO(MES, TRIMESTRE,SEMESTRE)
Dim ARR_TMP() As String                   ' DEPENDERA DEL ESTILO SOLO CARGARA LO QUE SELECCIONA
Dim ARR_TMP_1() As String                 ' ACUMULARA LOS TOTALES DE STOCK_ACTUAL, SALDO=STOCK_ACTUAL - TOTAL
                                          ' SE USA PARA DAR FORMATO DE LA GRILLA, SEGUN SELECCIONE EL USUARIO
Dim Q_TOTAL_ANYO As Integer               ' INDICA LA CANTIDAD DE AÑOS DE BUSQUEDA,
                                          ' EJ. 2004,2005 => Q_TOTAL_ANYO = 2
                                          ' EJ. 2004,2005,2006 => Q_TOTAL_ANYO = 3
Dim Q_COL_FILA As Integer                 ' INDICA LA CANTIDAD DE COLUMNAS ANTES DE LOS MESES
                                          ' EJ. IDCLI,CLIENTE => Q_COL_FILA=2
                                          ' IDCLI,ID_PTO_VTA,CLIENTE,PTO_VENTA => Q_COL_FILA=4
Dim Q_COL_FILA_ULTIMO As Integer          ' INDICA LA CANTIDAD DE COLUMNAS ADICIONALES QUE SE COLOCARAN DESPUES DEL TOTAL
Dim Q_POS_MES_INICIO As Integer           ' INDICA LA POSICION INICIAL DE LA COLUMNA DEL PRIMER MES, NO CAMBIA
                                          ' EJ. Q_POS_MES_INICIO = Q_COL_FILA +1
Dim Q_POS_MES As Integer                  ' INDICA LA POSICION DEL MES, ESTO CAMBIA UTIL PARA COLOCAR LOS DATOS EN EL GRID
Dim Q_COL_FILA_OCULTA As Integer          ' INDICA LAS COLUMNAS QUE CONTENDRAN LOS ID'S, ESTOS SE OCULTARAN
                                          ' -1 NO SE OCULTA, <> -1 SE PROCEDE A ACULTAR
                                          ' EJ. CLIENTE  vta_ventas.idcli,
                                          ' PUNTO DE VENTA vta_guia.idpunven
                                          ' PRODUCTO   alm_inventario.tippro
                                          ' ITEM       alm_inventario.id
                                          ' EMPLEADO   vta_ventas.idven
Dim Q_POSICION_TOTAL  As Integer          ' INDICA LA POCISION DE LA COLUMNA DONDE SE COLOCARA EL NOMBRE DEL TOTAL Y TOTAL_GRL OBTENDRA VALOR EN GENERAR_CONSULTA()
Dim Q_COL_COMPARAR_GRUPO As Integer       ' INDICA LA COLUMNA PARA COMPARAR EL GRUPO OBTENDRA VALOR EN GENERAR_CONSULTA()
Dim Q_COL_ARR_TOTAL As Integer            ' NOS INDICA EL TOTAL DE COLUMNAS QUE VA A CONTENER LOS ACUMULADOS OBTENDRA VALOR EN VALIDAR_CONSULTA()
                                          ' SI SEL MES: ENE, FEB, MAR => Q_COL_ARR_TOTAL= 2
                                          ' SI SEL TRI: ENE-MAR, ABR-JUN => Q_COL_ARR_TOTAL= 1 OBS: SE INICIA DESDE POS=0
Dim F_ES_COMPRA As Boolean                ' INDICA SI ES COMPRA O VENTA TRUE::ES COMPRA, FALSE::ES VENTA
Dim ID_PROGRAMA As String
Dim ID_RECETA As String
Dim TIPO_VENTANA As e_PROGRAMA
Dim ESTILO_VISTA As Integer

'*****************************************************************************************************
'* Nombre           : RECIBE_LINK_FRM
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE        |  TIPO         |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    ID_PROGRAMA1  |  String       |
'*                    ID_RECETA1    |  String       |
'*                    TIPO_VENTANA1 |  e_PROGRAMA   |
'*                    ARR_TMP1      |  String       |
'*                    ESTILO_VISTA1 |  Integer      |
'*                    D_EMISION     |  String       |
'*                    N_RESPONSABLE |  String       |
'*                    N_PERIODO     |  String       |
'* Devuelve         :
'*****************************************************************************************************
Public Sub RECIBE_LINK_FRM(ID_PROGRAMA1 As String, ID_RECETA1 As String, _
                            TIPO_VENTANA1 As e_PROGRAMA, ARR_TMP1() As String, ESTILO_VISTA1 As Integer, _
                            D_EMISION As String, N_RESPONSABLE As String, N_PERIODO As String)
    ID_PROGRAMA = ID_PROGRAMA1
    ID_RECETA = ID_RECETA1
    TIPO_VENTANA = TIPO_VENTANA1
    ESTILO_VISTA = ESTILO_VISTA1
    
    lbl(0).Caption = D_EMISION
    lbl(1).Caption = N_RESPONSABLE
    lbl(2).Caption = N_PERIODO
                                
    ' DEL NOMBRE DEL FRM
    Select Case TIPO_VENTANA
        Case 0: Me.Caption = "Consulta de Insumos"
        Case 1: Me.Caption = "Consulta de Tareas"
        Case 2: Me.Caption = "Consulta de Equipos"
    End Select
    Me.Caption = "Producción - Programa de Produccion - " & Me.Caption
                                
    On Error GoTo error
    
    Dim POS_ARR As Integer
    Q_COL_ARR_TOTAL = UBound(ARR_TMP1())
    Erase ARR_TMP()
    Erase ARR_TMP_1()
    ReDim ARR_TMP(Q_COL_ARR_TOTAL, 4)
    ReDim ARR_TMP_1(1, 1)   ' 0::STOCK=>> 0::TOTAL,1::TOTAL GEN
                            ' 1::SALDO=>> 0::TOTAL,1::TOTAL GEN
    POS_ARR = 0
    For POS_ARR = 0 To Q_COL_ARR_TOTAL
        ARR_TMP(POS_ARR, 0) = ARR_TMP1(POS_ARR, 0)
        ARR_TMP(POS_ARR, 1) = ARR_TMP1(POS_ARR, 1)
    Next POS_ARR
    
    pConsultar
    Exit Sub

SALIR:
    Exit Sub

error:
    SHOW_ERROR
End Sub

'*****************************************************************************************************
'* Nombre           : pConsultar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pConsultar()
    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    Dim CN_TMP As New ADODB.Connection    ' CONEX TEMPORAL
    Dim Rst_RUTA As New ADODB.Recordset   ' CARGA RUTAS DE BD'S
    Dim vStrSelect As String              ' RECIBIR LA CONSULTA
    ' CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    Dim N_ANYO As String
    Dim SQL_ANYO As String
    Dim k As Integer
    
    If Validar_Consulta() = False Then Exit Sub
    
    BAND_INTERRUMPIR = False
    
    ' CONFIGURAR LA PRESENTACION DE LA CONSULTA
    LimpiarGrid Me.Fg1, False, 1
    
    ' INVOCAR A ESTA FUNCION PARA OBTENER LOS VALORES DE
    '------------------------------------------------
    ' ENTRAR SOLO UNA VEZ
    vStrSelect = GENERAR_CONSULTA()
    Configurar_Grilla
        
    ' LIMPIAR ARRAY
    Limpiar_ARRAY_TOTAL True
    
    Me.MousePointer = vbHourglass
    DoEvents
    
    If vStrSelect = "" Then GoTo SALIR
    ' CARGADO EL RST
    RST_Busq rst_select, vStrSelect, xCon
   
    pCargarDatosGrid rst_select

SALIR:
    Set Rst_RUTA = Nothing
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    Set rst_select = Nothing
    SHOW_ERROR
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosGrid
'* Tipo             : FUNCION
'* Descripcion      : FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
'* Paranetros       : NOMBRE      |  TIPO            |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    RST_ORIGEN  |  ADODB.Recordset |
'* Devuelve         :
'*****************************************************************************************************
Private Function pCargarDatosGrid(RST_ORIGEN As ADODB.Recordset)
    Dim BAND_ADD_REG As Boolean
    
    BAND_ADD_REG = True
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    
    While Not RST_ORIGEN.EOF
    
    DoEvents
        ' SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then Exit Function
        
        Comparar_Grupo RST_ORIGEN, BAND_ADD_REG
        
        If RST_ORIGEN.Bookmark <> 1 Then ADD_REG Fg1
        ' ACUMULAR EN EL ARRAY_MES
        pCargarDatosArray RST_ORIGEN
        ' CARGAR A LA GRILLA
        pCargarDatosGridArrayTmp RST_ORIGEN, Fg1.Rows - 1
        
        RST_ORIGEN.MoveNext
        
        ' PONER TOTALES AL FINAL DE LA GRILLA
        If RST_ORIGEN.EOF Then
            pCargarDatosGridAddTotales BAND_ADD_REG, "Total:"
            Select Case ESTILO_VISTA
            Case 0, 1, 4, 5, 8, 9
            Case Else
                pCargarDatosGridAddTotales True, "Tot Gen:", True
            End Select
        End If
    Wend
End Function

'*****************************************************************************************************
'* Nombre           : Comparar_Grupo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FUNCION QUE NOS PERMITE ARMAR LOS GRUPOS, COMPARA CUANDO CAMBIAR DE GRUPO
'* Paranetros       : NOMBRE         |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    RST_ORIGEN     |  ADODB.Recordset  |
'*                    BAND_ADD_REG   |  Boolean          |
'*                    Q_COL_COMPARAR |  Integer          |
'* Devuelve         :
'*****************************************************************************************************
Private Sub Comparar_Grupo(RST_ORIGEN As ADODB.Recordset, _
                            BAND_ADD_REG As Boolean, _
                            Optional Q_COL_COMPARAR As Integer = -1)
    Dim RST_TEPM_1 As New ADODB.Recordset
    
    If Q_COL_COMPARAR_GRUPO = -1 Then
        If RST_ORIGEN.Bookmark = 1 Then ADD_REG Fg1, Fila_Ninguno
        GoTo SALIR
    End If
    
    If Q_COL_COMPARAR = -1 Then Q_COL_COMPARAR = Q_COL_COMPARAR_GRUPO
    
    If RST_ORIGEN.Bookmark = 1 Then
        ' SE CARGA EN GENERAR_CONSULTA() Q_COL_COMPARAR_GRUPO
        ADD_REG Fg1, Fila_grupo
        UNIR_CELDAS Fg1, Fg1.Rows - 1, 3, Fg1.Rows - 1, 6, INICIO_GRUPO + RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter:
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 3
        ADD_REG Fg1, Fila_Ninguno
    Else
        Set RST_TEPM_1 = RST_ORIGEN.Clone
        RST_TEPM_1.Bookmark = RST_ORIGEN.Bookmark
        RST_TEPM_1.MovePrevious
        
        If RST_TEPM_1.Fields(Q_COL_COMPARAR) <> RST_ORIGEN.Fields(Q_COL_COMPARAR) Then
            pCargarDatosGridAddTotales BAND_ADD_REG, "Total:"
            ADD_REG Fg1, Fila_en_Blanco
            Limpiar_ARRAY_TOTAL

            ADD_REG Fg1, Fila_grupo
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 3, Fg1.Rows - 1, 6, INICIO_GRUPO + RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 3
        End If
    End If
    
SALIR:
    Set RST_TEPM_1 = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosArray
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FUNCION QUE ACUMULARA EN EL ARRAY_TEMP
'* Paranetros       : NOMBRE      |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    RST_ORIGEN  |  ADODB.Recordset  |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosArray(RST_ORIGEN As ADODB.Recordset)
    Dim vStrCampo As String
    Dim Q_CAMPO As Integer
    Dim Q_POS As Integer
    Q_POS = 0
    
    ' ASIGNAR LOS DATOS AL RECORDSET TEMPORAL
    For Q_CAMPO = 0 To RST_ORIGEN.Fields.Count - 1
        ' SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then Exit Sub
        vStrCampo = RST_ORIGEN.Fields(Q_CAMPO).Name
        ' OBS: SE VA LLENAR EL ARRAY "TOTAL"
        
        If InStr(LCase(vStrCampo), "/") <> 0 Then ' indica las fechas
            For Q_POS = 0 To UBound(ARR_TMP())
                ' EL CAMPO DEBE DE COINCIDIR CON EL ENCABEZADO DE LA GRILLA
                If Replace(ARR_TMP(Q_POS, 0), "'", "") = LCase(vStrCampo) Then
                    ARR_TMP(Q_POS, 3) = ARR_TMP(Q_POS, 3) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                    Exit For
                End If
            Next Q_POS
        ElseIf LCase(vStrCampo) = "total" Then
            ARR_TMP(Q_COL_ARR_TOTAL, 3) = ARR_TMP(Q_COL_ARR_TOTAL, 3) + NulosN(RST_ORIGEN.Fields(vStrCampo))
        ElseIf LCase(vStrCampo) = "stckact" Then
            ARR_TMP_1(0, 0) = ARR_TMP_1(0, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
        ElseIf LCase(vStrCampo) = "saldo" Then
            ARR_TMP_1(1, 0) = ARR_TMP_1(1, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
        Else
        End If
    Next Q_CAMPO
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosGridArrayTmp
'* Tipo             : FUNCION
'* Descripcion      : FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
'* Paranetros       : NOMBRE       |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    RST_ORIGEN   |  ADODB.Recordset  |
'*                    Q_ROW        |  Integer          |
'* Devuelve         :
'*****************************************************************************************************
Private Function pCargarDatosGridArrayTmp(RST_ORIGEN As ADODB.Recordset, _
                                         Q_ROW As Integer)
    Dim Q_INCREMENTO_X_COL As Integer   ' SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer     ' POSICIONA A LA COLUMNA DEL TOTAL X AÑO
    Dim Q_POS As Integer
    Dim Q_CAMPO As Integer
    Dim vStrCampo As String
    
    ' IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
    Q_INCREMENTO_X_COL = 0
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    
    DoEvents
    
    For Q_CAMPO = 0 To RST_ORIGEN.Fields.Count - 1
        If BAND_INTERRUMPIR = True Then Exit Function
        vStrCampo = RST_ORIGEN.Fields(Q_CAMPO).Name
        
        If InStr(LCase(vStrCampo), "/") <> 0 Then ' indica las fechas
            ' ARR_TMP(0, 2) INDICA LA PRIMERA COLUMNA A MOSTRAR
            If LCase(vStrCampo) = ARR_TMP(0, 2) Then Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    
            Fg1.TextMatrix(Q_ROW, Q_POS_MES) = PONER_FORMATO(NulosN(RST_ORIGEN.Fields(vStrCampo)), , Q_ROW)
            Q_POS_MES = Q_POS_MES + 1
        ElseIf LCase(vStrCampo) = "total" Then
            Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL) * 1
            ' TOTAL AÑO
            Fg1.TextMatrix(Q_ROW, Q_POS_MES_TOTAL) = PONER_FORMATO(NulosN(RST_ORIGEN.Fields(vStrCampo)), , Q_COL_ARR_TOTAL + 1)
        ElseIf LCase(vStrCampo) = "canpro" Then
            ' TOTAL AÑO
            Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PU)
        ElseIf LCase(vStrCampo) = "stckact" Then
            Fg1.TextMatrix(Q_ROW, Fg1.Cols - 2) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_CANTIDAD)
        ElseIf LCase(vStrCampo) = "saldo" Then
            Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_CANTIDAD)
        Else
            ' AGREGAR LOS DEMAS DATOS
            Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
        End If
    Next
End Function

'*****************************************************************************************************
'* Nombre           : pImprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pImprimir()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    If F_ES_COMPRA = False Then T_RPT_TITULO = Replace(T_RPT_TITULO, "COMPRA", "VENTA")
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO + " ", "Responsable: " + lbl(1).Caption, lbl(2).Caption, False, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR
End Sub

Private Sub Fg1_DblClick()
    'Fg1_KeyDown 13, 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo error
    
    ' CARGAR DATOS
    Dim DATOS_EMP As New SGI2_funciones.Varias
    DATOS_EMP.CargaDatosEmpresa xCon, NomEmp, NumRUC
    Set DATOS_EMP = Nothing
    CentrarFrm Me
    Exit Sub

error:
    SHOW_ERROR
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase ARR_TMP
End Sub

'*****************************************************************************************************
'* Nombre           : Validar_Consulta
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FUNCION QUE VALIDARA LA CONSULTA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Function Validar_Consulta() As Boolean
    Validar_Consulta = True
End Function

'*****************************************************************************************************
'* Nombre           : GENERAR_CONSULTA
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Function GENERAR_CONSULTA() As String
    Dim vStrSelect As String            ' CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    Dim vStrFiltro As String
    Dim k As Integer
    
    ' DEL PROGRAMA
    vStrFiltro = " pro_programadet.idprod = " + ID_PROGRAMA + " "
    
    ' DE LA RECETA
    If ID_RECETA <> "-1" Then vStrFiltro = vStrFiltro + " AND pro_programadet.idrec= " + ID_RECETA + " "
    
    ' GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim N_VALOR As String
    Dim N_CAMPOS As String
    Dim N_CAMPOS_ULTIMO  As String      ' INDICA LOS CAMPOS QUE SE COLOCARAN DESPPUES DEL TOTAL (EJ. TOTAL,STOCKACTUAL,SALDO)
    Dim N_WHERE As String
    Dim N_FROM As String
    Dim N_GROUP_BY As String
    Dim N_ORDER_BY As String
    Dim N_PIVOT As String
    Dim N_PIVOT_SALIDA As String        ' ORDENA LOS VALORE MES A MES(ENE,FEB,MAR,ETC.)
    N_WHERE = vStrFiltro
    Q_COL_FILA_ULTIMO = -1
    
    Select Case ESTILO_VISTA   ' PARAMETRO
        Case 0, 1, 2    ' 0 = INSUMO X PRODUCTO TODA PROGRAMACION   1 = INSUMO X PRODUCTO DIA ACTUAL   2 = INSUMO TODO PROD TODA PROGRAMACION
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 7:        Q_POSICION_TOTAL = 7:        Q_COL_COMPARAR_GRUPO = 2
            T_RPT_TITULO = "SDFSDF"
            N_CAMPOS = " pro_programadet.idrec, pro_recetains.iditem, pro_receta.descripcion AS recdesc, mae_tipoproducto.descripcion as tipprodesc,alm_inventario.descripcion AS itemdesc, mae_unidades.abrev, pro_recetains.canpro "
            Q_COL_FILA_ULTIMO = 2
            N_CAMPOS_ULTIMO = ", alm_inventario.stckact, ([alm_inventario].[stckact]-Sum([pro_programadet].[canpro]*[pro_recetains].[canpro])) AS saldo"
            N_GROUP_BY = " pro_programadet.idrec, pro_recetains.iditem, pro_receta.descripcion, mae_tipoproducto.descripcion,alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro,alm_inventario.stckact "
            N_ORDER_BY = " pro_receta.descripcion,mae_tipoproducto.descripcion, alm_inventario.descripcion "
        
        Case 3, 4       ' 3 = INSUMO TODO PROD DIA ACTUAL     4 = INSUMO LOS PRODUCTOS RESUMEN
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 4:        Q_POSICION_TOTAL = 4:        Q_COL_COMPARAR_GRUPO = -1
            N_CAMPOS = "  alm_inventario.id, mae_tipoproducto.descripcion as tipprodesc,alm_inventario.descripcion AS itemdesc, mae_unidades.abrev "
            Q_COL_FILA_ULTIMO = 2
            N_CAMPOS_ULTIMO = ", alm_inventario.stckact, ([alm_inventario].[stckact]-Sum([pro_programadet].[canpro]*[pro_recetains].[canpro])) AS saldo"
            N_GROUP_BY = " alm_inventario.id, mae_tipoproducto.descripcion,alm_inventario.descripcion, mae_unidades.abrev,alm_inventario.stckact "
            N_ORDER_BY = " mae_tipoproducto.descripcion,alm_inventario.descripcion "
            
        Case 5, 6, 7    ' 5 = TAREA X PRODUCTO TODA PROGRAMACION   ' 6 = TAREA X PRODUCTO DIA ACTUAL    7 = TAREA TODO PROD TODA PROGRAMACION
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 6:        Q_POSICION_TOTAL = 6:        Q_COL_COMPARAR_GRUPO = 2
            T_RPT_TITULO = "SDFSDF"
            N_CAMPOS = " pro_receta.id, pro_recetatar.idtar, pro_receta.descripcion AS recdesc, pro_tareas.descripcion AS tardesc, mae_unidades.abrev, pro_recetatar.cantidad "
            N_GROUP_BY = " pro_receta.id, pro_recetatar.idtar, pro_receta.descripcion, pro_tareas.descripcion, pro_recetatar.orden, mae_unidades.abrev, pro_recetatar.cantidad "
            N_ORDER_BY = " pro_receta.descripcion, pro_tareas.descripcion "
        
        Case 8, 9       ' 8 = TAREA TODO PROD DIA ACTUAL        9 = TAREA TODO PROD RESUMEN
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 3:        Q_POSICION_TOTAL = 3:        Q_COL_COMPARAR_GRUPO = -1
            N_CAMPOS = " pro_recetatar.idtar, pro_tareas.descripcion AS tardesc, mae_unidades.abrev "
            N_GROUP_BY = " pro_recetatar.idtar, pro_tareas.descripcion, pro_recetatar.orden, mae_unidades.abrev "
            N_ORDER_BY = " pro_recetatar.orden "
        
        Case 10, 11, 12 ' 10::EQUIPO X PRODUCTO TODA PROGRAMACION
                        ' 11::EQUIPO X PRODUCTO DIA ACTUAL
                        ' 12::EQUIPO TODO PROD TODA PROGRAMACION
        
        Case 13, 14     ' 13::EQUIPO TODO PROD DIA ACTUAL
                        ' 14::EQUIPO TODO PROD RESUMEN
    End Select
    
    Select Case TIPO_VENTANA
        Case 0 ' INSUMO
            T_RPT_TITULO = "REPORTE DE INSUMOS"
            N_FROM = " mae_tipoproducto INNER JOIN ((pro_receta INNER JOIN pro_programadet ON pro_receta.id = pro_programadet.idrec) INNER JOIN (mae_unidades INNER JOIN (alm_inventario INNER JOIN pro_recetains ON alm_inventario.id = pro_recetains.iditem) ON mae_unidades.id = pro_recetains.idunimed) ON pro_receta.id = pro_recetains.idrec) ON mae_tipoproducto.id = alm_inventario.tippro "
            N_VALOR = " Sum(pro_programadet.canpro*pro_recetains.canpro) "
        
        Case 1 ' TAREA
            T_RPT_TITULO = "REPORTE DE TAREAS"
            N_FROM = " pro_tareas INNER JOIN ((pro_receta INNER JOIN pro_programadet ON pro_receta.id = pro_programadet.idrec) INNER JOIN (mae_unidades INNER JOIN pro_recetatar ON mae_unidades.id = pro_recetatar.idunimed) ON pro_receta.id = pro_recetatar.idrec) ON pro_tareas.id = pro_recetatar.idtar "
            N_VALOR = " Sum(pro_programadet.canpro*pro_recetatar.cantidad) "
        
        Case 2 ' EQUIPO
    End Select
        
    ' DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    Q_POS_MES_INICIO = Q_COL_FILA
    If Q_COL_FILA_ULTIMO <> -1 Then Q_COL_FILA = Q_COL_FILA + Q_COL_FILA_ULTIMO
    
    N_PIVOT = " Format(pro_programadet.dia,'dd/mm/yy') "
    N_PIVOT_SALIDA = ""
    
    ' DEL PIVOT
    For k = 0 To UBound(ARR_TMP()) - 1       ' MENOS TOTAL
        N_PIVOT_SALIDA = N_PIVOT_SALIDA + ARR_TMP(k, 1) + ","
    Next k
    
    N_PIVOT_SALIDA = " IN (" + Left(N_PIVOT_SALIDA, Len(N_PIVOT_SALIDA) - 1) + ") "
    N_WHERE = N_WHERE + " AND " + N_PIVOT + N_PIVOT_SALIDA
    
    ' GENERANDO LA CONSULTA
    vStrSelect = " TRANSFORM " + N_VALOR + _
        vbCr + " SELECT " + N_CAMPOS + "," + N_VALOR + " AS total " + N_CAMPOS_ULTIMO + _
        vbCr + " FROM " + N_FROM + _
        vbCr + " WHERE " + N_WHERE + _
        vbCr + " GROUP BY " + N_GROUP_BY + _
        vbCr + " ORDER BY " + N_ORDER_BY + _
        vbCr + " PIVOT " + N_PIVOT + N_PIVOT_SALIDA
    GENERAR_CONSULTA = vStrSelect
End Function

Private Sub Limpiar_ARRAY_TOTAL(Optional F_LIMPIA_TOT_GRL As Boolean = False)
    Dim k As Integer
    
    For k = 0 To UBound(ARR_TMP())
        ARR_TMP(k, 3) = 0
        If F_LIMPIA_TOT_GRL = True Then ARR_TMP(k, 4) = 0
    Next
    
    ARR_TMP_1(0, 0) = 0 ' STOCK
    ARR_TMP_1(1, 0) = 0 ' SALDO
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosGridAddTotales
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA LOS TOTALES POR CADA GRUPO Y EL TOTAL GENERAL, ACUMULA LOS TOTALES EN EL
'*                    TOTAL GENERAL
'* Paranetros       : NOMBRE          |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    BAND_ADD_TOTAL  |  Boolean     |
'*                    Nombre_total    |  String      |
'*                    Band_Total_gral |  Boolean     |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosGridAddTotales(BAND_ADD_TOTAL As Boolean, _
                                            Nombre_total As String, _
                                            Optional Band_Total_gral As Boolean = False)
    Dim Q_MES As Integer
    Dim X_ROW As Long
    
    X_ROW = Fg1.Rows
    If BAND_ADD_TOTAL = True Then
        ' AGREAGNDO NUEVA FILA
        ADD_REG Fg1, IIf(Band_Total_gral = False, Fila_Total, Fila_Total_grl)

        ' PONIENDO LOS NOMBRES DE LOS TOTALES  Q_POSICION_TOTAL SE OBTIENE DE GENERAR_CONSULTA()
        Fg1.TextMatrix(X_ROW, Q_POSICION_TOTAL) = Nombre_total
        FORMATO_CELDA Fg1, X_ROW, Q_POSICION_TOTAL
    End If
    
    ' ACUMULANDO LOS TOTALES GRLES
    If Band_Total_gral = False Then
        For Q_MES = 0 To Q_COL_ARR_TOTAL
            ARR_TMP(Q_MES, 4) = NulosN(ARR_TMP(Q_MES, 4)) + NulosN(ARR_TMP(Q_MES, 3))
        Next Q_MES
        If Q_COL_FILA_ULTIMO <> -1 Then
            ARR_TMP_1(0, 1) = NulosN(ARR_TMP_1(0, 1)) + NulosN(ARR_TMP_1(0, 0)) '--STOCK
            ARR_TMP_1(1, 1) = NulosN(ARR_TMP_1(1, 1)) + NulosN(ARR_TMP_1(1, 0)) '--SALDO
        End If
    End If
    
    Dim Q_INCREMENTO_X_COL As Integer   ' SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer     ' POSICIONA A LA COLUMNA DEL TOTAL X AÑO
    
    ' IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
    Q_INCREMENTO_X_COL = 0
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    
    For Q_MES = 0 To Q_COL_ARR_TOTAL
        ' INTERRUMPIR EL PROCESO
        If BAND_INTERRUMPIR = True Then Exit Sub
        Fg1.TextMatrix(X_ROW, Q_POS_MES) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP(Q_MES, 3), ARR_TMP(Q_MES, 4)), Band_Total_gral, Q_MES)
        FORMATO_CELDA Fg1, X_ROW, Q_POS_MES
        Q_POS_MES = Q_POS_MES + 1
    Next Q_MES
    
    If Q_COL_FILA_ULTIMO <> -1 Then
        ' STOCK
        Fg1.TextMatrix(X_ROW, Fg1.Cols - 2) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP_1(0, 0), ARR_TMP_1(0, 1)), Band_Total_gral, Fg1.Cols - 2)
        FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 2, RGB(128, 0, 0)
        ' SALDO
        Fg1.TextMatrix(X_ROW, Fg1.Cols - 1) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP_1(1, 0), ARR_TMP_1(1, 1)), Band_Total_gral, Fg1.Cols - 1)
        FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 1, vbRed
    End If
    Err.Clear
End Sub

'*****************************************************************************************************
'* Nombre           : Configurar_Grilla
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA DE ACUERDO A LO QUE SE SELECCIONA
'* Paranetros       : NOMBRE               |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    F_CONSERVAR_FORMATO  |  Boolean   |
'* Devuelve         :
'*****************************************************************************************************
Private Sub Configurar_Grilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    Dim M_ANCHO_COL_MES As Integer      ' DEPENDERA DEL TIPO DE PRESENTACION EN DECIMALES, EN MILES
    Dim k, j As Integer
    
    If F_CONSERVAR_FORMATO = True Then Fg1.Clear
    Fg1.FrozenCols = 0
    M_ANCHO_COL_MES = 900
    
    With Fg1
        Fg1.Cols = Q_COL_FILA + (Q_COL_ARR_TOTAL + 1)
        Q_POS_MES = Q_POS_MES_INICIO
        ' DATOS DE COLUMNAS
        For k = 0 To Q_COL_ARR_TOTAL              ' TODOS LOS DIAS + EL TOTAL
            .ColAlignment(Q_POS_MES) = flexAlignRightCenter
            ' COLOCANDO EL TOTAL
            If k = Q_COL_ARR_TOTAL Then
                .TextMatrix(0, Q_POS_MES) = ARR_TMP(k, 0): .ColWidth(Q_POS_MES) = M_ANCHO_COL_MES + 200
            Else                                  ' COLOCANDO LOS DEMAS DIAS
                .TextMatrix(0, Q_POS_MES) = ARR_TMP(k, 0): .ColWidth(Q_POS_MES) = M_ANCHO_COL_MES
            End If
            Q_POS_MES = Q_POS_MES + 1
        Next k
        
        .FrozenCols = Q_POS_MES_INICIO - 1
        .ColWidth(0) = 200
        
        ' DATOS DE FILA
        Select Case ESTILO_VISTA      ' PARAMETRO
            Case 0, 1, 2  '0 = INSUMO X PRODUCTO TODA PROGRAMACION    1 = INSUMO X PRODUCTO DIA ACTUAL     2 = INSUMO TODO PROD TODA PROGRAMACION
                .TextMatrix(0, 3) = "Receta":           .ColWidth(3) = 2000:    .ColAlignment(3) = flexAlignLeftCenter
                .TextMatrix(0, 4) = "Tipo Producto":    .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftCenter
                .TextMatrix(0, 5) = "Insumo":           .ColWidth(5) = 2500:    .ColAlignment(5) = flexAlignLeftCenter
                .TextMatrix(0, 6) = "U.M.":             .ColWidth(6) = 500:     .ColAlignment(6) = flexAlignLeftCenter
                .TextMatrix(0, 7) = "Unid.":            .ColWidth(7) = 1000:    .ColAlignment(7) = flexAlignRightBottom
                 If Q_COL_FILA_ULTIMO <> -1 Then
                    .TextMatrix(0, Fg1.Cols - 2) = "Stock Act.":    .ColWidth(Fg1.Cols - 2) = 1000:   .ColAlignment(Fg1.Cols - 2) = flexAlignRightCenter
                    .TextMatrix(0, Fg1.Cols - 1) = "Saldo":         .ColWidth(Fg1.Cols - 1) = 1000:   .ColAlignment(Fg1.Cols - 1) = flexAlignRightCenter
                End If
                
            Case 3, 4     ' 3 = INSUMO TODO PROD DIA ACTUAL      4 = INSUMO LOS PRODUCTOS RESUMEN
                .TextMatrix(0, 2) = "Tipo Producto":    .ColWidth(2) = 1200:     .ColAlignment(2) = flexAlignLeftCenter
                .TextMatrix(0, 3) = "Insumo":           .ColWidth(3) = 2200:     .ColAlignment(3) = flexAlignLeftCenter
                .TextMatrix(0, 4) = "U.M.":             .ColWidth(4) = 500:      .ColAlignment(4) = flexAlignLeftCenter
                If Q_COL_FILA_ULTIMO <> -1 Then
                    .TextMatrix(0, Fg1.Cols - 2) = "Stock Act.":    .ColWidth(Fg1.Cols - 2) = 1000:   .ColAlignment(Fg1.Cols - 2) = flexAlignRightCenter
                    .TextMatrix(0, Fg1.Cols - 1) = "Saldo":         .ColWidth(Fg1.Cols - 1) = 1000:   .ColAlignment(Fg1.Cols - 1) = flexAlignRightCenter
                End If
                If ESTILO_VISTA = 4 Then .ColWidth(3) = 4000
            
            Case 5, 6, 7  ' 4 = TAREA X PRODUCTO TODA PROGRAMACION    5 = TAREA X PRODUCTO DIA ACTUAL    6 = TAREA TODO PROD TODA PROGRAMACION
                .TextMatrix(0, 3) = "Receta":      .ColWidth(3) = 2000:     .ColAlignment(3) = flexAlignLeftCenter
                .TextMatrix(0, 4) = "Tarea":      .ColWidth(4) = 2500:      .ColAlignment(4) = flexAlignLeftCenter
                .TextMatrix(0, 5) = "U.M.":        .ColWidth(5) = 500
                .TextMatrix(0, 6) = "Unid.":        .ColWidth(6) = 1000:    .ColAlignment(6) = flexAlignRightBottom
                
            Case 8, 9     ' 8 = TAREA TODO PROD DIA ACTUAL    9 = TAREA TODO PROD RESUMEN
                .TextMatrix(0, 2) = "Tarea":      .ColWidth(2) = 2500:      .ColAlignment(2) = flexAlignLeftCenter
                .TextMatrix(0, 3) = "U.M.":       .ColWidth(3) = 500:       .ColAlignment(3) = flexAlignLeftCenter
                If ESTILO_VISTA = 9 Then .ColWidth(2) = 4000
            
            Case 10, 11, 12 ' 10 = EQUIPO X PRODUCTO TODA PROGRAMACION    11 = EQUIPO X PRODUCTO DIA ACTUAL    12 = EQUIPO TODO PROD TODA PROGRAMACION
                .TextMatrix(0, 3) = "Receta":      .ColWidth(3) = 2000:     .ColAlignment(3) = flexAlignLeftCenter
                .TextMatrix(0, 4) = "Equipo":      .ColWidth(4) = 2500:     .ColAlignment(4) = flexAlignLeftCenter
                .TextMatrix(0, 5) = "U.M.":        .ColWidth(5) = 500:      .ColAlignment(5) = flexAlignLeftCenter
                .TextMatrix(0, 6) = "Unid.":        .ColWidth(6) = 1000:    .ColAlignment(6) = flexAlignRightBottom
            
            Case 13, 14   ' 13 = EQUIPO TODO PROD DIA ACTUAL     14 = EQUIPO TODO PROD RESUMEN
                .TextMatrix(0, 2) = "Equipo":      .ColWidth(2) = 2500:     .ColAlignment(2) = flexAlignLeftCenter
                .TextMatrix(0, 3) = "U.M.":        .ColWidth(3) = 500:      .ColAlignment(3) = flexAlignLeftCenter
        End Select

        If Q_COL_COMPARAR_GRUPO <> -1 Then .ColWidth(3) = 0
        
        ' DE LOS ID'S
        For k = 1 To Q_COL_FILA_OCULTA
            .TextMatrix(0, k) = "ID" + CStr(k):         .ColWidth(k) = 500
        Next
        
        If Q_COL_FILA_OCULTA <> -1 Then OCULTAR_COL Fg1, 1, Q_COL_FILA_OCULTA
    End With
    DoEvents
End Sub

'*****************************************************************************************************
'* Nombre           : PONER_FORMATO
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ESTA FUNCION CONVERTIRA AL FORMATO
'* Paranetros       : NOMBRE          |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    S_MONTO         |  Double     |
'*                    Band_Total_gral |  Boolean    |
'*                    Q_POS           |  Integer    |
'* Devuelve         :
'*****************************************************************************************************
Private Function PONER_FORMATO(S_MONTO As Double, _
                        Optional Band_Total_gral As Boolean = False, _
                        Optional Q_POS As Integer = -1) As String
    If S_MONTO = 0 Then
            PONER_FORMATO = "0.00"
        Exit Function
    End If
    
    PONER_FORMATO = Format(S_MONTO, FORMAT_MONTO)
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    
    If Button.Index = 3 Then pExportar
    
    If Button.Index = 4 Then pImprimir
    
    If Button.Index = 6 Then
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pExportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Programa de Producción", lbl(2).Caption, "Responsable: " & lbl(1).Caption, "Programa de Producción"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub

