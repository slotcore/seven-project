VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmConsProgramarPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja y Bancos - Consulta Programar Pagos"
   ClientHeight    =   7665
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11910
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   3375
      TabIndex        =   6
      Top             =   2910
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   7
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
         TabIndex        =   10
         Top             =   150
         Width           =   1530
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Consulta"
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
         TabIndex        =   8
         Top             =   150
         Width           =   1770
      End
      Begin VB.Shape Shape1 
         Height          =   750
         Left            =   90
         Top             =   60
         Width           =   5805
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1185
      Left            =   30
      TabIndex        =   0
      Top             =   375
      Width           =   11835
      Begin VB.Frame Frame6 
         Caption         =   "[ Seleccione la Moneda ]"
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
         Height          =   870
         Left            =   7290
         TabIndex        =   18
         Top             =   240
         Width           =   2565
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   540
            Picture         =   "FrmConsProgramarPagos.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   495
            Width           =   240
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   105
            MaxLength       =   1
            TabIndex        =   20
            Text            =   "TxtIdMon"
            Top             =   465
            Width           =   705
         End
         Begin VB.Label LblMoneda 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblMoneda"
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
            Left            =   810
            TabIndex        =   22
            Top             =   480
            Width           =   1635
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   21
            Top             =   255
            Width           =   585
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "[ Fecha de Pago ]"
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
         Height          =   870
         Left            =   2100
         TabIndex        =   12
         Top             =   240
         Width           =   3225
         Begin AspaTextBoxFecha.TextBoxFecha TxtFec1 
            Height          =   300
            Left            =   150
            TabIndex        =   13
            Top             =   465
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
            Left            =   1695
            TabIndex        =   14
            Top             =   465
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
            Left            =   1695
            TabIndex        =   16
            Top             =   255
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   255
            Width           =   510
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "[ Tipo de Consulta ]"
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
         Height          =   870
         Left            =   90
         TabIndex        =   5
         Top             =   240
         Width           =   1935
         Begin VB.OptionButton opt_consulta 
            Caption         =   "&Detallado"
            Height          =   195
            Index           =   1
            Left            =   255
            TabIndex        =   2
            Top             =   555
            Width           =   1065
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "&Resumen"
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   1
            Top             =   285
            Value           =   -1  'True
            Width           =   1185
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "[ Agrupar Por ]"
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
         Height          =   870
         Left            =   5385
         TabIndex        =   11
         Top             =   240
         Width           =   1860
         Begin VB.OptionButton opt_tipo 
            Caption         =   "x N° &Orden Pago"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   3
            Top             =   285
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.OptionButton opt_tipo 
            Caption         =   "x &Proveedor"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   555
            Width           =   1455
         End
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Align           =   2  'Align Bottom
      Height          =   6075
      Left            =   0
      TabIndex        =   9
      Top             =   1590
      Width           =   11910
      _cx             =   21008
      _cy             =   10716
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
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsProgramarPagos.frx":0132
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
      Height          =   345
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a Excel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3285
         Top             =   0
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
               Picture         =   "FrmConsProgramarPagos.frx":0343
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":0887
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":0C19
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":0D73
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":1105
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":1289
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":16DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":17F5
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":1D39
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":227D
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":2391
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":24A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":28F9
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProgramarPagos.frx":2A65
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmConsProgramarPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-- ALMACENAR LOS TOTALES DE TODA LA CONSULTA
'--ARR_TMP(?,4)= Arr_Totales_cols() As Double '--ALMACENAR TOTALES POR TODAS LAS FILAS
'--ARR_TMP(?,3)= Arr_Totales_col() As Double     '--ALMACENAR TOTALES POR COLUMNA, SE LIMPIA DESPUES DE CAMBIO DE GRUPO


Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
'--DE LA IMPRESION
Dim T_RPT_PERIODO As String '--PERIODO DEL REPORTE
Dim T_RPT_TITULO As String  '--TITULO DE REPORTE
'------------
Dim ARR_ANYO() As String    '--ARRAY DE AÑOS SELECCIONADOS
Dim ARR_XX() As String      '--SE CARGARA CUANDO SE CARGA EL FORMULARIO Y CUANDO SE CAMBIE EL ESTILO(MES, TRIMESTRE,SEMESTRE)

Dim ARR_TMP(1, 1) As Double '--0::TOTAL IMP=>> 0::TOTAL,1::TOTAL GEN
                            '--1::TOTAL ACUENTA=>> 0::TOTAL,1::TOTAL GEN

Dim Q_TOTAL_ANYO As Integer '--INDICA LA CANTIDAD DE AÑOS DE BUSQUEDA,
                            '--EJ. 2004,2005 => Q_TOTAL_ANYO = 2
                            '--EJ. 2004,2005,2006 => Q_TOTAL_ANYO = 3
                            
Dim Q_COL_FILA As Integer   '--INDICA LA CANTIDAD DE COLUMNAS ANTES DE LOS MESES
                            '--EJ. IDCLI,CLIENTE => Q_COL_FILA=2
                            '--    IDCLI,ID_PTO_VTA,CLIENTE,PTO_VENTA => Q_COL_FILA=4
                            
                            
Dim Q_COL_FILA_ULTIMO As Integer '--INDICA LA CANTIDAD DE COLUMNAS ADICIONALES QUE SE COLOCARAN DESPUES DEL TOTAL
                            
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
                                 '--OBTENDRA VALOR EN GENERAR_CONSULTA()

Dim Q_COL_COMPARAR_GRUPO As Integer '--INDICA LA COLUMNA PARA COMPARAR EL GRUPO
                                    '--OBTENDRA VALOR EN GENERAR_CONSULTA()

Dim Q_COL_GRUPO_ADD As Integer  '--ADICIONAR DATOS AL GRID EN EL GRUPO (EJ. Q_COL_GRUPO_ADD=2 =>> NOMBRE_GRUPO|COLUM1|COLUM2)
                                '--FNUCIONA SI Q_COL_GRUPO_ADD<>-1
                                
Dim Q_COL_GRUPO_TERMINA     As Integer  '--INDICA EL TERMINO DEL GRUPO, UNE LAS CELDAS DE 1 HASTA Q_COL_GRUPO_TERMINA
'------------

Dim Q_COL_ARR_TOTAL As Integer  '--NOS INDICA EL TOTAL DE COLUMNAS QUE VA A CONTENER LOS ACUMULADOS
                                '--OBTENDRA VALOR EN VALIDAR_CONSULTA()
                                '--SI SEL MES: ENE, FEB, MAR => Q_COL_ARR_TOTAL= 2
                                '--SI SEL TRI: ENE-MAR, ABR-JUN => Q_COL_ARR_TOTAL= 1 OBS: SE INICIA DESDE POS=0

Dim F_ES_COMPRA As Boolean '--INDICA SI ES COMPRA O VENTA
                            '--TRUE::ES COMPRA, FALSE::ES VENTA

Dim ID_PROGRAMA As String
Dim ID_RECETA As String
Dim ESTILO_VISTA As Integer
'-------
Dim N_VALOR_FONDO           As String '--AMACENA EL VALOR PARA COMPARAR
Dim N_VALOR_FONDO_COLOR     As Long '--AMACENA EL VALOR DEL COLOR PARA EL FONDO DE LA FILA
Dim F_CAMIAR_FONDO          As Boolean  '--FALSE::SE CONSERVA EL FONDO ACTUAL, TRUE::CAMBIA DE FONDO
Dim Q_COL_COMPARAR_FONDO    As Integer  '--INDICA LA COLUMNA DEL RECORDSET QUE DEBERA DE COMPARAR PARA CAMBIAR DE FONDO
                                        '-- -1=NO HACER NADA

'------------


Private Sub CONSULTAR()
'    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    '--
    Dim vStrSelect As String '--RECIBIR LA CONSULTA
    
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
    If vStrSelect = "" Then GoTo Salir
    PosicionarProgBar
    DoEvents
    '--CARGADO EL RST
    RST_Busq rst_select, vStrSelect, xCon
   '--------------------------------------
    CARGAR_DATOS_GRILLA rst_select
   '--------------------------------------
   '
Salir:
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

Private Sub CARGAR_DATOS_GRILLA_FONDO(RST_ORIGEN As ADODB.Recordset, _
                                        X_ROW1 As Long, X_COL1 As Integer, _
                                        X_ROW2 As Long, X_COL2 As Integer)
    ''--PONER COLOR FONDO
    If Q_COL_COMPARAR_FONDO = -1 Then Exit Sub
        If IsNumeric(Fg1.TextMatrix(X_ROW1, 1)) = False Then Exit Sub
        If Fg1.TextMatrix(X_ROW1, 1) = e_ESTADO_ROW_GRID.Fila_grupo Then
            '--SI SE DESEA PONER COLOR AL GRUPO
            'GRID_COLOR_FONDO Fg1, X_ROW1, X_COL1, X_ROW2, X_COL2, RGB(0, 185, 185)
        ElseIf Fg1.TextMatrix(X_ROW1, 1) = e_ESTADO_ROW_GRID.Fila_Total Then
        ElseIf Fg1.TextMatrix(X_ROW1, 1) = e_ESTADO_ROW_GRID.Fila_Total_grl Then
        ElseIf Fg1.TextMatrix(X_ROW1, 1) = e_ESTADO_ROW_GRID.Fila_en_Blanco Then
        Else
           If RST_ORIGEN.Bookmark = 1 Then
                N_VALOR_FONDO = RST_ORIGEN.Fields(Q_COL_COMPARAR_FONDO) & ""
                N_VALOR_FONDO_COLOR = &HE0FEFE
                F_CAMIAR_FONDO = False
            End If
    
            If N_VALOR_FONDO = RST_ORIGEN.Fields(Q_COL_COMPARAR_FONDO) Then
                N_VALOR_FONDO_COLOR = N_VALOR_FONDO_COLOR
            Else
                N_VALOR_FONDO = RST_ORIGEN.Fields(Q_COL_COMPARAR_FONDO)
                If F_CAMIAR_FONDO = True Then
                    N_VALOR_FONDO_COLOR = &HE0FEFE
                    F_CAMIAR_FONDO = False
                Else
                    N_VALOR_FONDO_COLOR = &HFDFFFF
                    F_CAMIAR_FONDO = True
                End If
            End If
            GRID_COLOR_FONDO Fg1, X_ROW1, X_COL1, X_ROW2, X_COL2, N_VALOR_FONDO_COLOR
        End If
    
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
        ''--PONER COLOR FONDO
        If Q_COL_COMPARAR_FONDO <> -1 Then CARGAR_DATOS_GRILLA_FONDO RST_ORIGEN, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            
        '---------------------------------------------------------
        '---------------------------------------------------------
        RST_ORIGEN.MoveNext
'        --PONER TOTALES AL FINAL DE LA GRILLA
        
        If RST_ORIGEN.EOF Then
            CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:"
            Select Case ESTILO_VISTA
            Case 1, 2
                CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Tot Gen:", True
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
        If RST_ORIGEN.Bookmark = 1 Then ADD_REG Fg1, Fila_Ninguno
        GoTo Salir
    End If
    '---------------------------------------------------------
    If Q_COL_COMPARAR = -1 Then Q_COL_COMPARAR = 1
    
    
    If Q_COL_GRUPO_ADD <> -1 Then
        For Q_POS = 1 To Q_COL_GRUPO_ADD
            N_GRUPO_ADD = RST_ORIGEN.Fields(Q_COL_COMPARAR + Q_POS) & "  " + N_GRUPO_ADD
        Next Q_POS
        N_GRUPO_ADD = "   " + N_GRUPO_ADD
    End If
    
    If RST_ORIGEN.Bookmark = 1 Then
        '--SE CARGA EN GENERAR_CONSULTA() Q_COL_COMPARAR_GRUPO
        ADD_REG Fg1, Fila_grupo
        UNIR_CELDAS Fg1, Fg1.Rows - 1, 2, Fg1.Rows - 1, Q_COL_GRUPO_TERMINA, INICIO_GRUPO + RST_ORIGEN.Fields(Q_COL_COMPARAR) + N_GRUPO_ADD, flexAlignLeftCenter
        'Fg1.MergeCells = flexMergeRestrictRows
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 2
        If Q_COL_COMPARAR_FONDO <> -1 Then CARGAR_DATOS_GRILLA_FONDO RST_ORIGEN, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
        
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

            ADD_REG Fg1, Fila_grupo
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 2, Fg1.Rows - 1, Q_COL_GRUPO_TERMINA, INICIO_GRUPO + RST_ORIGEN.Fields(Q_COL_COMPARAR) + N_GRUPO_ADD, flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 2
            
            If Q_COL_COMPARAR_FONDO <> -1 Then CARGAR_DATOS_GRILLA_FONDO RST_ORIGEN, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            
        End If
    End If

    
    
Salir:
    Set RST_TEPM_1 = Nothing
End Sub

Private Sub CARGAR_DATOS_ARRAY(RST_ORIGEN As ADODB.Recordset)
    '--FUNCION QUE ACUMULARA EN EL ARRAY_TEMP
    Dim nCampo As String
    Dim mCampo&
    '--ASIGNAR LOS DATOS AL RECORDSET TEMPORAL
    For mCampo = 0 To RST_ORIGEN.Fields.Count - 1
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then Exit Sub
        nCampo = RST_ORIGEN.Fields(mCampo).Name
        '--OBS: SE VA LLENAR EL ARRAY "TOTAL"
        
        If LCase(nCampo) = "totimp" Then
            ARR_TMP(0, 0) = ARR_TMP(0, 0) + NulosN(RST_ORIGEN.Fields(nCampo))
        ElseIf LCase(nCampo) = "totacuenta" Then
            ARR_TMP(1, 0) = ARR_TMP(1, 0) + NulosN(RST_ORIGEN.Fields(nCampo))
        End If
        
    Next mCampo
    
End Sub

Private Function CARGAR_DATOS_GRILLA_ARRAY_TMP(RST_ORIGEN As ADODB.Recordset, _
                                         Q_ROW As Long)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim Q_POS As Integer
    Dim mCampo As Integer
    Dim nCampo As String
    
    '-----------
    DoEvents
    For mCampo = 0 To RST_ORIGEN.Fields.Count - 1
        If BAND_INTERRUMPIR = True Then Exit Function
        nCampo = RST_ORIGEN.Fields(mCampo).Name
        Select Case LCase(nCampo)
            Case "totimp"
                Fg1.TextMatrix(Q_ROW, Fg1.Cols - 2) = Format(NulosN(RST_ORIGEN.Fields(nCampo)), FORMAT_MONTO)
            Case "totacuenta"
                Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1) = Format(NulosN(RST_ORIGEN.Fields(nCampo)), FORMAT_MONTO)
            Case "fchemi", "fchdoc", "fchven"
                Fg1.TextMatrix(Q_ROW, mCampo + 1) = Format(NulosC(RST_ORIGEN.Fields(nCampo)), FORMAT_DATE)
            Case Else
                '--AGREGAR LOS DEMAS DATOS
                Fg1.TextMatrix(Q_ROW, mCampo + 1) = RST_ORIGEN.Fields(nCampo) & ""
        End Select
    Next
End Function


Private Sub pImprimir()

    On Error GoTo error
    Dim oPrint As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    If F_ES_COMPRA = False Then T_RPT_TITULO = Replace(T_RPT_TITULO, "COMPRA", "VENTA")
    oPrint.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO + " ", "", T_RPT_PERIODO, False, True
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub


Private Sub Fg1_DblClick()
    Fg1_KeyDown 13, 0
End Sub

Private Sub Fg1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> 13 Then Exit Sub
'    If Fg1.Rows = 1 Then Exit Sub
'    If Fg1.Row = 0 Or Fg1.Row = Fg1.Rows - 1 Or Fg1.Col = 0 Or Fg1.Col = 1 Or Fg1.Col = 2 Or Fg1.Col = Fg1.Cols - 1 Then
'        MsgBox "Selecione una Celda Correcta..", vbInformation, "Mensaje"
'        Exit Sub
'    End If
'    If txt(5).Text = "" Or IsNumeric(txt(5).Text) = False Then
'        MsgBox "Ingrese un número a mostrar", vbInformation, "Mensaje..."
'        txt(5).SetFocus
'        Exit Sub
'    End If
'    If IsNumeric(Fg1.TextMatrix(Fg1.Row, Fg1.Col)) = False Then
'        MsgBox "La celda no es numérico", vbInformation, "Mensaje..."
'        Exit Sub
'    End If
    
'    With FrmAnalizaPrecio_Item
'        .RECIBE_ID_ITEM Fg1.TextMatrix(Fg1.Row, 1), Fg1.TextMatrix(1, Fg1.Col), ARR_TMP(), F_ES_COMPRA
'        .Show 1
'    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo error
    CentrarFrm Me
    TxtIdMon.Text = ""
    TxtFec1.Valor = CDate("01/01/" + CStr(Year(Date)))
    TxtFec2.Valor = Date
    GENERAR_CONSULTA
    Configurar_Grilla
    TxtIdMon.Text = 1
    TxtIdMon_Validate False
    Exit Sub
error:
    SHOW_ERROR
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    BAND_INTERRUMPIR = True
    Erase ARR_TMP
End Sub

'------
Private Function Validar_Consulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    '--DE LA FECHA ES NULL
    If TxtFec1.Valor = "" Or TxtFec2.Valor = "" Then
        MsgBox "Ingrese una fecha", vbExclamation, xTitulo
        If TxtFec1.Valor = "" Then TxtFec1.SetFocus Else TxtFec2.SetFocus
        Exit Function
    End If
    If CDate(TxtFec1.Valor) > CDate(TxtFec2.Valor) Then
        MsgBox "La fecha inicial es superior al Final", vbExclamation, xTitulo
        TxtFec1.SetFocus
        Exit Function
    End If
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "La especificar la Moneda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    Validar_Consulta = True
End Function

Private Function GENERAR_CONSULTA() As String
    '--FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    '--
    Dim vStrSelect As String            '--CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    
    Dim nSQL As String
    Dim nSQL_1 As String      '--ESTE FILTRO SERVIRA PARA CONSULTAR EN EL SUB_SELECT
    
    Dim k&
    
    Dim SQL_PROD As String
    Dim SQL_INSUMO As String
    Dim T_CONSULTA As Integer '--DEL TIPO DE CONSULTA, SE FORMARA EL ENCABEZADO DEL GRID
    
        '--DE LA FECHA
    If CDate(TxtFec1.Valor) < CDate(TxtFec2.Valor) Then
        nSQL = " ( con_ordenpago.fchpag >=CDATE ('" + TxtFec1.Valor + "') AND con_ordenpago.fchpag <= CDATE('" + TxtFec2.Valor + "') ) "
        T_RPT_PERIODO = " Del: " + CStr(TxtFec1.Valor) + " Al: " + CStr(TxtFec2.Valor)
    Else
        nSQL = " con_ordenpago.fchpag = CDATE('" + TxtFec1.Valor + "') "
         T_RPT_PERIODO = "Al: " + CStr(TxtFec2.Valor)
    End If
    
    nSQL = nSQL + " AND con_ordenpago.idmon = " & NulosN(TxtIdMon.Text) & " AND con_ordenpagodet.aprobado=-1 "
    
    '--GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim N_VALOR As String
    Dim N_CAMPOS As String
    Dim N_WHERE As String
    Dim N_FROM As String
    Dim N_GROUP_BY As String
    Dim N_ORDER_BY As String
    
    N_WHERE = nSQL
    
    pEstiloConsulta
    Q_COL_COMPARAR_FONDO = -1
    Select Case ESTILO_VISTA
        Case 0 '--RESUMIDO
        
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 6:        Q_POSICION_TOTAL = 4:        Q_COL_COMPARAR_GRUPO = -1
            Q_COL_GRUPO_ADD = -1 '--ADICIONAR DATOS AL GRID EN EL GRUPO (NOMBRE_GRUPO|COLUM1|COLUM2)
            Q_COL_GRUPO_TERMINA = -1
            T_RPT_TITULO = "RESUMEN DE PROGRAMACIÓN DE PAGOS"
            N_CAMPOS = " mae_prov.id, mae_prov.numruc, mae_prov.nombre, Count(com_compras.id) AS canreg, mae_moneda.simbolo, Sum(com_compras.imptot) AS totimp, Sum(con_ordenpagodet.acuenta) AS totacuenta "
            N_GROUP_BY = " mae_prov.id, mae_prov.numruc, mae_prov.nombre, mae_moneda.simbolo, con_ordenpagodet.aprobado "
            N_ORDER_BY = "  mae_prov.nombre; "
        Case 1 '--DETALLE/ORDEN PAGO
        
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 11:        Q_POSICION_TOTAL = 9:        Q_COL_COMPARAR_GRUPO = 1
            Q_COL_GRUPO_ADD = -1
            Q_COL_GRUPO_TERMINA = 12
            T_RPT_TITULO = "DETALLE DE ORDEN DE PAGO AGRUPADO POR N° ORDEN"
            N_CAMPOS = " con_ordenpago.id, 'N° Orden:   ' & [con_ordenpago].[numdoc] AS numorden, IIf([con_ordenpago].[tipope]=1,'Caja','Banco') AS operacion, Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Mid([com_compras].[numreg],3) AS registro, mae_documento.abrev, [com_compras].[numser] & '-' & [com_compras].[numdoc] AS numerodoc, mae_prov.nombre, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, com_compras.imptot AS totimp, con_ordenpagodet.acuenta AS totacuenta "
            N_GROUP_BY = ""
            N_ORDER_BY = " 'N° Orden:   ' & [con_ordenpago].[numdoc], mae_prov.nombre; "
            
        Case 2 '--DETALLE/PROVEEDOR
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 11:        Q_POSICION_TOTAL = 9:        Q_COL_COMPARAR_GRUPO = 1
            Q_COL_GRUPO_ADD = 1
            Q_COL_GRUPO_TERMINA = 12
            T_RPT_TITULO = "DETALLE DE ORDEN DE PAGO AGRUPADO POR PROVEEDOR"
            N_CAMPOS = " con_ordenpago.id, mae_prov.numruc, mae_prov.nombre, con_ordenpago.numdoc AS numorden, IIf([con_ordenpago].[tipope]=1,'Caja','Banco') AS operacion, Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Mid([com_compras].[numreg],3) AS registro, mae_documento.abrev, [com_compras].[numser] & '-' & [com_compras].[numdoc] AS numerodoc, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, com_compras.imptot AS totimp, con_ordenpagodet.acuenta AS totacuenta "
            N_GROUP_BY = ""
            N_ORDER_BY = " mae_prov.numruc, con_ordenpago.numdoc "
    
    End Select
    
    '--DEL FROM
    Select Case ESTILO_VISTA
        Case 0 '--RESUMEN
            N_FROM = "  (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN com_compras ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) RIGHT JOIN ((con_ordenpago LEFT JOIN con_tc ON con_ordenpago.fchpag = con_tc.fecha) INNER JOIN con_ordenpagodet ON con_ordenpago.id = con_ordenpagodet.idord) ON com_compras.id = con_ordenpagodet.idcom "
        
        Case 1, 2 '--DETALLE
            N_FROM = "  (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) RIGHT JOIN ((con_ordenpago LEFT JOIN con_tc ON con_ordenpago.fchpag = con_tc.fecha) INNER JOIN con_ordenpagodet ON con_ordenpago.id = con_ordenpagodet.idord) ON com_compras.id = con_ordenpagodet.idcom "
        
    End Select
    '--DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    Q_POS_MES_INICIO = Q_COL_FILA '--Q_COL_FILA + CAMPO_TOTAL
    
    '------------------------------------------
    '--GENERANDO LA CONSULTA
    
    vStrSelect = "SELECT " + N_CAMPOS + _
        vbCr + " FROM " + N_FROM + _
        vbCr + " WHERE " + N_WHERE + _
        IIf(N_GROUP_BY <> "", vbCr + " GROUP BY " + N_GROUP_BY, "") + _
        vbCr + " ORDER BY " + N_ORDER_BY

    '-------------------------------------------
    
    GENERAR_CONSULTA = vStrSelect
    
End Function



Private Sub Limpiar_ARRAY_TOTAL(Optional F_LIMPIA_TOT_GRL As Boolean = False)
    Dim k&
    For k = 0 To UBound(ARR_TMP())
        ARR_TMP(k, 0) = 0
        If F_LIMPIA_TOT_GRL = True Then ARR_TMP(k, 1) = 0
    Next
                            
End Sub
'''
Private Sub CARGAR_DATOS_GRILLA_ADD_TOTALES(BAND_ADD_TOTAL As Boolean, _
                                            Nombre_total As String, _
                                            Optional Band_Total_gral As Boolean = False)
                
    Dim Q_MES As Integer
    '--AGREGA LOS TOTALES POR CADA GRUPO Y EL TOTAL GENERAL
    '--ACUMULA LOS TOTALES EN EL TOTAL GENERAL
    Dim X_ROW&
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
        ARR_TMP(0, 1) = NulosN(ARR_TMP(0, 1)) + NulosN(ARR_TMP(0, 0))
        ARR_TMP(1, 1) = NulosN(ARR_TMP(1, 1)) + NulosN(ARR_TMP(1, 0))
    End If
    
    '--------------------------
    '--INTERRUMPIR EL PROCESO
    If BAND_INTERRUMPIR = True Then Exit Sub
    Fg1.TextMatrix(X_ROW, Fg1.Cols - 2) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP(0, 0), ARR_TMP(0, 1)), Band_Total_gral, Fg1.Cols - 2)
    Fg1.TextMatrix(X_ROW, Fg1.Cols - 1) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP(1, 0), ARR_TMP(1, 1)), Band_Total_gral, Fg1.Cols - 1)
    
    FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 2
    FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 1
    Err.Clear
End Sub

Private Sub Configurar_Grilla()
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
    Dim M_ANCHO_COL& '--DEPENDERA DEL TIPO DE CONSULTA
                                   
    Dim k&, j&
    Dim T_CONSULTA&
    
    
    Fg1.FrozenCols = 0
    
    M_ANCHO_COL = 0

    With Fg1
        '-----
    .Cols = Q_COL_FILA_OCULTA + Q_COL_FILA
                 
    Q_POS_MES = Q_POS_MES_INICIO
        
    '.FrozenCols = Q_POS_MES_INICIO - 1
    .ColWidth(0) = 200
    '--DATOS DE FILA
        
    pEstiloConsulta
    Select Case ESTILO_VISTA
        Case 0 '--RESUMIDO / X PRODUCTO
            .TextMatrix(0, 2) = "Ruc":           .ColWidth(2) = 1200:    .ColAlignment(2) = flexAlignCenterCenter:  .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 3) = "Proveedor":     .ColWidth(3) = 4000:    .ColAlignment(3) = flexAlignLeftCenter:    .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
            .TextMatrix(0, 4) = "Cant.Doc.":     .ColWidth(4) = 800:     .ColAlignment(4) = flexAlignRightBottom:   .Row = 0: .Col = 4: .CellAlignment = flexAlignRightBottom
            .TextMatrix(0, 5) = "M":             .ColWidth(5) = 500:     .ColAlignment(5) = flexAlignRightBottom:   .Row = 0: .Col = 5: .CellAlignment = flexAlignRightBottom
        Case 1 '--DETALLE / GRUPO NUM ORDEN
            .TextMatrix(0, 2) = "Num.Orden":     .ColWidth(2) = 0:    .ColAlignment(3) = flexAlignCenterCenter:     .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 3) = "Operación":     .ColWidth(3) = 850:    .ColAlignment(3) = flexAlignLeftBottom:     .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
            .TextMatrix(0, 4) = "Num.Reg":       .ColWidth(4) = 1000:     .ColAlignment(4) = flexAlignCenterCenter: .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 5) = "T.D.":          .ColWidth(5) = 500:     .ColAlignment(5) = flexAlignLeftCenter:    .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
            
            .TextMatrix(0, 6) = "Num.Documento": .ColWidth(6) = 1500:    .ColAlignment(6) = flexAlignCenterCenter:  .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 7) = "Proveedor":     .ColWidth(7) = 3000:    .ColAlignment(7) = flexAlignLeftCenter:    .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(0, 8) = "Fch.Emi.":      .ColWidth(8) = 900:     .ColAlignment(8) = flexAlignCenterCenter:  .Row = 0: .Col = 8: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 9) = "Fch.Venc":      .ColWidth(9) = 900:     .ColAlignment(9) = flexAlignCenterCenter:  .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 10) = "M":            .ColWidth(10) = 450:    .ColAlignment(10) = flexAlignCenterCenter: .Row = 0: .Col = 10: .CellAlignment = flexAlignCenterCenter
                
        Case 2 '--DETALLE / GRUPO PROVEEDOR
            .TextMatrix(0, 2) = "Ruc":          .ColWidth(2) = 0:       .ColAlignment(3) = flexAlignCenterCenter:   .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 3) = "Proveedor":    .ColWidth(3) = 0:       .ColAlignment(3) = flexAlignLeftBottom:     .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
            .TextMatrix(0, 4) = "Num.Orden":    .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftBottom:     .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
            .TextMatrix(0, 5) = "Operación":    .ColWidth(5) = 850:     .ColAlignment(5) = flexAlignLeftCenter:     .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
            
            .TextMatrix(0, 6) = "Num.Reg":      .ColWidth(6) = 1100:    .ColAlignment(6) = flexAlignCenterCenter:   .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 7) = "T.D.":         .ColWidth(7) = 500:     .ColAlignment(7) = flexAlignLeftCenter:     .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(0, 8) = "Num.Documento": .ColWidth(8) = 1600:   .ColAlignment(8) = flexAlignCenterCenter:   .Row = 0: .Col = 7: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 9) = "Fch.Emi.":      .ColWidth(9) = 1000:    .ColAlignment(9) = flexAlignCenterCenter:  .Row = 0: .Col = 8: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 10) = "Fch.Venc":     .ColWidth(10) = 1000:   .ColAlignment(10) = flexAlignCenterCenter: .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 11) = "M":            .ColWidth(11) = 450:   .ColAlignment(11) = flexAlignCenterCenter:  .Row = 0: .Col = 10: .CellAlignment = flexAlignCenterCenter
    
        End Select

        .TextMatrix(0, .Cols - 2) = "Importe":      .ColWidth(.Cols - 2) = 900:     .ColAlignment(.Cols - 2) = flexAlignRightBottom:    .Row = 0: .Col = .Cols - 2: .CellAlignment = flexAlignRightBottom
        .TextMatrix(0, .Cols - 1) = "A Cuenta":     .ColWidth(.Cols - 1) = 900:     .ColAlignment(.Cols - 1) = flexAlignRightBottom:    .Row = 0: .Col = .Cols - 1: .CellAlignment = flexAlignRightBottom
        
        'If Q_COL_COMPARAR_GRUPO <> -1 Then .ColWidth(3) = 0
        
'        --DE LOS ID'S
        For k = 1 To Q_COL_FILA_OCULTA
            .TextMatrix(0, k) = "ID" + CStr(k):         .ColWidth(k) = 500
        Next
        If Q_COL_FILA_OCULTA <> -1 Then OCULTAR_COL Fg1, 1, Q_COL_FILA_OCULTA

        If Q_COL_GRUPO_ADD <> -1 Then OCULTAR_COL Fg1, Q_COL_COMPARAR_GRUPO + 1, Q_COL_COMPARAR_GRUPO + Q_COL_GRUPO_ADD + 1
    
        
    End With
    DoEvents
End Sub


Private Function PONER_FORMATO(S_MONTO As Double, _
                        Optional Band_Total_gral As Boolean = False, _
                        Optional Q_POS As Integer = -1) As String
                        
    '--ESTA FUNCION CONVERTIRA AL FORMATO
    If S_MONTO = 0 Then
            PONER_FORMATO = "0.00"
        Exit Function
    End If
    
    PONER_FORMATO = Format(S_MONTO, FORMAT_MONTO)
    
End Function

Private Function pEstiloConsulta()
    If opt_consulta(0).Value = True Then '--RESUMEN
        ESTILO_VISTA = 0
    Else '--DETALLE
        If opt_tipo(0).Value = True Then ESTILO_VISTA = 1 '--X ORDEN PAGO
        If opt_tipo(1).Value = True Then ESTILO_VISTA = 2 '--X PROVEEDOR
    End If
End Function

Sub PosicionarProgBar()
    '--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
    '    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    Me.PgBar.Value = 1
    FraProgreso.Visible = True
End Sub

'--------
Private Sub EXPORTAR()
On Error GoTo error
    Dim oPrint As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    oPrint.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO + " ", T_RPT_PERIODO, , "Producción"
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

'**********************************

Private Sub opt_consulta_Click(Index As Integer)
    If Index = 0 Then
        opt_tipo(0).Value = False
        opt_tipo(1).Value = False
        habilitar opt_tipo, False
    Else
        opt_tipo(0).Value = True
        habilitar opt_tipo, True
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then CONSULTAR
    If Button.Index = 3 Then pImprimir
    If Button.Index = 4 Then EXPORTAR
    If Button.Index = 6 Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub TxtIdMon_Change()
    If Trim(TxtIdMon.Text) = "" Then LblMoneda.Caption = ""
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosC(TxtIdMon.Text) <> "" Then
        LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
        If NulosC(LblMoneda.Caption) = "" Then
            TxtIdMon.Text = ""
        End If
    End If
End Sub

Private Sub CmdBusMon_Click()
    
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":   xCampos(1, 1) = "id":              xCampos(1, 2) = "500":      xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_moneda ORDER BY descripcion ;", xCampos(), "Buscando Moneda", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    TxtIdMon.Text = NulosN(xRs("id"))
    LblMoneda.Caption = NulosC(xRs("descripcion"))
    
Salir:
    Set xRs = Nothing
End Sub

'**********************************

