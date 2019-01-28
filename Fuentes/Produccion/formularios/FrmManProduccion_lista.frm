VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManProduccion_lista 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5715
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   11700
   Begin VB.Frame fr 
      Height          =   990
      Index           =   5
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   11655
      Begin VB.Frame fr 
         Height          =   810
         Index           =   0
         Left            =   5970
         TabIndex        =   6
         Top             =   105
         Width           =   2445
         Begin VB.CheckBox chk 
            Caption         =   "% D&esvio"
            Height          =   270
            Index           =   4
            Left            =   1290
            TabIndex        =   15
            Top             =   480
            Value           =   1  'Checked
            Width           =   990
         End
         Begin VB.CheckBox chk 
            Caption         =   "&Desvio"
            Height          =   270
            Index           =   3
            Left            =   1290
            TabIndex        =   10
            Top             =   180
            Value           =   1  'Checked
            Width           =   945
         End
         Begin VB.CheckBox chk 
            Caption         =   "&Real"
            Height          =   270
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox chk 
            Caption         =   "&Teórico"
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   180
            Value           =   1  'Checked
            Width           =   840
         End
         Begin VB.CheckBox chk 
            Caption         =   "&Programado"
            Height          =   270
            Index           =   0
            Left            =   540
            TabIndex        =   7
            Top             =   900
            Visible         =   0   'False
            Width           =   1245
         End
      End
      Begin VB.Shape Shape1 
         Height          =   675
         Left            =   8730
         Top             =   225
         Width           =   2520
      End
      Begin VB.Label lbl_color 
         AutoSize        =   -1  'True
         Caption         =   "Consumo Adicional"
         Height          =   195
         Index           =   3
         Left            =   9255
         TabIndex        =   14
         Top             =   660
         Width           =   1350
      End
      Begin VB.Label lbl_color 
         AutoSize        =   -1  'True
         Caption         =   "Consumo Ahorrado"
         Height          =   195
         Index           =   2
         Left            =   9255
         TabIndex        =   13
         Top             =   300
         Width           =   1350
      End
      Begin VB.Label lbl_color 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   9015
         TabIndex        =   12
         Top             =   660
         Width           =   210
      End
      Begin VB.Label lbl_color 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   9015
         TabIndex        =   11
         Top             =   300
         Width           =   210
      End
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
         Left            =   1320
         TabIndex        =   5
         Top             =   555
         Width           =   4530
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
         Left            =   1320
         TabIndex        =   4
         Top             =   195
         Width           =   1785
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Supervizado Por"
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   3
         Top             =   645
         Width           =   1170
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Fch Producción"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   1125
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Align           =   2  'Align Bottom
      Height          =   4365
      Left            =   0
      TabIndex        =   1
      Top             =   1350
      Width           =   11700
      _cx             =   20637
      _cy             =   7699
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
      FormatString    =   $"FrmManProduccion_lista.frx":0000
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
      Height          =   600
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   1058
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
               Picture         =   "FrmManProduccion_lista.frx":003C
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":0580
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":0912
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":0A96
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":0EEA
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":1002
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":1546
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":1A8A
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":1B9E
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":1CB2
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":2106
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":2272
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":27BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManProduccion_lista.frx":2AD4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmManProduccion_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANPRODUCCION_LISTA.FRM
'* Tipo             : FORMULARIO
'* Descripcion      :
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 05/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim BAND_INTERRUMPIR As Boolean     ' SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                    ' TRUE SE INTERRUMPE
' DE LA IMPRESION
Dim T_RPT_PERIODO As String         ' PERIODO DEL REPORTE
Dim T_RPT_TITULO As String          ' TITULO DE REPORTE
Dim ARR_TMP() As String             ' ACUMULARA LOS TOTALES DE PROGRAMADO, PRODUCIDO

                                    ' SE USA PARA DAR FORMATO DE LA GRILLA, SEGUN SELECCIONE EL USUARIO
Dim Q_TOTAL_ANYO As Integer         ' INDICA LA CANTIDAD DE AÑOS DE BUSQUEDA,
                                    ' EJ. 2004,2005 => Q_TOTAL_ANYO = 2
                                    ' EJ. 2004,2005,2006 => Q_TOTAL_ANYO = 3
Dim Q_COL_FILA As Integer           ' INDICA LA CANTIDAD DE COLUMNAS ANTES DE LOS MESES
                                    ' EJ. IDCLI,CLIENTE => Q_COL_FILA=2
                                    ' IDCLI,ID_PTO_VTA,CLIENTE,PTO_VENTA => Q_COL_FILA=4
Dim Q_COL_FILA_ULTIMO As Integer    ' INDICA LA CANTIDAD DE COLUMNAS ADICIONALES QUE SE COLOCARAN DESPUES DEL TOTAL
Dim Q_POS_MES_INICIO As Integer     ' INDICA LA POSICION INICIAL DE LA COLUMNA DEL PRIMER MES, NO CAMBIA
                                    ' EJ. Q_POS_MES_INICIO = Q_COL_FILA +1
Dim Q_POS_MES As Integer            ' INDICA LA POSICION DEL MES, ESTO CAMBIA
                                    ' UTIL PARA COLOCAR LOS DATOS EN EL GRID
Dim Q_COL_FILA_OCULTA As Integer    ' INDICA LAS COLUMNAS QUE CONTENDRAN LOS ID'S, ESTOS SE OCULTARAN
                                    ' -1 NO SE OCULTA, <> -1 SE PROCEDE A ACULTAR
                                    ' EJ. CLIENTE  vta_ventas.idcli,
                                    ' PUNTO DE VENTA vta_guia.idpunven
                                    ' PRODUCTO   alm_inventario.tippro
                                    ' ITEM       alm_inventario.id
                                    ' EMPLEADO   vta_ventas.idven
Dim Q_POSICION_TOTAL  As Integer    ' INDICA LA POCISION DE LA COLUMNA DONDE SE COLOCARA EL NOMBRE DEL TOTAL Y TOTAL_GRL
                                    ' OBTENDRA VALOR EN GENERAR_CONSULTA()
Dim Q_COL_COMPARAR_GRUPO As Integer ' INDICA LA COLUMNA PARA COMPARAR EL GRUPO
                                    ' OBTENDRA VALOR EN GENERAR_CONSULTA()
Dim Q_COL_ARR_TOTAL As Integer      ' NOS INDICA EL TOTAL DE COLUMNAS QUE VA A CONTENER LOS ACUMULADOS
                                    ' OBTENDRA VALOR EN VALIDAR_CONSULTA()
                                    ' SI SEL MES: ENE, FEB, MAR => Q_COL_ARR_TOTAL= 2
                                    ' SI SEL TRI: ENE-MAR, ABR-JUN => Q_COL_ARR_TOTAL= 1 OBS: SE INICIA DESDE POS=0
Dim F_ES_COMPRA As Boolean          ' INDICA SI ES COMPRA O VENTA
                                    ' TRUE::ES COMPRA, FALSE::ES VENTA
Dim ID_PRODUCCION As String
Dim ID_PRODUCCION_DET As String
Dim ID_RECETA As String
Dim TIPO_VENTANA As e_PROGRAMA
Dim ESTILO_VISTA As Integer

Public Sub RECIBE_LINK_FRM(ID_PRODUCCION1 As String, ID_PRODUCCION_DET1 As String, ID_RECETA1 As String, _
                            TIPO_VENTANA1 As e_PROGRAMA, ESTILO_VISTA1 As Integer, _
                            D_EMISION As String, N_SUPERVISOR As String)
    ID_PRODUCCION = ID_PRODUCCION1
    ID_PRODUCCION_DET = ID_PRODUCCION_DET1
    ID_RECETA = ID_RECETA1
    TIPO_VENTANA = TIPO_VENTANA1
    ESTILO_VISTA = ESTILO_VISTA1
    
    lbl(0).Caption = D_EMISION
    lbl(2).Caption = N_SUPERVISOR
    
    ' DEL NOMBRE DEL FRM
    Select Case TIPO_VENTANA
        Case 0: Me.Caption = "Consulta de Insumos"
        Case 1: Me.Caption = "Consulta de Tarea"
        Case 1: Me.Caption = "Consulta de Equipos"
    End Select
    Me.Caption = "Producción - " & Me.Caption
    On Error GoTo error
    
    Dim POS_ARR As Integer
    Erase ARR_TMP()
    ReDim ARR_TMP(3, 1)     ' 0 PROGRAMADO=>> 0::TOTAL,1::TOTAL GEN
                            ' 1 TEORICO=>> 0::TOTAL,1::TOTAL GEN
                            ' 2 REAL=>> 0::TOTAL,1::TOTAL GEN
                            ' 3 DIF=>> 0::TOTAL,1::TOTAL GEN
    Q_COL_ARR_TOTAL = 0
    ' Consultar
    pConsultar
    Exit Sub

SALIR:
    Exit Sub

error:
    SHOW_ERROR
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--pConsultar
            pConsultar
        
        Case 1 '--pImprimir
            pImprimir
        
        Case 3 '--SALIR
            Unload Me
    End Select
End Sub

Private Sub pConsultar()
    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    Dim CN_TMP As New ADODB.Connection       ' CONEX TEMPORAL
    Dim Rst_RUTA As New ADODB.Recordset      ' CARGA RUTAS DE BD'S
    Dim nSQL As String                       ' RECIBIR LA CONSULTA
    ' CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    Dim N_ANYO As String
    Dim SQL_ANYO As String
    Dim k As Integer
    
    If Validar_Consulta() = False Then Exit Sub
    
    BAND_INTERRUMPIR = False
    ' CONFIGURAR LA PRESENTACION DE LA CONSULTA
    LimpiarGrid Me.Fg1, False, 1
    ' INVOCAR A ESTA FUNCION PARA OBTENER LOS VALORES DE
    ' ENTRAR SOLO UNA VEZ
    nSQL = GENERAR_CONSULTA()
    Configurar_Grilla
        
    ' LIMPIAR ARRAY
    Limpiar_ARRAY_TOTAL True
    
    Me.MousePointer = vbHourglass
    DoEvents
    
    If nSQL = "" Then GoTo SALIR
    ' CARGADO EL RST
    RST_Busq rst_select, nSQL, xCon
    pCargarDatosGrid rst_select

SALIR:
    Set Rst_RUTA = Nothing
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    BAND_INTERRUMPIR = False
    Exit Sub

error:
    BAND_INTERRUMPIR = False
    Me.MousePointer = vbDefault
    Set rst_select = Nothing
    SHOW_ERROR Me.Name, "pConsultar"
End Sub


'*****************************************************************************************************
'* Nombre           : pCargarDatosGrid
'* Tipo             : FUNCION
'* Descripcion      : FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
'* Paranetros       : NOMBRE      |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    RST_ORIGEN  |  ADODB.Recordset  |
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
                Case 1
                    pCargarDatosGridAddTotales True, "Tot Gen:", True
            End Select
        End If
    Wend
End Function

'*****************************************************************************************************
'* Nombre           : Comparar_Grupo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FUNCION QUE NOS PERMITE ARMAR LOS GRUPOS, COMPARA CUANDO CAMBIAR DE GRUPO
'* Paranetros       : NOMBRE           |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    RST_ORIGEN       |  ADODB.Recordse   |
'*                    BAND_ADD_REG     |  Boolean          |
'*                    Q_COL_COMPARAR   |  Integer          |
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
        
        If LCase(vStrCampo) = "canprog" Then
            ARR_TMP(0, 0) = ARR_TMP(0, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
            
        ElseIf LCase(vStrCampo) = "canteo" Then
            ARR_TMP(1, 0) = ARR_TMP(1, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
        
        ElseIf LCase(vStrCampo) = "canreal" Then
            ARR_TMP(2, 0) = ARR_TMP(2, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
            
        ElseIf LCase(vStrCampo) = "dif" Then
            ARR_TMP(3, 0) = ARR_TMP(3, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
        Else
                        
        End If
    Next Q_CAMPO
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosGridArrayTmp
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
'* Paranetros       : NOMBRE      |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    RST_ORIGEN  |  ADODB.Recordset  |
'*                    Q_ROW       |  Long             |
'* Devuelve         :
'*****************************************************************************************************
Private Function pCargarDatosGridArrayTmp(RST_ORIGEN As ADODB.Recordset, _
                                         Q_ROW As Long)
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
       
        Select Case vStrCampo
            Case "canprog", "canreal"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_CANTIDAD)
            
            Case "unid"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PU)
            
            Case "dif"
                If NulosN(RST_ORIGEN.Fields(vStrCampo)) > 0 Then '--azul (consumo ahorrado)
                    FORMATO_CELDA Fg1, Q_ROW, Q_CAMPO + 1, &HFF0000, False, &HFFFFFF, Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_CANTIDAD)
                ElseIf NulosN(RST_ORIGEN.Fields(vStrCampo)) < 0 Then '--rojo (consumo adicional)
                    FORMATO_CELDA Fg1, Q_ROW, Q_CAMPO + 1, &HFF, False, &HFFFFFF, Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_CANTIDAD)
                Else
                    Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                End If
            
            Case "percendesvio"
                If NulosN(RST_ORIGEN.Fields(vStrCampo)) > 0 Then '--azul (consumo ahorrado)
                    FORMATO_CELDA Fg1, Q_ROW, Q_CAMPO + 1, &HFF0000, False, &HFFFFFF, Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PORCENTAJE) + "%"
                ElseIf NulosN(RST_ORIGEN.Fields(vStrCampo)) < 0 Then '--rojo (consumo adicional)
                    FORMATO_CELDA Fg1, Q_ROW, Q_CAMPO + 1, &HFF, False, &HFFFFFF, Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PORCENTAJE) + "%"
                Else
                    Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PORCENTAJE) + "%"
                End If
            
            Case Else
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
        End Select
    Next
End Function

'*****************************************************************************************************
'* Nombre           : pImprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pImprimir()
    On Error GoTo error
    Dim oPrint As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    If F_ES_COMPRA = False Then T_RPT_TITULO = Replace(T_RPT_TITULO, "COMPRA", "VENTA")
    oPrint.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO + " ", "Supervizado por: " + lbl(2).Caption, "Fch.Producción " & lbl(0).Caption, False, True
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If BAND_INTERRUMPIR = False Then Exit Sub
    If KeyCode = vbKeyEscape And Shift = 0 Then Unload Me
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    CentrarFrm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase ARR_TMP
End Sub

'*****************************************************************************************************
'* Nombre           : Validar_Consulta
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FUNCION QUE VALIDARA LA CONSULTA
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function Validar_Consulta() As Boolean
    Validar_Consulta = True
End Function

'*****************************************************************************************************
'* Nombre           : GENERAR_CONSULTA
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
'* Paranetros       :
'* Devuelve         : String
'*****************************************************************************************************
Private Function GENERAR_CONSULTA() As String
    Dim nSQL As String            ' CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    Dim nSQLFiltro As String
    Dim k&
    
    ' DEL PROGRAMA
    nSQLFiltro = " pro_producciondet.idpro = " + ID_PRODUCCION + " "
    
    ' DE LA RECETA
    If ID_RECETA <> "-1" Then nSQLFiltro = nSQLFiltro + " AND pro_producciondet.idrec= " + ID_RECETA + " "
    
    ' GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim N_CAMPOS As String
    Dim N_WHERE As String
    Dim N_FROM As String
    Dim N_GROUP_BY As String
    Dim N_HAVING As String
    Dim N_ORDER_BY As String
    
    Select Case ESTILO_VISTA      ' PARAMETRO
        Case 0, 1   ' 0 = INSUMO X PRODUCTO    1 = INSUMO TODOS LOS PRODUCTO
            Q_COL_FILA_OCULTA = 3:         Q_COL_FILA = 9:        Q_POSICION_TOTAL = 7:        Q_COL_COMPARAR_GRUPO = 3
            N_CAMPOS = "  pro_producciondet.idrec, alm_inventario.tippro, pro_producciondetins.iditem, alm_inventario_1.descripcion AS proddesc, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.descripcion, mae_unidades.abrev, Sum(IIf(pro_programadet.canpro Is Null,0,pro_programadet.canpro*pro_recetains.canpro)) AS canprog, Sum(IIf(pro_producciondet.cantidad Is Null,0,(pro_producciondet.cantidad*pro_recetains.canpro))) AS canteo, Sum(pro_producciondetins.canutil) AS canreal, [canteo]-[canreal] AS dif, IIf([canteo]=0 Or [dif]=0,0,[dif]/[canteo]*100) AS percendesvio "
            N_GROUP_BY = " pro_producciondet.idrec, alm_inventario.tippro, pro_producciondetins.iditem, alm_inventario_1.descripcion, mae_tipoproducto.descripcion, alm_inventario.descripcion, mae_unidades.abrev, pro_producciondet.idpro "
            N_ORDER_BY = " alm_inventario_1.descripcion, mae_tipoproducto.descripcion, alm_inventario.descripcion "
        
        Case 2      ' 2 = INSUMO RESUMEN
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 8:        Q_POSICION_TOTAL = 5:        Q_COL_COMPARAR_GRUPO = -1
            N_CAMPOS = " alm_inventario.tippro, pro_producciondetins.iditem, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.descripcion, mae_unidades.abrev, Sum(IIf(pro_programadet.canpro Is Null,0,pro_programadet.canpro*pro_recetains.canpro)) AS canprog, Sum(IIf([pro_producciondet].[cantidad] Is Null,0,([pro_producciondet].[cantidad]*[pro_recetains].[canpro]))) AS canteo, Sum(pro_producciondetins.canutil) AS canreal, [canteo]-[canreal] AS dif, IIf([canteo]=0 Or [dif]=0,0,[dif]/[canteo]*100) AS percendesvio  "
            N_GROUP_BY = " alm_inventario.tippro, pro_producciondetins.iditem, mae_tipoproducto.descripcion, alm_inventario.descripcion, mae_unidades.abrev, pro_producciondet.idpro "
            N_ORDER_BY = " mae_tipoproducto.descripcion, alm_inventario.descripcion; "
            
        Case 3      ' 3 = INSUMO:NUM DE PRODUCCION
            Q_COL_FILA_OCULTA = 3:         Q_COL_FILA = 10:        Q_POSICION_TOTAL = 7:        Q_COL_COMPARAR_GRUPO = -1
            N_CAMPOS = " pro_producciondet.idrec, alm_inventario.tippro, pro_producciondetins.iditem, [pro_receta].[descripcion] & '  - N° Prod.: ' & [pro_producciondet].[numparte] AS recdesc, mae_tipoproducto.descripcion AS tipprodesc, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro AS unid, IIf(pro_programadet.canpro Is Null,0,pro_programadet.canpro*pro_recetains.canpro) AS canprog, IIf([pro_producciondet].[cantidad] Is Null,0,([pro_producciondet].[cantidad]*[pro_recetains].[canpro])) AS canteo, pro_producciondetins.canutil AS canreal, [canteo]-[canreal] AS dif,IIf([canteo]=0 Or [dif]=0,0,[dif]/[canteo]*100) AS percendesvio  "
            N_GROUP_BY = " pro_producciondet.idrec, alm_inventario.tippro, pro_producciondetins.iditem, [pro_receta].[descripcion] & '  - N° Prod.: ' & [pro_producciondet].[numparte], mae_tipoproducto.descripcion, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, IIf(pro_programadet.canpro Is Null,0,pro_programadet.canpro*pro_recetains.canpro), IIf(pro_producciondet.cantidad Is Null,0,(pro_producciondet.cantidad*pro_recetains.canpro)), pro_producciondetins.canutil, pro_producciondet.idpro,pro_producciondetins.numparte "
            N_ORDER_BY = " mae_tipoproducto.descripcion, alm_inventario.descripcion; "
            
        Case 5, 6, 7  ' 5 = TAREA X PRODUCTO TODA PROGRAMACION    6 = TAREA X PRODUCTO DIA ACTUAL    7 = TAREA TODO PROD TODA PROGRAMACION
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 6:        Q_POSICION_TOTAL = 6:        Q_COL_COMPARAR_GRUPO = 2
            T_RPT_TITULO = "SDFSDF"
            N_CAMPOS = " pro_receta.id, pro_recetatar.idtar, pro_receta.descripcion AS recdesc, pro_tareas.descripcion AS tardesc, mae_unidades.abrev, pro_recetatar.cantidad "
            N_GROUP_BY = " pro_receta.id, pro_recetatar.idtar, pro_receta.descripcion, pro_tareas.descripcion, pro_recetatar.orden, mae_unidades.abrev, pro_recetatar.cantidad "
            N_ORDER_BY = " pro_receta.descripcion, pro_tareas.descripcion "
        
        Case 8, 9   ' 8 = TAREA TODO PROD DIA ACTUAL     9 = TAREA TODO PROD RESUMEN
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 3:        Q_POSICION_TOTAL = 3:        Q_COL_COMPARAR_GRUPO = -1
            N_CAMPOS = " pro_recetatar.idtar, pro_tareas.descripcion AS tardesc, mae_unidades.abrev "
            N_GROUP_BY = " pro_recetatar.idtar, pro_tareas.descripcion, pro_recetatar.orden, mae_unidades.abrev "
            N_ORDER_BY = " pro_recetatar.orden "
            
        Case 10, 11, 12  ' 10::EQUIPO X PRODUCTO TODA PROGRAMACION
                         ' 11::EQUIPO X PRODUCTO DIA ACTUAL
                         ' 12::EQUIPO TODO PROD TODA PROGRAMACION
        
        Case 13, 14      ' 13::EQUIPO TODO PROD DIA ACTUAL
                         ' 14::EQUIPO TODO PROD RESUMEN
    End Select
    
    Select Case TIPO_VENTANA
        Case 0 ' INSUMO
            ' DEL PARTE DE PRODUCCION
            If ID_PRODUCCION_DET <> "-1" Then nSQLFiltro = nSQLFiltro + " AND pro_producciondetins.numparte= '" + ID_PRODUCCION_DET + "' "
            T_RPT_TITULO = "REPORTE DE INSUMOS"
            N_FROM = "  (pro_producciondet LEFT JOIN pro_programadet ON (pro_producciondet.idrec = pro_programadet.idrec) AND (pro_producciondet.idpro = pro_programadet.idpro)) INNER JOIN (mae_tipoproducto RIGHT JOIN (((((pro_producciondetins LEFT JOIN mae_unidades ON pro_producciondetins.idunimed = mae_unidades.id) LEFT JOIN pro_recetains ON (pro_producciondetins.idrec = pro_recetains.idrec) AND (pro_producciondetins.iditem = pro_recetains.iditem)) LEFT JOIN alm_inventario ON pro_producciondetins.iditem = alm_inventario.id) LEFT JOIN pro_receta ON pro_producciondetins.idrec = pro_receta.id) LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_receta.iditem = alm_inventario_1.id) ON mae_tipoproducto.id = alm_inventario.tippro) ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro) "
       
        Case 1 ' TAREA
            T_RPT_TITULO = "REPORTE DE TAREAS"
            N_FROM = " pro_tareas INNER JOIN ((pro_receta INNER JOIN pro_programadet ON pro_receta.id = pro_programadet.idrec) INNER JOIN (mae_unidades INNER JOIN pro_recetatar ON mae_unidades.id = pro_recetatar.idunimed) ON pro_receta.id = pro_recetatar.idrec) ON pro_tareas.id = pro_recetatar.idtar "
        
        Case 2 ' EQUIPO
    End Select
        
    N_HAVING = nSQLFiltro
  
    ' DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    Q_POS_MES_INICIO = Q_COL_FILA
    
    ' GENERANDO LA CONSULTA
    nSQL = "SELECT " + N_CAMPOS + _
        vbCr + " FROM " + N_FROM + _
        vbCr + " GROUP BY " + N_GROUP_BY + _
        vbCr + " HAVING " + N_HAVING + _
        vbCr + " ORDER BY " + N_ORDER_BY
    
    GENERAR_CONSULTA = nSQL
End Function



Private Sub Limpiar_ARRAY_TOTAL(Optional F_LIMPIA_TOT_GRL As Boolean = False)
    Dim k As Integer
    For k = 0 To UBound(ARR_TMP())
        ARR_TMP(k, 0) = 0
        If F_LIMPIA_TOT_GRL = True Then ARR_TMP(k, 1) = 0
    Next
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
    Dim Q_POS As Integer
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
        For Q_POS = 0 To UBound(ARR_TMP())
            ARR_TMP(Q_POS, 1) = NulosN(ARR_TMP(Q_POS, 1)) + NulosN(ARR_TMP(Q_POS, 0))
        Next Q_POS
    End If
    
    FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 4, , , , Format(IIf(Band_Total_gral = False, ARR_TMP(0, 0), ARR_TMP(0, 1)), FORMAT_CANTIDAD)
    FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 3, , , , Format(IIf(Band_Total_gral = False, ARR_TMP(1, 0), ARR_TMP(1, 1)), FORMAT_CANTIDAD)
    FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 2, , , , Format(IIf(Band_Total_gral = False, ARR_TMP(2, 0), ARR_TMP(2, 1)), FORMAT_CANTIDAD)
    FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 1, , , , Format(IIf(Band_Total_gral = False, ARR_TMP(3, 0), ARR_TMP(3, 1)), FORMAT_CANTIDAD)
    Err.Clear
End Sub

'*****************************************************************************************************
'* Nombre           : Configurar_Grilla
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA DE ACUERDO A LO QUE SE SELECCIONA
'* Paranetros       : NOMBRE              |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    F_CONSERVAR_FORMATO |  Boolean   |
'* Devuelve         :
'*****************************************************************************************************
Private Sub Configurar_Grilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    Dim M_ANCHO_COL_MES As Integer       ' DEPENDERA DEL TIPO DE PRESENTACION EN DECIMALES, EN MILES
    Dim k, j As Integer
    
    If F_CONSERVAR_FORMATO = True Then Fg1.Clear
    
    Fg1.FrozenCols = 0
    
    M_ANCHO_COL_MES = 1100

    With Fg1
        Fg1.Cols = Q_COL_FILA_OCULTA + Q_COL_FILA
        Q_POS_MES = Q_POS_MES_INICIO
        .TextMatrix(0, Fg1.Cols - 5) = "Programado":    .ColWidth(Fg1.Cols - 5) = M_ANCHO_COL_MES:         .ColAlignment(Fg1.Cols - 5) = flexAlignRightCenter
        .TextMatrix(0, Fg1.Cols - 4) = "Teórico":       .ColWidth(Fg1.Cols - 4) = M_ANCHO_COL_MES::        .ColAlignment(Fg1.Cols - 4) = flexAlignRightCenter
        .TextMatrix(0, Fg1.Cols - 3) = "Real":          .ColWidth(Fg1.Cols - 3) = M_ANCHO_COL_MES:         .ColAlignment(Fg1.Cols - 3) = flexAlignRightCenter
        .TextMatrix(0, Fg1.Cols - 2) = "Desvio":        .ColWidth(Fg1.Cols - 2) = M_ANCHO_COL_MES:         .ColAlignment(Fg1.Cols - 2) = flexAlignRightCenter
        .TextMatrix(0, Fg1.Cols - 1) = "% Desvio":      .ColWidth(Fg1.Cols - 1) = M_ANCHO_COL_MES:         .ColAlignment(Fg1.Cols - 1) = flexAlignRightCenter
        
        If chk(0).Value = 0 Then .ColWidth(Fg1.Cols - 5) = 0
        If chk(1).Value = 0 Then .ColWidth(Fg1.Cols - 4) = 0
        If chk(2).Value = 0 Then .ColWidth(Fg1.Cols - 3) = 0
        If chk(3).Value = 0 Then .ColWidth(Fg1.Cols - 2) = 0
        If chk(4).Value = 0 Then .ColWidth(Fg1.Cols - 1) = 0
        
        .ColWidth(0) = 200
        
        ' DATOS DE FILA
        Select Case ESTILO_VISTA '--PARAMETRO
            Case 0, 1, 3    '0 = INSUMO X PRODUCTO     1 = INSUMO X TODOS LOS PRODUCTOS      3 = INSUMO X PARTE DE PRODUCCION
                .TextMatrix(0, 4) = "Receta":           .ColWidth(4) = 0:       .ColAlignment(4) = flexAlignLeftCenter
                .TextMatrix(0, 5) = "Tipo Producto":    .ColWidth(5) = 1200:    .ColAlignment(5) = flexAlignLeftCenter
                .TextMatrix(0, 6) = "Descripción":      .ColWidth(6) = 4000:    .ColAlignment(6) = flexAlignLeftCenter
                .TextMatrix(0, 7) = "U.M.":             .ColWidth(7) = 500:     .ColAlignment(7) = flexAlignLeftCenter
                .TextMatrix(0, 8) = "Unid.":            .ColWidth(8) = 0:       .ColAlignment(8) = flexAlignRightBottom
                .FrozenCols = 8
                If ESTILO_VISTA <> 3 Then .ColWidth(6) = 5000
                
            Case 2          ' 3::INSUMO RESUMEN
                .TextMatrix(0, 3) = "Tipo Producto":    .ColWidth(3) = 1500:     .ColAlignment(3) = flexAlignLeftCenter
                .TextMatrix(0, 4) = "Descripción":      .ColWidth(4) = 4200:     .ColAlignment(4) = flexAlignLeftCenter
                .TextMatrix(0, 5) = "U.M.":             .ColWidth(5) = 500:      .ColAlignment(5) = flexAlignLeftCenter
                .FrozenCols = 5
                
            Case 5, 6, 7    '--4::TAREA X PRODUCTO TODA PROGRAMACION
                            '--5::TAREA X PRODUCTO DIA ACTUAL
                            '--6::TAREA TODO PROD TODA PROGRAMACION
                
            Case 8, 9   '--8::TAREA TODO PROD DIA ACTUAL
                        '--9::TAREA TODO PROD RESUMEN
                
            Case 10, 11, 12 '--10::EQUIPO X PRODUCTO TODA PROGRAMACION
                            '--11::EQUIPO X PRODUCTO DIA ACTUAL
                            '--12::EQUIPO TODO PROD TODA PROGRAMACION
            
            Case 13, 14 '--13::EQUIPO TODO PROD DIA ACTUAL
                        '--14::EQUIPO TODO PROD RESUMEN
        
        End Select

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
'* Paranetros       : NOMBRE           |  TIPO         |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Band_Total_gral  |  Boolean      |
'*                    Q_POS            |  Integer      |
'* Devuelve         : String
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

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If Index <> 5 Then Exit Sub
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

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
'* Descripcion      : EXPORTA A EXCEL LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Orden de Producción", "Fch. Producción: " & lbl(0).Caption, "Supervizado por : " & lbl(2).Caption, "Orden de Producción"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub
