VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmConsIngAlmacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de ingreso/salida de almacen"
   ClientHeight    =   7620
   ClientLeft      =   -195
   ClientTop       =   2865
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleMode       =   0  'User
   ScaleWidth      =   12737.97
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6465
      Left            =   15
      TabIndex        =   7
      Top             =   1155
      Width           =   11895
      _cx             =   20981
      _cy             =   11404
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
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsIngAlmacen.frx":0000
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
   Begin VB.Frame Frame2 
      Height          =   1155
      Left            =   9495
      TabIndex        =   6
      Top             =   -30
      Width           =   2415
      Begin VB.CommandButton CmdSalir 
         Height          =   600
         Left            =   1530
         Picture         =   "FrmConsIngAlmacen.frx":017F
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Salir"
         Top             =   345
         Width           =   570
      End
      Begin VB.CommandButton CmdConsultar 
         Height          =   600
         Left            =   315
         Picture         =   "FrmConsIngAlmacen.frx":0489
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar"
         Top             =   345
         Width           =   570
      End
      Begin VB.CommandButton CmdImprimir 
         Height          =   600
         Left            =   930
         Picture         =   "FrmConsIngAlmacen.frx":08CB
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir"
         Top             =   345
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   2985
      TabIndex        =   2
      Top             =   -30
      Width           =   6450
      Begin VB.CommandButton CmdBusProducto 
         Height          =   225
         Left            =   2025
         Picture         =   "FrmConsIngAlmacen.frx":0BD5
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   675
         Width           =   225
      End
      Begin VB.Frame FraSolic 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   5940
         TabIndex        =   12
         Top             =   915
         Visible         =   0   'False
         Width           =   6210
         Begin VB.CommandButton CmdBusSolic 
            Height          =   240
            Left            =   1950
            Picture         =   "FrmConsIngAlmacen.frx":0D07
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   225
         End
         Begin VB.TextBox TxtIdSolicitante 
            Height          =   300
            Left            =   1290
            MaxLength       =   4
            TabIndex        =   14
            Text            =   "TxtIdSolicitante"
            Top             =   90
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   135
            Width           =   735
         End
         Begin VB.Label LblSolicitante 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblSolicitante"
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
            Left            =   2250
            TabIndex        =   15
            Top             =   90
            Width           =   3885
         End
      End
      Begin VB.TextBox TxtIdProducto 
         Height          =   300
         Left            =   1365
         MaxLength       =   5
         TabIndex        =   17
         Text            =   "TxtIdProducto"
         Top             =   645
         Width           =   915
      End
      Begin VB.Frame FraProvee 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   60
         TabIndex        =   8
         Top             =   195
         Width           =   6210
         Begin VB.CommandButton CmdBusProvDeAlmIngreso 
            Height          =   240
            Left            =   1950
            Picture         =   "FrmConsIngAlmacen.frx":0E39
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   225
         End
         Begin VB.TextBox TxtIdProveedor 
            Height          =   300
            Left            =   1290
            MaxLength       =   3
            TabIndex        =   10
            Text            =   "TxtIdProve"
            Top             =   90
            Width           =   915
         End
         Begin VB.Label LblProv 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   135
            Width           =   735
         End
         Begin VB.Label LblProveedor 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProveedor"
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
            Left            =   2250
            TabIndex        =   11
            Top             =   90
            Width           =   3885
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   195
         Left            =   210
         TabIndex        =   21
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblProducto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblProducto"
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
         Left            =   2325
         TabIndex        =   18
         Top             =   645
         Width           =   3900
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1155
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   2895
      Begin VB.OptionButton OptIng 
         Caption         =   "Ingreso"
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
         Left            =   435
         TabIndex        =   1
         Top             =   855
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.OptionButton OptSal 
         Caption         =   "Salida"
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
         Left            =   1545
         TabIndex        =   3
         Top             =   855
         Width           =   885
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec1 
         Height          =   300
         Left            =   930
         TabIndex        =   22
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
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
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec2 
         Height          =   300
         Left            =   930
         TabIndex        =   24
         Top             =   495
         Width           =   1275
         _ExtentX        =   2249
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
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   270
         TabIndex        =   25
         Top             =   495
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   210
         Width           =   465
      End
   End
End
Attribute VB_Name = "FrmConsIngAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCONSINGALMACEN.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO QUE MUESTRA UN REPORTE DE TODOS LOS INGRESOS Y SALIDAS DEL ALMACEN,
'*                    EN FUNCION A CRITERIOS ESPECIFICADOS POR EL USUARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 15/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim vStr As String, vFormatString As String, CaracteresNumericos As String
Dim RstConsIngSal As New ADODB.Recordset
Dim SeEjecuto As Boolean

'*****************************************************************************************************
'* Nombre           : SumaCant
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : SUMA LAS CONTIDADES DE LA COLUMNA 9 DE CONTROL FlexGrid Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub SumaCant()
    Dim i As Long, vTotal As Double
    For i = 1 To Fg1.Rows - 1
        vTotal = vTotal + Val(Format(Fg1.TextMatrix(i, 9), "#####0.00"))
    Next
    If Trim(Fg1.TextMatrix(1, 1)) <> "" Then
        Fg1.AddItem ""
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = "Total:"
        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(vTotal, "#,###0.00")
        Fg1.Row = Fg1.Rows - 1: Fg1.Col = 4
        Fg1.CellForeColor = &H800000
        Fg1.Row = Fg1.Rows - 1: Fg1.Col = 9
        Fg1.CellForeColor = &H800000
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ConfigIni
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LA CONFIGURACION INICIAL PARA EL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ConfigIni()
    TxtFec1.Valor = Date
    TxtFec2.Valor = Date
    FraSolic.Top = 150: FraSolic.Left = 75
    TxtIdSolicitante.Text = ""
    LblSolicitante.Caption = ""
    TxtIdProveedor.Text = ""
    LblProveedor.Caption = ""
    TxtIdProducto.Text = ""
    lblProducto.Caption = ""
    Fg1.ColWidth(11) = 0: Fg1.ColWidth(10) = 0
    Fg1.TextMatrix(0, 5) = "Proveedor"
    FraProvee.Top = 150: FraProvee.Left = 75
    FraProvee.Visible = True
    LimpiarGrid
    Fg1.FrozenCols = 4
End Sub

'*****************************************************************************************************
'* Nombre           : LimpiarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL GRID PARA MOSTRAR NUEVOS REGISTROS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub LimpiarGrid()
    Fg1.Clear
    Fg1.Rows = 2
    Fg1.FormatString = vFormatString
    If OptIng.Value = True Then
        Fg1.TextMatrix(0, 5) = "Proveedor"
        Fg1.ColWidth(6) = 0
        Fg1.ColWidth(10) = 0
        Fg1.ColWidth(11) = 0
    Else
        Fg1.TextMatrix(0, 5) = "Cliente"
        Fg1.ColWidth(6) = 0
        Fg1.ColWidth(10) = 0
        Fg1.ColWidth(11) = 2055
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : fFuncJalarDatos
'* Tipo             : FUNCION
'* Descripcion      :
'* Paranetros       : NOMBRE           |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pId              |  LONG      |
'*                    pEspecificTabla  |  STRING    |
'* Devuelve         : STRING
'*****************************************************************************************************
Private Function fFuncJalarDatos(pId As Long, pEspecificTabla As String) As String
    Dim RsFunEsp As New ADODB.Recordset
    RsFunEsp.CursorLocation = adUseClient
    Select Case Trim(pEspecificTabla)
        Case "SOLIC" 'SOLICITANTE
            RST_Busq RsFunEsp, "SELECT id, LTRIM(UCASE(ape)) + ', ' + LTRIM(nom) AS DatAJalar FROM pla_empleados WHERE id = " & pId & "", xCon
        Case "UMED" 'UNIDAD DE MEDIDA
            RST_Busq RsFunEsp, "SELECT id, abrev as DatAJalar FROM mae_unidades WHERE id = " & pId & "", xCon
    End Select
    If RsFunEsp.RecordCount > 0 Then
        If NulosC(RsFunEsp("DatAJalar")) <> "" Then
            fFuncJalarDatos = Trim(RsFunEsp("DatAJalar"))
        Else
            fFuncJalarDatos = ""
        End If
    Else
        fFuncJalarDatos = ""
    End If
    Set RsFunEsp = Nothing
End Function

Private Sub CmdBusProducto_Click()
    ' BUSCA UN PRODUCTO EN LA TABLA Alm_inventario
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Producto":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "id":        xCampos(1, 1) = "id":            xCampos(1, 2) = "1000":    xCampos(1, 3) = "N"
    xCampos(2, 0) = "Codigo":    xCampos(1, 1) = "codpro":        xCampos(2, 2) = "1200":    xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT id, codpro, descripcion FROM alm_inventario"
    
    xform.Titulo = "Buscando Productos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdProducto.Text = xRs("id")
        lblProducto.Caption = xRs("descripcion")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusProvDeAlmIngreso_Click()
    ' BUSCA UN PROVEEDOR O UN CLIENTE EN LOS INGRESOS O SALIDAS REGISTRADAS
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "nombre":     xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":               xCampos(1, 1) = "idpro":      xCampos(1, 2) = "2000":   xCampos(1, 3) = "N"
    
    If OptIng.Value = True Then 'INGRESO
        xform.SQLCad = "SELECT DISTINCT idpro, nombre FROM alm_ingreso WHERE tipmov = -1"
    Else 'SALIDA
        xform.SQLCad = "SELECT DISTINCT idpro, nombre FROM alm_ingreso WHERE tipmov = 0"
    End If
    
    xform.Titulo = "Buscando Proveedores"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If NulosN(xRs("idpro")) <= 0 Then
            TxtIdProveedor.Text = ""
        Else
            TxtIdProveedor.Text = xRs("idpro")
        End If
        LblProveedor.Caption = xRs("nombre")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusSolic_Click()
    ' BUSCA PERSONAL SOLICITANTE DEL INGRESO O SALIDA
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "apenom":     xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":               xCampos(1, 1) = "id":         xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT UCase([pla_empleados]![ape])+', '+[pla_empleados]![nom] AS apenom, pla_empleados.id " _
        & " From pla_empleados ORDER BY UCase([pla_empleados]![ape])+', '+[pla_empleados]![nom]"
    
    xform.Titulo = "Buscando Solicitantes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "apenom"
    xform.CampoBusca = "apenom"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdSolicitante.Text = xRs("id")
        LblSolicitante.Caption = xRs("apenom")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : CmdConsultar_Click
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LOS INGRESOS O SALIDAS EN EL CONTROL FlexGrid Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub CmdConsultar_Click()
    LimpiarGrid
    vStr = "SELECT alm_ingreso.*, mae_documento.abrev, mae_documento.descripcion, alm_ingreso!numser+'-'+alm_ingreso!numdoc AS numdoc2, " _
        & " UCase([pla_empleados]![apepat] & ' ' & [pla_empleados]![apepat]) & '-' & [pla_empleados]![nom] AS nomres, IIf(alm_ingreso!tipmov=-1,'Ingreso','Salida') AS movi, " _
        & " alm_almacenes.descripcion AS descalm, alm_ingreso.idsol AS idsolicitante, alm_ingreso.idare AS idarea, pla_area.descripcion AS nomarea, alm_inventario.descripcion AS nomproducto, " _
        & " alm_ingresodet.cantidad, alm_inventario.idunimed FROM ((((alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN " _
        & " pla_empleados ON alm_ingreso.idres = pla_empleados.id) LEFT JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id) LEFT JOIN pla_area " _
        & " ON alm_ingreso.idare = pla_area.id) LEFT JOIN (alm_inventario RIGHT JOIN alm_ingresodet ON alm_inventario.id = alm_ingresodet.iditem) " _
        & " ON alm_ingreso.id = alm_ingresodet.id "

    If OptIng.Value = True Then
        vStr = vStr & " WHERE alm_ingreso.tipmov = -1 "
    Else
        vStr = vStr & " WHERE alm_ingreso.tipmov = 0 "
    End If
    vStr = vStr & " AND alm_ingreso.fching BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "') "
    
    If Trim(TxtIdProveedor.Text) <> "" Then
        vStr = vStr & " AND alm_ingreso.idpro = " & Val(TxtIdProveedor.Text) & " "
    ElseIf Trim(TxtIdProveedor.Text) = "" And Trim(LblProveedor.Caption) <> "" Then
        vStr = vStr & " AND alm_ingreso.nombre = '" + Trim(LblProveedor.Caption) + "'" & " "
    End If
    If Trim(TxtIdProducto.Text) <> "" Then
        vStr = vStr & " AND alm_inventario.id = " & Val(TxtIdProducto.Text) & " "
    End If

    vStr = vStr & " ORDER BY alm_ingreso.fching, alm_ingreso.numdoc"
    Set RstConsIngSal = New ADODB.Recordset
    RST_Busq RstConsIngSal, vStr, xCon
    
    With RstConsIngSal
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If Fg1.TextMatrix(1, 1) <> "" Then Fg1.AddItem ""
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = Format(NulosC(Trim(.Fields("fching"))), "dd/mm/yyyy")  ' FECHA OPERAC(INGRESO O SALIDA)
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Trim(.Fields("abrev")))                         ' TIPO DOC
                Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Trim(.Fields("numser"))) & "-" & NulosC(Trim(.Fields("numdoc"))) 'NUMDOC
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosC(Trim(.Fields("fchdoc"))), "dd/mm/yyyy")  ' FECHA EMIS
                Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(Trim(.Fields("nombre")))                        ' CLIE/PROVEE
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(Trim(.Fields("nomres")))                        ' RESPONSABLE
                Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(Trim(.Fields("nomproducto")))                   ' PRODUCTO
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = fFuncJalarDatos(NulosN(.Fields("idunimed")), "UMED")   ' UNID MEDIDA
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosN(Trim(.Fields("cantidad"))), "#,###0.00") ' CANTIDAD
                
                If NulosC(.Fields("nomarea")) <> "" Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 11) = Trim(.Fields("nomarea"))                          ' NOMBRE AREA
                End If
                .MoveNext
            Loop
        End If
    End With
    SumaCant
End Sub

Private Sub CmdImprimir_Click()
    FrmPrintIngEgreso.Show
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A AJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    vFormatString = Fg1.FormatString
    CaracteresNumericos = "0123456789." & Chr(8)
    ConfigIni
    FraProvee.BackColor = &H8000000F
    FraSolic.BackColor = &H8000000F
    TxtFec1.Valor = "01/" & Month(Date) & "/" & Year(Date)
End Sub

Private Sub OptIng_Click()
    LimpiarGrid
    FraProvee.Top = 150
    FraProvee.Left = 75
    FraProvee.Visible = True
    FraSolic.Visible = False
    Fg1.TextMatrix(0, 5) = "Proveedor"
    LblProv.Caption = "Proveedor"
End Sub

Private Sub OptSal_Click()
    LimpiarGrid
    FraProvee.Top = 150
    FraProvee.Left = 75
    FraProvee.Visible = True
    FraSolic.Visible = False
    Fg1.TextMatrix(0, 5) = "Cliente"
    LblProv.Caption = "Cliente"
End Sub

Private Sub TxtIdProducto_Change()
    If Trim(TxtIdProducto.Text) = "" Then
        lblProducto.Caption = ""
    End If
End Sub

Private Sub TxtIdProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        If NulosC(TxtIdProducto.Text) = "" Then Exit Sub
        Dim xRs As New ADODB.Recordset
        xRs.CursorLocation = adUseClient
        RST_Busq xRs, "SELECT id, codpro, descripcion FROM alm_inventario WHERE id = " & Val(TxtIdProducto.Text) & "", xCon
        
        If xRs.RecordCount = 0 Then
            TxtIdProducto.Text = ""
            lblProducto.Caption = ""
            CmdConsultar.SetFocus
        Else
            lblProducto.Caption = xRs("descripcion")
        End If
        Set xRs = Nothing
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdProducto_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProducto.Value = True
    End If
End Sub

Private Sub TxtIdProveedor_Change()
    If Trim(TxtIdProveedor.Text) = "" Then
        LblProveedor.Caption = ""
    End If
End Sub

Private Sub TxtIdProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 46 'SUPRIMIR
            If Trim(TxtIdProducto.Text) = "" And Trim(LblProv.Caption) <> "" Then
                LblProveedor.Caption = ""
            End If
    End Select
End Sub

Private Sub TxtIdProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        
        If NulosC(TxtIdProveedor.Text) = "" Then Exit Sub
        Dim xRs As New ADODB.Recordset
        xRs.CursorLocation = adUseClient
        RST_Busq xRs, "SELECT DISTINCT idpro, nombre FROM alm_ingreso WHERE idpro = " & Val(TxtIdProveedor.Text) & " AND tipmov = -1", xCon
        
        If xRs.RecordCount = 0 Then
            TxtIdProveedor.Text = ""
            LblProveedor.Caption = ""
            CmdConsultar.SetFocus
        Else
            LblProveedor.Caption = xRs("nombre")
        End If
        Set xRs = Nothing
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdProveedor_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProvDeAlmIngreso.Value = True
    End If
End Sub

Private Sub TxtIdSolicitante_Change()
    If Trim(TxtIdSolicitante.Text) = "" Then
        LblSolicitante.Caption = ""
    End If
End Sub

Private Sub TxtIdSolicitante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        If NulosC(TxtIdSolicitante.Text) = "" Then Exit Sub
        Dim xRs As New ADODB.Recordset
        xRs.CursorLocation = adUseClient
        RST_Busq xRs, "SELECT id, ltrim(UCASE(ape)) + ', ' + ltrim(nom) AS nomsolic FROM pla_empleados WHERE id = " & Val(TxtIdSolicitante.Text) & "", xCon
        
        If xRs.RecordCount = 0 Then
            TxtIdSolicitante.Text = ""
            LblSolicitante.Caption = ""
            CmdConsultar.SetFocus
        Else
            LblSolicitante.Caption = xRs("nomsolic")
        End If
        Set xRs = Nothing
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdSolicitante_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusSolic.Value = True
    End If
End Sub
