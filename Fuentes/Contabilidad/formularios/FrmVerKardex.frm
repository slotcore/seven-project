VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVerKardex 
   Caption         =   "Contabilidad - Kardex"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   990
      Left            =   3390
      TabIndex        =   24
      Top             =   3270
      Visible         =   0   'False
      Width           =   5145
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   150
         TabIndex        =   25
         Top             =   435
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label7 
         Caption         =   "Cargando el Kardex"
         Height          =   180
         Left            =   165
         TabIndex        =   26
         Top             =   165
         Width           =   1650
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   960
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   5130
         X2              =   5130
         Y1              =   15
         Y2              =   945
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   5115
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5160
         Y1              =   975
         Y2              =   975
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
            Picture         =   "FrmVerKardex.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":08D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":0A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":0DC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":0F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":139A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":14B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":19F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":1F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":204E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":2162
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":25B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex.frx":2722
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
   Begin VB.Frame Frame2 
      Height          =   1170
      Left            =   0
      TabIndex        =   11
      Top             =   285
      Width           =   7755
      Begin VB.TextBox TxtDesc 
         Height          =   300
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "TxtDesc"
         Top             =   480
         Width           =   6525
      End
      Begin VB.TextBox TxtUnidad 
         Height          =   300
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "TxtUnidad"
         Top             =   165
         Width           =   1080
      End
      Begin VB.TextBox TxtSaldo 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "TxtSaldo"
         Top             =   795
         Width           =   1080
      End
      Begin VB.CommandButton CmdProducto 
         Height          =   240
         Left            =   3345
         Picture         =   "FrmVerKardex.frx":2C6A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   195
         Width           =   240
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   1065
         TabIndex        =   3
         Top             =   795
         Width           =   1305
         _ExtentX        =   2302
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
         Valor           =   "23/03/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   3600
         TabIndex        =   4
         Top             =   795
         Width           =   1305
         _ExtentX        =   2302
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
         Valor           =   "23/03/2007"
      End
      Begin VB.TextBox txtCodItem 
         Height          =   300
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "txtCodItem"
         Top             =   165
         Width           =   2550
      End
      Begin VB.Label LblIdProducto 
         AutoSize        =   -1  'True
         Caption         =   "LblIdProducto"
         Height          =   195
         Left            =   3705
         TabIndex        =   22
         Top             =   210
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   525
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   210
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   " Inicial"
         Height          =   195
         Left            =   105
         TabIndex        =   16
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Final"
         Height          =   195
         Left            =   3000
         TabIndex        =   15
         Top             =   840
         Width           =   330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Unidad"
         Height          =   195
         Left            =   5880
         TabIndex        =   14
         Top             =   210
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Stock Actual"
         Height          =   195
         Left            =   5475
         TabIndex        =   13
         Top             =   840
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1170
      Left            =   7800
      TabIndex        =   6
      Top             =   285
      Width           =   4095
      Begin VB.OptionButton Opt1 
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   75
         TabIndex        =   10
         Top             =   285
         Width           =   1245
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "Insumos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   75
         TabIndex        =   9
         Top             =   570
         Width           =   1245
      End
      Begin VB.Frame Frame8 
         Caption         =   "[  Metodo Valorizacion  ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1005
         Left            =   1620
         TabIndex        =   8
         Top             =   120
         Width           =   2400
         Begin VB.CommandButton CmdTodos 
            Caption         =   "Ver Todos"
            Height          =   255
            Left            =   930
            TabIndex        =   28
            Top             =   720
            Width           =   1425
         End
         Begin VB.OptionButton Option2 
            Caption         =   "P.E.P.S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   20
            Top             =   495
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.OptionButton OptVal1 
            Caption         =   "Promedio Ponderado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   19
            Top             =   225
            Width           =   2190
         End
      End
      Begin VB.OptionButton Opt3 
         Caption         =   "Mercaderia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   75
         TabIndex        =   7
         Top             =   855
         Width           =   1290
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         X1              =   1485
         X2              =   1485
         Y1              =   150
         Y2              =   1140
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   1470
         X2              =   1470
         Y1              =   135
         Y2              =   1125
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   5805
      Left            =   0
      TabIndex        =   21
      Top             =   1485
      Width           =   11880
      _cx             =   20955
      _cy             =   10239
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
      BackColor       =   14745342
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   19
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmVerKardex.frx":2D9C
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
      FrozenCols      =   10
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label LblDescripcion 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblDescripcion"
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
      Left            =   0
      TabIndex        =   27
      Top             =   7320
      Width           =   11880
   End
End
Attribute VB_Name = "FrmVerKardex"
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

Dim rst As New ADODB.Recordset            ' RECORSET QUE ALAMCENARA LOS MOVIMIENTOS DEL ITEM
Dim SeEjecuto As Boolean                  ' VARIABLE QUE CONTROLARA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim StockIni As Double                    ' ALMACENA EL STOCK INICIAL DEL ITEM
Dim xPrecioIni As Double                  ' ALMACENA EL PRECIO INICIAL DEL ITEM
Dim MuestraRpt As Integer

Public Sub pCargarRpt()
    ' se usara cunado se desee imprimir todos los productos desde la pantalla del kardex resumen
    MuestraRpt = 1
    Form_Activate
End Sub


Private Sub CmdProducto_Click()
    ' EJECUTA LA BUSQUEDA DE UN ITEM, ESPECIFICANDO EL TIPO DE ITEM QUE SE BUSCARA
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
   
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Producto":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":     xCampos(1, 1) = "codpro":        xCampos(1, 2) = "2000":    xCampos(1, 3) = "C"
    
    Dim nSQLActivos  As String
    If FrmResuMov.chkActivos.Value = 1 Then nSQLActivos = " and alm_inventario.activo =-1 "
        
    If Opt1.Value = True Then
        ' buscamos producto
        xform.SqlCad = "SELECT alm_inventario.*, mae_unidades.abrev FROM mae_unidades RIGHT JOIN alm_inventario ON " _
            & " mae_unidades.id = alm_inventario.idunimed Where  ((alm_inventario.tippro) = 3) " & nSQLActivos _
            & " ORDER BY alm_inventario.descripcion"

        xform.Titulo = "Buscando Producto"
    End If
    If Opt2.Value = True Then
        ' buscacomos materia prima /insumos
        xform.SqlCad = "SELECT alm_inventario.*, mae_unidades.abrev FROM mae_unidades RIGHT JOIN alm_inventario ON " _
            & " mae_unidades.id = alm_inventario.idunimed Where ((alm_inventario.tippro) = 1 " & nSQLActivos _
            & " Or (alm_inventario.tippro) = 4) ORDER BY alm_inventario.descripcion"

        xform.Titulo = "Buscando Materia Prima / Insumos"
    End If
    If Opt3.Value = True Then
        ' buscamos mercaderias
        xform.SqlCad = "SELECT alm_inventario.*, mae_unidades.abrev FROM mae_unidades RIGHT JOIN alm_inventario ON " _
            & " mae_unidades.id = alm_inventario.idunimed Where ((alm_inventario.tippro) = 2)  " & nSQLActivos _
            & " ORDER BY alm_inventario.descripcion"
        
        xform.Titulo = "Buscando Mercaderia"
    End If
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblIdProducto.Caption = xRs("id")
        txtCodItem.Text = NulosC(xRs("codpro"))
        TxtDesc.Text = NulosC(xRs("descripcion"))
        TxtUnidad.Text = NulosC(xRs("abrev"))
        TxtSaldo.Text = Format(NulosN(xRs("stckact")), "0.00")
        StockIni = NulosN(xRs("stckini"))
        
        If xRs("idmon") = 2 Then
            ' hallamos el precio inicial en funcion al precio inicial del producto
            Dim xTCDia As Double
            xTCDia = HallaTipoCambio(TxtFchIni.Valor, "2", Venta, xCon)
            xPrecioIni = NulosN(xRs("preini") * xTCDia)
        Else
            xPrecioIni = NulosN(xRs("preini"))
        End If
        txtCodItem.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TextBox PARA EL INGRESO DE DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    txtCodItem.Text = ""
    TxtUnidad.Text = ""
    TxtDesc.Text = ""
    TxtSaldo.Text = ""
End Sub

Private Sub Fg1_RowColChange()
    LblDescripcion.Caption = Fg1.TextMatrix(Fg1.Row, 13)
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        TxtSaldo.Text = "0"
        If MostrarValorizado = False Then
            Me.Caption = "Almacén - Kardex"
        Else
            Me.Caption = "Contabilidad - Kardex Valorizado"
        End If
        Opt1.Value = True
        Setea
        If UCase(txtCodItem.Text) <> "TXTCODITEM" Then
            Dim rst As New ADODB.Recordset
            Dim SqlCad As String
            OptVal1.Value = True
            
            SqlCad = "SELECT alm_inventario.*, mae_unidades.abrev FROM mae_unidades RIGHT JOIN alm_inventario ON " _
                & " mae_unidades.id = alm_inventario.idunimed Where ((alm_inventario.activo = -1) AND (alm_inventario.id = " & Val(LblIdProducto.Caption) & ")) " _
                & " ORDER BY alm_inventario.descripcion"
                        
            RST_Busq rst, SqlCad, xCon
            If rst.RecordCount <> 0 Then
                LblIdProducto.Caption = rst("id")
                txtCodItem.Text = NulosC(rst("codpro"))
                TxtDesc.Text = NulosC(rst("descripcion"))
                TxtUnidad.Text = NulosC(rst("abrev"))
                TxtSaldo.Text = Format(NulosN(rst("stckact")), "0.00")
                StockIni = NulosN(rst("stckini"))
                
                If rst("idmon") = 2 Then
                    ' hallamos el precio inicial en funcion al precio inicial del producto
                    Dim xTCDia As Double
                    xTCDia = HallaTipoCambio(TxtFchIni.Valor, "2", Venta, xCon)
                    xPrecioIni = NulosN(rst("preini") * xTCDia)
                Else
                    xPrecioIni = NulosN(rst("preini"))
                End If
                
                If MuestraRpt <> 1 Then txtCodItem.SetFocus
                
                If OptVal1.Value = True Then
                    MuestraKardexProm NulosN(LblIdProducto.Caption)
                End If
            End If
        Else
            Opt1.Value = True
            Blanquea
        End If
        OptVal1.Value = True
        SeEjecuto = True
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : MostrarDocumentos
'* Tipo             : FUNCCION
'* Descripcion      : DEVUELVE LOS NUMEROS DE LOS DOCUMENTO DE COMPRA O VENTA  VINCULADOS AL INGRESO
'*                    O SALIDA DE ALMACEN, ESTA FUNCION DEVUELVE UNA CADENA
'* Paranetros       : NOMBRE      |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    IdDocumento |  INTEGER   |  ESPECIFICA EL ID DEL DOCUMENTO QUE SE ESTA BUSCANDO
'*                    DondeBuscar |  String    |  ESPECIFICA DONDE SE EFECTUARA LA BUSQUEA
'*                                                AI Almacen Ingreso, GR Guia de Remision 'ventas
'* Devuelve         : String
'*****************************************************************************************************
Private Function MostrarDocumentos(IdDocumento, DondeBuscar As String) As String
    Dim rst As New ADODB.Recordset
    Dim xCad As String
    Dim nSQL As String
    
    If DondeBuscar = "AI" Then
        nSQL = "SELECT [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, mae_prov.nombre FROM mae_prov RIGHT JOIN (alm_ingresodoc LEFT JOIN " _
        & " com_compras ON alm_ingresodoc.iddoc = com_compras.id) ON mae_prov.id = com_compras.idpro WHERE (((alm_ingresodoc.id)=" & IdDocumento & "))"
        
    ElseIf DondeBuscar = "GR" Then
        
        nSQL = "SELECT [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, mae_cliente.nombre " _
            + vbCr + " FROM (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) RIGHT JOIN vta_guia ON vta_ventas.id = vta_guia.iddocven " _
            + vbCr + " WHERE (((vta_guia.id)=" & IdDocumento & "));"
    End If
    
    RST_Busq rst, nSQL, xCon
    
    Do While Not rst.EOF
        xCad = xCad + NulosC(rst("numdoc")) + " " + NulosC(rst("nombre")) + ", "
        rst.MoveNext
    Loop
    If xCad <> "" Then xCad = Mid(xCad, 1, Len(xCad) - 2)
    
    MostrarDocumentos = xCad
    
End Function

'**********************************************
' Creado: 02/05/2012 - Jose Chacon
' Halla el numero de Solicitud de MAteriales
' Halla el numero de Registro de Produccion
' de una salida de Almacen
'**********************************************
Private Sub hallarNumProd(IDING_ As Integer, GRID_ As VSFlexGrid, FILA_ As Integer, COLUMNA_ As Integer)
'    Dim xRs As New ADODB.Recordset
'    Dim cSQL As String
'
'    cSQL = "SELECT alm_ingreso.id, alm_ingreso.idorddet, pro_solicitudmat.numdoc, pro_producciondet.corr AS idprocorr, pro_producciondet.numparte " _
'        + vbCr + "FROM (alm_ingreso INNER JOIN pro_ordenproddet ON alm_ingreso.idorddet = pro_ordenproddet.id) INNER JOIN pro_producciondet ON pro_ordenproddet.idprocorr = pro_producciondet.corr " _
'        + vbCr + "WHERE (((alm_ingreso.id)=" & IDING_ & "));"
'
'    RST_Busq xRs, cSQL, xCon
'
'    If xRs.State = 0 Then Exit Sub
'    If xRs.RecordCount = 0 Then Exit Sub
'
'    GRID_.TextMatrix(FILA_, COLUMNA_) = NulosC(xRs("numparte"))
    GRID_.TextMatrix(FILA_, COLUMNA_) = ""
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraKardexProm
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EN FORMA DETALLADA TODOS LOS MOVIMIENTOS DEL ITEM SELECCIONADO, TAMBIEN
'*                    MUESTRA EL PRECIO PROMEDIO DE CADA OPERACION
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraKardexProm(IdItem As Double)
'--Modificado 18/02/11 Johan Castro
'--           Mostrar el saldo inicial o stock incial se tomara hasta un dia anterior a la fecha de inicio
'--           de la consulta, Si la fecha de inicio es el 01 de enero se tomara el stock incial de alm_inventario
'--Modificado 06/05/11 Johan Castro
'            Agregar parametro al evento, IdItem codigo del item de compra o venta
'            Agregar columna registro, numero de cuenta, nombre de cuenta
'            Modificar la presentacion en compras - nota credito antes ingreso ahora salida
'--Modificado 01/07/11 Johan Castro
'            Agregar variable xPrecioUni que almacenara el precio unitario x registro
'            Asignar a variable xPrecioUni el valor obtenido de evento PrecioUni() solo cuando ingreso sea (guia de remision, ventas)
'            y salida sea(almacen-compras), caso contrario tendra valor del recordset
'            Variable UltPreCosto se deja de usar momentaneamente dado que su valor es muy elevado en el calculo
'            Modificar subconsulta (AS, AI) campo "modulo" [iif(..='') a iif(..<>'0')]
'--Modificado 05/01/12 Enrique Pollongo
'            Agregar codigo para obtener el precio incial del item "xPrecioIni"
'--Creado 20/01/12 Johan Castro
'            Reemplazar Consulta que muestra detalle de kardex por evento KardexMovimientoSQL() que muestra lo mismo ademas se agrega lo siguiente.
'            Agregar filtro en almacen ingreso, salida "AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 "
'            Agregar filtro en produccion insumos solo se mostrara materia prima,lo demas sale de almacen "AND pro_producciondet.estado Not In (1,4,5) AND pro_producciondetins.canutil<>0 AND alm_inventario.tippro=3"
'            Agregar filtro en produccion productos terminado, solo mostrará los registros que esten aprobado o procesado "AND pro_producciondet.estado Not In (1,4,5)"

    DoEvents
    Dim xCadSQL As String
    Dim UltPreCosto As Double

    Dim mInicioGrupo As Long '--indica la fila inicial de un grupo, cambia cuando cambia de item
    Dim xPrecioUni As Double '--Indica el precio unitario de cada registro

    ' AI = Almacen Ingreso
    ' AS = Almacen Salida
    ' C =  Compras
    ' SM = SOLICUTID DE MATERIALES
    ' PP = PARTE DE PRODUCCION
    
    'GR = GUIAS DE REMISION
    'PS =
    
    TxtSaldo.Text = 0
    
''''    ' PREPARAMOS LA SELECT PARA ARMAR EL KARDEX
''''    xCadSQL = "SELECT alm_ingreso.id, alm_ingresodet.iditem, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc, " _
''''                & " alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, 'AI' AS tipo, alm_ingreso.nombre AS entidad, 0 AS aa, " _
''''                & " (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos ,'Almacén' & iif(cstr(numdocumentos) <>'0', ' - Compras','') as modulo, '' AS registro, '' AS ctanum, '' AS ctanom " _
''''        + vbCr + " FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) " _
''''                & " ON alm_ingreso.id = alm_ingresodet.id " _
''''        + vbCr + " WHERE (((alm_ingresodet.iditem)=" & IdItem & ") AND ((alm_ingreso.fching)>=CDate('" & TxtFchIni.Valor & "') " _
''''                & " And (alm_ingreso.fching)<=CDate('" & TxtFchFin.Valor & "')) AND ((alm_ingreso.tipmov)=-1)) " _
''''        + vbCr + " Union All " _
''''        + vbCr + " SELECT alm_ingreso.id, alm_ingresodet.iditem, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, [alm_ingreso]![numser]+'-'+[alm_ingreso]![numdoc] AS numdoc, " _
''''                & " alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, 'AS' AS tipo, alm_ingreso.nombre AS entidad, 0 AS aa, " _
''''                & " (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos ,'Almacén' & iif(cstr(numdocumentos) <>'0', ' - Compras','') as modulo, '' AS registro,'' AS ctanum, '' AS ctanom  " _
''''        + vbCr + " FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) " _
''''                & " ON alm_ingreso.id = alm_ingresodet.id  " _
''''        + vbCr + " WHERE (((alm_ingresodet.iditem)=" & IdItem & ") AND ((alm_ingreso.fching)>=CDate('" & TxtFchIni.Valor & "') " _
''''                & " And (alm_ingreso.fching)<=CDate('" & TxtFchFin.Valor & "')) AND ((alm_ingreso.tipmov)=0))" _
''''        + vbCr + " Union All " _
''''        + vbCr + " SELECT com_compras.id, com_comprasdet.iditem, alm_inventario.descripcion, com_compras.fchdoc, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, " _
''''                & " com_comprasdet.canpro, IIf([com_compras]![idmon]=2,[com_comprasdet]![preuni]*[con_tc]![impcom],[com_comprasdet]![preuni]) AS preuni, mae_documento.abrev AS descdoc, " _
''''                & " 'C' AS Tipo, mae_prov.nombre AS entidad, 0 AS aa, 0 AS numdocumentos,'Compras' as modulo,com_compras.numreg AS registro,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom " _
''''        + vbCr + " FROM (alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc)  " _
''''                & " LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
''''        + vbCr + " WHERE (((com_comprasdet.iditem)=" & IdItem & ") AND " _
''''                & " ((com_compras.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (com_compras.fchdoc)<=CDate('" & TxtFchFin.Valor & "')) AND ((com_compras.tipcom)=1))"
''''
''''    xCadSQL = xCadSQL _
''''        + vbCr + "  Union All" _
''''        + vbCr + " SELECT vta_guia.id, vta_guiadet.iditem, alm_inventario.descripcion, vta_guia.fecgiro, [vta_guia]![numser]+'-'+[vta_guia]![numdoc] AS numdoc, vta_guiadet.canpro, " _
''''                & " 0 AS preuni, mae_documento.abrev AS desdoc, 'GR' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, IIf([vta_guia]![iddocven]<>0,1,0) AS numdocumentos,'Guia de Remisión' as modulo, '' AS registro,'' AS ctanum, '' AS ctanom  " _
''''        + vbCr + " FROM ((mae_cliente RIGHT JOIN vta_guia ON mae_cliente.id = vta_guia.idcli) LEFT JOIN mae_documento ON vta_guia.tipdoc = mae_documento.id) LEFT JOIN (vta_guiadet " _
''''                & " LEFT JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id) ON vta_guia.id = vta_guiadet.idgui " _
''''        + vbCr + " WHERE (((vta_guiadet.iditem)=" & IdItem & ") " _
''''                & " AND ((vta_guia.fecgiro)>=CDate('" & TxtFchIni.Valor & "') And (vta_guia.fecgiro)<=CDate('" & TxtFchFin.Valor & "'))) " _
''''        + vbCr + " Union All " _
''''        + vbCr + " SELECT pro_produccion.id, pro_producciondetins.iditem, alm_inventario.descripcion, pro_produccion.dia, pro_producciondetins.numparte, pro_producciondetins.canutil, " _
''''                & " 0 AS preuni, 'SM' AS desdoc, 'PS' AS tipo, [alm_inventario_1].[descripcion] AS entidad, pro_producciondet.iditem AS aa, 0 AS numdocumentos,'Producción' as modulo, '' AS registro,'' AS ctanum, '' AS ctanom  " _
''''        + vbCr + " FROM (((pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN (pro_producciondetins LEFT JOIN alm_inventario ON pro_producciondetins.iditem = alm_inventario.id) ON (pro_producciondet.idrec = pro_producciondetins.idrec) " _
''''                & " AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro)) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_receta.iditem = alm_inventario_1.id " _
''''        + vbCr + " WHERE (((pro_producciondetins.iditem)=" & IdItem & ") AND ((pro_produccion.dia)>=CDate('" & TxtFchIni.Valor & "') And (pro_produccion.dia)<=CDate('" & TxtFchFin.Valor & "')))" _
''''        + vbCr + " Union All " _
''''        + vbCr & " SELECT pro_produccion.id, pro_producciondet.iditem, alm_inventario.descripcion, pro_produccion.dia, pro_producciondet.numparte, pro_producciondet.cantidad, " _
''''                & " 0 AS preuni, 'PP' AS desdoc, 'P' AS tipo, 'Producción' AS entidad, pro_producciondet.iditem AS aa, 0 AS numdocumentos ,'Producción' as modulo, '' as registro,'' AS ctanum, '' AS ctanom  " _
''''        + vbCr & " FROM pro_produccion INNER JOIN (pro_producciondet LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) ON pro_produccion.id = pro_producciondet.idpro " _
''''        + vbCr & " WHERE (((pro_producciondet.iditem)=" & IdItem & ") AND ((pro_produccion.dia)>=CDate('" & TxtFchIni.Valor & "') And (pro_produccion.dia)<=CDate('" & TxtFchFin.Valor & "'))) "
''''
''''    xCadSQL = xCadSQL + " UNION All " _
''''        + vbCr + " SELECT vta_ventas.id, vta_ventasdet.iditem, alm_inventario.descripcion, vta_ventas.fchdoc, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, " _
''''                    & " vta_ventasdet.canpro, vta_ventasdet.preuni, mae_documento.abrev AS descdoc, 'V' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, " _
''''                    & " 'Ventas' as modulo, vta_ventas.numreg AS registro,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom " _
''''        + vbCr + " FROM ((mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) RIGHT JOIN (vta_ventasdet  " _
''''                    & " LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
''''        + vbCr + " WHERE (((vta_ventasdet.iditem)=" & IdItem & ") " _
''''                    & " AND ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) " _
''''                    & " AND ((vta_ventas.iddocref) Is Null Or (vta_ventas.iddocref)=0) )" _
''''        + vbCr + " UNION All " _
''''        + vbCr + " SELECT vta_ventas.id, vta_ventasdet.iditem, alm_inventario.descripcion, vta_ventas.fchdoc, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, " _
''''                    & " vta_ventasdet.canpro, vta_ventasdet.preuni, mae_documento.abrev AS descdoc, 'V' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, " _
''''                    & " 'Ventas NC' AS modulo, vta_ventas.numreg AS registro,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom " _
''''        + vbCr + " FROM ((mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) RIGHT JOIN (vta_ventasdet " _
''''                    & " LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
''''        + vbCr + " WHERE (((vta_ventasdet.iditem)=" & IdItem & ") AND ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')) " _
''''                    & " AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref)<>0) AND ((vta_ventas.idmotnotcre)=4))"

    
    '--Generar la consulta SQL para obtener el detalle de movimientos del kardex
    xCadSQL = KardexMovimientoSQL(IdItem, 0, TxtFchIni.Valor, TxtFchFin.Valor)
        
    DoEvents
        
    RST_Busq rst, xCadSQL, xCon
    
    rst.Sort = "fchdoc, Tipo, numdoc"
        
    Dim A&
    Dim xFila As Integer
    Dim xTotSal, xTotEnt As Double
    
    '--agregar columna
    Fg1.Rows = Fg1.Rows + 1
    xFila = Fg1.Rows - 1
    
    '--indicando el inicio de grupo
    mInicioGrupo = xFila
        
    '--obtener el saldo inicial
    If CDate(TxtFchIni.Valor) <> CDate("01/01/" & AnoTra) Then
        StockIni = SaldoActual(IdItem, NulosC("01/01/" & AnoTra), NulosC(CDate(TxtFchIni.Valor) - 1), xCon)
    Else
        StockIni = NulosN(Busca_Codigo("id", NulosC(IdItem), "stckini", "alm_inventario", "N", xCon))
        xPrecioIni = NulosN(Busca_Codigo("id", NulosC(IdItem), "preini", "alm_inventario", "N", xCon))
    End If
        
    '--------------------------------------
    
    Fg1.TextMatrix(xFila, 3) = "Saldo Inicial"
    Fg1.TextMatrix(xFila, 4) = Format(StockIni, FORMAT_MONTO)
    Fg1.TextMatrix(xFila, 6) = Format(StockIni, FORMAT_MONTO)
    Fg1.TextMatrix(xFila, 7) = Format(xPrecioIni, "0.000000")
    Fg1.TextMatrix(xFila, 8) = Format(StockIni * xPrecioIni, FORMAT_MONTO)
    Fg1.TextMatrix(xFila, 10) = Format(StockIni * xPrecioIni, FORMAT_MONTO)
    
    If NulosN(Fg1.TextMatrix(xFila, 6)) <> 0 And NulosN(Fg1.TextMatrix(xFila, 10)) <> 0 Then
        Fg1.TextMatrix(xFila, 11) = NulosN(Fg1.TextMatrix(xFila, 10)) / NulosN(Fg1.TextMatrix(xFila, 6))
    End If
    
    Fg1.TextMatrix(xFila, 11) = Format(Fg1.TextMatrix(xFila, 11), "0.000000")
    
    UltPreCosto = NulosN(Fg1.TextMatrix(xFila, 11))
    
    ' Colocando el saldo inicial en stock actual
    TxtSaldo.Text = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 6))
        
    Dim xSaldo As Double
    Dim xSaldoImp As Double
        
    xSaldo = StockIni
    xSaldoImp = xSaldo * xPrecioIni
    xTotEnt = xTotEnt + StockIni
    
    '--agregando fila para proceder a ingresar los datos
    Fg1.Rows = Fg1.Rows + 1
    xFila = Fg1.Rows - 1
    
    If rst.RecordCount <> 0 Then
        rst.MoveFirst

        For A = 1 To rst.RecordCount
            'Fg1.Rows = Fg1.Rows + 1
            
            Fg1.TextMatrix(xFila, 1) = Format(rst("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(xFila, 2) = NulosC(rst("descdoc"))
            Fg1.TextMatrix(xFila, 3) = NulosC(rst("numdoc"))
            Fg1.TextMatrix(xFila, 12) = NulosC(rst("entidad"))
            Fg1.TextMatrix(xFila, 14) = NulosC(rst("modulo"))
            Fg1.TextMatrix(xFila, 15) = NulosC(rst("registro"))
            Fg1.TextMatrix(xFila, 16) = NulosC(rst("ctanum"))
            Fg1.TextMatrix(xFila, 17) = NulosC(rst("ctanom"))
            
'''            If Format(Rst("fchdoc"), "dd/mm/yy") = "07/02/09" Then
'''                MsgBox ""
'''            End If
            
            ' Si es un registro de Almacen
            ' Se halla su numero de Solicitud de Materiales
            ' y tambien su numero de Registro de Produccion
            If NulosC(rst("modulo")) = "Almacén" Then
                hallarNumProd NulosN(rst("id")), Fg1, xFila, 18
            End If

            If rst("tipo") = "C" Or rst("tipo") = "AI" Or rst("tipo") = "P" Then
                If rst("tipo") = "AI" Then
                    Fg1.TextMatrix(xFila, 13) = MostrarDocumentos(rst("id"), rst("tipo"))
                End If

                If rst("descdoc") = "NC" Then
                    Fg1.TextMatrix(xFila, 5) = Format(NulosN(rst("canpro")), FORMAT_MONTO)
                    xSaldo = xSaldo - NulosN(rst("canpro"))
                    xTotSal = xTotSal + NulosN(rst("canpro"))
                    
                Else
                    Fg1.TextMatrix(xFila, 4) = Format(NulosN(rst("canpro")), FORMAT_MONTO)
                    xSaldo = xSaldo + NulosN(rst("canpro"))
                    xTotEnt = xTotEnt + NulosN(rst("canpro"))
                    
                End If
                
                
                Fg1.TextMatrix(xFila, 6) = Format(xSaldo, FORMAT_MONTO)
                
                '--obtener el precio
                If rst("tipo") = "AI" And rst("numdocumentos") <> 0 Then
                    xPrecioUni = PrecioUni(rst("id"), IdItem, NulosC(rst("tipo")))
                Else
                    xPrecioUni = NulosN(rst("preuni"))
                End If
                
                Fg1.TextMatrix(xFila, 7) = Format(xPrecioUni, "0.000000")
                
                
                If rst("descdoc") = "NC" Then
                    Fg1.TextMatrix(xFila, 9) = Format(NulosN(rst("canpro")) * xPrecioUni, FORMAT_MONTO)
                    xSaldoImp = xSaldoImp - (NulosN(rst("canpro")) * xPrecioUni)
                    
                Else
                    Fg1.TextMatrix(xFila, 8) = Format(NulosN(rst("canpro")) * xPrecioUni, FORMAT_MONTO)
                    xSaldoImp = xSaldoImp + (NulosN(rst("canpro")) * xPrecioUni)
                    
                End If
                
                
                Fg1.TextMatrix(xFila, 10) = Format(xSaldoImp, "0.0000")
                
                
                If NulosN(Fg1.TextMatrix(xFila, 6)) <> 0 Then
                    Fg1.TextMatrix(xFila, 11) = Abs(NulosN(Fg1.TextMatrix(xFila, 10))) / NulosN(Fg1.TextMatrix(xFila, 6))
                End If
                
                Fg1.TextMatrix(xFila, 11) = Format(Fg1.TextMatrix(xFila, 11), "0.000000")
                
                UltPreCosto = NulosN(Fg1.TextMatrix(xFila, 11))
                
            Else
                If rst("tipo") = "GR" Then
                    Fg1.TextMatrix(xFila, 13) = MostrarDocumentos(rst("id"), rst("tipo"))
                End If
                
                If rst("descdoc") = "NC" Then
                    Fg1.TextMatrix(xFila, 4) = Format(NulosN(rst("canpro")), FORMAT_MONTO)
                    xSaldo = xSaldo + NulosN(rst("canpro"))
                    xTotEnt = xTotEnt + NulosN(rst("canpro"))
                Else
                    Fg1.TextMatrix(xFila, 5) = Format(NulosN(rst("canpro")), FORMAT_MONTO)
                    xSaldo = xSaldo - NulosN(rst("canpro"))
                    xTotSal = xTotSal + NulosN(rst("canpro"))
                End If
                
                '--saldo x cantidad
                Fg1.TextMatrix(xFila, 6) = Format(xSaldo, FORMAT_MONTO)
                
                '--obtener el precio
                If rst("tipo") = "GR" And rst("numdocumentos") <> 0 Then
                    xPrecioUni = PrecioUni(rst("id"), IdItem, NulosC(rst("tipo")))
                Else
'                    xPrecioUni = UltPreCosto
                    xPrecioUni = NulosN(rst("preuni"))
                    
                End If
                
                '--precio
                Fg1.TextMatrix(xFila, 7) = Format(xPrecioUni, "0.000000")
                
                If rst("descdoc") = "NC" Then
                    Fg1.TextMatrix(xFila, 8) = Format(NulosN(rst("canpro")) * xPrecioUni, FORMAT_MONTO)
                    xSaldoImp = xSaldoImp + (NulosN(rst("canpro")) * xPrecioUni)
                    
                Else
                    Fg1.TextMatrix(xFila, 9) = Format(NulosN(rst("canpro")) * xPrecioUni, FORMAT_MONTO)
                    xSaldoImp = xSaldoImp - (NulosN(rst("canpro")) * xPrecioUni)
                    
                End If
                '--saldo
                Fg1.TextMatrix(xFila, 10) = Format(xSaldoImp, "0.0000")
                
                
                If NulosN(Fg1.TextMatrix(xFila, 6)) <> 0 Then
                    Fg1.TextMatrix(xFila, 11) = Abs(NulosN(Fg1.TextMatrix(xFila, 10))) / NulosN(Fg1.TextMatrix(xFila, 6))
                    Fg1.TextMatrix(xFila, 11) = Format(Fg1.TextMatrix(xFila, 11), "0.000000")
                Else
                    Fg1.TextMatrix(xFila, 11) = "0.00"
                End If
                
'                If Rst("descdoc") = "NC" Then
'                    xTotEnt = xTotEnt + NulosN(Rst("canpro"))
'                Else
'                    xTotSal = xTotSal + NulosN(Rst("canpro"))
'                End If
            End If
           
            rst.MoveNext
            If rst.EOF = True Then
                Exit For
            End If
            Fg1.Rows = Fg1.Rows + 1
            xFila = xFila + 1
        Next A
        TxtSaldo.Text = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 6))
    End If
    
    ' Actualizar el saldo actual del producto
    xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckact = " & NulosN(TxtSaldo.Text) & " WHERE (((alm_inventario.id)=" & IdItem & "));"
    
    Fg1.Rows = Fg1.Rows + 1
    
'    Fg1.TextMatrix(Fg1.Rows - 1, 3) = "TOTALES ==>"
'    Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(xTotEnt, FORMAT_MONTO)
'    Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(xTotSal, FORMAT_MONTO)
    
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &HBB, True, , "TOTALES ==>"
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &HBB, True, , Format(GRID_SUMAR_COL(Fg1, 4, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &HBB, True, , Format(GRID_SUMAR_COL(Fg1, 5, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
    
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, &HBB, True, , Format(GRID_SUMAR_COL(Fg1, 8, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &HBB, True, , Format(GRID_SUMAR_COL(Fg1, 9, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
    
    
    
End Sub

'*****************************************************************************************************
'* Nombre           : Setea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CONFIGURA LAS COLUMNAS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Setea()
    'usamos la columna 19 para almacenar el destino de cada cuenta en la hoja de trabajo
    Fg1.Rows = 2
    Fg1.TextMatrix(0, 0) = "          "
    Fg1.TextMatrix(1, 0) = "          "
    Fg1.TextMatrix(0, 1) = "Fecha"
    Fg1.TextMatrix(1, 1) = "Fecha"
    Fg1.TextMatrix(0, 2) = "TD"
    Fg1.TextMatrix(1, 2) = "TD"
    Fg1.TextMatrix(0, 3) = "Nº Documento"
    Fg1.TextMatrix(1, 3) = "Nº Documento"
    Fg1.TextMatrix(0, 7) = "P. Unitario"
    Fg1.TextMatrix(1, 7) = "P. Unitario"
    Fg1.TextMatrix(0, 11) = "P. Promedio"
    Fg1.TextMatrix(1, 11) = "P. Promedio"
    
    Fg1.Redraw = False
    Fg1.MergeCol(0) = True
    Fg1.MergeCol(1) = True
    Fg1.MergeCol(2) = True
    Fg1.MergeCol(3) = True
    Fg1.MergeCol(7) = True
    Fg1.MergeCol(11) = True
    
    Fg1.MergeCells = 2
    Fg1.Redraw = True
    
    With Fg1
        .MergeCells = flexMergeFree
        .MergeRow(-1) = True
        .Cell(flexcpText, 0, 4, 0, 6) = "Unidades"
        .Cell(flexcpText, 0, 8, 0, 10) = "Importes"
        .Cell(flexcpBackColor, 0, 0, Fg1.Rows - 1, Fg1.Cols - 1) = &H8000000F
    End With
   
    Fg1.ColWidth(4) = 900
    Fg1.ColWidth(5) = 900
    Fg1.ColWidth(6) = 900
    
    If MostrarValorizado = True Then
        Fg1.ColWidth(7) = 900
        Fg1.ColWidth(8) = 900
        Fg1.ColWidth(9) = 900
        Fg1.ColWidth(10) = 1000
        Fg1.ColWidth(11) = 900
    Else
        Fg1.ColWidth(7) = 0
        Fg1.ColWidth(8) = 0
        Fg1.ColWidth(9) = 0
        Fg1.ColWidth(10) = 0
        Fg1.ColWidth(11) = 0
    End If
        
    Fg1.TextMatrix(1, 4) = "Entradas"
    Fg1.TextMatrix(1, 5) = "Salidas"
    Fg1.TextMatrix(1, 6) = "Saldo"
    Fg1.TextMatrix(1, 8) = "Entradas"
    Fg1.TextMatrix(1, 9) = "Salidas"
    Fg1.TextMatrix(1, 10) = "Saldo"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then

        If OptVal1.Value = True Then
            
            If fValidarDatos() = False Then Exit Sub
            
            Fg1.Rows = 2
    
            MuestraKardexProm NulosN(LblIdProducto.Caption)
            
        End If
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    TxtFchIni.Valor = CDate("01/01/" & Year(Date))
    TxtFchFin.Valor = Date
    LblDescripcion.Caption = ""
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
End Sub



Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height > 3000 Then
        Fg1.Top = 1485
        Fg1.Width = Me.Width - 150
        Fg1.Height = Me.Height - 2200 '--165
        
        LblDescripcion.Top = Me.Height - 700
        LblDescripcion.Width = Me.Width - 150
        
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        If OptVal1.Value = True Then
            
            If fValidarDatos() = False Then Exit Sub
            
            Fg1.Rows = 2
    
            MuestraKardexProm NulosN(LblIdProducto.Caption)
            
        End If
    End If
    
    If Button.Index = 2 Then pExportar 'ExportarExcel
    
    If Button.Index = 3 Then pImprimir
    
    If Button.Index = 5 Then
        Set rst = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ExportarExcel()
    Dim A&
    Dim B&
    Dim xFilas&
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    ' determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    ' abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Add
    
    Frame4.Left = 2940
    Frame4.Top = 1890
    Label3.Caption = "Exportando Kardex"
    Frame4.Visible = True
    
    ProgressBar1.Max = Fg1.Rows - 1
    
    With objExcel.ActiveSheet
        
        .Cells(1, 2) = NomEmp
        .Cells(1, 13) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        
        xFilas = 7
        For B = 1 To Fg1.Cols - 1
            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(0, B)
        Next B
        
        xFilas = xFilas + 1
        For B = 1 To Fg1.Cols - 1
            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(1, B)
        Next B
        
        DoEvents
        xFilas = xFilas + 1
        For A = 2 To Fg1.Rows - 1
            ProgressBar1.Value = A
            DoEvents
            For B = 1 To Fg1.Cols - 1
                If B = 1 Or B = 2 Or B = 3 Or B = 12 Then
                    .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                Else
                    .Cells(xFilas, B + 1) = NulosN(Fg1.TextMatrix(A, B))
                End If
            Next B
            xFilas = xFilas + 1
        Next A
    End With
    
    Frame4.Visible = False
    MsgBox "El proceso de exportación terminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
End Sub

Private Sub txtCodItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then CmdProducto_Click
End Sub

Private Sub txtCodItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtSaldo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtUnidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pImprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MANDAN A IMPRIMIR LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pImprimir()
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay registros para imprimir", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    FrmPrintKardex.Cargar2
    Me.MousePointer = vbDefault
    FrmPrintKardex.Show


'    If fValidarDatos() = False Then Exit Sub
'
'    On Error GoTo error
'
'    Dim oPrint  As New SGI2_funciones.formularios
'    Dim mIndex As Integer
'    Dim nTitulo As String
'    Dim nPeriodo As String
'    Dim nTitulo1 As String
'    If MostrarValorizado = False Then
'        nTitulo = "Consulta Kardex"
'    Else
'        nTitulo = "Consulta de Kardex Valorizado"
'    End If
'
'    If CDate(TxtFchIni.Valor) = CDate(TxtFchIni.Valor) Then
'        nPeriodo = "Periodo Al " & CDate(TxtFchIni.Valor)
'    Else
'        nPeriodo = "Periodo Del " & CDate(TxtFchIni.Valor) & " Al " & CDate(TxtFchIni.Valor)
'    End If
'
'    If NulosN(LblIdProducto.Caption) <> 0 Then
'        nTitulo1 = IIf(Opt1.Value = True, "Producto", IIf(Opt2.Value = True, "Insumo", "Mercadería")) & " " & StrConv(TxtDesc.Text, 3)
'    End If
'
'    Me.MousePointer = vbHourglass
'    oPrint.Imprimir_x_VSFlexGrid Fg1, nTitulo, nTitulo1, nPeriodo, False, True
'    Set oPrint = Nothing
'    Me.MousePointer = vbDefault
'    Exit Sub
'error:
'    Set oPrint = Nothing
'    Me.MousePointer = vbDefault
'    SHOW_ERROR Me.Name, "Exportar"
End Sub

'*****************************************************************************************************
'* Nombre           : pExportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A EXCEL LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    If fValidarDatos() = False Then Exit Sub
    
    On Error GoTo error
    
    Dim oExport As New SGI2_funciones.formularios
    Dim mIndex As Integer
    Dim nTitulo As String
    Dim nPeriodo As String
    Dim nTitulo1 As String
    If MostrarValorizado = False Then
        nTitulo = "Consulta Kardex"
    Else
        nTitulo = "Consulta de Kardex Valorizado"
    End If
        
    If CDate(TxtFchIni.Valor) = CDate(TxtFchIni.Valor) Then
        nPeriodo = "Periodo Al " & CDate(TxtFchIni.Valor)
    Else
        nPeriodo = "Periodo Del " & CDate(TxtFchIni.Valor) & " Al " & CDate(TxtFchIni.Valor)
    End If
    
    If NulosN(LblIdProducto.Caption) <> 0 Then
        nTitulo1 = IIf(Opt1.Value = True, "Producto", IIf(Opt2.Value = True, "Insumo", "Mercadería")) & " " & StrConv(TxtDesc.Text, 3)
    End If

    Me.MousePointer = vbHourglass
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, nTitulo, nPeriodo, nTitulo1, "Consulta de Kardex"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VERIFICA QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    If Opt1.Value = False And Opt2.Value = False And Opt3.Value = False Then
        MsgBox "Seleccione el Tipo de Búsqueda" & vbCr & "Productos, Insumos, Mercadería", vbExclamation, xTitulo
        Exit Function
    End If
    
    ' si selecciono algun registro
    If NulosN(LblIdProducto.Caption) = 0 Then
        MsgBox "Seleccione " + IIf(Opt1.Value = True, "Producto", IIf(Opt2.Value = True, "Insumo", "Mercadería")), vbExclamation, xTitulo
        txtCodItem.SetFocus
        Exit Function
    End If
    
    ' si esta la fecha correcta
    If IsDate(TxtFchIni.Valor) = False Then
        MsgBox "Ingrese la Fecha de Inicio", vbExclamation, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    ElseIf IsDate(TxtFchFin.Valor) = False Then
        MsgBox "Ingrese la Fecha Final", vbExclamation, xTitulo
        TxtFchFin.SetFocus
        Exit Function
    ElseIf CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha Inicial es superior al Final" & vbCr & "Modifique el Intervalo de Fechas", vbExclamation, xTitulo
        TxtFchFin.SetFocus
        Exit Function
    End If
    fValidarDatos = True
End Function


Private Sub CmdTodos_Click()
    '===================================================================================================
    'Creado : 20/03/10 Por: Johan Castro
    'Propósito: Mostrar todo el kardex en un solo listado
    '
    'Entradas:  Ninguna
    '
    'Resultados: Lista del Kardex en pantalla
    '
    'Nota:       1.- Seleccionar el tipo de listado(Producto,Insumos o Mercaderia)
    '            2.- Clic en boton Ver Todos
    'Modificado 06/05/11 Johan Castro
    '           Eliminar codigo redundante para mostrar los datos del detalle,
    '           este evento invocara el evento MuestraKardexProm()
    '           Eliminar variables que no se usan
    '===================================================================================================


    Dim xRs As New ADODB.Recordset
    Dim nSQL As String

    'Dim xCadSQL As String
    'Dim UltPreCosto As Double

    'Dim A&
    'Dim xFila&
    'Dim xTotSal, xTotEnt As Double
    
    'Dim xSaldo As Double
    'Dim xSaldoImp As Double
        
    Dim xTCDia As Double

    'Dim mInicioGrupo As Long '--indica la fila inicial de un grupo, cambia cuando cambia de item
    
    '--cargamos todos los registros segun seleccion de usuario
    If Opt1.Value = True Then
        ' buscamos producto
        nSQL = "SELECT alm_inventario.*, mae_unidades.abrev FROM mae_unidades RIGHT JOIN alm_inventario ON " _
            & " mae_unidades.id = alm_inventario.idunimed Where  ((alm_inventario.tippro) = 3)  and alm_inventario.activo =-1 " _
            & " ORDER BY alm_inventario.descripcion"
    End If
    If Opt2.Value = True Then
        ' buscamos materia prima /insumos
        nSQL = "SELECT alm_inventario.*, mae_unidades.abrev FROM mae_unidades RIGHT JOIN alm_inventario ON " _
            & " mae_unidades.id = alm_inventario.idunimed Where ((alm_inventario.tippro) = 1  and alm_inventario.activo =-1 " _
            & " Or (alm_inventario.tippro) = 4) ORDER BY alm_inventario.descripcion"
    End If
    If Opt3.Value = True Then
        ' buscamos mercaderias
        nSQL = "SELECT alm_inventario.*, mae_unidades.abrev FROM mae_unidades RIGHT JOIN alm_inventario ON " _
            & " mae_unidades.id = alm_inventario.idunimed Where ((alm_inventario.tippro) = 2)   and alm_inventario.activo =-1 " _
            & " ORDER BY alm_inventario.descripcion"
    End If
    
   
    '--cargar los datos de la consulta
    RST_Busq xRs, nSQL, xCon
    
    If xRs.RecordCount <> 0 Then
        ProgressBar1.Max = xRs.RecordCount
        ProgressBar1.Min = 0
        ProgressBar1.Value = 0
        xRs.MoveFirst
    End If
    
    
    '--mostrando la barra de progreso
    Frame4.Left = 3170
    Frame4.Top = 3390
    Frame4.Visible = True
    
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    Fg1.Rows = Fg1.Rows + 1
    
    Do While Not xRs.EOF
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        '--separar item
        If xRs.Bookmark <> 1 Then Fg1.Rows = Fg1.Rows + 2
        
        '--mostrando datos del producto
        LblIdProducto.Caption = xRs("id")
        txtCodItem.Text = NulosC(xRs("codpro"))
        TxtDesc.Text = NulosC(xRs("descripcion"))
        TxtUnidad.Text = NulosC(xRs("abrev"))
        TxtSaldo.Text = Format(NulosN(xRs("stckact")), "0.00")
        StockIni = NulosN(xRs("stckini"))
        
        DoEvents
        
        If xRs("idmon") = 2 Then
            ' hallamos el precio inicial en funcion al precio inicial del producto
            xTCDia = HallaTipoCambio(TxtFchIni.Valor, "2", Venta, xCon)
            xPrecioIni = NulosN(xRs("preini") * xTCDia)
        Else
            xPrecioIni = NulosN(xRs("preini"))
        End If
        
        TxtSaldo.Text = 0

'        '--nombre del item
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 1, &H840000, True, , txtCodItem.Text
        GRID_COMBINAR Fg1, Fg1.Rows - 1, 2, Fg1.Rows - 1, 7, TxtDesc.Text, flexAlignLeftCenter, True, flexMergeFree, &H840000, , True
        
        MuestraKardexProm xRs("id")
    
        xRs.MoveNext
        
    Loop
    
    Set xRs = Nothing
    
    Frame4.Visible = False
    
End Sub



Private Function PrecioUni(IdDocumento, IdItem As Double, DondeBuscar As String) As Double
    '===================================================================================================
    'Creado:     01/07/11 Johan Castro
    'Propósito:  Obtener el Precio unitario del registro de compras vinculado con documentos (de ingreso de almacen, Guia Remision)
    '
    'Entradas:   IdDocumento = Código de Libro
    '            IdItem = Código del Item (Producto, Materia prima, Insumo, etc)
    '            DondeBuscar = Indica el origen del registro
    '
    'Resultados: Precio unitario del item segun el documento ingresado
    '===================================================================================================
    
    Dim xRst As New ADODB.Recordset
    Dim nSQL As String
    
    If DondeBuscar = "AI" Then
        nSQL = "SELECT Avg(com_comprasdet.preuni) AS preuniprom " _
            + vbCr + " FROM com_comprasdet INNER JOIN alm_ingresodoc ON com_comprasdet.idcom = alm_ingresodoc.iddoc " _
            + vbCr + " GROUP BY alm_ingresodoc.id, com_comprasdet.iditem " _
            + vbCr + " HAVING (((alm_ingresodoc.id)=" & IdDocumento & ") AND ((com_comprasdet.iditem)=" & IdItem & "))"
    ElseIf DondeBuscar = "GR" Then
        nSQL = "SELECT vta_guia.id, vta_ventasdet.iditem, Avg(vta_ventasdet.preuni) AS preuniprom " _
            + vbCr + " FROM vta_guia INNER JOIN vta_ventasdet ON vta_guia.iddocven = vta_ventasdet.idvta " _
            + vbCr + " GROUP BY vta_guia.id, vta_ventasdet.iditem " _
            + vbCr + " HAVING (((vta_guia.id)=" & IdDocumento & ") AND ((vta_ventasdet.iditem)=" & IdItem & ")); "
    Else
        PrecioUni = 0
        Exit Function
    End If
    
    RST_Busq xRst, nSQL, xCon
    
    If rst.State = 1 Then
        If xRst.RecordCount <> 0 Then
            PrecioUni = NulosN(xRst("preuniprom"))
        Else
            PrecioUni = 0
        End If
    Else
        PrecioUni = 0
    End If
    
    Set xRst = Nothing
    
End Function
