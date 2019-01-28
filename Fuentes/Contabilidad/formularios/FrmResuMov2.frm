VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmResuMov2 
   Caption         =   "Contabilidad - Kardex"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "[ Opciones ]"
      Height          =   1000
      Left            =   9720
      TabIndex        =   17
      Top             =   360
      Width           =   2535
      Begin VB.CommandButton cmd 
         Caption         =   "Ver Detallado"
         Height          =   465
         Index           =   0
         Left            =   120
         Picture         =   "FrmResuMov2.frx":0000
         TabIndex        =   18
         ToolTipText     =   "Kardex"
         Top             =   360
         Width           =   2265
      End
   End
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   3390
      TabIndex        =   8
      Top             =   4410
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   9
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   720
         Width           =   1260
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Seleccionar ]"
      Height          =   1005
      Left            =   30
      TabIndex        =   2
      Top             =   350
      Width           =   4525
      Begin VB.CommandButton cmd 
         Height          =   240
         Index           =   1
         Left            =   1395
         Picture         =   "FrmResuMov2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   285
         Width           =   240
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   750
         TabIndex        =   5
         Top             =   600
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
         Left            =   3000
         TabIndex        =   6
         Top             =   600
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
      Begin VB.TextBox IdAlmacenText 
         Height          =   300
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "IdAlmacenText"
         Top             =   255
         Width           =   915
      End
      Begin VB.Label AlmacenLabel 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AlmacenLabel"
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
         Left            =   1710
         TabIndex        =   16
         Top             =   255
         Width           =   2725
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   15
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2400
         TabIndex        =   4
         Top             =   660
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   660
         Width           =   420
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
            Picture         =   "FrmResuMov2.frx":043C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":0980
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":0D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":0E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":11FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":1382
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":17D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":18EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":1E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":2376
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":248A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":259E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":29F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResuMov2.frx":2B5E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   609
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
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6135
      Left            =   30
      TabIndex        =   0
      Top             =   1440
      Width           =   12330
      _cx             =   21749
      _cy             =   10821
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
      Rows            =   3
      Cols            =   24
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmResuMov2.frx":30A6
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
   Begin VSFlex7Ctl.VSFlexGrid fg 
      Height          =   940
      Index           =   1
      Left            =   4600
      TabIndex        =   7
      ToolTipText     =   "Buscar Linea"
      Top             =   420
      Width           =   5105
      _cx             =   9005
      _cy             =   1658
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
      FormatString    =   $"FrmResuMov2.frx":3338
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
Attribute VB_Name = "FrmResuMov2"
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
Dim cSQL As String
Dim INDICE_ As Integer
Dim BAND_INTERRUMPIR As Boolean
Dim F As New SistemaLogica.Funciones

Private Enum COLUMNARESUMIDO_
    COLUMNATIPO_ = 1
    COLUMNACODIGO_
    COLUMNADESCRIPCION_
    COLUMNAUNIMED_
    COLUMNASTOCKINI_
    COLUMNAPRECIOINI_
    COLUMNAENTRADACANTIDAD_
    COLUMNAENTRADAIMPORTE_
    COLUMNASALIDACANTIDAD_
    COLUMNASALIDAIMPORTE_
    COLUMNASALDOCANTIDAD_
    COLUMNASALDOIMPORTE_
    COLUMNAPRECIOPROM_
    COLUMNAIDITEM_
    COLUMNAIDTIPPRO_
End Enum

Private Sub pIniciarCampos()
    TxtFchIni.Valor = CDate("01/01/" & Year(Date))
    TxtFchFin.Valor = Date
    BAND_INTERRUMPIR = False
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    GRID_COMBOLIST fg(1), 2
    fg(1).ColWidth(1) = 0
    
    Fg1.Rows = Fg1.FixedRows
    
    Blanquea
    pConfigurarGrid
End Sub

Private Sub pConfigurarGrid()
    GRID_COMBINAR Fg1, 0, COLUMNATIPO_, 0, COLUMNAUNIMED_, "DETALLES DEL ITEM", , , , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNATIPO_, 1, COLUMNATIPO_, "TIPO", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNACODIGO_, 1, COLUMNACODIGO_, "CODIGO", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNADESCRIPCION_, 1, COLUMNADESCRIPCION_, "DESCRIPCIÓN", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAUNIMED_, 1, COLUMNAUNIMED_, "U.M.", , False, , , &H8000000F, False
    
    GRID_COMBINAR Fg1, 0, COLUMNASTOCKINI_, 1, COLUMNASTOCKINI_, "STCK. INI.", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, COLUMNAPRECIOINI_, 1, COLUMNAPRECIOINI_, "PRECIO INI.", , False, , , &H8000000F, False
    
    GRID_COMBINAR Fg1, 0, COLUMNAENTRADACANTIDAD_, 0, COLUMNAENTRADAIMPORTE_, "ENTRADAS", , , , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAENTRADACANTIDAD_, 1, COLUMNAENTRADACANTIDAD_, "CANTIDAD", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAENTRADAIMPORTE_, 1, COLUMNAENTRADAIMPORTE_, "IMPORTE", , False, , , &H8000000F, False
    
    GRID_COMBINAR Fg1, 0, COLUMNASALIDACANTIDAD_, 0, COLUMNASALIDAIMPORTE_, "SALIDAS", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNASALIDACANTIDAD_, 1, COLUMNASALIDACANTIDAD_, "CANTIDAD", , True, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNASALIDAIMPORTE_, 1, COLUMNASALIDAIMPORTE_, "IMPORTE", , True, , , &H8000000F, False
    
    GRID_COMBINAR Fg1, 0, COLUMNASALDOCANTIDAD_, 0, COLUMNASALDOIMPORTE_, "SALDOS", , , , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNASALDOCANTIDAD_, 1, COLUMNASALDOCANTIDAD_, "CANTIDAD", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNASALDOIMPORTE_, 1, COLUMNASALDOIMPORTE_, "IMPORTE", , False, , , &H8000000F, False
        
    GRID_COMBINAR Fg1, 0, COLUMNAPRECIOPROM_, 1, COLUMNAPRECIOPROM_, "PRECIO PROM.", , False, , , &H8000000F, False
        
    Fg1.RowHeight(0) = 300
    Fg1.RowHeight(1) = 300
    Fg1.WordWrap = True
    ' -------------------------------------------DETALLES DE ITEM
    ' TIPO DE ITEM
    Fg1.ColWidth(COLUMNATIPO_) = 1500
    ' CODIGO DE ITEM
    Fg1.ColWidth(COLUMNACODIGO_) = 1900
    ' DESCRIPCION DE ITEM
    Fg1.ColWidth(COLUMNADESCRIPCION_) = 4500
    ' UNIDAD DE MEDIDA
    Fg1.ColWidth(COLUMNAUNIMED_) = 500
    ' -------------------------------------------STOCK INICIAL
    Fg1.ColWidth(COLUMNASTOCKINI_) = 900
    ' -------------------------------------------PRECIO INICIAL
    Fg1.ColWidth(COLUMNAPRECIOINI_) = 900
    ' -------------------------------------------ENTRADAS
    ' CANTIDAD
    Fg1.ColWidth(COLUMNAENTRADACANTIDAD_) = 1100
    ' COSTO TOTAL
    Fg1.ColWidth(COLUMNAENTRADAIMPORTE_) = 1100
    ' -------------------------------------------SALIDAS
    ' CANTIDAD
    Fg1.ColWidth(COLUMNASALIDACANTIDAD_) = 1100
    ' COSTO TOTAL
    Fg1.ColWidth(COLUMNASALIDAIMPORTE_) = 1100
    ' -------------------------------------------SALDOS
    ' CANTIDAD
    Fg1.ColWidth(COLUMNASALDOCANTIDAD_) = 1100
    ' IMPORTE
    Fg1.ColWidth(COLUMNASALDOIMPORTE_) = 1100
    ' -------------------------------------------PRECIO INICIAL
    Fg1.ColWidth(COLUMNAPRECIOPROM_) = 900
    
    OCULTAR_COL Fg1, 14, 23
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TextBox PARA EL INGRESO DE DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    IdAlmacenText.Text = ""
    AlmacenLabel.Caption = ""
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim xCampos() As String
    Dim cSQL As String
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    
    Select Case Index
        Case 0 ' VER KARDEX DETALLADO
            If Fg1.Rows = Fg1.FixedRows Then
                FrmVerKardex2.Show
                Exit Sub
            End If
            
            Unload FrmVerKardex2
            FrmVerKardex2.txtIdTipPro.Text = NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAIDTIPPRO_))
            FrmVerKardex2.lblTipPro.Caption = NulosC(Fg1.TextMatrix(Fg1.Row, COLUMNATIPO_))
            FrmVerKardex2.IdItemLabel.Caption = NulosN(Fg1.TextMatrix(Fg1.Row, COLUMNAIDITEM_))
            FrmVerKardex2.txtIdItem.Text = NulosC(Fg1.TextMatrix(Fg1.Row, COLUMNACODIGO_))
            FrmVerKardex2.lblItem.Caption = NulosC(Fg1.TextMatrix(Fg1.Row, COLUMNADESCRIPCION_))
            FrmVerKardex2.lblItem.ToolTipText = NulosC(Fg1.TextMatrix(Fg1.Row, COLUMNADESCRIPCION_))
            FrmVerKardex2.TxtFchIni.Valor = TxtFchIni.Valor
            FrmVerKardex2.TxtFchFin.Valor = TxtFchFin.Valor
            FrmVerKardex2.IdAlmacenText.Text = IdAlmacenText.Text
            FrmVerKardex2.AlmacenLabel.Caption = AlmacenLabel.Caption
            FrmVerKardex2.Fg1.Rows = FrmVerKardex2.Fg1.FixedRows
            FrmVerKardex2.Show
        
        Case 1 ' Almacen
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
            
            nTitulo = "Buscando Almacenes"
            cSQL = "SELECT alm_almacenes.* FROM alm_almacenes"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IdAlmacenText.Text = NulosN(xRs("id"))
            AlmacenLabel.Caption = UCase(NulosC(xRs("descripcion")))
            TxtFchIni.SetFocus
            Set xRs = Nothing
    End Select
End Sub

Private Sub fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    Dim nTitulo As String
    
    Select Case Index
        Case 0 ' TIPO DE ITEM
            
        Case 1 ' ITEM
            ReDim xCampos(2, 4) As String
            xCampos(0, 0) = "Código":       xCampos(0, 1) = "codpro":       xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "4000":    xCampos(1, 3) = "C"
            
            'nSQLId = GENERAR_SQL_ID(fg(0), 1, " AND alm_inventario.tippro", "IN", True)
            nSQLId = nSQLId & GENERAR_SQL_ID(fg(Index), 1, " AND alm_inventario.id", "NOT IN", True)
                
            cSQL = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion " _
                + vbCr + "From alm_inventario " _
                + vbCr + "WHERE (((alm_inventario.activo)=-1) AND alm_inventario.tippro <> 5)" & nSQLId
                             
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "codpro", "codpro", Principio
    
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            fg(Index).TextMatrix(fg(Index).Row, fg(Index).ColIndex("IDITEM")) = F.NuloNumeric(xRs("id"))
            fg(Index).TextMatrix(fg(Index).Row, fg(Index).ColIndex("CODPRO")) = F.NuloString(xRs("codpro"))
            fg(Index).TextMatrix(fg(Index).Row, fg(Index).ColIndex("ITEM")) = F.NuloString(xRs("descripcion"))
            fg(Index).Rows = fg(Index).Rows + 1
            fg(Index).TopRow = fg(Index).Rows - 1
            fg(Index).Row = fg(Index).Rows - 1
            
    End Select
End Sub

Private Sub Fg1_DblClick()
    cmd_Click 0
End Sub

Private Sub menu00_Click()
    If fg(INDICE_).Rows > 2 Then fg(INDICE_).TopRow = fg(INDICE_).Rows - 2
    fg_CellButtonClick INDICE_, fg(INDICE_).Rows - 1, 1
End Sub

Private Sub menu01_Click()
    If fg(INDICE_).Row < fg(INDICE_).FixedRows Then Exit Sub
    fg(INDICE_).RemoveItem fg(INDICE_).Row
    
    If fg(INDICE_).Rows = fg(INDICE_).FixedRows Then fg(INDICE_).Rows = fg(INDICE_).Rows + 1
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    INDICE_ = Index
    If Button <> 2 Then Exit Sub
    Select Case Index
        Case 0, 1
            PopupMenu menu
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 0, 1
            If KeyCode = vbKeyInsert Then ' Agregar
                menu00_Click
            End If
            
            If KeyCode = vbKeyDelete Then ' Eliminar
                menu01_Click
            End If
        Case Else
            Exit Sub
    End Select

End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        If MostrarValorizado = False Then
            Me.Caption = "Almacén - Kardex"
        Else
            Me.Caption = "Contabilidad - Kardex Valorizado"
        End If
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
    Dim xRs As New ADODB.Recordset
    Dim cSQL As String

    cSQL = "SELECT alm_ingreso.id, alm_ingreso.idorddet, pro_ordenproddet.numdoc, pro_producciondet.corr AS idprocorr, pro_producciondet.numparte " _
        + vbCr + "FROM (alm_ingreso INNER JOIN pro_ordenproddet ON alm_ingreso.idorddet = pro_ordenproddet.id) INNER JOIN pro_producciondet ON pro_ordenproddet.idprocorr = pro_producciondet.corr " _
        + vbCr + "WHERE (((alm_ingreso.id)=" & IDING_ & "));"
        
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    GRID_.TextMatrix(FILA_, COLUMNA_) = NulosC(xRs("numparte"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        If fValidarDatos() = False Then Exit Sub
        Fg1.Rows = Fg1.FixedRows
        pCargarDatos
    End If
    
    If KeyCode = vbKeyEscape Then
        '--interrumpir
        BAND_INTERRUMPIR = True
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    pIniciarCampos
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    Fg1.Top = 1485
    Fg1.Width = Me.Width - 315
    Fg1.Height = Me.Height - 2200
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pCargarDatos
    
    If Button.Index = 2 Then ExportarExcel Fg1
    
    If Button.Index = 3 Then pImprimir
    
    If Button.Index = 5 Then
        Set rst = Nothing
        Unload Me
    End If
End Sub

Private Sub pCargarDatos()
    If fValidarDatos() = False Then Exit Sub
    pCargarResumido
End Sub

Private Sub pCargarResumido()
    Dim xRs As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim A&, B&, C&
    Dim xTotal As Double
    Dim xCad As String
    Dim nSQLId As String       ' Para mostrar solo items activos
    Dim mCad_IdItem_In As String
    Dim StockIni As Double          '--stock incial, depende de la fecha de inicio de consulta
    Dim PrecioIni As Double
    Dim xPrecioPromedio As Double
    Dim INGRESOCANTIDAD_ As Double
    Dim INGRESOIMPORTE_ As Double
    Dim SALIDACANTIDAD_ As Double
    Dim SALIDAIMPORTE_ As Double
    Dim PRECIOPROMEDIO_ As Double
    Dim IdAlmacen As Long
    Dim FchInicio As Date
    Dim FchFinal As Date
    Dim QUERYA As String
    Dim QUERYB As String
    Dim QUERYC As String
    
    
    Me.MousePointer = vbHourglass
    ' FILTROS DE BUSQUEDA
    nSQLId = GENERAR_SQL_ID(fg(1), fg(1).ColIndex("IDITEM"), " M.iditem")
    If nSQLId <> "" Then nSQLId = vbCr + "WHERE " & nSQLId
    mCad_IdItem_In = F.GenerarSQLInGrid(fg(1), fg(1).ColIndex("IDITEM"), "", False)
        
    BAND_INTERRUMPIR = False
    IdAlmacen = F.NuloNumeric(IdAlmacenText.Text)
    FchInicio = TxtFchIni.Valor
    FchFinal = TxtFchFin.Valor
               
    cSQL = F.SQL_MovTotalizado(mCad_IdItem_In, IdAlmacen, FchInicio, FchFinal, xCon, True, True)
           
    RST_Busq xRs, cSQL, xCon
    
    Fg1.Rows = Fg1.FixedRows
    If xRs.RecordCount = 0 Then Me.MousePointer = vbDefault: Exit Sub
    xRs.MoveFirst
    
    Me.MousePointer = vbDefault
    CentrarFrm FraProgreso
    FraProgreso.Visible = True
    Fg1.Rows = Fg1.FixedRows
    Fg1.MergeCells = flexMergeFixedOnly
    PgBar.Min = 0
    PgBar.Max = xRs.RecordCount
    PgBar.Value = 0
    
    While Not xRs.EOF
        With Fg1
            xPrecioPromedio = 0
            '--Validar la interrupcion de la consulta
            If BAND_INTERRUMPIR = True Then GoTo xSalir
            PgBar.Value = PgBar.Value + 1
            .Rows = .Rows + 1
            LblProg.Caption = NulosC(xRs("item"))
            
            .TextMatrix(.Rows - 1, COLUMNATIPO_) = UCase(NulosC(xRs("tippro")))
            .TextMatrix(.Rows - 1, COLUMNACODIGO_) = NulosC(xRs("coditem"))
            .TextMatrix(.Rows - 1, COLUMNADESCRIPCION_) = NulosC(xRs("item"))
            .TextMatrix(.Rows - 1, COLUMNAUNIMED_) = NulosC(xRs("unimed"))
            .TextMatrix(.Rows - 1, COLUMNASTOCKINI_) = Format(F.NuloNumeric(xRs("canini")), "0.00")
            .TextMatrix(.Rows - 1, COLUMNAPRECIOINI_) = Format(F.NuloNumeric(xRs("costoini")), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNAIDITEM_) = NulosN(xRs("iditem"))
            .TextMatrix(.Rows - 1, COLUMNAIDTIPPRO_) = NulosN(xRs("idtippro"))
            .TextMatrix(.Rows - 1, COLUMNAENTRADACANTIDAD_) = Format(F.NuloNumeric(xRs("canent")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, COLUMNASALIDACANTIDAD_) = Format(F.NuloNumeric(xRs("cansal")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, COLUMNASALDOCANTIDAD_) = F.NuloNumeric(xRs("canini")) + F.NuloNumeric(xRs("canent")) - F.NuloNumeric(xRs("cansal"))
            .TextMatrix(.Rows - 1, COLUMNASALDOCANTIDAD_) = Format(.TextMatrix(.Rows - 1, COLUMNASALDOCANTIDAD_), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNAENTRADAIMPORTE_) = Format(F.NuloNumeric(xRs("costoent")), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNASALIDAIMPORTE_) = Format(F.NuloNumeric(xRs("costosal")), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNASALDOIMPORTE_) = F.NuloNumeric(xRs("costoini")) + F.NuloNumeric(xRs("costoent")) - F.NuloNumeric(xRs("costosal"))
            .TextMatrix(.Rows - 1, COLUMNASALDOIMPORTE_) = Format(.TextMatrix(.Rows - 1, COLUMNASALDOIMPORTE_), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, COLUMNAPRECIOPROM_) = Format(F.NuloNumeric(xRs("costouniprom")), FORMAT_MONTO)
            
            xRs.MoveNext
        End With
    Wend
    
    ' SE PINTAN LOS NEGATIVOS
    pintarGrid Fg1, COLUMNAENTRADACANTIDAD_, &H0&, &HFF&
    pintarGrid Fg1, COLUMNAENTRADAIMPORTE_, &H0&, &HFF&
    pintarGrid Fg1, COLUMNASALIDACANTIDAD_, &H0&, &HFF&
    pintarGrid Fg1, COLUMNASALIDAIMPORTE_, &H0&, &HFF&
    pintarGrid Fg1, COLUMNASALDOCANTIDAD_, &H0&, &HFF&
    pintarGrid Fg1, COLUMNASALDOIMPORTE_, &H0&, &HFF&
    ' SE DA COLOR A LA CELDA
    Fg1.Select Fg1.FixedRows, COLUMNASALDOCANTIDAD_, Fg1.Rows - 1, COLUMNASALDOIMPORTE_
    Fg1.FillStyle = flexFillRepeat
    Fg1.CellBackColor = &HDDFFFF
    Fg1.Select 1, 1
    
    ' SE AGREGAN LOS TOTALES
    Fg1.Rows = Fg1.Rows + 1
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNADESCRIPCION_) = "TOTAL"
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNASTOCKINI_) = Format(GRID_SUMAR_COL(Fg1, COLUMNASTOCKINI_), FORMAT_CANTIDAD)
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAPRECIOINI_) = Format(GRID_SUMAR_COL(Fg1, COLUMNAPRECIOINI_), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAENTRADACANTIDAD_) = Format(GRID_SUMAR_COL(Fg1, COLUMNAENTRADACANTIDAD_), FORMAT_CANTIDAD)
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAENTRADAIMPORTE_) = Format(GRID_SUMAR_COL(Fg1, COLUMNAENTRADAIMPORTE_), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNASALIDACANTIDAD_) = Format(GRID_SUMAR_COL(Fg1, COLUMNASALIDACANTIDAD_), FORMAT_CANTIDAD)
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNASALIDAIMPORTE_) = Format(GRID_SUMAR_COL(Fg1, COLUMNASALIDAIMPORTE_), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNASALDOCANTIDAD_) = Format(GRID_SUMAR_COL(Fg1, COLUMNASALDOCANTIDAD_), FORMAT_CANTIDAD)
    Fg1.TextMatrix(Fg1.Rows - 1, COLUMNASALDOIMPORTE_) = Format(GRID_SUMAR_COL(Fg1, COLUMNASALDOIMPORTE_), FORMAT_MONTO)
    Fg1.Select Fg1.Rows - 1, COLUMNADESCRIPCION_, Fg1.Rows - 1, COLUMNASALDOIMPORTE_
    Fg1.FillStyle = flexFillRepeat
    Fg1.CellFontBold = True
    Fg1.Select 1, 1
    Fg1.TopRow = Fg1.Rows - 1
xSalir:
    Set xRs = Nothing
    Set RstDet = Nothing
    FraProgreso.Visible = False
End Sub

Private Sub pintarGrid(GRID_ As VSFlexGrid, COLUMNA_ As Integer, COLOR1_ As String, COLOR2_ As String)
    Dim A As Integer
    
    With GRID_
        For A = GRID_.FixedRows To .Rows - 1
            .Select A, COLUMNA_
            If NulosN(.TextMatrix(A, COLUMNA_)) >= 0 Then
                .CellForeColor = COLOR1_
            Else
                .CellForeColor = COLOR2_
            End If
        Next A
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
    
    TITULO_ = "REPORTE RESUMEN DE KARDEX"

    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, GRID_, TITULO_, "", "", TITULO_
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
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

Private Function GRID_NUMEROREGISTROS(GRID_ As VSFlexGrid, Optional COLUMNA_ As Integer = 1, Optional FILAINICIO_ As Integer = 1) As Integer
    Dim A As Integer
    Dim CONTADOR_ As Integer
    
    CONTADOR_ = 0
    For A = FILAINICIO_ To GRID_.Rows - 1
        If NulosC(GRID_.TextMatrix(A, COLUMNA_)) <> "" Then CONTADOR_ = CONTADOR_ + 1
    Next
    
    GRID_NUMEROREGISTROS = CONTADOR_
End Function

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VERIFICA QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    Dim NUMEROREGISTROSTIPO_ As Integer
    Dim NUMEROREGISTROSITEM_ As Integer
    
    NUMEROREGISTROSITEM_ = GRID_NUMEROREGISTROS(fg(1))
    ' si esta la fecha correcta
    If F.NuloNumeric(IdAlmacenText.Text) = 0 Then
        MsgBox "Ingrese el almacén", vbExclamation, xTitulo
        IdAlmacenText.SetFocus
        Exit Function
    End If
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
