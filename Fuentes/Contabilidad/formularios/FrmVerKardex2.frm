VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVerKardex2 
   Caption         =   "Contabilidad - Kardex"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   12420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "[ Detalles ]"
      Height          =   1005
      Left            =   4450
      TabIndex        =   19
      Top             =   350
      Width           =   6450
      Begin VB.CommandButton cmd 
         Height          =   240
         Index           =   1
         Left            =   1950
         Picture         =   "FrmVerKardex2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   630
         Width           =   240
      End
      Begin VB.CommandButton cmd 
         Height          =   240
         Index           =   0
         Left            =   1800
         Picture         =   "FrmVerKardex2.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   240
      End
      Begin VB.TextBox txtIdTipPro 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   5
         Text            =   "txtIdTipPro"
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtIdItem 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   7
         Text            =   "txtIdItem"
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label IdItemLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IdItemLabel"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5400
         TabIndex        =   26
         Top             =   645
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ítem"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   660
         Width           =   300
      End
      Begin VB.Label lblItem 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblItem"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2205
         TabIndex        =   22
         Top             =   600
         Width           =   4130
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Ítem"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   885
      End
      Begin VB.Label lblTipPro 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTipPro"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2085
         TabIndex        =   20
         Top             =   240
         Width           =   4250
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Seleccionar ]"
      ForeColor       =   &H00000080&
      Height          =   1005
      Left            =   30
      TabIndex        =   16
      Top             =   350
      Width           =   4405
      Begin VB.CommandButton cmd 
         Height          =   240
         Index           =   2
         Left            =   1455
         Picture         =   "FrmVerKardex2.frx":0264
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   630
         Width           =   240
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   810
         TabIndex        =   0
         Top             =   240
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
         Left            =   2820
         TabIndex        =   2
         Top             =   270
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
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "IdAlmacenText"
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         Height          =   195
         Index           =   1
         Left            =   70
         TabIndex        =   25
         Top             =   645
         Width           =   615
      End
      Begin VB.Label AlmacenLabel 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AlmacenLabel"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1770
         TabIndex        =   24
         Top             =   600
         Width           =   2545
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fin:"
         Height          =   195
         Left            =   2340
         TabIndex        =   18
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   70
         TabIndex        =   17
         Top             =   300
         Width           =   420
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "[  Met. Val. ]"
      ForeColor       =   &H00000080&
      Height          =   1005
      Left            =   10920
      TabIndex        =   13
      Top             =   350
      Width           =   1455
      Begin VB.OptionButton OptVal1 
         Caption         =   "Promedio Ponderado"
         Height          =   435
         Left            =   105
         TabIndex        =   15
         Top             =   225
         Width           =   1210
      End
      Begin VB.OptionButton Option2 
         Caption         =   "P.E.P.S"
         Height          =   195
         Left            =   105
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   990
      Left            =   3390
      TabIndex        =   10
      Top             =   3270
      Visible         =   0   'False
      Width           =   5145
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   150
         TabIndex        =   11
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
         TabIndex        =   12
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
            Picture         =   "FrmVerKardex2.frx":0396
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":0C6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":0DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":1158
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":12DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":1730
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":1848
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":1D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":22D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":23E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":24F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":294C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerKardex2.frx":2AB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12420
      _ExtentX        =   21908
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
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6050
      Left            =   30
      TabIndex        =   1
      Top             =   1440
      Width           =   12330
      _cx             =   21749
      _cy             =   10672
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
      BackColorSel    =   -2147483635
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
      Cols            =   25
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmVerKardex2.frx":3000
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
Attribute VB_Name = "FrmVerKardex2"
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
    COLUMNAENTRADACANTIDAD_
    COLUMNAENTRADAIMPORTE_
    COLUMNASALIDACANTIDAD_
    COLUMNASALIDAIMPORTE_
    COLUMNASALDOCANTIDAD_
    COLUMNASALDOIMPORTE_
    COLUMNAIDITEM_
End Enum

Private Enum COLUMNADETALLADO_
    COLUMNAIDMOVDET_ = 1
    COLUMNAFECHA_
    COLUMNATIPODOC_
    COLUMNANUMSER_
    COLUMNANUMDOC_
    COLUMNATIPOPERACION_
    COLUMNAUNIENTRADA_
    COLUMNAPRECUNIENT_
    COLUMNAIMPENTRADA_
    COLUMNAUNISALIDA_
    COLUMNAPRECUNISAL_
    COLUMNAIMPSALIDA_
    COLUMNAUNISALDO_
    COLUMNAPRECIOUNI_
    COLUMNAIMPSALDO_
    COLUMNAPRECIOPROM_
    COLUMNACLIENTE_
    COLUMNANUMDOCREF_
    COLUMNAMODULO_
    COLUMNANUMREG_
    COLUMNACTANUM_
    COLUMNACTANOMBRE_
End Enum

Public Sub pCargarRpt()
    ' se usara cunado se desee imprimir todos los productos desde la pantalla del kardex resumen
    MuestraRpt = 1
    Form_Activate
End Sub

Private Sub pIniciarCampos()
    TxtFchIni.Valor = CDate("01/01/" & Year(Date))
    TxtFchFin.Valor = Date
    BAND_INTERRUMPIR = False
    
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    
    OptVal1.Value = True
    pConfigurarGrid
    Fg1.Rows = Fg1.FixedRows
End Sub

Private Sub pConfigurarGrid()
    ' --------------DETALLADO
    GRID_COMBINAR Fg1, 0, COLUMNAFECHA_, 0, COLUMNANUMDOC_, "DOC. DE TRASLADO, COMPROBANTE DE PAGO, DOC. INTERNO O SIMILAR", , , , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAFECHA_, 1, COLUMNAFECHA_, "FECHA", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNATIPODOC_, 1, COLUMNATIPODOC_, "TIPO", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNANUMSER_, 1, COLUMNANUMSER_, "SERIE", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNANUMDOC_, 1, COLUMNANUMDOC_, "NUMERO", , False, , , &H8000000F, False
    
    GRID_COMBINAR Fg1, 0, COLUMNATIPOPERACION_, 1, COLUMNATIPOPERACION_, "TIPO OPERACION", , False, , , &H8000000F, False
    
    GRID_COMBINAR Fg1, 0, COLUMNAUNIENTRADA_, 0, COLUMNAIMPENTRADA_, "ENTRADAS", , , , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAUNIENTRADA_, 1, COLUMNAUNIENTRADA_, "CANTIDAD", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAPRECUNIENT_, 1, COLUMNAPRECUNIENT_, "COSTO UNITARIO", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAIMPENTRADA_, 1, COLUMNAIMPENTRADA_, "COSTO TOTAL", , False, , , &H8000000F, False
    
    GRID_COMBINAR Fg1, 0, COLUMNAUNISALIDA_, 0, COLUMNAIMPSALIDA_, "SALIDAS", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAUNISALIDA_, 1, COLUMNAUNISALIDA_, "CANTIDAD", , True, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAPRECUNISAL_, 1, COLUMNAPRECUNISAL_, "COSTO UNITARIO", , True, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAIMPSALIDA_, 1, COLUMNAIMPSALIDA_, "COSTO TOTAL", , True, , , &H8000000F, False
    
    GRID_COMBINAR Fg1, 0, COLUMNAUNISALDO_, 0, COLUMNAIMPSALDO_, "SALDO FINAL", , , , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAUNISALDO_, 1, COLUMNAUNISALDO_, "CANTIDAD", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAPRECIOUNI_, 1, COLUMNAPRECIOUNI_, "COSTO UNITARIO", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 1, COLUMNAIMPSALDO_, 1, COLUMNAIMPSALDO_, "COSTO TOTAL", , False, , , &H8000000F, False
    
    GRID_COMBINAR Fg1, 0, COLUMNAPRECIOPROM_, 1, COLUMNAPRECIOPROM_, "PRECIO PROM.", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, COLUMNACLIENTE_, 1, COLUMNACLIENTE_, "Cliente", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, COLUMNANUMDOCREF_, 1, COLUMNANUMDOCREF_, "Nº Doc. Ref.", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, COLUMNAMODULO_, 1, COLUMNAMODULO_, "Módulo", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, COLUMNANUMREG_, 1, COLUMNANUMREG_, "Nº Reg.", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, COLUMNACTANUM_, 1, COLUMNACTANUM_, "Cta. Num.", , False, , , &H8000000F, False
    GRID_COMBINAR Fg1, 0, COLUMNACTANOMBRE_, 1, COLUMNACTANOMBRE_, "Cta. Nom.", , False, , , &H8000000F, False
    
    Fg1.RowHeight(0) = 430
    Fg1.RowHeight(1) = 430
    Fg1.WordWrap = True
    
    Fg1.ColWidth(COLUMNAFECHA_) = 900
    Fg1.ColWidth(COLUMNATIPODOC_) = 600
    Fg1.ColWidth(COLUMNANUMSER_) = 900
    Fg1.ColWidth(COLUMNANUMDOC_) = 1000
    
    Fg1.ColWidth(COLUMNATIPOPERACION_) = 2000
    
    Fg1.ColWidth(COLUMNAUNIENTRADA_) = 1000
    Fg1.ColWidth(COLUMNAPRECUNIENT_) = 900
    Fg1.ColWidth(COLUMNAIMPENTRADA_) = 900
    
    Fg1.ColWidth(COLUMNAUNISALIDA_) = 1000
    Fg1.ColWidth(COLUMNAPRECIOUNI_) = 900
    Fg1.ColWidth(COLUMNAIMPSALIDA_) = 900
    
    Fg1.ColWidth(COLUMNAUNISALDO_) = 1000
    Fg1.ColWidth(COLUMNAIMPSALDO_) = 900
    
    Fg1.ColWidth(COLUMNAPRECIOPROM_) = 900
    Fg1.ColWidth(COLUMNACLIENTE_) = 2000
    Fg1.ColWidth(COLUMNANUMDOCREF_) = 1200
    Fg1.ColWidth(COLUMNAMODULO_) = 1200
    Fg1.ColWidth(COLUMNANUMREG_) = 1200
    Fg1.ColWidth(COLUMNACTANUM_) = 1200
    Fg1.ColWidth(COLUMNACTANOMBRE_) = 1200
    Fg1.ColWidth(Fg1.ColIndex("IDMOVDET")) = 0
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TextBox PARA EL INGRESO DE DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    txtIdTipPro.Text = ""
    lblTipPro.Caption = ""
    IdItemLabel.Caption = 0
    txtIdItem.Text = ""
    lblItem.Caption = ""
    IdAlmacenText.Text = ""
    AlmacenLabel.Caption = ""
End Sub


Private Sub cmd_Click(Index As Integer)
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    
    Select Case Index
        Case 0 ' TIPO DE PRODUCTO
            ReDim xCampos(1, 4) As String
            xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "3500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
                        
            cSQL = "SELECT mae_tipoproducto.id, mae_tipoproducto.descripcion " _
                + vbCr + "FROM mae_tipoproducto " _
                + vbCr + "WHERE mae_tipoproducto.descripcion <> '' AND mae_tipoproducto.id <> 5"
                            
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "descripcion", "descripcion", Principio
    
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
    
            txtIdTipPro.Text = NulosN(xRs("id"))
            lblTipPro.Caption = NulosC(xRs("descripcion"))
            
        Case 1 ' ITEM
            If NulosN(txtIdTipPro.Text) = 0 Then ' TIPO DE PRODUCTO
                MsgBox "Seleccione por lo menos un tipo de ítem", vbExclamation, xTitulo
                txtIdTipPro.SetFocus
                Exit Sub
            End If
    
            ReDim xCampos(2, 4) As String
            xCampos(0, 0) = "Código":       xCampos(0, 1) = "codpro":       xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "4000":    xCampos(1, 3) = "C"
            
            cSQL = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion " _
                + vbCr + "FROM alm_inventario " _
                + vbCr + "WHERE (((alm_inventario.activo)=-1) AND ((alm_inventario.tippro)=" & NulosN(txtIdTipPro.Text) & "))"
                             
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "codpro", "codpro", Principio
    
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IdItemLabel.Caption = F.NuloNumeric(xRs("id"))
            txtIdItem.Text = F.NuloString(xRs("codpro"))
            lblItem.Caption = F.NuloString(xRs("descripcion"))
            lblItem.ToolTipText = F.NuloString(xRs("descripcion"))
        
        Case 2 ' Almacen
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
            txtIdTipPro.SetFocus
            Set xRs = Nothing
            
    End Select
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        
        If MostrarValorizado = False Then
            Me.Caption = "Almacén - Kardex"
        Else
            Me.Caption = "Contabilidad - Kardex Valorizado"
        End If
        
        If NulosN(IdItemLabel.Caption) = 0 Then
            Blanquea
        Else
            pCargarDetallado NulosN(IdItemLabel.Caption)
        End If
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
        nSQL = "SELECT [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, mae_prov.nombre " _
        + vbCr + "FROM mae_prov RIGHT JOIN (alm_ingresodoc LEFT JOIN com_compras ON alm_ingresodoc.iddoc = com_compras.id) ON mae_prov.id = com_compras.idpro " _
        + vbCr + "WHERE (((alm_ingresodoc.id)=" & IdDocumento & "))"
        
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

Private Function precioPromedio(IDITEM_ As Integer, FECHA_ As String, XCON_ As ADODB.Connection, _
                                Optional TIPO_ As Integer = 1, Optional TIPODOCUMENTO_ As String, _
                                Optional IDDOCUMENTO_ As Integer, Optional CANTIDAD_ As Double) As Double
    Dim cSQL As String
    Dim PRECIOPROMEDIO_ As Double
    Dim PRECIOUNITARIO_ As Double
    Dim A As Integer
    Dim STOCKINICIAL_ As Double
    Dim PRECIOINICIAL_ As Double
    Dim TOTALSALIDAS_ As Double
    Dim TOTALENTRADAS_ As Double
    Dim CANTIDADACUMULADA_ As Double
    Dim IMPORTEACUMULADO_ As Double
    Dim TIPOPRODUCTO_ As Integer
    Dim RECORDSET_ As New ADODB.Recordset
    
    '---------------DETALLE DE MOVIMIENTOS
    cSQL = KardexMovimientoSQL(CDbl(IDITEM_), 0, "01/01/" & Year(CDate(FECHA_)), CDate(FECHA_))
            
    RST_Busq RECORDSET_, cSQL, xCon
    RECORDSET_.Sort = "fchdoc, Tipo, numdoc"
    
    ' --------------STOCK Y PRECIO INICIAL
    STOCKINICIAL_ = SaldoActual(CDbl(IDITEM_), NulosC("01/01/" & AnoTra), NulosC(CDate(TxtFchIni.Valor) - 1), xCon)
    PRECIOINICIAL_ = NulosN(Busca_Codigo("id", NulosC(IDITEM_), "preini", "alm_inventario", "N", xCon))
    PRECIOPROMEDIO_ = PRECIOINICIAL_
    CANTIDADACUMULADA_ = STOCKINICIAL_
    IMPORTEACUMULADO_ = CANTIDADACUMULADA_ * PRECIOINICIAL_
    TOTALENTRADAS_ = TOTALENTRADAS_ + STOCKINICIAL_
        
    Select Case TIPO_
        Case 0
                ' ----------------------------------------------------------INGRESOS
                If TIPODOCUMENTO_ = "C" Or TIPODOCUMENTO_ = "AI" Or TIPODOCUMENTO_ = "P" Then
                    ' -------------------------------------
                    ' ----------------------COSTO DE TAREAS
                    ' -------------------------------------
                    ' ----------------------INSUMOS DE LA PRODUCCION
                    cSQL = "SELECT pro_producciondetins.iditem AS idins, pro_producciondetins.canutil AS cantidad " _
                        + vbCr + "FROM (pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN pro_producciondetins ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro) " _
                        + vbCr + "WHERE (((pro_produccion.id)=" & IDDOCUMENTO_ & ") AND ((pro_producciondet.iditem)=" & IDITEM_ & "));"
                    
                    Set RECORDSET_ = Nothing
                    RST_Busq RECORDSET_, cSQL, XCON_
                    If RECORDSET_.State = 0 Then precioPromedio = 0: Exit Function
                    If RECORDSET_.RecordCount = 0 Then precioPromedio = 0: Exit Function
                    
                    RECORDSET_.MoveFirst
                    While Not RECORDSET_.EOF
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ + (precioPromedio(RECORDSET_("idins"), FECHA_, XCON_) * RECORDSET_("cantidad"))
                        RECORDSET_.MoveNext
                    Wend
                    '-----------------------------------------
                    ' -----------------------COSTO DE PLANILLA
                    '-----------------------------------------
                    Dim DURACPRODUCCION_ As Double
                    Dim DURHORASARREGLO() As String
                    ' ---------------DURACION DE LA PRODUCCION
                    cSQL = "SELECT CDate([pro_producciondet].[horfin]-[pro_producciondet].[horini]) AS dur " _
                        + vbCr + "FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
                        + vbCr + "WHERE (((pro_producciondet.iditem)=" & IDITEM_ & ") AND ((pro_produccion.id)=" & IDDOCUMENTO_ & "));"
                        
                    Set RECORDSET_ = Nothing
                    RST_Busq RECORDSET_, cSQL, XCON_
                    
                    DURHORASARREGLO = Split(Format(RECORDSET_("dur"), "HH:mm"), ":")
                    DURACPRODUCCION_ = NulosN(DURHORASARREGLO(0)) + (NulosN(DURHORASARREGLO(1)) / 60)
                    
                    ' ---------------TOTAL PLANILLA DEL DIA
                    Dim TOTALPLANILLA_ As Double
                    
                    cSQL = "SELECT Sum(pro_pagos.imptot) AS montotot " _
                        + vbCr + "FROM pro_pagos " _
                        + vbCr + "GROUP BY pro_pagos.fchtra " _
                        + vbCr + "HAVING (((pro_pagos.fchtra)=CDate('" & FECHA_ & "')));"
                    
                    Set RECORDSET_ = Nothing
                    RST_Busq RECORDSET_, cSQL, XCON_
                    TOTALPLANILLA_ = NulosN(RECORDSET_("montotot"))
                    
                    ' ---------------TOTAL HORAS DE PRODUCCION DEL DIA
                    Dim TOTALHORASPRODUCCION_ As Double
                    Dim DURHORASNUMERICO_ As Double
                    
                    cSQL = "SELECT pro_producciondet.iditem, CDate([pro_producciondet].[horfin]-[pro_producciondet].[horini]) AS dur " _
                        + vbCr + "FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
                        + vbCr + "WHERE (((pro_produccion.dia)=CDate('" & FECHA_ & "')));"

                    Set RECORDSET_ = Nothing
                    RST_Busq RECORDSET_, cSQL, XCON_
                    
                    RECORDSET_.MoveFirst
                    While Not RECORDSET_.EOF
                        DURHORASARREGLO = Split(Format(RECORDSET_("dur"), "HH:mm"), ":")
                        DURHORASNUMERICO_ = NulosN(DURHORASARREGLO(0)) + (NulosN(DURHORASARREGLO(1)) / 60)
                        TOTALHORASPRODUCCION_ = TOTALHORASPRODUCCION_ + DURHORASNUMERICO_
                        RECORDSET_.MoveNext
                    Wend
                    ' ---------------COSTO PROMEDIO POR HORA
                    Dim COSTOPROMHORA_ As Double
                    COSTOPROMHORA_ = TOTALPLANILLA_ / TOTALHORASPRODUCCION_
                    
                    IMPORTEACUMULADO_ = IMPORTEACUMULADO_ + (COSTOPROMHORA_ * DURACPRODUCCION_)
                    PRECIOPROMEDIO_ = IMPORTEACUMULADO_ / CANTIDAD_
                ' ----------------------------------------------------------SALIDAS
                Else
                End If
                
        Case 1
            If RECORDSET_.RecordCount = 0 Then precioPromedio = PRECIOINICIAL_: Exit Function
            RECORDSET_.MoveFirst
            While Not RECORDSET_.EOF
                ' ----------------------------------------------------------INGRESOS
                If RECORDSET_("tipo") = "C" Or RECORDSET_("tipo") = "AI" Or RECORDSET_("tipo") = "P" Then
                    ' --------------------------------SALDO Y TOTALES
                    If RECORDSET_("descdoc") = "NC" Then
                        CANTIDADACUMULADA_ = CANTIDADACUMULADA_ - NulosN(RECORDSET_("canpro"))
                        TOTALSALIDAS_ = TOTALSALIDAS_ + NulosN(RECORDSET_("canpro"))
                    Else
                        CANTIDADACUMULADA_ = CANTIDADACUMULADA_ + NulosN(RECORDSET_("canpro"))
                        TOTALENTRADAS_ = TOTALENTRADAS_ + NulosN(RECORDSET_("canpro"))
                    End If
                    '---------------------------------PRECIO UNITARIO
                    If RECORDSET_("tipo") = "AI" And RECORDSET_("numdocumentos") <> 0 Then
                        PRECIOUNITARIO_ = PrecioUni(RECORDSET_("id"), CDbl(IDITEM_), NulosC(RECORDSET_("tipo")))
                    Else
                        ' --------------TIPO DE ITEM
                        TIPOPRODUCTO_ = Busca_Codigo(IDITEM_, "id", "tippro", "alm_inventario", "N", XCON_)
                        Select Case TIPOPRODUCTO_
                            Case 3
                                PRECIOUNITARIO_ = precioPromedio(IDITEM_, FECHA_, XCON_, 0, RECORDSET_("tipo"), RECORDSET_("id"), NulosN(RECORDSET_("canpro")))
                            Case Else
                                PRECIOUNITARIO_ = NulosN(RECORDSET_("preuni"))
                                
                        End Select
                    End If
                    ' --------------------------------IMPORTE ACUMULADO
                    If RECORDSET_("descdoc") = "NC" Then
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ - (NulosN(RECORDSET_("canpro")) * PRECIOUNITARIO_)
                    Else
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ + (NulosN(RECORDSET_("canpro")) * PRECIOUNITARIO_)
                    End If
                    ' --------------------------------PRECIO PROMEDIO
                    PRECIOPROMEDIO_ = IMPORTEACUMULADO_ / CANTIDADACUMULADA_
                ' ----------------------------------------------------------SALIDAS
                Else
                    ' --------------------------------SALDO Y TOTALES
                    If RECORDSET_("descdoc") = "NC" Then
                        CANTIDADACUMULADA_ = CANTIDADACUMULADA_ + NulosN(RECORDSET_("canpro"))
                        TOTALENTRADAS_ = TOTALENTRADAS_ + NulosN(RECORDSET_("canpro"))
                    Else
                        CANTIDADACUMULADA_ = CANTIDADACUMULADA_ - NulosN(RECORDSET_("canpro"))
                        TOTALSALIDAS_ = TOTALSALIDAS_ + NulosN(RECORDSET_("canpro"))
                    End If
                    '---------------------------------PRECIO UNITARIO
                    If RECORDSET_("tipo") = "GR" And RECORDSET_("numdocumentos") <> 0 Then
                        PRECIOUNITARIO_ = PrecioUni(RECORDSET_("id"), CDbl(IDITEM_), NulosC(RECORDSET_("tipo")))
                    Else
                        PRECIOUNITARIO_ = PRECIOPROMEDIO_
                    End If
                    ' --------------------------------IMPORTE ACUMULADO
                    If RECORDSET_("descdoc") = "NC" Then
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ + (NulosN(RECORDSET_("canpro")) * PRECIOUNITARIO_)
                    Else
                        IMPORTEACUMULADO_ = IMPORTEACUMULADO_ - (NulosN(RECORDSET_("canpro")) * PRECIOUNITARIO_)
                    End If
                End If
               
                RECORDSET_.MoveNext
            Wend
            
    End Select
    precioPromedio = PRECIOPROMEDIO_
End Function

Private Sub pHallarDatos(IDITEM_ As Integer, FCHINI_ As Date, FCHFIN_ As Date, ByRef INGRESOCANTIDAD_ As Double, ByRef INGRESOIMPORTE_ As Double, _
                            ByRef SALIDACANTIDAD_ As Double, ByRef SALIDAIMPORTE_ As Double)
    Dim xCadSQL As String
    Dim UltPreCosto As Double
    Dim mInicioGrupo As Long '--indica la fila inicial de un grupo, cambia cuando cambia de item
    Dim xPrecioUni As Double '--Indica el precio unitario de cada registro
    Dim xRs As New ADODB.Recordset
    Dim TIPOPRODUCTO_ As Integer
    Dim xSaldo As Double
    Dim xSaldoImp As Double
    Dim A&
    Dim xFila As Integer
    Dim xTotSal, xTotEnt As Double
    Dim xImpSal, xImpEnt As Double
    
    ' AI = Almacen Ingreso
    ' AS = Almacen Salida
    ' C =  Compras
    ' SM = SOLICUTID DE MATERIALES
    ' PP = PARTE DE PRODUCCION
    ' GR = GUIAS DE REMISION
    ' PS =
    
    '--Generar la consulta SQL para obtener el detalle de movimientos del kardex
    xCadSQL = F.KardexMovimientoSQL(CLng(IDITEM_), 0, FCHINI_, FCHFIN_, xCon)
            
    RST_Busq rst, xCadSQL, xCon
    rst.Sort = "fchdoc, Tipo, numdoc"
                
    '--obtener el saldo inicial
    If CDate(FCHINI_) <> CDate("01/01/" & AnoTra) Then
        StockIni = SaldoActual(CDbl(IDITEM_), NulosC("01/01/" & AnoTra), NulosC(CDate(FCHINI_) - 1), xCon)
    Else
        StockIni = NulosN(Busca_Codigo("id", NulosC(IDITEM_), "stckini", "alm_inventario", "N", xCon))
        xPrecioIni = NulosN(Busca_Codigo("id", NulosC(IDITEM_), "preini", "alm_inventario", "N", xCon))
    End If
            
    'UltPreCosto = NulosN(Fg1.TextMatrix(xFila, COLUMNAPRECIOPROM_))
    
    xSaldo = StockIni
    xSaldoImp = xSaldo * xPrecioIni
    xTotEnt = xTotEnt + StockIni
    
    If rst.RecordCount = 0 Then Exit Sub
    rst.MoveFirst

    While Not rst.EOF
        ProgressBar1.Value = A
        ' ----------------------------------------------INGRESOS
        If rst("tipo") = "C" Or rst("tipo") = "AI" Or rst("tipo") = "P" Then
            If rst("descdoc") = "NC" Then
                xSaldo = xSaldo - NulosN(rst("canpro"))
                xTotSal = xTotSal + NulosN(rst("canpro"))
            Else
                xSaldo = xSaldo + NulosN(rst("canpro"))
                xTotEnt = xTotEnt + NulosN(rst("canpro"))
            End If
                            
            '--obtener el precio
            If rst("tipo") = "AI" And rst("numdocumentos") <> 0 Then
                xPrecioUni = PrecioUni(rst("id"), CDbl(IDITEM_), NulosC(rst("tipo")))
            Else
                xPrecioUni = NulosN(rst("preuni"))
            End If
            
            TIPOPRODUCTO_ = Busca_Codigo(IDITEM_, "id", "tippro", "alm_inventario", "N", xCon)
            
            If TIPOPRODUCTO_ = 3 Then
                xPrecioUni = 0
                cSQL = "SELECT con_centrocostopreuni.premprima, con_centrocostopreuni.premobra, con_centrocostopreuni.pregfabrica, con_centrocostopreuni.preuni " _
                    + vbCr + "FROM con_centrocostopreuni " _
                    + vbCr + "WHERE (((con_centrocostopreuni.fecha)=CDate('" & rst("fchdoc") & "')) AND ((con_centrocostopreuni.iditem)=" & IDITEM_ & "));"
                
                Set xRs = Nothing
                RST_Busq xRs, cSQL, xCon
                
                If xRs.State = 0 Then Exit Sub
                If xRs.RecordCount > 0 Then
                    xPrecioUni = NulosN(xRs("premprima")) + NulosN(xRs("premobra"))
                End If
            End If
                                   
            If rst("descdoc") = "NC" Then
                xSaldoImp = xSaldoImp - (NulosN(rst("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(rst("canpro")) * xPrecioUni)
            Else
                xSaldoImp = xSaldoImp + (NulosN(rst("canpro")) * xPrecioUni)
                xImpEnt = xImpEnt + (NulosN(rst("canpro")) * xPrecioUni)
            End If
                                
            UltPreCosto = xPrecioUni
        ' ----------------------------------------------------------SALIDAS
        Else
            If rst("descdoc") = "NC" Then
                xSaldo = xSaldo + NulosN(rst("canpro"))
                xTotEnt = xTotEnt + NulosN(rst("canpro"))
            Else
                xSaldo = xSaldo - NulosN(rst("canpro"))
                xTotSal = xTotSal + NulosN(rst("canpro"))
            End If
            
            xPrecioUni = UltPreCosto
                            
            If rst("descdoc") = "NC" Then
                xSaldoImp = xSaldoImp + (NulosN(rst("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(rst("canpro")) * xPrecioUni)
            Else
                xSaldoImp = xSaldoImp - (NulosN(rst("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(rst("canpro")) * xPrecioUni)
            End If
        End If
        
        rst.MoveNext
    Wend
    
    SALIDACANTIDAD_ = xTotSal
    SALIDAIMPORTE_ = xImpSal
    INGRESOCANTIDAD_ = xTotEnt
    INGRESOIMPORTE_ = xImpEnt
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDetallado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EN FORMA DETALLADA TODOS LOS MOVIMIENTOS DEL ITEM SELECCIONADO, TAMBIEN
'*                    MUESTRA EL PRECIO PROMEDIO DE CADA OPERACION
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub pCargarDetallado(IdItem As Long)
    Dim xCadSQL As String
    Dim UltPreCosto As Double
    Dim mInicioGrupo As Long '--indica la fila inicial de un grupo, cambia cuando cambia de item
    Dim xPrecioUni As Double '--Indica el precio unitario de cada registro
    Dim xPrecioUniProm As Double
    Dim xRs As New ADODB.Recordset
    Dim TIPOPRODUCTO_ As Integer
    Dim xSaldo As Double
    Dim xSaldoImp As Double
    Dim A&
    Dim xFila As Integer
    Dim xTotSal, xTotEnt As Double

    ' AI = Almacen Ingreso
    ' AS = Almacen Salida
    ' C =  Compras
    ' SM = SOLICUTID DE MATERIALES
    ' PP = PARTE DE PRODUCCION
    ' GR = GUIAS DE REMISION
    ' PS =
    
    '********************
    ' Saldo Inicial
    '********************
    xCadSQL = F.SQL_MovHistoricoTotalizado(F.NuloNumeric(IdAlmacenText.Text), CDate(TxtFchIni.Valor) - 1, CStr(IdItem), xCon, True)
    Set rst = Nothing
    Set rst = F.GeneraRstSQL(xCadSQL, xCon)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNATIPOPERACION_) = "SALDO INICIAL"
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAUNIENTRADA_) = Format(NulosN(rst("canini")) + NulosN(rst("canent")) - NulosN(rst("cansal")), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAUNISALDO_) = Format(NulosN(rst("canini")) + NulosN(rst("canent")) - NulosN(rst("cansal")), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAPRECUNIENT_) = Format(NulosN(rst("costouniprom")), FORMAT_IMPORTEKARDEX)
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIMPENTRADA_) = Format(NulosN(rst("costoini")) + NulosN(rst("costoent")) - NulosN(rst("costosal")), FORMAT_IMPORTEKARDEX)
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAPRECIOUNI_) = Format(NulosN(rst("costouniprom")), FORMAT_IMPORTEKARDEX)
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAIMPSALDO_) = Format(NulosN(rst("costoini")) + NulosN(rst("costoent")) - NulosN(rst("costosal")), FORMAT_IMPORTEKARDEX)
        Fg1.TextMatrix(Fg1.Rows - 1, COLUMNAPRECIOPROM_) = Format(NulosN(rst("costouniprom")), FORMAT_IMPORTEKARDEX)
        StockIni = NulosN(rst("canini")) + NulosN(rst("canent")) - NulosN(rst("cansal"))
        xPrecioIni = NulosN(rst("costouniprom"))
    Else
        StockIni = 0
        xPrecioIni = 0
    End If
    
    '*************
    ' Movimientos
    '*************
    xCadSQL = F.SQL_MovDetallado(CStr(IdItem), F.NuloNumeric(IdAlmacenText.Text), TxtFchIni.Valor, TxtFchFin.Valor, xCon, , False, , , True)
    Set rst = Nothing
    Set rst = F.GeneraRstSQL(xCadSQL, xCon)
    
    '--indicando el inicio de grupo
    mInicioGrupo = xFila
    
    UltPreCosto = xPrecioIni
    xPrecioUniProm = xPrecioIni
    xPrecioUni = xPrecioIni
    xSaldo = StockIni
    xSaldoImp = xSaldo * xPrecioIni
    xTotEnt = xTotEnt + StockIni
    
    CentrarFrm Frame4
    Frame4.Visible = True
    Fg1.MergeCells = flexMergeFixedOnly
    ProgressBar1.Min = 0
    If rst.RecordCount = 0 Then
        ProgressBar1.Max = 1
    Else
        ProgressBar1.Max = NulosN(rst.RecordCount)
    End If
    
    '--agregando fila para proceder a ingresar los datos
    Fg1.Rows = Fg1.Rows + 1
    xFila = Fg1.Rows - 1
    
    If rst.RecordCount <> 0 Then
        rst.MoveFirst

        For A = 1 To rst.RecordCount
            ProgressBar1.Value = A
            
            Fg1.TextMatrix(xFila, Fg1.ColIndex("IDMOVDET")) = NulosN(rst("idmovdet"))
            Fg1.TextMatrix(xFila, COLUMNAFECHA_) = Format(rst("fchmov"), "dd/mm/yy")
            Fg1.TextMatrix(xFila, COLUMNATIPODOC_) = NulosC(rst("doc"))
            Fg1.TextMatrix(xFila, COLUMNANUMSER_) = NulosC(rst("numser"))
            Fg1.TextMatrix(xFila, COLUMNANUMDOC_) = NulosC(rst("numdoc"))
            
            If NulosN(rst("tipmov") = 0) Then
                Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = "ALMACEN SALIDA"
            Else
                Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = "ALMACEN INGRESO"
            End If
            
            ' ----------------------------------------------INGRESOS
            If NulosN(rst("tipmov")) = -1 Then
            
                Fg1.TextMatrix(xFila, COLUMNAUNIENTRADA_) = Format(NulosN(rst("cantidad")), FORMAT_MONTO)
                xSaldo = xSaldo + NulosN(rst("cantidad"))
                xTotEnt = xTotEnt + NulosN(rst("cantidad"))
                                
                Fg1.TextMatrix(xFila, COLUMNAUNISALDO_) = Format(xSaldo, FORMAT_MONTO)
                
                If F.NuloNumeric(rst("cantidad")) > 0 Then
                    xPrecioUni = F.NuloNumeric(rst("costo")) / F.NuloNumeric(rst("cantidad"))
                Else
                    xPrecioUni = 0
                End If
                                                
                ' --------------------------------PRECIO UNITARIO
                Fg1.TextMatrix(xFila, COLUMNAPRECUNIENT_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
                
                Fg1.TextMatrix(xFila, COLUMNAIMPENTRADA_) = Format(NulosN(rst("costo")), FORMAT_IMPORTEKARDEX)
                xSaldoImp = xSaldoImp + NulosN(rst("costo"))
                
                Fg1.TextMatrix(xFila, COLUMNAPRECIOUNI_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
                Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_) = Format(xSaldoImp, FORMAT_IMPORTEKARDEX)
                
                ' --------------------------------PRECIO PROMEDIO
                If xSaldo > 0 Then
                    xPrecioUniProm = xSaldoImp / xSaldo
                Else
                    xPrecioUniProm = 0
                End If
                Fg1.TextMatrix(xFila, COLUMNAPRECIOPROM_) = Format(xPrecioUniProm, FORMAT_IMPORTEKARDEX)
                UltPreCosto = xPrecioUni
            
            ' ----------------------------------------------------------SALIDAS
            Else
                If F.NuloNumeric(rst("cantidad")) > 0 Then
                    xPrecioUni = F.NuloNumeric(rst("costo")) / F.NuloNumeric(rst("cantidad"))
                Else
                    xPrecioUni = 0
                End If
                
                Fg1.TextMatrix(xFila, COLUMNAUNISALIDA_) = Format(NulosN(rst("cantidad")), FORMAT_MONTO)
                xSaldo = xSaldo - NulosN(rst("cantidad"))
                xTotSal = xTotSal + NulosN(rst("cantidad"))
                
                '--saldo x cantidad
                Fg1.TextMatrix(xFila, COLUMNAUNISALDO_) = Format(xSaldo, FORMAT_MONTO)
                    
                ' ----------------------PRECIO UNITARIO
                Fg1.TextMatrix(xFila, COLUMNAPRECIOUNI_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
                Fg1.TextMatrix(xFila, COLUMNAPRECUNISAL_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
                
                Fg1.TextMatrix(xFila, COLUMNAIMPSALIDA_) = Format(NulosN(rst("costo")), FORMAT_IMPORTEKARDEX)
                xSaldoImp = xSaldoImp - (NulosN(rst("cantidad")) * xPrecioUni)
                '--saldo
                Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_) = Format(xSaldoImp, FORMAT_IMPORTEKARDEX)
                
                ' -------------PRECIO PROMEDIO
                Fg1.TextMatrix(xFila, COLUMNAPRECIOPROM_) = Format(xPrecioUniProm, FORMAT_IMPORTEKARDEX)
                
                If xSaldo = 0 Then
                    xPrecioUniProm = 0
                End If
            End If
            rst.MoveNext
            If rst.EOF = True Then
                Exit For
            End If
            
            Fg1.Rows = Fg1.Rows + 1
            xFila = xFila + 1
        Next A
    End If
    
    Fg1.Rows = Fg1.Rows + 1
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNATIPOPERACION_, &HBB, True, &H8000000F, "TOTALES"
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAUNIENTRADA_, &HBB, True, &H8000000F, Format(GRID_SUMAR_COL(Fg1, COLUMNAUNIENTRADA_, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAUNISALIDA_, &HBB, True, &H8000000F, Format(GRID_SUMAR_COL(Fg1, COLUMNAUNISALIDA_, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)

    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAIMPENTRADA_, &HBB, True, &H8000000F, Format(GRID_SUMAR_COL(Fg1, COLUMNAIMPENTRADA_, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAIMPSALIDA_, &HBB, True, &H8000000F, Format(GRID_SUMAR_COL(Fg1, COLUMNAIMPSALIDA_, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
    
    Frame4.Visible = False
    Fg1.TopRow = Fg1.Rows - 1
    
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
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then

        If OptVal1.Value = True Then
            
            If fValidarDatos() = False Then Exit Sub
            
            Fg1.Rows = 2
    
            'pCargarDetallado NulosN(LblIdProducto.Caption)
            
        End If
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
    Fg1.Width = Me.Width - 330
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
    If OptVal1.Value = True Then ' COSTO PROMEDIO
        If fValidarDatos() = False Then Exit Sub
        Fg1.Rows = Fg1.FixedRows
        pCargarDetallado NulosN(IdItemLabel.Caption)
    End If
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
    
    TITULO_ = "RESUMEN DE MOVIMIENTOS"

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
'    Dim NUMEROREGISTROSTIPO_ As Integer
'    Dim NUMEROREGISTROSITEM_ As Integer
        
'    If opt_consulta(0).Value = False And opt_consulta(1).Value = False Then
'        MsgBox "Seleccione el Tipo de Ítem para la consulta", vbExclamation, xTitulo
'        Exit Function
'    End If
    
    If NulosN(txtIdTipPro.Text) = 0 Then ' TIPO DE PRODUCTO
        MsgBox "Seleccione por lo menos un tipo de ítem", vbExclamation, xTitulo
        txtIdTipPro.SetFocus
        Exit Function
    End If
    
    If NulosN(IdItemLabel.Caption) = 0 Then ' ITEM
        MsgBox "Seleccione por lo menos un ítem", vbExclamation, xTitulo
        txtIdItem.SetFocus
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
