VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form FrmVerCostoVenta 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Contabilidad - Costo de Venta Detallado"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "[  Opcion  ]"
      ForeColor       =   &H00000080&
      Height          =   1005
      Left            =   9000
      TabIndex        =   22
      Top             =   360
      Width           =   2055
      Begin VB.OptionButton OpTipProd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo de Producto"
         Height          =   195
         Left            =   105
         TabIndex        =   24
         Top             =   225
         Value           =   -1  'True
         Width           =   1830
      End
      Begin VB.OptionButton OpAlmacenes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Almacenes"
         Height          =   195
         Left            =   105
         TabIndex        =   23
         Top             =   495
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "[ Detalles ]"
      Height          =   1005
      Left            =   2040
      TabIndex        =   8
      Top             =   360
      Width           =   6885
      Begin VB.CommandButton cmd 
         Height          =   240
         Index           =   1
         Left            =   1800
         Picture         =   "FrmVerCostoVenta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   630
         Width           =   240
      End
      Begin VB.CommandButton cmd 
         Height          =   240
         Index           =   0
         Left            =   1800
         MaskColor       =   &H00E0E0E0&
         Picture         =   "FrmVerCostoVenta.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   270
         Width           =   240
      End
      Begin VB.TextBox txtIdItem 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   12
         Text            =   "txtIdItem"
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox txtIdTipPro 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   11
         Text            =   "txtIdTipPro"
         Top             =   240
         Width           =   915
      End
      Begin XtremeSuiteControls.Label lblDetalle 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo"
      End
      Begin VB.Label lblTipPro 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTipPro"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2085
         TabIndex        =   15
         Top             =   240
         Width           =   4710
      End
      Begin VB.Label lblItem 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblItem"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2085
         TabIndex        =   14
         Top             =   600
         Width           =   4710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ítem"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   660
         Width           =   300
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "[ Seleccionar ]"
      ForeColor       =   &H00000080&
      Height          =   1005
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   1905
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   540
         TabIndex        =   4
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
         CalendarTitleBackColor=   16444898
         CalendarTrailingForeColor=   12632256
         Valor           =   "23/03/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   540
         TabIndex        =   5
         Top             =   630
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
         CalendarTitleBackColor=   16444898
         CalendarTrailingForeColor=   12632256
         Valor           =   "23/03/2007"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fin:"
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   660
         Width           =   255
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "[  Metodo Valorizacion  ]"
      ForeColor       =   &H00000080&
      Height          =   1005
      Left            =   11160
      TabIndex        =   0
      Top             =   360
      Width           =   2055
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "P.E.P.S"
         Height          =   195
         Left            =   105
         TabIndex        =   2
         Top             =   495
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton OptVal1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Promedio Ponderado"
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   225
         Width           =   1830
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
            Picture         =   "FrmVerCostoVenta.frx":0264
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":07A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":0B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":0C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":1026
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":11AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":15FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":1716
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":1C5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":219E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":22B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":23C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":281A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVerCostoVenta.frx":2986
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   14505
      _ExtentX        =   25585
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
      Left            =   0
      TabIndex        =   18
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
      Cols            =   28
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmVerCostoVenta.frx":2ECE
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00FAEDE2&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   990
         Left            =   3960
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   5145
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   315
            Left            =   150
            TabIndex        =   20
            Top             =   435
            Width           =   4830
            _ExtentX        =   8520
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00404040&
            BorderWidth     =   2
            Index           =   0
            X1              =   120
            X2              =   5265
            Y1              =   1080
            Y2              =   1080
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
         Begin VB.Line Line6 
            BorderColor     =   &H00404040&
            BorderWidth     =   2
            Index           =   0
            X1              =   5130
            X2              =   5130
            Y1              =   15
            Y2              =   945
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
         Begin VB.Label Label7 
            BackColor       =   &H00FAEDE2&
            Caption         =   "Cargando "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   165
            TabIndex        =   21
            Top             =   150
            Width           =   1650
         End
      End
   End
   Begin VB.Label LblTipItem 
      Caption         =   "Label1"
      Height          =   375
      Left            =   13800
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "FrmVerCostoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FrmVerCostoVenta.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : MUESTRA EL VINCAR DEL ITEM SELECCIONADO, ADEMAS PERMITE COSTEAS LAS VENTAS
'*                    MEDIANTE EL METODO PROMEDIO PONDERADO
'* DISEÑADO POR     :
'* ULTIMA REVISION  :
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim Rst As New ADODB.Recordset            ' RECORSET QUE ALAMCENARA LOS MOVIMIENTOS DEL ITEM
Dim SeEjecuto As Boolean                  ' VARIABLE QUE CONTROLARA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim StockIni As Double                    ' ALMACENA EL STOCK INICIAL DEL ITEM
Dim xPrecioIni As Double                  ' ALMACENA EL PRECIO INICIAL DEL ITEM
Dim MuestraRpt As Integer
Dim cSQL As String
Dim INDICE_ As Integer
Dim BAND_INTERRUMPIR As Boolean
'***
Dim xtippro  As Integer
Dim xfam As Integer
Dim xcla As Integer
Dim xsubcla As Integer
'**
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
    COLUMNAFECHA_ = 1
    COLUMNATIPODOC_
    COLUMNANUMSER_
    COLUMNANUMDOC_
    COLUMNATIPOPERACION_
    COLUMNAUNIENTRADA_
    COLUMNAPRECUNIENT_
    COLUMNAIMPENTRADA_
    '***
        COLUMNAPRECUNIENTPROD_
        COLUMNAIMPENTRADAPROD_
    '***
    COLUMNAUNISALIDA_
    COLUMNAPRECUNISAL_
    COLUMNAIMPSALIDA_
'    '***
'        COLUMNAPRECUNISALPROD_
'        COLUMNAIMPSALIDAPROD_
    '***
    
    
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
    GRID_COMBINAR Fg1, 0, COLUMNAFECHA_, 0, COLUMNANUMDOC_, "DOC. DE TRASLADO, COMPROBANTE DE PAGO, DOC. INTERNO O SIMILAR", , , , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNAFECHA_, 1, COLUMNAFECHA_, "FECHA", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNATIPODOC_, 1, COLUMNATIPODOC_, "TIPO", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNANUMSER_, 1, COLUMNANUMSER_, "SERIE", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNANUMDOC_, 1, COLUMNANUMDOC_, "NUMERO", , False, , , &HE0E0E0, False
    
    GRID_COMBINAR Fg1, 0, COLUMNATIPOPERACION_, 1, COLUMNATIPOPERACION_, "TIPO OPERACION", , False, , , &HE0E0E0, False
    'COSTO DE VENTA
    GRID_COMBINAR Fg1, 0, COLUMNAUNIENTRADA_, 0, COLUMNAIMPENTRADAPROD_, "COSTO DE VENTA", , , , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNAUNIENTRADA_, 1, COLUMNAUNIENTRADA_, "CANTIDAD", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNAPRECUNIENT_, 1, COLUMNAPRECUNIENT_, "COSTO KARDEX", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNAIMPENTRADA_, 1, COLUMNAIMPENTRADA_, "COSTO TOTAL", , False, , , &HE0E0E0, False
    '*****
    GRID_COMBINAR Fg1, 1, COLUMNAPRECUNIENTPROD_, 1, COLUMNAPRECUNIENTPROD_, "PRECIO VENTA", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNAIMPENTRADAPROD_, 1, COLUMNAIMPENTRADAPROD_, "TOTAL VENTA", , False, , , &HE0E0E0, False
    '*****
    
    GRID_COMBINAR Fg1, 0, COLUMNAUNISALIDA_, 0, COLUMNAIMPSALIDA_, "DEVOLUCIONES", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNAUNISALIDA_, 1, COLUMNAUNISALIDA_, "CANTIDAD", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNAPRECUNISAL_, 1, COLUMNAPRECUNISAL_, "COSTO UNITARIO", , True, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNAIMPSALIDA_, 1, COLUMNAIMPSALIDA_, "COSTO TOTAL", , True, , , &HE0E0E0, False
    '*****
'    GRID_COMBINAR Fg1, 1, COLUMNAPRECUNISALPROD_, 1, COLUMNAPRECUNISALPROD_, "COSTO VENTA", , False, , , &HE0E0E0, False
'    GRID_COMBINAR Fg1, 1, COLUMNAIMPSALIDAPROD_, 1, COLUMNAIMPSALIDAPROD_, "TOTAL VENTA", , False, , , &HE0E0E0, False
    '*****
    
    '/// DIFERENCIAS
    GRID_COMBINAR Fg1, 0, COLUMNAUNISALDO_, 0, COLUMNAIMPSALDO_, "DIFERENCIA", , , , , &HE0E0E0, False
'    GRID_COMBINAR Fg1, 1, COLUMNAUNISALDO_, 1, COLUMNAUNISALDO_, "CANTIDAD", , False, , , &HE0E0E0, False
'    GRID_COMBINAR Fg1, 1, COLUMNAPRECIOUNI_, 1, COLUMNAPRECIOUNI_, "COSTO UNITARIO", , False, , , &HE0E0E0, False
'    GRID_COMBINAR Fg1, 1, COLUMNAIMPSALDO_, 1, COLUMNAIMPSALDO_, "COSTO TOTAL", , False, , , &HE0E0E0, False
    
    GRID_COMBINAR Fg1, 1, COLUMNAUNISALDO_, 1, COLUMNAUNISALDO_, "DIF VENTAS", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNAPRECIOUNI_, 1, COLUMNAPRECIOUNI_, "DIF DEVOL", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 1, COLUMNAIMPSALDO_, 1, COLUMNAIMPSALDO_, "TOTAL", , False, , , &HE0E0E0, False
    
    
    GRID_COMBINAR Fg1, 0, COLUMNAPRECIOPROM_, 1, COLUMNAPRECIOPROM_, "PRECIO PROM.", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 0, COLUMNACLIENTE_, 1, COLUMNACLIENTE_, "Cliente", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 0, COLUMNANUMDOCREF_, 1, COLUMNANUMDOCREF_, "Nº Doc. Ref.", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 0, COLUMNAMODULO_, 1, COLUMNAMODULO_, "Módulo", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 0, COLUMNANUMREG_, 1, COLUMNANUMREG_, "Nº Reg.", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 0, COLUMNACTANUM_, 1, COLUMNACTANUM_, "Cta. Num.", , False, , , &HE0E0E0, False
    GRID_COMBINAR Fg1, 0, COLUMNACTANOMBRE_, 1, COLUMNACTANOMBRE_, "Cta. Nom.", , False, , , &HE0E0E0, False
    
    Fg1.RowHeight(0) = 430
    Fg1.RowHeight(1) = 430
    Fg1.WordWrap = True
    
    Fg1.ColWidth(COLUMNAFECHA_) = 900
    Fg1.ColWidth(COLUMNATIPODOC_) = 600
    Fg1.ColWidth(COLUMNANUMSER_) = 900
    Fg1.ColWidth(COLUMNANUMDOC_) = 900
    
    Fg1.ColWidth(COLUMNATIPOPERACION_) = 2000
    
    Fg1.ColWidth(COLUMNAUNIENTRADA_) = 1000
    Fg1.ColWidth(COLUMNAPRECUNIENT_) = 900
    Fg1.ColWidth(COLUMNAIMPENTRADA_) = 1000
    '****
    Fg1.ColWidth(COLUMNAPRECUNIENTPROD_) = 900
    Fg1.ColWidth(COLUMNAIMPENTRADAPROD_) = 1000
    '****
    
    Fg1.ColWidth(COLUMNAUNISALIDA_) = 1000
   
    Fg1.ColWidth(COLUMNAPRECUNISAL_) = 1000
    Fg1.ColWidth(COLUMNAIMPSALIDA_) = 1000

    
    Fg1.ColWidth(COLUMNAUNISALDO_) = 1100
    Fg1.ColWidth(COLUMNAIMPSALDO_) = 1500
    
    Fg1.ColWidth(COLUMNAPRECIOPROM_) = 0
    '***
    Fg1.ColWidth(COLUMNAPRECIOUNI_) = 1000
    '***
    Fg1.ColWidth(COLUMNACLIENTE_) = 2000
    Fg1.ColWidth(COLUMNANUMDOCREF_) = 1200
    Fg1.ColWidth(COLUMNAMODULO_) = 1200
    Fg1.ColWidth(COLUMNANUMREG_) = 1200
    Fg1.ColWidth(COLUMNACTANUM_) = 1200
   
    Fg1.ColWidth(COLUMNACTANOMBRE_) = 1200
    
    '*********
        Fg1.ColWidth(26) = 0
        Fg1.ColWidth(27) = 0
    '*********

    
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
    txtIdItem.Text = ""
    lblItem.Caption = ""
End Sub


Private Sub Cmd_Click(Index As Integer)
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    
    Dim cSQLAl As String
    
    Select Case Index
        Case 0 ' ALMACENES
            ReDim xCampos(1, 4) As String
            
                      
            xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "3500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
                        
             If OpAlmacenes.Value = True Then
                        
             cSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion , alm_almacenes.idtippro, alm_almacenes.idfam, " _
                + vbCr + " alm_almacenes.idclas, alm_almacenes.idsubclas  " _
                + vbCr + " FROM alm_almacenes " _
                + vbCr + " WHERE alm_almacenes.idtippro <> 0 and alm_almacenes.descripcion <> ''"
            Else
                          
             cSQL = "SELECT mae_tipoproducto.id, mae_tipoproducto.descripcion " _
                + vbCr + " FROM mae_tipoproducto " _
                + vbCr + " WHERE mae_tipoproducto.descripcion <> ''"
       
            End If
            
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "descripcion", "descripcion", Principio
    
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            '****
            txtIdItem.Text = ""
            lblItem.Caption = ""
            '****
            
            If OpAlmacenes.Value = True Then
            txtIdTipPro.Text = NulosN(xRs("id"))
            lblTipPro.Caption = NulosC(xRs("descripcion"))
            xtippro = NulosN(xRs("idtippro"))
            xfam = NulosN(xRs("idfam"))
            xcla = NulosN(xRs("idclas"))
            xsubcla = NulosN(xRs("idsubclas"))
            LblTipItem.Caption = NulosN(xRs("idtippro"))
            Else
            txtIdTipPro.Text = NulosN(xRs("id"))
            lblTipPro.Caption = NulosC(xRs("descripcion"))
            xtippro = NulosN(xRs("id"))
            
            End If
            
            
        Case 1 ' ITEM
            If NulosN(txtIdTipPro.Text) = 0 Then ' TIPO DE PRODUCTO
                MsgBox "Seleccione por lo menos un Almacen", vbExclamation, xTitulo
                txtIdTipPro.SetFocus
                Exit Sub
            End If
    
            ReDim xCampos(1, 4) As String
            xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6000":    xCampos(0, 3) = "C"
            
            cSQLAl = " AND alm_inventario.tippro = " & xtippro
            If xfam <> 0 Then
            cSQLAl = cSQLAl + " AND alm_inventario.idfam = " & xfam
            End If
            If xcla <> 0 Then
            cSQLAl = cSQLAl + " AND alm_inventario.idclas = " & xcla
            End If
            If xsubcla <> 0 Then
            cSQLAl = cSQLAl + " AND alm_inventario.idsubclas = " & xsubcla
            End If
            
            cSQL = "SELECT alm_inventario.id, alm_inventario.descripcion " _
                + vbCr + "FROM alm_inventario " _
                + vbCr + "WHERE alm_inventario.activo=-1 " & cSQLAl
                             
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "descripcion", "descripcion", Principio
    
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            txtIdItem.Text = NulosN(xRs("id"))
            lblItem.Caption = NulosC(xRs("descripcion"))
            
    End Select
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        
        If MostrarValorizado = False Then
            Me.Caption = "Contabilidad - Costo de Venta"
        Else
            'Me.Caption = "Contabilidad - Kardex Valorizado"
        End If
        
        If NulosN(txtIdItem.Text) = 0 Then
            Blanquea
        Else
            pCargarDetallado NulosN(txtIdItem.Text)
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
    Dim Rst As New ADODB.Recordset
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
    
    RST_Busq Rst, nSQL, xCon
    
    Do While Not Rst.EOF
        xCad = xCad + NulosC(Rst("numdoc")) + " " + NulosC(Rst("nombre")) + ", "
        Rst.MoveNext
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
    
    'TIPOPRODUCTO_ = Busca_Codigo(IDITEM_, "id", "tippro", "alm_inventario", "N", xCon)
    
    cSQL = KardexMovimientoSQL(CDbl(IDITEM_), "01/01/" & Year(CDate(FECHA_)), CDate(FECHA_))
            
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
    
    'TIPOPRODUCTO_ = Busca_Codigo(IDITEM_, "id", "tippro", "alm_inventario", "N", xCon)
    
    
    '--Generar la consulta SQL para obtener el detalle de movimientos del kardex
    xCadSQL = KardexMovimientoSQL(CDbl(IDITEM_), FCHINI_, FCHFIN_)
            
    RST_Busq Rst, xCadSQL, xCon
    Rst.Sort = "fchdoc, Tipo, numdoc"
                
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
    
    If Rst.RecordCount = 0 Then Exit Sub
    Rst.MoveFirst

    While Not Rst.EOF
        ProgressBar1.Value = A
        ' ----------------------------------------------INGRESOS
        If Rst("tipo") = "C" Or Rst("tipo") = "AI" Or Rst("tipo") = "P" Then
            If Rst("descdoc") = "NC" Then
                xSaldo = xSaldo - NulosN(Rst("canpro"))
                xTotSal = xTotSal + NulosN(Rst("canpro"))
            Else
                xSaldo = xSaldo + NulosN(Rst("canpro"))
                xTotEnt = xTotEnt + NulosN(Rst("canpro"))
            End If
                            
            '--obtener el precio
            If Rst("tipo") = "AI" And Rst("numdocumentos") <> 0 Then
                xPrecioUni = PrecioUni(Rst("id"), CDbl(IDITEM_), NulosC(Rst("tipo")))
            Else
                xPrecioUni = NulosN(Rst("preuni"))
            End If
            
            TIPOPRODUCTO_ = Busca_Codigo(IDITEM_, "id", "tippro", "alm_inventario", "N", xCon)
            
            If TIPOPRODUCTO_ = 3 Then
                xPrecioUni = 0
                cSQL = "SELECT con_centrocostopreuni.premprima, con_centrocostopreuni.premobra, con_centrocostopreuni.pregfabrica, con_centrocostopreuni.preuni " _
                    + vbCr + "FROM con_centrocostopreuni " _
                    + vbCr + "WHERE (((con_centrocostopreuni.fecha)=CDate('" & Rst("fchdoc") & "')) AND ((con_centrocostopreuni.iditem)=" & IDITEM_ & "));"
                
                Set xRs = Nothing
                RST_Busq xRs, cSQL, xCon
                
                If xRs.State = 0 Then Exit Sub
                If xRs.RecordCount > 0 Then
                    xPrecioUni = NulosN(xRs("premprima")) + NulosN(xRs("premobra"))
                End If
            End If
                                   
            If Rst("descdoc") = "NC" Then
                xSaldoImp = xSaldoImp - (NulosN(Rst("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(Rst("canpro")) * xPrecioUni)
            Else
                xSaldoImp = xSaldoImp + (NulosN(Rst("canpro")) * xPrecioUni)
                xImpEnt = xImpEnt + (NulosN(Rst("canpro")) * xPrecioUni)
            End If
                                
            UltPreCosto = xPrecioUni
        ' ----------------------------------------------------------SALIDAS
        Else
            If Rst("descdoc") = "NC" Then
                xSaldo = xSaldo + NulosN(Rst("canpro"))
                xTotEnt = xTotEnt + NulosN(Rst("canpro"))
            Else
                xSaldo = xSaldo - NulosN(Rst("canpro"))
                xTotSal = xTotSal + NulosN(Rst("canpro"))
            End If
            
            xPrecioUni = UltPreCosto
                            
            If Rst("descdoc") = "NC" Then
                xSaldoImp = xSaldoImp + (NulosN(Rst("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(Rst("canpro")) * xPrecioUni)
            Else
                xSaldoImp = xSaldoImp - (NulosN(Rst("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(Rst("canpro")) * xPrecioUni)
            End If
        End If
        
        Rst.MoveNext
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
Sub pCargarDetallado(IdItem As Double)
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

    ' AI = Almacen Ingreso
    ' AS = Almacen Salida
    ' C =  Compras
    ' SM = SOLICUTID DE MATERIALES
    ' PP = PARTE DE PRODUCCION
    ' GR = GUIAS DE REMISION
    ' PS =
        
        'TIPOPRODUCTO_ = Busca_Codigo(IdItem, "id", "tippro", "alm_inventario", "N", xCon)
        
    '--Generar la consulta SQL para obtener el detalle de movimientos del kardex
    xCadSQL = KardexMovimientoSQLV(IdItem, TxtFchIni.Valor, TxtFchFin.Valor)
            
    RST_Busq Rst, xCadSQL, xCon
    
    Rst.Sort = "fchdoc, Tipo, numdoc"
    
    '--agregar columna
    'Fg1.Rows = Fg1.Rows + 1
    'xFila = Fg1.Rows - 1
    
    '--indicando el inicio de grupo
    'mInicioGrupo = xFila
        
    '--obtener el saldo inicial
    If CDate(TxtFchIni.Valor) <> CDate("01/01/" & AnoTra) Then
        StockIni = SaldoActual(IdItem, NulosC("01/01/" & AnoTra), NulosC(CDate(TxtFchIni.Valor) - 1), xCon)
    Else
        StockIni = NulosN(Busca_Codigo("id", NulosC(IdItem), "stckini", "alm_inventario", "N", xCon))
    End If
    xPrecioIni = pHallarPrecioInicial(CInt(IdItem), TxtFchIni.Valor, CInt(AnoTra))
    
    '--------------------------------------
    'fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = "SALDO INICIAL"
    '** SIN COMENTAR
'    FORMATO_CELDA Fg1, CLng(xFila), COLUMNATIPOPERACION_, &HBB, True, , "SALDO INICIAL"
'    Fg1.TextMatrix(xFila, COLUMNAUNIENTRADA_) = Format(StockIni, FORMAT_MONTO)
'    Fg1.TextMatrix(xFila, COLUMNAUNISALDO_) = Format(StockIni, FORMAT_MONTO)
'    Fg1.TextMatrix(xFila, COLUMNAPRECIOUNI_) = Format(xPrecioIni, FORMAT_MONTO)
'    Fg1.TextMatrix(xFila, COLUMNAIMPENTRADA_) = Format(StockIni * xPrecioIni, FORMAT_MONTO)
'    Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_) = Format(StockIni * xPrecioIni, FORMAT_MONTO)
    '**
    If NulosN(Fg1.TextMatrix(xFila, COLUMNAUNISALDO_)) <> 0 And NulosN(Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_)) <> 0 Then
        Fg1.TextMatrix(xFila, COLUMNAPRECIOPROM_) = NulosN(Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_)) / NulosN(Fg1.TextMatrix(xFila, COLUMNAUNISALDO_))
    End If
    
'    Fg1.TextMatrix(xFila, COLUMNAPRECUNIENT_) = Format(Fg1.TextMatrix(xFila, COLUMNAPRECIOPROM_), FORMAT_MONTO)
'    Fg1.TextMatrix(xFila, COLUMNAPRECIOPROM_) = Format(Fg1.TextMatrix(xFila, COLUMNAPRECIOPROM_), FORMAT_MONTO)
    
    UltPreCosto = xPrecioIni
    
    xSaldo = StockIni
    xSaldoImp = xSaldo * xPrecioIni
    xTotEnt = xTotEnt + StockIni
    
    CentrarFrm Frame4
    Frame4.Visible = True
    Fg1.MergeCells = flexMergeFixedOnly
    ProgressBar1.Min = 0
    If Rst.RecordCount = 0 Then
        ProgressBar1.Max = 1
    Else
        ProgressBar1.Max = NulosN(Rst.RecordCount)
    End If
    
    '--agregando fila para proceder a ingresar los datos
    Fg1.Rows = Fg1.Rows + 1
    xFila = Fg1.Rows - 1
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst

        For A = 1 To Rst.RecordCount
            ProgressBar1.Value = A
            
            Fg1.TextMatrix(xFila, COLUMNAFECHA_) = Format(Rst("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(xFila, COLUMNATIPODOC_) = NulosC(Rst("descdoc"))
            Fg1.TextMatrix(xFila, COLUMNANUMSER_) = NulosC(Rst("numser"))
            Fg1.TextMatrix(xFila, COLUMNANUMDOC_) = NulosC(Rst("numdoc"))
            Fg1.TextMatrix(xFila, COLUMNACLIENTE_) = NulosC(Rst("entidad"))
            Fg1.TextMatrix(xFila, COLUMNAMODULO_) = NulosC(Rst("modulo"))
            Fg1.TextMatrix(xFila, COLUMNANUMREG_) = NulosC(Rst("registro"))
            Fg1.TextMatrix(xFila, COLUMNACTANUM_) = NulosC(Rst("ctanum"))
            Fg1.TextMatrix(xFila, COLUMNACTANOMBRE_) = NulosC(Rst("ctanom"))
            
            If NulosC(Rst("desope")) = "" Then
                Select Case Rst("tipo")
'                    Case "AI"
'                        Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = "ALMACEN INGRESO"
'
'                    Case "AS"
'                        Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = "ALMACEN SALIDA"
'
'                    Case "C"
'                        Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = "COMPRA"
'
'                    Case "SM"
'                        Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = "SOLICITUD DE MATERIALES"
'
'                    Case "PP"
'                        Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = "PARTE DE PRODUCCION"
'
'                    Case "GR"
'                        Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = "GUIAS DE REMISION"
                        
                    Case "V"
                        Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = "VENTA"
                        
'                    Case "P"
'                        Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = "PRODUCCION"
                        
                    Case Else
                        Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = ""
                        
                End Select
            Else
                Fg1.TextMatrix(xFila, COLUMNATIPOPERACION_) = NulosC(Rst("desope"))
            End If
            
            ' ----------------------------------------------INGRESOS
            'If Rst("tipo") = "C" Or Rst("tipo") = "AI" Or Rst("tipo") = "P" Then
            If Rst("tipo") = "V" Then

                If Rst("descdoc") = "NC" Then
                    Fg1.TextMatrix(xFila, COLUMNAUNISALIDA_) = Format(NulosN(Rst("canpro")), FORMAT_MONTO)
                    xSaldo = xSaldo - NulosN(Rst("canpro"))
                    xTotSal = xTotSal + NulosN(Rst("canpro"))
                Else
                    Fg1.TextMatrix(xFila, COLUMNAUNIENTRADA_) = Format(NulosN(Rst("canpro")), FORMAT_MONTO)
                    xSaldo = xSaldo + NulosN(Rst("canpro"))
                    xTotEnt = xTotEnt + NulosN(Rst("canpro"))
                End If
                               
              
                '--obtener el precio
              
                    '*************
                    If Rst("tipo") = "V" Then
                          xPrecioUni = pHallarPrecioInicial(CInt(IdItem), Rst("fchdoc"), CInt(AnoTra))
                    Else
                          xPrecioUni = NulosN(Rst("preuni"))
                    End If
                    '*************
                              
               
                If Rst("tipo") = "V" Then
                    'xPrecioUni = 0
                    '******************
'                    cSQL = " SELECT DISTINCT TOP 1 con_librocostodet.id, con_librocostodet.idlibro, con_librocostodet.idprod, con_librocostodet.iditem, pro_produccion.dia, con_librocostodet.impmprima, con_librocostodet.impmanobr, con_librocostodet.impgasfab, con_librocostodet.cantidad" _
'                        & " FROM (con_librocostodet INNER JOIN con_librocostomatpri ON con_librocostodet.id = con_librocostomatpri.idlibrodet) INNER JOIN pro_produccion ON con_librocostodet.idprod = pro_produccion.id " _
'                        & " WHERE (((con_librocostodet.iditem)= " & CInt(IdItem) & ") AND ((pro_produccion.dia)>=CDate('" & TxtFchIni.Valor & "') And (pro_produccion.dia)<=CDate(' " & Rst("fchdoc") & " '))) " _
'                        & " ORDER BY pro_produccion.dia DESC;"
'                    '******************'
'                    Set xRs = Nothing
'                    RST_Busq xRs, cSQL, xCon
'
'                    If xRs.State = 0 Then Exit Sub
'                    If xRs.RecordCount > 0 Then
'                        xPrecioUni = (NulosN(xRs("impmprima")) + NulosN(xRs("impmanobr")) + NulosN(xRs("impgasfab"))) / NulosN(xRs("cantidad"))
'                    End If
                End If
                                
                                
                ' --------------------------------PRECIO UNITARIO
                If Rst("descdoc") = "NC" Then
                    Fg1.TextMatrix(xFila, COLUMNAPRECUNISAL_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
                    
                    Fg1.TextMatrix(xFila, COLUMNAIMPSALIDA_) = Format(NulosN(Rst("canpro")) * xPrecioUni, FORMAT_IMPORTEKARDEX)
                    xSaldoImp = xSaldoImp - (NulosN(Rst("canpro")) * xPrecioUni)
                    
                                        
                Else
                    Fg1.TextMatrix(xFila, COLUMNAPRECUNIENT_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
                    
                    Fg1.TextMatrix(xFila, COLUMNAIMPENTRADA_) = Format(NulosN(Rst("canpro")) * xPrecioUni, FORMAT_IMPORTEKARDEX)
                    xSaldoImp = xSaldoImp + (NulosN(Rst("canpro")) * xPrecioUni)
                    
                '***************
                 ' -------------PRECIO DE VENTA
                   Fg1.TextMatrix(xFila, COLUMNAPRECUNIENTPROD_) = Format(NulosN(Rst("preven")), FORMAT_IMPORTEKARDEX)
                   Fg1.TextMatrix(xFila, COLUMNAIMPENTRADAPROD_) = Format(NulosN(Rst("canpro")) * NulosN(Rst("preven")), FORMAT_IMPORTEKARDEX)
                '***************
                    
                End If
              
        If NulosC(Rst("descdoc")) = "FA" Then

              Fg1.TextMatrix(xFila, COLUMNAUNISALDO_) = Format(Fg1.TextMatrix(xFila, COLUMNAIMPENTRADAPROD_) - NulosN(Fg1.TextMatrix(xFila, COLUMNAIMPENTRADA_)), FORMAT_IMPORTEKARDEX)
             ' Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_) = Format(Fg1.TextMatrix(xFila, COLUMNAUNISALDO_), FORMAT_IMPORTEKARDEX)

        Else
              Fg1.TextMatrix(xFila, COLUMNAPRECIOUNI_) = Format(Fg1.TextMatrix(xFila, COLUMNAIMPSALIDA_), FORMAT_IMPORTEKARDEX)
              'Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_) = Format(0 - NulosN(Fg1.TextMatrix(xFila, COLUMNAPRECIOUNI_)), FORMAT_IMPORTEKARDEX)

        End If
                          
                          
              If xFila = 2 Then
                If NulosC(Rst("descdoc")) = "FA" Then
                Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_) = Format(Fg1.TextMatrix(xFila, COLUMNAUNISALDO_), FORMAT_IMPORTEKARDEX)
                Else
                Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_) = Format(0 - NulosN(Fg1.TextMatrix(xFila, COLUMNAPRECIOUNI_)), FORMAT_IMPORTEKARDEX)
                End If
                
              Else
              Dim totalant As Double
              Dim nuevoimp As Double
              totalant = Fg1.TextMatrix(xFila - 1, COLUMNAIMPSALDO_)
             
                If NulosC(Rst("descdoc")) = "FA" Then
                nuevoimp = Fg1.TextMatrix(xFila, COLUMNAUNISALDO_)
                nuevoimp = totalant + nuevoimp
                Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_) = Format(nuevoimp, FORMAT_IMPORTEKARDEX)
                Else
                nuevoimp = Fg1.TextMatrix(xFila, COLUMNAPRECIOUNI_)
                nuevoimp = totalant - nuevoimp
                Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_) = Format(nuevoimp, FORMAT_IMPORTEKARDEX)
               
                End If
                
                
              End If
                          
                  
                
                ' --------------------------------PRECIO PROMEDIO
                If xSaldo = 0 Then
                    xPrecioUni = 0
                Else
                    xPrecioUni = xSaldoImp / xSaldo
                End If
                Fg1.TextMatrix(xFila, COLUMNAPRECIOPROM_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
                UltPreCosto = xPrecioUni
            
            ' ----------------------------------------------------------SALIDAS
            Else
'                If Rst("tipo") = "GR" Then
'                    Fg1.TextMatrix(xFila, COLUMNANUMDOCREF_) = MostrarDocumentos(Rst("id"), Rst("tipo"))
'                End If
'
'                If xSaldo = 0 Then
'                    xPrecioUni = 0
'                Else
'                    xPrecioUni = xSaldoImp / xSaldo
'                End If
'
'                If Rst("descdoc") = "NC" Then
'                    Fg1.TextMatrix(xFila, COLUMNAUNIENTRADA_) = Format(NulosN(Rst("canpro")), FORMAT_MONTO)
'                    xSaldo = xSaldo + NulosN(Rst("canpro"))
'                    xTotEnt = xTotEnt + NulosN(Rst("canpro"))
'
'                    '*************************
''                    Fg1.TextMatrix(xFila, COLUMNAUNISALIDA_) = Format(NulosN(Rst("canpro")), FORMAT_MONTO)
''                    xSaldo = xSaldo + NulosN(Rst("canpro"))
''                    xTotEnt = xTotEnt + NulosN(Rst("canpro"))
'                    '*************************
'
'                Else
'                    Fg1.TextMatrix(xFila, COLUMNAUNISALIDA_) = Format(NulosN(Rst("canpro")), FORMAT_MONTO)
'                    xSaldo = xSaldo - NulosN(Rst("canpro"))
'                    xTotSal = xTotSal + NulosN(Rst("canpro"))
'
'                    '***********************************
'
''                    Fg1.TextMatrix(xFila, COLUMNAUNIENTRADA_) = Format(NulosN(Rst("canpro")), FORMAT_MONTO)
''                    xSaldo = xSaldo - NulosN(Rst("canpro"))
''                    xTotSal = xTotSal + NulosN(Rst("canpro"))
'                    '***********************************
'
'                End If
'
'                '--saldo x cantidad
'                Fg1.TextMatrix(xFila, COLUMNAUNISALDO_) = Format(xSaldo, FORMAT_MONTO)
'
'                ' ----------------------PRECIO UNITARIO
'                Fg1.TextMatrix(xFila, COLUMNAPRECIOUNI_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
'
'                If Rst("descdoc") = "NC" Then
'                     Fg1.TextMatrix(xFila, COLUMNAPRECUNIENT_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
'                    '*************
''                     Fg1.TextMatrix(xFila, COLUMNAPRECUNISAL_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
'                    '*************
'
'
'                '***************
'                ' -------------PRECIO DE VENTA
''                   Fg1.TextMatrix(xFila, COLUMNAPRECUNISALPROD_) = Format(NulosN(Rst("preuni")), FORMAT_IMPORTEKARDEX)
''                   Fg1.TextMatrix(xFila, COLUMNAIMPSALIDAPROD_) = Format(NulosN(Rst("canpro")) * NulosN(Rst("preuni")), FORMAT_IMPORTEKARDEX)
'                '***************
'
'
'                Else
'                    Fg1.TextMatrix(xFila, COLUMNAPRECUNISAL_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
'                    '********
''                    Fg1.TextMatrix(xFila, COLUMNAPRECUNIENT_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
'                    '********
'
'
'                 '***************
'                ' -------------PRECIO DE VENTA
'                   Fg1.TextMatrix(xFila, COLUMNAPRECUNIENTPROD_) = Format(NulosN(Rst("preuni")), FORMAT_IMPORTEKARDEX)
'                   Fg1.TextMatrix(xFila, COLUMNAIMPENTRADAPROD_) = Format(NulosN(Rst("canpro")) * NulosN(Rst("preuni")), FORMAT_IMPORTEKARDEX)
'                '***************
'
'
'                End If
'
'                If Rst("descdoc") = "NC" Then
'                    Fg1.TextMatrix(xFila, COLUMNAIMPENTRADA_) = Format(NulosN(Rst("canpro")) * xPrecioUni, FORMAT_IMPORTEKARDEX)
'                    xSaldoImp = xSaldoImp + (NulosN(Rst("canpro")) * xPrecioUni)
'                    '**************
''                    Fg1.TextMatrix(xFila, COLUMNAIMPSALIDA_) = Format(NulosN(Rst("canpro")) * xPrecioUni, FORMAT_IMPORTEKARDEX)
''                    xSaldoImp = xSaldoImp + (NulosN(Rst("canpro")) * xPrecioUni)
'                    '**************
'
'                Else
'                    Fg1.TextMatrix(xFila, COLUMNAIMPSALIDA_) = Format(NulosN(Rst("canpro")) * xPrecioUni, FORMAT_IMPORTEKARDEX)
'                    xSaldoImp = xSaldoImp - (NulosN(Rst("canpro")) * xPrecioUni)
'
'                    '***************
''                    Fg1.TextMatrix(xFila, COLUMNAIMPENTRADA_) = Format(NulosN(Rst("canpro")) * xPrecioUni, FORMAT_IMPORTEKARDEX)
''                    xSaldoImp = xSaldoImp - (NulosN(Rst("canpro")) * xPrecioUni)
'                    '***************
'
'                End If
'                '--saldo
'                Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_) = Format(xSaldoImp, FORMAT_IMPORTEKARDEX)
'
'                ' -------------PRECIO PROMEDIO
'                Fg1.TextMatrix(xFila, COLUMNAPRECIOPROM_) = Format(xPrecioUni, FORMAT_IMPORTEKARDEX)
'
                
               
                
                
            End If
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
            
            DoEvents
            Fg1.Rows = Fg1.Rows + 1
            xFila = xFila + 1
        Next A
    End If
    
    ' Actualizar el saldo actual del producto
    'xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckact = " & NulosN(TxtSaldo.Text) & " WHERE (((alm_inventario.id)=" & IdItem & "));"
    
    Fg1.Rows = Fg1.Rows + 1
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNATIPOPERACION_, &HBB, True, &H8000000F, "TOTALES"
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAUNIENTRADA_, &HBB, True, &H8000000F, Format(GRID_SUMAR_COL(Fg1, COLUMNAUNIENTRADA_, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAUNISALIDA_, &HBB, True, &H8000000F, Format(GRID_SUMAR_COL(Fg1, COLUMNAUNISALIDA_, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)

    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAIMPENTRADA_, &HBB, True, &H8000000F, Format(GRID_SUMAR_COL(Fg1, COLUMNAIMPENTRADA_, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAIMPSALIDA_, &HBB, True, &H8000000F, Format(GRID_SUMAR_COL(Fg1, COLUMNAIMPSALIDA_, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAIMPENTRADAPROD_, &HBB, True, &H8000000F, Format(GRID_SUMAR_COL(Fg1, COLUMNAIMPENTRADAPROD_, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
   
    
    '*************
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAUNISALDO_, &HBB, True, &H8000000F, Format(GRID_SUMAR_COL(Fg1, COLUMNAUNISALDO_, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAPRECIOUNI_, &HBB, True, &H8000000F, Format(GRID_SUMAR_COL(Fg1, COLUMNAPRECIOUNI_, mInicioGrupo, Fg1.Rows - 2), FORMAT_MONTO)
    'FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAIMPSALDO_, &HBB, True, &H8000000F, Format(Fg1, COLUMNAIMPSALDO_, FORMAT_MONTO)
   
    FORMATO_CELDA Fg1, Fg1.Rows - 1, COLUMNAIMPSALDO_, &HBB, True, &H8000000F, Format(Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_), FORMAT_MONTO)
   
    'Fg1.TextMatrix(xFila, COLUMNAIMPSALDO_)'
    '*************
    
    
    Frame4.Visible = False
    
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
'    Fg1.TextMatrix(0, 0) = "          "
'    Fg1.TextMatrix(1, 0) = "          "
'    Fg1.TextMatrix(0, 1) = "Fecha"
'    Fg1.TextMatrix(1, 1) = "Fecha"
'    Fg1.TextMatrix(0, 2) = "TD"
'    Fg1.TextMatrix(1, 2) = "TD"
'    Fg1.TextMatrix(0, 3) = "Nº Documento"
'    Fg1.TextMatrix(1, 3) = "Nº Documento"
'    Fg1.TextMatrix(0, 7) = "P. Unitario"
'    Fg1.TextMatrix(1, 7) = "P. Unitario"
'    Fg1.TextMatrix(0, 11) = "P. Promedio"
'    Fg1.TextMatrix(1, 11) = "P. Promedio"
    
'    Fg1.Redraw = False
'    Fg1.MergeCol(0) = True
'    Fg1.MergeCol(1) = True
'    Fg1.MergeCol(2) = True
'    Fg1.MergeCol(3) = True
'    Fg1.MergeCol(7) = True
'    Fg1.MergeCol(11) = True
    
'    Fg1.MergeCells = 2
'    Fg1.Redraw = True
    
'    With Fg1
'        .MergeCells = flexMergeFree
'        .MergeRow(-1) = True
'        .Cell(flexcpText, 0, 4, 0, 6) = "Unidades"
'        .Cell(flexcpText, 0, 8, 0, 10) = "Importes"
'        .Cell(flexcpBackColor, 0, 0, Fg1.Rows - 1, Fg1.Cols - 1) = &H8000000F
'    End With
   
'    fg1.ColWidth(2) = 500
'    fg1.ColWidth(4) = 900
'    fg1.ColWidth(5) = 900
'    fg1.ColWidth(6) = 900
'
''    If MostrarValorizado = True Then
'        fg1.ColWidth(7) = 900
'        fg1.ColWidth(8) = 900
'        fg1.ColWidth(9) = 900
'        fg1.ColWidth(10) = 1000
'        fg1.ColWidth(11) = 1100
'    Else
'        Fg1.ColWidth(7) = 0
'        Fg1.ColWidth(8) = 0
'        Fg1.ColWidth(9) = 0
'        Fg1.ColWidth(10) = 0
'        Fg1.ColWidth(11) = 0
'    End If
        
'    Fg1.TextMatrix(1, 4) = "Entradas"
'    Fg1.TextMatrix(1, 5) = "Salidas"
'    Fg1.TextMatrix(1, 6) = "Saldo"
'    Fg1.TextMatrix(1, 8) = "Entradas"
'    Fg1.TextMatrix(1, 9) = "Salidas"
'    Fg1.TextMatrix(1, 10) = "Saldo"
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
    If Me.Height > 3000 Then
        Fg1.Top = 1485
        Fg1.Width = Me.Width - 150
        Fg1.Height = Me.Height - 2200 '--165
        
'        LblDescripcion.Top = Me.Height - 700
'        LblDescripcion.Width = Me.Width - 150
        
    End If
End Sub

Private Sub OpAlmacenes_Click()
    lblDetalle.Caption = "Almacen"
    
    txtIdTipPro.Text = ""
    lblTipPro.Caption = ""
    txtIdItem.Text = ""
    lblItem.Caption = ""
    
End Sub

Private Sub OpTipProd_Click()
    lblDetalle.Caption = "Tipo"
    
    txtIdTipPro.Text = ""
    lblTipPro.Caption = ""
    txtIdItem.Text = ""
    lblItem.Caption = ""
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pCargarDatos
    
    If Button.Index = 2 Then ExportarExcel Fg1
    
    If Button.Index = 3 Then pImprimir
    
    If Button.Index = 5 Then
        Set Rst = Nothing
        Unload Me
    End If
End Sub

Private Sub pCargarDatos()
    If OptVal1.Value = True Then ' COSTO PROMEDIO
        If fValidarDatos() = False Then Exit Sub
        Fg1.Rows = Fg1.FixedRows
        pCargarDetallado NulosN(txtIdItem.Text)
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
    
    TITULO_ = "RESUMEN DE MOVIMIENTOS "

    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, GRID_, TITULO_ & "  -  " & lblItem.Caption, "", "", TITULO_
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
    
    'SE DESABILITA ESTA OPCION PORQUE SE ACTUALIZARON LOS FORMULARIOS DE DE VISUALIZACION DEL KARDEX
'    Me.MousePointer = vbHourglass
'    FrmPrintKardex.Cargar2
'    Me.MousePointer = vbDefault
'    FrmPrintKardex.Show
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
    
    If NulosN(txtIdItem.Text) = 0 Then ' ITEM
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


