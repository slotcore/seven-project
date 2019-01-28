VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsProduccion 
   Caption         =   "Producción - Consulta de Producción"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   615
   ClientWidth     =   14235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   14235
   Begin VB.Frame Frame5 
      Caption         =   "[ Agrupar Por ]"
      Height          =   990
      Left            =   3390
      TabIndex        =   17
      Top             =   360
      Width           =   1470
      Begin VB.OptionButton opt_tipo 
         Caption         =   "x &Familia"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   945
      End
      Begin VB.OptionButton opt_tipo 
         Caption         =   "x &Producto"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton opt_tipo 
         Caption         =   "x &Insumo"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Tipo ]"
      Height          =   960
      Left            =   2100
      TabIndex        =   12
      Top             =   375
      Width           =   1275
      Begin VB.OptionButton opt_consulta 
         Caption         =   "&Resumen"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton opt_consulta 
         Caption         =   "&Detallado"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   600
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Seleccionar ]"
      Height          =   960
      Left            =   30
      TabIndex        =   7
      Top             =   375
      Width           =   2055
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec1 
         Height          =   300
         Left            =   630
         TabIndex        =   8
         Top             =   240
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
         Left            =   630
         TabIndex        =   9
         Top             =   570
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   600
         Width           =   465
      End
   End
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2760
      TabIndex        =   1
      Top             =   4065
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   2
         Top             =   345
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   3
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   -60
         X2              =   6360
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5745
         X2              =   5745
         Y1              =   -90
         Y2              =   4800
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
         Left            =   4140
         TabIndex        =   5
         Top             =   75
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
         Left            =   90
         TabIndex        =   4
         Top             =   75
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Producción"
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
         Left            =   1185
         TabIndex        =   3
         Top             =   75
         Width           =   930
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14235
      _ExtentX        =   25109
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4860
         Top             =   90
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
               Picture         =   "FrmConsProduccion.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProduccion.frx":2A98
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6165
      Left            =   60
      TabIndex        =   6
      Top             =   1470
      Width           =   14160
      _cx             =   24977
      _cy             =   10874
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
      FormatString    =   $"FrmConsProduccion.frx":2E2A
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
   Begin VSFlex7Ctl.VSFlexGrid Fg 
      Height          =   915
      Index           =   0
      Left            =   4890
      TabIndex        =   15
      ToolTipText     =   "Buscar Producto"
      Top             =   450
      Width           =   3750
      _cx             =   6615
      _cy             =   1614
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
      FormatString    =   $"FrmConsProduccion.frx":303B
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
   Begin VSFlex7Ctl.VSFlexGrid Fg 
      Height          =   915
      Index           =   1
      Left            =   8640
      TabIndex        =   16
      ToolTipText     =   "Buscar Insumo"
      Top             =   450
      Width           =   3210
      _cx             =   5662
      _cy             =   1614
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsProduccion.frx":30B1
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
   Begin VSFlex7Ctl.VSFlexGrid Fg 
      Height          =   915
      Index           =   2
      Left            =   11850
      TabIndex        =   21
      Top             =   450
      Width           =   2355
      _cx             =   4154
      _cy             =   1614
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsProduccion.frx":3164
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
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmConsProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCONSPRODUCCION.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO QUE MUESTRA LA PRODUCCION DEL PERIODO ESPECIFICADO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 29/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim BAND_INTERRUMPIR As Boolean     ' SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA TRUE SE INTERRUMPE
Dim T_RPT_PERIODO As String         ' PERIODO DEL REPORTE
Dim T_RPT_TITULO As String          ' TITULO DE REPORTE
Dim ARR_ANYO() As String            ' ARRAY DE AÑOS SELECCIONADOS
Dim ARR_XX() As String              ' SE CARGARA CUANDO SE CARGA EL FORMULARIO Y CUANDO SE CAMBIE EL ESTILO(MES, TRIMESTRE,SEMESTRE)
Dim ARR_TMP(3, 1) As String         ' 0 PROGRAMADO=>> 0::TOTAL,1::TOTAL GEN
                                    ' 1 TEORICO=>> 0::TOTAL,1::TOTAL GEN
                                    ' 2 REAL=>> 0::TOTAL,1::TOTAL GEN
                                    ' 3 DIF=>> 0::TOTAL,1::TOTAL GEN
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
                                    ' OBTENDRA VALOR EN fGenerarConsulta()
Dim Q_COL_COMPARAR_GRUPO As Integer ' INDICA LA COLUMNA PARA COMPARAR EL GRUPO
                                    ' OBTENDRA VALOR EN fGenerarConsulta()
Dim Q_COL_GRUPO_ADD As Integer      ' ADICIONAR DATOS AL GRID EN EL GRUPO (EJ. Q_COL_GRUPO_ADD=2 =>> NOMBRE_GRUPO|COLUM1|COLUM2)
                                    ' FNUCIONA SI Q_COL_GRUPO_ADD<>-1
Dim Q_COL_GRUPO_TERMINA As Integer  ' INDICA EL TERMINO DEL GRUPO, UNE LAS CELDAS DE 1 HASTA Q_COL_GRUPO_TERMINA
Dim Q_COL_ARR_TOTAL As Integer      ' NOS INDICA EL TOTAL DE COLUMNAS QUE VA A CONTENER LOS ACUMULADOS
                                    ' OBTENDRA VALOR EN VALIDAR_CONSULTA()
                                    ' SI SEL MES: ENE, FEB, MAR => Q_COL_ARR_TOTAL= 2
                                    ' SI SEL TRI: ENE-MAR, ABR-JUN => Q_COL_ARR_TOTAL= 1 OBS: SE INICIA DESDE POS=0
Dim F_ES_COMPRA As Boolean          ' INDICA SI ES COMPRA O VENTA
                                    ' TRUE::ES COMPRA, FALSE::ES VENTA
Dim ID_PROGRAMA As String
Dim ID_RECETA As String
Dim TIPO_VENTANA As e_PROGRAMA
Dim ESTILO_VISTA As Integer
Dim N_VALOR_FONDO As String         ' AMACENA EL VALOR PARA COMPARAR
Dim N_VALOR_FONDO_COLOR As Long     ' AMACENA EL VALOR DEL COLOR PARA EL FONDO DE LA FILA
Dim F_CAMIAR_FONDO As Boolean       ' FALSE::SE CONSERVA EL FONDO ACTUAL, TRUE::CAMBIA DE FONDO
Dim Q_COL_COMPARAR_FONDO As Integer ' INDICA LA COLUMNA DEL RECORDSET QUE DEBERA DE COMPARAR PARA CAMBIAR DE FONDO
                                    ' -1=NO HACER NADA
Dim cSQL As String

'*****************************************************************************************************
'* Nombre           : pConsultar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LA PRODUCCION DEL PERIODO ESPECIFICADO EN FUNCION A LOS CRITERIOS
'*                    APLICADOS POR EL USUARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pConsultar()
    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    Dim CN_TMP As New ADODB.Connection     ' CONEX TEMPORAL
    Dim Rst_RUTA As New ADODB.Recordset    ' CARGA RUTAS DE BD'S
    Dim vStrSelect As String               ' RECIBIR LA CONSULTA
    ' CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    Dim N_ANYO As String
    Dim SQL_ANYO As String
    Dim k As Integer
    
    If Validar_Consulta() = False Then Exit Sub
    
    BAND_INTERRUMPIR = False
    ' CONFIGURAR LA PRESENTACION DE LA CONSULTA
    LimpiarGrid Me.Fg1, False, 1
    
    ' ENTRAR SOLO UNA VEZ
    vStrSelect = fGenerarConsulta()
    Configurar_Grilla
        
    ' LIMPIAR ARRAY
    Limpiar_ARRAY_TOTAL True
    Me.MousePointer = vbHourglass
    DoEvents
    
    If vStrSelect = "" Then GoTo SALIR
    PosicionarProgBar
    DoEvents
    ' CARGADO EL RST
    RST_Busq rst_select, vStrSelect, xCon
    pCargarDatosGrid rst_select

SALIR:
    FraProgreso.Visible = False
    Set Rst_RUTA = Nothing
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Set rst_select = Nothing
    SHOW_ERROR Me.Name, "pConsultar"
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosGridFondo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CAMBIA EL COLOR ALAS FILAS DEL CONTROL Fg1
'* Paranetros       : NOMBRE      |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    RST_ORIGEN  |  ADODB.Recordset  |  RECORDSET A MOSTRAR EN EL CONTROL Fg1
'*                    X_ROW1      |  Long             |  ESPECIFICA LA FILA DE INICIO
'*                    X_COL1      |  Integer          |  ESPECIFICA LA COLUMNA DE INICIO
'*                    X_ROW2      |  Long             |  ESPECIFICA LA FILA FINAL
'*                    X_COL2      |  Integer          |  ESPECIFICA LA COLUMNA FINAL
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosGridFondo(RST_ORIGEN As ADODB.Recordset, X_ROW1 As Long, X_COL1 As Integer, X_ROW2 As Long, X_COL2 As Integer)
    ''--PONER COLOR FONDO
    If Q_COL_COMPARAR_FONDO = -1 Then Exit Sub
        If IsNumeric(Fg1.TextMatrix(X_ROW1, 1)) = False Then Exit Sub
        If Fg1.TextMatrix(X_ROW1, 1) = e_ESTADO_ROW_GRID.Fila_grupo Then
            ' SI SE DESEA PONER COLOR AL GRUPO
            ' GRID_COLOR_FONDO Fg1, X_ROW1, X_COL1, X_ROW2, X_COL2, RGB(0, 185, 185)
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

'*****************************************************************************************************
'* Nombre           : pCargarDatosGrid
'* Tipo             : FUNCION
'* Descripcion      : FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
'* Paranetros       : NOMBRE     |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    RST_ORIGEN |  ADODB.Recordset  |  RECORDSET QUE SE CARGARA EN EL CONTROL Fg1
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
    PgBar.Min = 0
    PgBar.Max = RST_ORIGEN.RecordCount
    
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
        ' PONER COLOR FONDO
        If Q_COL_COMPARAR_FONDO <> -1 Then pCargarDatosGridFondo RST_ORIGEN, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            
        RST_ORIGEN.MoveNext
        
        ' PONER TOTALES AL FINAL DE LA GRILLA
        If RST_ORIGEN.EOF Then
            pCargarDatosGridAddTotales BAND_ADD_REG, "Total:"
            Select Case ESTILO_VISTA
            Case 0, 1, 4, 5, 8, 9
            Case Else
                pCargarDatosGridAddTotales True, "Tot Gen:", True
            End Select
        Else
            PgBar.Value = CLng(RST_ORIGEN.Bookmark)
        End If
    Wend
End Function

'*****************************************************************************************************
'* Nombre           : Comparar_Grupo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FUNCION QUE NOS PERMITE ARMAR LOS GRUPOS, COMPARA CUANDO CAMBIAR DE GRUPO
'* Paranetros       : NOMBRE          |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    RST_ORIGEN      |  ADODB.Recordset  |  RECORDSET CON DATOS
'*                    BAND_ADD_REG    |  Boolean          |
'*                    Q_COL_COMPARAR  |  Integer          |
'* Devuelve         :
'*****************************************************************************************************
Private Sub Comparar_Grupo(RST_ORIGEN As ADODB.Recordset, BAND_ADD_REG As Boolean, Optional Q_COL_COMPARAR As Integer = -1)
    Dim RST_TEPM_1 As New ADODB.Recordset
    Dim N_GRUPO_ADD As String
    Dim Q_POS As Integer
    
    If Q_COL_COMPARAR_GRUPO = -1 Then
        If RST_ORIGEN.Bookmark = 1 Then ADD_REG Fg1, Fila_Ninguno
        GoTo SALIR
    End If
    
    If Q_COL_COMPARAR = -1 Then Q_COL_COMPARAR = Q_COL_COMPARAR_GRUPO
    
    If Q_COL_GRUPO_ADD <> -1 Then
        For Q_POS = 1 To Q_COL_GRUPO_ADD
            If LCase(RST_ORIGEN.Fields(Q_COL_COMPARAR + Q_POS).Name) = "instot" Then
                N_GRUPO_ADD = Format(NulosN(RST_ORIGEN.Fields(Q_COL_COMPARAR + Q_POS)), FORMAT_MONTO) + " " + N_GRUPO_ADD
            Else
                N_GRUPO_ADD = RST_ORIGEN.Fields(Q_COL_COMPARAR + Q_POS) & "  " + N_GRUPO_ADD
            End If
        Next Q_POS
        N_GRUPO_ADD = "  =>>   " + N_GRUPO_ADD
    End If
    
    If RST_ORIGEN.Bookmark = 1 Then
        ' SE CARGA EN fGenerarConsulta() Q_COL_COMPARAR_GRUPO
        ADD_REG Fg1, Fila_grupo
        UNIR_CELDAS Fg1, Fg1.Rows - 1, 3, Fg1.Rows - 1, Q_COL_GRUPO_TERMINA, INICIO_GRUPO + RST_ORIGEN.Fields(Q_COL_COMPARAR) + N_GRUPO_ADD, flexAlignLeftCenter:
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 3
        If Q_COL_COMPARAR_FONDO <> -1 Then pCargarDatosGridFondo RST_ORIGEN, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
        
        ADD_REG Fg1, Fila_Ninguno
        UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, " ", flexAlignLeftCenter
    Else
        Set RST_TEPM_1 = RST_ORIGEN.Clone
        RST_TEPM_1.Bookmark = RST_ORIGEN.Bookmark
        RST_TEPM_1.MovePrevious
        
        If RST_TEPM_1.Fields(Q_COL_COMPARAR) <> RST_ORIGEN.Fields(Q_COL_COMPARAR) Then
            pCargarDatosGridAddTotales BAND_ADD_REG, "Total:"
            ADD_REG Fg1, Fila_en_Blanco
            UNIR_CELDAS Fg1, Fg1.Rows - 1, IIf(Q_COL_FILA_OCULTA = -1, 1, Q_COL_FILA_OCULTA + 1), Fg1.Rows - 1, Fg1.Cols - 1, " ", flexAlignLeftCenter
            Limpiar_ARRAY_TOTAL
            ADD_REG Fg1, Fila_grupo
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 3, Fg1.Rows - 1, Q_COL_GRUPO_TERMINA, INICIO_GRUPO + RST_ORIGEN.Fields(Q_COL_COMPARAR) + N_GRUPO_ADD, flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 3
            If Q_COL_COMPARAR_FONDO <> -1 Then pCargarDatosGridFondo RST_ORIGEN, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
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
'*                    Q_ROW        |  Long             |
'* Devuelve         :
'*****************************************************************************************************
Private Function pCargarDatosGridArrayTmp(RST_ORIGEN As ADODB.Recordset, Q_ROW As Long)
    Dim Q_INCREMENTO_X_COL As Integer      ' SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer        ' POSICIONA A LA COLUMNA DEL TOTAL X AÑO
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
        
        Select Case LCase(vStrCampo)
            Case "canteo", "canreal", "instot", "prodtotreal"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
            
            Case "dif"
                If NulosN(RST_ORIGEN.Fields(vStrCampo)) > 0 Then     ' azul (consumo ahorrado)
                    FORMATO_CELDA Fg1, Q_ROW, Q_CAMPO + 1, &HFF0000, False, &HFFFFFF, Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_CANTIDAD)
                ElseIf NulosN(RST_ORIGEN.Fields(vStrCampo)) < 0 Then ' rojo (consumo adicional)
                    FORMATO_CELDA Fg1, Q_ROW, Q_CAMPO + 1, &HFF, False, &HFFFFFF, Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_CANTIDAD)
                Else
                    Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                End If
                
            Case "percendesvio"
                If NulosN(RST_ORIGEN.Fields(vStrCampo)) > 0 Then     ' azul (consumo ahorrado)
                    FORMATO_CELDA Fg1, Q_ROW, Q_CAMPO + 1, &HFF0000, False, &HFFFFFF, Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PORCENTAJE) + "%"
                ElseIf NulosN(RST_ORIGEN.Fields(vStrCampo)) < 0 Then ' rojo (consumo adicional)
                    FORMATO_CELDA Fg1, Q_ROW, Q_CAMPO + 1, &HFF, False, &HFFFFFF, Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PORCENTAJE) + "%"
                Else
                    Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PORCENTAJE) + "%"
                End If
            
            Case "percenreal"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PORCENTAJE) + "%"
            
            Case "dia"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(RST_ORIGEN.Fields(vStrCampo), FORMAT_DATE)
                
            Case Else
                ' AGREGAR LOS DEMAS DATOS
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
    Dim X_PRINT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    If F_ES_COMPRA = False Then T_RPT_TITULO = Replace(T_RPT_TITULO, "COMPRA", "VENTA")
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO + " ", "", T_RPT_PERIODO, True, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub Fg1_DblClick()
    'Fg1_KeyDown 13, 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    On Error GoTo error
    
    Me.WindowState = 2
    Me.Height = 7950
    Me.Width = 11910
    
    CentrarFrm Me
    
    ' FORMATO DE LAS GRILLAS
    GRID_COMBOLIST fg(0), 2:        fg(0).Tag = fg(0).FormatString
    GRID_COMBOLIST fg(1), 2:        fg(1).Tag = fg(1).FormatString
    '**************************************************************
    GRID_COMBOLIST fg(2), 2:        fg(2).Tag = fg(2).FormatString
    '**************************************************************
    TxtFec1.valor = CDate("01/01/" + CStr(Year(Date)))
    TxtFec2.valor = Date
    fGenerarConsulta
    Configurar_Grilla
    Exit Sub

error:
    SHOW_ERROR
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    Fg1.Width = Me.Width - 150
    Fg1.Height = Me.Height - 1905
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BAND_INTERRUMPIR = True
    Erase ARR_TMP
End Sub

'*****************************************************************************************************
'* Nombre           : Validar_Consulta
'* Tipo             : FUNCION
'* Descripcion      : FUNCION QUE VALIDARA LA CONSULTA DE LA FECHA ES NULL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Function Validar_Consulta() As Boolean
    If TxtFec1.valor = "" Or TxtFec2.valor = "" Then
        MsgBox "Ingrese una fecha", vbExclamation, xTitulo
        If TxtFec1.valor = "" Then TxtFec1.SetFocus Else TxtFec2.SetFocus
        Exit Function
    End If
    
    If CDate(TxtFec1.valor) > CDate(TxtFec2.valor) Then
        MsgBox "La fecha inicial es superior al Final", vbExclamation, xTitulo
        TxtFec1.SetFocus
        Exit Function
    End If
    
    Validar_Consulta = True
End Function

'*****************************************************************************************************
'* Nombre           : fGenerarConsulta
'* Tipo             : FUNCCION
'* Descripcion      : FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL
'*                    USUARIO
'* Paranetros       :
'* Devuelve         : String
'*****************************************************************************************************
Private Function fGenerarConsulta() As String
    Dim vStrSelect As String            ' CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    Dim vStrFiltro As String
    Dim vStrFiltro_1 As String          ' ESTE FILTRO SERVIRA PARA CONSULTAR EN EL SUB_SELECT
    Dim k As Integer
    Dim SQL_PROD As String
    Dim SQL_INSUMO As String
    Dim SQL_ESTADO As String
    Dim T_CONSULTA As Integer           ' DEL TIPO DE CONSULTA, SE FORMARA EL ENCABEZADO DEL GRID
    
    ' DE LA FECHA
    If CDate(TxtFec1.valor) < CDate(TxtFec2.valor) Then
        vStrFiltro = " ( pro_produccion.fchdoc >=CDATE ('" + TxtFec1.valor + "') AND pro_produccion.fchdoc <= CDATE('" + TxtFec2.valor + "') ) "
        T_RPT_PERIODO = " Del: " + CStr(TxtFec1.valor) + " Al: " + CStr(TxtFec2.valor)
    Else
        vStrFiltro = " pro_produccion.dia = CDATE('" + TxtFec1.valor + "') "
         T_RPT_PERIODO = "Al: " + CStr(TxtFec2.valor)
    End If
    
    '-------------------------------------------------------------
    ' ----------------------CANTIDADES POSITIVAS
    'vStrFiltro = vStrFiltro & " AND pro_producciondetins.canutil>0"
    '-------------------------------------------------------------
    
    ' DE LOS PRODUCTOS
    SQL_PROD = GENERAR_SQL_ID(fg(0), 1, "pro_receta.iditem", "IN")
    If SQL_PROD <> "" Then SQL_PROD = " AND " + SQL_PROD
    ' DE LOS INSUMOS
    SQL_INSUMO = GENERAR_SQL_ID(fg(1), 1, "pro_producciondetins.iditem", "IN")
    If SQL_INSUMO <> "" Then SQL_INSUMO = " AND " + SQL_INSUMO
    ' DE LOS ESTADOS
    SQL_ESTADO = GENERAR_SQL_ID(fg(2), 1, "pro_producciondet.estado", "IN")
    If SQL_ESTADO <> "" Then SQL_ESTADO = " AND " + SQL_ESTADO
    
    vStrFiltro = vStrFiltro + SQL_PROD + SQL_INSUMO + SQL_ESTADO
    ' UTIL PARA REPORTE X INSUMO
    vStrFiltro_1 = Replace(vStrFiltro, "pro_produccion.", "pro_produccion5.")
    vStrFiltro_1 = Replace(vStrFiltro_1, "pro_receta.", "pro_receta5.")
    vStrFiltro_1 = Replace(vStrFiltro_1, "pro_producciondetins.", "pro_producciondetins5.")
    
    ' GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim N_VALOR As String
    Dim N_CAMPOS As String
    Dim N_WHERE As String
    Dim N_FROM As String
    Dim N_GROUP_BY As String
    Dim N_ORDER_BY As String
    N_WHERE = vStrFiltro
    T_CONSULTA = ESTILO_CONSULTA()
    Q_COL_COMPARAR_FONDO = -1
    Select Case T_CONSULTA
        Case 0 ' RESUMIDO / X PRODUCTO
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 10:        Q_POSICION_TOTAL = 8:        Q_COL_COMPARAR_GRUPO = 2
            Q_COL_GRUPO_ADD = -1 '--ADICIONAR DATOS AL GRID EN EL GRUPO (NOMBRE_GRUPO|COLUM1|COLUM2)
            Q_COL_GRUPO_TERMINA = 12
            T_RPT_TITULO = "RESUMEN DE PRODUCCIÓN AGRUPADO POR PRODUCTO"
            N_CAMPOS = " pro_receta.iditem AS proid, alm_inventario.descripcion AS proddesc, mae_unidades.abrev AS produnidabrev, Sum(pro_producciondet.cantidad) AS prodtotreal, mae_tipoproducto.descripcion AS instipprodesc, alm_inventario_1.descripcion AS insdesc, mae_unidades_1.abrev AS insunidabrev, Sum(IIf([Condoc2].[canpro] Is Null,0,[Condoc2].[canpro])) AS canteo, Sum(Condoc2.canutil) AS canreal, Sum(IIf([Condoc2].[canpro] Is Null,0,[Condoc2].[canutil]-[Condoc2].[canpro])) AS dif, IIf([canteo]=0 Or [dif]=0,0,[dif]/[canteo]*100) AS percendesvio, pro_producciondet.idrec, pro_producciondet.idord "
            N_GROUP_BY = " pro_ordenprod.id, pro_ordenprod.idrec, condoc1.idsol, pro_recetains.iditem, pro_ordenprod.idunimed, condoc1.cantteo, condoc1.cantidad, condoc1.idordprod "
            N_ORDER_BY = " alm_inventario.descripcion, mae_tipoproducto.descripcion, alm_inventario_1.descripcion; "
        
        Case 1 ' RESUMIDO / X FAMILIA
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 8:        Q_POSICION_TOTAL = 6:        Q_COL_COMPARAR_GRUPO = 2
            Q_COL_GRUPO_ADD = -1
            Q_COL_GRUPO_TERMINA = 10
            T_RPT_TITULO = "RESUMEN DE PRODUCCIÓN AGRUPADO POR FAMILIA"
            N_CAMPOS = " mae_familia.id AS famid, pro_producciondetins.iditem AS insid, mae_familia.descripcion AS famdesc, mae_tipoproducto.descripcion AS instipprodesc, alm_inventario_1.descripcion AS insdesc, mae_unidades_1.abrev AS insdunidabrev, Sum(IIf(pro_producciondetins.canpro Is Null,0,(pro_producciondetins.canpro*pro_producciondet.cantidad))) AS canteo, Sum(pro_producciondetins.canutil) AS canreal, Sum(IIf(pro_producciondetins.canpro Is Null,0,(pro_producciondetins.canpro*pro_producciondet.cantidad))-pro_producciondetins.canutil) AS dif, IIf([canteo]=0 Or [dif]=0,0,([dif]/[canteo])*100) AS percendesvio "
            N_GROUP_BY = " mae_familia.id, pro_producciondetins.iditem, mae_familia.descripcion, mae_tipoproducto.descripcion, alm_inventario_1.descripcion, mae_unidades_1.abrev "
            N_ORDER_BY = " mae_familia.descripcion, mae_tipoproducto.descripcion, alm_inventario_1.descripcion "
        
        Case 2 ' RESUMEN / X INSUMO
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 12:        Q_POSICION_TOTAL = 7:        Q_COL_COMPARAR_GRUPO = 3
            Q_COL_GRUPO_ADD = 2
            Q_COL_GRUPO_TERMINA = 14
            T_RPT_TITULO = "RESUMEN DE PRODUCCIÓN AGRUPADO POR INSUMO"
            N_CAMPOS = " pro_producciondetins.iditem AS insid, pro_receta.iditem AS proid, mae_tipoproducto.descripcion AS instipprodesc, alm_inventario_1.descripcion AS insdesc, mae_unidades_1.abrev AS insunidabrev, " _
                + vbCr + " (SELECT Sum([pro_producciondetins5].canutil) AS instot FROM pro_produccion AS pro_produccion5 INNER JOIN ((pro_receta AS pro_receta5 INNER JOIN pro_producciondet AS pro_producciondet5 ON pro_receta5.id = pro_producciondet5.idrec) INNER JOIN pro_producciondetins AS pro_producciondetins5 ON (pro_producciondet5.idrec = pro_producciondetins5.idrec) AND (pro_producciondet5.numparte = pro_producciondetins5.numparte) AND (pro_producciondet5.idpro = pro_producciondetins5.idpro)) ON pro_produccion5.id = pro_producciondet5.idpro  " _
                + vbCr + " WHERE pro_producciondetins5.iditem=pro_producciondetins.iditem  AND " + vStrFiltro_1 + " ) AS instot, " _
                + vbCr + " alm_inventario.descripcion AS proddesc, mae_unidades.abrev AS produnidabrev, Sum(pro_producciondet.cantidad) AS prodtotreal, Sum(IIf(pro_producciondetins.canpro Is Null,0,(pro_producciondetins.canpro*pro_producciondet.cantidad))) AS canteo, Sum(pro_producciondetins.canutil) AS canreal, Sum(IIf(pro_producciondetins.canpro Is Null,0,(pro_producciondetins.canpro*pro_producciondet.cantidad))-pro_producciondetins.canutil) AS dif, IIf([canteo]=0 Or [dif]=0,0,[dif]/[canteo]*100) AS percendesvio, IIf([instot]=0 Or [canreal]=0,0,([canreal]/[instot])*100) AS percenreal "
            N_GROUP_BY = " pro_producciondetins.iditem, pro_receta.iditem, mae_tipoproducto.descripcion, alm_inventario_1.descripcion, mae_unidades_1.abrev, alm_inventario.descripcion, mae_unidades.abrev "
            N_ORDER_BY = " mae_tipoproducto.descripcion, alm_inventario_1.descripcion, alm_inventario.descripcion "
        
        Case 3 ' DETALLADO / X PRODUCTO
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 14:        Q_POSICION_TOTAL = 6:        Q_COL_COMPARAR_GRUPO = 2
            T_RPT_TITULO = "DETALLE DE PRODUCCIÓN AGRUPADO POR PRODUCTO"
            Q_COL_GRUPO_ADD = -1
            Q_COL_GRUPO_TERMINA = 8
            Q_COL_COMPARAR_FONDO = 4
            N_CAMPOS = "  pro_receta.iditem AS proid, pro_producciondetins.iditem AS insid, alm_inventario.descripcion AS proddesc, pro_produccion.dia, pro_producciondet.numparte, mae_unidades.abrev AS produnidabrev, pro_producciondet.cantidad AS prodtotreal, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS res, pro_receta.codrec, mae_tipoproducto.descripcion AS instipprodesc, alm_inventario_1.descripcion AS insdesc, mae_unidades_1.abrev AS indunidabrev, IIf(pro_producciondetins.canpro Is Null,0,(pro_producciondetins.canpro*pro_producciondet.cantidad)) AS canteo, pro_producciondetins.canutil AS canreal, IIf(pro_producciondetins.canpro Is Null,0,(pro_producciondetins.canpro*pro_producciondet.cantidad)-pro_producciondetins.canutil) AS dif, IIf([canteo]=0 Or [dif]=0,0,[dif]/[canteo]*100) AS percendesvio "
            N_GROUP_BY = ""
            N_ORDER_BY = " alm_inventario.descripcion, pro_produccion.dia, pro_producciondet.numparte, mae_tipoproducto.descripcion, alm_inventario_1.descripcion; "
    
    End Select
    
    ' DEL FROM
    Select Case T_CONSULTA
        Case 0, 1, 2 ' RESUMEN
            N_FROM = "(((pro_produccion INNER JOIN ((alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) RIGHT JOIN (mae_unidades RIGHT JOIN (pro_receta INNER JOIN pro_producciondet ON pro_receta.id = pro_producciondet.idrec) ON mae_unidades.id = pro_producciondet.idunimed) ON alm_inventario.id = pro_receta.iditem) ON pro_produccion.id = pro_producciondet.idpro) LEFT JOIN " _
                        & "( " _
                        & "SELECT pro_ordenprod.id, pro_ordenprod.idrec, condoc1.idsol, pro_recetains.iditem, pro_ordenprod.idunimed, condoc1.cantteo AS canpro, condoc1.cantidad AS canutil " _
                        & "FROM (pro_ordenprod LEFT JOIN (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) ON pro_ordenprod.idrec = pro_recetains.idrec) LEFT JOIN " _
                        & "( " _
                        & "SELECT pro_ordenprod.id AS idordprod, pro_solicitudmat.id AS idsol, alm_ingresodet.iditem, alm_ingresodet.cantteo, alm_ingresodet.cantidad " _
                        & "FROM ((pro_ordenprod INNER JOIN pro_solicitudmat ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_ingreso ON pro_solicitudmat.id = alm_ingreso.iddocref) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id " _
                        & "WHERE (((pro_solicitudmat.idtipdocref) = 115)) " _
                        & ") " _
                        & "AS condoc1 ON pro_recetains.iditem = condoc1.iditem " _
                        & "GROUP BY pro_ordenprod.id, pro_ordenprod.idrec, condoc1.idsol, pro_recetains.iditem, pro_ordenprod.idunimed, condoc1.cantteo, condoc1.cantidad, condoc1.idordprod " _
                        & ") " _
                        & "AS Condoc2 ON pro_producciondet.idord = Condoc2.id) LEFT JOIN (alm_inventario AS alm_inventario_1 LEFT JOIN mae_tipoproducto ON alm_inventario_1.tippro = mae_tipoproducto.id) ON Condoc2.iditem = alm_inventario_1.id) LEFT JOIN mae_unidades AS mae_unidades_1 ON Condoc2.idunimed = mae_unidades_1.id "
        
        Case 3, 4, 5 ' DETALLE
            N_FROM = "  (pro_produccion INNER JOIN (pla_empleados RIGHT JOIN ((alm_inventario INNER JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) INNER JOIN (pro_emp RIGHT JOIN (mae_unidades INNER JOIN (pro_receta INNER JOIN pro_producciondet ON pro_receta.id = pro_producciondet.idrec) ON mae_unidades.id = pro_producciondet.idunimed) ON pro_emp.id = pro_producciondet.idres) ON alm_inventario.id = pro_receta.iditem) ON pla_empleados.id = pro_emp.idemp) ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN ((mae_unidades AS mae_unidades_1 INNER JOIN (pro_producciondetins INNER JOIN alm_inventario AS alm_inventario_1 ON pro_producciondetins.iditem = alm_inventario_1.id) ON mae_unidades_1.id = pro_producciondetins.idunimed) INNER JOIN mae_tipoproducto ON alm_inventario_1.tippro = mae_tipoproducto.id) ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro) "
        
        Case 2       ' EQUIPO
        
    End Select
    
    cSQL = "SELECT pro_receta.iditem AS proid, alm_inventario.descripcion AS proddesc, mae_unidades.abrev AS produnidabrev, Sum(pro_producciondet.cantidad) AS prodtotreal, mae_tipoproducto.descripcion AS instipprodesc, alm_inventario_1.descripcion AS insdesc, mae_unidades_1.abrev AS insunidabrev, Sum(IIf([Condoc2].[canpro] Is Null,0,[Condoc2].[canpro])) AS canteo, Sum(Condoc2.canutil) AS canreal, Sum(IIf([Condoc2].[canpro] Is Null,0,[Condoc2].[canutil]-[Condoc2].[canpro])) AS dif, IIf([canteo]=0 Or [dif]=0,0,[dif]/[canteo]*100) AS percendesvio, pro_producciondet.idrec, pro_producciondet.idord " _
& "(((pro_produccion INNER JOIN ((alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) RIGHT JOIN (mae_unidades RIGHT JOIN (pro_receta INNER JOIN pro_producciondet ON pro_receta.id = pro_producciondet.idrec) ON mae_unidades.id = pro_producciondet.idunimed) ON alm_inventario.id = pro_receta.iditem) ON pro_produccion.id = pro_producciondet.idpro) LEFT JOIN " _
& "( " _
& "SELECT pro_ordenprod.id, pro_ordenprod.idrec, condoc1.idsol, pro_recetains.iditem, pro_ordenprod.idunimed, condoc1.cantteo AS canpro, condoc1.cantidad AS canutil " _
& "FROM (pro_ordenprod LEFT JOIN (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) ON pro_ordenprod.idrec = pro_recetains.idrec) LEFT JOIN " _
& "( " _
& "SELECT pro_ordenprod.id AS idordprod, pro_solicitudmat.id AS idsol, alm_ingresodet.iditem, alm_ingresodet.cantteo, alm_ingresodet.cantidad " _
& "FROM ((pro_ordenprod INNER JOIN pro_solicitudmat ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_ingreso ON pro_solicitudmat.id = alm_ingreso.iddocref) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id " _
& "WHERE (((pro_solicitudmat.idtipdocref) = 115)) " _
& ") " _
& "AS condoc1 ON pro_recetains.iditem = condoc1.iditem " _
& "GROUP BY pro_ordenprod.id, pro_ordenprod.idrec, condoc1.idsol, pro_recetains.iditem, pro_ordenprod.idunimed, condoc1.cantteo, condoc1.cantidad, condoc1.idordprod " _
& "HAVING (((pro_ordenprod.ID) = 3953) And ((condoc1.idordprod) = 3953)) " _
& ") " _
& "AS Condoc2 ON pro_producciondet.idord = Condoc2.id) LEFT JOIN (alm_inventario AS alm_inventario_1 LEFT JOIN mae_tipoproducto ON alm_inventario_1.tippro = mae_tipoproducto.id) ON Condoc2.iditem = alm_inventario_1.id) LEFT JOIN mae_unidades AS mae_unidades_1 ON Condoc2.idunimed = mae_unidades_1.id " _
+ vbCr + "WHERE (((pro_produccion.fchdoc)>=CDate('01/01/2014') And (pro_produccion.fchdoc)<=CDate('01/06/2014'))) " _
+ vbCr + "GROUP BY pro_receta.iditem, alm_inventario.descripcion, mae_unidades.abrev, mae_tipoproducto.descripcion, alm_inventario_1.descripcion, mae_unidades_1.abrev, pro_producciondet.idrec, pro_producciondet.idord " _
+ vbCr + "ORDER BY alm_inventario.descripcion, mae_tipoproducto.descripcion, alm_inventario_1.descripcion"
    
    ' DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    Q_POS_MES_INICIO = Q_COL_FILA

    ' GENERANDO LA CONSULTA
    vStrSelect = "SELECT " + N_CAMPOS + _
    vbCr + " FROM " + N_FROM + _
    vbCr + " WHERE " + N_WHERE + _
    vbCr + IIf(N_GROUP_BY <> "", " GROUP BY ", "") + N_GROUP_BY + _
    vbCr + " ORDER BY " + N_ORDER_BY
    fGenerarConsulta = vStrSelect
End Function

'*****************************************************************************************************
'* Nombre           : Limpiar_ARRAY_TOTAL
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE           |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    F_LIMPIA_TOT_GRL |  Boolean   |
'* Devuelve         :
'*****************************************************************************************************
Private Sub Limpiar_ARRAY_TOTAL(Optional F_LIMPIA_TOT_GRL As Boolean = False)
    Dim k As Integer
    For k = 0 To UBound(ARR_TMP())
'        ARR_TMP(k, 3) = 0
'        If F_LIMPIA_TOT_GRL = True Then ARR_TMP(k, 4) = 0
    Next
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosGridAddTotales
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA LOS TOTALES POR CADA GRUPO Y EL TOTAL GENERAL
'* Paranetros       : NOMBRE          |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    BAND_ADD_TOTAL  |  Boolean   |
'*                    Nombre_total    |  String    |
'*                    Band_Total_gral |  Boolean   |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosGridAddTotales(BAND_ADD_TOTAL As Boolean, Nombre_total As String, Optional Band_Total_gral As Boolean = False)
    ' ACUMULA LOS TOTALES EN EL TOTAL GENERAL
    Dim Q_MES As Integer
    Dim X_ROW As Integer
''''''    'On Error Resume Next
''''''    X_ROW = Fg1.Rows
''''''    If BAND_ADD_TOTAL = True Then
''''''        '--AGREAGNDO NUEVA FILA
''''''        ADD_REG Fg1, IIf(Band_Total_gral = False, Fila_Total, Fila_Total_grl)
''''''
''''''        'PONIENDO LOS NOMBRES DE LOS TOTALES  Q_POSICION_TOTAL SE OBTIENE DE fGenerarConsulta()
''''''        Fg1.TextMatrix(X_ROW, Q_POSICION_TOTAL) = Nombre_total
''''''        FORMATO_CELDA Fg1, X_ROW, Q_POSICION_TOTAL
''''''    End If
''''''
''''''
''''''    '--ACUMULANDO LOS TOTALES GRLES
''''''    If Band_Total_gral = False Then
''''''        For Q_MES = 0 To Q_COL_ARR_TOTAL
''''''            ARR_TMP(Q_MES, 4) = NulosN(ARR_TMP(Q_MES, 4)) + NulosN(ARR_TMP(Q_MES, 3))
''''''        Next Q_MES
''''''        If Q_COL_FILA_ULTIMO <> -1 Then
''''''            ARR_TMP_1(0, 1) = NulosN(ARR_TMP_1(0, 1)) + NulosN(ARR_TMP_1(0, 0)) '--STOCK
''''''            ARR_TMP_1(1, 1) = NulosN(ARR_TMP_1(1, 1)) + NulosN(ARR_TMP_1(1, 0)) '--SALDO
''''''        End If
''''''    End If
''''''    '
'''''''--------------------------
''''''    Dim Q_INCREMENTO_X_COL As Integer   '--SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
''''''    Dim Q_POS_MES_TOTAL  As Integer     '--POSICIONA A LA COLUMNA DEL TOTAL X AÑO
''''''
''''''
''''''    '--IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
''''''    Q_INCREMENTO_X_COL = 0
''''''    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
''''''    '-----------
''''''
''''''    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
''''''
''''''    For Q_MES = 0 To Q_COL_ARR_TOTAL
''''''        '--INTERRUMPIR EL PROCESO
''''''        If BAND_INTERRUMPIR = True Then Exit Sub
''''''        Fg1.TextMatrix(X_ROW, Q_POS_MES) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP(Q_MES, 3), ARR_TMP(Q_MES, 4)), Band_Total_gral, Q_MES)
''''''        FORMATO_CELDA Fg1, X_ROW, Q_POS_MES
''''''        Q_POS_MES = Q_POS_MES + 1
''''''    Next Q_MES
''''''
''''''
''''''    If Q_COL_FILA_ULTIMO <> -1 Then
''''''        '--STOCK
''''''        Fg1.TextMatrix(X_ROW, Fg1.Cols - 2) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP_1(0, 0), ARR_TMP_1(0, 1)), Band_Total_gral, Fg1.Cols - 2)
''''''        FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 2, RGB(128, 0, 0)
''''''        '--SALDO
''''''        Fg1.TextMatrix(X_ROW, Fg1.Cols - 1) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP_1(1, 0), ARR_TMP_1(1, 1)), Band_Total_gral, Fg1.Cols - 1)
''''''        FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 1, vbRed
''''''    End If
''''''    Err.Clear
End Sub

'*****************************************************************************************************
'* Nombre           : Configurar_Grilla
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA DE ACUERDO A LO QUE SE SELECCIONA
'* Paranetros       : NOMBRE              |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    F_CONSERVAR_FORMATO |  Boolean    |
'* Devuelve         :
'*****************************************************************************************************
Private Sub Configurar_Grilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    Dim M_ANCHO_COL As Integer         ' DEPENDERA DEL TIPO DE CONSULTA
    Dim k, j As Integer
    Dim T_CONSULTA As Integer
    
    If F_CONSERVAR_FORMATO = True Then Fg1.Clear
    
    Fg1.FrozenCols = 0
    
    M_ANCHO_COL = 0

    With Fg1
        Fg1.Cols = Q_COL_FILA_OCULTA + Q_COL_FILA
        Q_POS_MES = Q_POS_MES_INICIO
        .ColWidth(0) = 200
        '--DATOS DE FILA
        
        T_CONSULTA = ESTILO_CONSULTA()
        
        Select Case T_CONSULTA
            Case 0 ' RESUMIDO / X PRODUCTO
                    .TextMatrix(0, 3) = "Producto":         .ColWidth(3) = 0:    .ColAlignment(3) = flexAlignLeftBottom:            .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 4) = "U.M.":             .ColWidth(4) = 450:     .ColAlignment(4) = flexAlignLeftBottom:         .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 5) = "Total":            .ColWidth(5) = 1000:     .ColAlignment(5) = flexAlignRightBottom:       .Row = 0: .Col = 5: .CellAlignment = flexAlignRightBottom
                    
                    .TextMatrix(0, 6) = "Tip. Prod":        .ColWidth(6) = 1000:    .ColAlignment(6) = flexAlignLeftBottom:         .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 7) = "Insumo":           .ColWidth(7) = 4500:    .ColAlignment(7) = flexAlignLeftBottom:         .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 8) = "U.M.":             .ColWidth(8) = 450:     .ColAlignment(8) = flexAlignLeftBottom:         .Row = 0: .Col = 8: .CellAlignment = flexAlignCenterBottom
                    M_ANCHO_COL = 0
            
            Case 1 ' RESUMIDO / X FAMILIA
                    .TextMatrix(0, 3) = "Familia":          .ColWidth(3) = 0:       .ColAlignment(3) = flexAlignLeftBottom:         .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 4) = "Tip. Prod":        .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftBottom:         .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 5) = "Insumo":           .ColWidth(5) = 4000:    .ColAlignment(5) = flexAlignLeftBottom:         .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 6) = "U.M.":             .ColWidth(6) = 450:     .ColAlignment(6) = flexAlignLeftBottom:         .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterBottom
                    M_ANCHO_COL = 200
                    
            Case 2 ' RESUMEN / X INSUMO
                    .TextMatrix(0, 3) = "Tip. Prod":        .ColWidth(3) = 0:       .ColAlignment(3) = flexAlignLeftBottom:         .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 4) = "Insumo":           .ColWidth(4) = 0:       .ColAlignment(4) = flexAlignLeftBottom:         .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 5) = "U.M.":             .ColWidth(5) = 0:       .ColAlignment(5) = flexAlignLeftBottom:         .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 6) = "Total Ins.":       .ColWidth(6) = 0:       .ColAlignment(6) = flexAlignRightBottom:        .Row = 0: .Col = 6: .CellAlignment = flexAlignRightBottom
                    
                    .TextMatrix(0, 7) = "Producto":         .ColWidth(7) = 4500:    .ColAlignment(7) = flexAlignLeftBottom:         .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 8) = "U.M.":             .ColWidth(8) = 450:     .ColAlignment(8) = flexAlignLeftBottom:         .Row = 0: .Col = 8: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 9) = "Total":            .ColWidth(9) = 1000:    .ColAlignment(9) = flexAlignRightBottom:        .Row = 0: .Col = 9: .CellAlignment = flexAlignRightBottom
                    M_ANCHO_COL = 0
                    
            Case 3 ' DETALLE / X PRODUCTO
                    .TextMatrix(0, 3) = "Producto":         .ColWidth(3) = 0:       .ColAlignment(3) = flexAlignLeftBottom:         .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 4) = "Dia":              .ColWidth(4) = 850:     .ColAlignment(4) = flexAlignCenterBottom:       .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterBottom
                    .TextMatrix(0, 5) = "Nº Prod.":       .ColWidth(5) = 900:     .ColAlignment(5) = flexAlignCenterBottom:         .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterBottom
                    .TextMatrix(0, 6) = "U.M.":             .ColWidth(6) = 450:     .ColAlignment(6) = flexAlignLeftBottom:         .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 7) = "Total":            .ColWidth(7) = 800:     .ColAlignment(7) = flexAlignRightBottom:        .Row = 0: .Col = 7: .CellAlignment = flexAlignRightBottom
                    .TextMatrix(0, 8) = "Responsable":      .ColWidth(8) = 1400:    .ColAlignment(8) = flexAlignLeftBottom:         .Row = 0: .Col = 8: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 9) = "Receta":           .ColWidth(9) = 1050:    .ColAlignment(9) = flexAlignCenterBottom:       .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterBottom
                    
                    .TextMatrix(0, 10) = "Tip. Prod":       .ColWidth(10) = 1000:   .ColAlignment(10) = flexAlignLeftBottom:        .Row = 0: .Col = 10: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 11) = "Insumo":          .ColWidth(11) = 3000:   .ColAlignment(11) = flexAlignLeftBottom:        .Row = 0: .Col = 11: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 12) = "U.M.":            .ColWidth(12) = 450:    .ColAlignment(12) = flexAlignLeftBottom:        .Row = 0: .Col = 12: .CellAlignment = flexAlignLeftBottom
                    '--------
                    M_ANCHO_COL = -50
        End Select

        Select Case T_CONSULTA
            Case 0, 1, 3
                    .TextMatrix(0, .Cols - 4) = "Cant.Teor":      .ColWidth(.Cols - 4) = 900 + M_ANCHO_COL: .ColAlignment(.Cols - 4) = flexAlignRightBottom:        .Row = 0: .Col = .Cols - 4: .CellAlignment = flexAlignRightBottom
                    .TextMatrix(0, .Cols - 3) = "Cant.Real":      .ColWidth(.Cols - 3) = 900 + M_ANCHO_COL: .ColAlignment(.Cols - 3) = flexAlignRightBottom:        .Row = 0: .Col = .Cols - 3: .CellAlignment = flexAlignRightBottom
                    .TextMatrix(0, .Cols - 2) = "Desvio":         .ColWidth(.Cols - 2) = 900 + M_ANCHO_COL: .ColAlignment(.Cols - 2) = flexAlignRightBottom:        .Row = 0: .Col = .Cols - 2: .CellAlignment = flexAlignRightBottom
                    .TextMatrix(0, .Cols - 1) = "% Desvio":       .ColWidth(.Cols - 1) = 900:  .ColAlignment(.Cols - 1) = flexAlignRightBottom:                     .Row = 0: .Col = .Cols - 1: .CellAlignment = flexAlignRightBottom
            
            Case 2
                    .TextMatrix(0, .Cols - 5) = "Cant.Teor":      .ColWidth(.Cols - 5) = 900 + M_ANCHO_COL: .ColAlignment(.Cols - 5) = flexAlignRightBottom:        .Row = 0: .Col = .Cols - 5: .CellAlignment = flexAlignRightBottom
                    .TextMatrix(0, .Cols - 4) = "Cant.Real":      .ColWidth(.Cols - 4) = 900 + M_ANCHO_COL: .ColAlignment(.Cols - 4) = flexAlignRightBottom:        .Row = 0: .Col = .Cols - 4: .CellAlignment = flexAlignRightBottom
                    .TextMatrix(0, .Cols - 3) = "Desvio":         .ColWidth(.Cols - 3) = 900 + M_ANCHO_COL: .ColAlignment(.Cols - 3) = flexAlignRightBottom:        .Row = 0: .Col = .Cols - 3: .CellAlignment = flexAlignRightBottom
                    .TextMatrix(0, .Cols - 2) = "% Desvio":       .ColWidth(.Cols - 2) = 900:  .ColAlignment(.Cols - 2) = flexAlignRightBottom:                     .Row = 0: .Col = .Cols - 2: .CellAlignment = flexAlignRightBottom
                    .TextMatrix(0, .Cols - 1) = "% Consumo":      .ColWidth(.Cols - 1) = 1000: .ColAlignment(.Cols - 1) = flexAlignRightBottom:                     .Row = 0: .Col = .Cols - 4: .CellAlignment = flexAlignRightBottom
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
'* Tipo             : FUNCION
'* Descripcion      : ESTA FUNCION CONVERTIRA AL FORMATO, ESTA FUNCION DEVUELVE UNA CADENA
'* Paranetros       : NOMBRE          |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    S_MONTO         |  Double    |  VALOR AL QUE SE LE DARA FORMATO
'*                    Band_Total_gral |  Boolean   |
'*                    Q_POS           |  Integer   |
'* Devuelve         : String
'*****************************************************************************************************
Private Function PONER_FORMATO(S_MONTO As Double, Optional Band_Total_gral As Boolean = False, Optional Q_POS As Integer = -1) As String
    If S_MONTO = 0 Then
            PONER_FORMATO = "0.00"
        Exit Function
    End If
    
    PONER_FORMATO = Format(S_MONTO, FORMAT_MONTO)
End Function

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nOrden As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nSQLNotIn As String
    Dim Q_ROW As Long
    If Col <> 2 Then Exit Sub
    
    ' DE LOS REGISTROS YA SELECCIONADOS
    nSQLNotIn = GENERAR_SQL_ID(fg(Index), 1, " AND alm_inventario.id", "NOT IN")
    
    If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
        nSQLNotIn = nSQLNotIn & " AND UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' "
    End If

    Select Case Index
        Case 0 ' producto
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "proddesc":   xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Familia":      xCampos(1, 1) = "famdesc":    xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
            
            nSQL = "SELECT alm_inventario.id, alm_inventario.descripcion AS proddesc, mae_familia.descripcion AS famdesc " _
                + vbCr + " FROM alm_inventario INNER JOIN mae_familia ON alm_inventario.idfam = mae_familia.id " _
                + vbCr + " WHERE (((alm_inventario.tippro) = 3)) AND alm_inventario.activo = -1 " + nSQLNotIn _
                + vbCr + " ORDER BY mae_familia.descripcion, alm_inventario.descripcion; "
            nTitulo = "Buscando Producto"
            nOrden = "proddesc"
            nCampoBusca = "proddesc"
        
        Case 1 ' INSUMO
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "insdesc":     xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Tip. Prod":        xCampos(1, 1) = "tipprodesc":      xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Familia":          xCampos(2, 1) = "famdesc":      xCampos(2, 2) = "1500":   xCampos(2, 3) = "C"
            
            nSQL = "SELECT alm_inventario.id, alm_inventario.descripcion AS insdesc, mae_tipoproducto.descripcion AS tipprodesc, mae_familia.descripcion AS famdesc " _
                 + vbCr + " FROM (pro_receta INNER JOIN ((mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) INNER JOIN pro_recetains ON alm_inventario.id = pro_recetains.iditem) ON pro_receta.id = pro_recetains.idrec) LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id " _
                + vbCr + " WHERE alm_inventario.activo = -1 " + nSQLNotIn _
                 + vbCr + " GROUP BY alm_inventario.id, alm_inventario.descripcion, mae_tipoproducto.descripcion, mae_familia.descripcion " _
                 + vbCr + " ORDER BY alm_inventario.descripcion, mae_tipoproducto.descripcion, mae_familia.descripcion;"
             nTitulo = "Buscando Insumos"
             nOrden = "insdesc"
             nCampoBusca = "insdesc"
        
        '*****************************************************************************
        Case 2 ' ESTADO
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "desestado":     xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":               xCampos(1, 1) = "idestado":   xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
            
            nSQLNotIn = GENERAR_SQL_ID(fg(Index), 1, " AND mae_estados.id", "NOT IN")
            
            nSQL = "SELECT mae_estados.id AS idestado, mae_estados.descripcion AS desestado " _
                + vbCr + "FROM mae_estados " _
                + vbCr + "WHERE (mae_estados.id is not null)" & nSQLNotIn
                
             nTitulo = "Buscando Estados"
             nOrden = "desestado"
             nCampoBusca = "desestado"
        
        '*****************************************************************************
    
    End Select

    Dim xRs As New ADODB.Recordset
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, nOrden, nCampoBusca, Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    For Q_ROW = 0 To xRs.Fields.Count - 1
        fg(Index).TextMatrix(fg(Index).Row, Q_ROW + 1) = xRs.Fields(Q_ROW) & ""
        
    Next Q_ROW
        
    If fg(Index).Row = fg(Index).Rows - 1 Then fg(Index).AddItem ""
    fg(Index).Row = fg(Index).Rows - 1: fg(Index).Col = 2
        
SALIR:
    Set xRs = Nothing
    Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Fg_CellButtonClick(" + CStr(Index) + ")", True, "Error..."
End Sub

Private Sub fg_DblClick(Index As Integer)
    Fg_CellButtonClick Index, fg(Index).Rows - 1, 2
End Sub

Private Sub Fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If fg(Index).Row = -2 Then Exit Sub
    Select Case KeyCode
        Case 45  'INSERTAR REGI
            fg(Index).AddItem ""
            fg(Index).Row = fg(Index).Rows - 1: fg(Index).Col = 2
        
        Case 46 'SUPRIMIR/DELETE
            If fg(Index).Rows - 1 >= 2 Then
                fg(Index).RemoveItem fg(Index).Row
                fg(Index).Row = fg(Index).Rows - 1: fg(Index).Col = 2
            Else
                LimpiarGrid fg(Index), True
                GRID_COMBOLIST fg(Index)
            End If
    End Select
End Sub

'*****************************************************************************************************
'* Nombre           : ESTILO_CONSULTA
'* Tipo             : FUNCION
'* Descripcion      :
'* Paranetros       :
'* Devuelve         : Integer
'*****************************************************************************************************
Private Function ESTILO_CONSULTA() As Integer
    Dim M_ESTILO As Integer
    If opt_consulta(0).Value = True Then ' RESUMEN
        If opt_tipo(0).Value = True Then M_ESTILO = 0 ' X PRODUCTO
        If opt_tipo(1).Value = True Then M_ESTILO = 1 ' X FAMILIA
        If opt_tipo(2).Value = True Then M_ESTILO = 2 ' X INSUMO
    Else                                 ' DETALLE
        If opt_tipo(0).Value = True Then M_ESTILO = 3 ' X PRODUCTO
        If opt_tipo(1).Value = True Then M_ESTILO = 4 ' X FAMILIA
        If opt_tipo(2).Value = True Then M_ESTILO = 5 ' X INSUMO
    End If
    ESTILO_CONSULTA = M_ESTILO
End Function

'*****************************************************************************************************
'* Nombre           : PosicionarProgBar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub PosicionarProgBar()
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    Me.PgBar.Value = 1
    FraProgreso.Visible = True
End Sub

Private Sub opt_consulta_Click(Index As Integer)
    If Index = 0 Then
        habilitar opt_tipo, True
    Else
        opt_tipo(0).Value = True
        opt_tipo(1).Enabled = False
        opt_tipo(2).Enabled = False
    End If
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Index = 0 Then PopupMenu Menu1
        If Index = 1 Then PopupMenu Menu2
    End If
End Sub

' DEL PRODUCTO
Private Sub Menu1_1_Click()
    fg_DblClick 0
End Sub

Private Sub Menu1_3_Click()
    Fg_KeyDown 0, 46, 0
End Sub

' DE LOS INSUMOS
Private Sub Menu2_1_Click()
    fg_DblClick 1
End Sub

Private Sub Menu2_3_Click()
    Fg_KeyDown 1, 46, 0
End Sub

'*****************************************************************************************************
'* Nombre           : pExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportarExcel()
On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO + " ", T_RPT_PERIODO, , "Producción"
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 5 Then pExportarExcel
    If Button.Index = 6 Then pImprimir
    If Button.Index = 8 Then
        BAND_INTERRUMPIR = True
        Unload Me
    End If
End Sub
