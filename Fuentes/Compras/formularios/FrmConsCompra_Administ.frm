VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsCompra_Administ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras - Consulta de Compras"
   ClientHeight    =   8010
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   1590
      TabIndex        =   20
      Top             =   4650
      Visible         =   0   'False
      Width           =   5730
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   21
         Top             =   420
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   27
         Top             =   120
         Width           =   1530
      End
      Begin VB.Shape Shape1 
         Height          =   750
         Left            =   45
         Top             =   30
         Width           =   5655
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Compras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1455
         TabIndex        =   23
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label6 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   150
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2205
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   11715
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec2 
         Height          =   300
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
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
         Valor           =   "16/10/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec1 
         Height          =   300
         Left            =   720
         TabIndex        =   1
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
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
         Valor           =   "16/10/2007"
      End
      Begin VB.Frame Frame5 
         Caption         =   "Seleccionar"
         Height          =   810
         Left            =   6375
         TabIndex        =   26
         Top             =   120
         Width           =   1395
         Begin VB.OptionButton OptPag 
            Caption         =   "Pagados"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   570
            Width           =   1215
         End
         Begin VB.OptionButton OptPend 
            Caption         =   "Pendientes"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   375
            Width           =   1095
         End
         Begin VB.OptionButton OptTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   180
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Seleccionar "
         Height          =   810
         Left            =   3700
         TabIndex        =   25
         Top             =   120
         Width           =   1290
         Begin VB.OptionButton OptVenc 
            Caption         =   "F. Venc."
            Height          =   255
            Left            =   135
            TabIndex        =   30
            Top             =   450
            Width           =   990
         End
         Begin VB.OptionButton OptEmi 
            Caption         =   "F. Emi."
            Height          =   195
            Left            =   135
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Consulta"
         Height          =   810
         Left            =   2130
         TabIndex        =   19
         Top             =   120
         Width           =   1470
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Detallado"
            Height          =   195
            Left            =   195
            TabIndex        =   29
            Top             =   465
            Width           =   1065
         End
         Begin VB.OptionButton OptResum 
            Caption         =   "Resumen"
            Height          =   195
            Left            =   195
            TabIndex        =   3
            Top             =   255
            Value           =   -1  'True
            Width           =   1155
         End
      End
      Begin VB.CheckBox ChkPagada 
         Caption         =   "Pagadas"
         Height          =   195
         Left            =   11745
         TabIndex        =   18
         Top             =   495
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CheckBox ChkPend 
         Caption         =   "Pedientes"
         Height          =   195
         Left            =   11745
         TabIndex        =   17
         Top             =   225
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Moneda"
         Height          =   810
         Left            =   5090
         TabIndex        =   16
         Top             =   120
         Width           =   1185
         Begin VB.OptionButton OptDol 
            Caption         =   "Dólares"
            Height          =   195
            Left            =   90
            TabIndex        =   32
            Top             =   570
            Width           =   915
         End
         Begin VB.OptionButton OptSol 
            Caption         =   "Soles"
            Height          =   195
            Left            =   90
            TabIndex        =   31
            Top             =   390
            Width           =   750
         End
         Begin VB.OptionButton OptMonTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   90
            TabIndex        =   5
            Top             =   195
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg3 
         Height          =   990
         Left            =   105
         TabIndex        =   10
         Top             =   1050
         Width           =   5700
         _cx             =   10054
         _cy             =   1746
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsCompra_Administ.frx":0000
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
      Begin VB.CheckBox ChkMostrarItem 
         Caption         =   "Mostrar item"
         Height          =   195
         Left            =   7830
         TabIndex        =   8
         Top             =   735
         Width           =   1275
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   990
         Left            =   5925
         TabIndex        =   9
         Top             =   1050
         Width           =   5700
         _cx             =   10054
         _cy             =   1746
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
         FormatString    =   $"FrmConsCompra_Administ.frx":0068
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
      Begin VB.CommandButton CmdBusProducto 
         Height          =   225
         Left            =   8295
         Picture         =   "FrmConsCompra_Administ.frx":00E5
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   435
         Width           =   225
      End
      Begin VB.TextBox TxtIdTipProd 
         Height          =   300
         Left            =   7830
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "TxtIdTipProd"
         Top             =   405
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   285
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   690
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Producto"
         Height          =   195
         Left            =   7830
         TabIndex        =   13
         Top             =   165
         Width           =   1230
      End
      Begin VB.Label lblTipProducto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTipProducto"
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
         Left            =   8550
         TabIndex        =   12
         Top             =   405
         Width           =   3060
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Align           =   2  'Align Bottom
      Height          =   5415
      Left            =   0
      TabIndex        =   24
      Top             =   2595
      Width           =   11775
      _cx             =   20770
      _cy             =   9551
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
      Rows            =   3
      Cols            =   24
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsCompra_Administ.frx":0217
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
      Height          =   570
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1005
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
               Picture         =   "FrmConsCompra_Administ.frx":04DA
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":0A1E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":0DB0
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":0F34
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":1388
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":14A0
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":19E4
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":1F28
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":203C
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":2150
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":25A4
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":2710
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":2C58
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCompra_Administ.frx":2F72
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmConsCompra_Administ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCONSCOMPRA.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA LA EMISION DE REPORTES DE COMPRA, PERMITE HACER LA CONSULTA POR
'*                    PROVEEDOR Y ITEM, ADEMAS DE PERMITIR SELECCIONAR RANGO DE FECHAS PARA LA CONSULTA
'* DISEÑADO POR     : JOHAN CASTRO
'* ULTIMA REVISION  : 15/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit
Dim RsCons As New ADODB.Recordset, rsTemp As New ADODB.Recordset
Dim RsProv As New ADODB.Recordset, RsItem As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset, Rs2 As New ADODB.Recordset, Rs3 As New ADODB.Recordset
Dim Rs4 As New ADODB.Recordset, Rs5 As New ADODB.Recordset, RsPreProm As New ADODB.Recordset
Dim vSubTotal_Sol As Double, vSubTotal_Dol As Double, vSumCol_Sol As Double, vSumCol_Dol As Double
Dim vSubTotalAbon_Sol As Double
Dim vSumColSal_Sol As Double, vSumColSal_Dol As Double, vSubTotSaldo_Dol As Double, vSubTotSaldo_Sol As Double
Dim vSumCol_AbonTot_Sol As Double, vSumCol_AbonAbono_Sol As Double, vSumCol_AboSaldo_Sol As Double
Dim vSumCol_Cant As Double, vNomProv As String, vAgregaSubTotal As String
Dim vSumColGen_SubTot_Dol As Double, vSumColGen_SubTot_Sol As Double, vFlag As Integer
Dim vSumColGen_AbonTot_Sol As Double, vSumColGen_AbonAbon As Double, vSumColGen_AbonSaldo_Sol As Double
Dim vSumColGen_Cant As Double

Dim vFlag2 As Integer, vCantDocum As Long, vCanItemRes As Double
Dim BAND_INTERRUMPIR As Boolean
Dim vIndicadorConsul As String, vIndicadorConsulHastaDetalleReal As String
Dim vStrCons As String, vFormatString As String, vFormatStrGridItem As String, vFormatGridProv As String
Dim CaracteresNumericos As String
Dim vIndicaPreProm As Integer, X As Integer
Dim vSubTotal_Dol_Convert_Sol As Double, vSubTotalSaldo_Dol_Convert_Sol As Double
Dim vSumTotGenSol As Double, vSumTotGenDol As Double, vSumTotGenSaldoSol As Double, vSumTotGenSaldoDol As Double
'--VARIABLES PARA EL REPORTE
Dim vtitreporte As String, vrangofec As String

'*****************************************************************************************************
'* Nombre           : fverifFecha
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VERIFICA QUE EL RANGO DE FECHAS SEAN LOS CORRECTOS, DEVUELVE FALSO SI EL RANGO
'*                    DE FECHAS NO ES VALIDO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function fverifFecha() As Boolean
    If IsDate(TxtFec1.Valor) = False Then
        MsgBox "Ingrese una fecha de inicio válida", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
        TxtFec1.SetFocus
        fverifFecha = False
        Exit Function
    Else
        fverifFecha = True
    End If
    If IsDate(TxtFec2.Valor) = False Then
        MsgBox "Ingrese una fecha final válida", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
        TxtFec2.SetFocus
        fverifFecha = False
        Exit Function
    Else
        fverifFecha = True
    End If
    If CDate(TxtFec1.Valor) > CDate(TxtFec2.Valor) Then
        MsgBox "La fecha de inicio es mayor que la ultima fecha", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
        TxtFec1.SetFocus
        fverifFecha = False
        Exit Function
    Else
        fverifFecha = True
    End If
End Function

'*****************************************************************************************************
'* Nombre           : VerifSiSeCalculaPreProm
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : VERIFICA SI SE CALCULARA EL PRECIO PROMEDIO CUANDO SE DETALLE LA COMPRA POR ITEM
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub VerifSiSeCalculaPreProm()
    vIndicaPreProm = 0
    Dim vContar_Local As Integer
    If ChkMostrarItem.Value = 1 Then
        For X = 1 To Fg2.Rows - 1
            If Trim(Fg2.TextMatrix(X, 3)) <> "" Then
                vContar_Local = vContar_Local + 1
            End If
        Next
        If vContar_Local = 0 Or vContar_Local > 1 Then 'NO CALCULAR EL PRE PROM
            vIndicaPreProm = 0
        Else
            vIndicaPreProm = 1
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : fSqlCalPreProm
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : HALLA EL PRECIO PROMEDIO DE UN ITEM
'* Paranetros       : NOMBRE     |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pIdPro     |  LONG        |  ESPECIFICA EL ID DEL PROVEEDOR
'*                    pIdTipProd |  LONG        |  ESPECIFICA EL TIPO DE ITEM QUE SE ESTA CONSULTANDO
'*                    pIdItem    |  LONG        |  ESPECIFICA EL ID DEL ITEM QUE SE ESTA CONSULTANDO
'*                    pFlagGen   |  BOOLEAN     |  ESPECIFICA SI SE TOTALIZARAN LOS PRECIOS PROMEDIOS
'* Devuelve         :
'*****************************************************************************************************
Private Sub fSqlCalPreProm(pIdPro As Long, pIdTipProd As Long, pIdItem As Long, Optional pFlagGen As Boolean = False)
'    vStrCons = "SELECT Avg(com_comprasdet.preuni) AS pre_prom" _
'        & " FROM (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) RIGHT JOIN (com_compras LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON alm_inventario.id = com_comprasdet.iditem" _
'        & " GROUP BY com_compras.idpro, mae_tipoproducto.id, alm_inventario.id" _
'        & " Having Avg(com_comprasdet.preuni) Is Not Null And com_compras.idpro = " & pIdPro & " And mae_tipoproducto.id = " & pIdTipProd & " And alm_inventario.id = " & pIdItem & ""
    'fSqlCalPreProm = vStrCons
    Set RsPreProm = Nothing
    vStrCons = pGenerarConsulta(pIdPro, pIdTipProd, pIdItem, pFlagGen)
    RST_Busq RsPreProm, vStrCons, xCon
End Sub

'*****************************************************************************************************
'* Nombre           : LimpiarVaribles
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : INICIALIZA A 0 LA VARIABLES NECESARIAS PARA EL FUNCIONAMIENTO DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub LimpiarVaribles()
    If OptResum.Value = True Then
    Else
        If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
            vAgregaSubTotal = ""
            vFlag = 0: vFlag2 = 0
            vSumColGen_SubTot_Dol = 0
            vSumColGen_SubTot_Sol = 0
            vSumTotGenSaldoDol = 0
            vSumTotGenSaldoSol = 0
            '--SUMTOTAL GENERAL PARTE DEL ABONO
            vSumColGen_AbonTot_Sol = 0
            vSumColGen_AbonAbon = 0
            vSumColGen_AbonSaldo_Sol = 0
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : FrozenCelGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE CONGELAR COLUMNAS DEL CONTROL FlexGrid Fg1
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pCol      |  INTEGER          |  ESPECIFICA LA COLUMNA QUE SE VA A CONGELAR
'* Devuelve         :
'*****************************************************************************************************
Sub FrozenCelGrid(pCol As Integer)
    Fg1.FrozenCols = pCol
End Sub

'*****************************************************************************************************
'* Nombre           : formatTextCeldaGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CAMBIA EL FORMATO DE LA CELDA ESPECIFICADA DEL CONTROL FlexGrid Fg1
'* Paranetros       : NOMBRE     |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pFila      |  LONG        |  ESPECIFICA LA FILA ACTUAL
'*                    pCol       |  INTEGER     |  ESPECIFICA LA COLUMNA ACTUAL
'*                    pColorText |  LONG        |  ESPECIFICA EL COLOR DE TEXTO QUE SE APLICARA A LA
'*                                                 CELDA
'* Devuelve         :
'*****************************************************************************************************
Sub formatTextCeldaGrid(pFila As Long, pCol As Integer, pColorText As Long)
    Fg1.Row = pFila: Fg1.Col = pCol
    Fg1.CellForeColor = pColorText
End Sub

'*****************************************************************************************************
'* Nombre           : UnirCeldas
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE COMBINAR CELDAS EN EL CONTROL FlexGrid Fg1
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pFila     |  LONG             |  ESPECIFICA EL ID DE LA FILA
'*                    pColRang1 |  INTEGER          |  ESPECIFICA EL RANGO INICIAL
'*                    pColRang2 |  INTEGER          |  ESPECIFICA EL RANGO FINAL
'*                    pCadena   |  STRING           |  ESPECIFICA LA CADENA QUE SE UTILIZARA PARA LAS
'*                              |                   |  CELDAS A UNIRSE
'* Devuelve         :
'*****************************************************************************************************
Sub UnirCeldas(pFila As Long, pColRang1 As Integer, pColRang2 As Integer, pCadena As String)
    With Fg1
        .MergeCells = flexMergeFree
        .Row = pFila
        .MergeRow(pFila) = True
        .Select pFila, pColRang1, pFila, pColRang2
        .CellAlignment = flexAlignLeftCenter
        .Cell(flexcpText, pFila, pColRang1, pFila, pColRang2) = pCadena
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : fSqlTipProdDife_Resume
'* Tipo             : FUNCCION
'* Descripcion      : DEVUELVE UNA CADENA SQL QUE MOSTRARA LOS ITEMS COMPRADOS EN EL RANGO DE FECHA
'*                    ESPECIFICADO, CLASIFICANDOLO POR TIPO DE PRODUCTO, DEVUELVE UNA CADENA
'* Paranetros       :
'* Devuelve         : String
'*****************************************************************************************************
Function fSqlTipProdDife_Resume() As String
    vStrCons = "SELECT DISTINCT mae_tipoproducto.id AS idtipprod, mae_tipoproducto.descripcion AS destipprod, alm_inventario.id AS iditem, alm_inventario.descripcion AS desitem, com_comprasdet.idunimed" _
       & " FROM (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) RIGHT JOIN (com_compras LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON alm_inventario.id = com_comprasdet.iditem" _
       & " Where (((mae_tipoproducto.id) Is Not Null)) AND com_compras.fchdoc BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "')"
       If Trim(TxtIdTipProd.Text) <> "" Then
            vStrCons = vStrCons & " AND mae_tipoproducto.id = " & Val(TxtIdTipProd.Text) & ""
       End If
       vStrCons = vStrCons & " ORDER BY mae_tipoproducto.descripcion, alm_inventario.descripcion"
    If OptVenc.Value = True Then
        vStrCons = Replace(vStrCons, "com_compras.fchdoc", "com_compras.fchven")
    End If
    fSqlTipProdDife_Resume = vStrCons
End Function

'*****************************************************************************************************
'* Nombre           : SumColumna
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : SUMA LAS COLUMNAS DEL CONTROL FlexGrid Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub SumColumna()
    vCantDocum = 0: vSumCol_Dol = 0: vSumColSal_Dol = 0: vSumCol_Sol = 0: vSumColSal_Sol = 0
    vSumCol_AbonTot_Sol = 0: vSumCol_AbonAbono_Sol = 0: vSumCol_AboSaldo_Sol = 0
    If OptResum.Value = True Then
'    1         2            3            4         5             6
'Tipo Doc., Cliente, Num. Documento, Fec. Doc., Fec. Venc., Cond. Pago,
'    7          8         9           10          11      12         13         14
'Dias Atras., Moneda, Tipo Cambio, Tipo Producto, Item, Uni. Med., Cantidad, Prec. Unit. $
'       15       16            17            18            19         20          21        22
'Imp. Total $, Saldo $, Prec. Unit S/., Imp. Total S/., Saldo S/., Total S/., Abono S/., Saldo S/.
        If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
            For X = 2 To Fg1.Rows - 1
                vCantDocum = vCantDocum + Val(Format(Fg1.TextMatrix(X, 4), "##0.00")) '
                vSumCol_Dol = vSumCol_Dol + Val(Format(Fg1.TextMatrix(X, 16), "#####0.00")) '
                vSumColSal_Dol = vSumColSal_Dol + Val(Format(Fg1.TextMatrix(X, 17), "#####0.00")) '
                vSumCol_Sol = vSumCol_Sol + Val(Format(Fg1.TextMatrix(X, 19), "#####0.00")) '
                vSumColSal_Sol = vSumColSal_Sol + Val(Format(Fg1.TextMatrix(X, 20), "#####0.00")) '
                vSumCol_AbonTot_Sol = vSumCol_AbonTot_Sol + Val(Format(Fg1.TextMatrix(X, 21), "#####0.00")) '
                vSumCol_AbonAbono_Sol = vSumCol_AbonAbono_Sol + Val(Format(Fg1.TextMatrix(X, 22), "#####0.00")) '
                vSumCol_AboSaldo_Sol = vSumCol_AboSaldo_Sol + Val(Format(Fg1.TextMatrix(X, 23), "#####0.00")) '
            Next
            Fg1.AddItem ""
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = "Total Gen.:" '
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = vCantDocum '
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(vSumCol_Dol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(vSumColSal_Dol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumCol_Sol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(vSumColSal_Sol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(vSumCol_AbonTot_Sol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 22) = Format(vSumCol_AbonAbono_Sol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 23) = Format(vSumCol_AboSaldo_Sol, FORMAT_MONTO) '
            
            formatTextCeldaGrid Fg1.Rows - 1, 3, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 4, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 16, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 17, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 19, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 20, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 21, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 22, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 23, &H800000 '
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ProcResumen
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ProcResumen()
    vFlag = 0: vFlag2 = 0
    If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
        vtitreporte = "REPORTE RESUMIDO"
        If Rs1.RecordCount > 0 Then ' RECORREMOS LOS PROVEDORES
            FraProgreso.Visible = True
            PgBar.Value = 0
            Rs1.MoveFirst
            Do While Not Rs1.EOF    ' RECORREMOS LOS PROVEDORES
                If BAND_INTERRUMPIR = True Then
                    FraProgreso.Visible = False
                    Exit Sub
                End If
                DoEvents
                '----
                PgBar.Max = Rs1.RecordCount
                '---
                vFlag = 0: vCantDocum = 0
                RsCons.Filter = adFilterNone
                RsCons.Filter = "idpro = " & Val(Rs1("idpro")) & ""
                If RsCons.RecordCount > 0 Then
                    vCantDocum = RsCons.RecordCount
                    vFlag = 1
                    RsCons.MoveFirst
                    vSubTotal_Dol = 0: vSubTotal_Sol = 0: vSubTotSaldo_Dol = 0: vSubTotSaldo_Sol = 0
                    vSubTotal_Dol_Convert_Sol = 0: vSubTotalSaldo_Dol_Convert_Sol = 0
                    Do While Not RsCons.EOF
                        vSubTotal_Dol = vSubTotal_Dol + NulosN(RsCons("subtotal_dol"))
                        If NulosN(RsCons("idmon")) = 2 Then
                            vSubTotal_Dol_Convert_Sol = vSubTotal_Dol_Convert_Sol + NulosN(RsCons("subtotal_sol"))
                        End If
                        If NulosN(RsCons("idmon")) = 1 Then
                            vSubTotal_Sol = vSubTotal_Sol + NulosN(RsCons("subtotal_sol"))
                        End If
                        vSubTotSaldo_Dol = vSubTotSaldo_Dol + NulosN(RsCons("saldo_dol"))
                        If NulosN(RsCons("idmon")) = 2 Then
                            vSubTotalSaldo_Dol_Convert_Sol = vSubTotalSaldo_Dol_Convert_Sol + NulosN(RsCons("saldo_sol"))
                        End If
                        If NulosN(RsCons("idmon")) = 1 Then
                            vSubTotSaldo_Sol = vSubTotSaldo_Sol + Val(RsCons("saldo_sol"))
                        End If
                        RsCons.MoveNext
                    Loop
                    ' AQUI PINTAR GRID PA RESUMEN
                    PintarGridResumen_Nuevo
                End If
                Rs1.MoveNext
                If PgBar.Value <= PgBar.Max Then
                    PgBar.Value = PgBar.Value + 1
                End If
            Loop
            FraProgreso.Visible = False
            If vFlag = 1 Then
                SumColumna
            End If
        End If
    ElseIf vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
        vtitreporte = "REPORTE RESUMIDO POR TIPO DE PRODUCTO"
        ' PROVEEDOR CON TIPO DE PRODUCTO
        vAgregaSubTotal = "": vSumColGen_SubTot_Dol = 0: vSumColGen_SubTot_Sol = 0: vSumColGen_AbonTot_Sol = 0
            
        If Rs1.RecordCount > 0 Then ' RECORREMOS LOS PROVEDORES
            If BAND_INTERRUMPIR = True Then
                FraProgreso.Visible = False
                Exit Sub
            End If
            
            FraProgreso.Visible = True
            PgBar.Max = Rs1.RecordCount
            PgBar.Value = 0
            vSumColGen_SubTot_Dol = 0: vSumColGen_SubTot_Sol = 0: vSumColGen_AbonTot_Sol = 0
            Rs1.MoveFirst
            
            Do While Not Rs1.EOF    ' RECORREMOS LOS PROVEDORES
                DoEvents
                If BAND_INTERRUMPIR = True Then
                    FraProgreso.Visible = False
                    Exit Sub
                End If
                If Rs4.RecordCount > 0 Then ' RECORREMOS LOS TIPO DE PRODU
                    vNomProv = "Proveedor: " & NulosC(Rs1("nombre"))
                    Rs4.MoveFirst
                    vSumCol_Dol = 0: vSumCol_Sol = 0: vSumCol_AbonTot_Sol = 0
                    
                    Do While Not Rs4.EOF    ' RECORREMOS LOS TIPO DE PRODU
                        vSumCol_Dol = 0: vSumCol_Sol = 0: vSumCol_AbonTot_Sol = 0
                        vFlag = 0
                        RsCons.Filter = adFilterNone
                        RsCons.Filter = "idpro = " & Val(Rs1("idpro")) & " AND idtipoprod = " & Val(Rs4("idtipprod")) & ""
                        
                        If RsCons.RecordCount > 0 Then
                            vFlag = 1
                            RsCons.MoveFirst
                            vSubTotal_Dol = 0: vSubTotal_Sol = 0: vSubTotalAbon_Sol = 0
                            
                            Do While Not RsCons.EOF
                                vSubTotal_Dol = vSubTotal_Dol + Val(RsCons("subtotal_dol"))
                                If NulosN(RsCons("idmon")) = 1 Then
                                    vSubTotal_Sol = vSubTotal_Sol + Val(RsCons("subtotal_sol"))
                                End If
                                vSubTotalAbon_Sol = vSubTotalAbon_Sol + NulosN(RsCons("subtotal_sol"))
                                RsCons.MoveNext
                            Loop
                            ' AQUI PINTAR GRID PA RESUMEN
                            PintarGridResumen_Nuevo
                            vSumCol_Dol = vSumCol_Dol + vSubTotal_Dol
                            vSumCol_Sol = vSumCol_Sol + vSubTotal_Sol
                            vSumCol_AbonTot_Sol = vSumCol_AbonTot_Sol + vSubTotalAbon_Sol
                        End If
                        Rs4.MoveNext
                    Loop
                    If vFlag = "1" Then
                        vFlag2 = 1
                        vSumColGen_SubTot_Dol = vSumColGen_SubTot_Dol + vSumCol_Dol
                        vSumColGen_SubTot_Sol = vSumColGen_SubTot_Sol + vSumCol_Sol
                        vSumColGen_AbonTot_Sol = vSumColGen_AbonTot_Sol + vSumCol_AbonTot_Sol
                    End If
                End If
                If PgBar.Value < PgBar.Max Then
                    PgBar.Value = PgBar.Value + 1
                End If
                Rs1.MoveNext
            Loop
            If vFlag2 = 1 Then ' AQUI TOTAL GENERAL
                vAgregaSubTotal = "TOTGEN"
                PintarGridResumen_Nuevo
            End If
            FraProgreso.Visible = False
        End If
    ElseIf vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
        vtitreporte = "REPORTE RESUMIDO POR ITEM"
        vSumColGen_SubTot_Dol = 0: vSumColGen_SubTot_Sol = 0: vSumColGen_AbonTot_Sol = 0
        vSumColGen_Cant = 0
        
        If Rs1.RecordCount > 0 Then ' RECORREMOS LOS PROVEEDORES
            FraProgreso.Visible = True
            PgBar.Max = Rs1.RecordCount
            PgBar.Value = 0
            Rs1.MoveFirst
            vNomProv = ""
            Do While Not Rs1.EOF    ' RECORREMOS LOS PROVEEDORES
                If BAND_INTERRUMPIR = True Then
                    FraProgreso.Visible = False
                    Exit Sub
                End If
                DoEvents
                vNomProv = NulosC(Rs1("nombre"))
                
                If Rs5.RecordCount > 0 Then  ' RECORREMOS LOS ITEM CON SUS TIP PROD
                    Rs5.MoveFirst
                    vFlag = 0
                    vSumCol_Dol = 0: vSumCol_Sol = 0: vSumCol_AbonTot_Sol = 0
                    vSumCol_Cant = 0
                    
                    Do While Not Rs5.EOF     ' RECORREMOS LOS ITEM CON SUS TIP PROD
                        vSubTotal_Dol = 0: vSubTotal_Sol = 0: vSubTotalAbon_Sol = 0: vCanItemRes = 0
                        RsCons.Filter = adFilterNone
                        RsCons.Filter = "idpro = " & Val(Rs1("idpro")) & " AND idtipoprod = " & Val(Rs5("idtipprod")) & " AND iditem = " & Val(Rs5("iditem")) & ""
                        
                        If RsCons.RecordCount > 0 Then
                            vFlag = 1
                            RsCons.MoveFirst
                            vSubTotal_Dol = 0: vSubTotal_Sol = 0: vSubTotalAbon_Sol = 0: vCanItemRes = 0
                            vSubTotal_Dol_Convert_Sol = 0
                            
                            Do While Not RsCons.EOF
                                vCanItemRes = vCanItemRes + NulosN(RsCons("canpro"))
                                vSubTotal_Dol = vSubTotal_Dol + Val(RsCons("subtotal_dol"))
                                If NulosN(RsCons("idmon")) = 2 Then
                                    vSubTotal_Dol_Convert_Sol = vSubTotal_Dol_Convert_Sol + NulosN(RsCons("subtotal_sol"))
                                End If
                                If NulosN(RsCons("idmon")) = 1 Then
                                    vSubTotal_Sol = vSubTotal_Sol + Val(RsCons("subtotal_sol"))
                                End If
                                vSubTotalAbon_Sol = vSubTotalAbon_Sol + NulosN(RsCons("subtotal_sol"))
                                RsCons.MoveNext
                            Loop
                            PintarGridResumen_Nuevo
                            vSumCol_Cant = vSumCol_Cant + vCanItemRes
                            vSumCol_Dol = vSumCol_Dol + vSubTotal_Dol
                            vSumCol_Sol = vSumCol_Sol + vSubTotal_Sol
                            vSumCol_AbonTot_Sol = vSumCol_AbonTot_Sol + vSubTotalAbon_Sol
                        End If
                        Rs5.MoveNext
                    Loop
                    If vFlag = 1 Then
                        vFlag2 = 1
                        vAgregaSubTotal = "SUBTOT"
                        PintarGridResumen_Nuevo
                        vSumColGen_Cant = vSumColGen_Cant + vSumCol_Cant
                        vSumColGen_SubTot_Dol = vSumColGen_SubTot_Dol + vSumCol_Dol
                        vSumColGen_SubTot_Sol = vSumColGen_SubTot_Sol + vSumCol_Sol
                        vSumColGen_AbonTot_Sol = vSumColGen_AbonTot_Sol + vSumCol_AbonTot_Sol
                    End If
                End If
                
                If PgBar.Value <= PgBar.Max Then
                    PgBar.Value = PgBar.Value + 1
                End If
                Rs1.MoveNext
            Loop ' FIN RECORREMOS LOS PROVEEDORES
            
            If vFlag2 = 1 Then
                vAgregaSubTotal = "TOTGEN"
                PintarGridResumen_Nuevo
            End If
            FraProgreso.Visible = False
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : PintarGridResumen_Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DA FORMATO AL CONTROL FlexGrid Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub PintarGridResumen_Nuevo()
    If vAgregaSubTotal = "SUBTOT" Then
        If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
            Fg1.AddItem ""
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = "Total:" '
            formatTextCeldaGrid Fg1.Rows - 1, 11, &H800000 '
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(vSumCol_Dol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumCol_Sol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(vSumCol_AbonTot_Sol, FORMAT_MONTO) '
            
            formatTextCeldaGrid Fg1.Rows - 1, 16, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 19, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 21, &H800000 '
        ElseIf vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
            Fg1.AddItem ""
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = "Total:" '
            formatTextCeldaGrid Fg1.Rows - 1, 11, &H800000 '
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(vSumCol_Cant, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(vSumCol_Dol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumCol_Sol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(vSumCol_AbonTot_Sol, FORMAT_MONTO) '
            
            formatTextCeldaGrid Fg1.Rows - 1, 14, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 16, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 19, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 21, &H800000 '
        End If
        vAgregaSubTotal = ""
        Exit Sub
    ElseIf vAgregaSubTotal = "TOTGEN" Then
        If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
            Fg1.AddItem ""
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = "Tot. Gen.:" '
            formatTextCeldaGrid Fg1.Rows - 1, 11, &H800000 '
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(vSumColGen_SubTot_Dol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumColGen_SubTot_Sol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(vSumColGen_AbonTot_Sol, FORMAT_MONTO) '
            
            formatTextCeldaGrid Fg1.Rows - 1, 16, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 19, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 21, &H800000 '
        ElseIf vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
            Fg1.AddItem ""
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = "Tot. Gen.:" '
            formatTextCeldaGrid Fg1.Rows - 1, 11, &H800000 '
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(vSumColGen_Cant, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(vSumColGen_SubTot_Dol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumColGen_SubTot_Sol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(vSumColGen_AbonTot_Sol, FORMAT_MONTO) '
            
            formatTextCeldaGrid Fg1.Rows - 1, 14, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 16, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 19, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 21, &H800000 '
            vSumColGen_Cant = 0
        End If
        vAgregaSubTotal = ""
        Exit Sub
    End If
    If vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
        If vNomProv <> "" Then
            If Fg1.TextMatrix(2, 1) = "" Then
                Fg1.TextMatrix(2, 1) = vNomProv
                formatTextCeldaGrid 2, 1, &H800000
            Else
                Fg1.AddItem ""
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = vNomProv
                formatTextCeldaGrid Fg1.Rows - 1, 1, &H800000
            End If
            vNomProv = ""
        End If
    End If
    If RsCons.RecordCount > 0 Then
        RsCons.MoveFirst
'    1         2            3            4         5             6
'Tipo Doc., Cliente, Num. Documento, Fec. Doc., Fec. Venc., Cond. Pago,
'    7          8         9           10          11      12         13         14
'Dias Atras., Moneda, Tipo Cambio, Tipo Producto, Item, Uni. Med., Cantidad, Prec. Unit. $
'       15       16            17            18            19         20          21        22
'Imp. Total $, Saldo $, Prec. Unit S/., Imp. Total S/., Saldo S/., Total S/., Abono S/., Saldo S/.
        With Fg1
             If .TextMatrix(2, 2) <> "" Then .AddItem "" '
            .TextMatrix(.Rows - 1, 2) = NulosC(RsCons("numruc")) 'AQUI VA EL RUC '
            .TextMatrix(.Rows - 1, 3) = NulosC(RsCons("nombre")) '
            .TextMatrix(.Rows - 1, 4) = vCantDocum  'AQUI CAMBIAR EL ENCABEZADO '
            .TextMatrix(.Rows - 1, 9) = NulosC(RsCons("simbolo")) '
            If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
                .TextMatrix(.Rows - 1, 11) = NulosC(RsCons("destipprod")) '
            End If
            If vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
                .TextMatrix(.Rows - 1, 11) = NulosC(RsCons("destipprod")) '
                .TextMatrix(.Rows - 1, 12) = NulosC(RsCons("desc_item")) '
                .TextMatrix(.Rows - 1, 13) = NulosC(RsCons("unid_abrev")) '
                .TextMatrix(.Rows - 1, 14) = Format(vCanItemRes, FORMAT_MONTO) '
            End If
            .TextMatrix(.Rows - 1, 16) = Format(vSubTotal_Dol, FORMAT_MONTO) '
            .TextMatrix(.Rows - 1, 17) = Format(vSubTotSaldo_Dol, FORMAT_MONTO) '
            .TextMatrix(.Rows - 1, 19) = Format(vSubTotal_Sol, FORMAT_MONTO) '
            .TextMatrix(.Rows - 1, 20) = Format(vSubTotSaldo_Sol, FORMAT_MONTO) '
            If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
                .TextMatrix(.Rows - 1, 21) = Format(vSubTotalAbon_Sol, FORMAT_MONTO) '
            Else
                .TextMatrix(.Rows - 1, 21) = Format(vSubTotal_Dol_Convert_Sol + vSubTotal_Sol, FORMAT_MONTO) '
            End If
            .TextMatrix(.Rows - 1, 22) = Format((vSubTotal_Dol_Convert_Sol + vSubTotal_Sol) - (vSubTotalSaldo_Dol_Convert_Sol + vSubTotSaldo_Sol), FORMAT_MONTO) '
            .TextMatrix(.Rows - 1, 23) = Format(vSubTotalSaldo_Dol_Convert_Sol + vSubTotSaldo_Sol, FORMAT_MONTO) '
        End With
    End If
End Sub


'*****************************************************************************************************
'* Nombre           : FormatGrid_Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DA FORMATO AL CONTROL FlexGrid Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub FormatGrid_Nuevo()
    UnirCeldas 0, 15, 17, "Dólares" '
    UnirCeldas 0, 18, 20, "Soles" '
    UnirCeldas 0, 21, 23, "Totales Soles" '
    If OptResum.Value = True Then
        With Fg1
            .ColWidth(1) = 0
            If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                .TextMatrix(1, 2) = "R.U.C.": .ColWidth(2) = 1100 '
            Else
                .ColWidth(2) = 0 '
            End If
            .ColWidth(3) = 2295 '
            .ColWidth(4) = 0 '
            If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                .TextMatrix(1, 4) = "# Doc.": .ColWidth(4) = 600 '
            End If
            .ColWidth(5) = 0 '
            .ColWidth(6) = 0 '
            .ColWidth(7) = 0 '
            .ColWidth(8) = 0 '
            .ColWidth(9) = 420 '
            .ColWidth(10) = 0 '
            If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                .ColWidth(11) = 0 '
            Else
                .ColWidth(11) = 1140 '
            End If
            If vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
                .ColWidth(12) = 2085 '
                .ColWidth(13) = 540 '
                .ColWidth(14) = 1065 '
            Else
                .ColWidth(12) = 0 '
                .ColWidth(13) = 0 '
                .ColWidth(14) = 0 '
            End If
            .ColWidth(15) = 0 '
'    1         2            3            4         5             6
'Tipo Doc., Cliente, Num. Documento, Fec. Doc., Fec. Venc., Cond. Pago,
'    7          8         9           10          11      12         13         14
'Dias Atras., Moneda, Tipo Cambio, Tipo Producto, Item, Uni. Med., Cantidad, Prec. Unit. $
'       15       16            17            18            19         20          21        22
'Imp. Total $, Saldo $, Prec. Unit S/., Imp. Total S/., Saldo S/., Total S/., Abono S/., Saldo S/.
            If OptMonTodos.Value = True Or OptDol.Value = True Then
                .ColWidth(16) = 1050 '
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                    .ColWidth(17) = 1050 '
                Else
                    .ColWidth(17) = 0 '
                End If
            Else
                .ColWidth(16) = 0 '
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                    .ColWidth(17) = 0 '
                Else
                    .ColWidth(17) = 1050 '
                End If
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
                    .ColWidth(17) = 0 '
                End If
            End If
            .ColWidth(18) = 0 '
            .ColWidth(19) = 1050 '
            If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                .ColWidth(20) = 1050 '
            Else
                .ColWidth(20) = 0 '
            End If
            If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                .ColWidth(21) = 1005 '
                .ColWidth(22) = 915 '
                .ColWidth(23) = 930 '
            Else
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                    .ColWidth(21) = 1050 '
                End If
                .ColWidth(22) = 0 '
                .ColWidth(23) = 0 '
            End If
        End With
        '--FROZEN
        If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
            FrozenCelGrid 9 '
        ElseIf vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
            FrozenCelGrid 0
        ElseIf vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
            FrozenCelGrid 12 '
        End If
    Else 'DETALLE
'    1         2            3            4         5             6
'Tipo Doc., Cliente, Num. Documento, Fec. Doc., Fec. Venc., Cond. Pago,
'    7          8         9           10          11      12         13         14
'Dias Atras., Moneda, Tipo Cambio, Tipo Producto, Item, Uni. Med., Cantidad, Prec. Unit. $
'       15       16            17            18            19         20          21        22
'Imp. Total $, Saldo $, Prec. Unit S/., Imp. Total S/., Saldo S/., Total S/., Abono S/., Saldo S/.

        With Fg1
            .ColWidth(1) = 830
            .ColWidth(2) = 405 '
            If vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
                .ColWidth(3) = 0 '
            ElseIf vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                .ColWidth(3) = 0 '
            Else
                .ColWidth(3) = 2295 '
            End If
            .TextMatrix(1, 4) = "Num. Documento"  '
            .ColWidth(4) = 1425  '
            .ColWidth(5) = 855 '
            .ColWidth(6) = 855 '
            .ColWidth(7) = 960 '
            If vIndicadorConsul = "AGR" Or vIndicadorConsul = "DET" Then
                If OptPend.Value = True Or OptTodos.Value = True Then
                    .ColWidth(8) = 900 '
                Else
                    .ColWidth(8) = 0 '
                End If
            Else
                If OptPend.Value = True Or OptTodos.Value = True Then
                    .ColWidth(8) = 960 '
                Else
                    .ColWidth(8) = 0 '
                End If
            End If
            .ColWidth(9) = 420 '
            .ColWidth(10) = 780 '
            .ColWidth(11) = 1140 '
            If vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
                .ColWidth(12) = 2085 '
                .ColWidth(13) = 540 '
                .ColWidth(14) = 1065 '
                If OptMonTodos.Value = True Or OptDol.Value = True Then
                    .ColWidth(15) = 720 '
                Else
                    .ColWidth(15) = 0 '
                End If
            Else
                .ColWidth(15) = 0 '
            End If
            If vIndicadorConsul = "AGR" Then
                .ColWidth(11) = 0 '
                .ColWidth(12) = 0 '
                .ColWidth(13) = 0 '
                .ColWidth(14) = 0 '
                .ColWidth(15) = 0 '
            End If
            If OptMonTodos.Value = True Or OptDol.Value = True Then
                .ColWidth(16) = 1050 '
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                    .ColWidth(17) = 1050 '
                Else
                    .ColWidth(17) = 0 '
                End If
            Else
                .ColWidth(16) = 0 '
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                    .ColWidth(17) = 0 '
                Else
                    .ColWidth(17) = 1050 '
                End If
                .ColWidth(17) = 0 '
            End If
            If vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
                .ColWidth(18) = 720 '
            Else
                .ColWidth(18) = 0 '
            End If
            .ColWidth(19) = 1050 '
            If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                .ColWidth(20) = 1050 '
            Else
                .ColWidth(20) = 0 '
            End If
            If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                .ColWidth(21) = 1005 '
                .ColWidth(22) = 915 '
                .ColWidth(23) = 930 '
            Else
                .ColWidth(21) = 0 '
                .ColWidth(22) = 0 '
                .ColWidth(23) = 0 '
            End If
        End With
        '--FROZEN
        FrozenCelGrid 6 '
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : fSqlTipProd_SinRepetir
'* Tipo             : FUNCCION
'* Descripcion      : ********************************************************************************
'* Paranetros       :
'* Devuelve         : STRING
'*****************************************************************************************************
Private Function fSqlTipProd_SinRepetir() As String
    vStrCons = "SELECT DISTINCT mae_tipoproducto.id AS idtipprod, mae_tipoproducto.descripcion" _
        & " FROM (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) RIGHT JOIN (com_compras LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON alm_inventario.id = com_comprasdet.iditem" _
        & " WHERE com_compras.fchdoc Between DateValue('" & Trim(TxtFec1.Valor) & "') And DateValue('" & Trim(TxtFec2.Valor) & "')"
    If Trim(TxtIdTipProd.Text) <> "" Then
        vStrCons = vStrCons & " AND mae_tipoproducto.id = " & Val(TxtIdTipProd.Text) & ""
    End If
    If OptVenc.Value = True Then
        vStrCons = Replace(vStrCons, "com_compras.fchdoc", "com_compras.fchven")
    End If
    fSqlTipProd_SinRepetir = vStrCons
End Function

'*****************************************************************************************************
'* Nombre           : fSqlTipProd_O_Item
'* Tipo             : FUNCCION
'* Descripcion      : ********************************************************************************
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pParam    |  INTEGER          |  ***********************************************
'* Devuelve         : STRING
'*****************************************************************************************************
Private Function fSqlTipProd_O_Item(pParam As Integer) As String
    If pParam = 1 Then ' SOLO LOS TIPO DE PRODUCTOS
        vStrCons = "SELECT DISTINCT com_compras.id AS idcomp, mae_tipoproducto.id AS idtipprod, mae_tipoproducto.descripcion" _
            & " FROM (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) RIGHT JOIN (com_compras LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON alm_inventario.id = com_comprasdet.iditem" _
            & " WHERE com_compras.fchdoc Between DateValue('" & Trim(TxtFec1.Valor) & "') And DateValue('" & Trim(TxtFec2.Valor) & "')"
        If Trim(TxtIdTipProd.Text) <> "" Then
            vStrCons = vStrCons & " AND mae_tipoproducto.id = " & Val(TxtIdTipProd.Text) & ""
        End If
        If OptVenc.Value = True Then
            vStrCons = Replace(vStrCons, "com_compras.fchdoc", "com_compras.fchven")
        End If
        vStrCons = vStrCons & " ORDER BY mae_tipoproducto.descripcion"
    ElseIf pParam = 2 Then
        vStrCons = "SELECT DISTINCT alm_inventario.id AS iditem, alm_inventario.descripcion AS desitem, mae_tipoproducto.id AS idtipprod, mae_tipoproducto.descripcion AS destipprod, com_compras.tipdoc, com_compras.numser, com_compras.numdoc" _
            & " FROM (mae_tipoproducto INNER JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) INNER JOIN (com_compras INNER JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON alm_inventario.id = com_comprasdet.iditem" _
            & " WHERE fchdoc BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "')" _
            & " ORDER BY com_compras.tipdoc, com_compras.numser, com_compras.numdoc, mae_tipoproducto.descripcion, alm_inventario.descripcion"
        If OptVenc.Value = True Then
            vStrCons = Replace(vStrCons, "com_compras.fchdoc", "com_compras.fchven")
        End If
    End If
    fSqlTipProd_O_Item = vStrCons
End Function

'*****************************************************************************************************
'* Nombre           : PintarGrid_OptDetalle
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ********************************************************************************
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pBool     |  BOOLEAN          |  ***********************************************
'*                    pCantReg  |  LONG             |  ***********************************************
'* Devuelve         :
'*****************************************************************************************************
Sub PintarGrid_OptDetalle(pBool As Boolean, pCantReg As Long)
    'PROC NUEVO
    If pBool = False And pCantReg >= 1 Then
        If vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
            If vAgregaSubTotal = "TOTGEN" Then
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = "Tot. Gen.:" '
                formatTextCeldaGrid Fg1.Rows - 1, 6, &H800000 '
                Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(vSumColGen_Cant, FORMAT_MONTO) '
                Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(vSumColGen_SubTot_Dol, FORMAT_MONTO) '
                Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumColGen_SubTot_Sol, FORMAT_MONTO) '
                formatTextCeldaGrid Fg1.Rows - 1, 14, &H800000 '
                formatTextCeldaGrid Fg1.Rows - 1, 16, &H800000 '
                formatTextCeldaGrid Fg1.Rows - 1, 19, &H800000 '
                vAgregaSubTotal = ""
                vSumColGen_Cant = 0
                vSumColGen_SubTot_Dol = 0: vSumColGen_SubTot_Sol = 0
                Exit Sub
            End If
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = "Proveedor: " & Trim(RsCons("nombre"))
            UnirCeldas Fg1.Rows - 1, 1, 5, Trim(RsCons("nombre")) '
            formatTextCeldaGrid Fg1.Rows - 1, 1, &H800000
        ElseIf vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
            If vAgregaSubTotal = "TOTGEN" Then
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = "Tot. Gen.:" '
                formatTextCeldaGrid Fg1.Rows - 1, 6, &H800000 '
                Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(vSumColGen_SubTot_Dol, FORMAT_MONTO) '
                Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(vSumTotGenSaldoDol, FORMAT_MONTO) '
                Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumColGen_SubTot_Sol, FORMAT_MONTO) '
                Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(vSumTotGenSaldoSol, FORMAT_MONTO) '
                
                formatTextCeldaGrid Fg1.Rows - 1, 16, &H800000 '
                formatTextCeldaGrid Fg1.Rows - 1, 17, &H800000 '
                formatTextCeldaGrid Fg1.Rows - 1, 19, &H800000 '
                formatTextCeldaGrid Fg1.Rows - 1, 20, &H800000 '
                '--TOTGEN PARTE DEL ABONO
                Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(vSumColGen_AbonTot_Sol, FORMAT_MONTO) '
                Fg1.TextMatrix(Fg1.Rows - 1, 22) = Format(vSumColGen_AbonAbon, FORMAT_MONTO) '
                Fg1.TextMatrix(Fg1.Rows - 1, 23) = Format(vSumColGen_AbonSaldo_Sol, FORMAT_MONTO) '
                formatTextCeldaGrid Fg1.Rows - 1, 21, &H800000 '
                formatTextCeldaGrid Fg1.Rows - 1, 22, &H800000 '
                formatTextCeldaGrid Fg1.Rows - 1, 23, &H800000 '
                vAgregaSubTotal = ""
                vSumColGen_SubTot_Dol = 0: vSumTotGenSaldoDol = 0: vSumColGen_SubTot_Sol = 0: vSumTotGenSaldoSol = 0
                vSumColGen_AbonTot_Sol = 0: vSumColGen_AbonAbon = 0: vSumColGen_AbonSaldo_Sol = 0
                Exit Sub
            End If
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = "Proveedor: " & Trim(RsCons("nombre"))
            UnirCeldas Fg1.Rows - 1, 1, 5, Trim(RsCons("nombre")) '
            formatTextCeldaGrid Fg1.Rows - 1, 1, &H800000
        ElseIf vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Trim(RsCons("nombre"))
            UnirCeldas Fg1.Rows - 1, 1, 5, Trim(RsCons("nombre")) '
            formatTextCeldaGrid Fg1.Rows - 1, 1, &H800000
        End If
        RsCons.MoveFirst
        vSumCol_Cant = 0: vSumCol_Dol = 0: vSumCol_Sol = 0: vSumColSal_Dol = 0: vSumColSal_Sol = 0
        vSumCol_AbonTot_Sol = 0: vSumCol_AbonAbono_Sol = 0: vSumCol_AboSaldo_Sol = 0
        vSubTotal_Dol_Convert_Sol = 0: vSubTotalSaldo_Dol_Convert_Sol = 0
        
        Do While Not RsCons.EOF
            If Trim(Fg1.TextMatrix(2, 2)) <> "" Then Fg1.AddItem "" '
            With Fg1
                .TextMatrix(.Rows - 1, 1) = NulosC(RsCons("registro"))
                .TextMatrix(.Rows - 1, 2) = NulosC(RsCons("abrev")) '
                .TextMatrix(.Rows - 1, 3) = NulosC(RsCons("nombre")) '
                .TextMatrix(.Rows - 1, 4) = NulosC(RsCons("numerodoc")) '
                .TextMatrix(.Rows - 1, 5) = Format(RsCons("fchdoc"), FORMAT_DATE) '
                .TextMatrix(.Rows - 1, 6) = Format(RsCons("fchven"), FORMAT_DATE) '
                .TextMatrix(.Rows - 1, 7) = NulosC(RsCons("desccond")) '
                If NulosN(RsCons("impsal")) > 0 Then
                    .TextMatrix(.Rows - 1, 8) = Abs(DateDiff("d", Date, NulosC(RsCons("fchven")))) '
                End If
                .TextMatrix(.Rows - 1, 9) = NulosC(RsCons("simbolo")) '
                .TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(RsCons("impcom")), FORMAT_IMPUESTO) '
                If vIndicadorConsulHastaDetalleReal = "DET" Then
                    .TextMatrix(.Rows - 1, 11) = NulosC(RsCons("destipprod")) '
                    .TextMatrix(.Rows - 1, 12) = NulosC(RsCons("desc_item")) '
                    .TextMatrix(.Rows - 1, 13) = NulosC(RsCons("unid_abrev")) '
                    .TextMatrix(.Rows - 1, 14) = Format(NulosN(RsCons("canpro")), FORMAT_MONTO) '
                    .TextMatrix(.Rows - 1, 15) = Format(NulosN(RsCons("preuni_dol")), FORMAT_MONTO) '
                End If
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
                    .TextMatrix(Fg1.Rows - 1, 16) = Format(vSubTotal_Dol, FORMAT_MONTO) '
                Else
                    .TextMatrix(Fg1.Rows - 1, 16) = Format(NulosN(RsCons("subtotal_dol")), FORMAT_MONTO) '
                End If
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                    .TextMatrix(.Rows - 1, 17) = Format(NulosN(RsCons("saldo_dol")), FORMAT_MONTO) '
                End If
                If vIndicadorConsulHastaDetalleReal = "DET" Then
                    .TextMatrix(.Rows - 1, 18) = Format(NulosN(RsCons("preuni_sol")), FORMAT_MONTO) '
                End If
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
                    .TextMatrix(.Rows - 1, 19) = Format(vSubTotal_Sol, FORMAT_MONTO) '
                Else
                    If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                        If NulosN(RsCons("idmon")) = 1 Then
                            .TextMatrix(.Rows - 1, 19) = Format(NulosN(RsCons("subtotal_sol")), FORMAT_MONTO) '
                        ElseIf NulosN(RsCons("idmon")) = 2 Then
                            .TextMatrix(.Rows - 1, 19) = "0.00" '
                        End If
                    Else
                        .TextMatrix(.Rows - 1, 19) = Format(NulosN(RsCons("subtotal_sol")), FORMAT_MONTO) '
                    End If
                End If

                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                    If NulosN(RsCons("idmon")) = 1 Then
                        .TextMatrix(.Rows - 1, 20) = Format(NulosN(RsCons("saldo_sol")), FORMAT_MONTO) '
                    ElseIf NulosN(RsCons("idmon")) = 2 Then
                        .TextMatrix(.Rows - 1, 20) = "0.00" '
                    End If
                    .TextMatrix(.Rows - 1, 21) = Format(RsCons("subtotal_sol"), FORMAT_MONTO) '
                    .TextMatrix(.Rows - 1, 22) = Format(NulosN(RsCons("subtotal_sol")) - NulosN(RsCons("saldo_sol")), FORMAT_MONTO) '
                    .TextMatrix(.Rows - 1, 23) = Format(NulosN(RsCons("saldo_sol")), FORMAT_MONTO) '
                End If
                ' SUMAR COLUMNAS DE GRUPO
                If vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
                    vFlag = 1
                    vSumCol_Cant = vSumCol_Cant + NulosN(RsCons("canpro"))
                    vSumCol_Dol = vSumCol_Dol + Val(NulosN(RsCons("subtotal_dol")))
                    vSumCol_Sol = vSumCol_Sol + Val(NulosN(RsCons("subtotal_sol")))
                ElseIf vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                    vFlag = 1
                    vSumCol_Dol = vSumCol_Dol + Val(NulosN(RsCons("subtotal_dol")))
                    If NulosN(RsCons("idmon")) = 1 Then
                        vSumCol_Sol = vSumCol_Sol + Val(NulosN(RsCons("subtotal_sol")))
                    End If
                    vSumColSal_Dol = vSumColSal_Dol + Val(NulosN(RsCons("saldo_dol")))
                    If NulosN(RsCons("idmon")) = 1 Then
                        vSumColSal_Sol = vSumColSal_Sol + Val(NulosN(RsCons("saldo_sol")))
                    End If
                    ' PARTE DEL ABONO
                    vSumCol_AbonTot_Sol = vSumCol_AbonTot_Sol + Val(Format(.TextMatrix(.Rows - 1, 21), "#####0.00")) '
                    vSumCol_AbonAbono_Sol = vSumCol_AbonAbono_Sol + Val(Format(.TextMatrix(.Rows - 1, 22), "#####0.00")) '
                    vSumCol_AboSaldo_Sol = vSumCol_AboSaldo_Sol + Val(Format(.TextMatrix(.Rows - 1, 23), "#####0.00")) '
                End If
            End With
            RsCons.MoveNext
        Loop
        
        If vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
            Fg1.AddItem ""
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = "Total: " '
            formatTextCeldaGrid Fg1.Rows - 1, 6, &H800000 '
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(vSumCol_Cant, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(vSumCol_Dol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumCol_Sol, FORMAT_MONTO) '
            ' SUMTOTAL GENERAL
            vSumColGen_Cant = vSumColGen_Cant + vSumCol_Cant
            vSumColGen_SubTot_Dol = vSumColGen_SubTot_Dol + vSumCol_Dol
            vSumColGen_SubTot_Sol = vSumColGen_SubTot_Sol + vSumCol_Sol
            ' FIN SUMTOTAL GENERAL
            formatTextCeldaGrid Fg1.Rows - 1, 14, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 16, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 19, &H800000 '
            Fg1.AddItem ""
            ' AQUI LLAMR LA CONSULT DEL CALCULO PRE PROM
            
            If vIndicaPreProm = 1 Then 'LLAMAR LA CONSULTA
                RsCons.MoveFirst
                fSqlCalPreProm Val(RsCons("idpro")), Val(RsCons("idtipoprod")), Val(RsCons("iditem"))
                ' ESTE PROCEDIMIENTO DEVUELVE EL RECORSET QUE CONTIENE EL PREC PROM
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = "P. Prom.:" '
                formatTextCeldaGrid Fg1.Rows - 1, 6, &H800000 '
                
                If NulosN(RsPreProm("pre_prom_sol")) > 0 Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(NulosN(RsPreProm("pre_prom_sol")), FORMAT_MONTO) '
                    formatTextCeldaGrid Fg1.Rows - 1, 18, &H800000 '
                End If
                
                If NulosN(RsPreProm("pre_prom_dol")) > 0 Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(NulosN(RsPreProm("pre_prom_dol")), FORMAT_MONTO) '
                    formatTextCeldaGrid Fg1.Rows - 1, 15, &H800000 '
                End If
                Fg1.AddItem ""
            End If
        ElseIf vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
            Fg1.AddItem ""
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = "Total: " '
            formatTextCeldaGrid Fg1.Rows - 1, 6, &H800000 '
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(vSumCol_Dol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumCol_Sol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(vSumColSal_Dol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(vSumColSal_Sol, FORMAT_MONTO) '
            formatTextCeldaGrid Fg1.Rows - 1, 16, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 17, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 19, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 20, &H800000 '
            ' PARTE DEL ABONO
            Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(vSumCol_AbonTot_Sol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 22) = Format(vSumCol_AbonAbono_Sol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 23) = Format(vSumCol_AboSaldo_Sol, FORMAT_MONTO) '
            formatTextCeldaGrid Fg1.Rows - 1, 21, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 22, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 23, &H800000 '
            ' SUMTOTAL GENERAL
            vSumColGen_SubTot_Dol = vSumColGen_SubTot_Dol + vSumCol_Dol
            vSumColGen_SubTot_Sol = vSumColGen_SubTot_Sol + vSumCol_Sol
            vSumTotGenSaldoDol = vSumTotGenSaldoDol + vSumColSal_Dol
            vSumTotGenSaldoSol = vSumTotGenSaldoSol + vSumColSal_Sol
            ' SUMTOTAL GENERAL PARTE DEL ABONO
            vSumColGen_AbonTot_Sol = vSumColGen_AbonTot_Sol + vSumCol_AbonTot_Sol
            vSumColGen_AbonAbon = vSumColGen_AbonAbon + vSumCol_AbonAbono_Sol
            vSumColGen_AbonSaldo_Sol = vSumColGen_AbonSaldo_Sol + vSumCol_AboSaldo_Sol
            Fg1.AddItem ""
        End If
    Else  'AGR - DET
'    1         2            3            4         5             6
'Tipo Doc., Cliente, Num. Documento, Fec. Doc., Fec. Venc., Cond. Pago,
'    7          8         9           10          11      12         13         14
'Dias Atras., Moneda, Tipo Cambio, Tipo Producto, Item, Uni. Med., Cantidad, Prec. Unit. $
'       15       16            17            18            19         20          21        22
'Imp. Total $, Saldo $, Prec. Unit S/., Imp. Total S/., Saldo S/., Total S/., Abono S/., Saldo S/.
        If vAgregaSubTotal = "SUBTOT" Then
            Fg1.AddItem ""
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = "Total: " '
            formatTextCeldaGrid Fg1.Rows - 1, 6, &H800000 '
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(vSumCol_Dol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumCol_Sol, FORMAT_MONTO) '
            formatTextCeldaGrid Fg1.Rows - 1, 16, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 19, &H800000 '
            Exit Sub
        ElseIf vAgregaSubTotal = "TOTGEN" Then
            Fg1.AddItem ""
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = "Tot. Gen.: " '
            formatTextCeldaGrid Fg1.Rows - 1, 6, &H800000 '
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = Format(vSumColGen_SubTot_Dol, FORMAT_MONTO) '
            Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumColGen_SubTot_Sol, FORMAT_MONTO) '
            formatTextCeldaGrid Fg1.Rows - 1, 16, &H800000 '
            formatTextCeldaGrid Fg1.Rows - 1, 19, &H800000 '
            Exit Sub
        End If
        
        If vNomProv <> "" Then
            If Fg1.TextMatrix(2, 2) = "" Then '
                Fg1.TextMatrix(2, 1) = vNomProv
                formatTextCeldaGrid 2, 1, &H800000
                UnirCeldas 2, 1, 3, Trim(vNomProv) '
            Else
                Fg1.AddItem ""
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = vNomProv
                formatTextCeldaGrid Fg1.Rows - 1, 1, &H800000
                UnirCeldas Fg1.Rows - 1, 1, 3, Trim(vNomProv) '
            End If
        End If
        
        RsCons.MoveFirst
            If Trim(Fg1.TextMatrix(2, 2)) <> "" Then Fg1.AddItem "" '
            With Fg1
                .TextMatrix(.Rows - 1, 1) = NulosC(RsCons("registro")) '
                .TextMatrix(.Rows - 1, 2) = NulosC(RsCons("abrev")) '
                .TextMatrix(.Rows - 1, 3) = NulosC(RsCons("nombre")) '
                .TextMatrix(.Rows - 1, 4) = NulosC(RsCons("numerodoc")) '
                .TextMatrix(.Rows - 1, 5) = NulosC(RsCons("fchdoc")) '
                .TextMatrix(.Rows - 1, 6) = NulosC(RsCons("fchven")) '
                .TextMatrix(.Rows - 1, 7) = NulosC(RsCons("desccond")) '
                If NulosN(RsCons("impsal")) > 0 Then
                    .TextMatrix(.Rows - 1, 8) = Abs(DateDiff("d", Date, NulosC(RsCons("fchven")))) '
                End If
                .TextMatrix(.Rows - 1, 9) = NulosC(RsCons("simbolo")) '
    '    1         2            3            4         5             6
    'Tipo Doc., Cliente, Num. Documento, Fec. Doc., Fec. Venc., Cond. Pago,
    '    7          8         9           10          11      12         13         14
    'Dias Atras., Moneda, Tipo Cambio, Tipo Producto, Item, Uni. Med., Cantidad, Prec. Unit. $
    '       15       16            17            18            19         20          21        22
    'Imp. Total $, Saldo $, Prec. Unit S/., Imp. Total S/., Saldo S/., Total S/., Abono S/., Saldo S/.
                .TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(RsCons("impcom")), FORMAT_IMPUESTO) '
                If vIndicadorConsulHastaDetalleReal = "DET" Then
                    .TextMatrix(.Rows - 1, 11) = NulosC(RsCons("destipprod")) '
                    .TextMatrix(.Rows - 1, 12) = NulosC(RsCons("desc_item")) '
                    .TextMatrix(.Rows - 1, 13) = NulosC(RsCons("unid_abrev")) '
                    .TextMatrix(.Rows - 1, 14) = Format(NulosN(RsCons("canpro")), FORMAT_MONTO) '
                    .TextMatrix(.Rows - 1, 15) = Format(NulosN(RsCons("preuni_dol")), FORMAT_MONTO) '
                End If
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
                    .TextMatrix(Fg1.Rows - 1, 16) = Format(vSubTotal_Dol, FORMAT_MONTO) '
                Else
                    .TextMatrix(Fg1.Rows - 1, 16) = Format(NulosN(RsCons("subtotal_dol")), FORMAT_MONTO) '
                End If
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                    .TextMatrix(.Rows - 1, 17) = Format(NulosN(RsCons("saldo_dol")), FORMAT_MONTO) '
                End If
                If vIndicadorConsulHastaDetalleReal = "DET" Then
                    .TextMatrix(.Rows - 1, 18) = Format(NulosN(RsCons("preuni_sol")), FORMAT_MONTO) '
                End If
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
                    .TextMatrix(.Rows - 1, 19) = Format(vSubTotal_Sol, FORMAT_MONTO) '
                Else
                    .TextMatrix(.Rows - 1, 19) = Format(NulosN(RsCons("subtotal_sol")), FORMAT_MONTO) '
                End If
                If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
                    .TextMatrix(.Rows - 1, 20) = Format(NulosN(RsCons("saldo_sol")), FORMAT_MONTO) '
                    .TextMatrix(.Rows - 1, 21) = .TextMatrix(.Rows - 1, 19) '
                    'Val (Format(.TextMatrix(Fg1.Rows - 1, 15), "#####0.00"))
                    .TextMatrix(.Rows - 1, 22) = Format(Val(Format(.TextMatrix(Fg1.Rows - 1, 19), "#####0.00")) - Val(Format(.TextMatrix(.Rows - 1, 20), "#####0.00")), FORMAT_MONTO) '
                    .TextMatrix(.Rows - 1, 22) = .TextMatrix(.Rows - 1, 20) '
                End If
            End With
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : fStrSqlProv
'* Tipo             : FUNCCION
'* Descripcion      : GENERA UNA CADENA SQL QUE DEVOLVERA LA LISTA DE COMPRAS POR PROVEEDOR REALIZADAS
'*                    EN EL PERIODO ESPECIFICADO, DEVUELVE UNA CADENA QUE CONTIENE UNA SENTENCIA SQL
'* Paranetros       :
'* Devuelve         : STRING
'*****************************************************************************************************
Private Function fStrSqlProv() As String
    Dim vStrFiltro_ITEM As String, k As Integer
    With Fg3
        If Fg3.TextMatrix(1, 1) <> "" Then
            For k = 1 To .Rows - 1
                If .TextMatrix(k, 1) <> "" Then
                    If CStr(.TextMatrix(k, 1)) <> "" Then vStrFiltro_ITEM = vStrFiltro_ITEM + CStr(.TextMatrix(k, 1)) + ","
                End If
            Next k
        End If
    End With
    If vStrFiltro_ITEM <> "" Then vStrFiltro_ITEM = " AND mae_prov.id IN (" + Left(vStrFiltro_ITEM, Len(vStrFiltro_ITEM) - 1) + ") "
    
    vStrCons = "SELECT DISTINCT com_compras.idpro, mae_prov.nombre FROM mae_prov INNER JOIN com_compras ON mae_prov.id = com_compras.idpro " _
        & "WHERE com_compras.fchdoc BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "') " & vStrFiltro_ITEM _
        & " ORDER BY mae_prov.nombre"
    If OptVenc.Value = True Then
        vStrCons = Replace(vStrCons, "com_compras.fchdoc", "com_compras.fchven")
    End If
    fStrSqlProv = vStrCons
End Function

'*****************************************************************************************************
'* Nombre           : fStrSql
'* Tipo             : FUNCCION
'* Descripcion      : ********************************************************************************
'* Paranetros       : NOMBRE        |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pAgrupadoSiNo |  INTEGER    |  *************************************************
'* Devuelve         : STRING
'*****************************************************************************************************
Private Function fStrSql(pAgrupadoSiNo As Integer) As String
    Dim vStrFiltro_ITEM As String, k As Integer
    If pAgrupadoSiNo = 1 Then  ' AGRUPADO
        With Fg3
            If Fg3.TextMatrix(1, 1) <> "" Then
                For k = 1 To .Rows - 1
                    If .TextMatrix(k, 1) <> "" Then
                        If CStr(.TextMatrix(k, 1)) <> "" Then vStrFiltro_ITEM = vStrFiltro_ITEM + CStr(.TextMatrix(k, 1)) + ","
                    End If
                Next k
            End If
        End With
        If vStrFiltro_ITEM <> "" Then vStrFiltro_ITEM = " AND mae_prov.id IN (" + Left(vStrFiltro_ITEM, Len(vStrFiltro_ITEM) - 1) + ") "
        
        vStrCons = "SELECT com_compras.*,IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='',[com_compras].[numreg],Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4)) AS registro, mae_prov.numruc, mae_prov.nombre, com_compras!numser+'-'+com_compras!numdoc AS numerodoc, mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_prov.numruc, mae_moneda.descripcion AS descmon, mae_moneda.simbolo, con_tc.impcom, " _
            & " IIf(com_compras.idmon=1, com_compras.imptot,IIf(com_compras.idmon=2,IIf(con_tc.impcom Is Null,0, com_compras.imptot * con_tc.impcom),0)) AS subtotal_sol, " _
            & " IIf(com_compras.idmon = 2, com_compras.imptot, 0) As subtotal_dol, " _
            & " IIf([com_compras].[idmon]=1,[com_compras].[impsal],IIf([com_compras].[idmon]=2,IIf([con_tc].[impcom] Is Null,0,[com_compras].[impsal]*[con_tc].[impcom]),0)) AS saldo_sol, " _
            & " IIf([com_compras].[idmon] = 2, [com_compras].[impsal], 0) As saldo_dol " _
            & " FROM (mae_condpago RIGHT JOIN (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) ON mae_condpago.id = com_compras.idconpag) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id " _
            & " WHERE com_compras.fchdoc BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "') " & vStrFiltro_ITEM
            
        If OptVenc.Value = True Then
            vStrCons = Replace(vStrCons, "com_compras.fchdoc BETWEEN", "com_compras.fchven BETWEEN ")
        End If
        
        If NulosN(TxtIdTipProd.Text) <> 0 Then vStrCons = vStrCons & " AND alm_inventario.tippro= " & NulosN(TxtIdTipProd.Text)
        
        ' VERIFICAR MONEDA SELECCIONADA
        If OptSol.Value = True Then
            vStrCons = vStrCons & " AND com_compras.idmon = 1 "
        ElseIf OptDol.Value = True Then
            vStrCons = vStrCons & " AND com_compras.idmon = 2 "
        End If
        ' VERIFICAR SI LAS COMPRAS ESTAN PENDIENTES O CANCELADAS
        If OptPend.Value = True Then
            vStrCons = vStrCons & " AND com_compras.impsal > 0 "
        ElseIf OptPag.Value = True Then
            vStrCons = vStrCons & " AND com_compras.impsal <= 0 "
        End If

        vStrCons = vStrCons & " ORDER BY com_compras.fchdoc"
        
    Else ' DETALLADO
        With Fg2
            For k = 0 To .Rows - 1
                If Me.ChkMostrarItem.Value = 0 Then Exit For '--SALIR SI NO SELECCIONA MOSTRAR ITEM
                If k + 1 = .Rows Then Exit For
                If CStr(.TextMatrix(k + 1, 3)) <> "" Then vStrFiltro_ITEM = vStrFiltro_ITEM + CStr(.TextMatrix(k + 1, 3)) + ","
            Next k
        End With
        If vStrFiltro_ITEM <> "" Then vStrFiltro_ITEM = " AND alm_inventario.id IN (" + Left(vStrFiltro_ITEM, Len(vStrFiltro_ITEM) - 1) + ") "
        
        vStrCons = "SELECT com_compras.*,IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='',[com_compras].[numreg],Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4)) AS registro, mae_prov.nombre, com_compras!numser+'-'+com_compras!numdoc AS numerodoc, mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_prov.numruc, " _
            & " mae_moneda.descripcion AS descmon, mae_moneda.simbolo, con_tc.impcom, com_comprasdet.iditem, alm_inventario.descripcion AS desc_item, mae_unidades.abrev AS unid_abrev, com_comprasdet.canpro, IIF(com_compras.idmon = 1, com_comprasdet.preuni, IIF(com_compras.idmon = 2, IIF(con_tc.impcom is null, 0,  com_comprasdet.preuni * con_tc.impcom), 0)) AS preuni_sol, IIF(com_compras.idmon = 1, com_comprasdet.imptot, IIF(com_compras.idmon = 2, IIF(con_tc.impcom is null, 0,  com_comprasdet.imptot * con_tc.impcom), 0)) AS subtotal_sol, IIF(com_compras.idmon = 2, com_comprasdet.preuni, 0) AS preuni_dol, IIF(com_compras.idmon = 2, com_comprasdet.imptot, 0) AS subtotal_dol,  mae_tipoproducto.id AS idtipoprod, mae_tipoproducto.descripcion AS destipprod " _
            & " FROM ((mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) RIGHT JOIN (mae_condpago RIGHT JOIN (mae_prov RIGHT JOIN (mae_unidades RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_moneda.id = com_compras.idmon) ON mae_unidades.id = com_comprasdet.idunimed) ON mae_prov.id = com_compras.idpro) ON mae_condpago.id = com_compras.idconpag) ON alm_inventario.id = com_comprasdet.iditem) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id " _
            & " WHERE com_compras.fchdoc BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "') " & vStrFiltro_ITEM
        If OptVenc.Value = True Then
            vStrCons = Replace(vStrCons, "com_compras.fchdoc BETWEEN", "com_compras.fchven BETWEEN ")
        End If

        ' del tipo de producto
        If NulosN(TxtIdTipProd.Text) <> 0 Then vStrCons = vStrCons & " AND alm_inventario.tippro= " & NulosN(TxtIdTipProd.Text)

        ' VERIFICAR MONEDA SELECCIONADA
        If OptSol.Value = True Then
            vStrCons = vStrCons & " AND com_compras.idmon = 1 "
        ElseIf OptDol.Value = True Then
            vStrCons = vStrCons & " AND com_compras.idmon = 2 "
        End If
        
        ' VERIFICAR SI LAS COMPRAS ESTAN PENDIENTES O CANCELADAS
        If OptPend.Value = True Then
            vStrCons = vStrCons & " AND com_compras.impsal > 0 "
        ElseIf OptPag.Value = True Then
            vStrCons = vStrCons & " AND com_compras.impsal <= 0 "
        End If

        vStrCons = vStrCons & " ORDER BY com_compras.fchdoc"
    End If
    fStrSql = vStrCons
End Function

'*****************************************************************************************************
'* Nombre           : PosicionarProgBar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CENTRA EN EL FORMULARIO EL FRAME CONTROL FraProgreso PARA MOSTRAR EL CONTROL
'                     PgBar
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pCantMax  |  LONG             |  ESPECIFICA LA CANTIDAD MAXIMA DEL CONTADOR PARA
'*                              |                   |  EL CONTROL ProgressBar PgBar
'* Devuelve         :
'*****************************************************************************************************
Sub PosicionarProgBar(pCantMax As Long)
    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    FraProgreso.Visible = True
    PgBar.Max = pCantMax
End Sub

'*****************************************************************************************************
'* Nombre           : FormatGridSubTotales
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DA FORMATO PARA MOSTRAR SUBTOTALES EN EL CONTROL FlexGrid Fg1
'* Paranetros       : NOMBRE    |  TIPO         |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pFila     |  LONG         |  ESPECIFICA EL ID DE LA FILA QUE SE DARA FORMATO
'*                    pCol1     |  INTEGER      |  ESPECIFICA EL ID DE LA COLUMAN QUE SE DARA FORMATO
'*                    pCol2     |  INTEGER      |  ESPECIFICA EL ID DE LA COLUMAN QUE SE DARA FORMATO
'*                    pCol3     |  INTEGER      |  ESPECIFICA EL ID DE LA COLUMAN QUE SE DARA FORMATO
'*                    pCol4     |  INTEGER      |  ESPECIFICA EL ID DE LA COLUMAN QUE SE DARA FORMATO
'*                    pCol5     |  INTEGER      |  ESPECIFICA EL ID DE LA COLUMAN QUE SE DARA FORMATO
'* Devuelve         :
'*****************************************************************************************************
Sub FormatGridSubTotales(pFila As Long, pCol1 As Integer, pCol2 As Integer, pCol3 As Integer, pCol4 As Integer, pCol5 As Integer)
    Fg1.Row = pFila: Fg1.Col = pCol1
    Fg1.CellFontBold = True
    
    Fg1.Col = pCol2
    Fg1.CellFontBold = True
    
    Fg1.Col = pCol3
    Fg1.CellFontBold = True
    
    Fg1.Col = pCol4
    Fg1.CellFontBold = True
    
    Fg1.Col = pCol5
    Fg1.CellFontBold = True
End Sub

'*****************************************************************************************************
'* Nombre           : SumarTotalGeneral
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ********************************************************************************
'* Paranetros       : NOMBRE      |  TIPO         |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pConDetalle |  INTEGER      |  *************************************************
'* Devuelve         :
'*****************************************************************************************************
Sub SumarTotalGeneral(pConDetalle As Integer)
    vSumTotGenSol = 0: vSumTotGenDol = 0: vSumTotGenSaldoSol = 0: vSumTotGenSaldoDol = 0
    If pConDetalle = 1 Then 'SIN DETALLE
        If Fg1.Rows - 1 >= 2 Then
'    1            2             3          4           5          6            7               8        9         10
'Fec. Doc., Num. Document., Tipo Doc., Cond. Pago, Proveedor, Fec. Venc., Dias de Atraso, Fec. Pago, Moneda, Tipo Compra,
'     11            12             13        14      15        16           17             18              19            20          21         22
'Tipo Cambio, Imp. Total S/., Imp. Total $, Item, Cantidad, Uni. Med., Prec. Unit S/., Imp. Total S/., Prec. Unit $, Imp. Total $, Saldo S/., Saldo $
            For X = 1 To Fg1.Rows - 1
                With Fg1
                    If Trim(.TextMatrix(X, 1)) = "" And Trim(.TextMatrix(X, 12)) <> "" Then
                        vSumTotGenSol = vSumTotGenSol + NulosN(.TextMatrix(X, 12))
                        vSumTotGenDol = vSumTotGenDol + NulosN(.TextMatrix(X, 13))
                        vSumTotGenSaldoSol = vSumTotGenSaldoSol + NulosN(.TextMatrix(X, 21))
                        vSumTotGenSaldoDol = vSumTotGenSaldoDol + NulosN(.TextMatrix(X, 22))
                    End If
                End With
            Next
            Fg1.AddItem ""
            If OptMonTodos.Value = True Or OptDol.Value = True Then
                Fg1.TextMatrix(Fg1.Rows - 1, 11) = "Tot. Gen.:"
                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 11
                Fg1.CellFontBold = True
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = "Tot. Gen."
                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 10
                Fg1.CellFontBold = True
            End If
            
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(vSumTotGenSol, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(vSumTotGenDol, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(vSumTotGenSaldoSol, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 22) = Format(vSumTotGenSaldoDol, FORMAT_MONTO)
            Fg1.Row = Fg1.Rows - 1: Fg1.Col = 12
            Fg1.CellFontBold = True
            Fg1.Row = Fg1.Rows - 1: Fg1.Col = 13
            Fg1.CellFontBold = True
            
            Fg1.Row = Fg1.Rows - 1: Fg1.Col = 21
            Fg1.CellFontBold = True
            Fg1.Row = Fg1.Rows - 1: Fg1.Col = 22
            Fg1.CellFontBold = True
        End If
    ElseIf pConDetalle = 2 Then
        If Fg1.Rows - 1 >= 2 Then
            For X = 1 To Fg1.Rows - 1
                With Fg1
                    If Trim(.TextMatrix(X, 1)) = "" And Trim(.TextMatrix(X, 18)) <> "" Then
                        vSumTotGenSol = vSumTotGenSol + Val(Format(.TextMatrix(X, 18), FORMAT_MONTO))
                        vSumTotGenDol = vSumTotGenDol + Val(Format(.TextMatrix(X, 20), FORMAT_MONTO))
                    End If
                End With
            Next
'    1            2             3          4           5          6            7               8        9         10
'Fec. Doc., Num. Document., Tipo Doc., Cond. Pago, Proveedor, Fec. Venc., Dias de Atraso, Fec. Pago, Moneda, Tipo Compra,
'     11            12             13        14      15        16           17             18              19            20          21         22
'Tipo Cambio, Imp. Total S/., Imp. Total $, Item, Cantidad, Uni. Med., Prec. Unit S/., Imp. Total S/., Prec. Unit $, Imp. Total $, Saldo S/., Saldo $
            Fg1.AddItem ""
            If OptMonTodos.Value = True Or OptDol.Value = True Then
                Fg1.TextMatrix(Fg1.Rows - 1, 11) = "Tot. Gen.:"
                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 11
                Fg1.CellFontBold = True
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = "Tot. Gen."
                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 10
                Fg1.CellFontBold = True
            End If
            Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(vSumTotGenSol, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(vSumTotGenDol, FORMAT_MONTO)
            Fg1.Row = Fg1.Rows - 1: Fg1.Col = 18
            Fg1.CellFontBold = True
            Fg1.Row = Fg1.Rows - 1: Fg1.Col = 20
            Fg1.CellFontBold = True
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : LlenarGrillaResumen
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ********************************************************************************
'* Paranetros       : NOMBRE              |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pImpTotSol          |  DOUBLE     |  ESPECIFICA EL TOTAL EN SOLES
'*                    pImpTotDol          |  DOUBLE     |  ESPECIFICA EL TOTAL EN DOLARES
'*                    pImpTotSaldoSol     |  DOUBLE     |  ESPECIFICA EL SALDO EN SOLES
'*                    pImpTotSaldoDol     |  DOUBLE     |  ESPECIFICA EL SALDO EN DOLARES
'*                    pTotGenSiNo         |  STRING     |  ESPECIFICA SI SE MOSTRARA EL TOTAL GENERAL
'*                                        |             |  SI = MOSTRAR TOTAL GENERAL
'*                                        |             |  NO = NO MOSTRAR TOTAL GENERAL
'*                    pResumDetalladoSiNo |  STRING     |  ESPECIFICA SI SE TOTALIZARA EL RESUMEN
'*                                        |             |  SI = TOTALIZAR RESUMEN
'*                                        |             |  NO = NO TOTALIZAR RESUMEN
'* Devuelve         :
'*****************************************************************************************************
Private Sub LlenarGrillaResumen(pImpTotSol As Double, pImpTotDol As Double, pImpTotSaldoSol As Double, pImpTotSaldoDol As Double, pTotGenSiNo As String, pResumDetalladoSiNo As String)
    If pTotGenSiNo = "NO" Then
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
'    1            2             3          4           5          6            7               8        9         10
'Fec. Doc., Num. Document., Tipo Doc., Cond. Pago, Proveedor, Fec. Venc., Dias de Atraso, Fec. Pago, Moneda, Tipo Compra,
'     11            12             13        14      15        16           17             18              19            20          21         22
'Tipo Cambio, Imp. Total S/., Imp. Total $, Item, Cantidad, Uni. Med., Prec. Unit S/., Imp. Total S/., Prec. Unit $, Imp. Total $, Saldo S/., Saldo $
            If Fg1.TextMatrix(1, 5) <> "" Then Fg1.AddItem ""
            With Fg1
                .TextMatrix(.Rows - 1, 5) = NulosC(rsTemp("nombre")) 'PROVEEDOR
                .TextMatrix(.Rows - 1, 9) = NulosC(rsTemp("simbolo"))
                .TextMatrix(.Rows - 1, 10) = NulosC(rsTemp("desctipcom"))
                If pResumDetalladoSiNo = "SI" Then
                    .TextMatrix(.Rows - 1, 14) = NulosC(rsTemp("desc_item"))
                End If
                If NulosN(rsTemp("idmon")) > 0 Then  ' PARA EL SALDO
                    If rsTemp("idmon") = 1 Then      ' SOLES
                        .TextMatrix(.Rows - 1, 18) = Format(pImpTotSol, FORMAT_MONTO)
                        .TextMatrix(.Rows - 1, 20) = 0
                        .TextMatrix(.Rows - 1, 21) = Format(pImpTotSaldoSol, FORMAT_MONTO)
                        .TextMatrix(.Rows - 1, 22) = Format(0, "0.00")
                    ElseIf rsTemp("idmon") = 2 Then  ' DOLARES
                        .TextMatrix(.Rows - 1, 18) = Format(pImpTotSol, FORMAT_MONTO)
                        .TextMatrix(.Rows - 1, 20) = Format(pImpTotDol, FORMAT_MONTO)
                        .TextMatrix(.Rows - 1, 21) = Format(pImpTotSaldoSol, FORMAT_MONTO)
                        .TextMatrix(.Rows - 1, 22) = Format(pImpTotSaldoDol, FORMAT_MONTO)
                    End If
                End If
             End With
        End If
    Else
        ' SOLO PARA EL TOTAL GENERAL
        If Fg1.TextMatrix(1, 5) <> "" Then Fg1.AddItem ""
        If pResumDetalladoSiNo = "SI" Then
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = "Tot. Gen.: "
            ' PARA EL CASO DE TOTAL GENERAL(DEL RESUMEN)
            FormatGridSubTotales Fg1.Rows - 1, 14, 18, 20, 21, 22
        Else
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = "Tot. Gen.: "
            ' PARA EL CASO DE TOTAL GENERAL(DEL RESUMEN)
            FormatGridSubTotales Fg1.Rows - 1, 10, 18, 20, 21, 22
        End If
        Fg1.CellFontBold = True
        With Fg1
            .TextMatrix(.Rows - 1, 18) = Format(vSumTotGenSol, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, 20) = Format(vSumTotGenDol, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, 21) = Format(vSumTotGenSaldoSol, FORMAT_MONTO)
            .TextMatrix(.Rows - 1, 22) = Format(vSumTotGenSaldoDol, FORMAT_MONTO)
        End With
    End If
End Sub


'*****************************************************************************************************
'* Nombre           : LlenarGrilla
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ********************************************************************************
'* Paranetros       : NOMBRE                  |  TIPO     |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pConDetalleO_SinDetalle |  INTEGER  |  *****************************************
'* Devuelve         :
'*****************************************************************************************************
Private Sub LlenarGrilla(pConDetalleO_SinDetalle As Integer)
    Dim vSumTotImpTotSol As Double, vSumTotImpTotDol As Double
    Dim vSumPreUniSol As Double, vSumPreUnitDol As Double
    Dim vSumSaldoSol As Double, vSumSaldoDol As Double
    vSumTotImpTotSol = 0: vSumTotImpTotDol = 0: vSumPreUniSol = 0: vSumPreUnitDol = 0
    vSumSaldoSol = 0: vSumSaldoDol = 0
    If pConDetalleO_SinDetalle = 1 Then 'SIN DETALLE
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                DoEvents
                If Fg1.TextMatrix(1, 1) <> "" Then Fg1.AddItem ""
                With Fg1
                    .TextMatrix(.Rows - 1, 1) = Format(rsTemp("fchdoc"), FORMAT_DATE)  ' ERA INDEX 1
                    .TextMatrix(.Rows - 1, 2) = NulosC(rsTemp("numerodoc"))            ' ERA INDEX 2
                    .TextMatrix(.Rows - 1, 3) = NulosC(rsTemp("nomdoc"))               ' ERA INDEX 3
                    .TextMatrix(.Rows - 1, 4) = NulosC(rsTemp("desccond"))             ' ERA INDEX 4
                    .TextMatrix(.Rows - 1, 5) = NulosC(rsTemp("nombre"))               ' ERA INDEX 5
                    .TextMatrix(.Rows - 1, 6) = Format(rsTemp("fchven"), FORMAT_DATE)  ' FECHA DE VENCIM 'ERA INDEX 6
                    If Val(rsTemp("impsal")) > 0 Then
                    .TextMatrix(.Rows - 1, 7) = Abs(DateDiff("d", Date, rsTemp("fchven"))) ' ERA INDEX 7
                    End If
                    .TextMatrix(.Rows - 1, 8) = Format(rsTemp("fchpag"), FORMAT_DATE)  ' ERA INDEX 8
                    .TextMatrix(.Rows - 1, 9) = NulosC(rsTemp("simbolo"))              ' ERA INDEX 9
                    .TextMatrix(.Rows - 1, 10) = NulosC(rsTemp("desctipcom"))          ' ERA INDEX 10
                    .TextMatrix(.Rows - 1, 11) = NulosN(rsTemp("impcom")) 'TIPO CAMBIO ' ERA INDEX 11
                    .TextMatrix(.Rows - 1, 12) = Format(NulosN(rsTemp("imptotal_sol")), FORMAT_MONTO) ' ERA INDEX 12
                    .TextMatrix(.Rows - 1, 13) = Format(NulosN(rsTemp("imptot_dol")), FORMAT_MONTO)   ' ERA INDEX 13
                    If NulosN(rsTemp("idmon")) > 0 Then  ' PARA EL SALDO
                        If rsTemp("idmon") = 1 Then      ' SOLES
                            .TextMatrix(.Rows - 1, 21) = Format(NulosN(rsTemp("impsal")), FORMAT_MONTO) ' ERA INDEX 21
                            .TextMatrix(.Rows - 1, 22) = Format(0, "0.00") ' ERA INDEX 22
                        ElseIf rsTemp("idmon") = 2 Then  ' DOLARES
                            .TextMatrix(.Rows - 1, 21) = Format(NulosN(rsTemp("impsal")) * NulosN(rsTemp("impcom")), FORMAT_MONTO) 'ERA INDEX 21
                            .TextMatrix(.Rows - 1, 22) = Format(NulosN(rsTemp("impsal")), FORMAT_MONTO) 'ERA INDEX 22
                        End If
                    End If
                    vSumTotImpTotSol = vSumTotImpTotSol + NulosN(Fg1.TextMatrix(.Rows - 1, 12)) ' ERA INDEX 12
                    vSumTotImpTotDol = vSumTotImpTotDol + NulosN(Fg1.TextMatrix(.Rows - 1, 13)) ' ERA INDEX 13
                    vSumSaldoSol = vSumSaldoSol + NulosN(.TextMatrix(.Rows - 1, 21)) ' ERA INDEX 21
                    vSumSaldoDol = vSumSaldoDol + NulosN(.TextMatrix(.Rows - 1, 22)) ' ERA INDEX 22
                End With
                rsTemp.MoveNext
            Loop
            Fg1.AddItem ""
            If OptSol.Value = True Then
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = "Total: "  'ERA INDEX 10
                FormatGridSubTotales Fg1.Rows - 1, 10, 12, 13, 21, 22
            ElseIf OptMonTodos.Value = True Or OptDol.Value = True Then
                Fg1.TextMatrix(Fg1.Rows - 1, 11) = "Total: "
                FormatGridSubTotales Fg1.Rows - 1, 11, 12, 13, 21, 22
            End If
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(vSumTotImpTotSol, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(vSumTotImpTotDol, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(vSumSaldoSol, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 22) = Format(vSumSaldoDol, FORMAT_MONTO)
        End If
    ElseIf pConDetalleO_SinDetalle = 2 Then 'CON DETALLE
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                DoEvents
                If Fg1.TextMatrix(1, 1) <> "" Then Fg1.AddItem ""
'    1            2             3          4           5          6        7         8
'Fec. Doc., Num. Document., Tipo Doc., Cond. Pago, Proveedor, Fec. Pago, Moneda, Tipo Compra,
'      9              10        11      12        13           14              15              16            17
'Imp. Total S/., Imp. Total $, Item, Cantidad, Uni. Med., Prec. Unit S/., Imp. Total S/., Prec. Unit $, Imp. Total $
                With Fg1
                    .TextMatrix(.Rows - 1, 1) = Format(rsTemp("fchdoc"), FORMAT_DATE) ' FEC DOC
                    .TextMatrix(.Rows - 1, 2) = NulosC(rsTemp("numerodoc"))           ' NUM DOC
                    .TextMatrix(.Rows - 1, 3) = NulosC(rsTemp("nomdoc"))              ' TIPO DOC
                    .TextMatrix(.Rows - 1, 4) = NulosC(rsTemp("desccond"))            ' COND PAGO
                    .TextMatrix(.Rows - 1, 5) = NulosC(rsTemp("nombre"))              ' PROVEEDOR
                    .TextMatrix(.Rows - 1, 6) = Format(rsTemp("fchven"), FORMAT_DATE) ' FECHA DE VENCIM
                    .TextMatrix(.Rows - 1, 7) = Abs(DateDiff("d", Date, rsTemp("fchven"))) ' DIAS DE ATRASO
                    .TextMatrix(.Rows - 1, 8) = Format(rsTemp("fchpag"), FORMAT_DATE) ' FEC PAGO
                    .TextMatrix(.Rows - 1, 9) = NulosC(rsTemp("simbolo"))             ' SIMBOL MONEDA
                    .TextMatrix(.Rows - 1, 10) = NulosC(rsTemp("desctipcom"))         ' TIPO COMPRA
                    .TextMatrix(.Rows - 1, 11) = NulosN(rsTemp("impcom"))             ' TIPO CAMBIO
                    .TextMatrix(.Rows - 1, 14) = NulosC(rsTemp("desc_item"))          ' ITEM
                    .TextMatrix(.Rows - 1, 15) = NulosN(rsTemp("canpro"))             ' CANTIDAD
                    .TextMatrix(.Rows - 1, 16) = NulosC(rsTemp("unid_abrev"))         ' UNID MED
                    .TextMatrix(.Rows - 1, 17) = Format(NulosN(rsTemp("preunit_sol")), FORMAT_MONTO) ' PREC. UNIT SOL
                    .TextMatrix(.Rows - 1, 18) = Format(NulosN(rsTemp("imptot_sol")), FORMAT_MONTO)  ' IMP TOT SOL
                    .TextMatrix(.Rows - 1, 19) = Format(NulosN(rsTemp("preunit_dol")), FORMAT_MONTO) ' PREC. UNIT DOL
                    .TextMatrix(.Rows - 1, 20) = Format(NulosN(rsTemp("imptot_dol")), FORMAT_MONTO)  ' IMP TOT DOL
                    
                    vSumPreUniSol = vSumPreUniSol + NulosN(.TextMatrix(.Rows - 1, 17))
                    vSumTotImpTotSol = vSumTotImpTotSol + NulosN(.TextMatrix(.Rows - 1, 18))
                    vSumPreUnitDol = vSumPreUnitDol + NulosN(.TextMatrix(.Rows - 1, 19))
                    vSumTotImpTotDol = vSumTotImpTotDol + NulosN(.TextMatrix(.Rows - 1, 20))
                End With
                rsTemp.MoveNext
            Loop
            Fg1.AddItem ""
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = "Total: "
            Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(vSumPreUniSol, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(vSumTotImpTotSol, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(vSumPreUnitDol, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(vSumTotImpTotDol, FORMAT_MONTO)
            FormatGridSubTotales Fg1.Rows - 1, 16, 17, 18, 19, 20
        End If
    End If
    Set rsTemp = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : proc5
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ********************************************************************************
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub proc5()
    vSumTotGenSol = 0: vSumTotGenDol = 0: vSumTotGenSaldoSol = 0: vSumTotGenSaldoDol = 0
    Dim Rs1 As New ADODB.Recordset
    Dim vSumTotImpTotSol As Double, vSumTotImpTotDol As Double
    Dim vSumPreUniSol As Double, vSumPreUnitDol As Double, vSumTotSaldoSol As Double, vSumTotSaldoDol As Double
    Dim i As Integer, X As Integer
    If Trim(TxtIdTipProd.Text) = "" And Trim(lblTipProducto.Caption) = "" And Trim(Fg3.TextMatrix(1, 1)) <> "" Then
        LimpiarGrid
        If OptDetalle.Value = True Then
            Call ConfiguraGridCons(1)
        Else
            ConfiguraGridResumen "NO"
        End If
        ' NO SE SELECCIONO EL TIPO DE PRODUCTO, SI HAY PROVEEDORES SELECIONADOS
        ' ESTA CONSULTA ES SIN DETALLE
        Set RsCons = New ADODB.Recordset
        RsCons.CursorLocation = adUseClient
        vStrCons = "SELECT com_compras.*, mae_prov.nombre, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numerodoc, mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_prov.numruc, mae_moneda.descripcion AS descmon, mae_moneda.simbolo, mae_tipoproducto.descripcion AS desctipcom, con_tc.impcom, "
        vStrCons = vStrCons & "IIF(com_compras.idmon = 1, com_compras.imptot, IIF(com_compras.idmon = 2, com_compras.imptot * con_tc.impcom, 0)) as imptotal_sol, iif(com_compras.idmon = 2, com_compras.imptot, 0) as imptot_dol "
        vStrCons = vStrCons & "FROM mae_condpago RIGHT JOIN (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((com_compras LEFT JOIN mae_tipoproducto ON com_compras.idtipo = mae_tipoproducto.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) ON mae_condpago.id = com_compras.idconpag "
        vStrCons = vStrCons & "WHERE com_compras.fchdoc BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "') "
        ' VERIFICAR MONEDA SELECCIONADA
        If OptSol.Value = True Then
            vStrCons = vStrCons & "AND com_compras.idmon = 1 "
        ElseIf OptDol.Value = True Then
            vStrCons = vStrCons & "AND com_compras.idmon = 2 "
        End If
        ' VERIFICAR SI LAS COMPRAS ESTAN PENDIENTES O CANCELADAS
        If ChkPend.Value = 1 And ChkPagada.Value = 0 Then 'COMPRA CON SALDO
            vStrCons = vStrCons & "AND com_compras.impsal > 0 "
        ElseIf ChkPend.Value = 0 And ChkPagada.Value = 1 Then
            vStrCons = vStrCons & "AND com_compras.impsal <= 0 "
        End If
        vStrCons = vStrCons & "ORDER BY com_compras.fchdoc, com_compras.numser, com_compras.numdoc, mae_moneda.simbolo"
        RST_Busq RsCons, vStrCons, xCon
        
        Set Rs1 = New ADODB.Recordset
        Rs1.CursorLocation = adUseClient
        vStrCons = "SELECT DISTINCT com_compras.idtipo, mae_tipoproducto.descripcion FROM mae_tipoproducto INNER JOIN com_compras ON mae_tipoproducto.id = com_compras.idtipo " _
            & "WHERE com_compras.fchdoc BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "') " _
            & "ORDER BY mae_tipoproducto.descripcion"
        RST_Busq Rs1, vStrCons, xCon
        
        For i = 1 To Fg3.Rows - 1
            If Trim(Fg3.TextMatrix(i, 1)) <> "" Then
                If Rs1.RecordCount > 0 Then ' RECORREMOS LOS TIPOS DE ITEM
                    Rs1.MoveFirst
                    Do While Not Rs1.EOF
                        DoEvents
                        '''''''
                        RsCons.Filter = "idpro = " & Val(Fg3.TextMatrix(i, 1)) & " AND idtipo = " & Val(Rs1("idtipo")) & ""
                        If OptDetalle.Value = True Then
                            If RsCons.RecordCount > 0 Then
                                Set rsTemp = Nothing
                                Set rsTemp = RsCons
                                LlenarGrilla 1
                            End If
                        Else ' PARA EL RESUMEN
                            If RsCons.RecordCount > 0 Then
                                vSumTotImpTotSol = 0: vSumTotImpTotDol = 0: vSumTotSaldoSol = 0: vSumTotSaldoDol = 0
                                Set rsTemp = Nothing
                                Set rsTemp = RsCons
                                RsCons.MoveFirst
                                Do While Not RsCons.EOF
                                    vSumTotImpTotSol = vSumTotImpTotSol + NulosN(RsCons("imptotal_sol"))
                                    vSumTotImpTotDol = vSumTotImpTotDol + NulosN(RsCons("imptot_dol"))
                                    If NulosN(RsCons("idmon")) > 0 Then 'PARA EL SALDO
                                        If RsCons("idmon") = 1 Then
                                            vSumTotSaldoSol = vSumTotSaldoSol + NulosN(RsCons("impsal"))
                                            vSumTotSaldoDol = 0
                                        ElseIf RsCons("idmon") = 2 Then
                                            vSumTotSaldoSol = vSumTotSaldoSol + (NulosN(RsCons("impsal")) * NulosN(RsCons("impcom")))
                                            vSumTotSaldoDol = vSumTotSaldoDol + NulosN(RsCons("impsal"))
                                        End If
                                    End If
                                    RsCons.MoveNext
                                Loop
                                ' AQUI LLAMAR CONSULTA RESUMIDA
                                LlenarGrillaResumen vSumTotImpTotSol, vSumTotImpTotDol, vSumTotSaldoSol, vSumTotSaldoDol, "NO", "NO"
                                vSumTotGenSol = vSumTotGenSol + vSumTotImpTotSol
                                vSumTotGenDol = vSumTotGenDol + vSumTotImpTotDol
                                vSumTotGenSaldoSol = vSumTotGenSaldoSol + vSumTotSaldoSol
                                vSumTotGenSaldoDol = vSumTotGenSaldoDol + vSumTotSaldoDol
                            End If
                        End If
                        RsCons.Filter = adFilterNone
                        Rs1.MoveNext
                    Loop
                    If OptResum.Value = True Then
                        If Fg1.TextMatrix(1, 5) <> "" Then
                            LlenarGrillaResumen vSumTotImpTotSol, vSumTotImpTotDol, vSumTotSaldoSol, vSumTotSaldoDol, "SI", "NO"
                        End If
                    End If
                End If
            End If
        Next
    End If
    If OptDetalle.Value = True Then SumarTotalGeneral 1
    Set RsCons = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : proc2
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ********************************************************************************
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub proc2()
    vSumTotGenSol = 0: vSumTotGenDol = 0: vSumTotGenSaldoSol = 0: vSumTotGenSaldoDol = 0
    Dim Rs1 As New ADODB.Recordset, Rs2 As ADODB.Recordset
    Dim vCantRegItem As Long, vCantRegProv As Long, vSumTotImpTotSol As Double, vSumTotImpTotDol As Double
    Dim vSumPreUniSol As Double, vSumPreUnitDol As Double, vSumTotSaldoSol As Double, vSumTotSaldoDol As Double
    Dim i As Integer, X As Integer
    vCantRegItem = Fg2.Rows - 1: vCantRegProv = Fg3.Rows - 1
    If Trim(TxtIdTipProd.Text) <> "" And Trim(lblTipProducto.Caption) <> "" And ChkMostrarItem.Value = 0 And Fg3.TextMatrix(1, 1) = "" Then 'CON TIPO DE PRODUCTO, SIN MARCAR MOSTRAR ITEM, MONED TODOS, PAGO PENDIENTE, SIN PROVEEDOR
        ' CON TIPO DE PRODUCTO PERO SIN MARCAR LA OPCION DE ITEM, SIN PROVEEDOR
        LimpiarGrid
        Call ConfiguraGridCons(1)
        ' CONSULTA AGRUPADO SOLO POR DOC COMPRA
        Set Rs1 = New ADODB.Recordset
        Rs1.CursorLocation = adUseClient
        vStrCons = "SELECT com_compras.*, mae_prov.nombre, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numerodoc, mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_prov.numruc, mae_moneda.descripcion AS descmon, mae_moneda.simbolo, mae_tipoproducto.descripcion AS desctipcom, con_tc.impcom, "
        vStrCons = vStrCons & "IIF(com_compras.idmon = 1, com_compras.imptot, IIF(com_compras.idmon = 2, com_compras.imptot * con_tc.impcom, 0)) as imptotal_sol, iif(com_compras.idmon = 2, com_compras.imptot, 0) as imptot_dol "
        vStrCons = vStrCons & "FROM mae_condpago RIGHT JOIN (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((com_compras LEFT JOIN mae_tipoproducto ON com_compras.idtipo = mae_tipoproducto.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) ON mae_condpago.id = com_compras.idconpag "
        vStrCons = vStrCons & "WHERE com_compras.fchdoc BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "') AND mae_tipoproducto.id = " & Val(TxtIdTipProd.Text) & " "
        
        ' VERIFICAR MONEDA SELECCIONADA
        If OptSol.Value = True Then
            vStrCons = vStrCons & "AND com_compras.idmon = 1 "
        ElseIf OptDol.Value = True Then
            vStrCons = vStrCons & "AND com_compras.idmon = 2 "
        End If
        ' VERIFICAR SI LAS COMPRAS ESTAN PENDIENTES O CANCELADAS
        If ChkPend.Value = 1 And ChkPagada.Value = 0 Then 'COMPRA CON SALDO
            vStrCons = vStrCons & "AND com_compras.impsal > 0 "
        ElseIf ChkPend.Value = 0 And ChkPagada.Value = 1 Then
            vStrCons = vStrCons & "AND com_compras.impsal <= 0 "
        End If
        ''''''''
        vStrCons = vStrCons & "ORDER BY com_compras.fchdoc, com_compras.numser, com_compras.numdoc, mae_moneda.simbolo"
        RST_Busq Rs1, vStrCons, xCon
                
        Set Rs2 = New ADODB.Recordset
        Rs2.CursorLocation = adUseClient
        vStrCons = "SELECT DISTINCT com_compras.idpro, mae_prov.nombre FROM mae_prov INNER JOIN com_compras ON mae_prov.id = com_compras.idpro " _
            & "WHERE com_compras.fchdoc BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "') " _
            & "ORDER BY mae_prov.nombre"
        RST_Busq Rs2, vStrCons, xCon
        If OptResum.Value = True Then
            ConfiguraGridResumen "NO"
        End If
        
        If Rs2.RecordCount > 0 Then ' RECORREMOS LOS PROVEEDORES
            PosicionarProgBar Rs2.RecordCount
            Rs2.MoveFirst
            Do While Not Rs2.EOF
                DoEvents
                Rs1.Filter = adFilterNone
                Rs1.Filter = "idpro = " & Val(Rs2("idpro")) & ""
                If OptDetalle.Value = True Then
                    If Rs1.RecordCount > 0 Then
                        Set rsTemp = Nothing
                        Set rsTemp = Rs1
                        LlenarGrilla 1
                    End If
                Else ' PARA EL RESUMEN
                    If Rs1.RecordCount > 0 Then
                        vSumTotImpTotSol = 0: vSumTotImpTotDol = 0: vSumTotSaldoSol = 0: vSumTotSaldoDol = 0
                        Set rsTemp = Nothing
                        Set rsTemp = Rs1
                        Rs1.MoveFirst
                        Do While Not Rs1.EOF
                            vSumTotImpTotSol = vSumTotImpTotSol + NulosN(Rs1("imptotal_sol"))
                            vSumTotImpTotDol = vSumTotImpTotDol + NulosN(Rs1("imptot_dol"))
                            If NulosN(Rs1("idmon")) > 0 Then ' PARA EL SALDO
                                If Rs1("idmon") = 1 Then
                                    vSumTotSaldoSol = vSumTotSaldoSol + NulosN(Rs1("impsal"))
                                    vSumTotSaldoDol = 0
                                ElseIf Rs1("idmon") = 2 Then
                                    vSumTotSaldoSol = vSumTotSaldoSol + (NulosN(Rs1("impsal")) * NulosN(Rs1("impcom")))
                                    vSumTotSaldoDol = vSumTotSaldoDol + NulosN(Rs1("impsal"))
                                End If
                            End If
                            Rs1.MoveNext
                        Loop
                        ' AQUI LLAMAR CONSULTA RESUMIDA
                        LlenarGrillaResumen vSumTotImpTotSol, vSumTotImpTotDol, vSumTotSaldoSol, vSumTotSaldoDol, "NO", "NO"
                        vSumTotGenSol = vSumTotGenSol + vSumTotImpTotSol
                        vSumTotGenDol = vSumTotGenDol + vSumTotImpTotDol
                        vSumTotGenSaldoSol = vSumTotGenSaldoSol + vSumTotSaldoSol
                        vSumTotGenSaldoDol = vSumTotGenSaldoDol + vSumTotSaldoDol
                    End If
                    ''''''''
                End If
                PgBar.Value = PgBar.Value + 1
                Rs2.MoveNext
            Loop ' RECORREMOS LOS PROVEEDORES
            FraProgreso.Visible = False
            If OptResum.Value = True Then
                If Fg1.TextMatrix(1, 5) <> "" Then
                    LlenarGrillaResumen vSumTotImpTotSol, vSumTotImpTotDol, vSumTotSaldoSol, vSumTotSaldoDol, "SI", "NO"
                End If
            End If
        Else
            MsgBox "No se encontraron registros...!", vbInformation, "Mensaje...!"
            TxtFec1.SetFocus
            Exit Sub
        End If
    End If
    If OptDetalle.Value = True Then SumarTotalGeneral 1
    Fg1.SetFocus
    Fg1.Row = Fg1.Rows - 1: Fg1.Col = 22
    SendKeys "{down}"
    Set Rs1 = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : ConfiguraGridCons
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pParam    |  INTEGER    |
'* Devuelve         :
'*****************************************************************************************************
Sub ConfiguraGridCons(pParam As Integer) 'SIN DETALLE = 1, CON DETALLE = 2
    If OptMonTodos.Value = True Or OptDol.Value = True Then
        Fg1.ColWidth(11) = 1050
    Else
        Fg1.ColWidth(11) = 0
    End If
    If ChkPend.Value = 1 And ChkPagada.Value = 0 Then
        Fg1.ColWidth(6) = 960
        Fg1.ColWidth(7) = 1230
    ElseIf ChkPend.Value = 0 And ChkPagada.Value = 1 Then
        Fg1.ColWidth(6) = 0
        Fg1.ColWidth(7) = 0
    End If
    If pParam = 1 Then ' SIN DETALLE
        ' VERIFICAR MONEDA SELECCIONADA
        Fg1.ColWidth(12) = 1365    ' IMPO. TOTAL S/.
        Fg1.ColWidth(13) = 1365    ' IMPO TOTAL $
        If OptSol.Value = True Then
            Fg1.ColWidth(12) = 1365
            Fg1.ColWidth(13) = 0
        ElseIf OptDol.Value = True Then
            Fg1.ColWidth(12) = 1365
            Fg1.ColWidth(13) = 1365
        End If
        Fg1.ColWidth(14) = 0
        Fg1.ColWidth(15) = 0
        Fg1.ColWidth(16) = 0
        Fg1.ColWidth(17) = 0
        Fg1.ColWidth(18) = 0
        Fg1.ColWidth(19) = 0
        Fg1.ColWidth(20) = 0
        Fg1.ColWidth(21) = 1230
        If OptDol.Value = True Then
            Fg1.ColWidth(22) = 1230
        Else
            Fg1.ColWidth(22) = 0
        End If
    ElseIf pParam = 2 Then       ' CON DETALLE
        Fg1.ColWidth(12) = 0     ' IMPO. TOTAL S/.
        Fg1.ColWidth(13) = 0     ' IMPO TOTAL $
        Fg1.ColWidth(14) = 1965  ' ITEM
        Fg1.ColWidth(15) = 1110  ' CANTID
        Fg1.ColWidth(16) = 960   ' UNID MED
        Fg1.ColWidth(17) = 1230  ' PRE UNIT S/.
        Fg1.ColWidth(18) = 1230  ' IMP. TOT S/.(SUB TOTAL)
        Fg1.ColWidth(19) = 1230  ' PREC UNIT $
        Fg1.ColWidth(20) = 1230  ' IMP TOTAL $(SUB TOTAL)
        Fg1.ColWidth(21) = 0
        Fg1.ColWidth(22) = 0
        If OptSol.Value = True Then
            Fg1.ColWidth(17) = 1230
            Fg1.ColWidth(18) = 1230
            Fg1.ColWidth(19) = 0
            Fg1.ColWidth(20) = 0
        ElseIf OptDol.Value = True Then
            Fg1.ColWidth(17) = 1230
            Fg1.ColWidth(18) = 1230
            Fg1.ColWidth(19) = 1230
            Fg1.ColWidth(20) = 1230
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : LimpiarGridItem
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LIMPIA LAS FILAS DEL CONTROL FlexGrid F2
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub LimpiarGridItem()
    Fg2.Clear
    Fg2.Rows = 2
    Fg2.FormatString = vFormatStrGridItem
    Fg2.ColWidth(1) = 0
    Fg2.ColWidth(3) = 0
    Fg2.ColComboList(2) = "|..."
    Fg2.Editable = flexEDKbdMouse
    Fg2.SelectionMode = flexSelectionFree
End Sub

'*****************************************************************************************************
'* Nombre           : LimpiarGridProv
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LIMPIA LAS FILAS DEL CONTROL FlexGrid F3
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub LimpiarGridProv()
    Fg3.Clear
    Fg3.Rows = 2
    Fg3.FormatString = vFormatGridProv
    Fg3.ColWidth(1) = 0
    Fg3.ColComboList(2) = "|..."
    Fg3.Editable = flexEDKbdMouse
    Fg3.SelectionMode = flexSelectionFree
End Sub

'*****************************************************************************************************
'* Nombre           : CargarProv
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LA LISTA DE PROVEEDORES DISPONIBLES
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargarProv()
    Dim xfrm As New eps_librerias.FormSeleccion
    Dim xCampos(2, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim xRs1 As New ADODB.Recordset
    
    xCampos(0, 0) = "Nombre":    xCampos(0, 1) = "nombre":   xCampos(0, 2) = "4000":   xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
    xCampos(1, 0) = "Id":        xCampos(1, 1) = "id":       xCampos(1, 2) = "1000":   xCampos(1, 3) = "C":     xCampos(1, 4) = "S"
        
    xfrm.SQLCad = "SELECT id, nombre, numruc FROM mae_prov ORDER BY nombre"
    xfrm.Titulo = "Buscando proveedores"
    
    Set xfrm.Coneccion = xCon
    Set xRs = xfrm.Seleccionar(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount = 0 Then
            Set xRs = Nothing
            Exit Sub
        End If
        
        Dim xCadWHERE As String
        Dim A As Integer
        Dim Rst As New ADODB.Recordset
        
        'CARGAMOS LOS DOCUMENTOS ADJUNTOS Y LO MOSTRAMOS EN LA LISTA DE "DOCUMENTOS ADJUNTOS"
        For A = 1 To xRs.RecordCount
        Next A
    End If
    Set xfrm = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : LimpiarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA LAS FILAS DEL CONTROL FlexGrid Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub LimpiarGrid()
    Fg1.Clear
    Fg1.Rows = 3
    Fg1.FormatString = vFormatString
End Sub

'*****************************************************************************************************
'* Nombre           : fConvertMoneda
'* Tipo             : FUNCCION
'* Descripcion      : COMVIERTE UN IMPORTE EN UNA MONEDA A OTRA
'* Paranetros       : NOMBRE    |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pIdMon    |  INTEGER     |  ESPECIFICA EL ID DE LA MONEDA
'*                    pTipoCam  |  DOUBLE      |  ESPECIFICA EL TIPO DE CAMBIO ACTUAL PARA LA CONVERSION
'*                    pvalor    |  DOUBLE      |  ESPECIFICA EL VALOR A CONVERTIR
'* Devuelve         : DOUBLE
'*****************************************************************************************************
Private Function fConvertMoneda(pIdMon As Integer, pTipoCam As Double, pvalor As Double) As Double
    Dim vValorConvert As Double
    Select Case pIdMon
        Case 1  'SOLES
            vValorConvert = pvalor / pTipoCam
        Case 2 'DOLARES
            vValorConvert = pvalor * pTipoCam
    End Select
    fConvertMoneda = Format(vValorConvert, "#####0.00")
End Function

Private Sub CmdBusProducto_Click()
    ' BUSCA EL TIPO DE PRODUCTO
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "800":          xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT id, descripcion FROM mae_tipoproducto"
    
    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdTipProd.Text = NulosN(xRs("id"))
        lblTipProducto.Caption = NulosC(xRs("descripcion"))
        ChkMostrarItem.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : pConsultar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pConsultar()
    If fverifFecha = False Then
        Exit Sub
    End If
    ' CONFIGURAR LA CONSULTA
    Set RsCons = Nothing
    Set RsProv = Nothing
    BAND_INTERRUMPIR = False
    vrangofec = "Desde: " & Trim(TxtFec1.Valor) & " Hasta: " & Trim(TxtFec2.Valor)
    If OptResum.Value = True Then
        If Trim(TxtIdTipProd.Text) <> "" Or ChkMostrarItem.Value = 1 Then
            ' MOSTRAR CONSULTA DETALLADA
            RST_Busq RsCons, fStrSql(2), xCon
            vIndicadorConsulHastaDetalleReal = "DET" ' CONSULTA HASTA EL DETALLE
            If ChkMostrarItem.Value = 0 Then
                vIndicadorConsul = "AGR"             ' AGRUPADO OSEA SIN MOSTRAR LOS ITEM
            Else
                vIndicadorConsul = "DET"             ' DETALLE ACA SI MOSTRAR LOS ITEM
            End If
        Else
            RST_Busq RsCons, fStrSql(1), xCon
            vIndicadorConsul = "AGR"                 ' AGRUPADO
            vIndicadorConsulHastaDetalleReal = "AGR" ' CONSULTA SOLO AGRUPADO SOLO PROVEEDOR
            'SIN LOS TIPO DE PRODUCTOS
        End If
        Set Rs4 = Nothing
        RST_Busq Rs4, fSqlTipProd_SinRepetir, xCon
        
        Set Rs5 = Nothing
        RST_Busq Rs5, fSqlTipProdDife_Resume, xCon
        
        RST_Busq Rs1, fStrSqlProv, xCon
        LimpiarGrid
        FormatGrid_Nuevo
        ProcResumen
    ElseIf OptDetalle.Value = True Then
        If Trim(TxtIdTipProd.Text) <> "" Or ChkMostrarItem.Value = 1 Then
            ' MOSTRAR CONSULTA DETALLADA
            RST_Busq RsCons, fStrSql(2), xCon
            vIndicadorConsulHastaDetalleReal = "DET" ' CONSULTA HASTA EL DETALLE
            If ChkMostrarItem.Value = 0 Then
                vIndicadorConsul = "AGR"             ' AGRUPADO OSEA SIN MOSTRAR LOS ITEM
            Else
                vIndicadorConsul = "DET"             ' DETALLE ACA SI MOSTRAR LOS ITEM
            End If
        Else
            RST_Busq RsCons, fStrSql(1), xCon
            vIndicadorConsul = "AGR"                 ' AGRUPADO
            vIndicadorConsulHastaDetalleReal = "AGR" ' CONSULTA SOLO AGRUPADO SOLO PROVEEDOR
            'SIN LOS TIPO DE PRODUCTOS
        End If
        Set Rs4 = Nothing
        RST_Busq Rs4, fSqlTipProd_SinRepetir, xCon
        
        RST_Busq Rs1, fStrSqlProv, xCon
        If Trim(TxtIdTipProd.Text) <> "" Then
            RST_Busq Rs2, fSqlTipProd_O_Item(1), xCon
        End If
        If ChkMostrarItem.Value = 1 Then
            RST_Busq Rs3, fSqlTipProd_O_Item(2), xCon
        End If
        LimpiarGrid
        FormatGrid_Nuevo
'    1         2            3            4         5             6
'Tipo Doc., Cliente, Num. Documento, Fec. Doc., Fec. Venc., Cond. Pago,
'    7          8         9           10          11      12         13         14
'Dias Atras., Moneda, Tipo Cambio, Tipo Producto, Item, Uni. Med., Cantidad, Prec. Unit. $
'       15       16            17            18            19         20          21        22
'Imp. Total $, Saldo $, Prec. Unit S/., Imp. Total S/., Saldo S/., Total S/., Abono S/., Saldo S/.
        If vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "DET" Then
            vtitreporte = "REPORTE DETALLADO POR TIPO DE PRODUCTO"
            vSumColGen_SubTot_Dol = 0: vSumColGen_SubTot_Sol = 0: vFlag2 = 0
            ' SOLO SI MARCO EL TIPO DE PRODUCTO
            vSumColGen_SubTot_Dol = 0: vSumColGen_SubTot_Sol = 0: vFlag2 = 0
            vAgregaSubTotal = ""
            
            If Rs1.RecordCount > 0 Then     ' RECORREMOS LOS PROVEEDORES
                PgBar.Max = Rs1.RecordCount
                FraProgreso.Visible = True
                PgBar.Value = 0
                vAgregaSubTotal = "": vSumCol_Dol = 0: vSumCol_Sol = 0
                Rs1.MoveFirst
                Do While Not Rs1.EOF        ' RECORREMOS LOS PROVEEDORES
                    If BAND_INTERRUMPIR = True Then
                        FraProgreso.Visible = False
                        Exit Sub
                    End If
                    DoEvents
                    vNomProv = NulosC(Rs1("nombre"))
                    If Rs2.RecordCount > 0 Then ' RECORREMOS LOS TIPO DE PRODUCTOS
                        Rs2.MoveFirst

                        vFlag = 0: vAgregaSubTotal = "": vSumCol_Dol = 0: vSumCol_Sol = 0
                        Do While Not Rs2.EOF    ' RECORREMOS LOS TIPO DE PRODUCTOS
                            RsCons.Filter = adFilterNone
                            RsCons.Filter = "idpro = " & Val(Rs1("idpro")) & " AND id = " & Val(Rs2("idcomp")) & " AND idtipoprod = " & Val(Rs2("idtipprod")) & ""
                            If RsCons.RecordCount > 0 Then
                                vSubTotal_Sol = 0: vSubTotal_Dol = 0
                                vFlag = 1
                                RsCons.MoveFirst
                                Do While Not RsCons.EOF
                                    vSubTotal_Sol = vSubTotal_Sol + RsCons("subtotal_sol")
                                    vSubTotal_Dol = vSubTotal_Dol + RsCons("subtotal_dol")
                                    RsCons.MoveNext
                                Loop
                                PintarGrid_OptDetalle True, 1
                                ' SUMA PARA EL SUB TOTAL DE COLUMNAS
                                vSumCol_Dol = vSumCol_Dol + vSubTotal_Dol
                                vSumCol_Sol = vSumCol_Sol + vSubTotal_Sol
                                vNomProv = ""
                            End If
                            Rs2.MoveNext
                        Loop ' FIN RECORREMOS LOS TIPO DE PRODUCTOS
                        ' AQUI YA CAMBIA EL PROVEEDOR
                        If vFlag = 1 Then
                            vFlag2 = 1
                            vAgregaSubTotal = "SUBTOT"
                            PintarGrid_OptDetalle True, 1
                            vSumColGen_SubTot_Dol = vSumColGen_SubTot_Dol + vSumCol_Dol
                            vSumColGen_SubTot_Sol = vSumColGen_SubTot_Sol + vSumCol_Sol
                        End If
                    End If
                    If PgBar.Value < PgBar.Max Then
                        PgBar.Value = PgBar.Value + 1
                    End If
                    Rs1.MoveNext
                Loop
                FraProgreso.Visible = False
            End If
                
            If vFlag2 = 1 Then
                vAgregaSubTotal = "TOTGEN"
                PintarGrid_OptDetalle True, 1
            End If
        ElseIf vIndicadorConsul = "DET" And vIndicadorConsulHastaDetalleReal = "DET" Then
            vtitreporte = "REPORTE DETALLADO POR ITEM"
            vIndicaPreProm = 0
            VerifSiSeCalculaPreProm
            ' AQUI MOSTRAR CON TODOS LOS ITEM, TIPO DE PRODUCTO Y LOS PROVEDORES
            vFlag = 0: vFlag2 = 0: vAgregaSubTotal = "": vSumColGen_SubTot_Dol = 0: vSumColGen_SubTot_Sol = 0
            
            If Rs1.RecordCount > 0 Then  ' RECORREMOS LOS PROVEDORES
                FraProgreso.Visible = True
                PgBar.Max = Rs1.RecordCount
                PgBar.Value = 0
                Rs1.MoveFirst
                Do While Not Rs1.EOF     ' RECORREMOS LOS PROVEDORES
                    If BAND_INTERRUMPIR = True Then
                        FraProgreso.Visible = False
                        Exit Sub
                    End If
                    DoEvents
                    RsCons.Filter = adFilterNone
                    If Trim(TxtIdTipProd.Text) <> "" And Fg3.TextMatrix(1, 1) = "" Then
                        RsCons.Filter = adFilterNone
                        RsCons.Filter = "idpro = " & Val(Rs1("idpro")) & " AND idtipoprod = " & Val(TxtIdTipProd.Text) & ""
                        If RsCons.RecordCount > 0 Then
                            ' PA EL PREC PROM
                            If vIndicaPreProm = 1 Then
                                Set rsTemp = RsCons
                            End If
                            PintarGrid_OptDetalle False, RsCons.RecordCount
                        End If
                    Else
                        RsCons.Filter = adFilterNone
                        RsCons.Filter = "idpro = " & Val(Rs1("idpro")) & ""
                        If RsCons.RecordCount > 0 Then
                            PintarGrid_OptDetalle False, RsCons.RecordCount
                        End If
                    End If
                    If PgBar.Value < PgBar.Max Then
                        PgBar.Value = PgBar.Value + 1
                    End If
                    Rs1.MoveNext
                Loop
                
                If vFlag = 1 Then
                    vAgregaSubTotal = "TOTGEN"
                    PintarGrid_OptDetalle False, 1
                    vFlag = 0
                End If
                If vIndicaPreProm = 1 Then
                    fSqlCalPreProm 0, 0, 0, True
                    If RsPreProm.RecordCount > 0 Then
                        ' ESTE PROCEDIMIENTO DEVUELVE EL RECORSET QUE CONTIENE EL PREC PROM
                        Fg1.AddItem ""
                        Fg1.TextMatrix(Fg1.Rows - 1, 6) = "P. Prom. Gen:" '
                        formatTextCeldaGrid Fg1.Rows - 1, 6, &H800000 '
                        If NulosN(RsPreProm("pre_prom_sol")) > 0 Then
                            Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(NulosN(RsPreProm("pre_prom_sol")), FORMAT_MONTO) '
                            formatTextCeldaGrid Fg1.Rows - 1, 18, &H800000 '
                        End If
                        If NulosN(RsPreProm("pre_prom_dol")) > 0 Then
                            Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(NulosN(RsPreProm("pre_prom_dol")), FORMAT_MONTO) '
                            formatTextCeldaGrid Fg1.Rows - 1, 15, &H800000 '
                        End If
                        Fg1.AddItem ""
                    End If
                End If
                FraProgreso.Visible = False
            End If
        ElseIf vIndicadorConsul = "AGR" And vIndicadorConsulHastaDetalleReal = "AGR" Then
            vtitreporte = "REPORTE DETALLADO"
            ' ACA SOLO SI NO SE ESPECIFICO EL TIPO DE PRODUCTO Y LOS ITEM
            LimpiarVaribles
            If Rs1.RecordCount > 0 Then 'RECORREMOS LOS PROVEDORES
                FraProgreso.Visible = True
                PgBar.Max = Rs1.RecordCount
                PgBar.Value = 0
                Rs1.MoveFirst
                Do While Not Rs1.EOF
                    If BAND_INTERRUMPIR = True Then
                        FraProgreso.Visible = False
                        Exit Sub
                    End If
                    DoEvents
                    RsCons.Filter = adFilterNone
                    RsCons.Filter = "idpro = " & Val(Rs1("idpro")) & ""
                    If RsCons.RecordCount > 0 Then
                        PintarGrid_OptDetalle False, RsCons.RecordCount
                    End If
                    If PgBar.Value < PgBar.Max Then
                        PgBar.Value = PgBar.Value + 1
                    End If
                    Rs1.MoveNext
                Loop
                If vFlag = 1 Then
                    vAgregaSubTotal = "TOTGEN"
                    PintarGrid_OptDetalle False, 1
                    vFlag = 0
                End If
                FraProgreso.Visible = False
            End If
        End If
    End If
    If Fg1.Rows - 1 = 2 Then
        MsgBox "No se encontró informacion", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pImprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIMIR LA CONSULTA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pImprimir()
    Dim xform As New FrmPrintConsComAdmin
    xform.propTitulo1 = vtitreporte
    xform.proptitulo2 = vrangofec
    FrmPrintConsComAdmin.capform Me
    FrmPrintConsComAdmin.Show
    xform.capform Me
    xform.Show
    Set xform = Nothing
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim nSQLNotIn  As String
    If TxtIdTipProd.Text = "" Then
        MsgBox "Falta especificar el tipo de item...!", vbExclamation, xTitulo
        TxtIdTipProd.SetFocus
        Exit Sub
    End If
     If ChkMostrarItem.Value = 0 Then
        MsgBox "Seleccione la opcion de mostrar item.", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If Col = 2 Then
        Dim xform As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        Dim xCampos(3, 4) As String
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Cod. Prod.":    xCampos(1, 1) = "codpro":         xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":        xCampos(2, 1) = "id":             xCampos(2, 2) = "800":         xCampos(2, 3) = "N"
        
        nSQLNotIn = GRID_GENERAR_SQL_ID(Fg2, 3, " AND alm_inventario.id", "NOT IN", True)
        
        '--si se ingresa algun filtro adicional
        If NulosC(Fg2.TextMatrix(Row, Col)) <> "" Then
            nSQLNotIn = nSQLNotIn & " AND (UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(Fg2.TextMatrix(Row, Col))) & "%' OR UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(Fg2.TextMatrix(Row, Col))) & "%' ) "
        End If
        
        xform.SQLCad = "SELECT id, codpro, descripcion FROM alm_inventario WHERE tippro = " & NulosN(TxtIdTipProd.Text) & "" & nSQLNotIn
        Fg2.TextMatrix(Row, 2) = ""
        xform.Titulo = "Buscando Tipo de Item"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Fg2.TextMatrix(Row, 1) = NulosC(xRs("codpro"))
            Fg2.TextMatrix(Row, 2) = NulosC(xRs("descripcion"))
            Fg2.TextMatrix(Row, 3) = Trim(xRs("id"))
            If Trim(Fg2.TextMatrix(Row, 2)) <> "" And Trim(Fg2.TextMatrix(Row, 3)) <> "" Then
                If Trim(Fg2.TextMatrix(Fg2.Rows - 1, 3)) <> "" Then Fg2.AddItem ""
                Fg2.Row = Fg2.Rows - 1: Fg2.Col = 2
            End If
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 45  'INSERTAR REGI
            If Fg2.TextMatrix(1, 1) <> "" And Trim(Fg2.TextMatrix(Fg2.Rows - 1, 1)) <> "" Then
                Fg2.AddItem ""
                Fg2.Row = Fg2.Rows - 1: Fg2.Col = 1
            End If
        Case 46 'SUPRIMIR/DELETE
            If Fg2.Row < 1 Then Exit Sub
            If Fg2.Rows - 1 >= 2 Then
                Fg2.RemoveItem Fg2.Row
            Else
                LimpiarGridItem
            End If
    End Select
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub Fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        Dim xform As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        Dim nSQLNotIn As String
        Dim xCampos(3, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Ruc":   xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1500":   xCampos(1, 3) = "C"
        xCampos(2, 0) = "Id":   xCampos(2, 1) = "id":        xCampos(2, 2) = "800":   xCampos(2, 3) = "N"
        
        nSQLNotIn = GRID_GENERAR_SQL_ID(Fg3, 1, " WHERE mae_prov.id", "NOT IN", True)

        ' si se ingresa algun filtro adicional
        If NulosC(Fg3.TextMatrix(Row, Col)) <> "" Then
            nSQLNotIn = IIf(nSQLNotIn = "", " WHERE ", nSQLNotIn & " AND ") & "  (UCASE(mae_prov.nombre) LIKE '%" & UCase(NulosC(Fg3.TextMatrix(Row, Col))) & "%' OR UCASE(mae_prov.nombre) LIKE '%" & UCase(NulosC(Fg3.TextMatrix(Row, Col))) & "%' ) "
        End If
        Fg3.TextMatrix(Row, Col) = ""
        xform.SQLCad = "SELECT * FROM mae_prov " & nSQLNotIn
        
        xform.Titulo = "Buscando proveedores"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "nombre"
        xform.CampoBusca = "nombre"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Fg3.TextMatrix(Row, 1) = NulosN(xRs("id"))
            Fg3.TextMatrix(Row, 2) = NulosC(xRs("nombre"))
            If Trim(Fg3.TextMatrix(Row, 1)) <> "" Then
                If Trim(Fg3.TextMatrix(Fg3.Rows - 1, 1)) <> "" Then Fg3.AddItem ""
                Fg3.Row = Fg3.Rows - 1: Fg3.Col = 2
            End If
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 45  ' INSERTAR REGI
            If Fg3.TextMatrix(1, 1) <> "" And Trim(Fg3.TextMatrix(Fg3.Rows - 1, 1)) <> "" Then
                Fg3.AddItem ""
                Fg3.Row = Fg3.Rows - 1: Fg3.Col = 1
            End If
        Case 46  ' SUMPRIMIR/DELETE REGIST
            If Fg3.Row < 1 Then Exit Sub
            If Fg3.Rows - 1 >= 2 Then
                Fg3.RemoveItem Fg3.Row
            Else
                LimpiarGridProv
            End If
    End Select
End Sub

Private Sub Fg3_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        ' interrumpir
        BAND_INTERRUMPIR = True
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO DEL FORMULARIO A EJECUTARSE
    FormatGrid_Nuevo
    vFormatString = Fg1.FormatString
    vFormatStrGridItem = Fg2.FormatString
    
    vFormatGridProv = Fg3.FormatString
    LimpiarGridItem
    
    Fg3.ColComboList(1) = "|..."
    Fg3.Editable = flexEDKbdMouse
    Fg3.SelectionMode = flexSelectionFree
    LimpiarGridProv
    
    TxtIdTipProd.Text = ""
    lblTipProducto.Caption = ""
    CaracteresNumericos = "0123456789." & Chr(8)
    
    TxtFec1.Valor = CDate("01/01/" + CStr(Year(Date)))
    TxtFec2.Valor = Date
    
    FraProgreso.Width = 5745
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    
    UnirCeldas 0, 14, 16, "Dólares"
    UnirCeldas 0, 17, 19, "Soles"
    UnirCeldas 0, 20, 22, "Totales Soles"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsCons = Nothing: Set rsTemp = Nothing
    Set RsProv = Nothing: Set RsItem = Nothing
    Set Rs1 = Nothing: Set Rs2 = Nothing: Set Rs3 = Nothing
    Set Rs4 = Nothing: Set Rs5 = Nothing: Set RsPreProm = Nothing
    
End Sub

Private Sub TxtIdTipProd_Change()
    If Trim(TxtIdTipProd.Text) = "" Then
        lblTipProducto.Caption = ""
        LimpiarGridItem
    End If
End Sub

Private Sub TxtIdTipProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim RsTipProd As New ADODB.Recordset
        RsTipProd.CursorLocation = adUseClient
        If TxtIdTipProd.Text <> "" Then
            Set RsTipProd = BuscaConCriterio("SELECT id, descripcion FROM mae_tipoproducto WHERE id =" & Val(TxtIdTipProd.Text) & "", xCon)
            If RsTipProd.RecordCount <> 0 Then
                lblTipProducto.Caption = RsTipProd("descripcion")
            Else
                lblTipProducto.Caption = ""
                TxtIdTipProd.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdTipProd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then  'TECHAL F5
        CmdBusProducto.Value = True
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pGenerarConsulta
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
'* Paranetros       : NOMBRE      |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    pProv       |  LONG             |  ESPECIFICA EL ID DEL PROVEEDOR
'*                    pIdTipProd  |  LONG             |  ESPECIFICA EL ID DEL TIPO DE PRODUCTO
'*                    pIdItem     |  LONG             |  ESPECIFICA EL ID DEL ITEM
'*                    pFlagTot    |  BOOLEAN          |  ESPECIFICA SI SE TOTALIZAR EL RESULTADO
'* Devuelve         : STRING
'*****************************************************************************************************
Function pGenerarConsulta(pProv As Long, pIdTipProd As Long, pIdItem As Long, Optional pFlagTot As Boolean = False) As String
    Dim k As Integer
    ' FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    Dim vStrSelect As String        ' CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    Dim vStrFiltro_ITEM As String   ' SOLO ITEM
    Dim vStrFiltro_CLI As String    ' SOLO CLIENTES
    Dim vStrFiltro As String
    Dim vStrFiltro_1 As String      ' ESTE FILTRO SERVIRA PARA CONSULTAR EN EL SUB_SELECT
    ' DE LA FECHA
    If CDate(TxtFec1.Valor) < CDate(TxtFec2.Valor) Then
        vStrFiltro = " com_compras.fchdoc BETWEEN #" + Format(TxtFec1.Valor, "mm/dd/yyyy") + "# AND #" + Format(TxtFec2.Valor, "mm/dd/yyyy") + "# "
    Else
        vStrFiltro = " com_compras.fchdoc = #" + Format(TxtFec1.Valor, "mm/dd/yyyy") + "# "
    End If
    ' SI SE SELECCIONA LA OPCION DE SELECCIONAR POR FECHA DE VENCIMIENTO
    If Me.OptVenc.Value = True Then vStrFiltro = Replace(vStrFiltro, "com_compras.fchdoc", "com_compras.fchven")
    
    ' DEL TIPO DE PRODUCTO
    If TxtIdTipProd.Text <> "" Then vStrFiltro = vStrFiltro + " AND alm_inventario.tippro = " + CStr(TxtIdTipProd.Text) + " "
    ' DEL ITEM
    With Fg2
        For k = 0 To .Rows - 1
            If Me.ChkMostrarItem.Value = 0 Then Exit For ' SALIR SI NO SELECCIONA MOSTRAR ITEM
            If k + 1 = .Rows Then Exit For
            If CStr(.TextMatrix(k + 1, 3)) <> "" Then vStrFiltro_ITEM = vStrFiltro_ITEM + CStr(.TextMatrix(k + 1, 3)) + ","
        Next k
    End With
    If vStrFiltro_ITEM <> "" Then vStrFiltro_ITEM = " AND alm_inventario.id IN (" + Left(vStrFiltro_ITEM, Len(vStrFiltro_ITEM) - 1) + ") "
    ' DEL CLIENTE
    
    With Fg3
        For k = 0 To .Rows - 1
            If k + 1 = .Rows Then Exit For
            If CStr(.TextMatrix(k + 1, 1)) <> "" Then vStrFiltro_CLI = vStrFiltro_CLI + CStr(.TextMatrix(k + 1, 1)) + ","
        Next k
    End With
    If vStrFiltro_CLI <> "" Then vStrFiltro_CLI = " AND com_compras.idpro IN (" + Left(vStrFiltro_CLI, Len(vStrFiltro_CLI) - 1) + ") "
    If pFlagTot = False Then vStrFiltro_CLI = " AND com_compras.idpro = " & pProv & ""
    ' CONCATENAR FECHA + ITEM + CLIENTE
    vStrFiltro = vStrFiltro + vStrFiltro_ITEM + vStrFiltro_CLI

    ' DE LA MONEDA
    If OptSol.Value = True Then vStrFiltro = vStrFiltro + " AND com_compras.idmon= 1 "     ' SOLES
    If Me.OptDol.Value = True Then vStrFiltro = vStrFiltro + " AND com_compras.idmon= 2 "  ' DOLARES
    
    ' SI ES CANCELADO
    If OptPag.Value = True Then vStrFiltro = vStrFiltro + " AND com_compras.impsal = 0 "
    ' SI ES PENDIENTE DE PAGO
    If OptPend.Value = True Then vStrFiltro = vStrFiltro + " AND com_compras.impsal <> 0 "
    
    vStrFiltro_1 = Replace(vStrFiltro, "com_compras.", "com_compras1.")
    vStrFiltro_1 = Replace(vStrFiltro_1, "alm_inventario.", "alm_inventario1.")
    If pFlagTot = False Then
        vStrSelect = "SELECT DISTINCT mae_prov.numruc, mae_prov.nombre AS nomcliente, mae_tipoproducto.descripcion AS desctipcom, alm_inventario.descripcion " _
            & vbCr & " , (SELECT Avg(IIf([com_compras1].[idmon]=2,[com_comprasdet1].[preuni],0)) AS total_dol_d " _
            & vbCr & " FROM com_compras AS com_compras1 INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN com_comprasdet AS com_comprasdet1 ON alm_inventario1.id = com_comprasdet1.iditem) ON com_compras1.id = com_comprasdet1.idcom " _
            & vbCr & " Where " & vStrFiltro_1 & "" _
            & vbCr & " GROUP BY com_compras1.idpro,  alm_inventario1.tippro, alm_inventario1.id " _
            & vbCr & " Having com_compras1.idpro = com_compras.idpro And alm_inventario1.tippro = alm_inventario.tippro And alm_inventario1.id = alm_inventario.id " _
            & vbCr & " ) AS pre_prom_dol, " _
            & vbCr & " (SELECT Avg(IIf([com_compras1].[idmon]=1,[com_comprasdet1].[preuni],0)) AS total_mn_d " _
            & vbCr & " FROM com_compras AS com_compras1 INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN com_comprasdet AS com_comprasdet1 ON alm_inventario1.id = com_comprasdet1.iditem) ON com_compras1.id = com_comprasdet1.idcom " _
            & vbCr & " Where " & vStrFiltro_1 & "" _
            & vbCr & " GROUP BY com_compras1.idpro, alm_inventario1.tippro,alm_inventario1.id " _
            & vbCr & " HAVING com_compras1.idpro=com_compras.idpro  AND alm_inventario1.tippro=alm_inventario.tippro and alm_inventario1.id =alm_inventario.id )  as pre_prom_sol "
        vStrSelect = vStrSelect _
        & vbCr & " FROM mae_unidades RIGHT JOIN ((mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN com_compras ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) INNER JOIN (mae_tipoproducto RIGHT JOIN (alm_inventario RIGHT JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem) ON mae_tipoproducto.id = alm_inventario.tippro) ON com_compras.id = com_comprasdet.idcom) ON mae_unidades.id = alm_inventario.idunimed " _
        & vbCr & " Where " & vStrFiltro _
        & vbCr & " ORDER BY mae_prov.nombre, mae_tipoproducto.descripcion, alm_inventario.descripcion"
    Else
        vStrSelect = "SELECT DISTINCT  mae_tipoproducto.descripcion AS desctipcom, alm_inventario.descripcion " _
            & vbCr & " , (SELECT Avg(IIf([com_compras1].[idmon]=2,[com_comprasdet1].[preuni],0)) AS total_dol_d " _
            & vbCr & " FROM com_compras AS com_compras1 INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN com_comprasdet AS com_comprasdet1 ON alm_inventario1.id = com_comprasdet1.iditem) ON com_compras1.id = com_comprasdet1.idcom " _
            & vbCr & " Where " & vStrFiltro_1 & "" _
            & vbCr & " GROUP BY  alm_inventario1.tippro, alm_inventario1.id " _
            & vbCr & " Having alm_inventario1.tippro = alm_inventario.tippro And alm_inventario1.id = alm_inventario.id " _
            & vbCr & " ) AS pre_prom_dol, " _
            & vbCr & " (SELECT Avg(IIf([com_compras1].[idmon]=1,[com_comprasdet1].[preuni],0)) AS total_mn_d " _
            & vbCr & " FROM com_compras AS com_compras1 INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN com_comprasdet AS com_comprasdet1 ON alm_inventario1.id = com_comprasdet1.iditem) ON com_compras1.id = com_comprasdet1.idcom " _
            & vbCr & " Where " & vStrFiltro_1 & "" _
            & vbCr & " GROUP BY  alm_inventario1.tippro,alm_inventario1.id " _
            & vbCr & " HAVING  alm_inventario1.tippro=alm_inventario.tippro and alm_inventario1.id =alm_inventario.id )  as pre_prom_sol "
        vStrSelect = vStrSelect _
            & vbCr & " FROM mae_unidades RIGHT JOIN ((mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN com_compras ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) INNER JOIN (mae_tipoproducto RIGHT JOIN (alm_inventario RIGHT JOIN com_comprasdet ON alm_inventario.id = com_comprasdet.iditem) ON mae_tipoproducto.id = alm_inventario.tippro) ON com_compras.id = com_comprasdet.idcom) ON mae_unidades.id = alm_inventario.idunimed " _
            & vbCr & " Where " & vStrFiltro & "" _
            & vbCr & " ORDER BY mae_tipoproducto.descripcion, alm_inventario.descripcion"
    End If
    pGenerarConsulta = vStrSelect
End Function

'*****************************************************************************************************
'* Nombre           : pExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ESPORTA A EXCEL LA CONSULTA REALIZADA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportarExcel()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim T_RPT_PERIODO As String
    Dim T_RPT_TITULO As String
    If CDate(TxtFec1.Valor) < CDate(TxtFec2.Valor) Then
        T_RPT_PERIODO = " Del: " + TxtFec1.Valor + " Al: " + TxtFec2.Valor
    Else
        T_RPT_PERIODO = "Al: " + TxtFec2.Valor
    End If
    If OptResum.Value = True Then
        T_RPT_TITULO = "Resumen de Compras"
    Else
        T_RPT_TITULO = "Consulta Detallado de Compras"
    End If
    Me.MousePointer = vbHourglass
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO + " ", "", T_RPT_PERIODO, "Compras"
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportar"
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

Sub ConfiguraGridResumen(pConDetalleSiNo As String)
    ' NO ELIMINAR ESTA FUNCION AUNQUE NO HACE NADA SE INVOCA EN VARIAS PARTES DEL FORMULARIO
End Sub
