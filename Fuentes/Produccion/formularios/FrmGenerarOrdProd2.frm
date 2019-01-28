VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmGenerarOrdProd2 
   Caption         =   "Produccion - Solicitud de Materiales"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frm4 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   3600
      Left            =   5700
      TabIndex        =   21
      Top             =   3780
      Visible         =   0   'False
      Width           =   6120
      Begin VB.Frame Frame8 
         Height          =   495
         Left            =   70
         TabIndex        =   24
         Top             =   3000
         Width           =   5950
         Begin VB.CommandButton Cmd 
            Caption         =   "Elimi&nar"
            Enabled         =   0   'False
            Height          =   330
            Index           =   5
            Left            =   2400
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Personal"
            Top             =   135
            Width           =   1035
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "Eliminar Todos"
            Enabled         =   0   'False
            Height          =   330
            Index           =   6
            Left            =   3435
            TabIndex        =   27
            ToolTipText     =   "Agregar Personal"
            Top             =   135
            Width           =   1200
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Agregar"
            Enabled         =   0   'False
            Height          =   330
            Index           =   3
            Left            =   60
            TabIndex        =   26
            ToolTipText     =   "Agregar Personal"
            Top             =   135
            Width           =   1065
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Seleccionar"
            Enabled         =   0   'False
            Height          =   330
            Index           =   4
            Left            =   1140
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Personal"
            Top             =   135
            Width           =   1065
         End
      End
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   5830
         Picture         =   "FrmGenerarOrdProd2.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   22
         ToolTipText     =   "Cerrar"
         Top             =   50
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   2640
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   360
         Width           =   5925
         _cx             =   10451
         _cy             =   4657
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmGenerarOrdProd2.frx":02EC
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccion de Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   29
         Top             =   45
         Width           =   1635
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   7
         X1              =   -60
         X2              =   6090
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   6
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   4
         X1              =   6090
         X2              =   6090
         Y1              =   0
         Y2              =   3570
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   50
         Top             =   30
         Width           =   6000
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6570
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":03AE
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":08F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":0C84
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":0E08
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":125C
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":1374
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":18B8
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":1DFC
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":1F10
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":2024
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":2478
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":25E4
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2.frx":2B2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Materiales"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Linea"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7080
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   11895
      _cx             =   20981
      _cy             =   12488
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   8388608
      Caption         =   "  &Consulta  |   &Detalle  "
      Align           =   0
      CurrTab         =   1
      FirstTab        =   0
      Style           =   0
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6660
         Left            =   -12450
         TabIndex        =   4
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6135
            Left            =   0
            TabIndex        =   11
            Top             =   420
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   10821
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Fecha"
            Columns(0).DataField=   "fchpro"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Numero"
            Columns(1).DataField=   "numsol"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Producto"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Encargado"
            Columns(3).DataField=   "nomresp"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Tipo"
            Columns(4).DataField=   "destippro"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Procedencia"
            Columns(5).DataField=   "proc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1693"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1614"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1296"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1217"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=6562"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6482"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=5900"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=5821"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1826"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1746"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2223"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2143"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   1.5
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.fgcolor=&H0&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(64)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(67)  =   ":id=35,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   "Named:id=36:Selected"
            _StyleDefs(69)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(70)  =   "Named:id=37:Caption"
            _StyleDefs(71)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(72)  =   "Named:id=38:HighlightRow"
            _StyleDefs(73)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(74)  =   "Named:id=39:EvenRow"
            _StyleDefs(75)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(76)  =   "Named:id=40:OddRow"
            _StyleDefs(77)  =   ":id=40,.parent=33"
            _StyleDefs(78)  =   "Named:id=41:RecordSelector"
            _StyleDefs(79)  =   ":id=41,.parent=34"
            _StyleDefs(80)  =   "Named:id=42:FilterBar"
            _StyleDefs(81)  =   ":id=42,.parent=33"
         End
         Begin VB.Label LblMes 
            AutoSize        =   -1  'True
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   10020
            TabIndex        =   12
            Top             =   90
            Width           =   720
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Solicitud de Materiales"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   45
            TabIndex        =   5
            Top             =   45
            Width           =   11685
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6660
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   11805
         Begin VB.Frame Frame6 
            Height          =   945
            Left            =   30
            TabIndex        =   13
            Top             =   270
            Width           =   11745
            Begin VB.CommandButton CmdBusSup 
               Height          =   240
               Left            =   2070
               Picture         =   "FrmGenerarOrdProd2.frx":2EBE
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   240
               Width           =   240
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchPro 
               Height          =   300
               Left            =   1020
               TabIndex        =   16
               Top             =   540
               Width           =   1320
               _ExtentX        =   2328
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
               Locked          =   -1  'True
            End
            Begin VB.TextBox TxtIdProg 
               Height          =   300
               Left            =   1020
               Locked          =   -1  'True
               MaxLength       =   11
               TabIndex        =   15
               Text            =   "TxtIdProg"
               Top             =   210
               Width           =   1320
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fecha"
               Height          =   195
               Left            =   90
               TabIndex        =   20
               Top             =   615
               Width           =   450
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Supervisor"
               Height          =   195
               Left            =   90
               TabIndex        =   19
               Top             =   255
               Width           =   750
            End
            Begin VB.Label LblNomProg 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblNomProg"
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
               Left            =   2355
               TabIndex        =   18
               Top             =   210
               Width           =   9225
            End
            Begin VB.Label LblProc 
               Caption         =   "LblProc"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   9990
               TabIndex        =   17
               Top             =   570
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   4890
            Index           =   0
            Left            =   30
            TabIndex        =   6
            Top             =   1260
            Width           =   11700
            _cx             =   20637
            _cy             =   8625
            _ConvInfo       =   1
            Appearance      =   2
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmGenerarOrdProd2.frx":2FF0
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
         Begin VB.Frame Frame3 
            Height          =   540
            Left            =   30
            TabIndex        =   7
            Top             =   6100
            Width           =   11730
            Begin VB.CommandButton Cmd 
               Caption         =   "Eliminar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   1
               Left            =   1140
               TabIndex        =   10
               Top             =   130
               Width           =   1065
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "Agregar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   0
               Left            =   45
               TabIndex        =   9
               Top             =   130
               Width           =   1065
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "Detalle"
               Height          =   330
               Index           =   2
               Left            =   2250
               TabIndex        =   8
               Top             =   130
               Width           =   1065
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Solicitud de Materiales"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   45
            TabIndex        =   3
            Top             =   75
            Width           =   11685
         End
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Insertar"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Ver Receta"
      End
   End
End
Attribute VB_Name = "FrmGenerarOrdProd2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim RstOrd As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean               ' para saber cuando se este agregando FILAS AL CONTROL grid de productos
Dim IdMenuActivo As Integer            ' INDICA EL CODIGO DEL MENU ACTIVO
Dim agregados As Integer
Dim xHorIni As Date                    ' ESPECIFICA LA HORA DE INICIO
Dim mCorrelativo As Long               ' para diferenciar la fecha de entrega del pedido cuando se necesite modificar
Dim mIdRegistro&                       ' identificador del registro
Dim mMesActivo As Integer              ' indica el mes activo
Dim cSQL As String
' Definicion de columnas
Dim COLUMNA_SELECCIONADO As Integer
Dim COLUMNA_TIPO As Integer
Dim COLUMNA_ITEM As Integer
Dim COLUMNA_RESPONSABLE As Integer
Dim COLUMNA_UM As Integer
Dim columna_cantidad As Integer
Dim COLUMNA_RECETA As Integer
Dim COLUMNA_LOTE As Integer
Dim columna_idpro As Integer
Dim COLUMNA_IDREC As Integer
Dim COLUMNA_IDUNIMED As Integer
Dim COLUMNA_NUMORDEN As Integer
Dim COLUMNA_IDRESP As Integer
Dim COLUMNA_ID As Integer
Dim COLUMNA_IDCRDET As Integer

Dim NUMERO_CORRELATIVO As Double

Dim numSolMax As Integer
Dim RstValores As New ADODB.Recordset

'Para mover el frame
Dim OrigFX As Long
Dim OrigFY As Long

Private Sub centrarFrm(ByRef frm As Frame)
    With frm
        .Left = ((Me.Width - .Width) / 2)
        .Top = ((Me.Height - .Height) / 2)
    End With
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim A As Integer
    Dim num As Integer
    Dim Rpta As Integer
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    Dim xform As New eps_librerias.FormSeleccion
    
    Select Case Index
        Case 0 ' Agregar Solicitud
            If QueHace = 3 Then Exit Sub
            fg(0).Rows = fg(0).Rows + 1
            fg(0).Select fg(0).Rows - 1, 1
            Frm4.Visible = False
            
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNA_TIPO) = 3
            
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNA_NUMORDEN) = Format(NulosN(numSolMax), "000000")
            numSolMax = numSolMax + 1
            
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNA_ID) = NulosN(NUMERO_CORRELATIVO)
            NUMERO_CORRELATIVO = NUMERO_CORRELATIVO + 1
            
        Case 1 ' Eliminar Solicitud
            If QueHace = 3 Then Exit Sub
            If fg(0).Rows = 2 Then Exit Sub
            
            Rpta = MsgBox("¿Esta seguro de Eliminar esta Solicitud?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            
            If Rpta = vbYes Then
                RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID))
                limpiarRST RstValores, False
                ' Si se elimina el ultimo numero de solicitud se disminuye en uno
                If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_NUMORDEN)) = (numSolMax - 1) Then
                    numSolMax = numSolMax - 1
                End If
                fg(0).RemoveItem fg(0).Row
            End If
            
        Case 2 ' Ver detalle solicitud
            verDetalle
            
        Case 3 ' Agregar item en detalle
            ' Si viene de una receta no se hace nada
            If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 3 Then GoTo ERROR_AL_MOSTRAR
            
            'descripcion                  'campo                           'tamaño                         'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Producto":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":     xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
                    
            cSQL = "SELECT alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.tippro, alm_inventario.id, mae_unidades.abrev AS unimed, alm_inventario.idunimed " _
                + vbCr + "FROM alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                + vbCr + "Where (((alm_inventario.tippro) = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) & ")) " _
                + vbCr + "ORDER BY alm_inventario.descripcion;"
                
            nTitulo = "Buscando Items"
                    
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            If RstValores.State = 0 Then Exit Sub
            
            RstValores.AddNew
            RstValores("idorddet") = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID))    ' Id orden de solicitud
            RstValores("iditem") = NulosN(xRs("id"))                                    ' iD item
            RstValores("descripcion") = NulosC(xRs("descripcion"))                      ' Descripcion
            RstValores("abrev") = NulosC(xRs("unimed"))                                 ' Abrev de UM
            RstValores("cantidad") = 0
            RstValores.Update
            
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
            pCargarValores
            Exit Sub
                
ERROR_AL_MOSTRAR:
            nTitulo = "Error al Procesar Items"
            MsgBox "La receta de un Producto no se puede cambiar", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
            Exit Sub
                    
        Case 4 ' Listar items en detalle
            If QueHace = 3 Then Exit Sub
            
            ' Si viene de una receta no se hace nada
            If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 3 Then GoTo ERROR_AL_MOSTRAR
            
            'descripcion                  'campo                           'tamaño                         'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Producto":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":     xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
                                    
            ' generar la consulta
            cSQL = "SELECT 0 AS xsel, alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.tippro, alm_inventario.id, mae_unidades.abrev AS unimed, alm_inventario.idunimed " _
                + vbCr + "FROM alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                + vbCr + "Where (((alm_inventario.tippro) = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) & ")) " _
                + vbCr + "ORDER BY alm_inventario.descripcion;"
                
            nTitulo = "Buscando Items"
        
            xform.SQLCad = cSQL
                
            xform.titulo = nTitulo
            Set xform.Coneccion = xCon
            Set xRs = Nothing
            Set xRs = xform.seleccionar(xCampos)
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            While Not xRs.EOF
                ' agregando los datos al rst temporal
                RstValores.AddNew
                RstValores("idorddet") = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID))    ' Id orden de solicitud
                RstValores("iditem") = NulosN(xRs("id"))                                    ' iD item
                RstValores("descripcion") = NulosC(xRs("descripcion"))                      ' Descripcion
                RstValores("abrev") = NulosC(xRs("unimed"))                                 ' Abrev de UM
                RstValores("cantidad") = 0
                RstValores.Update
                
                xRs.MoveNext
            Wend
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
            pCargarValores
            
            Set xform = Nothing
            Set xRs = Nothing
        Case 5 ' Eliminar item en detalle
            ' Si viene de una receta no se hace nada
            If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 3 Then GoTo ERROR_AL_MOSTRAR
            
            Agregando = True
            eliminarRegistro
            
        Case 6 ' Eliminar todos los items en detalle
            ' Si viene de una receta no se hace nada
            If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 3 Then GoTo ERROR_AL_MOSTRAR
            
            num = fg(1).Rows - 1
            For A = 1 To num
                Agregando = False
                If fg(1).Rows > fg(1).FixedRows Then
                    fg(1).Select 1, 1
                    eliminarRegistro
                End If
            Next A
            pCargarValores
    End Select
End Sub

Private Sub eliminarRegistro()
    If fg(1).Row < 1 Then
        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(1).SetFocus
        Exit Sub
    End If
    
    If fg(1).Rows = 1 Then
        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(1).SetFocus
        Exit Sub
    End If
    
    If fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO) = 3 Then Exit Sub
    
    If Agregando Then
        If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    End If
    
    If RstValores.RecordCount <> 0 Then RstValores.MoveFirst
        
    Do While Not RstValores.EOF
        If RstValores.RecordCount = 0 Then Exit Do
        If NulosN(RstValores("iditem")) = NulosN(fg(1).TextMatrix(fg(1).Row, 5)) Then
            RstValores.Delete
            Exit Do
        End If
        RstValores.MoveNext
    Loop
    pCargarValores
End Sub

Private Sub pCargarValores()
    Agregando = True
    
    If RstValores.State = 0 Then Exit Sub
    If RstValores.RecordCount = 0 Then Exit Sub
    
    RstValores.MoveFirst
    
    With fg(1)
        .Rows = 1
        Do While Not RstValores.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = NulosN(RstValores("activo"))
            .TextMatrix(.Rows - 1, 2) = NulosC(RstValores("descripcion"))
            .TextMatrix(.Rows - 1, 3) = NulosC(RstValores("abrev"))
            .TextMatrix(.Rows - 1, 4) = Format(NulosN(RstValores("cantidad")), "0.00")
            .TextMatrix(.Rows - 1, 5) = NulosC(RstValores("idorddet"))
            .TextMatrix(.Rows - 1, 6) = NulosC(RstValores("iditem"))
            
            RstValores.MoveNext
        Loop
    End With
    
    Agregando = False
End Sub

Private Sub CmdBusSup_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'Dim cSQL As String
    Dim xCampos(2, 4) As String
    
    If QueHace = 3 Then Exit Sub
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":                xCampos(1, 1) = "id":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    cSQL = "SELECT pro_emp.*, pla_empleados.nombre " _
        + vbCr + "FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
        + vbCr + "Where (((pro_empdet.idfun) = 2)) " _
        + vbCr + "ORDER BY pla_empleados.nombre;"
    
    xform.SQLCad = cSQL
    
    xform.titulo = "Buscando Supervisores"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdProg.Text = xRs("id")
            LblNomProg.Caption = xRs("nombre")
            TxtFchPro.valor = Date
            TxtFchPro.SetFocus
        End If
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Sub ActivaEntorno(xQueFue As Boolean)
    TabOne1.Enabled = xQueFue
    Toolbar1.Enabled = xQueFue
End Sub

Private Sub PosicionarFrm(ByRef FRM_ As Frame)
    With FRM_
        .Top = Me.Height - 4155
        .Left = Me.Width - 6300
    End With
End Sub

Private Sub verDetalle()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
            
    ' Si no se han agregado productos
    If fg(0).Rows = 1 Then
        'xTitulo = "Error al mostrar "
        MsgBox "No hay items que mostrar, agreguelos Productos para procesarlos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).SetFocus
        
        PosicionarFrm Frm4
        Frm4.Visible = True
        fg(1).Rows = 1
    End If
    
    ' Si no se ha escogido Tipo
    If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 0 Then
        'xTitulo = "Error al mostrar "
        MsgBox "No hay items que mostrar, agregue Tipo para procesarlos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).SetFocus
        Exit Sub
    End If
    
    ' Si no hay Productos escogidos
    If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 3 And _
                                    NulosN(fg(0).TextMatrix(fg(0).Row, columna_idpro)) = 0 Then
        'xTitulo = "Error al mostrar "
        MsgBox "No hay items que mostrar, agregue Item para procesarlos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).SetFocus
        Exit Sub
    End If
    
    PosicionarFrm Frm4
    Frm4.Visible = True
    
    If RstValores.State = 0 Then Exit Sub
    RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID))
    
    If RstValores.RecordCount = 0 Then
        cargarReceta NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREC)), NulosN(fg(0).TextMatrix(fg(0).Row, columna_cantidad))
    End If
    
    RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
    pCargarValores
End Sub

Private Sub cargarReceta(ID_RECETA As Integer, cantidad As Double)
    Dim A As Integer
    
    cSQL = "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]* " & cantidad & " AS canreq " _
        + vbCr + " FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        + vbCr + " WHERE (((pro_recetains.idrec)=" & ID_RECETA & "));"
        
    RST_Busq Rst, cSQL, xCon

    If Rst.State = 0 Then Exit Sub
    If Rst.RecordCount = 0 Then Exit Sub
    
    If RstValores.State = 0 Then Exit Sub
    
    For A = 1 To Rst.RecordCount
        RstValores.AddNew
        RstValores("activo") = -1                                                   ' Activo
        RstValores("idorddet") = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID))    ' Id orden de solicitud
        RstValores("iditem") = NulosN(Rst("iditem"))                                ' ID item
        RstValores("descripcion") = NulosC(Rst("descripcion"))                      ' Descripcion
        RstValores("abrev") = NulosC(Rst("abrev"))                                  ' Abrev de UM
        RstValores("cantidad") = NulosN(Rst("canreq"))                              ' Cantidad
        RstValores.Update
        
        Rst.MoveNext
        If Rst.EOF = True Then Exit For
    Next A
End Sub

Sub CrearCabeceraVS(numPag As Integer)
    Dim xCad As String

    FrmVsPrinter.Vs.TextAlign = taLeftTop
    FrmVsPrinter.Vs.FontName = "Courier New"
    FrmVsPrinter.Vs.FontBold = True
    FrmVsPrinter.Vs.FontSize = 9

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 200
    FrmVsPrinter.Vs.Paragraph = "EMPRESA   : " & NomEmp

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 200
    FrmVsPrinter.Vs.Paragraph = "FECHA        : " & Format(Date, "dd/mm/yy")

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 400
    FrmVsPrinter.Vs.Paragraph = "Nº R.U.C. : " & NumRUC

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 400
    FrmVsPrinter.Vs.Paragraph = "Nº Pagina    : " & Format(numPag, "0000")

    FrmVsPrinter.Vs.DrawLine 1000, 650, 11000, 650
End Sub

Private Sub ImprimirSolicitud()
    Dim A As Integer
    Dim B As Integer
    Dim Rst As New ADODB.Recordset
    
    With FrmVsPrinter.Vs
        Dim xLinea As Integer
        Dim NUMEROPAG_ As Integer
        Dim INICIOPAG_ As Integer
        Dim TOPEPAG_ As Integer
        
        .BrushColor = &H80000005
        .FontSize = 11
        .TextAlign = taCenterMiddle
        .StartDoc
        
        xLinea = 1000
        NUMEROPAG_ = 1
        INICIOPAG_ = 1000
        TOPEPAG_ = 14500
        CrearCabeceraVS NUMEROPAG_
        
        For A = 1 To fg(0).Rows - 1
            If xLinea >= TOPEPAG_ Then
                xLinea = INICIOPAG_
                .NewPage
                NUMEROPAG_ = NUMEROPAG_ + 1
                CrearCabeceraVS NUMEROPAG_
            End If
            
            If fg(0).TextMatrix(A, COLUMNA_SELECCIONADO) <> -1 Then GoTo SIGUIENTE
            'LADO A
            .FontSize = 13
            .TextAlign = taCenterMiddle
            .TextBox "SOLICITUD DE MATERIALES", 500, xLinea, 6700, 500, True, False, True
            .FontSize = 10
            .TextAlign = taCenterTop
            .TextBox "Nº ", 7300, xLinea, 1700, 250, True, False, True
            xLinea = xLinea + 240
            .TextBox "0001" & "-" & fg(0).TextMatrix(A, COLUMNA_NUMORDEN), 7300, xLinea, 1700, 250, True, False, True
            
            .TextAlign = taLeftMiddle
            .FontSize = 9
            
            xLinea = xLinea + 400
            If fg(0).TextMatrix(A, COLUMNA_TIPO) = 3 Then
                .TextBox "Producto    ", 500, xLinea, 1500, 250, True, False, False
                .TextBox fg(0).TextMatrix(A, COLUMNA_ITEM), 2000, xLinea, 6000, 250, True, False, False
                xLinea = xLinea + 250
            End If
            
            .TextBox "Programador    ", 500, xLinea, 1500, 250, True, False, False
            If NulosN(xIdUsuario) = 0 Then
                .TextBox NulosC(LblNomProg.Caption), 2000, xLinea, 6000, 250, True, False, False
            Else
                Dim Nombre As String
                Nombre = Busca_Codigo(xIdUsuario, "id", "nomusu", "mae_usuarios", "N", xCon)
                .TextBox Nombre, 2000, xLinea, 6000, 250, True, False, False
            End If
            
            .TextBox "Fch. Prod.   ", 6000, xLinea, 1500, 250, True, False, False
            .TextBox TxtFchPro.valor, 7200, xLinea, 6000, 250, True, False, False
            
            If fg(0).TextMatrix(A, COLUMNA_TIPO) = 3 Then
                xLinea = xLinea + 250
                .TextBox "Receta ", 500, xLinea, 1500, 250, True, False, False
                .TextBox fg(0).TextMatrix(A, COLUMNA_RECETA), 2000, xLinea, 1500, 250, True, False, False
            
                .TextBox "Cantidad   ", 6000, xLinea, 1500, 250, True, False, False
                .TextBox fg(0).TextMatrix(A, columna_cantidad), 7200, xLinea, 6000, 250, True, False, False
            End If
            
            xLinea = xLinea + 250
            .TextBox "Lote   ", 500, xLinea, 1500, 250, True, False, False
            .TextBox fg(0).TextMatrix(A, COLUMNA_LOTE), 2000, xLinea, 1500, 250, True, False, False
            
            xLinea = xLinea + 300
            
            If xLinea >= TOPEPAG_ Then
                xLinea = INICIOPAG_
                .NewPage
                NUMEROPAG_ = NUMEROPAG_ + 1
                CrearCabeceraVS NUMEROPAG_
            End If
            
            .TextAlign = taCenterMiddle
            .TextBox "Item", 500, xLinea, 400, 500, True, False, True
            .TextBox "INSUMO / PRODUCTO / MP", 900, xLinea, 3700, 500, True, False, True
            .TextBox "U.M.", 4600, xLinea, 400, 500, True, False, True
            .TextBox "Cantidad Teorica", 5000, xLinea, 1000, 500, True, False, True
            .TextBox "Cantidad Real", 6000, xLinea, 1000, 500, True, False, True
            .TextBox "Adicional", 7000, xLinea, 1000, 500, True, False, True
            .TextBox "Devolucion", 8000, xLinea, 1000, 500, True, False, True
                        
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(A, COLUMNA_ID)) & ""
            
            If RstValores.RecordCount = 0 Then
                cSQL = "SELECT pro_ordenproddetins.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_ordenproddetins.cantidad, pro_ordenproddetins.activo " _
                    + vbCr + "FROM ((pro_ordenproddet RIGHT JOIN pro_ordenproddetins ON pro_ordenproddet.id = pro_ordenproddetins.idorddet) LEFT JOIN alm_inventario ON pro_ordenproddetins.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                    + vbCr + "Where (((pro_ordenproddet.numDoc) = '" & NulosC(fg(0).TextMatrix(A, COLUMNA_NUMORDEN)) & "')) " _
                    + vbCr + "GROUP BY pro_ordenproddetins.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_ordenproddetins.cantidad, pro_ordenproddetins.activo;"
                
                RST_Busq Rst, cSQL, xCon
            Else
                DEFINIR_RST_TMP Rst, RstValores
                CARGAR_RST_TMP Rst, RstValores
            End If
            
            If Rst.State = 0 Then Exit Sub
            
            If Rst.RecordCount = 0 Then
                Set Rst = Nothing

                cSQL = "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]*2050 AS cantidad, -1 AS activo " _
                    + vbCr + "FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
                    + vbCr + "WHERE (((pro_recetains.idrec)=" & NulosN(fg(0).TextMatrix(A, COLUMNA_IDREC)) & "));"
                
                RST_Busq Rst, cSQL, xCon
            End If
            
            Rst.Filter = "activo = -1"
        
            If Rst.RecordCount <> 0 Then
                Dim xFila As Integer
                xLinea = xLinea + 500
                xFila = xLinea
                For B = 1 To Rst.RecordCount
                    .FontSize = 8
                    .TextAlign = taLeftMiddle
                    .TextBox " " & Format(B, "00"), 500, xLinea, 400, 250, True, False, True
                    .TextBox " " & Rst("descripcion"), 900, xLinea, 3700, 250, True, False, True
                    .TextAlign = taCenterMiddle
                    .TextBox Rst("abrev"), 4600, xLinea, 400, 250, True, False, True
                    .TextAlign = taRightMiddle
                    .TextBox Format(Rst("cantidad"), "0.000000"), 5000, xLinea, 1000, 250, True, False, True
                    .TextBox "", 6000, xLinea, 1000, 250, True, False, True
                    .TextBox "", 7000, xLinea, 1000, 250, True, False, True
                    .TextBox "", 8000, xLinea, 1000, 250, True, False, True
                    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                    
                    xLinea = xLinea + 250
                    
                    If xLinea >= TOPEPAG_ Then
                        xLinea = INICIOPAG_
                        .NewPage
                        NUMEROPAG_ = NUMEROPAG_ + 1
                        CrearCabeceraVS NUMEROPAG_
                    End If
                Next B
            End If
            
            ' POSICION ANTES DEL DETALLE + ALTO DE DE 10 ITEMS + 500 DE ESPACIO
            xLinea = xLinea + 500
            If xLinea >= TOPEPAG_ Then
                xLinea = INICIOPAG_
                .NewPage
                NUMEROPAG_ = NUMEROPAG_ + 1
                CrearCabeceraVS NUMEROPAG_
            End If
            'LADO A
            .TextBox "_______________________________", 900, xLinea, 3500, 200, True, False, False
            .TextBox "_______________________________", 5400, xLinea, 3500, 200, True, False, False
            
            xLinea = xLinea + 200
            
            .FontSize = 6
            .TextAlign = taCenterMiddle
            
            If NulosC(fg(0).TextMatrix(A, COLUMNA_RESPONSABLE)) = "" Then
                .TextBox "VºBº Ger. Prod. ", 1500, xLinea, 3000, 250, True, False, False
            Else
                .TextBox NulosC(fg(0).TextMatrix(A, COLUMNA_RESPONSABLE)), 1500, xLinea, 3000, 250, True, False, False
            End If
            
            .TextBox "Responsable de Almacen", 6000, xLinea, 3000, 250, True, False, False
            .FontSize = 8
            
            xLinea = xLinea + 500
SIGUIENTE:
        Next A
        .EndDoc
    End With
    'Muestra la preimagen de la impresion
    FrmVsPrinter.WindowState = 2
    FrmVsPrinter.Show
End Sub

'Sub CrearCabeceraVS(numPag As Integer)
'    Dim xCad As String
'
'    FrmVsPrinter.Vs.TextAlign = taLeftTop
'    FrmVsPrinter.Vs.FontName = "Courier New"
'    FrmVsPrinter.Vs.FontBold = True
'    FrmVsPrinter.Vs.FontSize = 9
'
'    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 600
'    FrmVsPrinter.Vs.Paragraph = "EMPRESA   : " & NomEmp
'
'    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 600
'    FrmVsPrinter.Vs.Paragraph = "FECHA        : " & Format(Date, "dd/mm/yy")
'
'    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 800
'    FrmVsPrinter.Vs.Paragraph = "Nº R.U.C. : " & NumRUC
'
'    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 800
'    FrmVsPrinter.Vs.Paragraph = "Nº Pagina    : " & Format(numPag, "0000")
'
'    FrmVsPrinter.Vs.DrawLine 1000, 1050, 11000, 1050
'End Sub

Private Sub ImprimirLinea()
    Dim A As Integer
    Dim numPag As Integer
    Dim Rst As New ADODB.Recordset
    Dim B As Integer
    Dim xLinea As Integer
    Dim xColumna As Integer
    Dim numper As Double
    
    With FrmVsPrinter.Vs
        numPag = 0
        .BrushColor = &H80000005
        .FontSize = 11
        .TextAlign = taCenterMiddle
        .StartDoc
        
        xColumna = 1000
'        xLinea = 700
'        CrearCabeceraVS numPag
'        xLinea = xLinea + 600
        For A = 1 To fg(0).Rows - 1
            If fg(0).TextMatrix(A, COLUMNA_TIPO) <> 3 Then GoTo SIGUIENTE
            
            If A < fg(0).Rows - 1 Then
                xLinea = 1300
                numPag = numPag + 1
                If A > 1 Then .NewPage
                CrearCabeceraVS numPag
            End If
                        
            '******************************************************************* Titulo
            .FontSize = 12
            .FontBold = True
            .TextAlign = taCenterMiddle
            
            .TextBox "LINEA DE PRODUCCION", xColumna, xLinea, 8000, 500, True, False, True
            .FontSize = 10
            .TextAlign = taCenterTop
            .TextBox "Nº ", xColumna + 8100, xLinea, 1900, 250, True, False, True
            xLinea = xLinea + 240
            .TextBox "001" & "-" & fg(0).TextMatrix(A, COLUMNA_NUMORDEN), xColumna + 8100, xLinea, 1900, 250, True, False, True
            
            .TextAlign = taLeftMiddle
            .FontSize = 9
            
            '******************************************************************* Detalle de la Linea
            xLinea = xLinea + 300
            .FontBold = True
            .TextBox "Detalles de la Solicitud", xColumna, xLinea, 3500, 250, True, False, False
            
            '*************************************************************************
            .FontBold = False
            xLinea = xLinea + 250
            .TextBox "Producto", xColumna, xLinea, 1500, 250, True, False, False
            .TextBox fg(0).TextMatrix(A, COLUMNA_ITEM), xColumna + 1500, xLinea, 7000, 250, True, False, False
            
            '*************************************************************************
            xLinea = xLinea + 250
            .TextBox "Fecha Prog.", xColumna, xLinea, 1500, 250, True, False, False
            .TextBox TxtFchPro.valor, xColumna + 1500, xLinea, 6000, 250, True, False, False
            
            '*************************************************************************
            xLinea = xLinea + 250
            .TextBox "Receta", xColumna, xLinea, 1500, 250, True, False, False
            .TextBox fg(0).TextMatrix(A, COLUMNA_RECETA), xColumna + 1500, xLinea, 6000, 250, True, False, False
            
            .TextBox "Cantidad", xColumna + 6500, xLinea, 1500, 250, True, False, False
            .TextBox fg(0).TextMatrix(A, columna_cantidad), xColumna + 7700, xLinea, 6000, 250, True, False, False
            
            '*************************************************************************
            xLinea = xLinea + 250
            .TextBox "Responsable ", xColumna, xLinea, 1500, 250, True, False, False
            .TextBox fg(0).TextMatrix(A, COLUMNA_RESPONSABLE), xColumna + 1500, xLinea, 6000, 250, True, False, False
            
            '*************************************************************************
            xLinea = xLinea + 350
            .TextAlign = taCenterMiddle
            .TextBox "Item", xColumna, xLinea, 500, 500, True, False, True
            .TextBox "INSUMO / PRODUCTO / MP", xColumna + 500, xLinea, 4500, 500, True, False, True
            .TextBox "U.M.", xColumna + 5000, xLinea, 500, 500, True, False, True
            .TextBox "Cantidad Teorica", xColumna + 5500, xLinea, 1125, 500, True, False, True
            .TextBox "Cantidad Real", xColumna + 6625, xLinea, 1125, 500, True, False, True
            .TextBox "Adicional", xColumna + 7750, xLinea, 1125, 500, True, False, True
            .TextBox "Devolucion", xColumna + 8875, xLinea, 1125, 500, True, False, True
            
            cSQL = "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]*" & NulosN(fg(0).TextMatrix(A, columna_cantidad)) & " AS canreq " _
                + vbCr + "FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
                + vbCr + "WHERE (((pro_recetains.idrec)=" & NulosN(fg(0).TextMatrix(A, COLUMNA_IDREC)) & "));"
            
            RST_Busq Rst, cSQL, xCon
        
            If Rst.RecordCount <> 0 Then
                Dim xFila As Integer
                xLinea = xLinea + 500
                xFila = xLinea
                For B = 1 To Rst.RecordCount
                    .FontSize = 8
                    .FontBold = False
                    .TextAlign = taLeftMiddle
                    
                    .TextBox " " & Format(B, "00"), xColumna, xLinea, 500, 250, True, False, True
                    .TextBox " " & Rst("descripcion"), xColumna + 500, xLinea, 4500, 250, True, False, True
                    .TextAlign = taCenterMiddle
                    .TextBox Rst("abrev"), xColumna + 5000, xLinea, 500, 250, True, False, True
                    .TextAlign = taRightMiddle
                    .TextBox Format(Rst("canreq"), "0.000000"), xColumna + 5500, xLinea, 1125, 250, True, False, True
                    .TextBox "", xColumna + 6625, xLinea, 1125, 250, True, False, True
                    .TextBox "", xColumna + 7750, xLinea, 1125, 250, True, False, True
                    .TextBox "", xColumna + 8875, xLinea, 1125, 250, True, False, True
                    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                    
                    xLinea = xLinea + 250
                    
                    If xLinea >= 16200 Then
                        xLinea = 1300
                        numPag = numPag + 1
                        .NewPage
                        CrearCabeceraVS numPag
                    End If
                Next B
            End If
            
             '******************************************************************* Detalle de la Linea
            xLinea = xLinea + 300
            .TextAlign = taLeftMiddle
            .FontBold = True
            .TextBox "Detalles de la Linea", xColumna, xLinea, 2500, 250, True, False, False
            '*************************************************************************
            
            .FontBold = False
            xLinea = xLinea + 350
            .TextAlign = taCenterMiddle
            .TextBox "Orden", xColumna, xLinea, 500, 500, True, False, True
            .TextBox "TAREA", xColumna + 500, xLinea, 3500, 500, True, False, True
            .TextBox "Durac.", xColumna + 4000, xLinea, 800, 500, True, False, True
            .TextBox "Hor.Ini.", xColumna + 4800, xLinea, 800, 500, True, False, True
            .TextBox "Hor.Fin", xColumna + 5600, xLinea, 800, 500, True, False, True
            .TextBox "# Per.", xColumna + 6400, xLinea, 500, 500, True, False, True
            .TextBox "Cant.", xColumna + 6900, xLinea, 750, 500, True, False, True
            .TextBox "%Rdto", xColumna + 7650, xLinea, 500, 500, True, False, True
            .TextBox "Cant. Proc.", xColumna + 8150, xLinea, 750, 500, True, False, True
            .TextBox "Fech. Ini.", xColumna + 8900, xLinea, 1100, 500, True, False, True
            
            cSQL = "SELECT pro_cronogramatarea.*, alm_inventario.descripcion AS matpri, alm_inventario_1.descripcion AS despro, pro_tareas.descripcion AS destar " _
                + vbCr + "FROM ((pro_cronogramatarea LEFT JOIN alm_inventario ON pro_cronogramatarea.iditem = alm_inventario.id) LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_cronogramatarea.idpro = alm_inventario_1.id) LEFT JOIN pro_tareas ON pro_cronogramatarea.idtar = pro_tareas.id " _
                + vbCr + "Where (((pro_cronogramatarea.idcrdet) = " & NulosN(fg(0).TextMatrix(A, COLUMNA_IDCRDET)) & ") And ((pro_cronogramatarea.activo) = -1)) " _
                + vbCr + "ORDER BY pro_cronogramatarea.fchpro, pro_cronogramatarea.horpro, pro_cronogramatarea.idcrdet, pro_cronogramatarea.orden;"
            
            Set Rst = Nothing
            RST_Busq Rst, cSQL, xCon
            numper = 0
            If Rst.RecordCount <> 0 Then
                xLinea = xLinea + 500
                xFila = xLinea
                For B = 1 To Rst.RecordCount
                    .FontSize = 8
                    .FontBold = False
                    
                    .TextAlign = taLeftMiddle
                    .TextBox " " & Format(Rst("orden"), "00"), xColumna, xLinea, 500, 250, True, False, True
                    .TextBox " " & Rst("destar"), xColumna + 500, xLinea, 3500, 250, True, False, True
                    .TextAlign = taCenterMiddle
                    .TextBox Format(Rst("durtar"), "HH:mm"), xColumna + 4000, xLinea, 800, 250, True, False, True
                    .TextBox Format(Rst("horinitar"), "HH:mm"), xColumna + 4800, xLinea, 800, 250, True, False, True
                    .TextBox Format(Rst("horfintar"), "HH:mm"), xColumna + 5600, xLinea, 800, 250, True, False, True
                    .TextBox Rst("numper"), xColumna + 6400, xLinea, 500, 250, True, False, True
                    .TextBox Rst("cantidad"), xColumna + 6900, xLinea, 750, 250, True, False, True
                    .TextBox Rst("aplpor"), xColumna + 7650, xLinea, 500, 250, True, False, True
                    .TextBox Rst("cantproc"), xColumna + 8150, xLinea, 750, 250, True, False, True
                    .TextBox Rst("fchini"), xColumna + 8900, xLinea, 1100, 250, True, False, True
                    
                    numper = numper + NulosN(Rst("numper"))
                    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                    
                    xLinea = xLinea + 250
                    
                    If xLinea >= 16200 Then
                        xLinea = 1300
                        numPag = numPag + 1
                        .NewPage
                        CrearCabeceraVS numPag
                    End If
                Next B
                xLinea = xLinea + 250
                .TextAlign = taRightMiddle
                .TextBox "TOTAL", xColumna, xLinea, 4000, 250, True, False, True
                .TextAlign = taCenterMiddle
                .TextBox numper, xColumna + 6400, xLinea, 500, 250, True, False, True
            Else
                xLinea = xLinea + 500
                xFila = xLinea
                For B = 1 To 10
                    .FontSize = 8
                    .FontBold = False
                    
                    .TextAlign = taLeftMiddle
                    .TextBox "", xColumna, xLinea, 500, 250, True, False, True
                    .TextBox "", xColumna + 500, xLinea, 3500, 250, True, False, True
                    .TextAlign = taCenterMiddle
                    .TextBox "", xColumna + 4000, xLinea, 800, 250, True, False, True
                    .TextBox "", xColumna + 4800, xLinea, 800, 250, True, False, True
                    .TextBox "", xColumna + 5600, xLinea, 800, 250, True, False, True
                    .TextBox "", xColumna + 6400, xLinea, 500, 250, True, False, True
                    .TextBox "", xColumna + 6900, xLinea, 750, 250, True, False, True
                    .TextBox "", xColumna + 7650, xLinea, 500, 250, True, False, True
                    .TextBox "", xColumna + 8150, xLinea, 750, 250, True, False, True
                    .TextBox "", xColumna + 8900, xLinea, 1100, 250, True, False, True
                    
                    xLinea = xLinea + 250
                    
                    If xLinea >= 16200 Then
                        xLinea = 1300
                        numPag = numPag + 1
                        .NewPage
                        CrearCabeceraVS numPag
                    End If
                Next B
                numper = 15
                .TextAlign = taRightMiddle
                .TextBox "TOTAL", xColumna, xLinea, 4000, 250, True, False, True
                .TextAlign = taCenterMiddle
                .TextBox "", xColumna + 6400, xLinea, 500, 250, True, False, True
            End If
            
            '****************************************************************************************
            '******************************************************************* Detalle del Personal
            '****************************************************************************************
            xLinea = xLinea + 300
            .TextAlign = taLeftMiddle
            .FontBold = True
            .TextBox "Detalles del Personal", xColumna, xLinea, 2500, 250, True, False, False
            '*************************************************************************
            
            .FontBold = False
            xLinea = xLinea + 250
            .TextBox "MP Entregada", xColumna, xLinea, 1500, 250, True, False, False
            
            .TextBox "Prod. Elaborado", xColumna + 6500, xLinea, 1500, 250, True, False, False
            
            '*************************************************************************
            xLinea = xLinea + 250
            .TextBox "Hora Ini.", xColumna, xLinea, 1500, 250, True, False, False
            
            .TextBox "Hora Fin", xColumna + 6500, xLinea, 900, 250, True, False, False
            '*************************************************************************
            
            xLinea = xLinea + 350
            .TextAlign = taCenterMiddle
            .TextBox "Item", xColumna, xLinea, 500, 500, True, False, True
            .TextBox "PERSONAL", xColumna + 500, xLinea, 3500, 500, True, False, True
            .TextBox "Codigo", xColumna + 4000, xLinea, 1000, 500, True, False, True
            .TextBox "Tarea", xColumna + 5000, xLinea, 3500, 500, True, False, True

            cSQL = "SELECT pro_cronogramapers.id, pro_cronogramapers.idcr, pro_cronogramapers.idcrdet, pro_cronogramapers.iditem, pro_cronogramapers.idpro, pro_cronogramapers.idtar, pro_cronogramapers.orden, pro_cronogramapers.idper, pla_empleados.codigo, pla_empleados.nombre, pro_cronogramapers.activo, pro_tareas.descripcion, pro_cronogramatarea.activo " _
                + vbCr + "FROM ((pro_cronogramapers LEFT JOIN pla_empleados ON pro_cronogramapers.idper = pla_empleados.id) LEFT JOIN pro_tareas ON pro_cronogramapers.idtar = pro_tareas.id) LEFT JOIN pro_cronogramatarea ON (pro_cronogramapers.idtar = pro_cronogramatarea.idtar) AND (pro_cronogramapers.idcrdet = pro_cronogramatarea.idcrdet) AND (pro_cronogramapers.idcr = pro_cronogramatarea.idcr) " _
                + vbCr + "GROUP BY pro_cronogramapers.id, pro_cronogramapers.idcr, pro_cronogramapers.idcrdet, pro_cronogramapers.iditem, pro_cronogramapers.idpro, pro_cronogramapers.idtar, pro_cronogramapers.orden, pro_cronogramapers.idper, pla_empleados.codigo, pla_empleados.nombre, pro_cronogramapers.activo, pro_tareas.descripcion, pro_cronogramatarea.activo " _
                + vbCr + "HAVING (((pro_cronogramapers.idcrdet) = " & NulosN(fg(0).TextMatrix(A, COLUMNA_IDCRDET)) & ") AND ((pro_cronogramatarea.activo)= true));"
                        
            Set Rst = Nothing
            RST_Busq Rst, cSQL, xCon
        
            If Rst.RecordCount <> 0 Then
                xLinea = xLinea + 500
                xFila = xLinea
                For B = 1 To numper
                    .FontSize = 8
                    .FontBold = False
                    .TextAlign = taLeftMiddle
                    
                    .TextBox " " & Format(B, "00"), xColumna, xLinea, 500, 250, True, False, True
                    If Not Rst.EOF Then
                        .TextBox " " & Rst("nombre"), xColumna + 500, xLinea, 3500, 250, True, False, True
                        .TextBox NulosC(Rst("codigo")), xColumna + 4000, xLinea, 1000, 250, True, False, True
                        .TextBox " " & Format(Rst("descripcion"), "HH:mm"), xColumna + 5000, xLinea, 3500, 250, True, False, True
                        Rst.MoveNext
                    Else
                        .TextBox "", xColumna + 500, xLinea, 3500, 250, True, False, True
                        .TextBox "", xColumna + 4000, xLinea, 1000, 250, True, False, True
                        .TextBox "", xColumna + 5000, xLinea, 3500, 250, True, False, True
                    End If
                    
                    xLinea = xLinea + 250
                    
                    If xLinea >= 16200 Then
                        xLinea = 1300
                        numPag = numPag + 1
                        .NewPage
                        CrearCabeceraVS numPag
                    End If
                Next B
            Else
                xLinea = xLinea + 500
                xFila = xLinea
                For B = 1 To numper
                    .FontSize = 8
                    .FontBold = False
                    .TextAlign = taLeftMiddle
                    
                    .TextBox " " & Format(B, "00"), xColumna, xLinea, 500, 250, True, False, True
                    .TextBox "", xColumna + 500, xLinea, 3500, 250, True, False, True
                    .TextBox "", xColumna + 4000, xLinea, 1000, 250, True, False, True
                    .TextBox "", xColumna + 5000, xLinea, 3500, 250, True, False, True
                                        
                    xLinea = xLinea + 250
                    
                    If xLinea >= 16200 Then
                        xLinea = 1300
                        numPag = numPag + 1
                        .NewPage
                        CrearCabeceraVS numPag
                    End If
                Next B
                xLinea = xLinea - 250
            End If
            
            
            '****************************************************************************************
             '******************************************************************* Observaciones
             '****************************************************************************************
            xLinea = xLinea + 300
            
            If xLinea >= 15500 Then
                xLinea = 1300
                numPag = numPag + 1
                .NewPage
                CrearCabeceraVS numPag
            End If
            
            .TextAlign = taLeftMiddle
            .FontBold = True
            .TextBox "Observaciones", xColumna, xLinea, 2500, 250, True, False, False
            '*************************************************************************
            xLinea = xLinea + 500
            FrmVsPrinter.Vs.DrawLine xColumna + 500, xLinea, 10000, xLinea
            xLinea = xLinea + 250
            FrmVsPrinter.Vs.DrawLine xColumna + 500, xLinea, 10000, xLinea
            xLinea = xLinea + 250
            FrmVsPrinter.Vs.DrawLine xColumna + 500, xLinea, 10000, xLinea
            xLinea = xLinea + 250
            FrmVsPrinter.Vs.DrawLine xColumna + 500, xLinea, 10000, xLinea
            
SIGUIENTE:
        Next A
        .EndDoc
    End With
    'Muestra la preimagen de la impresion
    FrmVsPrinter.WindowState = 2
    FrmVsPrinter.Show
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstOrd("id")), xCon
    End If
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstOrd
End Sub

Private Sub seleccionartodos(seleccionar As Boolean, Index As Integer)
    Dim A As Integer
    RstValores.Filter = adFilterNone
    For A = 1 To fg(Index).Rows - 1
        fg(Index).Select A, 1
        If seleccionar Then
            fg(Index).CellChecked = flexChecked
        Else
            fg(Index).CellChecked = flexUnchecked
        End If
        
        If Index = 1 Then
            RstValores.Filter = "idorddet = " & fg(0).TextMatrix(fg(0).Row, COLUMNA_ID) & " And iditem = " & fg(1).TextMatrix(A, 6)
            RstValores("activo") = seleccionar
        End If
    Next A
End Sub

Private Sub fg_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    If Col = 1 Then
        Order = 0
        If QueHace = 3 Then Exit Sub
        
        fg(Index).Select 0, 1
        If fg(Index).CellChecked = flexChecked Then
            fg(Index).CellChecked = flexUnchecked
            seleccionartodos False, Index
        Else
            fg(Index).CellChecked = flexChecked
            seleccionartodos True, Index
        End If
    End If
End Sub

Private Sub fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Index = 0 Then
        Dim xform As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        Dim xCampos(2, 4) As String
        Dim nTitulo As String
        
        If QueHace = 3 Then Exit Sub
        
        If Col = COLUMNA_ITEM Then ' Buscando Productos
            ' Si no se ha escogido el tipo se sale
            If NulosN(fg(0).TextMatrix(Row, COLUMNA_TIPO)) = 0 Then
                nTitulo = "Error al escoger Producto"
                MsgBox "No ha escogido el Tipo de Producto, seleccionelo y vuelva a intentar", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
                fg(0).Select Row, COLUMNA_TIPO
                Exit Sub
            End If
            
            xCampos(0, 0) = "Producto":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":     xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
            
            If NulosN(fg(0).TextMatrix(Row, COLUMNA_TIPO)) <> 3 Then Exit Sub
            
            ' Se escogen los Productos segun el tipo
            cSQL = "SELECT alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.tippro, alm_inventario.id, pro_receta.codrec, pro_receta.id AS idrec, mae_unidades.abrev AS unimed, alm_inventario.idunimed " _
                + vbCr + "FROM (alm_inventario LEFT JOIN pro_receta ON alm_inventario.id = pro_receta.iditem) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                + vbCr + "Where (((alm_inventario.tippro) = 3) And ((pro_receta.prirec) = 1)) " _
                + vbCr + "ORDER BY alm_inventario.descripcion;"
            
            xform.SQLCad = cSQL
    
            xform.titulo = "Buscando Producto"
            xform.FormaBusca = Principio
            xform.Criterio = ""
            xform.Ordenado = "descripcion"
            xform.CampoBusca = "descripcion"
            Set xform.Coneccion = xCon
            
            'Inicia tabla de busqueda
            Set xRs = xform.BuscarReg(xCampos)
            
            If xRs.State = 0 Then Exit Sub
            
            fg(0).TextMatrix(Row, COLUMNA_SELECCIONADO) = -1
            fg(0).TextMatrix(Row, COLUMNA_ITEM) = NulosC(xRs("descripcion"))            ' Descripcion del producto
            fg(0).TextMatrix(Row, COLUMNA_UM) = NulosC(xRs("unimed"))                   ' Descripcion de la UM
            fg(0).TextMatrix(Row, COLUMNA_RECETA) = NulosC(xRs("codrec"))               ' Codigo de la receta
            
            fg(0).TextMatrix(Row, columna_idpro) = NulosN(xRs("id"))                    ' ID del item
            fg(0).TextMatrix(Row, COLUMNA_IDREC) = NulosN(xRs("idrec"))                 ' ID de la receta
            fg(0).TextMatrix(Row, COLUMNA_IDUNIMED) = NulosN(xRs("idunimed"))           ' ID de la UM
            
            If RstValores.State = 0 Then Exit Sub
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
            limpiarRST RstValores, False
    
            cargarReceta NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREC)), NulosN(fg(0).TextMatrix(fg(0).Row, columna_cantidad))
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
            pCargarValores
            PosicionarFrm Frm4
            Frm4.Visible = True
        End If
        
        If Col = COLUMNA_RESPONSABLE Then ' Buscando Personal Responsable
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":                xCampos(1, 1) = "id":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
                    
            cSQL = "SELECT pro_emp.idemp, pro_emp.sup, pro_emp.prog, pro_emp.res, pla_empleados.nombre " _
            + vbCr + "FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
            + vbCr + "Where (((pro_empdet.idfun) = 3)) " _
            + vbCr + "GROUP BY pro_emp.idemp, pro_emp.sup, pro_emp.prog, pro_emp.res, pla_empleados.nombre " _
            + vbCr + "Having (((pla_empleados.nombre) Is Not Null)) " _
            + vbCr + "ORDER BY pla_empleados.nombre;"
                
            nTitulo = "Buscando Personal Encargado"
                    
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            fg(0).TextMatrix(fg(0).Row, COLUMNA_RESPONSABLE) = NulosC(xRs("nombre"))      ' codigo de la receta
            fg(0).TextMatrix(fg(0).Row, COLUMNA_IDRESP) = NulosN(xRs("idemp"))          ' ID de la receta
        End If
        
        If Col = COLUMNA_RECETA Then ' Buscando Recetas
            If NulosN(fg(0).TextMatrix(Row, COLUMNA_TIPO)) <> 3 Then Exit Sub
            
            xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Receta":     xCampos(1, 1) = "codrec":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
            
            cSQL = "SELECT pro_receta.codrec, pro_receta.descripcion, pro_receta.prirec, pro_receta.id " _
                + vbCr + "From pro_receta " _
                + vbCr + "Where (((pro_receta.iditem) = " & NulosN(fg(0).TextMatrix(fg(0).Row, columna_idpro)) & ")) " _
                + vbCr + "ORDER BY pro_receta.prirec;"
                
            nTitulo = "Buscando Recetas del Producto"
                    
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            fg(0).TextMatrix(fg(0).Row, COLUMNA_RECETA) = NulosC(xRs("codrec"))      ' codigo de la receta
            fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREC) = NulosN(xRs("id"))          ' ID de la receta
        End If
    End If
    
    If Index = 1 Then
        Cmd_Click 3
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Index = 0 Then
        If Agregando = True Then Exit Sub
        
        If Col = COLUMNA_TIPO Then 'Cambiar TIPO
            ' Se limpia el detalle relacionado con la fila
            If RstValores.State = 0 Then Exit Sub
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(Row, COLUMNA_ID)) & ""
            limpiarRST RstValores, False
            limpiarDatosFila
            
            If NulosN(fg(0).TextMatrix(Row, COLUMNA_TIPO)) = 3 Then
                Frm4.Visible = False
            Else
                fg(0).TextMatrix(Row, COLUMNA_ITEM) = "VER DETALLE"
                fg(0).TextMatrix(Row, columna_idpro) = 0
                fg(1).Rows = 1
                PosicionarFrm Frm4
                Frm4.Visible = True
            End If
        End If
        
        If Col = columna_cantidad Then 'Cambiar cantidad
            fg(0).TextMatrix(Row, Col) = Format(fg(0).TextMatrix(Row, Col), "0.00")
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
            limpiarRST RstValores, False
            cargarReceta NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREC)), NulosN(fg(0).TextMatrix(fg(0).Row, columna_cantidad))
            pCargarValores
        End If
    End If
    
    If Index = 1 Then
        If Agregando = True Then Exit Sub
        
        If Col = 4 Then ' Cantidad
            fg(Index).TextMatrix(Row, Col) = Format(fg(Index).TextMatrix(Row, Col), "0.00")
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & _
                                " And iditem = " & NulosN(fg(1).TextMatrix(fg(1).Row, 6)) & ""
                                
            If RstValores.RecordCount = 0 Then Exit Sub
            RstValores("cantidad") = NulosN(fg(Index).TextMatrix(Row, Col))
            RstValores.Update
        End If
    End If
End Sub

Private Sub Fg_Click(Index As Integer)
    If Index = 0 Then
        If QueHace <> 3 Then Exit Sub
        
        If fg(Index).Col <> COLUMNA_SELECCIONADO Then Exit Sub
        
        If fg(0).TextMatrix(fg(0).Row, COLUMNA_SELECCIONADO) = 0 Then
            fg(0).TextMatrix(fg(0).Row, COLUMNA_SELECCIONADO) = -1
        Else
            fg(0).TextMatrix(fg(0).Row, COLUMNA_SELECCIONADO) = 0
        End If
    End If
End Sub

Private Sub fg_ComboCloseUp(Index As Integer, ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    If Index = 0 Then
        Frm4.Visible = False
    End If
End Sub

Private Sub Fg_DblClick(Index As Integer)
    If Index = 0 Then
        verDetalle
    End If
End Sub

Private Sub Fg_EnterCell(Index As Integer)
    If QueHace = 3 Then
        fg(Index).Editable = flexEDNone
        fg(Index).SelectionMode = flexSelectionByRow
        Exit Sub
    End If
    fg(Index).Editable = flexEDKbdMouse
    fg(Index).SelectionMode = flexSelectionFree
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Index = 0 Then
        Select Case Col
            Case COLUMNA_TIPO, COLUMNA_ITEM, COLUMNA_RESPONSABLE, COLUMNA_UM, COLUMNA_RECETA, COLUMNA_NUMORDEN
                KeyAscii = 0
            
            Case columna_cantidad
                ' Si no es Producto entonces no se edita la cantidad y se sale del sub
                If NulosN(fg(0).TextMatrix(Row, COLUMNA_TIPO)) <> 3 Then KeyAscii = 0: Exit Sub
                ' Si no es un numero no se edita
                If validar_numero(KeyAscii) = False Then KeyAscii = 0
        End Select
    End If
End Sub

Private Sub fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        If QueHace = 3 Then Exit Sub
        If Button = 2 Then
            PopupMenu Menu1
        End If
    End If
End Sub

Private Sub fg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    If Index = 1 Then
        If QueHace = 3 Then Exit Sub
        
        'If fg(Index).Col <> 1 Then Exit Sub
        
        If fg(Index).Row > 0 Then
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & _
                                " And iditem = " & NulosN(fg(1).TextMatrix(fg(1).Row, 6)) & ""
                                
            If RstValores.RecordCount = 0 Then Exit Sub
            RstValores("activo") = NulosN(fg(Index).TextMatrix(fg(Index).Row, 1))
            RstValores.Update
        End If
        
        ' Se verifica si se hizo clic en la cabecera de seleccion
        If (X < 350 And X > 180) And (Y > 15 And Y < 180) Then
            fg(Index).Select 0, 1
            If fg(Index).CellChecked = flexChecked Then
                seleccionartodos True, Index
            Else
                seleccionartodos False, Index
            End If
        End If
    End If
End Sub

Private Sub fg_RowColChange(Index As Integer)
    If Frm4.Visible = False Then Exit Sub
    
    If Index = 0 Then
        RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
        
        If RstValores.RecordCount = 0 Then
            cargarReceta NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREC)), NulosN(fg(0).TextMatrix(fg(0).Row, columna_cantidad))
        End If
        
        pCargarValores
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    Agregando = False
    iniciarCampos
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDOS E CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
            
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        pCargarGrid
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then '--F3 Nuevo
        If QueHace <> 3 Then Exit Sub
        Nuevo
    End If
    
    If KeyCode = 115 Then '--F4 Modificar
        If QueHace <> 3 Then Exit Sub
        Modificar
    End If
    
    If KeyCode = 113 Then '--F2 Grabar
        If QueHace = 3 Then Exit Sub
        If Grabar = True Then
            QueHace = 3
            Set RstOrd = Nothing
        End If
    End If
    
    If KeyCode = 116 Then '--F5 actualizar
        Me.Refresh
    End If
    
    If KeyCode = 117 Then '--F6 '--cancelar
        If QueHace = 3 Then Exit Sub
        Cancelar
    End If
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 8100 Then Me.Height = 8100

    ' Se dimensiona el Contenido
    TabOne1.Width = Me.Width - 150
    TabOne1.Height = Me.Height - 900
    
    Label4(0).Width = Me.Width - 100
    LblMes.Left = TabOne1.Width - 1200
    Dg1.Width = TabOne1.Width - 100
    Dg1.Height = TabOne1.Height - 800
    
    ' Se dimensiona el Detalle
    Label5.Width = Me.Width - 100
    Frame6.Width = TabOne1.Width - 100
    LblNomProg.Width = Frame6.Width - 2500
    
    fg(0).Width = TabOne1.Width - 100
    fg(0).Height = TabOne1.Height - 2150
    
    Frame3.Top = TabOne1.Height - 950
    Frame3.Width = TabOne1.Width - 100
    
    PosicionarFrm Frm4
End Sub

Private Sub limpiarDatosFila()
    fg(0).TextMatrix(fg(0).Row, COLUMNA_ITEM) = ""
    fg(0).TextMatrix(fg(0).Row, columna_idpro) = ""
    fg(0).TextMatrix(fg(0).Row, columna_cantidad) = ""
    fg(0).TextMatrix(fg(0).Row, COLUMNA_UM) = ""
    fg(0).TextMatrix(fg(0).Row, COLUMNA_IDUNIMED) = ""
    fg(0).TextMatrix(fg(0).Row, COLUMNA_RECETA) = ""
    fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREC) = ""
    fg(0).TextMatrix(fg(0).Row, COLUMNA_LOTE) = ""
End Sub

Private Sub pCargarDatosRstTemp(idCodigo)
    Dim RstTmp As New ADODB.Recordset
    Set RstTmp = Nothing
    
    ' Definir la estructura de recordset
    cSQL = "SELECT pro_ordenproddetins.idorddet, pro_ordenproddetins.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_ordenproddetins.cantidad, pro_ordenproddetins.activo " _
        + vbCr + "FROM (pro_ordenproddetins LEFT JOIN alm_inventario ON pro_ordenproddetins.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "Where (((pro_ordenproddetins.idord) = " & idCodigo & ")) " _
        + vbCr + "GROUP BY pro_ordenproddetins.idorddet, pro_ordenproddetins.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_ordenproddetins.cantidad, pro_ordenproddetins.activo " _
        + vbCr + "ORDER BY pro_ordenproddetins.idorddet;"
            
    RST_Busq RstTmp, cSQL, xCon
    
    If RstValores.State = 0 Then DEFINIR_RST_TMP RstValores, RstTmp
    CARGAR_RST_TMP RstValores, RstTmp
    
    Set RstTmp = Nothing
End Sub

Private Sub iniciarCampos()
    COLUMNA_SELECCIONADO = 1
    COLUMNA_TIPO = 2
    COLUMNA_ITEM = 3
    COLUMNA_RESPONSABLE = 4
    COLUMNA_UM = 5
    columna_cantidad = 6
    COLUMNA_RECETA = 7
    COLUMNA_LOTE = 8
    COLUMNA_NUMORDEN = 9
    columna_idpro = 10
    COLUMNA_IDREC = 11
    COLUMNA_IDUNIMED = 12
    COLUMNA_IDRESP = 13
    COLUMNA_ID = 14
    COLUMNA_IDCRDET = 15
    
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).ExplorerBar = flexExSortShowAndMove
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).BackColorSel = &H80&
    fg(0).ForeColorSel = &H80000005
    
    fg(1).AllowUserResizing = flexResizeColumns
    fg(1).ExplorerBar = flexExSortShowAndMove
    fg(1).SelectionMode = flexSelectionByRow
    fg(1).BackColorSel = &H80&
    fg(1).ForeColorSel = &H80000005
    
    fg(0).Rows = 2
    fg(0).ColWidth(columna_idpro) = 0
    fg(0).ColWidth(COLUMNA_IDREC) = 0
    fg(0).ColWidth(COLUMNA_IDUNIMED) = 0
    fg(0).ColWidth(COLUMNA_IDRESP) = 0
    fg(0).ColWidth(COLUMNA_ID) = 0
    fg(0).ColWidth(COLUMNA_IDCRDET) = 0
    
    GRID_COMBOLIST fg(0), COLUMNA_TIPO
    GRID_COMBOLIST fg(0), COLUMNA_ITEM
    GRID_COMBOLIST fg(0), COLUMNA_RESPONSABLE
    GRID_COMBOLIST fg(0), COLUMNA_RECETA
    
    fg(1).ColWidth(5) = 0
    fg(1).ColWidth(6) = 0
    GRID_COMBOLIST fg(1), 2
          
    ' Se agregan los tipos de Items segun BD
    Dim RstAux As New ADODB.Recordset
    Dim CAMPOS As String
    Dim A As Integer
    
    Set RstAux = Nothing
    CAMPOS = ""
    
    cSQL = "SELECT mae_tipoproducto.id, mae_tipoproducto.descripcion, mae_tipoproducto.prefijo " _
        + vbCr + "FROM mae_tipoproducto;"
        
    RST_Busq RstAux, cSQL, xCon
    
    If RstAux.State = 0 Then Exit Sub
    If RstAux.RecordCount = 0 Then Exit Sub
    
    RstAux.MoveFirst
    For A = 1 To RstAux.RecordCount - 1
         CAMPOS = CAMPOS & "#" & A & ";" & RstAux("descripcion") & "|"
         RstAux.MoveNext
    Next A
    CAMPOS = CAMPOS & "#" & A & ";" & RstAux("descripcion")
    fg(0).ColComboList(2) = CAMPOS
    
    ' Se agrega el mes Activo
    mMesActivo = xMes
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    
    ' Se pone el cuadro de seleccion en la cabecera del flexgrid
    fg(0).Select 0, 1
    fg(0).CellChecked = flexChecked
    
    fg(1).Select 0, 1
    fg(1).CellChecked = flexUnchecked
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub pCargarGrid()
    Dim cSQL  As String
    Dim Rpta As Integer
    
    TDB_FiltroLimpiar Dg1
    
    cSQL = "SELECT pro_ordenprod.id, pro_ordenprod.idsup, pla_empleados_1.nombre AS nomsup, pro_ordenproddet.idresponsable AS idresp, pro_ordenproddet.tipo AS idtippro, pro_ordenprod.fchemi AS fchpro, mae_tipoproducto.descripcion AS destippro, pla_empleados.nombre AS nomresp, IIf([pro_ordenprod].[idcr] <> 0,'CRONOGRAMA','MANUAL') AS [proc], alm_inventario.descripcion, pro_ordenproddet.numdoc AS numsol " _
            + vbCr + "FROM (((((pro_ordenprod LEFT JOIN pro_ordenproddet ON pro_ordenprod.id = pro_ordenproddet.idord) LEFT JOIN alm_inventario ON pro_ordenproddet.iditem = alm_inventario.id) LEFT JOIN mae_tipoproducto ON pro_ordenproddet.tipo = mae_tipoproducto.id) LEFT JOIN pla_empleados ON pro_ordenproddet.idresponsable = pla_empleados.id) LEFT JOIN pro_emp ON pro_ordenprod.idsup = pro_emp.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_emp.idemp = pla_empleados_1.id " _
            + vbCr + "GROUP BY pro_ordenprod.id, pro_ordenprod.idsup, pla_empleados_1.nombre, pro_ordenproddet.idresponsable, pro_ordenproddet.tipo, pro_ordenprod.fchemi, mae_tipoproducto.descripcion, pla_empleados.nombre, IIf([pro_ordenprod].[idcr] <> 0,'CRONOGRAMA','MANUAL'), alm_inventario.descripcion, pro_ordenproddet.numdoc " _
            + vbCr + "Having ((Month(pro_ordenprod.fchemi) = " & mMesActivo & ") And (Year(pro_ordenprod.fchemi) = " & Val(AnoTra) & ")) " _
            + vbCr + "ORDER BY pro_ordenprod.fchemi DESC , pro_ordenproddet.numdoc DESC;"
    
    ' cargando datos
    Me.MousePointer = vbHourglass
    
    RST_Busq RstOrd, cSQL, xCon
    Set Dg1.DataSource = RstOrd
    
    Me.MousePointer = vbDefault
    
    If RstOrd.State = 0 Then Exit Sub
End Sub

Sub procesarCronograma()
    Dim xRs As New ADODB.Recordset
    Dim Rst As New ADODB.Recordset
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId As Double
    Dim numSol As Double
    Dim nSQL As String
    
    cSQL = "SELECT * FROM pro_ordenproddet WHERE (idord = " & RstOrd("id") & " AND fchprog = CDate('" & RstOrd("fchpro") & "'));"

    RST_Busq xRs, cSQL, xCon
    
    If xRs.RecordCount = 0 Then
        If NulosN(RstOrd("idtippro")) = 3 Then
            ' SI SE ESTAN PROCESANDO PRODUCTOS
            cSQL = "SELECT pro_cronogramadet.id, pro_cronogramadet.fchpro, pro_cronogramadet.iditem AS idpro, '' AS nommatpri, alm_inventario.descripcion AS nompro, pro_cronogramadet.cantidad, pro_receta.codrec, pro_receta.id AS idrec, pro_receta.idunimed " _
                + vbCr + "FROM (pro_cronogramadet LEFT JOIN pro_receta ON pro_cronogramadet.iditem = pro_receta.iditem) LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id" _
                + vbCr + "Where (((pro_cronogramadet.idcr) = " & RstOrd("id") & ") And ((pro_cronogramadet.fchpro) = CDate('" & RstOrd("fchpro") & "')) And ((pro_receta.prirec) <= 1)) " _
                + vbCr + "ORDER BY pro_cronogramadet.fchpro;"
        Else
            ' SI SE ESTA PROCESANDO MATERIA PRIMA
            cSQL = "SELECT pro_cronogramadetprod.id, pro_cronogramadetprod.fchpro, pro_cronogramadetprod.iditem, pro_cronogramadetprod.idpro, pro_cronogramadetprod.cantidad, pro_receta.codrec, pro_receta.id AS idrec, [idrec] AS id, pro_receta.idunimed " _
                + vbCr + "FROM pro_cronogramadetprod LEFT JOIN pro_receta ON pro_cronogramadetprod.idpro = pro_receta.iditem " _
                + vbCr + "Where (((pro_cronogramadetprod.idcr) = " & RstOrd("id") & ") AND ((pro_cronogramadetprod.fchpro)= CDate('" & RstOrd("fchpro") & "')) AND ((pro_cronogramadetprod.cantidad)<>0) AND ((pro_receta.prirec)<=1)) " _
                + vbCr + "ORDER BY pro_cronogramadetprod.fchpro"
        End If
        
        RST_Busq Rst, cSQL, xCon
    
        On Error GoTo LaCague
        
        xCon.BeginTrans
        
        xId = RstOrd("id")
        
        RST_Busq RstDet, "SELECT TOP 1 * FROM pro_ordenproddet", xCon
        
        mIdRegistro = xId
        numSol = HallaCodigoTabla("pro_ordenproddet", xCon, "numdoc")
        ' Detalle
        Rst.MoveFirst
        While Not Rst.EOF
            RstDet.AddNew
            RstDet("idord") = xId
            RstDet("idcrdet") = NulosN(Rst("id"))
            RstDet("iditem") = NulosN(Rst("idpro"))
            RstDet("fchprog") = NulosC(Rst("fchpro"))
            RstDet("idrec") = NulosN(Rst("idrec"))
            RstDet("idunimed") = NulosN(Rst("idunimed"))
            RstDet("cantidad") = NulosN(Rst("cantidad"))
            
            RstDet("numser") = "0001"
            RstDet("lote") = Format(xId, "000000")
            RstDet("numdoc") = Format(numSol, "000000")
            RstDet("proc") = 1
            RstDet("obs") = ""
            RstDet.Update
            
            Rst.MoveNext
            numSol = numSol + 1
        Wend
        'Grabamos el movimiento en la tabla var_edicion
        GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
        xCon.CommitTrans
        MsgBox "Se Proceso la Operacion con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstDet = Nothing
    End If
    Exit Sub
LaCague:
        xCon.RollbackTrans
        Set RstDet = Nothing
        MsgBox "No se pudo procesar el registro por el siguiente motivo :" + Trim(Err.Description)
        Exit Sub
End Sub

Private Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    Dim Rpta As Integer
    
    Blanquea
    
    If RstOrd.RecordCount = 0 Then Exit Sub
    If RstOrd.EOF = True Then Exit Sub
     
    Set RstDet = Nothing
    Agregando = True
    
    cSQL = "SELECT pro_ordenproddet.id, pro_ordenproddet.idcrdet, alm_inventario.descripcion AS prod, mae_unidades.abrev AS unid, pro_ordenproddet.cantidad, pro_receta.codrec, pro_ordenproddet.numser, pro_ordenproddet.lote, pro_ordenproddet.numdoc, pro_ordenproddet.iditem, pro_ordenproddet.idrec, pro_ordenproddet.idunimed, mae_tipoproducto.descripcion AS desctipo, pro_ordenproddet.tipo AS idtipo, pla_empleados.nombre, pro_ordenproddet.idresponsable " _
        + vbCr + "FROM ((((pro_ordenproddet LEFT JOIN alm_inventario ON pro_ordenproddet.iditem = alm_inventario.id) LEFT JOIN pro_receta ON pro_ordenproddet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_ordenproddet.idunimed = mae_unidades.id) LEFT JOIN mae_tipoproducto ON pro_ordenproddet.tipo = mae_tipoproducto.id) LEFT JOIN pla_empleados ON pro_ordenproddet.idresponsable = pla_empleados.id " _
        + vbCr + "WHERE (((pro_ordenproddet.idord)=" & NulosN(RstOrd("id")) & ") AND ((pro_ordenproddet.fchprog)=CDate('" & RstOrd("fchpro") & "')));"

    RST_Busq RstDet, cSQL, xCon
    
    If IsDate(RstOrd("fchpro")) = True Then TxtFchPro.valor = CDate(RstOrd("fchpro"))
    
    TxtIdProg.Text = NulosN(RstOrd("idsup"))
    LblNomProg.Caption = NulosC(RstOrd("nomsup"))
    
    'Se llena la referencia a la procedencia de la solicitud
    If RstOrd("proc") = "MANUAL" Then LblProc = 2 Else LblProc = 1
            
    fg(0).Rows = fg(0).FixedRows
        
    If RstDet.State = 0 Then Exit Sub
    Agregando = True
    If RstDet.RecordCount = 0 Then Exit Sub
    
    RstDet.MoveFirst
    While Not RstDet.EOF
        fg(0).Rows = fg(0).Rows + 1
        With fg(0)
            .TextMatrix(.Rows - 1, COLUMNA_ID) = NulosN(RstDet("id"))
            .TextMatrix(.Rows - 1, COLUMNA_IDCRDET) = NulosN(RstDet("idcrdet"))
            .TextMatrix(.Rows - 1, COLUMNA_SELECCIONADO) = -1                                        ' Se pone como activo el producto
            .TextMatrix(.Rows - 1, COLUMNA_TIPO) = NulosN(RstDet("idtipo"))
            .TextMatrix(.Rows - 1, COLUMNA_RESPONSABLE) = NulosC(RstDet("nombre"))
            .TextMatrix(.Rows - 1, COLUMNA_UM) = NulosC(RstDet("unid"))                             ' descripcion de la unidad de medida
            
            If NulosN(RstDet("idtipo")) = 3 Then
                .TextMatrix(.Rows - 1, columna_cantidad) = Format(NulosN(RstDet("cantidad")), "0.00")   ' cantidad de producto
            Else
                If NulosN(RstDet("cantidad")) = 0 Then
                    .TextMatrix(.Rows - 1, columna_cantidad) = ""   ' cantidad de producto
                End If
            End If
            
            .TextMatrix(.Rows - 1, COLUMNA_ITEM) = NulosC(RstDet("prod"))                           ' Descripcion del producto
            .TextMatrix(.Rows - 1, COLUMNA_RECETA) = NulosC(RstDet("codrec"))                       ' codigo de la receta
            .TextMatrix(.Rows - 1, COLUMNA_LOTE) = NulosC(RstDet("lote"))                           ' numero de produccion
            .TextMatrix(.Rows - 1, COLUMNA_NUMORDEN) = Format(NulosN(RstDet("numdoc")), "000000")   ' numero de documento
            .TextMatrix(.Rows - 1, columna_idpro) = NulosC(RstDet("iditem"))                        ' ID de producto
            .TextMatrix(.Rows - 1, COLUMNA_IDREC) = NulosC(RstDet("idrec"))                         ' ID de receta
            .TextMatrix(.Rows - 1, COLUMNA_IDUNIMED) = NulosC(RstDet("idunimed"))                   ' ID de unidad de medida
            
            .TextMatrix(.Rows - 1, COLUMNA_IDRESP) = NulosN(RstDet("idresponsable"))
        End With
        RstDet.MoveNext
    Wend
    
    fg(0).Row = 1
    Agregando = False
    
    pCargarDatosRstTemp NulosN(RstOrd("id"))
    
    Set RstDet = Nothing
    Agregando = False
End Sub

Sub Cancelar()
    Bloquea
    Label5.Caption = "Detalle de Solicitud de Materiales"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    fg(0).SelectionMode = flexSelectionByRow
     
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
    limpiarRST RstValores
End Sub

Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    Bloquea
    Blanquea
    agregados = 0
    fg(0).Rows = 1
    'fg(0).Rows = fg(0).Rows + 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    
    Label5.Caption = "Agregando Solicitud de Materiales"
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    
    numSolMax = HallaCodigoTabla("pro_ordenproddet", xCon, "numdoc")
    NUMERO_CORRELATIVO = 666
    
    If RstValores.State = 0 Then pCargarDatosRstTemp 0
End Sub

Sub Bloquea()
    TxtIdProg.Locked = Not TxtIdProg.Locked
    TxtFchPro.Locked = Not TxtFchPro.Locked
    habilitar cmd, Not TxtFchPro.Locked
End Sub

Sub Blanquea()
    TxtIdProg.Text = ""
    LblNomProg.Caption = ""
    TxtFchPro.valor = ""
End Sub

Function Grabar() As Boolean
    Dim A As Integer
    Dim B As Integer
    Dim xTot As Long
    Dim procSol As Double
    
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If TxtIdProg.Text = "" Then
        MsgBox "No ha especificado un Supervisor para la Solicitud", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdProg.SetFocus
        Exit Function
    End If
    
    If TxtFchPro.valor = "" Then
        MsgBox "No ha especificado fecha de SOlicitud", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchPro.SetFocus
        Exit Function
    End If

    If fg(0).Rows = 1 Then
        MsgBox "No ha especificado items para la Solicitud de Materiales", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).SetFocus
        Exit Function
    End If
    
    ' verificando el detalle
    Rst.Filter = adFilterNone ' quitando el filtro al rst para hacer las evaluaciones
    
    For A = 1 To fg(0).Rows - 2
        If NulosN(fg(0).TextMatrix(A, COLUMNA_TIPO)) = 3 Then
            If NulosN(fg(0).TextMatrix(A, columna_cantidad)) = 0 Then
                MsgBox "No se le ha asignado una cantidad para el item : " & Chr(13) _
                    & fg(0).TextMatrix(A, COLUMNA_ITEM) _
                    & ", asignele una cantidad", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                fg(0).Select A, columna_cantidad
                fg(0).SetFocus
                Exit Function
            End If
        End If
    Next A
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDetIns As New ADODB.Recordset
    Dim xId As Double
    Dim nSQL As String
    
    'On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        ' Obetenemos el Id del registro
        xId = HallaCodigoTabla("pro_ordenprod", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pro_ordenprod", xCon
        RstCab.AddNew
        RstCab("id") = xId
        procSol = 2 ' La procedencia de la solicitud se pone como manual
    Else
        ' SI SE ESTA MOFIGICANDO UN REGISTRO OBTENEMOS EL ID DEL REGISTRO ACTUAL
        xId = RstOrd("id")
        RST_Busq RstCab, "SELECT * FROM pro_ordenprod WHERE id = " & xId & "", xCon
        ' Eliminamos el detalle
        xCon.Execute "DELETE * FROM pro_ordenproddetins WHERE idord  = " & xId & ""
        xCon.Execute "DELETE * FROM pro_ordenproddet WHERE idord  = " & xId & ""
        procSol = LblProc
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM pro_ordenproddet", xCon
    RST_Busq RstDetIns, "SELECT TOP 1 * FROM pro_ordenproddetins", xCon
    
    mIdRegistro = xId
    
    If procSol = 2 Then
        RstCab("idsup") = NulosN(TxtIdProg.Text)
        RstCab("fchemi") = CDate(TxtFchPro.valor)
        RstCab.Update
    End If
    
    ' Detalle
    Dim identificador As Integer
    identificador = HallaCodigoTabla("pro_ordenproddet", xCon, "id")
    
    For A = 1 To fg(0).Rows - 1
        RstDet.AddNew
        RstDet("id") = identificador
        RstDet("idord") = xId
        RstDet("idcrdet") = NulosN(fg(0).TextMatrix(A, COLUMNA_IDCRDET))
        'Para evitar grabar campos vacios
        If fg(0).TextMatrix(A, columna_idpro) = "" Then Exit For
        RstDet("iditem") = NulosN(fg(0).TextMatrix(A, columna_idpro))
        RstDet("idrec") = NulosN(fg(0).TextMatrix(A, COLUMNA_IDREC))
        RstDet("idunimed") = NulosN(fg(0).TextMatrix(A, COLUMNA_IDUNIMED))
        RstDet("cantidad") = NulosN(fg(0).TextMatrix(A, columna_cantidad))
        RstDet("numser") = "0001"
        RstDet("numdoc") = NulosC(fg(0).TextMatrix(A, COLUMNA_NUMORDEN))
        RstDet("lote") = NulosC(fg(0).TextMatrix(A, COLUMNA_LOTE))
        RstDet("fchprog") = CDate(TxtFchPro.valor)
        RstDet("proc") = procSol
        RstDet("tipo") = NulosN(fg(0).TextMatrix(A, COLUMNA_TIPO))
        RstDet("idresponsable") = NulosN(fg(0).TextMatrix(A, COLUMNA_IDRESP))
        RstDet("obs") = ""
        RstDet.Update
        
        ' Detalle de Insumos
        RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(A, COLUMNA_ID))
        If RstValores.RecordCount <> 0 Then RstValores.MoveFirst
        For B = 1 To RstValores.RecordCount
            RstDetIns.AddNew
            RstDetIns("idord") = xId
            RstDetIns("idorddet") = identificador
            RstDetIns("activo") = NulosN(RstValores("activo"))
            RstDetIns("iditem") = NulosN(RstValores("iditem"))
            RstDetIns("cantidad") = NulosN(RstValores("cantidad"))
            RstDetIns.Update
            RstValores.MoveNext
            If RstValores.EOF Then Exit For
        Next B
        identificador = identificador + 1
    Next A
            
    'Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
   
    xCon.CommitTrans
    MsgBox "La Solicitud de Materiales se registró con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
    Grabar = True
    Exit Function
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
End Function

Sub Modificar()
    If RstOrd.RecordCount = 0 Then
        MsgBox "No hay Registros para Modificar", vbInformation, Me.Caption
        Exit Sub
    End If
   
    QueHace = 2
    xHorIni = Time
    Bloquea
    agregados = 0
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    mCorrelativo = 1
    limpiarRST RstValores
    MuestraSegundoTab
    
    Label5.Caption = "Modificando Solicitud de Materiales"
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    
    xHorIni = Time
    CmdBusSup.SetFocus
    numSolMax = HallaCodigoTabla("pro_ordenproddet", xCon, "numdoc")
    NUMERO_CORRELATIVO = 666
    Frm4.Visible = False
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    If RstOrd.RecordCount = 0 Then
        MsgBox "No hay documentos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    TabOne1.CurrTab = 0
    Rpta = MsgBox("¿ Esta seguro de eliminar el Registro seleccionado ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_ordenproddetins WHERE idord = " & RstOrd("id") & ""
        xCon.Execute "DELETE * FROM pro_ordenproddet WHERE (idord = " & RstOrd("id") & " AND fchprog=CDate('" & RstOrd("fchpro") & "'))"
        xCon.Execute "DELETE * FROM pro_ordenprod WHERE id = " & RstOrd("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstOrd("id") & " AND idform = " & IdMenuActivo
        
        MsgBox "El registro se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstOrd.Requery
        Dg1.Refresh
    End If
End Sub

Sub ExportarExcel()
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    Dim xCad As String
    Dim objExcel As Object
    
    Dim oExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    
    Set oExcel = New Excel.Application
    Set oWBook = oExcel.Workbooks.Add
    Screen.MousePointer = vbHourglass

    oExcel.WindowState = 2
    For A = 1 To fg(0).Rows - 1
        oWBook.ActiveSheet.Name = fg(0).TextMatrix(A, 10)
        With oWBook.ActiveSheet
            'Se llena cabecera
            .Cells(1, 2) = "SOLICITUD DE MATERIALES Nº:   " + "0001" & "-" & fg(0).TextMatrix(A, 10)
            .Range("B1", "H1").Merge
            .Cells(1, 2).HorizontalAlignment = xlHAlignCenterAcrossSelection
            .Cells(1, 2).Font.Bold = True
            .Cells(1, 2).Rows(1).Font.Size = 12
            
            .Cells(3, 2) = "Producción    Nº " + fg(0).TextMatrix(A, 6)
            .Cells(3, 2).Font.Bold = True
            .Cells(4, 2) = "Fch. Prod. :   " + TxtFchPro.valor
            .Cells(4, 2).Font.Bold = True
            .Cells(5, 2) = "Producto :   " + fg(0).TextMatrix(A, 2)
            .Cells(5, 2).Font.Bold = True
            .Cells(6, 2) = "Receta :   " + fg(0).TextMatrix(A, 5)
            .Cells(6, 2).Font.Bold = True
            
            .Cells(7, 2) = "Cantidad  :   " + fg(0).TextMatrix(A, 4)
            .Cells(7, 2).Font.Bold = True
            
            .Cells(9, 2) = "Item"
            .Cells(9, 2).Font.Bold = True
            .Cells(9, 3) = "INSUMO / PRODUCTO / MP"
            .Cells(9, 3).Font.Bold = True
            .Cells(9, 4) = "Uni. Med."
            .Cells(9, 4).Font.Bold = True
            .Cells(9, 5) = "Cantidad Teorica"
            .Cells(9, 5).Font.Bold = True
            .Cells(9, 6) = "Cantidad Real"
            .Cells(9, 6).Font.Bold = True
            .Cells(9, 7) = "Adicional"
            .Cells(9, 7).Font.Bold = True
            .Cells(9, 8) = "Devolucion"
            .Cells(9, 8).Font.Bold = True
            
            Dim Rst As New ADODB.Recordset
            
            RST_Busq Rst, "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]*" & NulosN(fg(0).TextMatrix(A, 3)) & " AS canreq " _
                + vbCr + "FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
                + vbCr + "WHERE (((pro_recetains.idrec)=" & NulosN(fg(0).TextMatrix(A, 8)) & "))", xCon
        
            If Rst.RecordCount <> 0 Then
                Dim xFila As Integer
                xFila = 10
                For B = 1 To Rst.RecordCount
                    .Cells(xFila, 2) = Format(B, "00")
                    .Cells(xFila, 3) = Rst("descripcion")
                    .Cells(xFila, 4) = Rst("abrev")
                    .Cells(xFila, 5) = Format(Rst("canreq"), "0.000000")
    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                    xFila = xFila + 1
                Next B
            End If
            .Cells(xFila + 5, 5) = "VºBº Ger. Prod. "
            .Cells(xFila + 5, 5).Font.Bold = True
            .Cells(xFila + 5, 7) = "Entregado Por "
            .Cells(xFila + 5, 7).Font.Bold = True
        End With
        If A < fg(0).Rows - 1 Then oWBook.Sheets.Add
    Next A
    
    oExcel.Visible = True
    
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Reporte de Pedidos"
    oExcel.WindowState = 1
    
    Set oExcel = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub limpiarRST(Rst As ADODB.Recordset, Optional TODO As Boolean = True)
    With Rst
        If .State <> 0 Then
            If TODO Then .Filter = adFilterNone
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
            End If
        End If
    End With
End Sub

Private Sub Menu1_1_Click() ' AGREGAR
    Cmd_Click 0
End Sub

Private Sub Menu1_2_Click() ' ELIMINAR
    Cmd_Click 1
End Sub

Private Sub Menu1_3_Click() ' VER DETALLE
    verDetalle
End Sub

Private Sub PbCerrar_Click(Index As Integer)
    If Index = 1 Then
        Frm4.Visible = False
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then
            If RstOrd.RecordCount = 0 Then
                MsgBox "No existe información para visualizar", vbInformation, Me.Caption
                Blanquea
                fg(0).Rows = 1
                Exit Sub
            Else
                MuestraSegundoTab
            End If
        End If
    Else
        limpiarRST RstValores
        Frm4.Visible = False
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then
        If RstOrd.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then
        If RstOrd.RecordCount = 0 Then
            MsgBox "No se han registrados Pedidos para realizar esta opción", vbInformation, Me.Caption
            Exit Sub
        End If
        Eliminar
    End If
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstOrd.Requery
            Dg1.Refresh
            If RstOrd.RecordCount <> 0 Then
                RstOrd.MoveFirst
                RstOrd.Find "numsol=" & Format(mIdRegistro, "000000")
                If RstOrd.EOF = True Then RstOrd.MoveFirst
            End If
        End If
    End If
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        If TabOne1.CurrTab = 0 Then RstOrd.Filter = "": TDB_FiltroLimpiar Dg1
    End If
    
    If Button.Index = 12 Then CambiarMes
    
    If Button.Index = 14 Then ExportarExcel
    If Button.Index = 15 Then
        If TabOne1.CurrTab = 0 Then Exit Sub
        ImprimirSolicitud
    End If
    If Button.Index = 17 Then Unload Me
End Sub

Private Sub CambiarMes()
    TabOne1.CurrTab = 0
    mMesActivo = SeleccionaMes(xCon)
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    pCargarGrid
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 15 Then
        If ButtonMenu.Index = 1 Then
            If TabOne1.CurrTab = 0 Then Exit Sub
            ImprimirSolicitud
        End If
        If ButtonMenu.Index = 2 Then
            If TabOne1.CurrTab = 0 Then Exit Sub
            ImprimirLinea
        End If
    End If
End Sub

Private Sub TxtIdProg_Validate(Cancel As Boolean)
    'Dim cSQL As String
    Dim xRs1 As New ADODB.Recordset
    
    If NulosC(TxtIdProg.Text) = "" Then
        LblNomProg.Caption = ""
        Exit Sub
    End If
    
    cSQL = "SELECT pro_emp.*, pla_empleados.nombre " _
        + vbCr + "FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
        + vbCr + "Where (((pro_empdet.idfun) = 2) AND ((pro_emp.id) =" & TxtIdProg.Text & ")) " _
        + vbCr + "ORDER BY pla_empleados.nombre;"
    
    RST_Busq xRs1, cSQL, xCon
    If xRs1.RecordCount <> 0 Then
        LblNomProg.Caption = xRs1("nombre")
        TxtFchPro.valor = Date
    Else
        TxtIdProg.Text = ""
        LblNomProg = ""
        TxtFchPro.valor = ""
    End If
    Set xRs1 = Nothing
End Sub

'Metodos para arrastrar el Frame
''''''''''''''''''''''''''''''''
Private Sub Frm4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frm4.ZOrder 0
End Sub

Private Sub Frm4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frm4
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub
