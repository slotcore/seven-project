VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmGenerarOrdProd2_1 
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
   WindowState     =   2  'Maximized
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
      Left            =   5220
      TabIndex        =   20
      Top             =   3630
      Visible         =   0   'False
      Width           =   6500
      Begin VB.Frame Frame8 
         Height          =   495
         Left            =   60
         TabIndex        =   23
         Top             =   3000
         Width           =   6375
         Begin VB.CommandButton Cmd 
            Caption         =   "Elimi&nar"
            Enabled         =   0   'False
            Height          =   330
            Index           =   5
            Left            =   2400
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Personal"
            Top             =   135
            Width           =   1065
         End
         Begin VB.Label Lblidorddet 
            AutoSize        =   -1  'True
            Caption         =   "Lblidorddet"
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
            Height          =   195
            Left            =   4860
            TabIndex        =   29
            Top             =   210
            Visible         =   0   'False
            Width           =   960
         End
      End
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   6210
         Picture         =   "FrmGenerarOrdProd2_1.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   21
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   2640
         Index           =   1
         Left            =   60
         TabIndex        =   22
         Top             =   390
         Width           =   6345
         _cx             =   11192
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmGenerarOrdProd2_1.frx":02EC
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
         TabIndex        =   28
         Top             =   45
         Width           =   1635
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   7
         X1              =   0
         X2              =   6470
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
         X1              =   6470
         X2              =   6470
         Y1              =   0
         Y2              =   3570
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   45
         Top             =   30
         Width           =   6390
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
            Picture         =   "FrmGenerarOrdProd2_1.frx":03E3
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":0927
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":0CB9
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":0E3D
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":1291
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":13A9
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":18ED
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":1E31
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":1F45
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":2059
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":24AD
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":2619
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGenerarOrdProd2_1.frx":2B61
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
      CurrTab         =   0
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
         Left            =   45
         TabIndex        =   4
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6135
            Left            =   30
            TabIndex        =   30
            Top             =   480
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   10821
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "iddet"
            Columns(0).DataField=   "iddet"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Id"
            Columns(1).DataField=   "id"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fecha"
            Columns(2).DataField=   "fchpro"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nº Orden"
            Columns(3).DataField=   "numsol"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Ítem"
            Columns(4).DataField=   "descripcion"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Encargado"
            Columns(5).DataField=   "nomresp"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Nº Reg. Prod."
            Columns(6).DataField=   "numregprod"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Tipo"
            Columns(7).DataField=   "destipo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Estado"
            Columns(8).DataField=   "desestado"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1455"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1376"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=1693"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=1614"
            Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(21)=   "Column(3).Width=1693"
            Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=1614"
            Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(27)=   "Column(4).Width=7673"
            Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=7594"
            Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(33)=   "Column(5).Width=6112"
            Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=6033"
            Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(39)=   "Column(6).Width=2196"
            Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=2117"
            Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(43)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(45)=   "Column(7).Width=2223"
            Splits(0)._ColumnProps(46)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._WidthInPix=2143"
            Splits(0)._ColumnProps(48)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(49)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(50)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(51)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(52)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(54)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(55)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(56)=   "Column(8).Order=9"
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
            _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=54,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Named:id=33:Normal"
            _StyleDefs(73)  =   ":id=33,.parent=0"
            _StyleDefs(74)  =   "Named:id=34:Heading"
            _StyleDefs(75)  =   ":id=34,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(76)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
            _StyleDefs(77)  =   "Named:id=35:Footing"
            _StyleDefs(78)  =   ":id=35,.parent=33,.alignment=3,.valignment=2,.bgcolor=&H8000000F&"
            _StyleDefs(79)  =   ":id=35,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   "Named:id=36:Selected"
            _StyleDefs(81)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(82)  =   "Named:id=37:Caption"
            _StyleDefs(83)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(84)  =   "Named:id=38:HighlightRow"
            _StyleDefs(85)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(86)  =   "Named:id=39:EvenRow"
            _StyleDefs(87)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(88)  =   "Named:id=40:OddRow"
            _StyleDefs(89)  =   ":id=40,.parent=33"
            _StyleDefs(90)  =   "Named:id=41:RecordSelector"
            _StyleDefs(91)  =   ":id=41,.parent=34"
            _StyleDefs(92)  =   "Named:id=42:FilterBar"
            _StyleDefs(93)  =   ":id=42,.parent=33"
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
            TabIndex        =   11
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
         Left            =   12540
         TabIndex        =   2
         Top             =   375
         Width           =   11805
         Begin VB.Frame Frame6 
            Height          =   945
            Left            =   30
            TabIndex        =   12
            Top             =   270
            Width           =   11745
            Begin VB.CommandButton CmdBusSup 
               Height          =   240
               Left            =   2070
               Picture         =   "FrmGenerarOrdProd2_1.frx":2EF3
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   240
               Width           =   240
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchPro 
               Height          =   300
               Left            =   1020
               TabIndex        =   15
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
               TabIndex        =   14
               Text            =   "TxtIdProg"
               Top             =   210
               Width           =   1320
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fecha"
               Height          =   195
               Left            =   90
               TabIndex        =   19
               Top             =   615
               Width           =   450
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Supervisor"
               Height          =   195
               Left            =   90
               TabIndex        =   18
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
               TabIndex        =   17
               Top             =   210
               Width           =   9225
            End
            Begin VB.Label LblProc 
               AutoSize        =   -1  'True
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
               Height          =   195
               Left            =   10890
               TabIndex        =   16
               Top             =   540
               Visible         =   0   'False
               Width           =   660
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
            Rows            =   2
            Cols            =   19
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmGenerarOrdProd2_1.frx":3025
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
            Left            =   60
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
Attribute VB_Name = "FrmGenerarOrdProd2_1"
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
Dim COLUMNA_REFERENCIA As Integer
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
Dim COLUMNA_IDREGPROD As Integer
Dim COLUMNA_IDREFERENCIA As Integer

Dim COLUMNA_DET_SELECCIONADO As Integer
Dim COLUMNA_DET_ITEM As Integer
Dim COLUMNA_DET_LOTE As Integer
Dim COLUMNA_DET_UM As Integer
Dim COLUMNA_DET_CANTIDAD As Integer
Dim COLUMNA_DET_IDORDEN As Integer
Dim COLUMNA_DET_IDITEM As Integer
Dim COLUMNA_DET_IDLOTEDET As Integer

Dim COLUMNAESTADO_ As Integer

Dim NUMERO_CORRELATIVO As Double

Dim numSolMax As Integer
Dim RstValores As New ADODB.Recordset

Dim CAMBIOGRABAR_ As Double
Dim ESTADOANTERIOR_ As Double

Const ESTADOPENDIENTE_ = 1
Const ESTADOPROCESADO_ = 2
Const ESTADOAPROBADO_ = 3
Const ESTADOANULADO_ = 4

'Para mover el frame
Dim OrigFX As Long
Dim OrigFY As Long

Private Sub cmd_Click(Index As Integer)
    Dim A As Integer
    Dim num As Integer
    Dim Rpta As Integer
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim MENSAJE_ As String
    
    Select Case Index
        Case 0 ' Agregar Solicitud
            If QueHace = 3 Then Exit Sub
            fg(0).Rows = fg(0).Rows + 1
            fg(0).Select fg(0).Rows - 1, 1
            Frm4.Visible = False
            
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNA_TIPO) = 1
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNA_SELECCIONADO) = -1
            
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNA_NUMORDEN) = Format(NulosN(numSolMax), "000000")
            numSolMax = numSolMax + 1
            
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNA_ID) = NulosN(NUMERO_CORRELATIVO)
            NUMERO_CORRELATIVO = NUMERO_CORRELATIVO + 1
            
            fg(0).TextMatrix(fg(0).Rows - 1, COLUMNAESTADO_) = ESTADOPROCESADO_
            
            
        Case 1 ' Eliminar Solicitud
            If QueHace = 3 Then Exit Sub
            If fg(0).Rows <= 0 Then Exit Sub
            
            If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAESTADO_)) >= ESTADOAPROBADO_ Then
                MsgBox "El registro esta aprobado y no se puede eliminar", vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
                Exit Sub
            End If
            
            ' Se verifica si hay registros con estado no pendiente
            If Not verificarCambioEstado(NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)), MENSAJE_) Then
                MsgBox MENSAJE_, vbInformation + vbOKOnly + vbDefaultButton1, nTitulo
                Exit Sub
            End If
            
            Rpta = MsgBox("¿Esta seguro de Eliminar esta Solicitud?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            
            If Rpta = vbYes Then
                RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID))
                limpiarRST RstValores, False
                ' Si se elimina el ultimo numero de solicitud se disminuye en uno
                If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_NUMORDEN)) = (numSolMax - 1) Then
                    numSolMax = numSolMax - 1
                End If
                                
                ' SE ELIMINAN LOS INGRESOS SALIDAS RELACIONADAS
                cSQL = "SELECT alm_ingreso.idprocorr, alm_ingreso.idorddet, alm_ingreso.id " _
                    + vbCr + "FROM alm_ingreso " _
                    + vbCr + "WHERE (((alm_ingreso.idorddet)=" & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & "));"
                
                RST_Busq xRs, cSQL, xCon
                
                If xRs.State = 0 Then GoTo SIGUIENTE_
                If xRs.RecordCount = 0 Then GoTo SIGUIENTE_
                
                xCon.Execute "DELETE * FROM alm_ingreso WHERE id = " & NulosN(xRs("id"))
                xCon.Execute "DELETE * FROM alm_ingresodet WHERE id = " & NulosN(xRs("id"))
                                
                CAMBIOGRABAR_ = -1
SIGUIENTE_:
                                
                fg(0).RemoveItem fg(0).Row
                Frm4.Visible = False
            End If
            
        Case 2 ' Ver detalle solicitud
            verDetalle
            
        Case 3 ' Agregar item en detalle
            If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 1 Then
                MsgBox "Esta operacion no esta permitida para este tipo de solicitud", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            'descripcion                  'campo                           'tamaño                         'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Producto":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":     xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
                            
            cSQL = "SELECT alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.tippro, alm_inventario.id, mae_unidades.abrev AS unimed, alm_inventario.idunimed " _
                + vbCr + "FROM alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                + vbCr + "Where ((alm_inventario.activo) = -1) " _
                + vbCr + "ORDER BY alm_inventario.descripcion;"
                
            nTitulo = "Buscando Items"
                    
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            If RstValores.State = 0 Then Exit Sub
            
            RstValores.AddNew
            RstValores("activo") = -1
            RstValores("idorddet") = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID))    ' Id orden de solicitud
            RstValores("iditem") = NulosN(xRs("id"))                                    ' iD item
            RstValores("descripcion") = NulosC(xRs("descripcion"))                      ' Descripcion
            RstValores("abrev") = NulosC(xRs("unimed"))                                 ' Abrev de UM
            RstValores("cantidad") = 0
            RstValores.Update
            
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
            pCargarValores
                    
        Case 4 ' Listar items en detalle
            If QueHace = 3 Then Exit Sub
            
            If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 1 Then
                MsgBox "Esta operacion no esta permitida para este tipo de solicitud", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            'descripcion                  'campo                           'tamaño                         'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Producto":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":     xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
                                    
            ' generar la consulta
            cSQL = "SELECT 0 AS xsel, alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.tippro, alm_inventario.id, mae_unidades.abrev AS unimed, alm_inventario.idunimed " _
                + vbCr + "FROM alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                + vbCr + "Where (((alm_inventario.tippro) = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREFERENCIA)) & ")) " _
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
                RstValores("activo") = -1
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
            
            If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 1 Then
                MsgBox "Esta operacion no esta permitida para este tipo de solicitud", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            Agregando = True
            eliminarRegistro
            
        Case 6 ' Eliminar todos los items en detalle
            If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 1 Then
                MsgBox "Esta operacion no esta permitida para este tipo de solicitud", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
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
    
    If fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREFERENCIA) = 3 Then Exit Sub
    
    If Agregando Then
        If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    End If
    
    If RstValores.RecordCount <> 0 Then RstValores.MoveFirst
        
    Do While Not RstValores.EOF
        If RstValores.RecordCount = 0 Then Exit Do
        If NulosN(RstValores("iditem")) = NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNA_DET_IDITEM)) Then
            RstValores.Delete
            Exit Do
        End If
        RstValores.MoveNext
    Loop
    
    RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID))
    pCargarValores
End Sub

Private Sub pCargarValores()
    Agregando = True
    
    fg(1).Rows = 1
    If RstValores.State = 0 Then Exit Sub
    If RstValores.RecordCount = 0 Then Exit Sub
    
    RstValores.MoveFirst
    With fg(1)
        Do While Not RstValores.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COLUMNA_DET_SELECCIONADO) = NulosN(RstValores("activo"))
            .TextMatrix(.Rows - 1, COLUMNA_DET_ITEM) = NulosC(RstValores("descripcion"))
            .TextMatrix(.Rows - 1, COLUMNA_DET_LOTE) = NulosC(RstValores("deslote"))
            .TextMatrix(.Rows - 1, COLUMNA_DET_UM) = NulosC(RstValores("abrev"))
            .TextMatrix(.Rows - 1, COLUMNA_DET_CANTIDAD) = Format(NulosN(RstValores("cantidad")), "0.0000")
            .TextMatrix(.Rows - 1, COLUMNA_DET_IDORDEN) = NulosC(RstValores("idorddet"))
            .TextMatrix(.Rows - 1, COLUMNA_DET_IDITEM) = NulosC(RstValores("iditem"))
            .TextMatrix(.Rows - 1, COLUMNA_DET_IDLOTEDET) = NulosC(RstValores("idlotedet"))
            
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

Private Sub PosicionarFrm(ByRef FRM_ As Frame)
    With FRM_
        .Top = Me.Height - 4155
        .Left = Me.Width - 6680
    End With
End Sub

Private Sub verDetalle()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim IDORDDET_ As Integer
    Dim IDRECETA_ As Integer
    Dim CANT_ As Double
            
    ' Si no se han agregado productos
    If fg(0).Rows = 1 Then
        MsgBox "No hay items que mostrar, agreguelos Productos para procesarlos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).SetFocus
        
        PosicionarFrm Frm4
        Frm4.Visible = True
        fg(1).Rows = 1
    End If
    
    ' Si no hay Productos escogidos
    If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 1 And _
                                    NulosN(fg(0).TextMatrix(fg(0).Row, columna_idpro)) = 0 Then
        MsgBox "No hay items que mostrar, agregue Item para procesarlos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).SetFocus
        Exit Sub
    End If
    
    IDORDDET_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID))
    Lblidorddet.Caption = IDORDDET_
    PosicionarFrm Frm4
    Frm4.Visible = True
    
    If RstValores.State = 0 Then Exit Sub
    RstValores.Filter = "idorddet = " & IDORDDET_
    
    ' Si no tiene insumos se carga de receta
    If RstValores.RecordCount = 0 Then
        IDRECETA_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREC))
        CANT_ = NulosN(fg(0).TextMatrix(fg(0).Row, columna_cantidad))
        cargarReceta IDRECETA_, CANT_
    End If
    
    RstValores.Filter = "idorddet = " & IDORDDET_
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

Private Sub ImprimirLinea()
    Dim A As Integer
    Dim numPag As Integer
    Dim Rst As New ADODB.Recordset
    Dim B As Integer
    Dim FILA_ As Integer
    Dim COLUMNA_ As Integer
    Dim numper As Double
    Dim xFila As Integer
    Dim Nombre As String
                
    With FrmVsPrinter.Vs
        numPag = 0
        .BrushColor = &H80000005
        .FontSize = 11
        .TextAlign = taCenterMiddle
        .StartDoc
        
        FILA_ = 800
        COLUMNA_ = 1000
        numPag = numPag + 1
        'If A > 1 Then .NewPage
        CrearCabeceraVS numPag
            
        For A = 1 To fg(0).Rows - 1
            
            If fg(0).TextMatrix(A, COLUMNA_SELECCIONADO) <> -1 Then GoTo SIGUIENTE
            
            If FILA_ >= 13000 Then
                .NewPage
                FILA_ = 800
                numPag = numPag + 1
                CrearCabeceraVS numPag
            End If
                        
            '******************************************************************* Titulo
            .FontSize = 12
            .FontBold = True
            .TextAlign = taCenterMiddle
            
            .TextBox "SOLICITUD DE MATERIALES", COLUMNA_, FILA_, 8000, 500, True, False, True
            .FontSize = 10
            .TextAlign = taCenterTop
            .TextBox "Nº ", COLUMNA_ + 8100, FILA_, 1900, 250, True, False, True
            FILA_ = FILA_ + 240
            .TextBox "001" & "-" & fg(0).TextMatrix(A, COLUMNA_NUMORDEN), COLUMNA_ + 8100, FILA_, 1900, 250, True, False, True
            
            .TextAlign = taLeftMiddle
            .FontSize = 9
            
            '*************************************************************************
            If fg(0).TextMatrix(A, COLUMNA_TIPO) = 1 Then
                FILA_ = FILA_ + 300
                .TextBox "Producto:", COLUMNA_, FILA_, 1500, 250, True, False, False
                .FontBold = False
                .TextBox fg(0).TextMatrix(A, COLUMNA_ITEM), COLUMNA_ + 1500, FILA_, 7000, 250, True, False, False
                
                .FontBold = True
                .TextBox "Lote:", COLUMNA_ + 7500, FILA_, 1500, 250, True, False, False
                .FontBold = False
                .TextBox fg(0).TextMatrix(A, COLUMNA_LOTE), COLUMNA_ + 9000, FILA_, 6000, 250, True, False, False
            End If
            
            '*************************************************************************
            FILA_ = FILA_ + 250
            
            .FontBold = True
            .TextBox "Programador:", COLUMNA_, FILA_, 1500, 250, True, False, False
            .FontBold = False
            If NulosN(xIdUsuario) = 0 Then
                .TextBox NulosC(LblNomProg.Caption), COLUMNA_ + 1500, FILA_, 6000, 250, True, False, False
            Else
                Nombre = Busca_Codigo(xIdUsuario, "id", "nomusu", "mae_usuarios", "N", xCon)
                .TextBox Nombre, COLUMNA_ + 1500, FILA_, 6000, 250, True, False, False
            End If
            
            .FontBold = True
            .TextBox "Fecha Prog.:", COLUMNA_ + 7500, FILA_, 1500, 250, True, False, False
            .FontBold = False
            .TextBox Format(TxtFchPro.valor, FORMAT_DATE), COLUMNA_ + 9000, FILA_, 6000, 250, True, False, False
            
            
            FILA_ = FILA_ + 250
            
            If fg(0).TextMatrix(A, COLUMNA_TIPO) = 1 Then ' Si no es solicitud
                .FontBold = True
                .TextBox "Receta:", COLUMNA_, FILA_, 1500, 250, True, False, False
                .FontBold = False
                .TextBox fg(0).TextMatrix(A, COLUMNA_RECETA), COLUMNA_ + 1500, FILA_, 6000, 250, True, False, False
            End If
            
            If fg(0).TextMatrix(A, COLUMNA_TIPO) = 1 Then
                .FontBold = True
                .TextBox "Can. Teo.:", COLUMNA_ + 3750, FILA_, 1500, 250, True, False, False
                .FontBold = False
                .TextBox fg(0).TextMatrix(A, columna_cantidad), COLUMNA_ + 5250, FILA_, 6000, 250, True, False, False
            End If
                
            If fg(0).TextMatrix(A, COLUMNA_TIPO) = 1 Or fg(0).TextMatrix(A, COLUMNA_TIPO) = 2 Then
                .FontBold = True
                .TextBox "Num. Prod.:", COLUMNA_ + 7500, FILA_, 1500, 250, True, False, False
                .FontBold = False
                .TextBox fg(0).TextMatrix(A, COLUMNA_REFERENCIA), COLUMNA_ + 9000, FILA_, 6000, 250, True, False, False
            Else
                FILA_ = FILA_ - 250
            End If
            
            If fg(0).TextMatrix(A, COLUMNA_TIPO) = 1 Then
                FILA_ = FILA_ + 400
                .FontBold = True
                .TextBox " Fch. Proc.", COLUMNA_, FILA_, 1250, 250, True, False, True
                .TextBox "", COLUMNA_ + 1250, FILA_, 1250, 250, True, False, True
                .TextBox " Can. Real", COLUMNA_ + 2500, FILA_, 1250, 250, True, False, True
                .TextBox "", COLUMNA_ + 3750, FILA_, 1250, 250, True, False, True
                .TextBox " Hor. Ini.", COLUMNA_ + 5000, FILA_, 1250, 250, True, False, True
                .TextBox "", COLUMNA_ + 6250, FILA_, 1250, 250, True, False, True
                .TextBox " Hor. Fin", COLUMNA_ + 7500, FILA_, 1250, 250, True, False, True
                .TextBox "", COLUMNA_ + 8750, FILA_, 1250, 250, True, False, True
                .FontBold = False
            End If
            
            '*************************************************************************
            FILA_ = FILA_ + 350
            .TextAlign = taCenterMiddle
            .TextBox "#", COLUMNA_, FILA_, 400, 500, True, False, True
            .TextBox "Lote", COLUMNA_ + 400, FILA_, 1100, 500, True, False, True
            .TextBox "Ítem", COLUMNA_ + 1500, FILA_, 3500, 500, True, False, True
            .TextBox "U.M.", COLUMNA_ + 5000, FILA_, 500, 500, True, False, True
            .TextBox "Cantidad Teorica", COLUMNA_ + 5500, FILA_, 1125, 500, True, False, True
            .TextBox "Cantidad Real", COLUMNA_ + 6625, FILA_, 1125, 500, True, False, True
            .TextBox "Adicional", COLUMNA_ + 7750, FILA_, 1125, 500, True, False, True
            .TextBox "Devolucion", COLUMNA_ + 8875, FILA_, 1125, 500, True, False, True
            
            ' Se filtra los Insumos utilizados
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(A, COLUMNA_ID)) & ""
            ' Si no hay insumos cargados se los carga de la BD
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
            
            ' Se verifica el estado del recordset
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
                FILA_ = FILA_ + 500
                xFila = FILA_
                For B = 1 To Rst.RecordCount
                    .FontSize = 8
                    .FontBold = False
                    .TextAlign = taLeftMiddle
                    
                    .TextBox " " & Format(B, "00"), COLUMNA_, FILA_, 400, 250, True, False, True
                    .FontSize = 7
                    .TextBox " " & Rst("deslote"), COLUMNA_ + 400, FILA_, 1100, 250, True, False, True
                    .TextBox " " & Rst("descripcion"), COLUMNA_ + 1500, FILA_, 3500, 250, True, False, True
                    .FontSize = 8
                    .TextAlign = taCenterMiddle
                    .TextBox Rst("abrev"), COLUMNA_ + 5000, FILA_, 500, 250, True, False, True
                    .TextAlign = taRightMiddle
                    .TextBox Format(Rst("cantidad"), "0.0000"), COLUMNA_ + 5500, FILA_, 1125, 250, True, False, True
                    .TextBox "", COLUMNA_ + 6625, FILA_, 1125, 250, True, False, True
                    .TextBox "", COLUMNA_ + 7750, FILA_, 1125, 250, True, False, True
                    .TextBox "", COLUMNA_ + 8875, FILA_, 1125, 250, True, False, True
                    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                    
                    FILA_ = FILA_ + 250
                    
                    If FILA_ >= 16200 Then
                        FILA_ = 800
                        numPag = numPag + 1
                        .NewPage
                        CrearCabeceraVS numPag
                    End If
                Next B
            End If
            
            FILA_ = FILA_ + 400
            If FILA_ >= 16000 Then
                FILA_ = 2000
                .NewPage
            End If
            
            'LADO A
            .TextBox "_______________________________", COLUMNA_ + 700, FILA_, 3500, 250, True, False, False
            .TextBox "_______________________________", COLUMNA_ + 5700, FILA_, 3500, 250, True, False, False
            
            FILA_ = FILA_ + 200
            
            .FontSize = 7
            .TextAlign = taCenterMiddle
            
            If NulosC(fg(0).TextMatrix(A, COLUMNA_RESPONSABLE)) = "" Then
                .TextBox "VºBº GER. PROD. ", COLUMNA_ + 1000, FILA_, 3500, 250, True, False, False
            Else
                .TextBox NulosC(fg(0).TextMatrix(A, COLUMNA_RESPONSABLE)), COLUMNA_ + 1000, FILA_, 3500, 250, True, False, False
            End If
            
            .TextBox "RESPONSABLE DE ALMACEN", COLUMNA_ + 6000, FILA_, 3500, 250, True, False, False
            .FontSize = 8
            
            FILA_ = FILA_ + 400
            
SIGUIENTE:
        Next A
        .EndDoc
    End With
    'Muestra la preimagen de la impresion
    FrmVsPrinter.WindowState = 2
    FrmVsPrinter.Show
End Sub

Private Sub ImprimirSolicitud()
    Dim A As Integer
    Dim B As Integer
    Dim Rst As New ADODB.Recordset
    Dim FILA_ As Integer
    Dim COLUMNA_ As Integer
    
    With FrmVsPrinter.Vs
        
        .BrushColor = &H80000005
        .FontSize = 11
        .TextAlign = taCenterMiddle
        .StartDoc
        
        FILA_ = 500
        COLUMNA_ = 700
        For A = 1 To fg(0).Rows - 1
            If FILA_ >= 13000 Then
                FILA_ = 500
                .NewPage
            End If
            If fg(0).TextMatrix(A, COLUMNA_SELECCIONADO) <> -1 Then GoTo SIGUIENTE
            'LADO A
            .FontSize = 13
            .TextAlign = taCenterMiddle
            .TextBox "SOLICITUD DE MATERIALES", COLUMNA_, FILA_, 8750, 500, True, False, True
            .FontSize = 10
            .TextAlign = taCenterTop
            .TextBox "Nº ", COLUMNA_ + 8800, FILA_, 1700, 250, True, False, True
            FILA_ = FILA_ + 240
            .TextBox "0001" & "-" & fg(0).TextMatrix(A, COLUMNA_NUMORDEN), COLUMNA_ + 8800, FILA_, 1700, 250, True, False, True
            
            .TextAlign = taLeftMiddle
            .FontSize = 9
            
            FILA_ = FILA_ + 400
            If fg(0).TextMatrix(A, COLUMNA_TIPO) = 1 Then
                .TextBox "Producto    ", COLUMNA_, FILA_, 1500, 250, True, False, False
                .TextBox fg(0).TextMatrix(A, COLUMNA_ITEM), COLUMNA_ + 1500, FILA_, 6000, 250, True, False, False
                FILA_ = FILA_ + 250
            End If
            
            .TextBox "Programador    ", COLUMNA_, FILA_, 1500, 250, True, False, False
            If NulosN(xIdUsuario) = 0 Then
                .TextBox NulosC(LblNomProg.Caption), COLUMNA_ + 1500, FILA_, 6000, 250, True, False, False
            Else
                Dim Nombre As String
                Nombre = Busca_Codigo(xIdUsuario, "id", "nomusu", "mae_usuarios", "N", xCon)
                .TextBox Nombre, COLUMNA_ + 1500, FILA_, 6000, 250, True, False, False
            End If
            
            .TextBox "Fch. Prod.   ", COLUMNA_ + 7500, FILA_, 1500, 250, True, False, False
            .TextBox TxtFchPro.valor, COLUMNA_ + 8700, FILA_, 6000, 250, True, False, False
            
            If fg(0).TextMatrix(A, COLUMNA_TIPO) = 1 Then
                FILA_ = FILA_ + 250
                .TextBox "Receta ", COLUMNA_, FILA_, 1500, 250, True, False, False
                .TextBox fg(0).TextMatrix(A, COLUMNA_RECETA), COLUMNA_ + 1500, FILA_, 1500, 250, True, False, False
            
                .TextBox "Cantidad   ", COLUMNA_ + 7500, FILA_, 1500, 250, True, False, False
                .TextBox fg(0).TextMatrix(A, columna_cantidad), COLUMNA_ + 8700, FILA_, 6000, 250, True, False, False
            End If
                        
            FILA_ = FILA_ + 250
            .TextBox "Lote   ", COLUMNA_, FILA_, 1500, 250, True, False, False
            .TextBox fg(0).TextMatrix(A, COLUMNA_LOTE), COLUMNA_ + 1500, FILA_, 4500, 250, True, False, False
            
            FILA_ = FILA_ + 300
            .TextAlign = taCenterMiddle
            .TextBox "Item", COLUMNA_, FILA_, 400, 500, True, False, True
            .TextBox "Insumo / Producto / MP", COLUMNA_ + 400, FILA_, 3700, 500, True, False, True
            .TextBox "U.M", COLUMNA_ + 4100, FILA_, 400, 500, True, False, True
            .TextBox "Lote", COLUMNA_ + 4500, FILA_, 2000, 500, True, False, True
            .TextBox "Cantidad Teorica", COLUMNA_ + 6500, FILA_, 1000, 500, True, False, True
            .TextBox "Cantidad Real", COLUMNA_ + 7500, FILA_, 1000, 500, True, False, True
            .TextBox "Adicional", COLUMNA_ + 8500, FILA_, 1000, 500, True, False, True
            .TextBox "Devolucion", COLUMNA_ + 9500, FILA_, 1000, 500, True, False, True
                 
            ' Se filtra los Insumos utilizados
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(A, COLUMNA_ID)) & ""
            ' Si no hay insumos cargados se los carga de la BD
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
            ' Se verifica el estado del recordset
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
                FILA_ = FILA_ + 500
                For B = 1 To Rst.RecordCount
                    .FontSize = 8
                    .TextAlign = taLeftMiddle
                    .TextBox " " & Format(B, "00"), COLUMNA_, FILA_, 400, 250, True, False, True
                    .TextBox " " & NulosC(Rst("descripcion")), COLUMNA_ + 400, FILA_, 3700, 250, True, False, True
                    .TextAlign = taCenterMiddle
                    .TextBox NulosC(Rst("abrev")), COLUMNA_ + 4100, FILA_, 400, 250, True, False, True
                    .TextBox NulosC(Rst("lote")), COLUMNA_ + 4500, FILA_, 2000, 250, True, False, True
                    .TextAlign = taRightMiddle
                    .TextBox Format(NulosN(Rst("cantidad")), "0.000000"), COLUMNA_ + 6500, FILA_, 1000, 250, True, False, True
                    .TextBox "", COLUMNA_ + 7500, FILA_, 1000, 250, True, False, True
                    .TextBox "", COLUMNA_ + 8500, FILA_, 1000, 250, True, False, True
                    .TextBox "", COLUMNA_ + 9500, FILA_, 1000, 250, True, False, True
                    
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                    
                    FILA_ = FILA_ + 250
                    
                    If FILA_ >= 16200 Then
                        FILA_ = 500
                        .NewPage
                    End If
                Next B
            End If
            
            ' POSICION ANTES DEL DETALLE + ALTO DE DE 10 ITEMS + 500 DE ESPACIO
            FILA_ = FILA_ + 500
            If FILA_ >= 16400 Then
                FILA_ = 2000
                .NewPage
            End If
            'LADO A
            .TextBox "_______________________________", COLUMNA_ + 700, FILA_, 3500, 200, True, False, False
            .TextBox "_______________________________", COLUMNA_ + 5700, FILA_, 3500, 200, True, False, False
            
            FILA_ = FILA_ + 200
            
            .FontSize = 6
            .TextAlign = taCenterMiddle
            
            If NulosC(fg(0).TextMatrix(A, COLUMNA_RESPONSABLE)) = "" Then
                .TextBox "VºBº Ger. Prod. ", COLUMNA_ + 1000, FILA_, 3500, 250, True, False, False
            Else
                .TextBox NulosC(fg(0).TextMatrix(A, COLUMNA_RESPONSABLE)), COLUMNA_ + 1000, FILA_, 3500, 250, True, False, False
            End If
            
            .TextBox "Responsable de Almacen", COLUMNA_ + 6000, FILA_, 3500, 250, True, False, False
            .FontSize = 8
            
            FILA_ = FILA_ + 500
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
            RstValores.Filter = "idorddet = " & fg(0).TextMatrix(fg(0).Row, COLUMNA_ID) & " And iditem = " & fg(1).TextMatrix(A, COLUMNA_DET_IDITEM)
            RstValores("activo") = seleccionar
        End If
    Next A
End Sub

Private Sub fg_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Index
        Case 0:
            If NulosN(fg(Index).TextMatrix(Row, COLUMNAESTADO_)) >= ESTADOAPROBADO_ Then Cancel = True
            
            Select Case Col
                Case COLUMNA_ITEM, COLUMNA_RECETA, COLUMNA_UM
                    Cancel = True
                
                Case COLUMNA_REFERENCIA
                    If NulosN(fg(0).TextMatrix(Row, COLUMNA_TIPO)) = 3 Then Cancel = True
                    
                Case columna_cantidad
                    If NulosN(fg(0).TextMatrix(Row, COLUMNA_TIPO)) = 2 Then Cancel = True
                    If NulosN(fg(0).TextMatrix(Row, COLUMNA_TIPO)) = 3 Then Cancel = True
                    
                Case COLUMNAESTADO_
                    ' Se llena el estado anterior
                    ESTADOANTERIOR_ = NulosN(fg(0).TextMatrix(Row, Col))
                    
            End Select
            
        Case 1:
            Select Case Col
                Case COLUMNA_DET_LOTE
                
                Case Else
                    If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 1 Then Cancel = True
                    If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAESTADO_)) >= ESTADOAPROBADO_ Then Cancel = True
            End Select
            
    End Select
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

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Index = 0 Then
        Dim xform As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        Dim nTitulo As String
        Dim xCampos() As String
        
        If QueHace = 3 Then Exit Sub
        
        If Col = COLUMNA_REFERENCIA Then ' Buscando Referencia
            cargarCampos False, True
        End If
        
        If Col = COLUMNA_RESPONSABLE Then ' Buscando Personal Responsable
            ReDim xCampos(2, 4) As String
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
    End If
    
    If Index = 1 Then
        If Col = COLUMNA_DET_ITEM Then
            cmd_Click 3
        End If
        
        If Col = COLUMNA_DET_LOTE Then
            ReDim xCampos(4, 4) As String
            
            ' Se verifica si se escogio el producto
            If NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNA_DET_IDITEM)) = 0 Then
                MsgBox "Seleccione el Ítem para el registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                fg(1).Col = COLUMNA_DET_ITEM
                Exit Sub
            End If
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Lote":         xCampos(0, 1) = "deslote":      xCampos(0, 2) = "2000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Fch. Ing.":    xCampos(1, 1) = "fching":       xCampos(1, 2) = "1000":         xCampos(1, 3) = "D"
            xCampos(2, 0) = "Almacen":      xCampos(2, 1) = "desalm":       xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
            xCampos(3, 0) = "Cantidad":     xCampos(3, 1) = "cantidad":     xCampos(3, 2) = "1000":         xCampos(3, 3) = "N"
                    
            nTitulo = "Buscando Lotes de " & NulosC(fg(1).TextMatrix(fg(1).Row, COLUMNA_DET_ITEM))
    
            cSQL = "SELECT alm_inventariolotedet.idlote, alm_inventariolotedet.id AS idlotedet, alm_inventariolote.iditem, alm_inventariolotedet.idalm, alm_inventariolote.fching, alm_almacenes.descripcion AS desalm, alm_inventariolotedet.cantidad, alm_inventariolote.descripcion AS deslote " _
                + vbCr + "FROM (alm_inventariolote LEFT JOIN alm_inventariolotedet ON alm_inventariolote.id = alm_inventariolotedet.idlote) LEFT JOIN alm_almacenes ON alm_inventariolotedet.idalm = alm_almacenes.id " _
                + vbCr + "WHERE (((alm_inventariolote.iditem)=" & NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNA_DET_IDITEM)) & "))"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "deslote", "deslote", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            If xRs("cantidad") < NulosN(fg(1).TextMatrix(Row, COLUMNA_DET_CANTIDAD)) Then
                MsgBox "El lote seleccionado no contiene stock suficiente", vbCritical + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            ' Lote
            fg(1).TextMatrix(Row, COLUMNA_DET_IDLOTEDET) = NulosN(xRs("idlotedet"))
            fg(1).TextMatrix(Row, COLUMNA_DET_LOTE) = NulosC(xRs("deslote"))
            
            Set xRs = Nothing
        End If
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
            
            Select Case NulosN(fg(0).TextMatrix(Row, Col))
                Case 2
                    fg(0).TextMatrix(Row, COLUMNA_ITEM) = "VER DETALLE"
                    
                Case 3
                    fg(0).TextMatrix(Row, COLUMNA_ITEM) = "VER DETALLE"
                    PosicionarFrm Frm4
                    fg(1).Rows = 1
                    Frm4.Visible = True
                
            End Select
        End If
        
        If Col = columna_cantidad Then 'Cambiar cantidad
            fg(0).TextMatrix(Row, Col) = Format(fg(0).TextMatrix(Row, Col), "0.00")
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
            limpiarRST RstValores, False
            ' Carga los Datos segun la nueva cantidad
            cargarReceta NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREC)), NulosN(fg(0).TextMatrix(fg(0).Row, columna_cantidad))
            ' Muestra los Datos
            pCargarValores
        End If
    End If
    
    If Index = 1 Then
        If Agregando = True Then Exit Sub
        
        If Col = COLUMNA_DET_LOTE Then ' Lote
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & _
                                " And iditem = " & NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNA_DET_IDITEM)) & ""
                                
            If RstValores.RecordCount = 0 Then Exit Sub
            RstValores("lote") = NulosC(fg(Index).TextMatrix(Row, Col))
            RstValores.Update
        End If
        
        If Col = COLUMNA_DET_CANTIDAD Then ' Cantidad
            fg(Index).TextMatrix(Row, Col) = Format(fg(Index).TextMatrix(Row, Col), "0.0000")
            RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & _
                                " And iditem = " & NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNA_DET_IDITEM)) & ""
                                
            If RstValores.RecordCount = 0 Then Exit Sub
            RstValores("cantidad") = NulosN(fg(Index).TextMatrix(Row, Col))
            RstValores.Update
        End If
    End If
End Sub

Private Function cambiarEstadoRelacionados(IDORDDET_ As Double, ESTADO_ As Double) As Boolean
    Dim ID_ As Double
    
    On Error GoTo ERROR_
    ' Salidas de Almacen
    cSQL = "UPDATE alm_ingreso SET alm_ingreso.estado = " & ESTADO_ & " " _
        + vbCr + "WHERE (((alm_ingreso.idorddet)=" & IDORDDET_ & "));"

    xCon.Execute cSQL
    
    ' GRABAMOS LOS MOVIMIENTOS
    ' INGRESOS Y SALIDAS DE ALMACEN
    ID_ = Busca_Codigo(IDORDDET_, "idorddet", "id", "alm_ingreso", "N", xCon)
    GrabarOperacion xIdUsuario, 8, 7, xHorIni, Time, Date, xCon, ID_
        
    cambiarEstadoRelacionados = True
    Exit Function
ERROR_:
    MsgBox "Ha ocurrido un error al tratar de cambiar de estado", vbInformation, xTitulo
    cambiarEstadoRelacionados = False
End Function

Private Sub Fg_Click(Index As Integer)
    Dim Rpta As Integer
    Dim NINGUNERROR_ As Boolean
    Dim MENSAJE_ As String
    Dim NUMSOL_ As String
    
    If Index = 0 Then
        If QueHace = 3 Then
            Select Case fg(0).Col
                Case COLUMNA_SELECCIONADO
                    If fg(0).TextMatrix(fg(0).Row, COLUMNA_SELECCIONADO) = 0 Then
                        fg(0).TextMatrix(fg(0).Row, COLUMNA_SELECCIONADO) = -1
                    Else
                        fg(0).TextMatrix(fg(0).Row, COLUMNA_SELECCIONADO) = 0
                    End If
            
'                Case COLUMNAESTADO_
'                    If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAESTADO_)) = 0 Then
'                        Rpta = MsgBox("¿Aprobar este Evento lo dejara bloqueado para su modificacion; desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
'
'                        If Rpta = vbNo Then Exit Sub
'                        NUMSOL_ = NulosC(fg(0).TextMatrix(fg(0).Row, COLUMNA_NUMORDEN))
'                        NINGUNERROR_ = GrabarAlmacen(NUMSOL_)
'                        MENSAJE_ = "Ha ocurrido un error al intentar crear el Registro de Ingreso de Produccion; se cancelara la operación"
'
'                        If Not NINGUNERROR_ Then
'                            MsgBox MENSAJE_, vbInformation, xTitulo
'                            Exit Sub
'                        End If
'
'                        ' Se actualiza el estado a cerrado
'                        cSQL = "UPDATE pro_ordenproddet " _
'                            + vbCr + "SET pro_ordenproddet.cerrado = -1 " _
'                            + vbCr + "WHERE (((pro_ordenproddet.numdoc)='" & NUMSOL_ & "'));"
'
'                        xCon.Execute cSQL
'
'                        fg(0).TextMatrix(fg(0).Row, COLUMNAESTADO_) >= ESTADOAPROBADO_
'                    End If
                    
            End Select
        End If
    End If
End Sub

Private Sub fg_ComboCloseUp(Index As Integer, ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    Dim IDORD_ As Double
    Dim ESTADO_ As Double
    Dim Rpta As Integer
    Dim MENSAJE_ As String
    
    If Index = 0 Then
        If Col = COLUMNAESTADO_ Then
            Rpta = MsgBox("¿ Esta seguro de cambiar el estado actual?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)

            If Rpta = vbNo Then
                fg(0).TextMatrix(fg(0).Row, Col) = ESTADOANTERIOR_
                Exit Sub
            End If
            
            IDORD_ = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID))
            ESTADO_ = NulosN(fg(0).TextMatrix(fg(0).Row, Col))
                
            If ESTADOANTERIOR_ > ESTADO_ Then
                MsgBox "Este cambio de estado no esta permitido", vbInformation, xTitulo
            Else
                If ESTADO_ <> ESTADOANULADO_ Then Exit Sub
                
                If verificarCambioEstado(IDORD_, MENSAJE_) Then
                    If cambiarEstadoRelacionados(IDORD_, ESTADO_) Then
                        CAMBIOGRABAR_ = -1
                    Else
                        fg(0).TextMatrix(fg(0).Row, Col) = ESTADOANTERIOR_
                    End If
                Else
                    MsgBox MENSAJE_, vbInformation, xTitulo
                    fg(0).TextMatrix(fg(0).Row, Col) = ESTADOANTERIOR_
                End If
            End If
        End If
        
        Frm4.Visible = False
    End If
End Sub

Private Function verificarCambioEstado(IDORDDET_ As Double, ByRef MENSAJE_ As String) As Boolean
    Dim xRs As New ADODB.Recordset
            
    ' Buscando Registros de Almacen
    cSQL = "SELECT alm_ingreso.idprocorr, alm_ingreso.estado " _
        + vbCr + "FROM alm_ingreso " _
        + vbCr + "WHERE (((alm_ingreso.idorddet)=" & IDORDDET_ & ") AND ((alm_ingreso.estado)>=2));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MENSAJE_ = "Registros de Almacen"
    
    If xRs.State = 0 Then verificarCambioEstado = False: GoTo SALIR_
    If xRs.RecordCount > 0 Then verificarCambioEstado = False: GoTo SALIR_
    
    verificarCambioEstado = True
    Exit Function
    
SALIR_:
    MENSAJE_ = "Se han encontrado " & MENSAJE_ & " que se encuentran en un estado no modificable; " _
    & vbCr & "verifique la condición de dichos Registros para completar esta acción."
End Function

Private Sub Fg_DblClick(Index As Integer)
    If Index = 0 Then
        bloquearControles
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
            Case COLUMNA_REFERENCIA, COLUMNA_ITEM, COLUMNA_RESPONSABLE, COLUMNA_UM, COLUMNA_RECETA, COLUMNA_NUMORDEN
                KeyAscii = 0
            
            Case columna_cantidad
                ' Si no es Producto entonces no se edita la cantidad y se sale del sub
                If NulosN(fg(0).TextMatrix(Row, COLUMNA_TIPO)) = 1 Then
                    ' Si no es un numero no se edita
                    If validar_numero(KeyAscii) = False Then KeyAscii = 0
                Else
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
    End If
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                                " And iditem = " & NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNA_DET_IDITEM)) & ""
                                
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
        bloquearControles
        
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

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 4000 Then Me.Height = 4000

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

Private Sub bloquearControles()
    Dim ESTADO_ As Boolean
    
    ' Se verifica el estado para bloquear
    If QueHace = 3 Then
        ESTADO_ = False
    Else
        If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAESTADO_)) >= ESTADOAPROBADO_ Then
            ESTADO_ = False
        Else
            ESTADO_ = True
        End If
        
        ' Si proviene de Receta se bloquea
        If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREFERENCIA)) = 3 Then
            ESTADO_ = False
        End If
    End If
    
    cmd(3).Enabled = ESTADO_ ' Agregar
    cmd(4).Enabled = ESTADO_ ' Seleccionar
    cmd(5).Enabled = ESTADO_ ' Eliminar
    cmd(6).Enabled = ESTADO_ ' Eliminar Todos
End Sub

Private Sub limpiarDatosFila()
    fg(0).TextMatrix(fg(0).Row, COLUMNA_ITEM) = ""
    fg(0).TextMatrix(fg(0).Row, columna_idpro) = ""
    fg(0).TextMatrix(fg(0).Row, COLUMNA_REFERENCIA) = ""
    fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREFERENCIA) = ""
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
    cSQL = "SELECT pro_ordenproddetins.idorddet, pro_ordenproddetins.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_ordenproddetins.cantidad, pro_ordenproddetins.activo, pro_ordenproddetins.idlotedet, alm_inventariolote.descripcion AS deslote " _
        + vbCr + "FROM (((pro_ordenproddetins LEFT JOIN alm_inventario ON pro_ordenproddetins.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN alm_inventariolotedet ON pro_ordenproddetins.idlotedet = alm_inventariolotedet.id) LEFT JOIN alm_inventariolote ON alm_inventariolotedet.idlote = alm_inventariolote.id " _
        + vbCr + "WHERE (((pro_ordenproddetins.idord) = " & idCodigo & ")) " _
        + vbCr + "ORDER BY pro_ordenproddetins.idorddet;"
          
    RST_Busq RstTmp, cSQL, xCon
    
    If RstValores.State = 0 Then DEFINIR_RST_TMP RstValores, RstTmp
    CARGAR_RST_TMP RstValores, RstTmp
    
    Set RstTmp = Nothing
End Sub

Private Sub iniciarCampos()
    COLUMNA_SELECCIONADO = 1
    '*********************
    COLUMNA_NUMORDEN = 2
    '*********************
    COLUMNA_TIPO = 3
    COLUMNA_ITEM = 4
    COLUMNA_RECETA = 5
    COLUMNA_UM = 6
    COLUMNA_RESPONSABLE = 7
    columna_cantidad = 8
    COLUMNA_LOTE = 9
    
    COLUMNAESTADO_ = 10
    '*************************
    COLUMNA_REFERENCIA = 11
    '*************************
    columna_idpro = 12
    COLUMNA_IDREC = 13
    COLUMNA_IDUNIMED = 14
    COLUMNA_IDRESP = 15
    COLUMNA_ID = 16
    COLUMNA_IDREGPROD = 17
    COLUMNA_IDREFERENCIA = 18
    
    COLUMNA_DET_SELECCIONADO = 1
    COLUMNA_DET_ITEM = 2
    COLUMNA_DET_LOTE = 3
    COLUMNA_DET_UM = 4
    COLUMNA_DET_CANTIDAD = 5
    COLUMNA_DET_IDORDEN = 6
    COLUMNA_DET_IDITEM = 7
    COLUMNA_DET_IDLOTEDET = 8
    
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
    fg(0).ColWidth(COLUMNA_IDREGPROD) = 0
    fg(0).ColWidth(COLUMNA_IDREFERENCIA) = 0
    
    GRID_COMBOLIST fg(0), COLUMNA_RESPONSABLE
    GRID_COMBOLIST fg(0), COLUMNA_RECETA
    GRID_COMBOLIST fg(0), COLUMNA_ITEM
    GRID_COMBOLIST fg(0), COLUMNA_REFERENCIA
    GRID_COMBOLIST fg(0), COLUMNA_TIPO
    
    GRID_COMBOLIST fg(1), COLUMNA_DET_ITEM
    GRID_COMBOLIST fg(1), COLUMNA_DET_LOTE
    
    fg(1).ColWidth(COLUMNA_DET_IDORDEN) = 0
    fg(1).ColWidth(COLUMNA_DET_IDITEM) = 0
    fg(1).ColWidth(COLUMNA_DET_IDLOTEDET) = 0
    
    Dg1.Columns("numsol").Alignment = dbgCenter
    Dg1.Columns("numregprod").Alignment = dbgCenter
    
    ' Se agregan los tipos
    ' Tipo
    cargarCampos True, False
    ' Referencia
    cargarCampos False, True
    ' Se agrega el mes Activo
    mMesActivo = xMes
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    
    ' Se pone el cuadro de seleccion en la cabecera del flexgrid
    fg(0).Select 0, 1
    fg(0).CellChecked = flexChecked
    
    fg(1).Select 0, 1
    fg(1).CellChecked = flexUnchecked
    
    CAMBIOGRABAR_ = 0
    ESTADOANTERIOR_ = 1
End Sub

Private Sub llenarEstados(ByRef FGGRID As VSFlexGrid, columna As Integer)
    Dim CAMPOS As String
    Dim xRs As New ADODB.Recordset
    
    cSQL = "SELECT * FROM mae_estados ORDER BY id"
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then
        MsgBox "No se ha encontrado estados, Ingrese estados", vbInformation, xTitulo
        Exit Sub
    End If
    
    xRs.MoveFirst
    CAMPOS = "#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
    xRs.MoveNext
    While Not xRs.EOF
        CAMPOS = CAMPOS & "|#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
        xRs.MoveNext
    Wend
    FGGRID.ColComboList(columna) = CAMPOS
End Sub

Private Sub cargarCampos(TIPO_ As Boolean, REFERENCIA_ As Boolean)
    Dim xRs As New ADODB.Recordset
    Dim CAMPOS As String
    Dim A As Integer
    Dim xCampos() As String
    Dim nTitulo As String
    
    If TIPO_ Then
        CAMPOS = "#1;PRODUCCION|#2;SOLICITUD|#3;OTRO"
        fg(0).ColComboList(COLUMNA_TIPO) = CAMPOS
    End If
    
    If REFERENCIA_ Then
        ' Si no es de tipo 3 sale
        Set xRs = Nothing
        
        Select Case NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO))
            ' Si viene de una Produccion
            Case 1, 2
                ReDim xCampos(6, 4) As String
                'descripcion                        'campo                          'tamaño                         'tipo = Numerico, caracter, fecha
                xCampos(0, 0) = "Num. Prod.":       xCampos(0, 1) = "numparte":     xCampos(0, 2) = "1200":         xCampos(0, 3) = "C"
                xCampos(1, 0) = "Descripcion":      xCampos(1, 1) = "despro":       xCampos(1, 2) = "3500":         xCampos(1, 3) = "C"
                xCampos(2, 0) = "Fech. Pro.":       xCampos(2, 1) = "dia":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
                xCampos(3, 0) = "Hor. Pro.":        xCampos(3, 1) = "horini":       xCampos(3, 2) = "900":          xCampos(3, 3) = "C"
                xCampos(4, 0) = "U.M":              xCampos(4, 1) = "abrev":        xCampos(4, 2) = "500":          xCampos(4, 3) = "C"
                xCampos(5, 0) = "Cantidad":         xCampos(5, 1) = "cantidad":     xCampos(5, 2) = "1000":         xCampos(5, 3) = "N"
                    
                cSQL = "SELECT pro_produccion.dia, pro_receta.iditem, alm_inventario.descripcion AS despro, pro_producciondet.idrec, pro_receta.codrec, pro_producciondet.horini, pro_producciondet.horfin, pro_producciondet.numparte, pro_producciondet.cantidad, pro_receta.idunimed, mae_unidades.abrev, pro_emp.idemp AS idresp, pla_empleados.nombre, pro_producciondet.corr AS idregprod " _
                        + vbCr + "FROM pro_produccion LEFT JOIN (((((pro_producciondet LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) LEFT JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id) ON pro_produccion.id = pro_producciondet.idpro;"

                nTitulo = "Buscando Reg. Prod."
                        
                CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "dia", "numparte", CualquierParte
                
                If xRs.State = 0 Then Exit Sub
                If xRs.RecordCount = 0 Then Exit Sub
                
                Agregando = True
                fg(0).TextMatrix(fg(0).Row, COLUMNA_SELECCIONADO) = -1
                fg(0).TextMatrix(fg(0).Row, COLUMNA_REFERENCIA) = NulosC(xRs("numparte"))
                
                ' Tipo 1 :Se muestran todos los datos
                If NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_TIPO)) = 1 Then
                    fg(0).TextMatrix(fg(0).Row, COLUMNA_ITEM) = NulosC(xRs("despro"))
                    fg(0).TextMatrix(fg(0).Row, columna_idpro) = NulosN(xRs("iditem"))
                    fg(0).TextMatrix(fg(0).Row, COLUMNA_UM) = NulosC(xRs("abrev"))
                    fg(0).TextMatrix(fg(0).Row, COLUMNA_RECETA) = NulosC(xRs("codrec"))
                    fg(0).TextMatrix(fg(0).Row, columna_cantidad) = Format(NulosN(xRs("cantidad")), "0.00")
                    fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREC) = NulosN(xRs("idrec"))
                    fg(0).TextMatrix(fg(0).Row, COLUMNA_IDUNIMED) = NulosN(xRs("idunimed"))
                Else
                    ' Tipo 2, 3
                    fg(0).TextMatrix(fg(0).Row, COLUMNA_ITEM) = "VER DETALLE"
                    
                End If
                
                fg(0).TextMatrix(fg(0).Row, COLUMNA_RESPONSABLE) = NulosC(xRs("nombre"))
                fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREGPROD) = NulosN(xRs("idregprod"))
                fg(0).TextMatrix(fg(0).Row, COLUMNA_IDRESP) = NulosN(xRs("idresp"))
                Agregando = False
                
                If RstValores.State = 0 Then Exit Sub
                RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
                limpiarRST RstValores, False
        
                cargarReceta NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREC)), NulosN(fg(0).TextMatrix(fg(0).Row, columna_cantidad))
                RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
                pCargarValores
                PosicionarFrm Frm4
                Frm4.Visible = True
                
        End Select
    End If
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

    cSQL = "SELECT pro_ordenprod.id, pro_ordenproddet.id AS iddet, pro_ordenprod.idsup, pla_empleados_1.nombre AS nomsup, pro_ordenproddet.idresp, pro_ordenproddet.tipo AS idtippro, [fchemi] & '' AS fchpro, pro_ordenproddet.idprocorr AS idregprod, pro_producciondet.numparte AS numregprod, pla_empleados.nombre AS nomresp, IIf([pro_ordenproddet].[tipo]=1,'PRODUCCION',IIf([pro_ordenproddet].[tipo]=2,'SOLICITUD','OTRO')) AS destipo, alm_inventario.descripcion, pro_ordenproddet.numdoc AS numsol, pro_ordenproddet.estado AS idestado, UCase([mae_estados].[descripcion]) AS desestado " _
            + vbCr + "FROM ((((((pro_ordenprod RIGHT JOIN pro_ordenproddet ON pro_ordenprod.id = pro_ordenproddet.idord) LEFT JOIN alm_inventario ON pro_ordenproddet.iditem = alm_inventario.id) LEFT JOIN pla_empleados ON pro_ordenproddet.idresp = pla_empleados.id) LEFT JOIN pro_emp ON pro_ordenprod.idsup = pro_emp.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_emp.idemp = pla_empleados_1.id) LEFT JOIN mae_estados ON pro_ordenproddet.estado = mae_estados.id) LEFT JOIN pro_producciondet ON pro_ordenproddet.idprocorr = pro_producciondet.corr " _
            + vbCr + "WHERE ((Month(pro_ordenprod.fchemi) = " & mMesActivo & ") And (Year(pro_ordenprod.fchemi) = " & Val(AnoTra) & ")) " _
            + vbCr + "ORDER BY [fchemi] & '' DESC , pro_ordenproddet.numdoc DESC;"
    
    ' cargando datos
    Me.MousePointer = vbHourglass
    
    RST_Busq RstOrd, cSQL, xCon
    Set Dg1.DataSource = RstOrd
    
    Me.MousePointer = vbDefault
    
    If RstOrd.State = 0 Then Exit Sub
End Sub

Private Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    Dim Rpta As Integer
    
    Blanquea
    llenarEstados fg(0), COLUMNAESTADO_
    
    If RstOrd.RecordCount = 0 Then Exit Sub
    If RstOrd.EOF = True Then Exit Sub
     
    Set RstDet = Nothing
    Agregando = True
    
    cSQL = "SELECT pro_ordenproddet.id, alm_inventario.descripcion AS prod, mae_unidades.abrev AS unid, pro_ordenproddet.cantidad, pro_receta.codrec, pro_ordenproddet.lote, pro_ordenproddet.numdoc, pro_ordenproddet.iditem, pro_ordenproddet.idrec, pro_ordenproddet.idunimed, pla_empleados.nombre, pro_ordenproddet.idresp, pro_ordenproddet.tipo, pro_ordenproddet.idprocorr AS idregprod, pro_producciondet.numparte AS numregprod, pro_ordenproddet.estado AS idestado " _
        + vbCr + "FROM ((((pro_ordenproddet LEFT JOIN alm_inventario ON pro_ordenproddet.iditem = alm_inventario.id) LEFT JOIN pro_receta ON pro_ordenproddet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_ordenproddet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_ordenproddet.idresp = pla_empleados.id) LEFT JOIN pro_producciondet ON pro_ordenproddet.idprocorr = pro_producciondet.corr " _
        + vbCr + "WHERE (((pro_ordenproddet.idord)=" & NulosN(RstOrd("id")) & ") AND ((pro_ordenproddet.fchprog)=CDate('" & RstOrd("fchpro") & "')));"
    
    RST_Busq RstDet, cSQL, xCon
    
    If IsDate(RstOrd("fchpro")) = True Then TxtFchPro.valor = CDate(RstOrd("fchpro"))
    
    TxtIdProg.Text = NulosN(RstOrd("idsup"))
    LblNomProg.Caption = NulosC(RstOrd("nomsup"))
    fg(0).Rows = fg(0).FixedRows
        
    If RstDet.State = 0 Then Exit Sub
    Agregando = True
    If RstDet.RecordCount = 0 Then Exit Sub
    
    RstDet.MoveFirst
    While Not RstDet.EOF
        fg(0).Rows = fg(0).Rows + 1
        With fg(0)
            .TextMatrix(.Rows - 1, COLUMNA_ID) = NulosN(RstDet("id"))
            .TextMatrix(.Rows - 1, COLUMNA_IDREGPROD) = NulosN(RstDet("idregprod"))
            .TextMatrix(.Rows - 1, COLUMNA_SELECCIONADO) = -1
            .TextMatrix(.Rows - 1, COLUMNA_TIPO) = NulosN(RstDet("tipo"))
            .TextMatrix(.Rows - 1, COLUMNA_RESPONSABLE) = NulosC(RstDet("nombre"))
            
            If NulosN(.TextMatrix(.Rows - 1, COLUMNA_TIPO)) = 1 Then
                .TextMatrix(.Rows - 1, columna_cantidad) = Format(NulosN(RstDet("cantidad")), "0.00")
                .TextMatrix(.Rows - 1, COLUMNA_ITEM) = NulosC(RstDet("prod"))
                .TextMatrix(.Rows - 1, COLUMNA_REFERENCIA) = NulosC(RstDet("numregprod"))
                .TextMatrix(.Rows - 1, COLUMNA_RECETA) = NulosC(RstDet("codrec"))
                .TextMatrix(.Rows - 1, COLUMNA_UM) = NulosC(RstDet("unid"))
                .TextMatrix(.Rows - 1, columna_idpro) = NulosN(RstDet("iditem"))
                .TextMatrix(.Rows - 1, COLUMNA_IDREC) = NulosN(RstDet("idrec"))
                .TextMatrix(.Rows - 1, COLUMNA_IDUNIMED) = NulosN(RstDet("idunimed"))
            Else
                If NulosN(.TextMatrix(.Rows - 1, COLUMNA_TIPO)) = 2 Then
                    .TextMatrix(.Rows - 1, COLUMNA_ITEM) = "VER DETALLE"
                    .TextMatrix(.Rows - 1, COLUMNA_RECETA) = ""
                    .TextMatrix(.Rows - 1, COLUMNA_UM) = ""
                    .TextMatrix(.Rows - 1, columna_cantidad) = ""
                    .TextMatrix(.Rows - 1, COLUMNA_REFERENCIA) = NulosC(RstDet("numregprod"))
                Else
                    If NulosN(RstDet("iditem")) = 0 Then
                        .TextMatrix(.Rows - 1, COLUMNA_ITEM) = "VER DETALLE"
                        .TextMatrix(.Rows - 1, COLUMNA_RECETA) = ""
                        .TextMatrix(.Rows - 1, COLUMNA_UM) = ""
                        .TextMatrix(.Rows - 1, columna_cantidad) = ""
                        .TextMatrix(.Rows - 1, COLUMNA_REFERENCIA) = ""
                    Else
                        .TextMatrix(.Rows - 1, columna_cantidad) = Format(NulosN(RstDet("cantidad")), "0.00")
                        .TextMatrix(.Rows - 1, COLUMNA_ITEM) = NulosC(RstDet("prod"))
                        .TextMatrix(.Rows - 1, COLUMNA_REFERENCIA) = NulosC(RstDet("numregprod"))
                        .TextMatrix(.Rows - 1, COLUMNA_RECETA) = NulosC(RstDet("codrec"))
                        .TextMatrix(.Rows - 1, COLUMNA_UM) = NulosC(RstDet("unid"))
                        .TextMatrix(.Rows - 1, columna_idpro) = NulosN(RstDet("iditem"))
                        .TextMatrix(.Rows - 1, COLUMNA_IDREC) = NulosN(RstDet("idrec"))
                        .TextMatrix(.Rows - 1, COLUMNA_IDUNIMED) = NulosN(RstDet("idunimed"))
                    End If
                End If
            End If
            
            .TextMatrix(.Rows - 1, COLUMNA_LOTE) = NulosC(RstDet("lote"))
            .TextMatrix(.Rows - 1, COLUMNA_NUMORDEN) = Format(NulosN(RstDet("numdoc")), "000000")
            .TextMatrix(.Rows - 1, COLUMNA_IDRESP) = NulosN(RstDet("idresp"))
            .TextMatrix(.Rows - 1, COLUMNAESTADO_) = NulosN(RstDet("idestado"))
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
    If CAMBIOGRABAR_ = -1 Then
        MsgBox "No se puede Cancelar la operación; Grabe los registros para continuar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
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
    
    llenarEstados fg(0), COLUMNAESTADO_
    
    numSolMax = HallaCodigoTabla("pro_ordenproddet", xCon, "numdoc")
    NUMERO_CORRELATIVO = -666
    
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

Function GrabarAlmacen(NUMEROSOL_ As String) As Boolean
    Dim xId As Double
    Dim A As Integer
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim NUMERODOC_ As String
    
On Error GoTo LaCague
    xCon.BeginTrans
    
    cSQL = "SELECT alm_ingreso.id, alm_ingreso.numord " _
        + vbCr + "From alm_ingreso " _
        + vbCr + "WHERE (((alm_ingreso.numord)='" & NUMEROSOL_ & "'));"
    
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then GrabarAlmacen = False: Exit Function
    
    If xRs.RecordCount = 0 Then ' NUEVO
        xId = HallaCodigoTabla("alm_ingreso", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM alm_ingreso", xCon
        RST_Busq RstDet, "SELECT TOP 1 * FROM alm_ingresodet", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else 'MODIFICAR
        xId = NulosN(xRs("id"))
        RST_Busq RstCab, "SELECT * FROM alm_ingreso WHERE id = " & xId, xCon
        xCon.Execute "DELETE * FROM alm_ingresodet WHERE (id = " & xId & ")"
        RST_Busq RstDet, "SELECT * FROM alm_ingresodet", xCon
    End If
    
    
    Dim xRsAux As New ADODB.Recordset
    
    cSQL = "SELECT Max(alm_ingreso.numdoc) AS maxnum " _
        + vbCr + "From alm_ingreso " _
        + vbCr + "GROUP BY alm_ingreso.tipdoc " _
        + vbCr + "HAVING (((alm_ingreso.tipdoc)=110));"
    
    RST_Busq xRsAux, cSQL, xCon
    
    If xRsAux.State = 0 Then GrabarAlmacen = False: Exit Function
    
    If xRsAux.RecordCount = 0 Then
        NUMERODOC_ = 1
    Else
        NUMERODOC_ = NulosN(xRsAux("maxnum")) + 1
    End If
    
    'NUMERODOC_ = HallaCodigoTabla("alm_ingreso", xCon, "numdoc")
    
    mIdRegistro = xId
    RstCab("tipdoc") = 110
    RstCab("fching") = Format(TxtFchPro.valor, "dd/mm/yyyy")
    RstCab("fchdoc") = Format(TxtFchPro.valor, "dd/mm/yyyy")
    RstCab("numser") = "0001"
    RstCab("numdoc") = Format(NUMERODOC_, "0000000000")
    RstCab("idres") = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_IDRESP))
    RstCab("idalm") = 1
    RstCab("nombre") = "PLANEAMIENTO DE PRODUCCION"
    RstCab("tipmov") = 0
    RstCab("idare") = 9
    RstCab("ano") = AnoTra
    RstCab("idmes") = Month(TxtFchPro.valor)
    RstCab("numprod") = NulosC(fg(0).TextMatrix(fg(0).Row, COLUMNA_REFERENCIA))
    RstCab("numord") = NUMEROSOL_
    RstCab.Update
            
    ' Se filtra los Insumos utilizados
    RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_ID)) & ""
    ' Si no hay insumos cargados se los carga de la BD
    If RstValores.RecordCount = 0 Then
        cSQL = "SELECT pro_ordenproddetins.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_ordenproddetins.cantidad, pro_ordenproddetins.activo " _
            + vbCr + "FROM ((pro_ordenproddet RIGHT JOIN pro_ordenproddetins ON pro_ordenproddet.id = pro_ordenproddetins.idorddet) LEFT JOIN alm_inventario ON pro_ordenproddetins.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
            + vbCr + "Where (((pro_ordenproddet.numDoc) = '" & NulosC(fg(0).TextMatrix(fg(0).Row, COLUMNA_NUMORDEN)) & "')) " _
            + vbCr + "GROUP BY pro_ordenproddetins.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_ordenproddetins.cantidad, pro_ordenproddetins.activo;"
        
        RST_Busq Rst, cSQL, xCon
    Else
        DEFINIR_RST_TMP Rst, RstValores
        CARGAR_RST_TMP Rst, RstValores
    End If
    ' Se verifica el estado del recordset
    If Rst.State = 0 Then GrabarAlmacen = False: Exit Function
    
    If Rst.RecordCount = 0 Then
        Set Rst = Nothing

        cSQL = "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]*2050 AS cantidad, -1 AS activo " _
            + vbCr + "FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((pro_recetains.idrec)=" & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNA_IDREC)) & "));"
        
        RST_Busq Rst, cSQL, xCon
    End If
    
    Rst.Filter = "activo = -1"
    Rst.MoveFirst
    For A = 1 To RstValores.RecordCount
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("iditem") = NulosN(Rst("iditem"))
        RstDet("cantteo") = NulosN(Rst("cantidad"))
        RstDet("idtipo") = Busca_Codigo(NulosN(Rst("iditem")), "id", "tippro", "alm_inventario", "N", xCon) 'NulosN(RstValores("tippro"))
        RstDet.Update
        
        Rst.MoveNext
    Next A
    
    xCon.CommitTrans
    GrabarAlmacen = True
    Exit Function

LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    GrabarAlmacen = False
    Exit Function
End Function

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
    
    For A = 1 To fg(0).Rows - 1
        If NulosN(fg(0).TextMatrix(A, COLUMNA_TIPO)) = 1 Then
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
    
    ' Detalle
    Dim CORRORDDET_ As Double
    Dim IDORDDET_ As Double
    
    CORRORDDET_ = HallaCodigoTabla("pro_ordenproddet", xCon, "id")
    
    If QueHace = 1 Then
        ' Obetenemos el Id del registro
        xId = HallaCodigoTabla("pro_ordenprod", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pro_ordenprod", xCon
        RstCab.AddNew
        RstCab("id") = xId
        'procSol = 2 ' La procedencia de la solicitud se pone como manual
    Else
        ' SI SE ESTA MOFIGICANDO UN REGISTRO OBTENEMOS EL ID DEL REGISTRO ACTUAL
        xId = RstOrd("id")
        RST_Busq RstCab, "SELECT * FROM pro_ordenprod WHERE id = " & xId & "", xCon
        ' Eliminamos el detalle
        xCon.Execute "DELETE * FROM pro_ordenproddetins WHERE idord  = " & xId & ""
        xCon.Execute "DELETE * FROM pro_ordenproddet WHERE idord  = " & xId & ""
        'procSol = LblProc
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM pro_ordenproddet", xCon
    RST_Busq RstDetIns, "SELECT TOP 1 * FROM pro_ordenproddetins", xCon
    
    mIdRegistro = xId
    
    RstCab("idsup") = NulosN(TxtIdProg.Text)
    RstCab("fchemi") = CDate(TxtFchPro.valor)
    RstCab.Update
        
    For A = 1 To fg(0).Rows - 1
        ' Se verifica si es registro nuevo o no
        If NulosN(fg(0).TextMatrix(A, COLUMNA_ID)) < 0 Then
            IDORDDET_ = CORRORDDET_
            CORRORDDET_ = CORRORDDET_ + 1
        Else
            IDORDDET_ = NulosN(fg(0).TextMatrix(A, COLUMNA_ID))
        End If
        
        RstDet.AddNew
        RstDet("id") = IDORDDET_
        RstDet("idord") = xId
        RstDet("iditem") = NulosN(fg(0).TextMatrix(A, columna_idpro))
        RstDet("idrec") = NulosN(fg(0).TextMatrix(A, COLUMNA_IDREC))
        RstDet("idunimed") = NulosN(fg(0).TextMatrix(A, COLUMNA_IDUNIMED))
        RstDet("cantidad") = NulosN(fg(0).TextMatrix(A, columna_cantidad))
        RstDet("numdoc") = NulosC(fg(0).TextMatrix(A, COLUMNA_NUMORDEN))
        RstDet("lote") = NulosC(fg(0).TextMatrix(A, COLUMNA_LOTE))
        RstDet("fchprog") = CDate(TxtFchPro.valor)
        RstDet("tipo") = NulosN(fg(0).TextMatrix(A, COLUMNA_TIPO))
        RstDet("idresp") = NulosN(fg(0).TextMatrix(A, COLUMNA_IDRESP))
        RstDet("idprocorr") = NulosN(fg(0).TextMatrix(A, COLUMNA_IDREGPROD))
        RstDet("estado") = NulosN(fg(0).TextMatrix(A, COLUMNAESTADO_))
        RstDet("obs") = ""
        RstDet.Update
        
        ' Detalle de Insumos
        RstValores.Filter = "idorddet = " & NulosN(fg(0).TextMatrix(A, COLUMNA_ID))
        If RstValores.RecordCount <> 0 Then RstValores.MoveFirst
        For B = 1 To RstValores.RecordCount
            RstDetIns.AddNew
            RstDetIns("idord") = xId
            RstDetIns("idorddet") = IDORDDET_
            RstDetIns("activo") = NulosN(RstValores("activo"))
            RstDetIns("iditem") = NulosN(RstValores("iditem"))
            RstDetIns("idlotedet") = NulosN(RstValores("idlotedet"))
            RstDetIns("cantidad") = NulosN(RstValores("cantidad"))
            RstDetIns.Update
            RstValores.MoveNext
            If RstValores.EOF Then Exit For
        Next B
    Next A
            
    'Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
   
    xCon.CommitTrans
    MsgBox "La Solicitud de Materiales se registró con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
    Grabar = True
    CAMBIOGRABAR_ = 0
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
    
    Label5.Caption = "Modificando Solicitud de Materiales"
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    
    xHorIni = Time
    CmdBusSup.SetFocus
    numSolMax = HallaCodigoTabla("pro_ordenproddet", xCon, "numdoc")
    NUMERO_CORRELATIVO = -666
    Frm4.Visible = False
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Dim xRs As New ADODB.Recordset
    
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

Private Sub Form_Unload(Cancel As Integer)
    If CAMBIOGRABAR_ = -1 Then
        MsgBox "No se puede Cancelar la operación; Grabe los registros para continuar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
    End If
End Sub

Private Sub Menu1_1_Click() ' AGREGAR
    cmd_Click 0
End Sub

Private Sub Menu1_2_Click() ' ELIMINAR
    cmd_Click 1
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
        If QueHace = 1 Then Exit Sub
        MuestraSegundoTab
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
                RstOrd.Find "id=" & mIdRegistro
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
        ImprimirLinea
        'ImprimirSolicitud
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
'        If ButtonMenu.Index = 2 Then
'            If TabOne1.CurrTab = 0 Then Exit Sub
'            ImprimirLinea
'        End If
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
