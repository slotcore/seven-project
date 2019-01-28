VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPlanAbastecimiento2 
   Caption         =   "Compras - Plan de Abastecimiento"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame12 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   7395
      Left            =   11100
      TabIndex        =   36
      Top             =   -6630
      Visible         =   0   'False
      Width           =   11985
      Begin VB.CommandButton CmdPrin 
         Height          =   540
         Left            =   10350
         Picture         =   "FrmPlanAbastecimiento2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   6750
         Width           =   735
      End
      Begin VB.CommandButton CmdSalir 
         Height          =   540
         Left            =   11145
         Picture         =   "FrmPlanAbastecimiento2.frx":0B0A
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   6750
         Width           =   735
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg5 
         Height          =   6300
         Left            =   45
         TabIndex        =   38
         Top             =   360
         Width           =   11820
         _cx             =   20849
         _cy             =   11112
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
         BackColor       =   14417405
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   128
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14417405
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
         Rows            =   1
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmPlanAbastecimiento2.frx":0E14
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consolidacion de Insumos"
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
         Left            =   120
         TabIndex        =   37
         Top             =   60
         Width           =   2220
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   45
         Top             =   30
         Width           =   11820
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   0
         X1              =   -30
         X2              =   12000
         Y1              =   7350
         Y2              =   7380
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   11625
         Y1              =   15
         Y2              =   0
      End
   End
   Begin VB.Frame FrmProgreso 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1210
      Left            =   3135
      TabIndex        =   32
      Top             =   3105
      Visible         =   0   'False
      Width           =   5625
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   90
         TabIndex        =   33
         Top             =   850
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label LabelDet 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   105
         TabIndex        =   42
         Top             =   630
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5610
         Y1              =   1195
         Y2              =   1195
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   5610
         X2              =   5610
         Y1              =   15
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   5610
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   1035
      End
      Begin VB.Label LblProcesa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando Datos"
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
         Left            =   105
         TabIndex        =   35
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   105
         TabIndex        =   34
         Top             =   75
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   30
         Top             =   30
         Width           =   5550
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento2.frx":103C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento2.frx":1580
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento2.frx":1704
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento2.frx":1B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento2.frx":1C70
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento2.frx":21B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento2.frx":26F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento2.frx":280C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento2.frx":2920
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento2.frx":2D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanAbastecimiento2.frx":2EE0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar plan de abastecimiento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar plan de abastecimiento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar plan de abastecimiento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Desactivar plan de abastecimiento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Programa de Produccion"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Lista total de insumos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7335
      Left            =   30
      TabIndex        =   4
      Top             =   360
      Width           =   11895
      _cx             =   20981
      _cy             =   12938
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
         Height          =   6915
         Left            =   -12450
         TabIndex        =   13
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6570
            Left            =   30
            TabIndex        =   14
            Top             =   345
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11589
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nº Proyecto"
            Columns(1).DataField=   "id"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripcion"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Ini"
            Columns(3).DataField=   "fchini"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Fin"
            Columns(4).DataField=   "fchfin"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Estado"
            Columns(5).DataField=   "estado"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2381"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2302"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=8202"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=8123"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1826"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1746"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1799"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1720"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2064"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1984"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H80&,.bold=-1"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta Plan de Abastecimiento"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   105
            TabIndex        =   15
            Top             =   30
            Width           =   11595
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6915
         Left            =   45
         TabIndex        =   5
         Top             =   375
         Width           =   11805
         Begin VB.CommandButton CmdProcesar 
            Caption         =   "Procesar"
            Height          =   465
            Left            =   6540
            TabIndex        =   47
            Top             =   450
            Width           =   1290
         End
         Begin VB.CommandButton CmdVerConsolidado 
            Caption         =   "&Ver Req. de Insumos Total"
            Height          =   465
            Left            =   9810
            TabIndex        =   44
            Top             =   420
            Width           =   1890
         End
         Begin VB.CommandButton CmdAdd 
            Caption         =   "Agregar Plan de Produccion"
            Height          =   465
            Left            =   7890
            TabIndex        =   43
            Top             =   420
            Width           =   1890
         End
         Begin VB.Frame Frame15 
            BorderStyle     =   0  'None
            Caption         =   "Frame15"
            Height          =   285
            Left            =   6090
            TabIndex        =   39
            Top             =   6600
            Width           =   5625
            Begin VB.Shape Shape4 
               BackColor       =   &H000000C0&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00800000&
               Height          =   180
               Left            =   3765
               Top             =   45
               Width           =   540
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "= Item sin Stock"
               Height          =   195
               Left            =   4410
               TabIndex        =   41
               Top             =   45
               Width           =   1140
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H00C00000&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00800000&
               Height          =   180
               Left            =   1470
               Top             =   45
               Width           =   540
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "= Item con Stock"
               Height          =   195
               Left            =   2115
               TabIndex        =   40
               Top             =   45
               Width           =   1215
            End
         End
         Begin VB.TextBox TxtDesc 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "TxtDesc"
            Top             =   405
            Width           =   5250
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   5835
            Left            =   30
            TabIndex        =   6
            Top             =   1065
            Width           =   11775
            _cx             =   20770
            _cy             =   10292
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   0
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   0
            Position        =   1
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
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Height          =   5475
               Left            =   15
               TabIndex        =   8
               Top             =   15
               Width           =   11745
               Begin SizerOneLibCtl.ElasticOne Eo1 
                  Height          =   5445
                  Left            =   15
                  TabIndex        =   16
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   9604
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   -1  'True
                  Appearance      =   0
                  MousePointer    =   0
                  _ConvInfo       =   1
                  Version         =   700
                  BackColor       =   -2147483644
                  ForeColor       =   -2147483630
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   ""
                  Align           =   0
                  AutoSizeChildren=   8
                  BorderWidth     =   2
                  ChildSpacing    =   2
                  Splitter        =   0   'False
                  FloodDirection  =   0
                  FloodPercent    =   0
                  CaptionPos      =   1
                  WordWrap        =   -1  'True
                  MaxChildSize    =   0
                  MinChildSize    =   0
                  TagWidth        =   0
                  TagPosition     =   0
                  Style           =   0
                  TagSplit        =   2
                  PicturePos      =   4
                  CaptionStyle    =   0
                  ResizeFonts     =   0   'False
                  GridRows        =   3
                  GridCols        =   1
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"FrmPlanAbastecimiento2.frx":3428
                  Begin VB.Frame Frame8 
                     BorderStyle     =   0  'None
                     Caption         =   "Frame8"
                     Height          =   240
                     Left            =   30
                     TabIndex        =   21
                     Top             =   2685
                     Width           =   11655
                     Begin VB.CommandButton Command4 
                        Height          =   225
                        Left            =   5835
                        Picture         =   "FrmPlanAbastecimiento2.frx":3478
                        Style           =   1  'Graphical
                        TabIndex        =   31
                        Top             =   15
                        Width           =   5790
                     End
                     Begin VB.CommandButton Command1 
                        Height          =   225
                        Left            =   30
                        Picture         =   "FrmPlanAbastecimiento2.frx":35B6
                        Style           =   1  'Graphical
                        TabIndex        =   22
                        Top             =   15
                        Width           =   5790
                     End
                  End
                  Begin VB.Frame Frame7 
                     BorderStyle     =   0  'None
                     Caption         =   "Frame7"
                     Height          =   2460
                     Left            =   30
                     TabIndex        =   19
                     Top             =   2955
                     Width           =   11655
                     Begin VSFlex7Ctl.VSFlexGrid Fg2 
                        Height          =   2430
                        Left            =   30
                        TabIndex        =   20
                        Top             =   15
                        Width           =   11595
                        _cx             =   20452
                        _cy             =   4286
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
                        BackColor       =   14613184
                        ForeColor       =   -2147483640
                        BackColorFixed  =   -2147483633
                        ForeColorFixed  =   -2147483630
                        BackColorSel    =   128
                        ForeColorSel    =   -2147483634
                        BackColorBkg    =   -2147483636
                        BackColorAlternate=   14613184
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
                        Rows            =   1
                        Cols            =   19
                        FixedRows       =   1
                        FixedCols       =   1
                        RowHeightMin    =   0
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   0   'False
                        FormatString    =   $"FrmPlanAbastecimiento2.frx":36F4
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
                  End
                  Begin VB.Frame Frame6 
                     BorderStyle     =   0  'None
                     Height          =   2625
                     Left            =   30
                     TabIndex        =   17
                     Top             =   30
                     Width           =   11655
                     Begin VSFlex7Ctl.VSFlexGrid Fg1 
                        Height          =   2595
                        Left            =   30
                        TabIndex        =   18
                        Top             =   15
                        Width           =   11595
                        _cx             =   20452
                        _cy             =   4577
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
                        BackColor       =   14417405
                        ForeColor       =   -2147483640
                        BackColorFixed  =   -2147483633
                        ForeColorFixed  =   -2147483630
                        BackColorSel    =   128
                        ForeColorSel    =   -2147483634
                        BackColorBkg    =   -2147483636
                        BackColorAlternate=   14417405
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
                        Rows            =   1
                        Cols            =   19
                        FixedRows       =   1
                        FixedCols       =   1
                        RowHeightMin    =   0
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   0   'False
                        FormatString    =   $"FrmPlanAbastecimiento2.frx":3936
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
                  End
               End
            End
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Height          =   5475
               Left            =   12390
               TabIndex        =   7
               Top             =   15
               Width           =   11745
               Begin SizerOneLibCtl.ElasticOne Eo2 
                  Height          =   5445
                  Left            =   15
                  TabIndex        =   23
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   9604
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   -1  'True
                  Appearance      =   0
                  MousePointer    =   0
                  _ConvInfo       =   1
                  Version         =   700
                  BackColor       =   -2147483644
                  ForeColor       =   -2147483630
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   ""
                  Align           =   0
                  AutoSizeChildren=   8
                  BorderWidth     =   2
                  ChildSpacing    =   2
                  Splitter        =   0   'False
                  FloodDirection  =   0
                  FloodPercent    =   0
                  CaptionPos      =   1
                  WordWrap        =   -1  'True
                  MaxChildSize    =   0
                  MinChildSize    =   0
                  TagWidth        =   0
                  TagPosition     =   0
                  Style           =   0
                  TagSplit        =   2
                  PicturePos      =   4
                  CaptionStyle    =   0
                  ResizeFonts     =   0   'False
                  GridRows        =   3
                  GridCols        =   1
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"FrmPlanAbastecimiento2.frx":3B78
                  Begin VB.Frame Frame11 
                     BorderStyle     =   0  'None
                     Height          =   2625
                     Left            =   30
                     TabIndex        =   28
                     Top             =   30
                     Width           =   11655
                     Begin VSFlex7Ctl.VSFlexGrid Fg3 
                        Height          =   2595
                        Left            =   30
                        TabIndex        =   29
                        Top             =   15
                        Width           =   11595
                        _cx             =   20452
                        _cy             =   4577
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
                        BackColor       =   14417405
                        ForeColor       =   -2147483640
                        BackColorFixed  =   -2147483633
                        ForeColorFixed  =   -2147483630
                        BackColorSel    =   128
                        ForeColorSel    =   -2147483634
                        BackColorBkg    =   -2147483636
                        BackColorAlternate=   14417405
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
                        Rows            =   1
                        Cols            =   19
                        FixedRows       =   1
                        FixedCols       =   1
                        RowHeightMin    =   0
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   0   'False
                        FormatString    =   $"FrmPlanAbastecimiento2.frx":3BC8
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
                  End
                  Begin VB.Frame Frame10 
                     BackColor       =   &H00FFC0C0&
                     BorderStyle     =   0  'None
                     Caption         =   "Frame7"
                     Height          =   2460
                     Left            =   30
                     TabIndex        =   26
                     Top             =   2955
                     Width           =   11655
                     Begin VSFlex7Ctl.VSFlexGrid Fg4 
                        Height          =   2430
                        Left            =   30
                        TabIndex        =   27
                        Top             =   15
                        Width           =   11595
                        _cx             =   20452
                        _cy             =   4286
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
                        BackColor       =   14613184
                        ForeColor       =   -2147483640
                        BackColorFixed  =   -2147483633
                        ForeColorFixed  =   -2147483630
                        BackColorSel    =   128
                        ForeColorSel    =   -2147483634
                        BackColorBkg    =   -2147483636
                        BackColorAlternate=   14613184
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
                        Rows            =   1
                        Cols            =   19
                        FixedRows       =   1
                        FixedCols       =   1
                        RowHeightMin    =   0
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   0   'False
                        FormatString    =   $"FrmPlanAbastecimiento2.frx":3E0A
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
                  End
                  Begin VB.Frame Frame9 
                     BorderStyle     =   0  'None
                     Caption         =   "Frame8"
                     Height          =   240
                     Left            =   30
                     TabIndex        =   24
                     Top             =   2685
                     Width           =   11655
                     Begin VB.CommandButton Command3 
                        Height          =   225
                        Left            =   5835
                        Picture         =   "FrmPlanAbastecimiento2.frx":404C
                        Style           =   1  'Graphical
                        TabIndex        =   30
                        Top             =   15
                        Width           =   5790
                     End
                     Begin VB.CommandButton Command2 
                        Height          =   225
                        Left            =   30
                        Picture         =   "FrmPlanAbastecimiento2.frx":418A
                        Style           =   1  'Graphical
                        TabIndex        =   25
                        Top             =   15
                        Width           =   5790
                     End
                  End
               End
            End
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
            Height          =   300
            Left            =   1155
            TabIndex        =   1
            Top             =   720
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
            Locked          =   -1  'True
            Valor           =   "06/02/2006"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
            Height          =   300
            Left            =   5070
            TabIndex        =   2
            Top             =   720
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
            Locked          =   -1  'True
            Valor           =   "06/02/2006"
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fch. Termino"
            Height          =   195
            Left            =   3900
            TabIndex        =   12
            Top             =   750
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle Plan de Abastecimiento"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   11610
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   60
            TabIndex        =   10
            Top             =   450
            Width           =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fch. Inicio"
            Height          =   195
            Left            =   60
            TabIndex        =   9
            Top             =   750
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "FrmPlanAbastecimiento2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMPLANABASTECIMIENTO
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA EL INGRESO Y EDICION DEL PLAN DE ABASTECIMIENTO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstInsumos As New ADODB.Recordset
Dim RstPlanAbas As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim xHorIni As Date                 'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xMesInicio As Integer


Private Sub iniciarCampos()
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.ExplorerBar = flexExSortShowAndMove
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ForeColorSel = &H80000005
    Fg1.BackColorSel = &H80&
    
    Fg2.AllowUserResizing = flexResizeColumns
    Fg2.AutoSearch = flexSearchFromTop
    Fg2.ExplorerBar = flexExSortShowAndMove
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.ForeColorSel = &H80000005
    Fg2.BackColorSel = &H80&
    
    Fg3.AllowUserResizing = flexResizeColumns
    Fg3.AutoSearch = flexSearchFromTop
    Fg3.ExplorerBar = flexExSortShowAndMove
    Fg3.SelectionMode = flexSelectionByRow
    Fg3.ForeColorSel = &H80000005
    Fg3.BackColorSel = &H80&
    
    Fg4.AllowUserResizing = flexResizeColumns
    Fg4.AutoSearch = flexSearchFromTop
    Fg4.ExplorerBar = flexExSortShowAndMove
    Fg4.SelectionMode = flexSelectionByRow
    Fg4.ForeColorSel = &H80000005
    Fg4.BackColorSel = &H80&
    
    Fg5.AllowUserResizing = flexResizeColumns
    Fg5.AutoSearch = flexSearchFromTop
    Fg5.ExplorerBar = flexExSortShowAndMove
    Fg5.SelectionMode = flexSelectionByRow
    Fg5.ForeColorSel = &H80000005
    Fg5.BackColorSel = &H80&
End Sub

Private Sub configurarVista()
    Frame12.Top = 330
    Frame12.Left = 0
    Frame12.Height = 7340
    Frame12.Width = 11900
End Sub

Private Function calcularIndicador(fchIni As String, fchFin As String) As Integer
    Dim indicador As Integer
    Dim idMesIni As Integer
    Dim idMesFin As Integer
    Dim idAñoIni As Integer
    Dim idAñoFin As Integer
    
    idMesIni = NulosN(Format(fchIni, "m"))
    idMesFin = NulosN(Format(fchFin, "m"))
    idAñoIni = NulosN(Format(fchIni, "yyyy"))
    idAñoFin = NulosN(Format(fchFin, "yyyy"))
    
    If idMesIni <> 0 And idAñoIni <> 0 Then
        If idAñoFin > idAñoIni Then
            indicador = (13 - idMesIni) + idMesFin
        Else
            indicador = idMesFin - idMesIni + 1
        End If
        
        If indicador > 12 Then indicador = 12
    End If
    
    calcularIndicador = indicador
End Function

Private Sub rellenarValores(ByRef xRst As ADODB.Recordset, ByRef fgx As VSFlexGrid, Fini As String, fFin As String)
    Dim A, B, xCol As Integer
    Dim xStock As Double
    Dim indicador As Integer
    Dim xMesIni As Integer
    Dim xMesAux As Integer
    
    indicador = calcularIndicador(Fini, fFin)
    xMesIni = Format(Fini, "m")
    FrmProgreso.Visible = True
    ProgressBar1.Max = xRst.RecordCount
    With fgx
        .Rows = 1
        'xCol = 0
        For A = 1 To xRst.RecordCount
            xCol = 0
            ProgressBar1.Value = A
            .Rows = .Rows + 1
            .TextMatrix(A, 0) = NulosC(xRst("codpro"))
            .TextMatrix(A, 1) = NulosC(xRst("descripcion"))
            LabelDet = NulosC(xRst("descripcion"))
            .TextMatrix(A, 2) = NulosC(xRst("codigo"))
            .TextMatrix(A, 3) = NulosC(xRst("abrev"))
            
            xMesAux = xMesIni
            For B = 1 To indicador
                xCol = B + 3
                .TextMatrix(A, xCol) = Format(NulosN(xRst("" & xMesAux & "")), FORMAT_MONTO)
                xMesAux = xMesAux + 1
                If xMesAux > 12 Then xMesAux = 1
            Next B
            xCol = xCol + 1
            .TextMatrix(A, xCol) = Format(NulosN(xRst("tot")), FORMAT_MONTO)
            xCol = xCol + 1
            xStock = NulosN(Busca_Codigo(NulosC(xRst("codpro")), "id", "stckact", "alm_inventario", "N", xCon))
            .TextMatrix(A, xCol) = Format(xStock, "0.00")
            xCol = xCol + 1
            .TextMatrix(A, xCol) = Format((NulosN(xRst("tot")) - xStock), "0.00")
            
            'MATIZAMOS LOS COLORES
            If (xStock - (NulosN(xRst("tot")))) < 0 Then
                'mostramos en rojo lo que falta producir
                .Select A, xCol, A, xCol: .FillStyle = flexFillRepeat: .CellForeColor = &HFF&
            Else
                'mostramos en rojo lo que se produccio de mas
                .TextMatrix(A, xCol) = Abs(NulosN(.TextMatrix(A, xCol)))
                .Select A, xCol, A, xCol: .FillStyle = flexFillRepeat: .CellForeColor = &HFF0000
            End If
            
            xRst.MoveNext
        Next A
        Set xRst = Nothing
    End With
    
    fgx.FillStyle = flexFillRepeat
    If fgx.Rows = 1 Then fgx.Rows = fgx.Rows + 1
    fgx.Select 1, fgx.Cols - 3, fgx.Rows - 1, fgx.Cols - 1
    fgx.CellBackColor = &HFEFBEB
    fgx.Select 1, 1, 1, 1
    
    LabelDet = ""
    FrmProgreso.Visible = False
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL PLAN DE ABASTECIMIENTO ACTUAL
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Bloquea
    TxtDesc.Text = RstPlanAbas("descripcion")
    TxtFchIni.Valor = Format(RstPlanAbas("fchini"), "dd/mm/yyyy")
    TxtFchFin.Valor = Format(RstPlanAbas("fchfin"), "dd/mm/yyyy")
    
    Dim Rst As New ADODB.Recordset
    Dim cSQL As String
    
    'MOSTRAMOS LOS PRODUCTOS FINALES
    
    cSQL = "TRANSFORM First(ges_planabapropro.cantidad) AS PrimeroDecantidad " _
        + vbCr + "SELECT ges_planabapropro.idpv, ges_planabapropro.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev, ges_planabapropro.tipo, Sum(ges_planabapropro.cantidad) AS tot " _
        + vbCr + "FROM (ges_planabapropro LEFT JOIN alm_inventario ON ges_planabapropro.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "Where (((ges_planabapropro.idpv)=" & RstPlanAbas("id") & ") AND ((ges_planabapropro.tipo) = 1)) " _
        + vbCr + "GROUP BY ges_planabapropro.idpv, ges_planabapropro.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ges_planabapropro.tipo " _
        + vbCr + "PIVOT ges_planabapropro.idmes;"
    
    RST_Busq Rst, cSQL, xCon
    LblProcesa = "Procesando Productos Finales"
    configurarGrid Fg1, TxtFchIni.Valor, TxtFchFin.Valor
    rellenarValores Rst, Fg1, TxtFchIni.Valor, TxtFchFin.Valor
    
    
    'MOSTRAMOS LOS PRODUCTOS INTERMEDIOS
    
    cSQL = "TRANSFORM First(ges_planabapropro.cantidad) AS PrimeroDecantidad " _
        + vbCr + "SELECT ges_planabapropro.idpv, ges_planabapropro.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev, ges_planabapropro.tipo, Sum(ges_planabapropro.cantidad) AS tot " _
        + vbCr + "FROM (ges_planabapropro LEFT JOIN alm_inventario ON ges_planabapropro.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "Where (((ges_planabapropro.idpv)=" & RstPlanAbas("id") & ") AND ((ges_planabapropro.tipo) = 2)) " _
        + vbCr + "GROUP BY ges_planabapropro.idpv, ges_planabapropro.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ges_planabapropro.tipo " _
        + vbCr + "PIVOT ges_planabapropro.idmes;"
    
    RST_Busq Rst, cSQL, xCon
    LblProcesa = "Procesando Productos Intermedios"
    configurarGrid Fg3, TxtFchIni.Valor, TxtFchFin.Valor
    rellenarValores Rst, Fg3, TxtFchIni.Valor, TxtFchFin.Valor
    
    'MOSTRAMOS LOS INSUMOS FINALES
    
    cSQL = "TRANSFORM First(ges_planabadet.cantidad) AS PrimeroDecantidad " _
        + vbCr + "SELECT ges_planabadet.idpv, ges_planabadet.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev, ges_planabadet.tipo, Sum(ges_planabadet.cantidad) AS tot " _
        + vbCr + "FROM (ges_planabadet LEFT JOIN alm_inventario ON ges_planabadet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "Where (((ges_planabadet.idpv)=" & RstPlanAbas("id") & ") AND ((ges_planabadet.tipo) = 1)) " _
        + vbCr + "GROUP BY ges_planabadet.idpv, ges_planabadet.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ges_planabadet.tipo " _
        + vbCr + "PIVOT ges_planabadet.idmes;"
    
    RST_Busq Rst, cSQL, xCon
    LblProcesa = "Procesando Insumos Finales"
    configurarGrid Fg2, TxtFchIni.Valor, TxtFchFin.Valor
    rellenarValores Rst, Fg2, TxtFchIni.Valor, TxtFchFin.Valor
'
    'MOSTRAMOS LOS INSUMOS PARA LOS PRODUCTOS INTERMEDIOS
    
    cSQL = "TRANSFORM First(ges_planabadet.cantidad) AS PrimeroDecantidad " _
        + vbCr + "SELECT ges_planabadet.idpv, ges_planabadet.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev, ges_planabadet.tipo, Sum(ges_planabadet.cantidad) AS tot " _
        + vbCr + "FROM (ges_planabadet LEFT JOIN alm_inventario ON ges_planabadet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "Where (((ges_planabadet.idpv)=" & RstPlanAbas("id") & ") AND ((ges_planabadet.tipo) = 2)) " _
        + vbCr + "GROUP BY ges_planabadet.idpv, ges_planabadet.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ges_planabadet.tipo " _
        + vbCr + "PIVOT ges_planabadet.idmes;"
    
    RST_Busq Rst, cSQL, xCon
    LblProcesa = "Procesando Insumos Intermedios"
    configurarGrid Fg4, TxtFchIni.Valor, TxtFchFin.Valor
    rellenarValores Rst, Fg4, TxtFchIni.Valor, TxtFchFin.Valor
    
    TabOne2.CurrTab = 0
End Sub

Sub configurarGrid(fgx As VSFlexGrid, fchIni As String, fchFin As String)
    Dim Rst As New ADODB.Recordset
    Dim idMesIni As Integer
    Dim idAñoIni As Integer
    Dim A As Integer
    Dim xMes As Integer
    Dim xAño As Integer
    Dim indicador As Integer
    
    xMes = CInt(Mid(Format(fchIni, "dd/mm/yyyy"), 4, 2))
    xAño = CInt(Mid(Format(fchIni, "dd/mm/yyyy"), 7, 4))
    
    xMesInicio = xMes
    indicador = calcularIndicador(fchIni, fchFin)
    
    fgx.Cols = 7 + indicador
    
    fgx.TextMatrix(0, 0) = "Id"
    fgx.ColWidth(0) = 0
    fgx.TextMatrix(0, 1) = "Producto"
    fgx.ColWidth(1) = 4800
    fgx.TextMatrix(0, 2) = "CodPro"
    fgx.ColWidth(2) = 0
    fgx.ColAlignment(2) = flexAlignLeftCenter
    fgx.TextMatrix(0, 3) = "Unidad"
    
    
    For A = 1 To indicador
        RST_Busq Rst, "SELECT DISTINCT con_meses.id, con_meses.descripcion " _
                    & "FROM con_meses " _
                    & "WHERE (((con_meses.id)=" & xMes & "))", xCon
        
        fgx.TextMatrix(0, A + 3) = Rst("descripcion") & " " & xAño
        fgx.ColWidth(A + 3) = 1250
        xMes = xMes + 1
        If xMes > 12 Then xMes = 1: xAño = xAño + 1
    Next A
    
    fgx.TextMatrix(0, A + 3) = "Programado"
    fgx.ColWidth(A + 3) = 1250
    fgx.TextMatrix(0, A + 4) = "Stock"
    fgx.TextMatrix(0, A + 5) = "Diferencia"
    
    If QueHace <> 3 Then fgx.ColWidth(A + 3) = 0: fgx.ColWidth(A + 4) = 0: fgx.ColWidth(A + 5) = 0
    
'    Fgx.FillStyle = flexFillRepeat
'    If Fgx.Rows = 1 Then Fgx.Rows = Fgx.Rows + 1
'    Fgx.Select 1, Fgx.Cols - 3, Fgx.Rows - 1, Fgx.Cols - 1
'    Fgx.CellBackColor = &HFEFBEB
'    Fgx.Select 1, 1, 1, 1
    fgx.FrozenCols = 3
    
    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : PintarTotales
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PINTa EN COLORES LOS TOTALES DE LOS CONTROLES Fg1, Fg2, Fg3,Fg4
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
'Sub PintarTotales()
'    With Fg1
'        .Select 1, Fg1.Cols - 3, Fg1.Rows - 1, Fg1.Cols - 1: .FillStyle = flexFillRepeat: .CellBackColor = &HFEFBEB
'        .Select 1, 1, 1, 1
'    End With
'
'    With Fg2
'        .Select 1, Fg2.Cols - 3, Fg2.Rows - 1, Fg2.Cols - 1: .FillStyle = flexFillRepeat: .CellBackColor = &HFEFBEB
'        .Select 1, 1, 1, 1
'    End With
'
'    With Fg3
'        .Select 1, Fg3.Cols - 3, Fg3.Rows - 1, Fg3.Cols - 1: .FillStyle = flexFillRepeat: .CellBackColor = &HFEFBEB
'        .Select 1, 1, 1, 1
'    End With
'
'    With Fg4
'        .Select 1, Fg4.Cols - 3, Fg4.Rows - 1, Fg4.Cols - 1: .FillStyle = flexFillRepeat: .CellBackColor = &HFEFBEB
'        .Select 1, 1, 1, 1
'    End With
'End Sub

Private Sub CmdAdd_Click()
    ' PERMITE AGREGAR UN PLAN DE PRODUCCION AL PLAN DE ABASTECIMIENTO QUE SE CREANDO O EDITANDO
    PreparaRST
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCodItem As String
    Dim cSQL As String
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":            xCampos(1, 2) = "2000":    xCampos(1, 3) = "N"

    cSQL = "SELECT ges_plaprod.id, ges_plaprod.descripcion, ges_plaprod.fchini, ges_plaprod.fchfin " _
        + vbCr + "From ges_plaprod " _
        + vbCr + "ORDER BY ges_plaprod.descripcion;"
    xform.SQLCad = cSQL
    
    xform.Titulo = "Buscando Plan de Produccion"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        Dim xId As Integer
        xId = xRs("id")
        TxtFchIni.Valor = xRs("fchini")
        TxtFchFin.Valor = xRs("fchfin")
        Set xform = Nothing
        Set xRs = Nothing
            
        MostrarDetallePlanProduccion xId
        MostrarInsumosProducto
        'PintarTotales
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

'*****************************************************************************************************
'* Nombre Archivo   : MostrarInsumosProducto
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LOS INSUMOS DEL PRODUCTO QUE SE ESTA CONSULTANDO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub MostrarInsumosProducto()
    Dim A, B, C As Integer
    Dim RstRec As New ADODB.Recordset
    Dim xTotal, xStock As Double
    Dim indicador As Integer
    Dim xMesIni As String
    
    'Procesamos insumos para los productos finales
    FrmProgreso.Visible = True
    LblProcesa.Caption = "Analizando Insumos para Productos Terminados"
    ProgressBar1.Max = Fg1.Rows - 1
    
    configurarGrid Fg2, TxtFchIni.Valor, TxtFchFin.Valor
    
    For A = 1 To Fg1.Rows - 1
        ProgressBar1.Value = A
        FrmProgreso.Refresh
        
        RST_Busq RstRec, "SELECT DISTINCT pro_receta.prirec, pro_recetains.iditem, alm_inventario.codpro, alm_inventario.descripcion, pro_recetains.canpro, alm_inventario.tippro, mae_unidades.abrev " _
            + vbCr + "FROM pro_receta RIGHT JOIN (alm_inventario RIGHT JOIN (pro_recetains LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id) ON alm_inventario.id = pro_recetains.iditem) ON pro_receta.id = pro_recetains.idrec " _
            + vbCr + "WHERE (((pro_receta.prirec)=1) AND ((pro_receta.iditem)=" & NulosN(Fg1.TextMatrix(A, 0)) & "))", xCon
        
        If RstRec.RecordCount <> 0 Then
            RstRec.MoveFirst
            indicador = calcularIndicador(TxtFchIni.Valor, TxtFchFin.Valor)
            
            For B = 1 To RstRec.RecordCount
                RstInsumos.Filter = adFilterNone
                RstInsumos.Filter = "idpro = " & NulosN(RstRec("iditem")) & ""
                If RstInsumos.RecordCount = 0 Then
                    RstInsumos.AddNew
                End If
                RstInsumos("idpro") = NulosN(RstRec("iditem"))
                RstInsumos("cod_item") = RstRec("codpro")
                RstInsumos("unimed") = RstRec("abrev")
                RstInsumos("descripcion") = RstRec("descripcion")
                LabelDet = RstRec("descripcion")
                
                xMesIni = Format(NulosC(TxtFchIni.Valor), "m")
                
                For C = 1 To indicador
                    RstInsumos("" & xMesIni & "") = RstInsumos("" & xMesIni & "") + (RstRec("canpro") * NulosN(Fg1.TextMatrix(A, C + 3)))
                    xMesIni = xMesIni + 1
                    If xMesIni > 12 Then xMesIni = 1
                Next C
                    
                RstRec.MoveNext
                If RstRec.EOF = True Then
                    Exit For
                End If
            Next B
        Else
            MsgBox "El producto " & NulosC(Fg1.TextMatrix(A, 1)) & " no tiene una receta asignada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            'Blanquea
            'Exit Sub
        End If
    Next A
    
    xTotal = 0
    RstInsumos.Filter = adFilterNone
    RstInsumos.Sort = "descripcion"
    RstInsumos.MoveFirst
    
    LblProcesa.Caption = "Procesando Insumos para Productos Terminados"
    ProgressBar1.Max = RstInsumos.RecordCount
    
    For A = 1 To RstInsumos.RecordCount
        ProgressBar1.Value = A
        FrmProgreso.Refresh
    
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(A, 0) = RstInsumos("idpro")
        Fg2.TextMatrix(A, 1) = RstInsumos("descripcion")
        LabelDet = RstInsumos("descripcion")
        Fg2.TextMatrix(A, 2) = RstInsumos("cod_item")
        Fg2.TextMatrix(A, 3) = RstInsumos("unimed")
        
        xMesIni = Format(NulosC(TxtFchIni.Valor), "m")
                
        For C = 1 To indicador
            Fg2.TextMatrix(A, C + 3) = Format(RstInsumos("" & xMesIni & ""), "0.0000")
            xTotal = xTotal + NulosN(RstInsumos("" & xMesIni & ""))
            'RstInsumos("" & xMesIni & "") = RstInsumos("" & xMesIni & "") + (RstRec("canpro") * Val(Fg1.TextMatrix(A, C + 3)))
            xMesIni = xMesIni + 1
            If xMesIni > 12 Then xMesIni = 1
        Next C
'
        Fg2.TextMatrix(A, Fg1.Cols - 3) = Format(xTotal, "0.00")
        
        xStock = 0
        xStock = Busca_Codigo(RstInsumos("idpro"), "id", "stckact", "alm_inventario", "N", xCon)
        Fg2.TextMatrix(A, Fg2.Cols - 2) = Format(xStock, "0.00")
        Fg2.TextMatrix(A, Fg2.Cols - 1) = Format((xTotal - xStock), "0.00")
        
        With Fg2
            If (xStock - xTotal) < 0 Then
                'mostramos en rojo lo que falta producir
                .Select A, Fg2.Cols - 1, A, Fg2.Cols - 1: .FillStyle = flexFillRepeat: .CellForeColor = &HFF&
            Else
                'mostramos en azul lo que se produccio de mas
                Fg2.TextMatrix(A, Fg2.Cols - 1) = Abs(NulosN(Fg2.TextMatrix(A, Fg2.Cols - 1)))
                .Select A, Fg2.Cols - 1, A, Fg2.Cols - 1: .FillStyle = flexFillRepeat: .CellForeColor = &HFF0000
            End If
        End With
        
        RstInsumos.MoveNext
        
        If RstInsumos.EOF = True Then
            Exit For
        End If
    Next A
    '-------------------------------------------------
    'Procesamos insumos para los productos Intermedios
    Set RstInsumos = Nothing
    PreparaRST
    
    LblProcesa.Caption = "Procesando Insumos para Productos Intermedios"
    ProgressBar1.Max = Fg3.Rows - 1
    
    configurarGrid Fg4, TxtFchIni.Valor, TxtFchFin.Valor
    
    For A = 1 To Fg3.Rows - 1
        ProgressBar1.Value = A
        FrmProgreso.Refresh
        
        RST_Busq RstRec, "SELECT DISTINCT pro_receta.prirec, pro_recetains.iditem, alm_inventario.codpro, alm_inventario.descripcion, pro_recetains.canpro, " _
            & " alm_inventario.tippro, mae_unidades.abrev FROM (alm_inventario RIGHT JOIN (pro_receta LEFT JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) " _
            & " ON alm_inventario.id = pro_recetains.iditem) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
            & " WHERE (((pro_receta.iditem)=" & Fg3.TextMatrix(A, 0) & ") AND ((pro_receta.prirec)=1) AND ((alm_inventario.tippro)<>3))", xCon
        
        If RstRec.RecordCount <> 0 Then
            RstRec.MoveFirst
            For B = 1 To RstRec.RecordCount
                RstInsumos.Filter = adFilterNone
                RstInsumos.Filter = "idpro = '" & RstRec("iditem") & "'"
                If RstInsumos.RecordCount = 0 Then
                    RstInsumos.AddNew
                End If
                RstInsumos("idpro") = RstRec("iditem")
                RstInsumos("cod_item") = RstRec("codpro")
                RstInsumos("unimed") = RstRec("abrev")
                RstInsumos("descripcion") = RstRec("descripcion")
                LabelDet = RstRec("descripcion")
                
                xMesIni = Format(NulosC(TxtFchIni.Valor), "m")
                
                For C = 1 To indicador
                    RstInsumos("" & xMesIni & "") = RstInsumos("" & xMesIni & "") + (RstRec("canpro") * NulosN(Fg3.TextMatrix(A, C + 3)))
                    xMesIni = xMesIni + 1
                    If xMesIni > 12 Then xMesIni = 1
                Next C
                    
                RstRec.MoveNext
                If RstRec.EOF = True Then
                    Exit For
                End If
            Next B
        End If
    Next A
    
    xTotal = 0
    RstInsumos.Filter = adFilterNone
    RstInsumos.Sort = "descripcion"
    RstInsumos.MoveFirst
    
    LblProcesa.Caption = "Procesando Insumos para Productos Intermedios"
    ProgressBar1.Max = RstInsumos.RecordCount
    
    For A = 1 To RstInsumos.RecordCount
        ProgressBar1.Value = A
        FrmProgreso.Refresh
    
        Fg4.Rows = Fg4.Rows + 1
        Fg4.TextMatrix(A, 0) = RstInsumos("idpro")
        Fg4.TextMatrix(A, 1) = RstInsumos("descripcion")
        LabelDet = RstInsumos("descripcion")
        Fg4.TextMatrix(A, 2) = RstInsumos("cod_item")
        Fg4.TextMatrix(A, 3) = RstInsumos("unimed")
        
        xMesIni = Format(NulosC(TxtFchIni.Valor), "m")
        
        For C = 1 To indicador
            Fg4.TextMatrix(A, C + 3) = Format(RstInsumos("" & xMesIni & ""), "0.0000")
            xTotal = xTotal + NulosN(RstInsumos("" & xMesIni & ""))
            xMesIni = xMesIni + 1
            If xMesIni > 12 Then xMesIni = 1
        Next C
        
        Fg4.TextMatrix(A, Fg1.Cols - 3) = Format(xTotal, "0.00")
        
        xStock = 0
        xStock = Busca_Codigo(RstInsumos("idpro"), "id", "stckact", "alm_inventario", "N", xCon)
        Fg4.TextMatrix(A, Fg2.Cols - 2) = Format(xStock, "0.00")
        Fg4.TextMatrix(A, Fg2.Cols - 1) = Format((xTotal - xStock), "0.00")
        
        RstInsumos.MoveNext
        
        If RstInsumos.EOF = True Then
            Exit For
        End If
    Next A
    
    FrmProgreso.Visible = False
    LabelDet = ""
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : MostrarDetallePlanProduccion
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DEL PLAN DE PRODUCCION SELECCIONADO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    xId          |  Integer   |  ESPECIFICA EL ID DEL PLAN DE PRODUCCION
'* DEVUELVE         :
'*****************************************************************************************************
Sub MostrarDetallePlanProduccion(xId As Integer)
    Dim RstProd As New ADODB.Recordset
    Dim RstDeta As New ADODB.Recordset
    Dim A, B, xCol As Integer
    Dim xTotal, xStock As Double
    Dim cSQL As String
    
    FrmProgreso.Left = 3135
    FrmProgreso.Top = 3105
    LblProcesa.Caption = "Mostrando Productos"
    FrmProgreso.Visible = True
    
    'mostramos los productos finales a producir
    
    cSQL = "TRANSFORM First(ges_plaproddet.cantidad) AS PrimeroDecantidad " _
        + vbCr + "SELECT ges_plaproddet.idpv, ges_plaproddet.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev, Sum(ges_plaproddet.cantidad) AS tot " _
        + vbCr + "FROM (ges_plaproddet LEFT JOIN alm_inventario ON ges_plaproddet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "Where (((ges_plaproddet.idpv) = " & xId & ")) " _
        + vbCr + "GROUP BY ges_plaproddet.idpv, ges_plaproddet.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev " _
        + vbCr + "PIVOT ges_plaproddet.idmes;"
        
    RST_Busq RstProd, cSQL, xCon
    configurarGrid Fg1, TxtFchIni.Valor, TxtFchFin.Valor
    rellenarValores RstProd, Fg1, TxtFchIni.Valor, TxtFchFin.Valor
    

    LblProcesa.Caption = "Mostrando Productos Intermedios"
    
    cSQL = "TRANSFORM First(ges_plaproddet2.cantidad) AS PrimeroDecantidad " _
        + vbCr + "SELECT ges_plaproddet2.idpv, ges_plaproddet2.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev, Sum(ges_plaproddet2.cantidad) AS tot " _
        + vbCr + "FROM (ges_plaproddet2 LEFT JOIN alm_inventario ON ges_plaproddet2.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "Where (((ges_plaproddet2.idpv) = " & xId & ")) " _
        + vbCr + "GROUP BY ges_plaproddet2.idpv, ges_plaproddet2.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev " _
        + vbCr + "PIVOT ges_plaproddet2.idmes;"
        
    RST_Busq RstProd, cSQL, xCon
    configurarGrid Fg3, TxtFchIni.Valor, TxtFchFin.Valor
    rellenarValores RstProd, Fg3, TxtFchIni.Valor, TxtFchFin.Valor

    FrmProgreso.Visible = False
End Sub

'Private Sub CmdMas_Click()
'    ' CAMBIA EL ANCHO DE LAS COLUMNAS DE  LOS CONTROLS Fg1, Fg2, Fg3, Fg4, INCREMENTANDO EL ANCHO EN 10 PIXEL
'    Fg1.ColWidth(1) = Fg1.ColWidth(1) + 100
'    Fg1.ColWidth(4) = Fg1.ColWidth(4) + 10
'    Fg1.ColWidth(5) = Fg1.ColWidth(5) + 10
'    Fg1.ColWidth(6) = Fg1.ColWidth(6) + 10
'    Fg1.ColWidth(7) = Fg1.ColWidth(7) + 10
'    Fg1.ColWidth(8) = Fg1.ColWidth(8) + 10
'    Fg1.ColWidth(9) = Fg1.ColWidth(9) + 10
'    Fg1.ColWidth(10) = Fg1.ColWidth(10) + 10
'    Fg1.ColWidth(11) = Fg1.ColWidth(11) + 10
'    Fg1.ColWidth(12) = Fg1.ColWidth(12) + 10
'    Fg1.ColWidth(13) = Fg1.ColWidth(13) + 10
'    Fg1.ColWidth(14) = Fg1.ColWidth(14) + 10
'    Fg1.ColWidth(15) = Fg1.ColWidth(15) + 10
'
'    Fg2.ColWidth(1) = Fg2.ColWidth(1) + 100
'    Fg2.ColWidth(4) = Fg2.ColWidth(4) + 10
'    Fg2.ColWidth(5) = Fg2.ColWidth(5) + 10
'    Fg2.ColWidth(6) = Fg2.ColWidth(6) + 10
'    Fg2.ColWidth(7) = Fg2.ColWidth(7) + 10
'    Fg2.ColWidth(8) = Fg2.ColWidth(8) + 10
'    Fg2.ColWidth(9) = Fg2.ColWidth(9) + 10
'    Fg2.ColWidth(10) = Fg2.ColWidth(10) + 10
'    Fg2.ColWidth(11) = Fg2.ColWidth(11) + 10
'    Fg2.ColWidth(12) = Fg2.ColWidth(12) + 10
'    Fg2.ColWidth(13) = Fg2.ColWidth(13) + 10
'    Fg2.ColWidth(14) = Fg2.ColWidth(14) + 10
'    Fg2.ColWidth(15) = Fg2.ColWidth(15) + 10
'
'    Fg3.ColWidth(1) = Fg3.ColWidth(1) + 100
'    Fg3.ColWidth(4) = Fg3.ColWidth(4) + 10
'    Fg3.ColWidth(5) = Fg3.ColWidth(5) + 10
'    Fg3.ColWidth(6) = Fg3.ColWidth(6) + 10
'    Fg3.ColWidth(7) = Fg3.ColWidth(7) + 10
'    Fg3.ColWidth(8) = Fg3.ColWidth(8) + 10
'    Fg3.ColWidth(9) = Fg3.ColWidth(9) + 10
'    Fg3.ColWidth(10) = Fg3.ColWidth(10) + 10
'    Fg3.ColWidth(11) = Fg3.ColWidth(11) + 10
'    Fg3.ColWidth(12) = Fg3.ColWidth(12) + 10
'    Fg3.ColWidth(13) = Fg3.ColWidth(13) + 10
'    Fg3.ColWidth(14) = Fg3.ColWidth(14) + 10
'    Fg3.ColWidth(15) = Fg3.ColWidth(15) + 10
'
'    Fg4.ColWidth(1) = Fg4.ColWidth(1) + 100
'    Fg4.ColWidth(4) = Fg4.ColWidth(4) + 10
'    Fg4.ColWidth(5) = Fg4.ColWidth(5) + 10
'    Fg4.ColWidth(6) = Fg4.ColWidth(6) + 10
'    Fg4.ColWidth(7) = Fg4.ColWidth(7) + 10
'    Fg4.ColWidth(8) = Fg4.ColWidth(8) + 10
'    Fg4.ColWidth(9) = Fg4.ColWidth(9) + 10
'    Fg4.ColWidth(10) = Fg4.ColWidth(10) + 10
'    Fg4.ColWidth(11) = Fg4.ColWidth(11) + 10
'    Fg4.ColWidth(12) = Fg4.ColWidth(12) + 10
'    Fg4.ColWidth(13) = Fg4.ColWidth(13) + 10
'    Fg4.ColWidth(14) = Fg4.ColWidth(14) + 10
'    Fg4.ColWidth(15) = Fg4.ColWidth(15) + 10
'End Sub

Private Sub CmdPrin_Click()
    ExportarExcel
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : ExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA LOS DATOS DEL CONTROL Fg7 A MS EXCEL
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub ExportarExcel()
    If Fg5.Rows = 1 Then
        MsgBox "No se ha procesado registros para el consolidados de insumos", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
        Exit Sub
    End If
    
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    Dim xCad As String
    Dim objExcel As Object
    
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    objExcel.WindowState = 2
    objExcel.Workbooks.Add
   
    With objExcel.ActiveSheet
        .Cells(1, 2) = NomEmp
        .Cells(1, 10) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        .Cells(3, 2) = "Consilidado de Insumos y Materia Prima"
        
        xFilas = 5
        For A = 0 To Fg5.Rows - 1
            For B = 1 To Fg5.Cols - 1
                If A = 0 Then
                    .Cells(xFilas, B + 1) = "'" + Fg5.TextMatrix(A, B)
                Else
                    If B <= 4 Then
                        .Cells(xFilas, B + 1) = "'" + Fg5.TextMatrix(A, B)
                    Else
                        .Cells(xFilas, B + 1) = Val(Fg5.TextMatrix(A, B))
                    End If
                End If
            Next B
            xFilas = xFilas + 1
        Next A
    End With
    
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
End Sub

'Private Sub CmdRes_Click()
'    ' CAMBIA EL ANCHO DE LAS COLUMNAS DE  LOS CONTROLS Fg1, Fg2, Fg3, Fg4, DECREMENTANDO EL ANCHO EN 10 PIXEL
'    Fg1.ColWidth(1) = Fg1.ColWidth(1) - 100
'    Fg1.ColWidth(4) = Fg1.ColWidth(4) - 10
'    Fg1.ColWidth(5) = Fg1.ColWidth(5) - 10
'    Fg1.ColWidth(6) = Fg1.ColWidth(6) - 10
'    Fg1.ColWidth(7) = Fg1.ColWidth(7) - 10
'    Fg1.ColWidth(8) = Fg1.ColWidth(8) - 10
'    Fg1.ColWidth(9) = Fg1.ColWidth(9) - 10
'    Fg1.ColWidth(10) = Fg1.ColWidth(10) - 10
'    Fg1.ColWidth(11) = Fg1.ColWidth(11) - 10
'    Fg1.ColWidth(12) = Fg1.ColWidth(12) - 10
'    Fg1.ColWidth(13) = Fg1.ColWidth(13) - 10
'    Fg1.ColWidth(14) = Fg1.ColWidth(14) - 10
'    Fg1.ColWidth(15) = Fg1.ColWidth(15) - 10
'
'    Fg2.ColWidth(1) = Fg2.ColWidth(1) - 100
'    Fg2.ColWidth(4) = Fg2.ColWidth(4) - 10
'    Fg2.ColWidth(5) = Fg2.ColWidth(5) - 10
'    Fg2.ColWidth(6) = Fg2.ColWidth(6) - 10
'    Fg2.ColWidth(7) = Fg2.ColWidth(7) - 10
'    Fg2.ColWidth(8) = Fg2.ColWidth(8) - 10
'    Fg2.ColWidth(9) = Fg2.ColWidth(9) - 10
'    Fg2.ColWidth(10) = Fg2.ColWidth(10) - 10
'    Fg2.ColWidth(11) = Fg2.ColWidth(11) - 10
'    Fg2.ColWidth(12) = Fg2.ColWidth(12) - 10
'    Fg2.ColWidth(13) = Fg2.ColWidth(13) - 10
'    Fg2.ColWidth(14) = Fg2.ColWidth(14) - 10
'    Fg2.ColWidth(15) = Fg2.ColWidth(15) - 10
'
'    Fg3.ColWidth(1) = Fg3.ColWidth(1) - 100
'    Fg3.ColWidth(4) = Fg3.ColWidth(4) - 10
'    Fg3.ColWidth(5) = Fg3.ColWidth(5) - 10
'    Fg3.ColWidth(6) = Fg3.ColWidth(6) - 10
'    Fg3.ColWidth(7) = Fg3.ColWidth(7) - 10
'    Fg3.ColWidth(8) = Fg3.ColWidth(8) - 10
'    Fg3.ColWidth(9) = Fg3.ColWidth(9) - 10
'    Fg3.ColWidth(10) = Fg3.ColWidth(10) - 10
'    Fg3.ColWidth(11) = Fg3.ColWidth(11) - 10
'    Fg3.ColWidth(12) = Fg3.ColWidth(12) - 10
'    Fg3.ColWidth(13) = Fg3.ColWidth(13) - 10
'    Fg3.ColWidth(14) = Fg3.ColWidth(14) - 10
'    Fg3.ColWidth(15) = Fg3.ColWidth(15) - 10
'
'    Fg4.ColWidth(1) = Fg4.ColWidth(1) - 100
'    Fg4.ColWidth(4) = Fg4.ColWidth(4) - 10
'    Fg4.ColWidth(5) = Fg4.ColWidth(5) - 10
'    Fg4.ColWidth(6) = Fg4.ColWidth(6) - 10
'    Fg4.ColWidth(7) = Fg4.ColWidth(7) - 10
'    Fg4.ColWidth(8) = Fg4.ColWidth(8) - 10
'    Fg4.ColWidth(9) = Fg4.ColWidth(9) - 10
'    Fg4.ColWidth(10) = Fg4.ColWidth(10) - 10
'    Fg4.ColWidth(11) = Fg4.ColWidth(11) - 10
'    Fg4.ColWidth(12) = Fg4.ColWidth(12) - 10
'    Fg4.ColWidth(13) = Fg4.ColWidth(13) - 10
'    Fg4.ColWidth(14) = Fg4.ColWidth(14) - 10
'    Fg4.ColWidth(15) = Fg4.ColWidth(15) - 10
'End Sub

Private Sub CmdSalir_Click()
    Toolbar1.Enabled = Not Toolbar1.Enabled
    Frame12.Visible = False
End Sub

Private Sub CmdVerConsolidado_Click()
    'MUESTRA EL REQUERIMIENTO TOTAL DE INSUMOS PARA EL PLAN DE ABASTECIMIENTO
    Dim A, B As Integer
    Dim Total, xStock As Double
    Dim indicador, xCol As Integer
    Dim xMesIni, xMesAux As String
    
    indicador = calcularIndicador(TxtFchIni.Valor, TxtFchFin.Valor)
    xMesIni = Format(TxtFchIni.Valor, "m")
    
    configurarGrid Fg5, TxtFchIni.Valor, TxtFchFin.Valor
    
    Fg5.Rows = 1
    PreparaRST
    For A = 1 To Fg2.Rows - 1
        RstInsumos.AddNew
        RstInsumos("descripcion") = Fg2.TextMatrix(A, 1)
        RstInsumos("cod_item") = Fg2.TextMatrix(A, 0)
        RstInsumos("unimed") = Fg2.TextMatrix(A, 3)
        
        xMesAux = xMesIni
        xCol = 0
        For B = 1 To indicador
            xCol = B + 3
            RstInsumos("" & xMesAux & "") = Fg2.TextMatrix(A, xCol)
            xMesAux = xMesAux + 1
            If xMesAux > 12 Then xMesAux = 1
        Next B
    Next A
    
    For A = 1 To Fg4.Rows - 1
        RstInsumos.MoveFirst
        RstInsumos.Filter = "cod_item = '" & Fg4.TextMatrix(A, 2) & "'"
        If RstInsumos.RecordCount = 0 Then
            RstInsumos.AddNew
            RstInsumos("descripcion") = Fg4.TextMatrix(A, 1)
            RstInsumos("cod_item") = Fg4.TextMatrix(A, 0)
            RstInsumos("unimed") = Fg4.TextMatrix(A, 3)
            
            xMesAux = xMesIni
            xCol = 0
            For B = 1 To indicador
                xCol = B + 3
                RstInsumos("" & xMesAux & "") = Fg4.TextMatrix(A, xCol)
                xMesAux = xMesAux + 1
                If xMesAux > 12 Then xMesAux = 1
            Next B
        Else
            If RstInsumos.RecordCount = 1 Then
                
                xMesAux = xMesIni
                xCol = 0
                For B = 1 To indicador
                    xCol = B + 3
                    RstInsumos("" & xMesAux & "") = RstInsumos("" & xMesAux & "") + Fg4.TextMatrix(A, xCol)
                    xMesAux = xMesAux + 1
                    If xMesAux > 12 Then xMesAux = 1
                Next B
            Else
                'este error nunca debe de ocurrir
                MsgBox "Hay mas de un items con el mismo codigo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
        End If
    Next A
    
    RstInsumos.Filter = adFilterNone
    RstInsumos.Sort = "descripcion"
    RstInsumos.MoveFirst
    For A = 1 To RstInsumos.RecordCount
        Fg5.Rows = Fg5.Rows + 1
        Fg5.TextMatrix(A, 1) = RstInsumos("descripcion")
        Fg5.TextMatrix(A, 2) = RstInsumos("cod_item")
        Fg5.TextMatrix(A, 3) = RstInsumos("unimed")
        
        xMesAux = xMesIni
        xCol = 0
        Total = 0
        For B = 1 To indicador
            xCol = B + 3
            Fg5.TextMatrix(A, xCol) = Format(RstInsumos("" & xMesAux & ""), "0.00")
            Total = Total + RstInsumos("" & xMesAux & "")
            xMesAux = xMesAux + 1
            If xMesAux > 12 Then xMesAux = 1
        Next B
        
        xCol = xCol + 1
        Fg5.TextMatrix(A, xCol) = Format(Total, "0.00")
        
        xCol = xCol + 1
        xStock = NulosN(Busca_Codigo(NulosC(RstInsumos("cod_item")), "id", "stckact", "alm_inventario", "N", xCon))
        Fg5.TextMatrix(A, xCol) = Format(xStock, "0.00")
        xCol = xCol + 1
        Fg5.TextMatrix(A, xCol) = Format((NulosN(Total) - xStock), "0.00")
        
        RstInsumos.MoveNext
        If RstInsumos.EOF = True Then
            Exit For
        End If
    Next A
    
    Toolbar1.Enabled = Not Toolbar1.Enabled
    Frame12.Left = 135
    Frame12.Top = 795
    configurarVista
    Frame12.Visible = True
End Sub

Private Sub Command1_Click()
    If Eo1.GRID(gsRowHeight, 0) <= 2670 Then
        Eo1.GRID(gsRowHeight, 0) = 3600
        Eo1.GRID(gsRowHeight, 2) = 0
        
        Fg1.Height = 5100
        Fg2.Height = 0
        On Error Resume Next
        Command1.Picture = LoadPicture(Trim(App.Path) + "\bmps\" + "flechaup.bmp")
        Err.Clear
        Command4.Visible = False
    Else
        Eo1.GRID(gsRowHeight, 0) = 2670
        Eo1.GRID(gsRowHeight, 1) = 270
        Eo1.GRID(gsRowHeight, 2) = 2490
    
        Eo1.GRID(gsRowHeight, 0) = 2670
        Eo1.GRID(gsRowHeight, 1) = 260
        Eo1.GRID(gsRowHeight, 2) = 2490
    
        Fg1.Height = 2595
        Fg2.Height = 2430
        On Error Resume Next
        Command1.Picture = LoadPicture(Trim(App.Path) + "\bmps\" + "flechadown.bmp")
        Err.Clear
        Command4.Visible = True
    End If
    Fg1.SetFocus
End Sub

Private Sub Command2_Click()
    If Eo2.GRID(gsRowHeight, 0) <= 2670 Then
        Eo2.GRID(gsRowHeight, 0) = 3600
        Eo2.GRID(gsRowHeight, 2) = 0
        
        Fg3.Height = 5100
        Fg4.Height = 0
        On Error Resume Next
        Command2.Picture = LoadPicture(Trim(App.Path) + "\bmps\" + "flechaup.bmp")
        Err.Clear
        Command3.Visible = False
    Else
        Eo2.GRID(gsRowHeight, 0) = 2670
        Eo2.GRID(gsRowHeight, 1) = 270
        Eo2.GRID(gsRowHeight, 2) = 2490
    
        Eo2.GRID(gsRowHeight, 0) = 2670
        Eo2.GRID(gsRowHeight, 1) = 260
        Eo2.GRID(gsRowHeight, 2) = 2490
    
        Fg3.Height = 2595
        Fg4.Height = 2430
        On Error Resume Next
        Command2.Picture = LoadPicture(Trim(App.Path) + "\bmps\" + "flechadown.bmp")
        Err.Clear
        Command3.Visible = True
    End If
    Fg3.SetFocus

End Sub

Private Sub Command3_Click()
    If Eo2.GRID(gsRowHeight, 2) = 2490 Then
        Eo2.GRID(gsRowHeight, 0) = 0
        Fg3.Height = 0
        Fg4.Height = 5000
        Eo2.GRID(gsRowHeight, 1) = 270
        On Error Resume Next
        Command3.Picture = LoadPicture(Trim(App.Path) + "\bmps\" + "flechadown.bmp")
        Err.Clear
        Command2.Visible = False
    
    Else
        Fg3.Height = 2595
        Fg4.Height = 2430
        Eo2.GRID(gsRowHeight, 0) = 2670
        Eo2.GRID(gsRowHeight, 1) = 270
        Eo2.GRID(gsRowHeight, 2) = 2490
        
        Eo2.GRID(gsRowHeight, 0) = 2670
        Eo2.GRID(gsRowHeight, 1) = 270
        Eo2.GRID(gsRowHeight, 2) = 2490
        On Error Resume Next
        Command3.Picture = LoadPicture(Trim(App.Path) + "\bmps\" + "flechaup.bmp")
        Err.Clear
        Command2.Visible = True
    End If
    Fg3.SetFocus
End Sub

Private Sub Command4_Click()
    If Eo1.GRID(gsRowHeight, 2) = 2490 Then
        Eo1.GRID(gsRowHeight, 0) = 0
        Fg1.Height = 0
        Fg2.Height = 5000
        Eo1.GRID(gsRowHeight, 1) = 270
        On Error Resume Next
        Command4.Picture = LoadPicture(Trim(App.Path) + "\bmps\" + "flechadown.bmp")
        Err.Clear
        Command1.Visible = False
    
    Else
        Fg1.Height = 2595
        Fg2.Height = 2430
        Eo1.GRID(gsRowHeight, 0) = 2670
        Eo1.GRID(gsRowHeight, 1) = 270
        Eo1.GRID(gsRowHeight, 2) = 2490
        
        Eo1.GRID(gsRowHeight, 0) = 2670
        Eo1.GRID(gsRowHeight, 1) = 270
        Eo1.GRID(gsRowHeight, 2) = 2490
        On Error Resume Next
        Command4.Picture = LoadPicture(Trim(App.Path) + "\bmps\" + "flechaup.bmp")
        Err.Clear
        Command1.Visible = True
    End If
    Fg1.SetFocus
End Sub

Private Sub CmdProcesar_Click()
    configurarGrid Fg1, TxtFchIni.Valor, TxtFchFin.Valor
    configurarGrid Fg2, TxtFchIni.Valor, TxtFchFin.Valor
    configurarGrid Fg3, TxtFchIni.Valor, TxtFchFin.Valor
    configurarGrid Fg4, TxtFchIni.Valor, TxtFchFin.Valor
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstPlanAbas("id")), xCon
    End If
End Sub

'Private Sub Command5_Click()
'    ' DECREMENTA EN 10 PIXEL EL ANCHO DE LAS COLUMNAS DEL CONTROL Fg5
'    Fg5.ColWidth(1) = Fg5.ColWidth(1) - 100
'    Fg5.ColWidth(4) = Fg5.ColWidth(4) - 10
'    Fg5.ColWidth(5) = Fg5.ColWidth(5) - 10
'    Fg5.ColWidth(6) = Fg5.ColWidth(6) - 10
'    Fg5.ColWidth(7) = Fg5.ColWidth(7) - 10
'    Fg5.ColWidth(8) = Fg5.ColWidth(8) - 10
'    Fg5.ColWidth(9) = Fg5.ColWidth(9) - 10
'    Fg5.ColWidth(10) = Fg5.ColWidth(10) - 10
'    Fg5.ColWidth(11) = Fg5.ColWidth(11) - 10
'    Fg5.ColWidth(12) = Fg5.ColWidth(12) - 10
'    Fg5.ColWidth(13) = Fg5.ColWidth(13) - 10
'    Fg5.ColWidth(14) = Fg5.ColWidth(14) - 10
'    Fg5.ColWidth(15) = Fg5.ColWidth(15) - 10
'
'End Sub

'Private Sub Command6_Click()
'    ' INCREMENTA EN 10 PIXEL EL ANCHO DE LAS COLUMNAS DEL CONTROL Fg5
'    Fg5.ColWidth(1) = Fg5.ColWidth(1) + 100
'    Fg5.ColWidth(4) = Fg5.ColWidth(4) + 10
'    Fg5.ColWidth(5) = Fg5.ColWidth(5) + 10
'    Fg5.ColWidth(6) = Fg5.ColWidth(6) + 10
'    Fg5.ColWidth(7) = Fg5.ColWidth(7) + 10
'    Fg5.ColWidth(8) = Fg5.ColWidth(8) + 10
'    Fg5.ColWidth(9) = Fg5.ColWidth(9) + 10
'    Fg5.ColWidth(10) = Fg5.ColWidth(10) + 10
'    Fg5.ColWidth(11) = Fg5.ColWidth(11) + 10
'    Fg5.ColWidth(12) = Fg5.ColWidth(12) + 10
'    Fg5.ColWidth(13) = Fg5.ColWidth(13) + 10
'    Fg5.ColWidth(14) = Fg5.ColWidth(14) + 10
'    Fg5.ColWidth(15) = Fg5.ColWidth(15) + 10
'End Sub

Private Sub Form_Activate()
'Modificado: 08/01/11 Johan Castro
'            Agregar linea de codigo para bloquear accesos de usuarios

    ' SEGUNDO EVENTO A EJECUTARSE DESPUES DE CARGARSE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        '----------------------------------------------
        
        RST_Busq RstPlanAbas, "SELECT IIf([ges_planaba]![activo]=0,'No Activo','Activo') AS estado, * " _
            & " From ges_planaba ORDER BY ges_planaba.id", xCon
        
        Set Dg1.DataSource = RstPlanAbas

    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    QueHace = 3
    iniciarCampos
    TabOne1.CurrTab = 0
End Sub

Private Sub Modificar()
    QueHace = 2
    xHorIni = Time
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label1.Caption = "Modificando Plan de Abastecimiento"
    Bloquea
    PreparaRST
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    Fg3.Editable = flexEDKbdMouse
    Fg4.Editable = flexEDKbdMouse
    
    TxtDesc.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    xHorIni = Time

    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label1.Caption = "Agregando Plan de Abastecimiento"
    Bloquea
    Blanquea
    PreparaRST
    Fg1.Rows = 1
    Fg2.Rows = 1
    Fg3.Rows = 1
    Fg4.Rows = 1
    
    TxtDesc.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TEXTBOX DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Bloquea()
'    TxtDesc.Locked = Not TxtDesc.Locked
''    TxtFchIni.Locked = Not TxtFchIni.Locked
''    TxtFchFin.Locked = Not TxtFchFin.Locked
'    CmdVerConsolidado.Enabled = Not CmdVerConsolidado.Enabled
    If QueHace <> 3 Then TxtDesc.Locked = False Else TxtDesc.Locked = True
    If QueHace <> 3 Then TxtFchIni.Locked = False Else TxtFchIni.Locked = True
    If QueHace <> 3 Then TxtFchFin.Locked = False Else TxtFchFin.Locked = True
    If QueHace <> 3 Then CmdAdd.Visible = True Else CmdAdd.Visible = False
    If QueHace <> 3 Then CmdProcesar.Visible = True Else CmdProcesar.Visible = False
    If QueHace <> 3 Then CmdVerConsolidado.Visible = False Else CmdVerConsolidado.Visible = True
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TEXTBOX PARA EL INGRESO DE UN REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Blanquea()
    TxtDesc.Text = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : PreparaRST
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA UN RECORDSET TEMPORAL
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub PreparaRST()
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(18, 3) As String

    xCampos(0, 0) = "cod_item":     xCampos(0, 1) = "C":      xCampos(0, 2) = "20"
    xCampos(1, 0) = "unimed":       xCampos(1, 1) = "C":      xCampos(1, 2) = "4"
    xCampos(2, 0) = "descripcion":  xCampos(2, 1) = "C":      xCampos(2, 2) = "200"
    
    xCampos(3, 0) = "1":          xCampos(3, 1) = "N":      xCampos(3, 2) = "2"
    xCampos(4, 0) = "2":          xCampos(4, 1) = "N":      xCampos(4, 2) = "2"
    xCampos(5, 0) = "3":          xCampos(5, 1) = "N":      xCampos(5, 2) = "2"
    xCampos(6, 0) = "4":          xCampos(6, 1) = "N":      xCampos(6, 2) = "2"
    xCampos(7, 0) = "5":          xCampos(7, 1) = "N":      xCampos(7, 2) = "2"
    xCampos(8, 0) = "6":          xCampos(8, 1) = "N":      xCampos(8, 2) = "2"
    xCampos(9, 0) = "7":          xCampos(9, 1) = "N":      xCampos(9, 2) = "2"
    xCampos(10, 0) = "8":         xCampos(10, 1) = "N":      xCampos(10, 2) = "2"
    xCampos(11, 0) = "9":         xCampos(11, 1) = "N":      xCampos(11, 2) = "2"
    xCampos(12, 0) = "10":         xCampos(12, 1) = "N":      xCampos(12, 2) = "2"
    xCampos(13, 0) = "11":         xCampos(13, 1) = "N":      xCampos(13, 2) = "2"
    xCampos(14, 0) = "12":         xCampos(14, 1) = "N":      xCampos(14, 2) = "2"
    
    xCampos(15, 0) = "ope":         xCampos(15, 1) = "N":      xCampos(15, 2) = "2"
    xCampos(16, 0) = "idpro":       xCampos(16, 1) = "N":      xCampos(16, 2) = "2"
    xCampos(17, 0) = "tippro":       xCampos(17, 1) = "C":      xCampos(17, 2) = "2"
    Set RstInsumos = xFun.CrearRstTMP(xCampos)
    RstInsumos.Open
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA ges_plaaba, ESTA FUNCION DEVUELVER VERDADERO
'*                    CUANDO TIENE EXITO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    If NulosC(TxtDesc.Text) = "" Then
        MsgBox "No ha especificado la descripcion del producto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDesc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Function
    End If
    
    On Error GoTo LaCague
    xCon.BeginTrans
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDet2 As New ADODB.Recordset
    Dim RstFue As New ADODB.Recordset
    Dim xId As Double
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT TOP 1 * FROM ges_planaba", xCon
        
        xId = HallaCodigoTabla("ges_planaba", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xId
    Else
        
        xId = RstPlanAbas("id")
        
        RST_Busq RstCab, "SELECT * FROM ges_planaba WHERE id=" & xId & " ", xCon
        xCon.Execute "DELETE * FROM ges_planabadet WHERE idpv = " & xId & ""
        xCon.Execute "DELETE * FROM ges_planabapropro WHERE idpv = " & xId & ""
        
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM ges_planabadet", xCon
    RST_Busq RstDet2, "SELECT TOP 1 * FROM ges_planabapropro", xCon
    
    RstCab("descripcion") = NulosC(TxtDesc.Text)
    RstCab("fchini") = NulosC(TxtFchIni.Valor)
    RstCab("fchfin") = NulosC(TxtFchFin.Valor)
    RstCab.Update
    
    Dim xFila, xCol, xMes As Integer
    
    'guardamos los insumos calculados
    'insumos para productos finales
    For xFila = 1 To Fg2.Rows - 1
        xMes = xMesInicio
        For xCol = 4 To Fg2.Cols - 4
            RstDet.AddNew
            RstDet("idpv") = xId
            RstDet("codpro") = Trim(Fg2.TextMatrix(xFila, 0))
            RstDet("idmes") = xMes
            RstDet("cantidad") = NulosN(Fg2.TextMatrix(xFila, xCol))
            RstDet("tipo") = 1
            RstDet.Update
            xMes = xMes + 1
            
            If xMes = 13 Then
                xMes = 1
            End If
        Next xCol
    Next xFila
    
    'insumos para productos intermedios
    For xFila = 1 To Fg4.Rows - 1
        xMes = xMesInicio
        For xCol = 4 To Fg4.Cols - 4
            RstDet.AddNew
            RstDet("idpv") = xId
            RstDet("codpro") = Trim(Fg4.TextMatrix(xFila, 0))
            RstDet("idmes") = xMes
            RstDet("cantidad") = NulosN(Fg4.TextMatrix(xFila, xCol))
            RstDet("tipo") = 2
            RstDet.Update
            xMes = xMes + 1
            
            If xMes = 13 Then
                xMes = 1
            End If
        Next xCol
    Next xFila
    
    'grabamos los productos del plan de produccion
    'productos finales
    For xFila = 1 To Fg1.Rows - 1
        xMes = xMesInicio
        For xCol = 4 To Fg1.Cols - 4
            RstDet2.AddNew
            RstDet2("idpv") = xId
            RstDet2("codpro") = Trim(Fg1.TextMatrix(xFila, 0))
            RstDet2("idmes") = xMes
            RstDet2("cantidad") = NulosN(Fg1.TextMatrix(xFila, xCol))
            RstDet2("tipo") = 1
            RstDet2.Update
            xMes = xMes + 1
            
            If xMes = 13 Then
                xMes = 1
            End If
        Next xCol
    Next xFila
    
    'productos intermedios
    For xFila = 1 To Fg3.Rows - 1
        xMes = xMesInicio
        For xCol = 4 To Fg3.Cols - 4
            RstDet2.AddNew
            RstDet2("idpv") = xId
            RstDet2("codpro") = Trim(Fg3.TextMatrix(xFila, 0))
            RstDet2("idmes") = xMes
            RstDet2("cantidad") = Fg3.TextMatrix(xFila, xCol)
            RstDet2("tipo") = 2
            RstDet2.Update
            xMes = xMes + 1
            
            If xMes = 13 Then
                xMes = 1
            End If
        Next xCol
    Next xFila
       
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
       
    xCon.CommitTrans
    MsgBox "El plan de abastecimiento se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function

LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
    
End Function

'*****************************************************************************************************
'* Nombre Archivo   : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE INGRESAR O MODIFICAR UN REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    ActivaTool
    TabOne1.TabEnabled(0) = True
    Bloquea
    Label1.Caption = "Detalle Plan de Abastecimiento"
    TabOne1.CurrTab = 0
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA ges_planaba
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Rpta = MsgBox("¿Esta seguro de eliminar el plan de abastecimiento seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM ges_planabapropro WHERE idpv =" & RstPlanAbas("id") & ""
        xCon.Execute "DELETE * FROM ges_planabadet WHERE idpv =" & RstPlanAbas("id") & ""
        xCon.Execute "DELETE * FROM ges_planaba WHERE id =" & RstPlanAbas("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstPlanAbas("id") & " AND idform = " & IdMenuActivo
        
        MsgBox "El plan de abastecimiento se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstPlanAbas.Requery
        Dg1.Refresh

    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            MuestraSegundoTab
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CambiarEstado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CAMBIA EL ESTADO DE UN REGISTRO DE LA TABLA ges_planaba
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Activado     |  Boolean   |  INDICA SI SE ACTIVA O DESACTIVA UN REGISTRO
'* DEVUELVE         :
'*****************************************************************************************************
Sub CambiarEstado(Activado As Boolean)
    Dim Rpta As Integer
    TabOne1.CurrTab = 0
    If Activado = False Then
        Rpta = MsgBox("Esta seguro de desactivar el plan de abastecimiento seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    Else
        Rpta = MsgBox("Esta seguro de activar el plan de abastecimiento seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    End If
    
    If Rpta = vbYes Then
        If Activado = False Then
            xCon.Execute "UPDATE ges_planaba SET ges_planaba.activo = 0 Where (((ges_planaba.id) = " & RstPlanAbas("id") & "))"
            MsgBox "El plan de abastecimiento se desactivo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Else
            xCon.Execute "UPDATE ges_planaba SET ges_planaba.activo = -1 Where (((ges_planaba.id) = " & RstPlanAbas("id") & "))"
            MsgBox "El plan de abastecimiento se activo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    RstPlanAbas.Requery
    Dg1.Refresh
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstPlanAbas.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 14 Then
        Set RstPlanAbas = Nothing
        Unload Me
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        'If ButtonMenu.Index = 1 Then Modificar
        If ButtonMenu.Index = 2 Then CambiarEstado True
    End If
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then Eliminar
        If ButtonMenu.Index = 2 Then CambiarEstado False
    End If
End Sub

Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
