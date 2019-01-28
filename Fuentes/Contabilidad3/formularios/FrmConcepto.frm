VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmConcepto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Concepto"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8235
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":1EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":23EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":2706
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":2B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":2FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConcepto.frx":32C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
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
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   30
      TabIndex        =   8
      Top             =   360
      Width           =   11745
      _cx             =   20717
      _cy             =   12726
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   12632256
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   12632256
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "   Consulta   |   Detalles   "
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
         Height          =   6795
         Left            =   45
         TabIndex        =   11
         Top             =   375
         Width           =   11655
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6405
            Left            =   45
            TabIndex        =   12
            Top             =   390
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   11298
            _LayoutType     =   4
            _RowHeight      =   14
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Concepto"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Variable"
            Columns(2).DataField=   "variable"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Categoría"
            Columns(3).DataField=   "catnombre"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fórmula"
            Columns(4).DataField=   "formula"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=2"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=7117"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=7038"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1931"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1852"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2461"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2381"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(24)=   "Column(4).Width=7805"
            Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=7726"
            Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=38,.bgcolor=&HDBFDFD&,.fgcolor=&H0&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.fgcolor=&HFF0000&,.bold=0"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(25)  =   ":id=13,.strikethrough=0,.charset=0"
            _StyleDefs(26)  =   ":id=13,.fontname=MS Sans Serif"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.namedParent=33,.fgcolor=&H800000&"
            _StyleDefs(29)  =   ":id=14,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(30)  =   ":id=14,.fontname=MS Sans Serif"
            _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&,.bold=0"
            _StyleDefs(34)  =   ":id=18,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(35)  =   ":id=18,.fontname=MS Sans Serif"
            _StyleDefs(36)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(37)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(38)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(39)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(40)  =   ":id=21,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(41)  =   ":id=21,.fontname=MS Sans Serif"
            _StyleDefs(42)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(43)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Conceptos"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   105
            TabIndex        =   13
            Top             =   60
            Width           =   11400
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6795
         Left            =   12390
         TabIndex        =   9
         Top             =   375
         Width           =   11655
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6315
            Left            =   -90
            TabIndex        =   14
            Top             =   420
            Width           =   11700
            _cx             =   20637
            _cy             =   11139
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
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
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "  Datos Principales  |   Listado de Cuenta   "
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
            Begin VB.Frame Frame5 
               BorderStyle     =   0  'None
               Height          =   5895
               Left            =   12345
               TabIndex        =   27
               Top             =   45
               Width           =   11610
               Begin VB.Frame fra 
                  Height          =   5175
                  Index           =   1
                  Left            =   9480
                  TabIndex        =   29
                  Top             =   540
                  Width           =   1830
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Eliminar"
                     Height          =   675
                     Index           =   4
                     Left            =   240
                     TabIndex        =   32
                     ToolTipText     =   "Eliminar Cuenta"
                     Top             =   2250
                     Width           =   1260
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "Agregar"
                     Height          =   675
                     Index           =   2
                     Left            =   270
                     TabIndex        =   31
                     ToolTipText     =   "Agregar Cuenta Contable"
                     Top             =   510
                     Width           =   1260
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "Seleccionar"
                     Height          =   675
                     Index           =   3
                     Left            =   270
                     TabIndex        =   30
                     ToolTipText     =   "Muestra Ctas utilizadas en libro diario"
                     Top             =   1320
                     Width           =   1260
                  End
               End
               Begin VB.Frame fra 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame12"
                  Height          =   405
                  Index           =   0
                  Left            =   120
                  TabIndex        =   28
                  Top             =   45
                  Width           =   11340
                  Begin VB.Line Line2 
                     BorderColor     =   &H00FFFFFF&
                     BorderWidth     =   2
                     Index           =   0
                     X1              =   15
                     X2              =   15
                     Y1              =   0
                     Y2              =   380
                  End
                  Begin VB.Line Line1 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   0
                     X1              =   11325
                     X2              =   11325
                     Y1              =   15
                     Y2              =   395
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H00808080&
                     BorderWidth     =   2
                     Index           =   1
                     X1              =   -30
                     X2              =   12000
                     Y1              =   390
                     Y2              =   390
                  End
                  Begin VB.Line lin 
                     BorderColor     =   &H80000009&
                     BorderWidth     =   2
                     Index           =   0
                     X1              =   -15
                     X2              =   12000
                     Y1              =   15
                     Y2              =   15
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   5175
                  Left            =   150
                  TabIndex        =   33
                  Top             =   570
                  Width           =   9195
                  _cx             =   16219
                  _cy             =   9128
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
                  Rows            =   2
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConcepto.frx":3656
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
                  Ellipsis        =   1
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
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   5895
               Left            =   45
               TabIndex        =   15
               Top             =   45
               Width           =   11610
               Begin VB.Frame FraDestinoCta 
                  Caption         =   "Seleccionar"
                  Enabled         =   0   'False
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
                  Height          =   645
                  Left            =   6180
                  TabIndex        =   35
                  Top             =   2010
                  Visible         =   0   'False
                  Width           =   4080
                  Begin VB.OptionButton OptPasivo 
                     Caption         =   "Pasivo y Patrimonio"
                     Height          =   285
                     Left            =   1890
                     TabIndex        =   37
                     Top             =   270
                     Width           =   1785
                  End
                  Begin VB.OptionButton OptActivo 
                     Caption         =   "Activo"
                     Height          =   285
                     Left            =   120
                     TabIndex        =   36
                     Top             =   270
                     Width           =   855
                  End
               End
               Begin VB.TextBox TxtId 
                  BackColor       =   &H0000FF00&
                  Height          =   285
                  Left            =   6990
                  Locked          =   -1  'True
                  TabIndex        =   34
                  Text            =   "TxtId"
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.TextBox TxtNombreCorto 
                  Height          =   315
                  Left            =   1425
                  Locked          =   -1  'True
                  MaxLength       =   30
                  TabIndex        =   3
                  Tag             =   "null"
                  Text            =   "TxtNombreCorto"
                  Top             =   1575
                  Width           =   3585
               End
               Begin VB.TextBox TxtVariable 
                  BackColor       =   &H8000000F&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1425
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   2
                  Text            =   "TxtVariable"
                  Top             =   1170
                  Width           =   4245
               End
               Begin VB.TextBox TxtDescripcion 
                  Height          =   315
                  Left            =   1425
                  Locked          =   -1  'True
                  MaxLength       =   200
                  TabIndex        =   1
                  Text            =   "TxtDescripcion"
                  Top             =   765
                  Width           =   9510
               End
               Begin VB.Frame FraTipo 
                  Caption         =   "[ Seleccionar Tipo ]"
                  Enabled         =   0   'False
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
                  Height          =   645
                  Left            =   150
                  TabIndex        =   24
                  Top             =   2010
                  Width           =   5760
                  Begin VB.OptionButton OptOrigen 
                     Caption         =   "Formula"
                     Height          =   225
                     Index           =   1
                     Left            =   3180
                     TabIndex        =   25
                     Top             =   300
                     Width           =   2100
                  End
                  Begin VB.OptionButton OptOrigen 
                     Caption         =   "Cuenta"
                     Height          =   225
                     Index           =   0
                     Left            =   420
                     TabIndex        =   4
                     Top             =   300
                     Value           =   -1  'True
                     Width           =   1905
                  End
               End
               Begin VB.CommandButton CmdBusCat 
                  Height          =   225
                  Left            =   1950
                  Picture         =   "FrmConcepto.frx":36D2
                  Style           =   1  'Graphical
                  TabIndex        =   21
                  ToolTipText     =   "Seleccione la Categoría de Concepto"
                  Top             =   420
                  Width           =   210
               End
               Begin VB.Frame Frame6 
                  BeginProperty Font 
                     Name            =   "Trebuchet MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1785
                  Left            =   180
                  TabIndex        =   16
                  Top             =   2760
                  Width           =   10815
                  Begin VB.CommandButton cmd_formula 
                     Caption         =   "Editar Formula"
                     Enabled         =   0   'False
                     Height          =   645
                     Left            =   120
                     Picture         =   "FrmConcepto.frx":3804
                     Style           =   1  'Graphical
                     TabIndex        =   5
                     Top             =   540
                     Width           =   1275
                  End
                  Begin VB.TextBox TxtFormula 
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1095
                     Left            =   1455
                     Locked          =   -1  'True
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   17
                     Text            =   "FrmConcepto.frx":3906
                     Top             =   270
                     Width           =   9045
                  End
               End
               Begin VB.TextBox TxtIdCat 
                  Height          =   300
                  Left            =   1425
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   0
                  Text            =   "TxtIdCat"
                  Top             =   390
                  Width           =   765
               End
               Begin VB.TextBox TxtComentario 
                  Height          =   870
                  Left            =   240
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   6
                  Tag             =   "null"
                  Text            =   "FrmConcepto.frx":3916
                  Top             =   4920
                  Width           =   10800
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nombre Corto"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   3
                  Left            =   150
                  TabIndex        =   26
                  ToolTipText     =   "Es nombre abreviado del concepto, con posibilidad de que aparesca en el reporte"
                  Top             =   1695
                  Width           =   975
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Categoría"
                  Height          =   195
                  Index           =   0
                  Left            =   240
                  TabIndex        =   22
                  Top             =   495
                  Width           =   705
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Comentario :"
                  Height          =   210
                  Index           =   5
                  Left            =   240
                  TabIndex        =   20
                  Top             =   4665
                  Width           =   900
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Variable"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   2
                  Left            =   150
                  TabIndex        =   19
                  Top             =   1305
                  Width           =   570
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Descripción"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   1
                  Left            =   180
                  TabIndex        =   18
                  Top             =   900
                  Width           =   840
               End
               Begin VB.Label LblDescCat 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblDescCategoria"
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
                  Height          =   285
                  Left            =   2190
                  TabIndex        =   23
                  Top             =   390
                  Width           =   3495
               End
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Conceptos"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   90
            TabIndex        =   10
            Top             =   60
            Width           =   11400
         End
      End
   End
End
Attribute VB_Name = "FrmConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--modificado
'Fecha      |Por            |Procedimiento      |Descripcion
'=========  ==============  =================== ===============
'11/11/09   |JCastro        |                   |Agregar campo desctabal; Identificar si es activo o pasivo los conceptos
'                                               cuando categoria sea balance
'
'
 
 
Option Explicit
Dim RstFrm As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim Agregando As Boolean

Dim fOrdenLista As Boolean ''--especifica el orden de la lista de la consulta
Dim mIdRegistro& '--identificador del registro

Dim xHorIni As Date
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO



Sub Cancelar()
    QueHace = 3
    Bloquea False
    ActivaTool
    Label5.Caption = "Detalle de Conceptos"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
End Sub




Private Sub cmd_formula_Click()
    FrmConceptoFormula.Show 1
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstFrm
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If ColIndex = 5 Then Exit Sub
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    'Modificado: 11/01/11 Johan Castro
    '            Agregar linea de codigo para bloquear accesos de usuarios
    
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        '----------------------------------------------
        
        pCargarGrid
        
    End If
End Sub


Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
End Sub


Sub Blanquea()
    TxtFormula.Text = ""
    TxtFormula.Tag = ""
    TxtNombreCorto.Text = ""
    TxtVariable.Text = ""
    TxtDescripcion.Text = ""
    TxtComentario.Text = ""
    
    TxtIdCat.Text = ""
    LblDescCat.Caption = ""
    
    TxtId.Text = ""
    
    Fg1.Rows = 1
End Sub

Sub Bloquea(band As Boolean)

    TxtIdCat.Locked = Not band
    TxtDescripcion.Locked = Not band
    TxtNombreCorto.Locked = Not band
    FraTipo.Enabled = band
    TxtComentario.Locked = Not band

    FraDestinoCta.Enabled = band

    If (QueHace = 1) Or (QueHace = 2 And NulosC(TxtFormula.Text) = "") Then
        TxtVariable.Enabled = True
        TxtVariable.BackColor = vbWhite
    Else
    
        Dim RstTmp As New ADODB.Recordset
        Dim nSQl As String
            
        If RstFrm.RecordCount <> 0 Then
            nSQl = "SELECT con_concepto.id FROM con_concepto WHERE ucase(con_concepto.formula) Like '%" & UCase(NulosC(RstFrm.Fields("variable"))) & "%' ;"
            RST_Busq RstTmp, nSQl, xCon
            If RstTmp.RecordCount <> 0 Then
                TxtVariable.Enabled = False
                TxtVariable.BackColor = &H8000000F
            Else
                TxtVariable.Enabled = True
                TxtVariable.BackColor = vbWhite
            End If
        Else
            TxtVariable.Enabled = True
            TxtVariable.BackColor = vbWhite
        End If
        Set RstTmp = Nothing
        
        
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set RstFrm = Nothing
End Sub



Private Sub OptOrigen_Click(Index As Integer)
    If OptOrigen(0).Value = True Then
        cmd_formula.Enabled = False
        TxtFormula.Text = ""
    Else
        cmd_formula.Enabled = True
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Function Grabar() As Boolean

    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " al Concepto ", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset '--relacionado a los aportes que se consideraran
    Dim nSQl As String
    Dim xId As Double
    Dim A&
    On Error GoTo LaCague

    xCon.BeginTrans

    If QueHace = 1 Then

        xId = HallaCodigoTabla("con_concepto", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_concepto", xCon

        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM con_concepto WHERE id = " & xId & ";", xCon
        
        '--eliminar los conceptos de aportes relacionado con el concepto
        xCon.Execute "DELETE FROM con_conceptodet WHERE idcpto = " & xId & ";"
        
    End If
    
    '******************
    mIdRegistro = xId
    '******************
    RST_Busq RstDet, "SELECT TOP 1 * FROM con_conceptodet", xCon
    
    
    RstCab("idcat") = NulosN(TxtIdCat.Text)
    RstCab("descripcion") = NulosC(TxtDescripcion.Text)
    RstCab("variable") = NulosC(TxtVariable.Text)
    RstCab("formula") = NulosC(TxtFormula.Text)
    RstCab("nomcorto") = NulosC(TxtNombreCorto.Text)
    RstCab("observacion") = TxtComentario.Text
    
    '--0 Cuenta; --1 Formula
    RstCab("origen") = IIf(OptOrigen(0).Value = True, 0, -1)
       
    '--cuando se crea un concepto; el sistema activara por defecto
    If QueHace = 1 Then RstCab("activo") = -1
    
    '--solo para balance, grabar destino de la cuenta
    RstCab("desctabal") = 0
    If NulosN(TxtIdCat.Text) = 1 Then
        If OptActivo.Value = True Then RstCab("desctabal") = 1
        If OptPasivo.Value = True Then RstCab("desctabal") = 2
    End If
    
    
    
    
    RstCab.Update
    
    '--
    '--grabar las cuentas relacionadas al concepto
    If OptOrigen(0).Value = True Then
        For A = Fg1.FixedRows To Fg1.Rows - 1
            DoEvents
            RstDet.AddNew
            RstDet("idcpto") = xId
            RstDet("idref") = NulosN(Fg1.TextMatrix(A, 1))
            RstDet.Update
        Next A
    Else
        '--insertar los conceptos relacionado con la formula
        xCon.Execute "INSERT INTO con_conceptodet (idcpto,idref) " _
                & " SELECT " & xId & " AS idcpto, con_concepto.id AS idref " _
                & " From con_concepto " _
                & " WHERE (((InStr('" & NulosC(TxtFormula.Text) & "',[con_concepto].[variable]))<>0));"

    End If
    xCon.CommitTrans
    Set RstCab = Nothing
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    
    MsgBox "Los datos del Concepto " + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo


    Grabar = True
    Exit Function
LaCague:
    Set RstCab = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar al Personal por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Sub Nuevo()
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea True
    Blanquea
    
    FraDestinoCta.Visible = False
    OptActivo.Value = False
    OptPasivo.Value = False
    
    Label5.Caption = "Agregando Conceptos"
    TabOne2.CurrTab = 0
    
    xHorIni = Time

    '-------------------------------------------
    TxtIdCat.SetFocus
    
End Sub

Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Label5.Caption = "Modificando Concepto"

    ActivaTool
    
    QueHace = 2
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
    End If
    
    Bloquea True
    
    TabOne1.TabEnabled(0) = False
    
    Agregando = False
    If TabOne2.CurrTab <> 0 Then TabOne2.CurrTab = 0
    
    xHorIni = Time
    
    TxtIdCat.SetFocus

End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Eliminar()
    TabOne1.CurrTab = 0
    Dim Rpta As Integer
    Dim xId&
    Dim nSQl As String
    
    Dim RstBus As New ADODB.Recordset
    xId = RstFrm.Fields("id")
    '---------------------------------------------------------------------------------
    '--validar que no este en formulas de otros conceptos
    nSQl = "SELECT con_concepto.id, con_concepto.descripcion, con_concepto.variable, con_concepto.formula " _
        + vbCr + " FROM con_concepto " _
        + vbCr + " WHERE (((con_concepto.id)<>" & xId & ") AND ((con_concepto.formula) Like '%" & RstFrm.Fields("variable") & "%'));"
        
    RST_Busq RstBus, nSQl, xCon
    If RstBus.RecordCount <> 0 Then
        MsgBox "No se puede eliminar el Concepto, figura en fórmulas de otros conceptos" + vbCr + "Ej. " & RstBus.Fields("descripcion") & vbCr & "Eliminar primero el Concepto Mensionado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstBus = Nothing
        Exit Sub
    End If
    
    Set RstBus = Nothing
    '---------------------------------------------------------------------------------
        
    nSQl = "SELECT TOP 1 con_informe.descripcion, con_informedet.idcpto " _
        + vbCr + " FROM con_informedet INNER JOIN con_informe ON con_informedet.idinf = con_informe.id " _
        + vbCr + " WHERE (((con_informedet.idcpto)=" & xId & ")); "
    
    RST_Busq RstBus, nSQl, xCon
    If RstBus.RecordCount <> 0 Then
        MsgBox "No se puede eliminar el Concepto, esta relacionado al Informe " + vbCr + "Ej. " & RstBus.Fields("descripcion") & vbCr & "Eliminar primero el Concepto del Informe", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstBus = Nothing
        Exit Sub
    End If
    
    Set RstBus = Nothing
    
    '---------------------------------------------------------------------------------
    
    Rpta = MsgBox("Esta seguro de eliminar al Concepto seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
    
        xCon.Execute "DELETE FROM con_conceptodet WHERE idcpto = " & xId & ";" '--replaciona al cuentas contables
        xCon.Execute "DELETE FROM con_concepto WHERE id = " & xId & ";"
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo

        
        MsgBox "El Concepto se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstFrm.Requery
        Dg1.Refresh
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    If Button.Index = 5 Then Cancelar
    If Button.Index = 6 Then
        If Grabar = True Then
            RstFrm.Requery
            Dg1.Refresh
            If RstFrm.State = 1 Then
                If RstFrm.RecordCount <> 0 Then
                    RstFrm.MoveFirst
                    RstFrm.Find "id =" & mIdRegistro
                    If RstFrm.EOF = True Then RstFrm.MoveFirst
                End If
            End If
            Cancelar
        End If
    End If

    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstFrm.Filter = adFilterNone
    End If
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 12 Then pExportar
    

    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Buscar()

End Sub

Sub Filtrar()

End Sub


'****************************************************************************************
'****************************************************************************************
'****************************************************************************************

Sub MuestraSegundoTab()
    Dim QueHaceTmp As Integer
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros para Mostrar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    Blanquea
    TabOne2.CurrTab = 0
    '--datos de concepto
    TxtId.Text = RstFrm("id")
    TxtIdCat.Text = RstFrm("idcat")
    LblDescCat.Caption = NulosC(RstFrm("catnombre"))
    TxtDescripcion.Text = NulosC(RstFrm("descripcion"))
    
    TxtVariable.Text = NulosC(RstFrm("variable"))
    TxtNombreCorto.Text = NulosC(RstFrm("nomcorto"))
    TxtFormula.Tag = ""
    TxtFormula.Text = NulosC(RstFrm("formula"))
    
    If NulosC(RstFrm("formula")) <> "" Then
        
        TxtFormula.Text = MostrarFormulaEquivalente(RstFrm("id")) 'RstFrm("formula")
        TxtFormula.Tag = RstFrm("formula")
        
    End If
    TxtComentario.Text = NulosC(RstFrm("observacion"))
    
    cmd_formula.Tag = TxtFormula.Text
    
    If NulosN(TxtIdCat.Text) = 1 Then
        
        FraDestinoCta.Visible = True
        
        If NulosN(RstFrm("desctabal")) = 1 Then
            OptActivo.Value = True
        ElseIf NulosN(RstFrm("desctabal")) = 2 Then
            OptPasivo.Value = True
        Else
            OptActivo.Value = False
            OptPasivo.Value = False
        End If
    Else
        FraDestinoCta.Visible = False
        
    End If
    
    '--0 Cuenta; --1 Formula
    If NulosN(RstFrm("origen")) = 0 Then '--x cuenta
        Fg1.TextMatrix(0, 2) = "N°.Cta"
        OptOrigen(0).Value = True
    Else '--x formula
        Fg1.TextMatrix(0, 2) = "Variable"
        OptOrigen(1).Value = True
    End If
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQl As String
    

    
    nSQl = "SELECT con_conceptodet.idref  AS refid, IIf(con_concepto_1.origen=-1,con_concepto.variable,con_planctas.cuenta) AS RefNombre1, IIf(con_concepto_1.origen=-1,con_concepto.descripcion,con_planctas.descripcion) AS RefNombre2 " _
        + vbCr + " FROM ((con_conceptodet LEFT JOIN con_planctas ON con_conceptodet.idref = con_planctas.id) LEFT JOIN con_concepto ON con_conceptodet.idref = con_concepto.id) INNER JOIN con_concepto AS con_concepto_1 ON con_conceptodet.idcpto = con_concepto_1.id " _
        + vbCr + " WHERE (((con_conceptodet.idcpto) = " & NulosN(RstFrm("id")) & ")) " _
        + vbCr + " ORDER BY IIf(con_concepto_1.origen=-1,con_concepto.variable,con_planctas.cuenta);"
    
    
    RST_Busq RstTmp, nSQl, xCon
    
    Fg1.Rows = 1
    DoEvents
    Do While Not RstTmp.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosN(RstTmp("refid"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstTmp("refnombre1"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstTmp("refnombre2"))
        RstTmp.MoveNext
    Loop

End Sub


Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQl As String
    
    Set RstFrm = Nothing
    
    TDB_FiltroLimpiar Dg1
    
    RstFrm.Filter = adFilterNone
    DoEvents

    nSQl = "SELECT con_conceptocat.descripcion AS catnombre, con_concepto.* " _
        + vbCr + " FROM con_concepto LEFT JOIN con_conceptocat ON con_concepto.idcat = con_conceptocat.id WHERE con_concepto.activo =-1"

    
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQl, xCon
    
    Set Dg1.DataSource = RstFrm
    TabOne1.CurrTab = 0
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Function fValidarDatos() As Boolean

    TabOne2.CurrTab = 0
    

    If NulosN(TxtIdCat.Text) = 0 Then
        MsgBox "Falta especificar la Categoria", vbExclamation, xTitulo
        TxtIdCat.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtDescripcion.Text) = "" Then
        MsgBox "Falta especificar la Descripción del Concepto", vbExclamation, xTitulo
        TxtDescripcion.SetFocus
        Exit Function
    End If

    If NulosC(TxtVariable.Text) = "" Then
        MsgBox "Falta especificar la Variable del Concepto", vbExclamation, xTitulo
        TxtVariable.SetFocus
        Exit Function
    End If
    
    '--verificar que la variable no tenga cietos caracteres
    
    Dim mCantCarateres&
    For mCantCarateres = 1 To Len(TxtVariable.Text)
        If InStr("()=+*-/[],: .'?¿!¡%&$#@<>áéíóúñ|°", Mid(TxtVariable.Text, mCantCarateres, 1)) <> 0 Then
            MsgBox "Caracter no Permitido: [ " & Mid(TxtVariable.Text, mCantCarateres, 1) & " ]" + vbCr + "Modifique la variable", vbInformation, xTitulo
            Exit Function
        End If
    Next
    
    
    If OptOrigen(0).Value = True And Fg1.Rows = Fg1.FixedRows Then
''        MsgBox "Falta seleccionar las Cuentas Contables para este concepto", vbExclamation, xTitulo
''        TabOne2.CurrTab = 1
''        Exit Function
    End If
    
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQl As String
    If QueHace = 1 Then
        nSQl = "SELECT con_concepto.descripcion, UCase([con_concepto].[variable])  FROM con_concepto WHERE (((UCase([con_concepto].[variable]))='" & UCase(TxtFormula.Text) & "'));"
    Else
        nSQl = "SELECT con_concepto.descripcion, UCase([con_concepto].[variable]) AS Expr1  FROM con_concepto WHERE (((UCase([con_concepto].[variable]))='" & UCase(TxtFormula.Text) & "') AND ((con_concepto.id)<>" & NulosN(RstFrm.Fields("id")) & "));"
    End If
    RST_Busq RstTmp, nSQl, xCon
    If RstTmp.RecordCount <> 0 Then
        MsgBox "Existe un Concepto que tiene asignado la misma variable" + vbCr + "Concepto: " + NulosC(RstFrm("descripcion")) + vbCr + "Cambie el nombre de la Variable", vbExclamation, xTitulo
        Exit Function
    End If
    '--
    fValidarDatos = True
    
End Function


Private Sub pExportar()
    TabOne1.CurrTab = 0
    
    Dim rst As New ADODB.Recordset
    Dim nSQl As String
    Dim oExport As New SGI2_funciones.formularios
    
    nSQl = "SELECT con_conceptodet.idcpto, con_conceptocat.descripcion AS categoria, con_concepto_1.variable, con_concepto_1.descripcion, con_concepto_1.formula, con_conceptodet.idref AS refid, IIf(con_concepto_1.origen=-1,con_concepto.variable,con_planctas.cuenta) AS RefNombre1, IIf(con_concepto_1.origen=-1,con_concepto.descripcion,con_planctas.descripcion) AS RefNombre2 " _
        + vbCr + " FROM con_conceptocat RIGHT JOIN (((con_conceptodet LEFT JOIN con_planctas ON con_conceptodet.idref = con_planctas.id) LEFT JOIN con_concepto AS con_concepto_1 ON con_conceptodet.idcpto = con_concepto_1.id) LEFT JOIN con_concepto ON con_conceptodet.idref = con_concepto.id) ON con_conceptocat.id = con_concepto_1.idcat " _
        + vbCr + " ORDER BY con_concepto_1.idcat, con_concepto_1.orden; "
    
    RST_Busq rst, nSQl, xCon
        
    Dim xCampos(6, 3) As String
    
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Id":           xCampos(0, 1) = "idcpto":       xCampos(0, 2) = 0:  xCampos(0, 3) = "500"
    xCampos(1, 0) = "Categoria":    xCampos(1, 1) = "categoria":    xCampos(1, 2) = 0:  xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Variable":     xCampos(2, 1) = "variable":     xCampos(2, 2) = 0:  xCampos(2, 3) = "750"
    xCampos(3, 0) = "Descripción":  xCampos(3, 1) = "descripcion":  xCampos(3, 2) = 0:  xCampos(3, 3) = "3300"
    xCampos(4, 0) = "Fórmula":      xCampos(4, 1) = "formula":      xCampos(4, 2) = 0:  xCampos(4, 3) = "900"
    xCampos(5, 0) = "Ref. Código":  xCampos(5, 1) = "refnombre1":   xCampos(5, 2) = 0:  xCampos(5, 3) = "1100"
    xCampos(6, 0) = "Ref. Nombre":  xCampos(6, 1) = "refnombre2":   xCampos(6, 2) = 0:  xCampos(6, 3) = "4500"
    
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Lista de Conceptos de EE.FF.", "", "", "Conceptos de EE.FF.", rst, xCampos
    Set oExport = Nothing
    Set rst = Nothing
    
End Sub



'************************************************************
'************************************************************

Private Sub cmd_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Select Case Index
        Case 0 '--ADD REG
            pRegistroAdd 0
        Case 1 '--DEL REG
            pRegistroDel 0
        '--DE LAS CUENTAS
        Case 2 '--ADD
            pRegistroAdd True
        Case 3 '--SEL
            pRegistroAdd 1, True
        Case 4 '--DEL
            pRegistroDel 1
            
    End Select
End Sub


Private Sub pRegistroAdd(Index As Integer, Optional fSelVarios As Boolean = False)
    '--salir si estas en consulta
    If QueHace = 3 Then Exit Sub
    '--solo agregara un registro si origen es por cuenta
    If OptOrigen(1).Value = True Then Exit Sub
    
    If NulosN(TxtIdCat.Text) = 0 Then
        MsgBox "Falta especificar la Categoría", vbInformation, xTitulo
        TxtIdCat.SetFocus
        Exit Sub
    End If
    
    Agregando = True
    
    '-------------------------------
    '--GENERAR EL WHERE DE LOS ID'S DE CUENTA PARA QUE NO SE REPITAN
    Dim nSQLId As String

    '--obtener la lista de cuentas para no considerar en al sgte conculta
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 1, " AND con_planctas.id", " not in ", True)
        
    '-------------------------------
    On Error GoTo error
    Dim xRs  As New ADODB.Recordset
    Dim nSQl As String
    Dim nSQLDistribucion As String '--Sentencia SQL que filtrara la distribucion de la cuenta
    
        
    If NulosN(TxtIdCat.Text) = 1 Then
        nSQLDistribucion = " AND IIf(con_planctas.iddes=1 Or con_planctas.iddes2=1,-1,0)=-1 "
        
    ElseIf NulosN(TxtIdCat.Text) = 2 Then
        nSQLDistribucion = " AND IIf(con_planctas.iddes=3 Or con_planctas.iddes2=3,-1,0)=-1 "
        
    ElseIf NulosN(TxtIdCat.Text) = 3 Then
        nSQLDistribucion = " AND IIf(con_planctas.iddes=2 Or con_planctas.iddes2=2,-1,0)=-1 "
        
    Else
        
    End If
       
    
    If fSelVarios = True Then
    
        ReDim xCampos(3, 5) As String
        xCampos(0, 0) = "N° Cta.":      xCampos(0, 1) = "cuenta":           xCampos(0, 2) = "1500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
        xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "6500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
        xCampos(2, 0) = "Pendiente":    xCampos(2, 1) = "pendiente":      xCampos(2, 2) = "850":     xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
        
        nSQl = "SELECT 0 as xsel, con_planctas.id as idcuenta, con_planctas.cuenta & '' AS cuenta, con_planctas.descripcion,IIf(ctacpto.idref Is Null,'Si','') AS pendiente " _
                + vbCr + " FROM (con_diario INNER JOIN con_planctas ON con_diario.idcue = con_planctas.id ) " _
                + vbCr + " LEFT JOIN ( SELECT con_concepto.idcat, con_concepto.origen, con_conceptodet.idref FROM con_concepto INNER JOIN con_conceptodet ON con_concepto.id = con_conceptodet.idcpto WHERE (((con_concepto.origen)=0)) and con_concepto.idcat=" & NulosN(TxtIdCat.Text) & " ) AS ctacpto ON con_planctas.id = ctacpto.idref " _
                + vbCr + " WHERE con_planctas.tipo=0 " & nSQLId & nSQLDistribucion _
                + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta & ' ', con_planctas.descripcion, con_planctas.cuenta ,ctacpto.idref, ctacpto.idcat " _
                + vbCr + " HAVING (((ctacpto.idcat)=" & NulosN(TxtIdCat.Text) & " Or (ctacpto.idcat) Is Null)) " _
                + vbCr + " ORDER BY con_planctas.cuenta ASC "
        
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQl, xCampos(), "Buscando Cuentas con Movimientos"
    Else
        ReDim xCampos(2, 5) As String
        xCampos(0, 0) = "N° Cta.":      xCampos(0, 1) = "cuenta":           xCampos(0, 2) = "1500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
        xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "6500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    
    
        nSQl = "SELECT con_planctas.id as idcuenta, con_planctas.cuenta & '' AS cuenta, con_planctas.descripcion " _
                + vbCr + " FROM con_planctas " _
                + vbCr + " WHERE con_planctas.tipo=0 " & nSQLId & nSQLDistribucion _
                + vbCr + " ORDER BY con_planctas.cuenta ASC "
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQl, xCampos(), "Agregando Cuentas", "cuenta", "cuenta", Principio
        
    End If
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR

    If fSelVarios = True Then xRs.MoveFirst
    Do While Not xRs.EOF
        With Fg1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = xRs.Fields("idcuenta") & ""
            .TextMatrix(.Rows - 1, 2) = xRs.Fields("cuenta") & ""
            .TextMatrix(.Rows - 1, 3) = xRs.Fields("descripcion") & ""
            '---
        End With
        If fSelVarios = False Then Exit Do
        xRs.MoveNext
    Loop
SALIR:
    Agregando = False
    Set xRs = Nothing
    '----
     Fg1.SetFocus
    Exit Sub
error:
    Agregando = False
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "pRegistroAdd"
End Sub


Private Sub pRegistroDel(Index As Integer)
    If QueHace = 3 Then Exit Sub
    '--solo eliminara un registro si origen es por cuenta
    If OptOrigen(1).Value = True Then Exit Sub
    If Fg1.Row <= 0 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una fila correcta", vbExclamation, xTitulo
        Exit Sub
    End If
    
    Fg1.RemoveItem (Fg1.Row)
    
End Sub


'************************************************************
'*************************************************************
Private Sub CmdBusCat_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripción":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Código":         xCampos(1, 1) = "id":               xCampos(1, 2) = "500":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM con_conceptocat ORDER BY descripcion"
    
    xform.Titulo = "Buscando Modalidad"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdCat.Text = xRs("id")
            LblDescCat.Caption = xRs("descripcion")
            
            '--mostrar cuando categoria sea balance
            If xRs("id") = 1 Then
                FraDestinoCta.Visible = True
            Else
                FraDestinoCta.Visible = False
            End If
            '----------------------------------
            TxtDescripcion.SetFocus
        End If
    End If
    '--estableciendo la nueva variable
    HallaVariable
    '-----------------
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub TxtIdCat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdCat_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCat_Click
    End If
End Sub

Private Sub TxtIdCat_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    'buscamos el codigo de la moneda         digitada
    RST_Busq xRs, "SELECT * FROM con_conceptocat WHERE id = " & NulosN(TxtIdCat.Text) & "", xCon
    
    If xRs.RecordCount = 0 Then
        TxtIdCat.Text = ""
        LblDescCat.Caption = ""
    Else
        LblDescCat.Caption = Trim(xRs("descripcion"))
        
        '--mostrar cuando categoria sea balance
        If xRs("id") = 1 Then
            FraDestinoCta.Visible = True
        Else
            FraDestinoCta.Visible = False
        End If
        '----------------------------------
        
        '--estableciendo la nueva variable
        HallaVariable
        '-----------------
    End If
    Set xRs = Nothing

End Sub

'*************************************************************




Private Sub HallaVariable()
    Dim rst As New ADODB.Recordset
    Dim nSQl As String
    Dim mCorr As Long
    
    '--Solo generara la variable cuando se agrege un registro, si se modifica no cambia la variable
    
    If QueHace <> 1 Then Exit Sub
    
    If NulosN(TxtIdCat.Text) = 0 Then
        MsgBox "Falta especificar la Categoria para obtener la variable", vbInformation, xTitulo
        TxtIdCat.SetFocus
        Exit Sub
    End If
    
    nSQl = "SELECT Count(con_concepto.idcat) AS totreg, con_conceptocat.prefijo,Last(con_concepto.variable) AS varultimo " _
        & " FROM con_concepto RIGHT JOIN con_conceptocat ON con_concepto.idcat = con_conceptocat.id " _
        & " GROUP BY con_conceptocat.id, con_conceptocat.prefijo HAVING (((con_conceptocat.id)=" & NulosN(TxtIdCat.Text) & ")); "
        
    RST_Busq rst, nSQl, xCon
    
    If rst.RecordCount <> 0 Then
        mCorr = NulosN(Mid(rst("varultimo"), Len(rst("prefijo")) + 1)) + 1
    Else
        mCorr = 1
    End If
    
    
    
    TxtVariable.Text = NulosC(rst("prefijo")) & Format(mCorr, "000")
    Set rst = Nothing

End Sub
