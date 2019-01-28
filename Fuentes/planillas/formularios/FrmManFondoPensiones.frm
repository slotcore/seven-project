VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmManFondoPensiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Aportes, Comisiones y Prima de Seguros "
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   9540
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":1EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":23EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManFondoPensiones.frx":277E
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
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   9585
      _cx             =   16907
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
      Caption         =   "   Consulta   |   Detalle   "
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
         Height          =   6795
         Left            =   -10140
         TabIndex        =   4
         Top             =   375
         Width           =   9495
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6405
            Left            =   45
            TabIndex        =   5
            Top             =   390
            Width           =   9390
            _ExtentX        =   16563
            _ExtentY        =   11298
            _LayoutType     =   4
            _RowHeight      =   14
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Mes"
            Columns(0).DataField=   "mes"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "AFP"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Aporte Obligatorio"
            Columns(2).DataField=   "aporte"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Comisión Variable"
            Columns(3).DataField=   "comision"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Prima de Seguros"
            Columns(4).DataField=   "prima"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1720"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1640"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=256"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=5398"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5318"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=256"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2858"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2778"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=770"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2805"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2725"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=770"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2778"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2699"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=770"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
            _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0,.bold=0,.fontsize=825"
            _StyleDefs(45)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(46)  =   ":id=28,.fontname=MS Sans Serif"
            _StyleDefs(47)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=0"
            _StyleDefs(48)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(51)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=0"
            _StyleDefs(52)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(55)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14,.alignment=1"
            _StyleDefs(56)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(59)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=1"
            _StyleDefs(60)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(63)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14,.alignment=1"
            _StyleDefs(64)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(66)  =   "Named:id=33:Normal"
            _StyleDefs(67)  =   ":id=33,.parent=0"
            _StyleDefs(68)  =   "Named:id=34:Heading"
            _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(70)  =   ":id=34,.wraptext=-1"
            _StyleDefs(71)  =   "Named:id=35:Footing"
            _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(73)  =   "Named:id=36:Selected"
            _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(75)  =   "Named:id=37:Caption"
            _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(77)  =   "Named:id=38:HighlightRow"
            _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=39:EvenRow"
            _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(81)  =   "Named:id=40:OddRow"
            _StyleDefs(82)  =   ":id=40,.parent=33"
            _StyleDefs(83)  =   "Named:id=41:RecordSelector"
            _StyleDefs(84)  =   ":id=41,.parent=34"
            _StyleDefs(85)  =   "Named:id=42:FilterBar"
            _StyleDefs(86)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Aportes, Comisiones y Prima de Seguros "
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   105
            TabIndex        =   6
            Top             =   60
            Width           =   9435
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6795
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   9495
         Begin VB.Frame Frame3 
            Height          =   6210
            Left            =   330
            TabIndex        =   9
            Top             =   450
            Width           =   8460
            Begin VB.Frame fra 
               BorderStyle     =   0  'None
               Caption         =   "Frame12"
               Height          =   585
               Index           =   0
               Left            =   525
               TabIndex        =   16
               Top             =   5460
               Width           =   7230
               Begin VB.CommandButton CmdDet 
                  Caption         =   "&Eliminar"
                  Enabled         =   0   'False
                  Height          =   435
                  Index           =   1
                  Left            =   1410
                  TabIndex        =   19
                  Top             =   75
                  Width           =   1260
               End
               Begin VB.CommandButton CmdDet 
                  Caption         =   "&Agregar"
                  Enabled         =   0   'False
                  Height          =   435
                  Index           =   0
                  Left            =   120
                  TabIndex        =   18
                  Top             =   75
                  Width           =   1260
               End
               Begin VB.CommandButton CmdDetRestablece 
                  Caption         =   "Duplicar de Último Periodo"
                  Enabled         =   0   'False
                  Height          =   435
                  Left            =   5085
                  TabIndex        =   17
                  Top             =   75
                  Width           =   2040
               End
               Begin VB.Line Line5 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   2
                  X1              =   15
                  X2              =   15
                  Y1              =   0
                  Y2              =   1000
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00808080&
                  BorderWidth     =   2
                  Index           =   3
                  X1              =   7215
                  X2              =   7215
                  Y1              =   15
                  Y2              =   1000
               End
               Begin VB.Line lin 
                  BorderColor     =   &H00808080&
                  BorderWidth     =   2
                  Index           =   3
                  X1              =   -30
                  X2              =   13000
                  Y1              =   570
                  Y2              =   570
               End
               Begin VB.Line lin 
                  BorderColor     =   &H80000009&
                  BorderWidth     =   2
                  Index           =   2
                  X1              =   -15
                  X2              =   13000
                  Y1              =   15
                  Y2              =   15
               End
            End
            Begin VB.Frame fra 
               BorderStyle     =   0  'None
               Caption         =   "Frame12"
               Height          =   480
               Index           =   1
               Left            =   525
               TabIndex        =   10
               Top             =   240
               Width           =   7230
               Begin MSComCtl2.UpDown UpDown1 
                  Height          =   300
                  Left            =   1290
                  TabIndex        =   11
                  Top             =   90
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   2008
                  OrigLeft        =   1155
                  OrigTop         =   195
                  OrigRight       =   1410
                  OrigBottom      =   495
                  Max             =   2999
                  Min             =   2007
                  Wrap            =   -1  'True
                  Enabled         =   0   'False
               End
               Begin MSDataListLib.DataCombo dtcb_periodo 
                  Height          =   315
                  Left            =   2595
                  TabIndex        =   12
                  Top             =   75
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Style           =   2
                  Text            =   "dtcb_periodo"
               End
               Begin VB.TextBox txt_ano 
                  Height          =   300
                  Left            =   540
                  Locked          =   -1  'True
                  TabIndex        =   13
                  Text            =   "txt_ano"
                  Top             =   90
                  Width           =   510
               End
               Begin VB.Line lin 
                  BorderColor     =   &H80000009&
                  BorderWidth     =   2
                  Index           =   0
                  X1              =   -15
                  X2              =   13000
                  Y1              =   15
                  Y2              =   15
               End
               Begin VB.Line lin 
                  BorderColor     =   &H00808080&
                  BorderWidth     =   2
                  Index           =   1
                  X1              =   -30
                  X2              =   13000
                  Y1              =   465
                  Y2              =   465
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H00808080&
                  BorderWidth     =   2
                  Index           =   0
                  X1              =   7215
                  X2              =   7215
                  Y1              =   15
                  Y2              =   1000
               End
               Begin VB.Line Line2 
                  BorderColor     =   &H00FFFFFF&
                  BorderWidth     =   2
                  X1              =   15
                  X2              =   15
                  Y1              =   0
                  Y2              =   1000
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Periodo"
                  Height          =   195
                  Index           =   0
                  Left            =   1950
                  TabIndex        =   15
                  Top             =   195
                  Width           =   540
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Año"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   14
                  Top             =   195
                  Width           =   285
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   4455
               Left            =   495
               TabIndex        =   20
               Top             =   780
               Width           =   7245
               _cx             =   12779
               _cy             =   7858
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
               Rows            =   1
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManFondoPensiones.frx":2B10
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
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   285
            Index           =   0
            Left            =   8580
            TabIndex        =   7
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   -15
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   8040
            TabIndex        =   8
            Top             =   75
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Aportes, Comisiones y Prima de Seguros "
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   60
            TabIndex        =   3
            Top             =   120
            Width           =   9375
         End
      End
   End
End
Attribute VB_Name = "FrmManFondoPensiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RstFrm As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim Agregando As Boolean

Dim fOrdenLista As Boolean    '--especfica el orden de la lista de la consulta
Dim CargarDefault As Boolean  '--true cuando se requiere que cargue los datos de manera predeterminada
                              '--false cuando se consulta desde la bd


Sub Cancelar()
    ActivaTool
    Label5.Caption = "Detalle del Asignación de Sueldo"
    QueHace = 3
    Bloquea False
    TabOne1.TabEnabled(0) = True
    
    Fg1.SelectionMode = flexSelectionByRow
    Me.MousePointer = vbDefault
    TabOne1.CurrTab = 0
End Sub




Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub


Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    Err.Clear
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Row < 1 Then Exit Sub
    If Col <> 2 Then Exit Sub
    On Error GoTo error
    Dim xCampos(3, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String '--almacenara los codigos de documentos ya seleccionados
    Dim nSQL As String

    xCampos(0, 0) = "Cod.Sunat":    xCampos(0, 1) = "codsun":       xCampos(0, 2) = "1000":   xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
    xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "3500":  xCampos(1, 3) = "C":     xCampos(1, 4) = "S"
    xCampos(2, 0) = "Id":           xCampos(2, 1) = "id":           xCampos(2, 2) = "600":   xCampos(2, 3) = "N":     xCampos(2, 4) = "N"
        
    '*************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 1, "mae_regimenpen.id", " NOT IN ")
    '*************************************************************
    nSQL = "SELECT mae_regimenpen.* FROM mae_regimenpen " _
        + vbCr + " WHERE mae_regimenpen.cuspp=-1 " _
        + vbCr + IIf(nSQLId = "", "", " AND " + nSQLId)
    '*************************************************************
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando SPP", "descripcion", "descripcion", Principio
    If xRs.State = 0 Then GoTo SALIR
    If xRs.RecordCount = 0 Then GoTo SALIR
    Agregando = True
    Do While Not xRs.EOF
        Fg1.TextMatrix(Row, 1) = NulosN(xRs("id"))
        Fg1.TextMatrix(Row, 2) = NulosC(xRs("descripcion"))
        xRs.MoveNext
    Loop
    Agregando = False
    Fg1.Row = Row: Fg1.Col = 2:  Fg1.SetFocus
SALIR:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "Fg1_CellButtonClick"
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    Select Case Col
        Case 3, 4, 5
            If NulosN(Fg1.TextMatrix(Row, Col)) = 0 Then Fg1.TextMatrix(Row, Col) = 0
            Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "0.00000")
    End Select
    Exit Sub
error:
    SHOW_ERROR Me.Name, "fg1_CellChanged"
End Sub



Private Sub Fg1_EnterCell()
    If QueHace = 3 Or Fg1.Row < 1 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Fg1.Col >= 2 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
    
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then
        pRegistroAdd
    End If
    If KeyCode = 46 Then
        pRegistroDel
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    pConfigurarGrilla
    pCargarGrid
    '--agregando los meses
    Dim RsCons As New ADODB.Recordset
    RST_Busq RsCons, "SELECT id, descripcion From con_meses WHERE (((con_meses.id) Not In (0,13))) ORDER BY id", xCon
    Set dtcb_periodo.RowSource = RsCons
    dtcb_periodo.ListField = "descripcion"
    dtcb_periodo.BoundColumn = "id"
    Set RsCons = Nothing
    
    '--asignado los valores por defecto
    dtcb_periodo.BoundText = 1
    txt_ano.Text = AnoTra
    '--------
    SeEjecuto = True
    pCargarDatosDet
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado SPP, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            nuevo
        End If
    End If
    '-------
End Sub


Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    QueHace = 3
    Dg1.Columns("aporte").NumberFormat = "0.0000"
    Dg1.Columns("comision").NumberFormat = "0.0000"
    Dg1.Columns("prima").NumberFormat = "0.0000"
    TabOne1.CurrTab = 0
    Fg1.ColWidth(1) = 0
End Sub

Sub Blanquea()
    LimpiaText txt
    Fg1.Rows = Fg1.FixedRows
End Sub

Sub Bloquea(band As Boolean)
    CmdDetRestablece.Enabled = False
    habilitar CmdDet, band
    dtcb_periodo.Enabled = band
    UpDown1.Enabled = band
    txt_ano.Locked = Not band
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set RstFrm = Nothing
End Sub


Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Function Grabar() As Boolean

    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("¿Seguro desea Grabar?", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstDet As New ADODB.Recordset
    
    Dim xId&, A&

    On Error GoTo LaCague

    xCon.BeginTrans

    xCon.Execute "DELETE * FROM mae_regimenpendet WHERE (((mae_regimenpendet.anno)=" & NulosN(txt_ano.Text) & ") AND ((mae_regimenpendet.idmes)=" & NulosN(dtcb_periodo.BoundText) & "));"

    RST_Busq RstDet, "SELECT TOP 1 * FROM mae_regimenpendet ; ", xCon
    
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idregpen") = NulosN(Fg1.TextMatrix(A, 1))     '--regimen pensionario
        RstDet("anno") = NulosN(txt_ano.Text)
        RstDet("idmes") = dtcb_periodo.BoundText
        RstDet("aporte") = NulosN(Fg1.TextMatrix(A, 3))     '--aporte obligatorio
        RstDet("comision") = NulosN(Fg1.TextMatrix(A, 4))   '--comision variable
        RstDet("prima") = NulosN(Fg1.TextMatrix(A, 5))      '--prima de seguros
        RstDet.Update
    Next A

    MsgBox "Los datos grabaron con éxito", vbInformation, xTitulo

    xCon.CommitTrans
    Set RstDet = Nothing
    Grabar = True
    Exit Function
LaCague:
    Set RstDet = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo


End Function

Sub nuevo()
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea True
    Blanquea
'    pConfigurarGrilla
    
    Fg1.SelectionMode = flexSelectionFree
    
    Label5.Caption = "Agregando Aportes, Comisiones y Prima de Seguros"
    dtcb_periodo.SetFocus
End Sub

Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Label5.Caption = "Modificando Aportes, Comisiones y Prima de Seguros"

    ActivaTool
    
    Bloquea True
    
    If TabOne1.CurrTab = 0 Then TabOne1.CurrTab = 1
        
    TabOne1.TabEnabled(0) = False
    
    Fg1.SelectionMode = flexSelectionFree
        
    QueHace = 2
        
    Agregando = False
    
    CargarDefault = True
    
    dtcb_periodo.SetFocus

End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Eliminar()
    TabOne1.CurrTab = 0
    
    If RstFrm.RecordCount = 0 Then
        MsgBox "Ho hay registros a eliminar", vbInformation, xTitulo
        Exit Sub
    End If

    Dim RstBus As New ADODB.Recordset
    Dim nSQL As String
    nSQL = "SELECT mae_regimenpendet.idregpen, [mae_regimenpendet].[anno] & ' - ' & Format(CDate('01/' & [mae_regimenpendet].[idmes] & '/' & [mae_regimenpendet].[anno]),'mmm') AS periodo, mae_regimenpendet.anno, mae_regimenpendet.idmes " _
        + vbCr + " FROM mae_regimenpendet " _
        + vbCr + " WHERE (((mae_regimenpendet.idregpen) = " & RstFrm("id") & ")) " _
        + vbCr + " ORDER BY mae_regimenpendet.anno, mae_regimenpendet.idmes; "

    RST_Busq RstBus, nSQL, xCon
    If RstBus.RecordCount <> 0 Then
        MsgBox "No se puede eliminar el registro especificado." & vbCr & "Tiene Aportes, Comisiones y Prima de Seguros " & vbCr & "Periodo:  " & RstBus.Fields("periodo"), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstBus = Nothing
        Exit Sub
    End If
    Set RstBus = Nothing

    Dim Rpta As Integer
    Rpta = MsgBox("Esta seguro de eliminar el registro seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        
        xCon.Execute "DELETE * FROM mae_regimenpen WHERE id = " & RstFrm("id") & ""
        MsgBox "El registro se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstFrm.Requery
        Dg1.Refresh
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    If Button.Index = 5 Then Cancelar
    If Button.Index = 6 Then
        If Grabar = True Then
            pCargarGrid
            Cancelar
        End If
    End If
    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then
        RstFrm.Filter = adFilterNone
    End If
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 13 Or Button.Index = 14 Then
    
        
        If Button.Index = 13 Then pExportarExcel
        If Button.Index = 14 Then pImprimir
    End If
    If Button.Index = 16 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0

    Dim xRs As New ADODB.Recordset
    Dim xCampos(5, 4) As String
    Dim nSQL As String
    
    xCampos(0, 0) = "Mes":      xCampos(0, 1) = "mes":        xCampos(0, 2) = "800":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "AFP":      xCampos(1, 1) = "descripcion": xCampos(1, 2) = "3500":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Aporte":   xCampos(2, 1) = "aporte":     xCampos(2, 2) = "1100":   xCampos(2, 3) = "n"
    xCampos(3, 0) = "Comisión": xCampos(3, 1) = "comision":   xCampos(3, 2) = "1100":   xCampos(3, 3) = "n"
    xCampos(4, 0) = "Prima":    xCampos(4, 1) = "prima":      xCampos(4, 2) = "1100":   xCampos(4, 3) = "N"
    
    nSQL = "SELECT Format(CDate('01/' & [mae_regimenpendet].[idmes] & '/' & [mae_regimenpendet].[anno]),'mmm') AS mes, [mae_regimenpendet].[idregpen] & [mae_regimenpendet].[idmes] AS codigo, mae_regimenpendet.idregpen, mae_regimenpendet.idmes, mae_regimenpen.codsun, mae_regimenpen.descripcion, mae_regimenpendet.aporte, mae_regimenpendet.comision, mae_regimenpendet.prima " _
        + vbCr + " FROM mae_regimenpen INNER JOIN mae_regimenpendet ON mae_regimenpen.id = mae_regimenpendet.idregpen " _
        + vbCr + " Where ((mae_regimenpendet.anno) = " & AnoTra & ") " _
        + vbCr + " ORDER BY mae_regimenpendet.idmes DESC , mae_regimenpen.descripcion;"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Aportes, Comisiones y Prima de Seguros ", "descripcion", "descripcion", Principio

    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    RstFrm.MoveFirst
    RstFrm.Find "codigo = '" & CStr(xRs("codigo")) & "'"
SALIR:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

Sub Filtrar()
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(5, 3) As String
    xCampos(0, 0) = "Mes":      xCampos(0, 1) = "mes":        xCampos(0, 2) = "C":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "AFP":      xCampos(1, 1) = "descrpcion": xCampos(1, 2) = "C":   xCampos(1, 3) = "3500"
    xCampos(2, 0) = "Aporte Obligatorio":   xCampos(2, 1) = "aporte":     xCampos(2, 2) = "N":   xCampos(2, 3) = "1100"
    xCampos(3, 0) = "Comisión Variable": xCampos(3, 1) = "comision":   xCampos(3, 2) = "N":   xCampos(3, 3) = "1100"
    xCampos(4, 0) = "Prima de Seguros":    xCampos(4, 1) = "prima":      xCampos(4, 2) = "N":   xCampos(4, 3) = "1100"

    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1

End Sub

Sub MuestraSegundoTab()
    Fg1.Rows = 1
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros para Mostrar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If

    '-------------------------------------
    txt_ano.Text = AnoTra
    dtcb_periodo.BoundText = NulosN(RstFrm.Fields("idmes"))
    '-------------------------------------
    pCargarDatosDet
    
End Sub


Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
       
    nSQL = "SELECT Format(CDate('01/' & [mae_regimenpendet].[idmes] & '/' & [mae_regimenpendet].[anno]),'mmm') AS mes, [mae_regimenpendet].[idregpen] & [mae_regimenpendet].[idmes] AS codigo, mae_regimenpendet.idregpen, mae_regimenpendet.idmes, mae_regimenpen.codsun, mae_regimenpen.descripcion, mae_regimenpendet.aporte, mae_regimenpendet.comision, mae_regimenpendet.prima " _
        + vbCr + " FROM mae_regimenpen INNER JOIN mae_regimenpendet ON mae_regimenpen.id = mae_regimenpendet.idregpen " _
        + vbCr + " WHERE ((mae_regimenpendet.anno) = " & AnoTra & ") " _
        + vbCr + " ORDER BY mae_regimenpendet.idmes DESC , mae_regimenpen.descripcion;"

    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg1.DataSource = RstFrm
    TabOne1.CurrTab = 0
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Function fValidarDatos() As Boolean
    Dim band As Integer
    '--de los porcentajes
    If NulosN(txt_ano.Text) = 0 Then
        MsgBox "Vuelva a seleccionar el Año", vbExclamation, xTitulo
        Exit Function
    End If
    If NulosN(dtcb_periodo.BoundText) = 0 Then
        MsgBox "Vuelva a seleccionar Periodo", vbExclamation, xTitulo
        Exit Function
    End If
'    If Fg1.Rows = 1 Then
'        MsgBox "No hay Registros", vbExclamation, xTitulo
'        CmdDet(0).SetFocus
'        Exit Function
'    End If
    Dim mRow&, mCol&
    mCol = -1
    For mRow = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 1)) = 0 Then '--descripcion
            MsgBox "Falta especificar el Sistema Privado de Pensiones", vbExclamation, xTitulo
            mCol = 2:          Exit For
        End If
        If Fg1.TextMatrix(mRow, 3) = "" Then  '--aporte obligatorio
            MsgBox "Falta especificar el Aporte Obligatorio" + vbCr + "Descripción: " & Fg1.TextMatrix(mRow, 2), vbExclamation, xTitulo
            mCol = 3:          Exit For
        End If
        If Fg1.TextMatrix(mRow, 4) = "" Then  '--Comision Variable
            MsgBox "Falta especificar la Comisión Variable" + vbCr + "Descripción: " & Fg1.TextMatrix(mRow, 2), vbExclamation, xTitulo
            mCol = 4:          Exit For
        End If
        If Fg1.TextMatrix(mRow, 5) = "" Then  '--aporte obligatorio
            MsgBox "Falta especificar la Prima de Seguros" + vbCr + "Descripción: " & Fg1.TextMatrix(mRow, 2), vbExclamation, xTitulo
            mCol = 5:          Exit For
        End If
    Next mRow
    If mCol <> -1 Then
        Agregando = True:  Fg1.Row = mRow: Fg1.Col = mCol: Agregando = False
        Fg1.SetFocus
        Exit Function
    End If
    '--
    fValidarDatos = True
    
End Function

Private Sub pConfigurarGrilla()
    With Fg1 '--
        .Rows = 1
        .Cols = 6
        .FixedRows = 1
        .RowHeight(0) = 500
        .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "Descripción":                    .ColWidth(2) = 3500: .ColAlignment(2) = flexAlignLeftCenter:   .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "Aporte" + vbCr + "Obligatorio":  .ColWidth(3) = 1100: .ColAlignment(3) = flexAlignRightCenter:  .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 4) = "Comisión" + vbCr + "Variable":   .ColWidth(4) = 1100: .ColAlignment(4) = flexAlignRightCenter:  .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 5) = "Prima" + vbCr + "de Seguros":    .ColWidth(5) = 1100: .ColAlignment(5) = flexAlignRightCenter:  .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterCenter
        .ColFormat(3) = "#.#####"
        .ColFormat(4) = "#.#####"
        .ColFormat(5) = "#.#####"
        
        .ColEditMask(3) = "#.#####"
        .ColEditMask(4) = "#.#####"
        .ColEditMask(5) = "#.#####"
        
        .SelectionMode = flexSelectionByRow
    End With
    
    '*****************************************
    GRID_COMBOLIST Fg1, 2
    DoEvents
End Sub

Private Sub pRegistroAdd()
    Dim mCol%
    If QueHace = 3 Then Exit Sub
    Agregando = True
    If Fg1.Rows > 1 Then
        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = 0 Then
            MsgBox "Falta Completar...", vbExclamation, xTitulo
        Else
            Fg1.AddItem ""
        End If
    Else
        Fg1.AddItem ""
    End If
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = 2
    Fg1.SetFocus
    Agregando = False
End Sub

Private Sub pRegistroDel()
    If Fg1.Rows = 1 Then Exit Sub
    If Fg1.Row < 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    Fg1.RemoveItem Fg1.Row
    If Fg1.Rows = 1 Then
        CmdDet(0).SetFocus
    Else
        Fg1.Row = Fg1.Rows - 1
        Fg1.Col = 2
        Fg1.SetFocus
    End If
End Sub

'****************************************************************************************

Private Sub pImprimir()
    On Error GoTo error
    Me.MousePointer = vbHourglass
    Dim oPrint As New SGI2_funciones.formularios
    oPrint.Imprimir_x_VSFlexGrid Fg1, "Consulta de Aportes, Comisiones y Prima de Seguros", , "Periodo: " + dtcb_periodo.Text & " - " & txt_ano.Text, False, True
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

Private Sub pExportarExcel()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Consulta de Aportes, Comisiones y Prima de Seguros", "Periodo: " + dtcb_periodo.Text & " - " & txt_ano.Text, "", "%"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub

Public Sub pCargarDatosDet()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    If SeEjecuto = False Then Exit Sub

    If NulosN(txt_ano.Text) = 0 Or NulosN(dtcb_periodo.BoundText) = 0 Then
        MsgBox "Vuelva a seleccionar El Año o Periodo", vbExclamation, xTitulo
        Exit Sub
    End If
    
    '--aportes, comisiones y prima de seguros
    nSQL = "SELECT mae_regimenpendet.idregpen,mae_regimenpen.descripcion, mae_regimenpendet.aporte, mae_regimenpendet.comision, mae_regimenpendet.prima " _
        + vbCr + " FROM mae_regimenpen INNER JOIN mae_regimenpendet ON mae_regimenpen.id = mae_regimenpendet.idregpen " _
        + vbCr + " Where (((mae_regimenpendet.anno) = " & NulosN(txt_ano.Text) & ") And ((mae_regimenpendet.idmes) = " & NulosN(dtcb_periodo.BoundText) & ")) " _
        + vbCr + " ORDER BY mae_regimenpen.descripcion;"

    RST_Busq RstTmp, nSQL, xCon
    Fg1.Rows = 1
    
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
        CmdDetRestablece.Enabled = False
    Else
        CmdDetRestablece.Enabled = True
    End If
    Do While Not RstTmp.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosN(RstTmp("idregpen"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstTmp("descripcion"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosN(RstTmp("aporte"))
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosN(RstTmp("comision"))
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosN(RstTmp("prima"))
        RstTmp.MoveNext
    Loop
    
    Set RstTmp = Nothing
End Sub



Private Sub CmdDet_Click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            pRegistroAdd
        Case 1 '--eliminar
            pRegistroDel
    End Select
End Sub

Private Sub txt_ano_Change()
    If SeEjecuto = False Then Exit Sub
    If NulosN(txt_ano.Text) < 2005 Then Exit Sub
    pCargarDatosDet
End Sub

Private Sub UpDown1_Change()
    txt_ano.Text = UpDown1.Value
End Sub

Private Sub dtcb_periodo_Change()
    If SeEjecuto = False Then Exit Sub
    If dtcb_periodo.MatchedWithList = False Then Exit Sub
    pCargarDatosDet
End Sub



Private Sub CmdDetRestablece_Click()
    Dim nPeriodo As String
    Dim nSQL As String
    Dim mMes&, mAnno&
    mMes = dtcb_periodo.BoundText
    mAnno = NulosN(txt_ano.Text)
    
    nPeriodo = "Año: " & IIf(mMes = 1, mAnno - 1, mAnno) & "   Mes: " & Busca_Codigo(IIf(mMes = 1, 12, mMes - 1), "id", "descripcion", "con_meses", "N", xCon)

'    '--eliminando registros si esta en bd
'    nSQL = "DELETE FROM mae_regimenpendet WHERE anno=" & mAnno & " AND idmes= " & mMes & "; "
'    xCon.Execute nSQL
'    '--insertando de nuevo los registros
'    nSQL = "INSERT INTO mae_regimenpendet (idregpen,anno,idmes,aporte,comision,prima) " _
'        + vbCr + " SELECT idregpen, IIF(idmes=12,anno+1,anno), IIF(idmes=12,1,idmes+1) ,aporte,comision,prima " _
'        + vbCr + " FROM mae_regimenpendet WHERE anno=" & IIf(mMes = 1, mAnno - 1, mAnno) & " AND idmes= " & IIf(mMes = 1, 12, mMes - 1) & ""
'    xCon.Execute nSQL
    
    
    Dim RstTmp As New ADODB.Recordset
    nSQL = "SELECT mae_regimenpendet.idregpen, mae_regimenpen.descripcion, IIf(idmes=12,anno+1,anno) AS Expr1, IIf(idmes=12,1,idmes+1) AS Expr2, mae_regimenpendet.aporte, mae_regimenpendet.comision, mae_regimenpendet.prima " _
        + vbCr + " FROM mae_regimenpen INNER JOIN mae_regimenpendet ON mae_regimenpen.id = mae_regimenpendet.idregpen WHERE anno=" & IIf(mMes = 1, mAnno - 1, mAnno) & " AND idmes= " & IIf(mMes = 1, 12, mMes - 1) & ""
    
    RST_Busq RstTmp, nSQL, xCon
    
    Fg1.Rows = 1
    
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
        CmdDetRestablece.Enabled = False
    Else
        CmdDetRestablece.Enabled = True
    End If
    
    Do While Not RstTmp.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosN(RstTmp("idregpen"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstTmp("descripcion"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosN(RstTmp("aporte"))
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosN(RstTmp("comision"))
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosN(RstTmp("prima"))
        RstTmp.MoveNext
    Loop
    
    Set RstTmp = Nothing
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay información en el " & vbCr & nPeriodo, vbInformation, xTitulo
    Else
        MsgBox "La Información que se muestra es Duplicado de..." + vbCr + nPeriodo, vbQuestion + vbYesNo + vbDefaultButton1, xTitulo
    End If
End Sub


