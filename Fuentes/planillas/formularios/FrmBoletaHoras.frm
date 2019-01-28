VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmBoletaHoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Control de Asistencia"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11850
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoletaHoras.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoletaHoras.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoletaHoras.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoletaHoras.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoletaHoras.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoletaHoras.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoletaHoras.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoletaHoras.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoletaHoras.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoletaHoras.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBoletaHoras.frx":1EA4
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
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   1005
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
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Height          =   7215
      Left            =   30
      TabIndex        =   1
      Top             =   360
      Width           =   11835
      _cx             =   20876
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
         Left            =   -12390
         TabIndex        =   4
         Top             =   375
         Width           =   11745
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6405
            Left            =   45
            TabIndex        =   5
            Top             =   390
            Width           =   11565
            _ExtentX        =   20399
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
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Proceso"
            Columns(2).DataField=   "proceso"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "N°"
            Columns(3).DataField=   "numero"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch.Inicio"
            Columns(4).DataField=   "fchini"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fch.Fin"
            Columns(5).DataField=   "fchfin"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=926"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=5424"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5345"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2461"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2381"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=1402"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1323"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=1852"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1773"
            Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(28)=   "Column(5).Width=1773"
            Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=1693"
            Splits(0)._ColumnProps(31)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
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
            _StyleDefs(47)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
            _StyleDefs(48)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(51)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(52)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(70)  =   "Named:id=33:Normal"
            _StyleDefs(71)  =   ":id=33,.parent=0"
            _StyleDefs(72)  =   "Named:id=34:Heading"
            _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(74)  =   ":id=34,.wraptext=-1"
            _StyleDefs(75)  =   "Named:id=35:Footing"
            _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(77)  =   "Named:id=36:Selected"
            _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=37:Caption"
            _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(81)  =   "Named:id=38:HighlightRow"
            _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(83)  =   "Named:id=39:EvenRow"
            _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(85)  =   "Named:id=40:OddRow"
            _StyleDefs(86)  =   ":id=40,.parent=33"
            _StyleDefs(87)  =   "Named:id=41:RecordSelector"
            _StyleDefs(88)  =   ":id=41,.parent=34"
            _StyleDefs(89)  =   "Named:id=42:FilterBar"
            _StyleDefs(90)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Control de Asistencia"
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
            TabIndex        =   6
            Top             =   60
            Width           =   11400
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6795
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   11745
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1125
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   29
            Text            =   "txt(1)"
            Top             =   1035
            Width           =   3615
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   5895
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   28
            Text            =   "txt(2)"
            Top             =   1035
            Width           =   765
         End
         Begin VB.CommandButton cmdCargarPersonal 
            Caption         =   "Obtener Datos de Asistencia"
            Enabled         =   0   'False
            Height          =   450
            Left            =   180
            TabIndex        =   27
            Top             =   1410
            Width           =   2535
         End
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   1
            Left            =   6420
            Picture         =   "FrmBoletaHoras.frx":23EC
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Seleccione la Categoría de Concepto"
            Top             =   735
            Width           =   210
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   1
            Left            =   5895
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   22
            Text            =   "txt_cb(1)"
            Top             =   705
            Width           =   765
         End
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   0
            Left            =   1650
            Picture         =   "FrmBoletaHoras.frx":251E
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Seleccione la Categoría de Concepto"
            Top             =   735
            Width           =   210
         End
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Caption         =   "Frame12"
            Height          =   585
            Index           =   0
            Left            =   180
            TabIndex        =   9
            Top             =   6135
            Width           =   11460
            Begin VB.CommandButton CmdDet 
               Caption         =   "&Eliminar"
               Enabled         =   0   'False
               Height          =   435
               Index           =   1
               Left            =   1560
               TabIndex        =   11
               Top             =   60
               Width           =   1395
            End
            Begin VB.CommandButton CmdDet 
               Caption         =   "&Agregar"
               Enabled         =   0   'False
               Height          =   435
               Index           =   0
               Left            =   120
               TabIndex        =   10
               Top             =   60
               Width           =   1395
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
            Begin VB.Line lin 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Index           =   3
               X1              =   -30
               X2              =   13000
               Y1              =   570
               Y2              =   570
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Index           =   3
               X1              =   11445
               X2              =   11445
               Y1              =   0
               Y2              =   985
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               X1              =   15
               X2              =   15
               Y1              =   0
               Y2              =   1000
            End
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   285
            Index           =   0
            Left            =   10665
            TabIndex        =   7
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   -15
            Visible         =   0   'False
            Width           =   945
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   4140
            Left            =   180
            TabIndex        =   12
            Top             =   1890
            Width           =   11460
            _cx             =   20214
            _cy             =   7302
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
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmBoletaHoras.frx":2650
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
         Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
            Height          =   300
            Index           =   0
            Left            =   1125
            TabIndex        =   18
            Top             =   360
            Width           =   1350
            _ExtentX        =   2381
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
            Valor           =   "14/11/2008"
         End
         Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
            Height          =   300
            Index           =   1
            Left            =   3390
            TabIndex        =   19
            Top             =   360
            Width           =   1350
            _ExtentX        =   2381
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
            Valor           =   "14/11/2008"
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   1125
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   14
            Text            =   "txt_cb(0)"
            Top             =   705
            Width           =   765
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   31
            Top             =   1140
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número"
            Height          =   195
            Index           =   2
            Left            =   4875
            TabIndex        =   30
            Top             =   1140
            Width           =   555
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cod(1)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   8445
            TabIndex        =   26
            Top             =   690
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar Trabajo"
            Height          =   195
            Index           =   1
            Left            =   4890
            TabIndex        =   25
            Top             =   810
            Width           =   990
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(1)"
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
            Index           =   1
            Left            =   6660
            TabIndex        =   24
            Top             =   705
            Width           =   3180
         End
         Begin VB.Label lbl_fecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fch. Inicio"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   21
            Top             =   495
            Width           =   735
         End
         Begin VB.Label lbl_fecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fch. Fin"
            Height          =   195
            Index           =   1
            Left            =   2625
            TabIndex        =   20
            Top             =   480
            Width           =   570
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(0)"
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
            Index           =   0
            Left            =   1890
            TabIndex        =   17
            Top             =   705
            Width           =   2850
         End
         Begin VB.Label lbl_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Pago"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   16
            Top             =   810
            Width           =   735
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cod(0)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   3645
            TabIndex        =   15
            Top             =   705
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   10125
            TabIndex        =   8
            Top             =   75
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Control de Asistencia"
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
            TabIndex        =   3
            Top             =   60
            Width           =   11400
         End
      End
   End
End
Attribute VB_Name = "FrmBoletaHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstFrm As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim Agregando As Boolean

Dim fOrdenLista As Boolean          '--especfica el orden de la lista de la consulta
Dim mColInicioEdicion As Integer    '--nos indica que la posicion de la columna donde empezara a editar la grilla

Private Enum eConfigurarGrid
    eEncabezado = 0
    eObtener = 1
    eMostrar = 2
End Enum


Sub Cancelar()
    ActivaTool
    Label5.Caption = "Detalle del Control de Asistencia"
    QueHace = 3
    Bloquea False
    Fg1.SelectionMode = flexSelectionByRow
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub



Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub


Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub


Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    Select Case Col
        Case mColInicioEdicion, mColInicioEdicion - 1
            If IsNumeric(Fg1.TextMatrix(Row, Col)) = False Then Fg1.TextMatrix(Row, Col) = ""
        Case Is > mColInicioEdicion
            Fg1.TextMatrix(Row, Col) = Replace(Fg1.TextMatrix(Row, Col), "::", ":")
            If Left(Fg1.TextMatrix(Row, Col), 1) = ":" Then Fg1.TextMatrix(Row, Col) = Mid(Fg1.TextMatrix(Row, Col), 2)
            If IsNumeric(Fg1.TextMatrix(Row, Col)) = True Then
                Fg1.TextMatrix(Row, Col) = Fg1.TextMatrix(Row, Col) & ":00:00"
            End If
            Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "HH:MM:SS")
    End Select
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col < mColInicioEdicion - 1 Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Col
        Case mColInicioEdicion, mColInicioEdicion - 1
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Is > mColInicioEdicion
            If validar_numero(KeyAscii) = False And KeyAscii <> 58 Then KeyAscii = 0
    End Select
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
   
    SeEjecuto = False
    '--similar a configurar grilla
    pCargarDatosDet CDate("01/01/00"), CDate("02/01/00"), eEncabezado

    pCargarGrid
    
    SeEjecuto = True
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado ningún Control de Asistencia, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            nuevo
        End If
    End If
    
    
End Sub


Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    QueHace = 3

    TabOne1.CurrTab = 0
    txtfecha(0).Valor = Date
    txtfecha(1).Valor = Date
End Sub


Sub Blanquea()
    LimpiaText txt_cb
    LimpiaText txtfecha
    LimpiaText txt
    Fg1.Rows = Fg1.FixedRows
End Sub

Sub Bloquea(band As Boolean)
    habilitar_Locked txt_cb, Not band
    habilitar_Locked txt, Not band
    habilitar_Locked txtfecha, Not band
    habilitar CmdDet, band
    cmdCargarPersonal.Enabled = band
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
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " el Control de Asistencia", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim RstDetEmp As New ADODB.Recordset
    Dim RstDetHor As New ADODB.Recordset
   
    Dim xId&, mRow&, mCol&

    On Error GoTo LaCague

    xCon.BeginTrans

    If QueHace = 1 Then

        xId = HallaCodigoTabla("pla_marcaresum", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_marcaresum", xCon

        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM pla_marcaresum WHERE id = " & xId & "", xCon
        '--eliminando las horas de los empleados relacionado al proceso
        xCon.Execute "Delete from pla_marcaresumhor where idres =  " & xId & " ;"
        '--eliminando los empleados relacionado al proceso
        xCon.Execute "Delete from pla_marcaresumemp where idres =  " & xId & " ;"
                
    End If

    RST_Busq RstDetEmp, "SELECT TOP 1 * FROM pla_marcaresumemp ; ", xCon
    RST_Busq RstDetHor, "SELECT TOP 1 * FROM pla_marcaresumhor ; ", xCon

    RstCab("ano") = AnoTra
    RstCab("fchini") = CDate(txtfecha(0).Valor)
    RstCab("fchfin") = CDate(txtfecha(1).Valor)
    RstCab("idproc") = NulosC(lbl_cod(0).Caption)
    RstCab("idlugtra") = NulosN(lbl_cod(1).Caption)
    RstCab("descripcion") = NulosC(txt(1).Text)
    RstCab("numero") = NulosN(txt(2).Text)
    
    
    
    RstCab.Update
    
    '--detalle de conceptos
    For mRow = Fg1.FixedRows To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 1)) <> 0 Then '--solo que tengan codigo de emplado
            RstDetEmp.AddNew
            RstDetEmp("idres") = xId
            RstDetEmp("idemp") = NulosN(Fg1.TextMatrix(mRow, 1))
            RstDetEmp("totdiatra") = NulosN(Fg1.TextMatrix(mRow, mColInicioEdicion - 1))
            RstDetEmp("totdiafer") = NulosN(Fg1.TextMatrix(mRow, mColInicioEdicion))
            RstDetEmp.Update
            '--detalle de las horas
            For mCol = mColInicioEdicion + 1 To Fg1.Cols - 2 'menos la ultima col=total
                If NulosN(Fg1.TextMatrix(2, mCol)) <> 0 And NulosC(Fg1.TextMatrix(mRow, mCol)) <> "" Then
                    RstDetHor.AddNew
                    RstDetHor("idres") = xId
                    RstDetHor("idemp") = NulosN(Fg1.TextMatrix(mRow, 1))
                    RstDetHor("idhora") = NulosN(Fg1.TextMatrix(2, mCol)) '--codigo del tipo de hora
                    RstDetHor("tothor") = NulosC(Fg1.TextMatrix(mRow, mCol))
                    RstDetHor("totseg") = ConvertSeg(Fg1.TextMatrix(mRow, mCol))
                    RstDetHor.Update
                End If
            Next mCol
        End If
    Next mRow
    xCon.CommitTrans
    
    MsgBox "Los datos del Control de Asistencia se " + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

    Set RstCab = Nothing:    Set RstDetEmp = Nothing:    Set RstDetHor = Nothing
    Grabar = True
    Exit Function
LaCague:
    Set RstCab = Nothing:    Set RstDetEmp = Nothing:    Set RstDetHor = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar registro por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Sub nuevo()
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea True
    Blanquea
    Fg1.SelectionMode = flexSelectionFree
    Label5.Caption = "Agregando Control de Asistencia"
    txtfecha(0).SetFocus
End Sub

Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    Label5.Caption = "Modificando del Control de Asistencia"
    
    ActivaTool
    
    Bloquea True
    Fg1.SelectionMode = flexSelectionFree
    If TabOne1.CurrTab = 0 Then TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    QueHace = 2
    
    Agregando = False
    
    txt(1).SetFocus
    
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
        MsgBox "No hay Registros para Eliminar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Dim RstBus As New ADODB.Recordset
    '--ver si hay documento
    RST_Busq RstBus, "SELECT * FROM pla_boleta WHERE idproc = " & RstFrm("id") & "", xCon
    If RstFrm.RecordCount <> 0 Then
        MsgBox "No se puede eliminar el documento, pues se ha generado pagos...", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstBus = Nothing
        Exit Sub
    End If
    '--ver si hay concepto asociado a algun proceso.
    RST_Busq RstBus, "SELECT * FROM pla_conceptoproc WHERE idproc = " & RstFrm("id") & "", xCon
    If RstFrm.RecordCount <> 0 Then
        If MsgBox("El Proceso que desea eliminar tiene Conceptos Asignados" + vbCr + "Seguro que desea eliminar ", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then
            Set RstBus = Nothing
            Exit Sub
        End If
    End If
    Set RstBus = Nothing
    
    Dim Rpta As Integer

    Rpta = MsgBox("Esta seguro de eliminar el Proceso seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE FROM pla_conceptoproc WHERE idproc=" & RstFrm.Fields("id") & " ;"
        xCon.Execute "DELETE FROM pla_proceso WHERE id = " & RstFrm("id") & ""
        MsgBox "El documento se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
            RstFrm.Requery
            Dg1.Refresh
            Cancelar
        End If
    End If
    If Button.Index = 10 Then Buscar
    If Button.Index = 14 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub


Sub Buscar()
    TabOne1.CurrTab = 0
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim xCampos(2, 4) As String

    xCampos(0, 0) = "Descripción": xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "7000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":          xCampos(1, 1) = "id":            xCampos(1, 2) = "700":    xCampos(1, 3) = "N"

    nSQL = "SELECT pla_proceso.* FROM pla_proceso; "

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscar Documentos", "descripcion", "descripcion", Principio
    
    If xRs.State = 1 Then
        RstFrm.MoveFirst
        RstFrm.Find "id = " & xRs("id") & ""
    End If
    Set xRs = Nothing
    
End Sub

Sub MuestraSegundoTab()
    Dim QueHaceTmp  As Integer
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros para Mostrar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    Blanquea
    QueHaceTmp = QueHace
    QueHace = -1
    txt(0).Text = NulosN(RstFrm("id"))
    txt(1).Text = NulosC(RstFrm("descripcion"))
    txt(2).Text = NulosC(RstFrm("numero"))
    If IsDate(RstFrm("fchini")) = True Then txtfecha(0).Valor = CDate(RstFrm("fchini"))
    If IsDate(RstFrm("fchfin")) = True Then txtfecha(1).Valor = CDate(RstFrm("fchfin"))
    '--proceso
    If NulosN(RstFrm("idproc")) <> 0 Then
        txt_cb(0).Text = NulosN(RstFrm("idproc"))
        txt_cb_Validate 0, False
    End If
    '--lugar de trabajo
    If NulosN(RstFrm("idlugtra")) <> 0 Then
        txt_cb(1).Text = NulosN(RstFrm("idlugtra"))
        txt_cb_Validate 1, False
    End If
    QueHace = QueHaceTmp
    pCargarDatosDet Date, Date, eMostrar
    
End Sub


Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL As String
    
    nSQL = "SELECT pla_marcaresum.*, pla_proceso.abrev AS proceso " _
        + vbCr + " FROM pla_marcaresum LEFT JOIN pla_proceso ON pla_marcaresum.idproc = pla_proceso.id " _
        + vbCr + " Where (((pla_marcaresum.ano) = " & AnoTra & ")) " _
        + vbCr + " ORDER BY pla_marcaresum.fchini;"


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
    If IsDate(txtfecha(0).Valor) = False Or IsDate(txtfecha(1).Valor) = False Then
        MsgBox "Ingrese la Fecha " & IIf(IsDate(txtfecha(0).Valor) = True, "de Inicio", "Final") & " ", vbExclamation, xTitulo
        txtfecha(0).SetFocus
        Exit Function
    End If
    
    band = Validar(txt)
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl(band).Caption, vbInformation, xTitulo
       txt(band).SetFocus
       Exit Function
    End If
    
    '--
    fValidarDatos = True
    
End Function

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub pCargarDatosDet(fchini As Date, fchfin As Date, eTipo As eConfigurarGrid)
    Dim RstTmp As New ADODB.Recordset
    Dim RstDiaTra As New ADODB.Recordset
    Dim RstDiaFer As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLIdTrabajo As String
    
    On Error GoTo error
    '--si hay lugar de trabajo
    If NulosN(lbl_cod(1).Caption) <> 0 Then nSQLIdTrabajo = " AND pla_empleados.idlugtra=" & NulosN(lbl_cod(1).Caption)
    '----
    If eTipo = eMostrar Then
        nSQL = "TRANSFORM Sum(pla_marcaresumhor.totseg) AS SumaDetotseg " _
            + vbCr + " SELECT  pla_marcaresumhor.idemp,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_cargo.descripcion AS cargo, mae_empresalugtra.descripcion AS lugartra, Last(pla_periodolaboral.fchini) AS ingreso,pla_marcaresumemp.totdiatra, pla_marcaresumemp.totdiafer,  Sum(pla_marcaresumhor.totseg) AS total " _
            + vbCr + " FROM (mae_tipohora RIGHT JOIN ((pla_marcaresumemp INNER JOIN (pla_marcaresumhor INNER JOIN (mae_cargo RIGHT JOIN pla_empleados ON mae_cargo.id = pla_empleados.idcargo) ON pla_marcaresumhor.idemp = pla_empleados.id) ON (pla_marcaresumemp.idres = pla_marcaresumhor.idres) AND (pla_marcaresumemp.idemp = pla_marcaresumhor.idemp)) LEFT JOIN mae_empresalugtra ON pla_empleados.idlugtra = mae_empresalugtra.id) ON mae_tipohora.id = pla_marcaresumhor.idhora) INNER JOIN pla_periodolaboral ON pla_empleados.id = pla_periodolaboral.idemp " _
            + vbCr + " Where (((pla_marcaresumemp.idres) = " & NulosN(RstFrm("id")) & ")) " _
            + vbCr + " GROUP BY  pla_marcaresumhor.idemp,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom], mae_cargo.descripcion, mae_empresalugtra.descripcion, pla_marcaresumemp.totdiatra, pla_marcaresumemp.totdiafer " _
            + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] " _
            + vbCr + " PIVOT mae_tipohora.id In (15,2,3,13,12,17,11);"
    Else
        '----------------------------------------
        '--cargar lista del personal con los dias trabajados
        '--Consignar la cantidad de días efectivamente trabajados durante el período. Asimismo,
        '--considerar como tales el día de cese del trabajador (caso de baja), los días feriados y los días de descanso semanal obligatorio.
        nSQL = "SELECT  diastra.idemp, count (diastra.trabajo) as tottra " _
            + vbCr + " From " _
            + vbCr + " (SELECT DISTINCT pla_marcacionhora.idemp, pla_marcacion.dia AS trabajo " _
            + vbCr + " FROM pla_marcacion INNER JOIN (mae_tipohora INNER JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora) ON pla_marcacion.id = pla_marcacionhora.idmarca " _
            + vbCr + " WHERE (((pla_marcacionhora.idhora) Not In (10)) AND ((mae_tipohora.hortrabajo)=-1)) " _
            + vbCr + " GROUP BY  pla_marcacion.dia,pla_marcacionhora.idemp " _
            + vbCr + " HAVING (((pla_marcacion.dia) Between CDate('" & fchini & "') And CDate('" & fchfin & "'))) " _
            + vbCr + " ) AS diastra " _
            + vbCr + " GROUP BY diastra.idemp; "
        RST_Busq RstDiaTra, nSQL, xCon
        '----------------------------------------
        '--cargar lista del personal con los dias feriado
        nSQL = "SELECT  diasfer.idemp,  count (diasfer.feriado) as totfer " _
            + vbCr + " From " _
            + vbCr + " (SELECT DISTINCT pla_marcacionhora.idemp, pla_marcacion.dia AS feriado " _
            + vbCr + " FROM mae_diasfestivos, pla_marcacion INNER JOIN (mae_tipohora INNER JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora) ON pla_marcacion.id = pla_marcacionhora.idmarca " _
            + vbCr + " WHERE (((mae_tipohora.id) In (10))) " _
            + vbCr + " GROUP BY pla_marcacionhora.idemp, pla_marcacion.dia " _
            + vbCr + " HAVING (((pla_marcacion.dia) Between CDate('" & fchini & "') And CDate('" & fchfin & "'))) " _
            + vbCr + " ) as diasfer " _
            + vbCr + " GROUP BY diasfer.idemp "
        RST_Busq RstDiaFer, nSQL, xCon
        '----------------------------------------
    
        nSQL = "TRANSFORM Sum(tipohora.imptot) AS SumaDeimptot " _
            + vbCr + " SELECT emp.id as idemp , emp.nombres, emp.cargo, emp.lugartra,emp.ingreso, emp.totdiatra, emp.totdiafer, Sum(tipohora.imptot) AS total " _
            + vbCr + " From " _
            + vbCr + " (SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_cargo.descripcion AS cargo, mae_empresalugtra.descripcion AS lugartra, Last(pla_periodolaboral.fchini) AS ingreso, '' AS totdiatra, '' AS totdiafer " _
            + vbCr + " FROM ((mae_cargo RIGHT JOIN pla_empleados ON mae_cargo.id = pla_empleados.idcargo) LEFT JOIN mae_empresalugtra ON pla_empleados.idlugtra = mae_empresalugtra.id) LEFT JOIN pla_periodolaboral ON pla_empleados.id = pla_periodolaboral.idemp " _
            + vbCr + " WHERE (((pla_empleados.idbolpag)=" & NulosN(lbl_cod(0).Caption) & ")) " & nSQLIdTrabajo & " AND pla_periodolaboral.fchfin Is Null GROUP BY pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom], mae_cargo.descripcion, mae_empresalugtra.descripcion ) AS emp " _
            + vbCr + " Left Join " _
            + vbCr + " (SELECT * FROM " _
            + vbCr + " (SELECT pla_marcacionhora.idemp, mae_tipohora.id AS idhora, mae_tipohora.nomcor AS concepto, Sum(pla_marcacionhora.totseg) AS imptot " _
            + vbCr + " FROM pla_marcacion RIGHT JOIN (mae_tipohora LEFT JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora) ON pla_marcacion.id = pla_marcacionhora.idmarca " _
            + vbCr + " WHERE (((pla_marcacion.dia) Between CDate('" & fchini & "') And CDate('" & fchfin & "')) AND ((mae_tipohora.id) In (2,3,11,12,13,17))) " _
            + vbCr + " GROUP BY pla_marcacionhora.idemp, mae_tipohora.id, mae_tipohora.nomcor, mae_tipohora.id, mae_tipohora.variable "
        nSQL = nSQL _
            + vbCr + " Union " _
            + vbCr + " SELECT pla_marcacionhora.idemp, mae_tipohora_1.id AS idhora, mae_tipohora_1.nomcor AS concepto, Sum(pla_marcacionhora.totseg) AS imptot " _
            + vbCr + " FROM mae_tipohora AS mae_tipohora_1, pla_marcacion INNER JOIN (mae_tipohora INNER JOIN pla_marcacionhora ON mae_tipohora.id = pla_marcacionhora.idhora) ON pla_marcacion.id = pla_marcacionhora.idmarca " _
            + vbCr + " WHERE (((pla_marcacion.dia) Between CDate('" & fchini & "') And CDate('" & fchfin & "')) AND ((mae_tipohora.hortrabajo)=-1)) " _
            + vbCr + " GROUP BY pla_marcacionhora.idemp, mae_tipohora_1.id, mae_tipohora_1.nomcor, mae_tipohora_1.variable, mae_tipohora_1.id " _
            + vbCr + " HAVING (((mae_tipohora_1.id)=15))) AS hora " _
            + vbCr + " ) AS tipohora " _
            + vbCr + " ON emp.id = tipohora.idemp " _
            + vbCr + " GROUP BY emp.id, emp.nombres, emp.cargo, emp.lugartra,emp.ingreso, emp.totdiatra, emp.totdiafer " _
            + vbCr + " PIVOT tipohora.idhora In (15,2,3,13,12,17,11);"
        
    End If
    
    RST_Busq RstTmp, nSQL, xCon
    
    '***************************************************************************************
    '****** CONFIGURAR LA GRILLA
    '***************************************************************************************
    Agregando = True
    Fg1.Cols = 1
    Fg1.Rows = 2
    Fg1.FixedRows = 2
    '--poner encabezado
    Dim mColRst As Integer, mColGrid As Integer, mRowGrid As Integer
    Dim mColTmp As Integer         '--almacena la ultima columna despues de datos,ingresos,descuentos,aportes
    Dim mColInicio As Integer      '--almacenar la posicion inicial donde empieza a colocar los acumulados
    Dim nTipoHora As String
    With Fg1
        .Rows = 3
        .FixedRows = 3
        .FrozenCols = 0
        .RowHeight(2) = 0
        .Cols = RstTmp.Fields.Count + 1
        For mColRst = 0 To RstTmp.Fields.Count - 1
            
            If LCase(RstTmp.Fields(mColRst).Name) <> "total" Then mColGrid = mColRst + 1
            
            Select Case LCase(RstTmp.Fields(mColRst).Name)
                Case "nombres"
                    .TextMatrix(1, mColGrid) = "Nombres": .ColWidth(mColGrid) = 2300: .ColAlignment(mColGrid) = flexAlignLeftCenter: .Row = 1: .Col = mColGrid: .CellAlignment = flexAlignLeftCenter
                Case "cargo"
                    .TextMatrix(1, mColGrid) = "Cargo": .ColWidth(mColGrid) = 1000:  .ColAlignment(mColGrid) = flexAlignLeftCenter:  .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignLeftCenter
                Case "lugartra"
                    .TextMatrix(1, mColGrid) = "Lugar Trabajo": .ColWidth(mColGrid) = 1300:  .ColAlignment(mColGrid) = flexAlignLeftCenter:  .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignLeftCenter
                Case "ingreso"
                    .TextMatrix(1, mColGrid) = "Ingreso": .ColWidth(mColGrid) = 850:  .ColAlignment(mColGrid) = flexAlignCenterCenter:  .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignCenterCenter
                Case "total"
                    .TextMatrix(1, .Cols - 1) = "Total" '--oculto
                    .ColWidth(.Cols - 1) = 0: .ColAlignment(.Cols - 1) = flexAlignCenterCenter: .Row = 1: .Col = .Cols - 1: .CellAlignment = flexAlignCenterCenter
                    mColInicio = mColGrid
                    mColInicioEdicion = mColGrid
                Case NulosN(RstTmp.Fields(mColRst).Name) <> 0
                    mColGrid = mColGrid - 1
                    nTipoHora = Busca_Codigo(NulosN(RstTmp.Fields(mColRst).Name), "id", "nomcor", "mae_tipohora", "N", xCon)
                    '*------------------------------
                    .TextMatrix(1, mColGrid) = Busca_Codigo(NulosN(RstTmp.Fields(mColRst).Name), "id", "nomcor", "mae_tipohora", "N", xCon)
                    .ColWidth(mColGrid) = 950: .ColAlignment(mColGrid) = flexAlignCenterCenter: .Row = 1: .Col = mColGrid: .CellAlignment = flexAlignCenterCenter
                    '--fila donde se colocaran los codigos de horas
                    .TextMatrix(2, mColGrid) = NulosN(RstTmp.Fields(mColRst).Name)
                Case "idemp"
                    .TextMatrix(1, mColGrid) = "IdEmp": .ColWidth(mColGrid) = 0: .ColAlignment(mColGrid) = flexAlignCenterCenter:  .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignCenterCenter
                Case "totdiatra"
                    .TextMatrix(1, mColGrid) = "Trab": .ColWidth(mColGrid) = 500:  .ColAlignment(mColGrid) = flexAlignRightCenter:  .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignRightCenter
                Case "totdiafer"
                    .TextMatrix(1, mColGrid) = "Fer": .ColWidth(mColGrid) = 500:  .ColAlignment(mColGrid) = flexAlignRightCenter:  .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignRightCenter
                Case Else
                    .TextMatrix(1, mColGrid) = RstTmp.Fields(mColRst).Name: .ColWidth(mColGrid) = 800: .ColAlignment(mColGrid) = flexAlignLeftCenter:  .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignLeftCenter
            End Select
            
        Next mColRst
        .FrozenCols = mColInicio - 2
        
        GRID_COMBINAR Fg1, 0, 1, 0, mColInicio - 1, "Datos del Personal", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC, True
        
        GRID_COMBINAR Fg1, 0, mColInicio - 1, 0, mColInicio, "Total Dias", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC, True
         
        GRID_COMBINAR Fg1, 0, mColInicio + 1, 0, .Cols - 1, "Total Horas", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC, True
    End With
    
    '***************************************************************************************
    '****** CARGAR LOS DATOS
    '***************************************************************************************
    Dim mDiaTotal&
    Fg1.Rows = Fg1.FixedRows
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        DoEvents
        Fg1.Rows = Fg1.Rows + 1
        mRowGrid = Fg1.Rows - 1
        mColGrid = 1
        For mColRst = 0 To RstTmp.Fields.Count - 1
            Select Case LCase(RstTmp.Fields(mColRst).Name)
                Case "nombres", "cargo"
                    Fg1.TextMatrix(mRowGrid, mColGrid) = NulosC(RstTmp.Fields(mColRst))
                Case "total"
                    Fg1.TextMatrix(mRowGrid, Fg1.Cols - 1) = ConvertHora(NulosN(RstTmp.Fields(mColRst)))
                Case "ingreso"
                    Fg1.TextMatrix(mRowGrid, mColGrid) = Format(NulosC(RstTmp.Fields(mColRst)), FORMAT_DATE)
                Case NulosN(RstTmp.Fields(mColRst).Name) <> 0
                    If NulosN(RstTmp.Fields(mColRst)) <> 0 Then
                        Fg1.TextMatrix(mRowGrid, mColGrid - 1) = ConvertHora(NulosN(RstTmp.Fields(mColRst)))
                    End If
                Case "totdiatra"
                    mDiaTotal = 0
                    If RstDiaTra.State = 1 Then
                        RstDiaTra.Filter = "idemp=" & NulosN(RstTmp.Fields("idemp"))
                        If RstDiaTra.RecordCount <> 0 Then mDiaTotal = NulosN(RstDiaTra.Fields("tottra"))
                    Else
                        mDiaTotal = NulosN(RstTmp.Fields("totdiatra"))
                    End If
                    Fg1.TextMatrix(mRowGrid, mColGrid) = mDiaTotal
                Case "totdiafer"
                    mDiaTotal = 0
                    If RstDiaFer.State = 1 Then
                        RstDiaFer.Filter = "idemp=" & NulosN(RstTmp.Fields("idemp"))
                        If RstDiaFer.RecordCount <> 0 Then mDiaTotal = NulosN(RstDiaFer.Fields("totfer"))
                    Else
                        mDiaTotal = NulosN(RstTmp.Fields("totdiafer"))
                    End If
                    Fg1.TextMatrix(mRowGrid, mColGrid) = mDiaTotal
                Case Else
                    Fg1.TextMatrix(mRowGrid, mColGrid) = NulosC(RstTmp.Fields(mColRst))
            End Select
            mColGrid = mColGrid + 1
        Next mColRst
        RstTmp.MoveNext
    Loop
    
    Agregando = False
    Set RstTmp = Nothing:    Set RstDiaTra = Nothing:    Set RstDiaFer = Nothing
    Exit Sub
error:
    Agregando = False
    Set RstTmp = Nothing:    Set RstDiaTra = Nothing:    Set RstDiaFer = Nothing
    SHOW_ERROR Me.Name, "pCargarDatosDet"
End Sub

'**************************

Private Sub pConfigurarGrilla()
    With Fg1 '--
        .Rows = 1
        .Cols = 8
        .FixedRows = 1
        .RowHeight(0) = 250
        .ColWidth(1) = 0:
        
        .TextMatrix(0, 2) = "CodSunat":      .ColWidth(2) = 800:  .ColAlignment(2) = flexAlignCenterCenter:  .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Categoría":     .ColWidth(3) = 1000:  .ColAlignment(3) = flexAlignLeftCenter:   .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 4) = "Tipo":          .ColWidth(4) = 2800: .ColAlignment(4) = flexAlignLeftCenter:    .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 5) = "Descripción":   .ColWidth(5) = 6200: .ColAlignment(5) = flexAlignLeftCenter:    .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 6) = "Variable":      .ColWidth(6) = 0: .ColAlignment(6) = flexAlignLeftCenter:    .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 7) = "Fórmula":       .ColWidth(7) = 0: .ColAlignment(7) = flexAlignLeftCenter:    .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftCenter
        
        .SelectionMode = flexSelectionByRow
    End With
    DoEvents
End Sub

Private Sub Fg1_EnterCell()
    If mColInicioEdicion = 0 Or Fg1.Col < mColInicioEdicion - 1 Or QueHace = 3 Then
        Fg1.Editable = flexEDNone
    Else
        Fg1.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then pRegistroAdd False
    If KeyCode = 46 Then pRegistroDel
End Sub

Private Sub CmdDet_Click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            pRegistroAdd False
        Case 1 '--eliminar
            pRegistroDel
        Case 2 '--seleccionar
            pRegistroAdd True
    End Select
End Sub

Private Sub pRegistroAdd(Optional fSeleccionVarios As Boolean = True)
    On Error GoTo error
    Dim xCampos(4, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String '--almacenara los codigos de documentos ya seleccionados
    Dim nSQLNotInDocumentos As String
    Dim nSQL As String
    Dim nTitulo As String
    xCampos(0, 0) = "CodSun":       xCampos(0, 1) = "codsun":       xCampos(0, 2) = "800":   xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    If fSeleccionVarios = True Then
        xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "5000":  xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
        xCampos(2, 0) = "Tipo":         xCampos(2, 1) = "tipnombre":    xCampos(2, 2) = "2800":  xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
        xCampos(3, 0) = "Categoría":    xCampos(3, 1) = "catnombre":    xCampos(3, 2) = "1700":  xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    Else
        xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "4200":  xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
        xCampos(2, 0) = "Tipo":         xCampos(2, 1) = "tipnombre":    xCampos(2, 2) = "2000":  xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
        xCampos(3, 0) = "Categoría":    xCampos(3, 1) = "catnombre":    xCampos(3, 2) = "1200":  xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    End If
    '*************************************************************
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 1, "pla_concepto.id", " NOT IN ")
    If nSQLId <> "" Then nSQLId = " WHERE " & nSQLId
    '*************************************************************
    nSQL = "SELECT pla_concepto.id, pla_conceptocat.descripcion AS catnombre, pla_conceptotipo.descripcion AS tipnombre, pla_concepto.codsun, pla_concepto.descripcion, pla_concepto.variable, pla_concepto.formula " _
        + vbCr + " FROM (pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
        + nSQLId _
        + vbCr + " ORDER BY pla_conceptocat.descripcion DESC,pla_concepto.id ; "

    nTitulo = "Buscando Conceptos"
    '*************************************************************
    If fSeleccionVarios = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), nTitulo
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "descripcion", "descripcion", CualquierParte
    End If
    If xRs.State = 0 Then GoTo salir
    If xRs.RecordCount = 0 Then GoTo salir
    If fSeleccionVarios = True Then xRs.MoveFirst
    Agregando = True
    Do While Not xRs.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("id"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("codsun"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRs("catnombre"))
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRs("tipnombre"))
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(xRs("descripcion"))
        '-----------
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(xRs("variable"))
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(xRs("formula"))
        '-----------
        If fSeleccionVarios = False Then Exit Do
        xRs.MoveNext
    Loop
    Agregando = False
    Fg1.Row = Fg1.Rows - 1: Fg1.Col = 6:  Fg1.SetFocus
salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "Fg1_CellButtonClick"
End Sub

Private Sub pRegistroDel()
    If Fg1.Rows = 1 Then Exit Sub
    If Fg1.Row < 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    Fg1.RemoveItem Fg1.Row
End Sub

'****************************************************************************************
Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error

    If QueHace = 3 Then Exit Sub

    Select Case Index
        Case 0 '--proceso
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"

            nTitulo = "Seleccionar Proceso"
            nSQL = "SELECT pla_proceso.id, pla_proceso.descripcion AS nombre, pla_proceso.id AS cod " _
                + vbCr + " FROM pla_proceso " _
                + vbCr + " WHERE (((pla_proceso.enproceso)=-1)); "
        Case 1 '--lugar de trabajo
            ReDim xCampos(4, 3) As String
            xCampos(0, 0) = "Establecimiento":  xCampos(0, 1) = "centro":  xCampos(0, 2) = "1500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Nombre":           xCampos(1, 1) = "nombre":           xCampos(1, 2) = "2500":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Dirección":        xCampos(2, 1) = "direc":            xCampos(2, 2) = "3000":   xCampos(2, 3) = "C"
            xCampos(3, 0) = "Id":               xCampos(3, 1) = "id":               xCampos(3, 2) = "400":    xCampos(3, 3) = "N"
            nTitulo = "Buscando Lugar de Trabajo"
            
            nSQL = "SELECT mae_empresalugtra.id, mae_empresalugtra.descripcion AS nombre, mae_empresalugtra.id AS cod, mae_empresalugtra.direc, mae_tipoestablecimiento.descripcion AS centro " _
                + vbCr + " FROM mae_tipoestablecimiento RIGHT JOIN mae_empresalugtra ON mae_tipoestablecimiento.id = mae_empresalugtra.idtipest;"
                

    End Select


    Dim xRs As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    lbl_cb(Index).ToolTipText = xRs.Fields(1) & "" '--NOMBRE
       
    If Trim(lbl_cod(Index).Tag) <> Trim(lbl_cod(Index).Caption) Then
        Select Case Index
            Case 0 '--manual tipo doc
                Fg1.Rows = Fg1.FixedRows
        End Select
    End If
    Select Case Index
        Case 0 '--proceso
            txtfecha(0).SetFocus
    End Select

salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
        Me.lbl_cb(Index).Tag = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub

    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--proceso
            nSQL = "SELECT pla_proceso.id, pla_proceso.descripcion AS nombre, pla_proceso.id AS cod " _
                + vbCr + " FROM pla_proceso " _
                + vbCr + " WHERE (((pla_proceso.enproceso)=-1)) and pla_proceso.id  = " & NulosN(txt_cb(Index).Text) & ""
        Case 1 '--lugar de trabajo
            nSQL = "SELECT mae_empresalugtra.id, mae_empresalugtra.descripcion AS nombre, mae_empresalugtra.id AS cod, mae_empresalugtra.direc, mae_tipoestablecimiento.descripcion AS centro " _
                + vbCr + " FROM mae_tipoestablecimiento RIGHT JOIN mae_empresalugtra ON mae_tipoestablecimiento.id = mae_empresalugtra.idtipest " _
                + vbCr + " WHERE mae_tipoestablecimiento.id = " & NulosN(txt_cb(Index).Text) & ";"
    End Select

    If xCon.State = 0 Then GoTo salir
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.State = 0 Then GoTo salir
    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
        lbl_cb(Index).ToolTipText = RstTmp.Fields(1) & "" '--NOMBRE
        '--------------
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    '--------------
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
salir:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub

'****************************************************************************************

Private Sub cmdCargarPersonal_Click()
    If IsDate(txtfecha(0).Valor) = False Or IsDate(txtfecha(1).Valor) = False Then
        MsgBox "Ingrese la Fecha " & IIf(IsDate(txtfecha(0).Valor) = False, "de Inicio", "Final") & " ", vbExclamation, xTitulo
       If IsDate(txtfecha(0).Valor) = False Then txtfecha(0).SetFocus Else txtfecha(1).SetFocus
        Exit Sub
    End If
    
    If CDate(txtfecha(0).Valor) > CDate(txtfecha(1).Valor) Then
        MsgBox "La fecha Inicial es Superior al Final" + vbCr + "Modifique la fecha", vbExclamation, xTitulo
        txtfecha(0).SetFocus
        Exit Sub
    End If
    If NulosN(lbl_cod(0).Caption) = 0 Then
        MsgBox "Seleccione el Tipo de Pago", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Sub
    End If
    
    pCargarDatosDet CDate(txtfecha(0).Valor), CDate(txtfecha(1).Valor), eEncabezado
End Sub
