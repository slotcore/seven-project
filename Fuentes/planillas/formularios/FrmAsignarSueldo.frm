VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmAsignarSueldo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Asignar Sueldo a Personal"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11790
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
            Picture         =   "FrmAsignarSueldo.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":1EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":23EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSueldo.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   609
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
      Top             =   390
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
         Left            =   -12390
         TabIndex        =   4
         Top             =   375
         Width           =   11745
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6405
            Left            =   45
            TabIndex        =   5
            Top             =   390
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   11298
            _LayoutType     =   4
            _RowHeight      =   14
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Personal"
            Columns(0).DataField=   "nombres"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Proceso"
            Columns(1).DataField=   "docnombre"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Categoría"
            Columns(2).DataField=   "catnombre"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Concepto"
            Columns(3).DataField=   "descripcion"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Importe"
            Columns(4).DataField=   "imptot"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=5239"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5159"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=256"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2249"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2170"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=256"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2170"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2090"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=256"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=7461"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=7382"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=256"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1826"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1746"
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
            _StyleDefs(54)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(55)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14,.alignment=0"
            _StyleDefs(56)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(59)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=0"
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
         Begin VB.Label lblperiodo 
            Alignment       =   1  'Right Justify
            Caption         =   "lblperiodo(0)"
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
            Height          =   300
            Index           =   0
            Left            =   9570
            TabIndex        =   11
            Top             =   75
            Width           =   1980
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Asignar Sueldo"
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
         Begin VB.Frame Frame4 
            Caption         =   "( Periodo )"
            Height          =   720
            Left            =   9525
            TabIndex        =   22
            Top             =   210
            Width           =   2010
            Begin VB.Label lblperiodo 
               Alignment       =   2  'Center
               Caption         =   "lblperiodo(1)"
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
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   23
               Top             =   330
               Width           =   1740
            End
         End
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Caption         =   "Frame12"
            Height          =   585
            Index           =   0
            Left            =   75
            TabIndex        =   17
            Top             =   6165
            Width           =   11460
            Begin VB.TextBox txttotal 
               Alignment       =   1  'Right Justify
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
               Height          =   315
               Left            =   9585
               Locked          =   -1  'True
               TabIndex        =   20
               Text            =   "txttotal"
               Top             =   105
               Width           =   1290
            End
            Begin VB.CommandButton CmdDet 
               Caption         =   "&Eliminar"
               Enabled         =   0   'False
               Height          =   435
               Index           =   1
               Left            =   1890
               TabIndex        =   19
               Top             =   60
               Width           =   1395
            End
            Begin VB.CommandButton CmdDet 
               Caption         =   "&Agregar"
               Enabled         =   0   'False
               Height          =   435
               Index           =   0
               Left            =   120
               TabIndex        =   18
               Top             =   60
               Width           =   1395
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Total"
               Height          =   240
               Left            =   8655
               TabIndex        =   21
               Top             =   195
               Width           =   765
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
         Begin VB.CommandButton CmdDefault 
            Caption         =   "Predeterminado"
            Enabled         =   0   'False
            Height          =   405
            Left            =   7875
            TabIndex        =   10
            Top             =   525
            Width           =   1575
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
         Begin VSFlex7Ctl.VSFlexGrid fg1 
            Height          =   5115
            Left            =   120
            TabIndex        =   9
            Top             =   1005
            Width           =   11415
            _cx             =   20135
            _cy             =   9022
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmAsignarSueldo.frx":2B10
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
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   0
            Left            =   1485
            Picture         =   "FrmAsignarSueldo.frx":2C03
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   660
            Width           =   210
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   975
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   13
            Text            =   "txt_cb(0)"
            Top             =   630
            Width           =   765
         End
         Begin VB.Label lbl_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Personal"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   15
            Top             =   735
            Width           =   615
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
            Left            =   5865
            TabIndex        =   14
            Top             =   645
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
            Caption         =   "Detalle de Asignar Sueldo"
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
            Left            =   90
            TabIndex        =   3
            Top             =   60
            Width           =   11400
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
            Left            =   1755
            TabIndex        =   16
            Top             =   645
            Width           =   4725
         End
      End
   End
End
Attribute VB_Name = "FrmAsignarSueldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstFrm As New ADODB.Recordset
Dim RstMarca As New ADODB.Recordset
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
    
    fg1.SelectionMode = flexSelectionByRow
    Me.MousePointer = vbDefault
    TabOne1.CurrTab = 0
End Sub

Private Sub CmdDefault_Click()
'    If IsDate(txtfecha.Valor) = False Then
'        MsgBox "La fecha no es correcta", vbExclamation, xTitulo
'        Exit Sub
'    End If
'    CargarDefault = True
'    pCargarDatosDet
    
End Sub

Private Sub CmdDet_Click(Index As Integer)
    Select Case Index
        Case 0 '--agregar
            pRegistroAdd
        Case 1 '--eliminar
            pRegistroDel
    End Select
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
    
    If QueHace = 3 Then Exit Sub
    
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String '--almacenará los codigos de documentos ya seleccionados
    Dim nSQL As String
    Dim nTitulo As String
        
    Select Case Col
        Case 4 '--proceso
            ReDim xCampos(2, 5) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "5500": xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":           xCampos(1, 2) = "800":  xCampos(1, 3) = "C":    xCampos(1, 4) = "N"

            nSQL = "SELECT pla_proceso.id, pla_proceso.descripcion FROM pla_proceso WHERE enproceso=-1; "
            nTitulo = "Buscando Proceso"
            
        Case 6 '--concepto
            If NulosN(fg1.TextMatrix(Row, 2)) = 0 Then
                MsgBox "Falta seleccionar la Categoría", vbExclamation, xTitulo
                Exit Sub
            End If
            ReDim xCampos(3, 5)
            xCampos(0, 0) = "CodSun":       xCampos(0, 1) = "codsun":      xCampos(0, 2) = "900":  xCampos(0, 3) = "C": xCampos(0, 4) = "S"
            xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion": xCampos(1, 2) = "5000": xCampos(1, 3) = "C": xCampos(1, 4) = "N"
            xCampos(2, 0) = "Tipo":         xCampos(2, 1) = "tipnombre":   xCampos(2, 2) = "2200": xCampos(2, 3) = "C": xCampos(2, 4) = "N"
    
            nSQLId = GRID_GENERAR_SQL_ID(fg1, 3, "pla_concepto.id", "NOT IN", True, 1, fg1.TextMatrix(Row, 1))
            If nSQLId <> "" Then nSQLId = " and " & nSQLId
    
            nSQL = "SELECT pla_concepto.id,pla_concepto.codsun, pla_concepto.descripcion, pla_conceptotipo.descripcion AS tipnombre " _
                + vbCr + " FROM pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
                + vbCr + " WHERE pla_concepto.variable is not null AND (((pla_conceptotipo.idcat)=" & NulosN(fg1.TextMatrix(Row, 2)) & ")) " & nSQLId & " and pla_concepto.activo=-1;"

            nTitulo = "Buscando Conceptos"
            
      End Select
    
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
    If xRs.State = 0 Then GoTo salir
    If xRs.RecordCount = 0 Then GoTo salir
    
    Agregando = True
    
    fg1.TextMatrix(fg1.Row, Col) = NulosC(xRs.Fields("descripcion"))
    fg1.TextMatrix(fg1.Row, Col - 3) = NulosN(xRs.Fields("id"))
    
salir:
    Set xRs = Nothing
    Agregando = False
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row <= fg1.FixedRows - 1 Then Exit Sub
    
    Select Case Col
        Case 5 '--categoria
            If NulosN(fg1.TextMatrix(Row, 2)) <> NulosN(fg1.Cell(flexcpText, Row, Col)) Then
                fg1.TextMatrix(Row, 2) = NulosN(fg1.Cell(flexcpText, Row, Col))
                
                fg1.TextMatrix(Row, 3) = ""
                fg1.TextMatrix(Row, 6) = ""
            End If
        Case 7 '--importe
            If IsNumeric(fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "Ingrese un Valor Numérico", vbInformation, xTitulo
                fg1.TextMatrix(Row, Col) = ""
            End If
    End Select
    '--totalizar
    pTotalizar
    Exit Sub
error:
    SHOW_ERROR Me.Name, "fg1_CellChanged"
End Sub

Private Sub Fg1_EnterCell()
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then
        fg1.Editable = flexEDNone
        Exit Sub
    End If
    
    Select Case fg1.Col
        Case Is > 3
            fg1.Editable = flexEDKbdMouse
            
        Case Else
            fg1.Editable = flexEDNone
    End Select
    
End Sub

Private Sub Fg1_KeyPress(KeyAscii As Integer)
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col <> 7 Then KeyAscii = 0
    If validar_numero(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then
        pRegistroAdd
    ElseIf KeyCode = 46 Then
        pRegistroDel
    End If
End Sub

Private Sub Fg1_RowColChange()
    If Agregando = True Then Exit Sub
    If fg1.Rows <= 1 Then Exit Sub

End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = False
    pConfigurarGrilla
    pCargarGrid
    fg1.SelectionMode = flexSelectionByRow
    SeEjecuto = True
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha Asignado Conceptos al Personal, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            nuevo
        End If
    End If
End Sub


Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    QueHace = 3
    
    Dg1.Columns("imptot").NumberFormat = FORMAT_MONTO
    
    TabOne1.CurrTab = 0
End Sub

Sub Blanquea()
    LimpiaText txt
    LimpiaText txt_cb
    fg1.Rows = fg1.FixedRows
    txttotal.Text = "0.00"
End Sub

Sub Bloquea(band As Boolean)
    habilitar_Locked txt_cb, Not band
    CmdDefault.Enabled = band
    habilitar CmdDet, band
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

    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " el Registro ", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstDet As New ADODB.Recordset
    Dim xId&
    On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    '--eliminar datos detalle de la asignacion de sueldos
    xCon.Execute "DELETE FROM pla_conceptoemp WHERE (((pla_conceptoemp.anno)=" & AnoTra & ") AND ((pla_conceptoemp.idmes)= " & xMes & " ) AND ((pla_conceptoemp.idemp)=" & NulosN(txt_cb(0).Text) & "));"

    RST_Busq RstDet, "SELECT TOP 1 * FROM pla_conceptoemp ; ", xCon

    '******************************************************
    Dim mRow&
    '--grabando el detalle
    For mRow = fg1.FixedRows To fg1.Rows - 1
        If NulosN(fg1.TextMatrix(mRow, 1)) <> 0 And NulosN(fg1.TextMatrix(mRow, 2)) <> 0 And NulosN(fg1.TextMatrix(mRow, 3)) <> 0 Then
            RstDet.AddNew
            RstDet("anno") = AnoTra
            RstDet("idmes") = xMes
            RstDet("idemp") = NulosN(txt_cb(0).Text)
            RstDet("idproc") = NulosN(fg1.TextMatrix(mRow, 1))
            RstDet("idcpto") = NulosN(fg1.TextMatrix(mRow, 3))
            RstDet("imptot") = NulosN(fg1.TextMatrix(mRow, 7))
            RstDet.Update
        End If
    Next mRow
    
    '******************************************************
    Me.MousePointer = vbDefault
    xCon.CommitTrans
    MsgBox "Los datos del registro " + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

    Set RstDet = Nothing
    Grabar = True
    Exit Function
LaCague:
    Me.MousePointer = vbDefault
    Set RstDet = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Sub nuevo()
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea True
    Blanquea
    pConfigurarGrilla
    
    fg1.SelectionMode = flexSelectionFree
    
    Label5.Caption = "Agregando Asignación de Sueldo"
    txt_cb(0).SetFocus
End Sub

Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Label5.Caption = "Modificando Asignación de Sueldo"

    ActivaTool
    
    Bloquea True
    
    If TabOne1.CurrTab = 0 Then TabOne1.CurrTab = 1
        
    TabOne1.TabEnabled(0) = False
    
    fg1.SelectionMode = flexSelectionFree
        
    QueHace = 2
        
    Agregando = False
    
    CargarDefault = True
    
    txt_cb(0).SetFocus

End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Eliminar()
    TabOne1.CurrTab = 0
    Dim Rpta As Integer

    Rpta = MsgBox("Esta seguro de eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Dim RstTmp As New ADODB.Recordset
        Dim nSQL As String
        Dim xCod&
        
        xCod = RstFrm.Fields("id")
        xCon.Execute "DELETE FROM pla_conceptoempdet WHERE (((pla_conceptoempdet.idmarca)=" & xCod & ") AND ((pla_conceptoempdet.idori) In (1,5,6,7)));"
        xCon.Execute "DELETE FROM pla_conceptoemphora WHERE (((pla_conceptoemphora.idmarca)=" & xCod & ") AND ((pla_conceptoemphora.idhora) In (1,2,3,9,10,11,12,13)));"
        nSQL = "SELECT pla_conceptoempdet.idmarca, pla_conceptoempdet.idemp FROM pla_conceptoempdet WHERE (((pla_conceptoempdet.idmarca)=" & xCod & ")) " _
            + vbCr + " Union " _
            + vbCr + " SELECT pla_conceptoemphora.idmarca,pla_conceptoemphora.idemp FROM pla_conceptoemphora WHERE (((pla_conceptoemphora.idmarca)=" & xCod & ")) "
        RST_Busq RstTmp, nSQL, xCon
        If RstTmp.RecordCount = 0 Then
            xCon.Execute "DELETE FROM pla_conceptoemp WHERE pla_conceptoemp.id=" & xCod
        End If
        Set RstTmp = Nothing
        MsgBox "El Registro se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        pCargarGrid
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
    If Button.Index = 11 Then CambiarMes
    
    If Button.Index = 13 Or Button.Index = 14 Then
    
        If TabOne1.CurrTab = 0 Then
            MsgBox "Primero muestre el detalle", vbInformation, xTitulo
            Exit Sub
        End If
        
        Dim mCol&
        Agregando = True
        fg1.Rows = fg1.Rows + 1
        FORMATO_CELDA fg1, fg1.Rows - 1, 6, , True, , "Total"
        FORMATO_CELDA fg1, fg1.Rows - 1, 7, , True, , txttotal.Text
        mCol = fg1.ColWidth(7)
        fg1.ColWidth(7) = 1000
        
        If Button.Index = 13 Then pExportarExcel
        If Button.Index = 14 Then pImprimir
        
        '--restaurando valores
        fg1.Rows = fg1.Rows - 1
        fg1.ColWidth(7) = mCol
        '------------
        Agregando = False
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
    
    xCampos(0, 0) = "Apellidos y Nombres": xCampos(0, 1) = "nombres":     xCampos(0, 2) = "2200":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Tipo Doc.":           xCampos(1, 1) = "docnombre":   xCampos(1, 2) = "1000":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Categoría":           xCampos(2, 1) = "catnombre":   xCampos(2, 2) = "900":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Concepto":            xCampos(3, 1) = "descripcion": xCampos(3, 2) = "3200":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "Importe":             xCampos(4, 1) = "imptot":      xCampos(4, 2) = "900":    xCampos(4, 3) = "N"
    
    nSQL = "SELECT pla_conceptoemp.*, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pla_proceso.abrev, pla_proceso.descripcion AS docnombre, pla_conceptocat.id AS catid, pla_conceptocat.descripcion AS catnombre, pla_concepto.descripcion " _
        + vbCr + " FROM (pla_conceptocat RIGHT JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) RIGHT JOIN (pla_empleados RIGHT JOIN (pla_concepto RIGHT JOIN (pla_conceptoemp LEFT JOIN pla_proceso ON pla_conceptoemp.idproc = pla_proceso.id) ON pla_concepto.id = pla_conceptoemp.idcpto) ON pla_empleados.id = pla_conceptoemp.idemp) ON pla_conceptotipo.id = pla_concepto.idtipo " _
        + vbCr + " WHERE (((pla_conceptoemp.idmes) = " & xMes & ") And ((pla_conceptoemp.anno) = " & AnoTra & ")) " _
        + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat], pla_proceso.descripcion, pla_conceptocat.descripcion DESC , pla_concepto.descripcion; "

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Asignación de Sueldo", "nombres", "nombres", Principio

    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir
    
    RstFrm.MoveFirst
    RstFrm.Find "idemp = " & CStr(xRs("idemp")) '& " and idproc = " & CStr(xRs("idproc")) & " and idcpto = " & CStr(xRs("idcpto"))
salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

Sub Filtrar()
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(5, 3) As String

    xCampos(0, 0) = "Apellidos y Nombres": xCampos(0, 1) = "nombres":     xCampos(0, 2) = "C":   xCampos(0, 3) = "2200"
    xCampos(1, 0) = "Tipo Doc.":           xCampos(1, 1) = "docnombre":   xCampos(1, 2) = "C":   xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Categoría":           xCampos(2, 1) = "catnombre":   xCampos(2, 2) = "C":    xCampos(2, 3) = "900"
    xCampos(3, 0) = "Concepto":            xCampos(3, 1) = "descripcion": xCampos(3, 2) = "C":   xCampos(3, 3) = "3200"
    xCampos(4, 0) = "Importe":             xCampos(4, 1) = "imptot":      xCampos(4, 2) = "N":    xCampos(4, 3) = "900"
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1

End Sub

Sub MuestraSegundoTab()
    Blanquea
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros para Mostrar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    QueHace = -1 '--comodin para entrar a [txt_cb_Validate]
    
    txt_cb(0).Text = RstFrm.Fields("idemp")
    lbl_cb(0).Caption = RstFrm.Fields("nombres")
    lbl_cod(0).Caption = RstFrm.Fields("idemp")
    
    CargarDefault = False
    pCargarDatosDet
    QueHace = 3
End Sub

Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL As String
    lblperiodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
    lblperiodo(1).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
    Dim RstTmp As New ADODB.Recordset
       
    nSQL = "SELECT pla_conceptoemp.*, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pla_proceso.abrev, pla_proceso.descripcion AS docnombre, pla_conceptocat.id AS catid, pla_conceptocat.descripcion AS catnombre, pla_concepto.descripcion " _
        + vbCr + " FROM (pla_conceptocat RIGHT JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) RIGHT JOIN (pla_empleados RIGHT JOIN (pla_concepto RIGHT JOIN (pla_conceptoemp LEFT JOIN pla_proceso ON pla_conceptoemp.idproc = pla_proceso.id) ON pla_concepto.id = pla_conceptoemp.idcpto) ON pla_empleados.id = pla_conceptoemp.idemp) ON pla_conceptotipo.id = pla_concepto.idtipo " _
        + vbCr + " WHERE (((pla_conceptoemp.idmes) = " & xMes & ") And ((pla_conceptoemp.anno) = " & AnoTra & ")) " _
        + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat], pla_proceso.descripcion, pla_conceptocat.descripcion DESC , pla_concepto.descripcion; "

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
    band = Validar(txt_cb)
    
    If NulosN(txt_cb(0).Text) = 0 Then
        MsgBox "Falta seleccionar un Personal", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    '
    If fg1.Rows = 1 Then
        MsgBox "Falta Asignar algún Sueldo", vbExclamation, xTitulo
        CmdDet(0).SetFocus
        Exit Function
    End If
    Dim mRow&, mCol&
    mCol = -1
    For mRow = fg1.FixedRows To fg1.Rows - 1
        If NulosN(fg1.TextMatrix(mRow, 1)) = 0 Then   '--tipo doc
            MsgBox "Falta especificar el tipo de documento", vbExclamation, xTitulo
            mCol = 4:          Exit For
        ElseIf NulosN(fg1.TextMatrix(mRow, 2)) = 0 Then '--categoria
            MsgBox "Falta especificar la Categoría del Concepto", vbExclamation, xTitulo
            mCol = 5:          Exit For
        ElseIf NulosN(fg1.TextMatrix(mRow, 3)) = 0 Then '--concepto
            MsgBox "Falta especificar el Concepto..." & vbCr & "Categoría: " & fg1.TextMatrix(mRow, 4), vbExclamation, xTitulo
            mCol = 6:          Exit For
        End If
    Next mRow
    If mCol <> -1 Then
        Agregando = True:  fg1.Row = mRow: fg1.Col = mCol: Agregando = False
        fg1.SetFocus
        Exit Function
    End If
    '--
    fValidarDatos = True
    
End Function


Private Sub pConfigurarGrilla()
    With fg1 '--del personal
        .Rows = 2
        .Cols = 8
        .FixedRows = 2
        .RowHeight(0) = 300
        .ColWidth(1) = 0:
        
        UNIR_CELDAS fg1, 0, 4, 1, 4, " ", flexAlignLeftCenter, True
        UNIR_CELDAS fg1, 0, 5, 0, 7, "Datos de los Conceptos", flexAlignCenterCenter
                
        .TextMatrix(1, 1) = "idproc":             .ColWidth(1) = 0:
        .TextMatrix(1, 2) = "IdCat":             .ColWidth(2) = 0:
        .TextMatrix(1, 3) = "IdCpto":            .ColWidth(3) = 0:
        
        .TextMatrix(1, 4) = "Proceso":           .ColWidth(4) = 2200:  .ColAlignment(4) = flexAlignLeftCenter:  .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 5) = "Categoría":         .ColWidth(5) = 1500:  .ColAlignment(5) = flexAlignLeftCenter:  .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 6) = "Concepto":          .ColWidth(6) = 5700:  .ColAlignment(6) = flexAlignLeftCenter:  .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 7) = "Importe":           .ColWidth(7) = 1300:  .ColAlignment(7) = flexAlignRightCenter: .Row = 1: .Col = 7: .CellAlignment = flexAlignRightCenter
        
        .ColFormat(7) = FORMAT_MONTO
        
        .SelectionMode = flexSelectionByRow
    End With
    
    GRID_COMBOLIST fg1, 4 '--tipo doc
    GRID_COMBOLIST fg1, 6 '--concepto
    
    '*****************************************
    '--categoria
    '--COMBOLIST CON VSFLEXGRID
    Dim RstTmp As New ADODB.Recordset
    Dim tFormat$
    RST_Busq RstTmp, "SELECT pla_conceptocat.id, pla_conceptocat.descripcion FROM pla_conceptocat WHERE pla_conceptocat.id IN (1,3) ;", xCon
    tFormat = fg1.BuildComboList(RstTmp, "descripcion", "id", vbYellow)
    fg1.ColComboList(5) = tFormat
    Set RstTmp = Nothing
    DoEvents
End Sub

Private Sub pRegistroAdd()
    Dim mCol%
    Dim fInsertar As Boolean
    If QueHace = 3 Then Exit Sub
    Agregando = True
    If fg1.Rows > fg1.FixedRows Then
        If NulosN(fg1.TextMatrix(fg1.Rows - 1, 1)) = 0 Then
            MsgBox "Seleccione el Tipo de Documento ", vbExclamation, xTitulo
        ElseIf NulosN(fg1.TextMatrix(fg1.Rows - 1, 2)) = 0 Then
            MsgBox "Seleccione la Categoría", vbExclamation, xTitulo
        ElseIf NulosN(fg1.TextMatrix(fg1.Rows - 1, 3)) = 0 Then
            MsgBox "Seleccione el Concepto", vbExclamation, xTitulo
        ElseIf NulosN(fg1.TextMatrix(fg1.Rows - 1, 7)) = 0 Then
            MsgBox "Ingrese el valor del Importe..." + vbCr + "Concepto: " & fg1.TextMatrix(fg1.Rows - 1, 6), vbExclamation, xTitulo
        Else
            fInsertar = True
        End If
    Else
        fInsertar = True
    End If
    
    If fg1.Rows > 2 And fInsertar = True Then
        '--ordenar segun tipo de
        GRID_ORDENAR fg1, 2, 4, 2, 5, flexSortStringAscending
        '--agrupar por tipo de documento
        GRID_AGRUPAR fg1, 1
    End If
    
    If fInsertar = True Then fg1.AddItem ""
    
    GRID_COLOR_FONDO fg1, fg1.Rows - 1, 1, fg1.Rows - 1, fg1.Cols - 1, &HC0C0FF
    
    fg1.Row = fg1.Rows - 1
    fg1.Col = 4
    
    fg1.SetFocus
    Agregando = False
End Sub

Private Sub pRegistroDel()
    Dim mRowDel&, mRow&
    If fg1.Rows = 1 Then Exit Sub
    If fg1.Row < 1 Then Exit Sub

    If MsgBox("Seguro desea eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    mRow = fg1.Row
    fg1.RemoveItem fg1.Row
    '--agrupar -----------------------
    Me.MousePointer = vbHourglass
    Agregando = True
    If fg1.Rows > 2 Then
        '--ordenar segun tipo de
        GRID_ORDENAR fg1, 2, 4, 2, 5, flexSortStringAscending
        '--agrupar por tipo de documento
        GRID_AGRUPAR fg1, 1
    Else
        Me.MousePointer = vbDefault
        CmdDet(0).SetFocus
        Exit Sub
    End If
    GRID_COLOR_FONDO fg1, fg1.Rows - 1, 1, fg1.Rows - 1, fg1.Cols - 1, &HC0C0FF
    Agregando = False
    
    If fg1.Rows > fg1.FixedRows Then
        fg1.Row = fg1.Rows - 1
    ElseIf fg1.Rows = fg1.FixedRows Then
        fg1.Row = fg1.FixedRows - 1
    End If
    '------------
    pTotalizar
    '------------
    fg1.Col = 4
    Me.MousePointer = vbDefault
    '-------------------------------
End Sub

Private Sub CambiarMes()
    xMes = SeleccionaMes(xCon)
    pCargarGrid
    TabOne1.CurrTab = 0
End Sub


Private Sub pCargarDatosDet()
    
    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    
    nSQL = "SELECT pla_conceptoemp.*, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS nombres, pla_proceso.descripcion AS docnombre, pla_conceptotipo.idcat, pla_conceptocat.descripcion AS catnombre, pla_concepto.codsun, pla_concepto.descripcion AS cptonombre " _
        + vbCr + " FROM (pla_conceptocat INNER JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) INNER JOIN (pla_empleados INNER JOIN (pla_concepto INNER JOIN (pla_proceso INNER JOIN pla_conceptoemp ON pla_proceso.id = pla_conceptoemp.idproc) ON pla_concepto.id = pla_conceptoemp.idcpto) ON pla_empleados.id = pla_conceptoemp.idemp) ON pla_conceptotipo.id = pla_concepto.idtipo " _
        + vbCr + " WHERE (((pla_conceptoemp.idmes)=" & xMes & ") AND ((pla_conceptoemp.anno)=" & AnoTra & ") AND ((pla_conceptoemp.idemp)=" & RstFrm.Fields("idemp") & " )) " _
        + vbCr + " ORDER BY pla_proceso.descripcion, pla_conceptocat.descripcion DESC , pla_concepto.descripcion;"

    RST_Busq RstTmp, nSQL, xCon
    
    '**********************************************************************************************************
    Agregando = True
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        fg1.Rows = fg1.Rows + 1
        '------
        fg1.TextMatrix(fg1.Rows - 1, 1) = NulosN(RstTmp("idproc"))
        fg1.TextMatrix(fg1.Rows - 1, 2) = NulosC(RstTmp("idcat"))
        fg1.TextMatrix(fg1.Rows - 1, 3) = NulosC(RstTmp("idcpto"))
        fg1.TextMatrix(fg1.Rows - 1, 4) = NulosC(RstTmp("docnombre"))
        fg1.TextMatrix(fg1.Rows - 1, 5) = NulosC(RstTmp("catnombre"))
        fg1.TextMatrix(fg1.Rows - 1, 6) = NulosC(RstTmp("cptonombre"))
        fg1.TextMatrix(fg1.Rows - 1, 7) = Format(NulosN(RstTmp("imptot")), FORMAT_MONTO)
        RstTmp.MoveNext
    Loop
    '--poner los colores en el grid
    '--agrupar -----------------------
    GRID_AGRUPAR fg1, 1
    Me.MousePointer = vbDefault
    '-------------------------------
    If fg1.Rows > fg1.FixedRows Then
        fg1.Row = fg1.FixedRows
        fg1.Col = 4
    End If
    Agregando = False
   
    pTotalizar
   
    Me.MousePointer = vbDefault
End Sub

Private Sub pTotalizar()
    '---total
    txttotal.Text = Format(GRID_SUMAR_COL(fg1, 7), FORMAT_MONTO)
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
        Case 0 '--pesonal
            pCargarPersonal
            Exit Sub
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
        Case 0 '--personal
           
            nSQL = "SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.id as cod,pla_empleados.numdoc, mae_dociden.abrev AS tipodoc, mae_sexo.abrev AS sexo, Format([pla_empleados].[fchnac],'dd/mm/yyyy') AS fchnac, pla_empleados.numtel, pla_empleados.email " _
                + vbCr + " FROM (mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex) LEFT JOIN pla_periodolaboral ON pla_empleados.id = pla_periodolaboral.idemp " _
                + vbCr + " WHERE pla_empleados.id  = " & NulosN(txt_cb(Index).Text) & "" _
                + vbCr + " and (pla_periodolaboral.fchfin Is Null) AND pla_empleados.id not in ( SELECT pla_conceptoemp.idemp From pla_conceptoemp WHERE (((pla_conceptoemp.anno)=" & AnoTra & " ) AND ((pla_conceptoemp.idmes)= " & xMes & " and pla_conceptoemp.idemp <> " & NulosN(txt_cb(0).Text) & "))  );"

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

Private Sub pImprimir()
    On Error GoTo error
    
    Me.MousePointer = vbHourglass
    Dim oPrint As New SGI2_funciones.formularios
    oPrint.Imprimir_x_VSFlexGrid fg1, "Asignación de Sueldos  - " & lblperiodo(0).Caption & " " & AnoTra, , "Personal: " & StrConv(lbl_cb(0).Caption, 3), False, True
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
    oExport.VSFlexGrid_Exportar_MSExcel xCon, fg1, "Asignación de Sueldos - " & lblperiodo(0).Caption & " " & AnoTra, "Personal: " & StrConv(lbl_cb(0).Caption, 3), , "Asignación de Sueldos"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub


'--------------
Private Sub pCargarPersonal()
    Dim xRs As New ADODB.Recordset
    pBuscarPersonal xRs, True
    If xRs.State = 1 Then
        txt_cb(0) = xRs.Fields("id") & "" '--TEXTO A MOSTRAR
        lbl_cb(0).Caption = xRs.Fields("nombres") & "" '--NOMBRE
        lbl_cod(0).Caption = xRs.Fields("id") & "" '--CODIGO
        lbl_cb(0).ToolTipText = xRs.Fields("nombres") & "" '--NOMBRE
        txt_cb(0).SetFocus
    End If
    Set xRs = Nothing
End Sub
'--------------

