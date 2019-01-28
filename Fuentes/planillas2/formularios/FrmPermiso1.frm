VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPermiso1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Registro de Permiso"
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
            Picture         =   "FrmPermiso1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":1EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":23EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPermiso1.frx":277E
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
      Width           =   11790
      _ExtentX        =   20796
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
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   11298
            _LayoutType     =   4
            _RowHeight      =   14
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Código"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fch Emi"
            Columns(1).DataField=   "fchemi"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Personal"
            Columns(2).DataField=   "nombres"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch Ini"
            Columns(3).DataField=   "fchini"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch Fin"
            Columns(4).DataField=   "fchfin"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Hor Ini"
            Columns(5).DataField=   "horini"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Hor Fin"
            Columns(6).DataField=   "horfin"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Motivo"
            Columns(7).DataField=   "permiso"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Glosa"
            Columns(8).DataField=   "observacion"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=770"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2170"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2090"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2170"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2090"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=2196"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2117"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(24)=   "Column(4).Width=2170"
            Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2090"
            Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(30)=   "Column(5).Width=1905"
            Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=1826"
            Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(36)=   "Column(6).Width=2355"
            Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2275"
            Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(42)=   "Column(7).Width=2328"
            Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2249"
            Splits(0)._ColumnProps(45)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(47)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(50)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
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
            _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1,.bold=0,.fontsize=825"
            _StyleDefs(45)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(46)  =   ":id=28,.fontname=MS Sans Serif"
            _StyleDefs(47)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=1"
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
            _StyleDefs(58)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(59)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=2"
            _StyleDefs(60)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(63)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14,.alignment=2"
            _StyleDefs(64)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14,.alignment=2"
            _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=0"
            _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14,.alignment=2"
            _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
            _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(82)  =   "Named:id=33:Normal"
            _StyleDefs(83)  =   ":id=33,.parent=0"
            _StyleDefs(84)  =   "Named:id=34:Heading"
            _StyleDefs(85)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(86)  =   ":id=34,.wraptext=-1"
            _StyleDefs(87)  =   "Named:id=35:Footing"
            _StyleDefs(88)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(89)  =   "Named:id=36:Selected"
            _StyleDefs(90)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(91)  =   "Named:id=37:Caption"
            _StyleDefs(92)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(93)  =   "Named:id=38:HighlightRow"
            _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(95)  =   "Named:id=39:EvenRow"
            _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(97)  =   "Named:id=40:OddRow"
            _StyleDefs(98)  =   ":id=40,.parent=33"
            _StyleDefs(99)  =   "Named:id=41:RecordSelector"
            _StyleDefs(100) =   ":id=41,.parent=34"
            _StyleDefs(101) =   "Named:id=42:FilterBar"
            _StyleDefs(102) =   ":id=42,.parent=33"
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
            TabIndex        =   9
            Top             =   75
            Width           =   1980
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Marcación de Asistencia"
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
            Left            =   9660
            TabIndex        =   10
            Top             =   0
            Width           =   2010
            Begin VB.Label lblperiodo 
               Alignment       =   2  'Center
               Caption         =   "lblperiodo(2)"
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
               Index           =   2
               Left            =   120
               TabIndex        =   11
               Top             =   330
               Width           =   1740
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
            Caption         =   "Detalle de Permiso"
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
      End
   End
End
Attribute VB_Name = "FrmPermiso1"
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

Dim mMesActivo As Integer '--indica el mes activo


Sub Cancelar()
    ActivaTool
    Label5.Caption = "Detalle de Marcación de Asistencia"
    QueHace = 3
    Bloquea False
    TabOne1.TabEnabled(0) = True
    Fg1.ColFormat(6) = FORMAT_HORA_AL_SEGUNDO
    Fg1.ColFormat(7) = FORMAT_HORA_AL_SEGUNDO
    
    Fg3.ColFormat(2) = FORMAT_HORA_SIN_SEGUNDO
    Fg3.ColFormat(3) = FORMAT_HORA_SIN_SEGUNDO
    
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
'    Dim nOrden As String
'    If fOrdenLista = False Then nOrden = "ASC"
'    If fOrdenLista = True Then nOrden = "DESC"
'    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
'    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = False
    mMesActivo = xMes
    pConfigurarGrilla
    pCargarGrid
    
    SeEjecuto = True
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado Asistencia, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            nuevo
        End If
    End If
End Sub


Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    QueHace = 3
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    
    '--dar formato al grid (presentacion)
    Dg1.Columns("fchemi").NumberFormat = FORMAT_DATE
    Dg1.Columns("fchfin").NumberFormat = FORMAT_DATE
    Dg1.Columns("horini").NumberFormat = FORMAT_HORA_SIN_SEGUNDO
    Dg1.Columns("horfin").NumberFormat = FORMAT_HORA_SIN_SEGUNDO
    
    TabOne1.CurrTab = 0
End Sub


Sub Blanquea()
    LimpiaText txt
    txtfecha.Valor = ""

    fg2.Rows = 1
    Fg1.Rows = 1
    Fg3.Rows = 1
    lblpersonal.Caption = ""
End Sub

Sub Bloquea(band As Boolean)
    txtfecha.Locked = Not band
    CmdDefault.Enabled = band
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

    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " al Personal ", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstHoras As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim xId&, A&
    Dim nSQL As String
    On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    If QueHace = 1 Then
        nSQL = "SELECT pla_marcacion.id FROM pla_marcacion WHERE pla_marcacion.dia = cdate('" & txtfecha.Valor & "');"
        RST_Busq RstTmp, nSQL, xCon
        If RstTmp.RecordCount <> 0 Then
            xId = RstTmp.Fields("id")
            RST_Busq RstCab, "SELECT * FROM pla_marcacion WHERE id = " & xId & "", xCon
            
        Else
            RST_Busq RstCab, "SELECT TOP 1 * FROM pla_marcacion", xCon
            RstCab.AddNew
            RstCab("id") = xId
            RstMarca("dia") = CDate(txtfecha.Valor)
            RstCab.Update
        End If
        
    Else
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM pla_marcacion WHERE id = " & xId & "", xCon
        '--eliminar datos de marcaciones
        xCon.Execute "DELETE * FROM pla_marcaciondet WHERE (((pla_marcaciondet.idori) In (1,5)) AND ((pla_marcaciondet.idmarca)=" & xId & "));"
        '--eliminar datos de tipos de horas
        xCon.Execute "DELETE *  FROM pla_marcacionhora WHERE (((pla_marcacionhora.idmarca)=" & xId & ") AND ((pla_marcacionhora.idhora) In (1,2,3,11,12,13)));"
    End If

    RST_Busq RstDet, "SELECT TOP 1 * FROM pla_marcaciondet ; ", xCon
    RST_Busq RstHoras, "SELECT TOP 1 * FROM pla_marcacionhora ; ", xCon

    '******************************************************
    Dim mRow&
    Dim nCompara As String
    Dim mCorr& '--el correlativo tanto para la marcacion como para los tipos de horas
    '--grabando el detalle
    For mRow = Fg1.FixedRows To Fg1.Rows - 1
        '--obtener el correlativo del detalle de la marcacion
        '--grabar solo asistencia, falta
        If NulosN(Fg1.TextMatrix(mRow, 3)) = 1 Or NulosN(Fg1.TextMatrix(mRow, 3)) = 5 Or NulosN(Fg1.TextMatrix(mRow, 3)) = -1 Then
            RST_Busq RstTmp, "SELECT TOP 1 pla_marcaciondet.corr FROM pla_marcaciondet " & _
                             "WHERE (((pla_marcaciondet.idmarca)=" & xId & ")) AND pla_marcaciondet.idemp = " & NulosN(Fg1.TextMatrix(mRow, 2)) & " " & _
                             "ORDER BY pla_marcaciondet.corr DESC; ", xCon
            If RstTmp.RecordCount <> 0 Then
                mCorr = NulosN(RstTmp.Fields(0)) + 1
            Else
                mCorr = 1
            End If
            Set RstTmp = Nothing
            '--
            RstDet.AddNew
            RstDet("idmarca") = xId
            RstDet("idemp") = NulosN(Fg1.TextMatrix(mRow, 2))
            RstDet("corr") = mCorr
            RstDet("hingreso") = TimeValue(Fg1.TextMatrix(mRow, 6))
            RstDet("hsalida") = TimeValue(Fg1.TextMatrix(mRow, 7))
            If NulosN(Fg1.TextMatrix(mRow, 3)) = -1 Then
                RstDet("idori") = 1
            Else
                RstDet("idori") = NulosN(Fg1.TextMatrix(mRow, 3))
            End If
            RstDet("tiporegistro") = 1
            RstDet.Update
        End If
    Next mRow
    
    '************************************************************
    Dim RstTipoHoras As New ADODB.Recordset
    
    For mRow = Fg1.FixedRows To Fg1.Rows - 1
        If nCompara <> NulosC(Fg1.TextMatrix(mRow, 2)) Then
            Set RstTipoHoras = pCalculoHoras(RstMarca, CDate(txtfecha.Valor), NulosN(Fg1.TextMatrix(mRow, 2)))
            RstTipoHoras.Filter = "idhora<4 or idhora>10"
            If RstTipoHoras.RecordCount <> 0 Then
                RstTipoHoras.MoveFirst
            End If
            Do While Not RstTipoHoras.EOF
                '--obtener el correlativo de los tipos de horas
                RST_Busq RstTmp, "SELECT TOP 1 pla_marcacionhora.corr FROM pla_marcacionhora " & _
                                 "WHERE (((pla_marcacionhora.idmarca)=" & xId & ")) AND pla_marcacionhora.idemp = " & NulosN(Fg1.TextMatrix(mRow, 2)) & " " & _
                                 "ORDER BY pla_marcacionhora.corr DESC; ", xCon
                                 
                If RstTmp.RecordCount <> 0 Then
                    mCorr = NulosN(RstTmp.Fields(0)) + 1
                Else
                    mCorr = 1
                End If
                Set RstTmp = Nothing
            
                RstHoras.AddNew
                RstHoras("idmarca") = xId
                RstHoras("idemp") = NulosN(Fg1.TextMatrix(mRow, 2))
                RstHoras("corr") = mCorr
                RstHoras("idhora") = NulosN(RstTipoHoras.Fields("idhora"))
                RstHoras("tothor") = TimeValue(RstTipoHoras.Fields("tothor"))
                RstHoras("totseg") = NulosN(RstTipoHoras.Fields("totseg"))
                RstHoras.Update
                
                RstTipoHoras.MoveNext
            Loop
            Set RstTipoHoras = Nothing
            nCompara = NulosC(Fg1.TextMatrix(mRow, 2))
        End If
    Next mRow
    
    Set RstTmp = Nothing
    '--calculando el resumen de horas por dia para actualizar
    nSQL = "SELECT pla_marcacion.id, pla_marcacion.dia, " _
        + vbCr + " (SELECT Sum(pla_marcacionhora1.totseg) AS TotHorNor FROM pla_marcacionhora AS pla_marcacionhora1 WHERE (((pla_marcacionhora1.idhora) In (1,4,5,6,7,8,9,10)))  GROUP BY pla_marcacionhora1.idmarca HAVING (((pla_marcacionhora1.idmarca)=pla_marcacion.id));) AS TotHorNor, " _
        + vbCr + " (SELECT Sum(pla_marcacionhora2.totseg) AS TotHorTar FROM pla_marcacionhora AS pla_marcacionhora2 WHERE (((pla_marcacionhora2.idhora)=2)) GROUP BY pla_marcacionhora2.idmarca HAVING (((pla_marcacionhora2.idmarca)=pla_marcacion.id)); ) AS TotHorTar, " _
        + vbCr + " (SELECT Sum(pla_marcacionhora3.totseg) AS TotHoEx FROM pla_marcacionhora AS pla_marcacionhora3 WHERE (((pla_marcacionhora3.idhora) In (11,12,13)))  GROUP BY pla_marcacionhora3.idmarca HAVING (((pla_marcacionhora3.idmarca)=pla_marcacion.id)); ) AS TotHoEx, " _
        + vbCr + " (SELECT Sum(pla_marcacionhora4.totseg) AS TotHorFal FROM pla_marcacionhora AS pla_marcacionhora4 WHERE (((pla_marcacionhora4.idhora)=3)) GROUP BY pla_marcacionhora4.idmarca HAVING (((pla_marcacionhora4.idmarca)=pla_marcacion.id)); ) AS TotHorFal " _
        + vbCr + " From pla_marcacion " _
        + vbCr + " WHERE (((pla_marcacion.id)=" & xId & "));"

    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then
        RstCab("tothn") = ConvertHora(NulosN(RstTmp.Fields("TotHorNor")))
        RstCab("totht") = ConvertHora(NulosN(RstTmp.Fields("TotHorTar")))
        RstCab("tothe") = ConvertHora(NulosN(RstTmp.Fields("TotHoEx")))
        RstCab("tothf") = ConvertHora(NulosN(RstTmp.Fields("TotHorFal")))
    End If
    RstCab.Update
    '---
    Set RstTmp = Nothing
    Set RstTipoHoras = Nothing
    '******************************************************
    Me.MousePointer = vbDefault
    MsgBox "Los datos de la Asistencia " + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstHoras = Nothing
    Grabar = True
    Exit Function
LaCague:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstHoras = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar la Asistencia por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Sub nuevo()
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea True
    Blanquea
    Fg1.ColFormat(6) = FORMAT_HORA_LARGO
    Fg1.ColFormat(7) = FORMAT_HORA_LARGO
    
    Fg1.SelectionMode = flexSelectionFree
    
    Label5.Caption = "Agregando Marcación de Asistencia"
End Sub

Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Label5.Caption = "Modificando Marcación de Asistencia"

    ActivaTool
    
    Bloquea True
    
    If TabOne1.CurrTab = 0 Then TabOne1.CurrTab = 1
        
    TabOne1.TabEnabled(0) = False
    
    Fg1.SelectionMode = flexSelectionFree
        
    QueHace = 2
    Fg1.ColFormat(6) = FORMAT_HORA_LARGO
    Fg1.ColFormat(7) = FORMAT_HORA_LARGO
    
    Fg3.ColFormat(2) = FORMAT_HORA_LARGO
    Fg3.ColFormat(3) = FORMAT_HORA_LARGO
    
    Agregando = False
    
    CargarDefault = True
    
    txtfecha.SetFocus

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
        xCon.Execute "DELETE FROM pla_marcaciondet WHERE (((pla_marcaciondet.idmarca)=" & xCod & ") AND ((pla_marcaciondet.idori) In (1,5,6,7)));"
        xCon.Execute "DELETE FROM pla_marcacionhora WHERE (((pla_marcacionhora.idmarca)=" & xCod & ") AND ((pla_marcacionhora.idhora) In (1,2,3,9,10,11,12,13)));"
        nSQL = "SELECT pla_marcaciondet.idmarca, pla_marcaciondet.idemp FROM pla_marcaciondet WHERE (((pla_marcaciondet.idmarca)=" & xCod & ")) " _
            + vbCr + " Union " _
            + vbCr + " SELECT pla_marcacionhora.idmarca,pla_marcacionhora.idemp FROM pla_marcacionhora WHERE (((pla_marcacionhora.idmarca)=" & xCod & ")) "
        RST_Busq RstTmp, nSQL, xCon
        If RstTmp.RecordCount = 0 Then
            xCon.Execute "DELETE FROM pla_marcacion WHERE pla_marcacion.id=" & xCod
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
    If Button.Index = 16 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Buscar()
'    Dim xform As New eps_librerias.FormBuscar
'    Dim xRs As New ADODB.Recordset
'
'    Dim xCampos(2, 4) As String
'
'    xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "apenom":      xCampos(0, 2) = "7000":    xCampos(0, 3) = "C"
'    xCampos(1, 0) = "Codigo":                xCampos(1, 1) = "codaut":      xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
'
'    xform.SQLCad = "SELECT UCase([pla_empleados]![ape])+', '+[pla_empleados]![nom] AS apenom, pla_empleados.codaut, pla_empleados.id " _
'        & " FROM pla_empleados WHERE activo = -1"
'
'    xform.Titulo = "Buscando Empleados"
'    xform.FormaBusca = Principio
'    xform.Criterio = ""
'    xform.Ordenado = "apenom"
'    xform.CampoBusca = "apenom"
'    Set xform.Coneccion = xCon
'    Set xRs = xform.BuscarReg(xCampos)
'    If xRs.State = 1 Then
'        RstFrm.MoveFirst
'        RstFrm.Find "id = " & xRs("id") & ""
'    End If
'    Set xRs = Nothing
'    Set xform = Nothing
End Sub

Sub Filtrar()
'    'Dim xform As New EPS_Buscar.Filtrar
'    Dim xform As New eps_librerias.FormFiltrar
'
'    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
'    Dim xCampos(4, 3) As String
'
'    xCampos(0, 0) = "Ape. y Nombres":     xCampos(0, 1) = "apenom":      xCampos(0, 2) = "C"
'    xCampos(1, 0) = "Cargo":              xCampos(1, 1) = "descargo":    xCampos(1, 2) = "C"
'    xCampos(2, 0) = "Tipo":               xCampos(2, 1) = "destipser":   xCampos(2, 2) = "C"
'    xCampos(3, 0) = "Basico":             xCampos(3, 1) = "basico":      xCampos(3, 2) = "N"
'
'    Set xform.rst = RstFrm
'    Set xform.Coneccion = xCon
'    xform.FiltrarReg xCampos
'    Set Dg1.DataSource = RstFrm
'    Dg1.Refresh

End Sub

Sub MuestraSegundoTab()
    
    Blanquea
    
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros para Mostrar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    QueHace = -1 '--comodin para entrar a [txt_cb_Validate]
    
    Agregando = True
    
    txtfecha(0).Valor = RstFrm("fchemi") '--fecha emision
    If NulosN(RstFrm("idemp")) <> 0 Then '--personal
        txt_cb(0).Text = NulosN(RstFrm("idemp"))
        txt_cb_Validate 0, False
    End If
    If NulosN(RstFrm("idper")) <> 0 Then '--motivo-permiso
        txt_cb(1).Text = NulosN(RstFrm("idper"))
        txt_cb_Validate 1, False
    End If
    If IsDate(RstFrm("fchini")) = True Then '--fecha inicio
        txtfecha(1).Valor = CDate(RstFrm("fchini"))
    Else
        txtfecha(1).Valor = ""
    End If
    If IsDate(RstFrm("fchfin")) = True Then '--fecha fin
        txtfecha(2).Valor = CDate(RstFrm("fchfin"))
    Else
        txtfecha(2).Valor = ""
    End If
    If IsDate(RstFrm("horini")) = True Then '--hora inicio
        dtpk(0).Value = CDate(RstFrm("horini"))
    Else
        dtpk(0).Value = ""
    End If
    If IsDate(RstFrm("horfin")) = True Then '--hora fin
        dtpk(1).Value = CDate(RstFrm("horfin"))
    Else
        dtpk(1).Value = ""
    End If
    
    txt(1).Text = RstFrm("observacion") '--observacion
    
    Agregando = False
    
    QueHace = 3
End Sub


Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL As String
    lblperiodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    lblperiodo(2).Caption = lblperiodo(0).Caption
    txtfecha.Valor = CDate("01/" & Format(mMesActivo, "00") & "/" & AnoTra)
    Dim RstTmp As New ADODB.Recordset
       
    nSQL = "SELECT pla_permiso.*, pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, mae_tipopermiso.descripcion AS permiso " _
        + vbCr + " FROM pla_empleados RIGHT JOIN (mae_tipopermiso RIGHT JOIN pla_permiso ON mae_tipopermiso.id = pla_permiso.idper) ON pla_empleados.id = pla_permiso.idemp " _
        + vbCr + " WHERE (((Month([fchemi])) = " & mMesActivo & ")) " _
        + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"

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

Private Sub CambiarMes()
    TabOne1.CurrTab = 0
    mMesActivo = SeleccionaMes(xCon)
    pCargarGrid
End Sub
