VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmAsistencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Marcación de Asistencia"
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
            Picture         =   "FrmAsistencia.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":1EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":23EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsistencia.frx":277E
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
            Columns(1).Caption=   "Fecha"
            Columns(1).DataField=   "dia"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Dia"
            Columns(2).DataField=   "dianom"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Total HN"
            Columns(3).DataField=   "tothn"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Total HT"
            Columns(4).DataField=   "totht"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Total HE"
            Columns(5).DataField=   "tothe"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Total HF"
            Columns(6).DataField=   "tothf"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1217"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1138"
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
            Splits(0)._ColumnProps(24)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(30)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(36)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
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
            _StyleDefs(74)  =   "Named:id=33:Normal"
            _StyleDefs(75)  =   ":id=33,.parent=0"
            _StyleDefs(76)  =   "Named:id=34:Heading"
            _StyleDefs(77)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(78)  =   ":id=34,.wraptext=-1"
            _StyleDefs(79)  =   "Named:id=35:Footing"
            _StyleDefs(80)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(81)  =   "Named:id=36:Selected"
            _StyleDefs(82)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(83)  =   "Named:id=37:Caption"
            _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(85)  =   "Named:id=38:HighlightRow"
            _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(87)  =   "Named:id=39:EvenRow"
            _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(89)  =   "Named:id=40:OddRow"
            _StyleDefs(90)  =   ":id=40,.parent=33"
            _StyleDefs(91)  =   "Named:id=41:RecordSelector"
            _StyleDefs(92)  =   ":id=41,.parent=34"
            _StyleDefs(93)  =   "Named:id=42:FilterBar"
            _StyleDefs(94)  =   ":id=42,.parent=33"
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
            TabIndex        =   15
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
            TabIndex        =   21
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
               TabIndex        =   22
               Top             =   330
               Width           =   1740
            End
         End
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Caption         =   "Frame12"
            Height          =   405
            Index           =   12
            Left            =   90
            TabIndex        =   19
            Top             =   6315
            Width           =   11550
            Begin VB.Label lblpersonal 
               AutoSize        =   -1  'True
               Caption         =   "lblpersonal"
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
               Height          =   195
               Left            =   75
               TabIndex        =   20
               Top             =   75
               Width           =   8655
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               Index           =   4
               X1              =   15
               X2              =   15
               Y1              =   0
               Y2              =   380
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               Index           =   2
               X1              =   11535
               X2              =   11535
               Y1              =   -15
               Y2              =   365
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
         Begin VB.CommandButton CmdDefault 
            Caption         =   "Predeterminado"
            Enabled         =   0   'False
            Height          =   300
            Left            =   2070
            TabIndex        =   12
            Top             =   450
            Width           =   1575
         End
         Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
            Height          =   300
            Left            =   585
            TabIndex        =   11
            Top             =   450
            Width           =   1395
            _ExtentX        =   2461
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
            Valor           =   "02/04/2008"
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
            Height          =   5130
            Left            =   120
            TabIndex        =   9
            Top             =   1035
            Width           =   8040
            _cx             =   14182
            _cy             =   9049
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
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmAsistencia.frx":2B10
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
         Begin VSFlex7Ctl.VSFlexGrid fg2 
            Height          =   2850
            Left            =   8280
            TabIndex        =   17
            Top             =   1035
            Width           =   3390
            _cx             =   5980
            _cy             =   5027
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
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmAsistencia.frx":2C08
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
         Begin VSFlex7Ctl.VSFlexGrid Fg3 
            Height          =   1905
            Left            =   8280
            TabIndex        =   18
            Top             =   4260
            Width           =   3390
            _cx             =   5980
            _cy             =   3360
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
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmAsistencia.frx":2C87
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
         Begin VB.Label lblgrid 
            BackStyle       =   0  'Transparent
            Caption         =   "Horario"
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
            Height          =   270
            Index           =   2
            Left            =   8280
            TabIndex        =   16
            Top             =   4035
            Width           =   1830
         End
         Begin VB.Label lblgrid 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipos de Horas"
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
            Height          =   270
            Index           =   1
            Left            =   8280
            TabIndex        =   14
            Top             =   825
            Width           =   1830
         End
         Begin VB.Label lblgrid 
            BackStyle       =   0  'Transparent
            Caption         =   "Personal"
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
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   825
            Width           =   1830
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dia"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   555
            Width           =   240
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
            Caption         =   "Detalle de Marcación de Asistencia"
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
Attribute VB_Name = "FrmAsistencia"
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
    Label5.Caption = "Detalle de Marcación de Asistencia"
    QueHace = 3
    Bloquea False
    TabOne1.TabEnabled(0) = True
    fg1.ColFormat(6) = FORMAT_HORA_AL_SEGUNDO
    fg1.ColFormat(7) = FORMAT_HORA_AL_SEGUNDO
    
    Fg3.ColFormat(2) = FORMAT_HORA_SIN_SEGUNDO
    Fg3.ColFormat(3) = FORMAT_HORA_SIN_SEGUNDO
    
    fg1.SelectionMode = flexSelectionByRow
    Me.MousePointer = vbDefault
    TabOne1.CurrTab = 0
End Sub

Private Sub CmdDefault_Click()
    If IsDate(txtfecha.Valor) = False Then
        MsgBox "La fecha no es correcta", vbExclamation, xTitulo
        Exit Sub
    End If
    CargarDefault = True
    pMuestraDetalle
    
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

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub
    If Col = 6 Or Col = 7 Then
        '--invocar al formulario de horas
        Dim obj As New SGI2_funciones.formularios
        obj.HoraSeleccionar fg1, Row, Col, fg1.TextMatrix(Row, Col)
        Set obj = Nothing
    End If
    Exit Sub
salir:
    Agregando = False
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    If Col < 4 Then Exit Sub
    RstMarca.Filter = "idgrid=" & fg1.TextMatrix(Row, 1)
    Select Case Col
        Case 6 '--hora inicial
            If IsDate(fg1.TextMatrix(Row, 6)) = True Then
                If IsDate(fg1.TextMatrix(Row, 7)) = True Then
                    If TimeValue(fg1.TextMatrix(Row, 7)) < TimeValue(fg1.TextMatrix(Row, 6)) Then
                        MsgBox "La Hora Inicial es Superior a la Hora Final", vbExclamation, xTitulo
                        fg1.TextMatrix(Row, Col) = ""
                        RstMarca.Fields("hini") = Null
                    Else
                        RstMarca.Fields("hini") = TimeValue(fg1.TextMatrix(Row, 6))
                    End If
                Else
                    RstMarca.Fields("hini") = TimeValue(fg1.TextMatrix(Row, 6))
                End If
            Else
                fg1.TextMatrix(Row, Col) = ""
                RstMarca.Fields("hini") = Null
            End If
            
        Case 7 '--hora final
            If IsDate(fg1.TextMatrix(Row, 7)) = True Then
                If IsDate(fg1.TextMatrix(Row, 6)) = True Then
                    If TimeValue(fg1.TextMatrix(Row, 7)) < TimeValue(fg1.TextMatrix(Row, 6)) Then
                        MsgBox "La Hora Final es Inferior a la Hora Inicial", vbExclamation, xTitulo
                        fg1.TextMatrix(Row, Col) = ""
                        RstMarca.Fields("hfin") = Null
                    Else
                        RstMarca.Fields("hfin") = TimeValue(fg1.TextMatrix(Row, 7))
                    End If
                End If
            Else
                fg1.TextMatrix(Row, Col) = ""
                RstMarca.Fields("hfin") = Null
            End If

        Case 5 '--Origen
            If NulosN(fg1.Cell(flexcpText, Row, Col)) = 0 Then
                fg1.TextMatrix(Row, 3) = ""
                fg1.TextMatrix(Row, 5) = ""
                RstMarca.Fields("idori") = Null
                RstMarca.Fields("origen") = Null
            Else
                If NulosN(fg1.Cell(flexcpText, Row, Col)) = 5 Then '--si selecciona origen =falta
                    RstMarca.Filter = "idemp=" & NulosN(fg1.TextMatrix(Row, 2)) & " and idori = 1 and idgrid <> " & NulosN(fg1.TextMatrix(Row, 1))
                    If RstMarca.RecordCount <> 0 Then
                        MsgBox "Cuando selecciona el Origen Falta, se considera que será todo el dia" + vbCr + "Este Personal Tiene Registros cuyo Origen no es Falta", vbExclamation, xTitulo
                        '--aplicando el filtro segun fila del grid
                        RstMarca.Filter = "idgrid=" & fg1.TextMatrix(Row, 1)
                        
                        RstMarca.Fields("idori") = Null
                        RstMarca.Fields("origen") = Null
                        fg1.TextMatrix(Row, 3) = ""
                        fg1.TextMatrix(Row, 5) = ""
                    Else
                        '--aplicando el filtro segun fila del grid
                        fg1.TextMatrix(Row, 3) = NulosN(fg1.Cell(flexcpText, Row, Col))
                        RstMarca.Filter = "idgrid=" & fg1.TextMatrix(Row, 1)
                        
                        RstMarca.Fields("idori") = NulosN(fg1.TextMatrix(Row, 3))
                        RstMarca.Fields("origen") = fg1.TextMatrix(Row, 5)
                    End If
                Else '--si selecciona origen =Asistencia
                    RstMarca.Filter = "idemp=" & NulosN(fg1.TextMatrix(Row, 2)) & " and idori = 5 and idgrid <> " & NulosN(fg1.TextMatrix(Row, 1))
                    If RstMarca.RecordCount <> 0 Then
                        MsgBox "Cuando selecciona el Origen Asistencia, No es posible registrar Otro Registro con Origen Falta" + vbCr + "Este Personal Tiene Registros cuyo Origen es Falta", vbExclamation, xTitulo
                        RstMarca.Fields("idori") = Null
                        RstMarca.Fields("origen") = Null
                        fg1.TextMatrix(Row, 3) = ""
                        fg1.TextMatrix(Row, 5) = ""
                    Else
                        '--aplicando el filtro segun fila del grid
                        fg1.TextMatrix(Row, 3) = NulosN(fg1.Cell(flexcpText, Row, Col))
                        
                        RstMarca.Filter = "idgrid=" & fg1.TextMatrix(Row, 1)
                        
                        RstMarca.Fields("idori") = NulosN(fg1.TextMatrix(Row, 3))
                        RstMarca.Fields("origen") = fg1.TextMatrix(Row, 5)
                    End If
                End If
            End If
    
    End Select
    If RstMarca.RecordCount <> 0 Then RstMarca.Update
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
        Case 5, 6, 7
            '--si es diferente a 2:permiso, 3:licencia, 4:vacaciones
            '--1:asistencia, 5:falta
            If NulosN(fg1.TextMatrix(fg1.Row, 3)) = 1 Or NulosN(fg1.TextMatrix(fg1.Row, 3)) = 5 Or NulosN(fg1.TextMatrix(fg1.Row, 3)) = 0 Then
                fg1.Editable = flexEDKbdMouse
            ElseIf NulosN(fg1.TextMatrix(fg1.Row, 3)) = -1 Then '--por defecto
                fg1.Editable = flexEDKbdMouse
            Else
                fg1.Editable = flexEDNone
            End If
        Case Else
            fg1.Editable = flexEDNone
    End Select
    
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
    
    pMuestraTipoHoras fg1.TextMatrix(fg1.Row, 2)
        
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = False
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
    TabOne1.CurrTab = 0
End Sub


Sub Blanquea()
    LimpiaText txt
    txtfecha.Valor = ""

    fg2.Rows = 1
    fg1.Rows = 1
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
    For mRow = fg1.FixedRows To fg1.Rows - 1
        '--obtener el correlativo del detalle de la marcacion
        '--grabar solo asistencia, falta
        If NulosN(fg1.TextMatrix(mRow, 3)) = 1 Or NulosN(fg1.TextMatrix(mRow, 3)) = 5 Or NulosN(fg1.TextMatrix(mRow, 3)) = -1 Then
            RST_Busq RstTmp, "SELECT TOP 1 pla_marcaciondet.corr FROM pla_marcaciondet " & _
                             "WHERE (((pla_marcaciondet.idmarca)=" & xId & ")) AND pla_marcaciondet.idemp = " & NulosN(fg1.TextMatrix(mRow, 2)) & " " & _
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
            RstDet("idemp") = NulosN(fg1.TextMatrix(mRow, 2))
            RstDet("corr") = mCorr
            RstDet("hingreso") = TimeValue(fg1.TextMatrix(mRow, 6))
            RstDet("hsalida") = TimeValue(fg1.TextMatrix(mRow, 7))
            If NulosN(fg1.TextMatrix(mRow, 3)) = -1 Then
                RstDet("idori") = 1
            Else
                RstDet("idori") = NulosN(fg1.TextMatrix(mRow, 3))
            End If
            RstDet("tiporegistro") = 1
            RstDet.Update
        End If
    Next mRow
    
    '************************************************************
    Dim RstTipoHoras As New ADODB.Recordset
    
    For mRow = fg1.FixedRows To fg1.Rows - 1
        If nCompara <> NulosC(fg1.TextMatrix(mRow, 2)) Then
            Set RstTipoHoras = pCalculoHoras(RstMarca, CDate(txtfecha.Valor), NulosN(fg1.TextMatrix(mRow, 2)))
            RstTipoHoras.Filter = "idhora<4 or idhora>10"
            If RstTipoHoras.RecordCount <> 0 Then
                RstTipoHoras.MoveFirst
            End If
            Do While Not RstTipoHoras.EOF
                '--obtener el correlativo de los tipos de horas
                RST_Busq RstTmp, "SELECT TOP 1 pla_marcacionhora.corr FROM pla_marcacionhora " & _
                                 "WHERE (((pla_marcacionhora.idmarca)=" & xId & ")) AND pla_marcacionhora.idemp = " & NulosN(fg1.TextMatrix(mRow, 2)) & " " & _
                                 "ORDER BY pla_marcacionhora.corr DESC; ", xCon
                                 
                If RstTmp.RecordCount <> 0 Then
                    mCorr = NulosN(RstTmp.Fields(0)) + 1
                Else
                    mCorr = 1
                End If
                Set RstTmp = Nothing
            
                RstHoras.AddNew
                RstHoras("idmarca") = xId
                RstHoras("idemp") = NulosN(fg1.TextMatrix(mRow, 2))
                RstHoras("corr") = mCorr
                RstHoras("idhora") = NulosN(RstTipoHoras.Fields("idhora"))
                RstHoras("tothor") = TimeValue(RstTipoHoras.Fields("tothor"))
                RstHoras("totseg") = NulosN(RstTipoHoras.Fields("totseg"))
                RstHoras.Update
                
                RstTipoHoras.MoveNext
            Loop
            Set RstTipoHoras = Nothing
            nCompara = NulosC(fg1.TextMatrix(mRow, 2))
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
    fg1.ColFormat(6) = FORMAT_HORA_LARGO
    fg1.ColFormat(7) = FORMAT_HORA_LARGO
    
    fg1.SelectionMode = flexSelectionFree
    
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
    
    fg1.SelectionMode = flexSelectionFree
        
    QueHace = 2
    fg1.ColFormat(6) = FORMAT_HORA_LARGO
    fg1.ColFormat(7) = FORMAT_HORA_LARGO
    
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
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros para Mostrar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    Blanquea
    QueHace = -1 '--comodin para entrar a [txt_cb_Validate]
    txtfecha.Valor = RstFrm.Fields("dia")
    CargarDefault = False
    pMuestraDetalle
    QueHace = 3
End Sub


Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL As String
    lblperiodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
    lblperiodo(2).Caption = lblperiodo(0).Caption
    txtfecha.Valor = CDate("01/" & Format(xMes, "00") & "/" & AnoTra)
    Dim RstTmp As New ADODB.Recordset
       
    nSQL = "SELECT pla_marcacion.id, pla_marcacion.dia, Format([pla_marcacion].[dia],'dddd') AS dianom,pla_marcacion.tothn, pla_marcacion.totht, pla_marcacion.tothe, pla_marcacion.tothf, " _
        + vbCr + " (SELECT Sum(pla_marcacionhora1.totseg) AS TotHorNor FROM pla_marcacionhora AS pla_marcacionhora1 WHERE (((pla_marcacionhora1.idhora) In (1,4,5,6,7,8,9,10)))  GROUP BY pla_marcacionhora1.idmarca HAVING (((pla_marcacionhora1.idmarca)=pla_marcacion.id));) AS TotHorNor, " _
        + vbCr + " (SELECT Sum(pla_marcacionhora2.totseg) AS TotHorTar FROM pla_marcacionhora AS pla_marcacionhora2 WHERE (((pla_marcacionhora2.idhora)=2)) GROUP BY pla_marcacionhora2.idmarca HAVING (((pla_marcacionhora2.idmarca)=pla_marcacion.id)); ) AS TotHorTar, " _
        + vbCr + " (SELECT Sum(pla_marcacionhora3.totseg) AS TotHoEx FROM pla_marcacionhora AS pla_marcacionhora3 WHERE (((pla_marcacionhora3.idhora) In (11,12,13)))  GROUP BY pla_marcacionhora3.idmarca HAVING (((pla_marcacionhora3.idmarca)=pla_marcacion.id)); ) AS TotHoEx, " _
        + vbCr + " (SELECT Sum(pla_marcacionhora4.totseg) AS TotHorFal FROM pla_marcacionhora AS pla_marcacionhora4 WHERE (((pla_marcacionhora4.idhora)=3)) GROUP BY pla_marcacionhora4.idmarca HAVING (((pla_marcacionhora4.idmarca)=pla_marcacion.id)); ) AS TotHorFal " _
        + vbCr + " From pla_marcacion " _
        + vbCr + " WHERE (((Year([pla_marcacion].[dia]))=" & AnoTra & ") AND ((Month([pla_marcacion].[dia]))= " & xMes & " )) " _
        + vbCr + " ORDER BY pla_marcacion.dia;"

    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    DEFINIR_RST_TMP RstTmp, RstFrm
    CARGAR_RST_TMP RstTmp, RstFrm
    Set RstFrm = RstTmp
    
    If RstFrm.RecordCount <> 0 Then RstFrm.MoveFirst
    Do While Not RstFrm.EOF
        RstFrm.Fields("tothn") = ConvertHora(NulosN(RstFrm.Fields("TotHorNor")))
        RstFrm.Fields("totht") = ConvertHora(NulosN(RstFrm.Fields("TotHorTar")))
        RstFrm.Fields("tothe") = ConvertHora(NulosN(RstFrm.Fields("TotHoEx")))
        RstFrm.Fields("tothf") = ConvertHora(NulosN(RstFrm.Fields("TotHorFal")))
        RstFrm.MoveNext
    Loop
    
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
    band = Validar(txt)
    
    If IsDate(txtfecha.Valor) = False Then
        MsgBox "Falta especificar la fecha de marcación", vbExclamation, xTitulo
        txtfecha.SetFocus
        Exit Function
    End If

    '--de las marcaciones
    If fg1.Rows = 1 Then
        MsgBox "Falta Ingresar la lista de Marcaciones" + vbCr + "Puede previsualizar las marcaciones por defecto luego proceda a modificarlos", vbExclamation, xTitulo
        CmdDefault.SetFocus
        Exit Function
    End If
    Dim mRow&, mCol&
    mCol = -1
    For mRow = 1 To fg1.Rows - 1
        If NulosC(fg1.TextMatrix(mRow, 5)) = "" Then   '--origen
            MsgBox "Falta especificar el Origen", vbExclamation, xTitulo
            mCol = 5:          Exit For
        ElseIf IsDate(fg1.TextMatrix(mRow, 6)) = False Then
            MsgBox "Falta especificar la Hora de Inicio de " & fg1.TextMatrix(mRow, 5), vbExclamation, xTitulo
            mCol = 6:          Exit For
        ElseIf IsDate(fg1.TextMatrix(mRow, 7)) = False Then
            MsgBox "Falta especificar la Hora Final de " & fg1.TextMatrix(mRow, 5), vbExclamation, xTitulo
            mCol = 7:          Exit For
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
        .Rows = 1
        .Cols = 9
        .FixedRows = 1
        .RowHeight(0) = 300
        .ColWidth(1) = 0:
        .TextMatrix(0, 1) = "Id":               .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "Idemp":            .ColWidth(2) = 0:
        .TextMatrix(0, 3) = "IdOri":            .ColWidth(3) = 0:
        
        .TextMatrix(0, 4) = "Apellidos y Nombres":  .ColWidth(4) = 3300:    .ColAlignment(4) = flexAlignLeftCenter:     .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 5) = "Origen":               .ColWidth(5) = 950:    .ColAlignment(5) = flexAlignLeftCenter:      .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 6) = "H.Ingreso":             .ColWidth(6) = 1350:    .ColAlignment(6) = flexAlignCenterCenter:   .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 7) = "H.Salida":                .ColWidth(7) = 1350:    .ColAlignment(7) = flexAlignCenterCenter:   .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 8) = "Tipo":                 .ColWidth(8) = 0:     .ColAlignment(8) = flexAlignLeftCenter:     .Row = 0: .Col = 8: .CellAlignment = flexAlignLeftCenter
        
        .ColFormat(6) = FORMAT_HORA_AL_SEGUNDO
        .ColFormat(7) = FORMAT_HORA_AL_SEGUNDO
        
        .ColEditMask(6) = "##:##:##"
        .ColEditMask(7) = "##:##:##"
        
        .SelectionMode = flexSelectionByRow
    End With
    
    GRID_COMBOLIST fg1, 6
    GRID_COMBOLIST fg1, 7
    '--combolist
    Dim RstTmp As New ADODB.Recordset
    Dim tFormat$
    RST_Busq RstTmp, "SELECT pla_origenes.id, pla_origenes.descripcion FROM pla_origenes WHERE (((pla_origenes.id) In (1,5)))  ORDER BY pla_origenes.descripcion;", xCon
    tFormat = fg1.BuildComboList(RstTmp, "descripcion", "id", vbYellow)
    fg1.ColComboList(5) = tFormat
    Set RstTmp = Nothing
    
    With fg2 '--tipos de horas
        .Rows = 1
        .Cols = 5
        .ColWidth(1) = 200
        .FixedRows = 1
        .RowHeight(0) = 300
        .ColWidth(1) = 0:
        .TextMatrix(0, 1) = "Idemp":        .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "Descripción":  .ColWidth(2) = 2200:   .ColAlignment(2) = flexAlignLeftCenter:         .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "Total":        .ColWidth(3) = 900:    .ColAlignment(3) = flexAlignLeftCenter:         .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 4) = "totseg":       .ColWidth(4) = 0:      .ColAlignment(4) = flexAlignRightCenter
        
        .ColFormat(3) = FORMAT_HORA_LARGO
        
        .SelectionMode = flexSelectionByRow
    End With
    
    With Fg3 '--horario
        .Rows = 1
        .Cols = 4
        .ColWidth(1) = 200
        .FixedRows = 1
        .RowHeight(0) = 300
        .ColWidth(1) = 0:
        .TextMatrix(0, 1) = "Descripción":  .ColWidth(1) = 1300:   .ColAlignment(1) = flexAlignLeftCenter:   .Row = 0: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 2) = "H.Ingreso":      .ColWidth(2) = 900:    .ColAlignment(2) = flexAlignCenterCenter:   .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "H.Salida":       .ColWidth(3) = 900:    .ColAlignment(3) = flexAlignCenterCenter:  .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter

        .ColFormat(2) = FORMAT_HORA_SIN_SEGUNDO
        .ColFormat(3) = FORMAT_HORA_SIN_SEGUNDO
        
        .SelectionMode = flexSelectionByRow
    End With
    '*****************************************
    DoEvents
End Sub

Private Sub pRegistroAdd()
    Dim mCol&
    Dim mRow&
'    If QueHace = 3 Then Exit Sub
    Agregando = True
    If fg1.Rows > 1 Then
        If IsDate(fg1.TextMatrix(fg1.Row, 6)) = False Then
            MsgBox "Falta ingresar la Hora de Inicio", vbExclamation, xTitulo
            mCol = 5
        ElseIf IsDate(fg1.TextMatrix(fg1.Row, 7)) = False Then
            MsgBox "Falta ingresar la Hora Final", vbExclamation, xTitulo
            mCol = 6
        ElseIf NulosC(fg1.TextMatrix(fg1.Row, 5)) = "" Then
            MsgBox "Falta ingresar el Origen de la Marcación", vbExclamation, xTitulo
            mCol = 7
        End If
        If mCol <> 0 Then
            fg1.Col = mCol
            fg1.SetFocus
            Agregando = False
            Exit Sub
        End If
        
        mRow = fg1.Row + 1
        
        GRID_INSERT fg1, mRow, e_Fila
        fg1.TextMatrix(mRow, 1) = fg1.Rows - 1
        fg1.TextMatrix(mRow, 2) = fg1.TextMatrix(mRow - 1, 2)
        fg1.TextMatrix(mRow, 3) = -1
        fg1.TextMatrix(mRow, 4) = fg1.TextMatrix(mRow - 1, 4)
        '--agregando registro
        RstMarca.AddNew
        RstMarca.Fields("idgrid") = fg1.Rows - 1
        RstMarca.Fields("idemp") = fg1.TextMatrix(mRow - 1, 2)
        RstMarca.Fields("idori") = -1
        RstMarca.Fields("nombres") = fg1.TextMatrix(mRow - 1, 4)
        RstMarca.Fields("origen") = Null
        RstMarca.Update
    End If
    '--agrupar -----------------------
    Me.MousePointer = vbHourglass
    GRID_AGRUPAR fg1, 2
    Me.MousePointer = vbDefault
    '-------------------------------
    fg1.Row = mRow
    fg1.Col = 4
    fg1.SetFocus
    Agregando = False
End Sub

Private Sub pRegistroDel()
    Dim mRowDel&, mRow&
    If fg1.Rows = 1 Then Exit Sub
    If fg1.Row < 1 Then Exit Sub
    
    If NulosN(fg1.TextMatrix(fg1.Row, 3)) = 1 Or NulosN(fg1.TextMatrix(fg1.Row, 3)) = 5 Then
        fg1.Editable = flexEDKbdMouse
    ElseIf NulosN(fg1.TextMatrix(fg1.Row, 3)) = -1 Then '--por defecto
        fg1.Editable = flexEDKbdMouse
    Else
        '--no se podra eliminar
        MsgBox "No se puede eliminar este registro" + vbCr + "Origen: " & fg1.TextMatrix(fg1.Row, 5), vbExclamation, xTitulo
        Exit Sub
    End If
    
    
    If MsgBox("Seguro desea eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    mRowDel = fg1.TextMatrix(fg1.Row, 1)
    mRow = fg1.Row
    fg1.RemoveItem fg1.Row
    '--eliminar el registro del recordset
    RstMarca.Filter = "idgrid=" & mRowDel
    RstMarca.Delete
    RstMarca.Update
    '--agrupar -----------------------
    Me.MousePointer = vbHourglass
    Agregando = True
    GRID_AGRUPAR fg1, 2
    Agregando = False
    
    If mRow > 1 Then
        fg1.Row = mRow - 1
    ElseIf mRow = 1 Then
        fg1.Row = 1
    End If
    fg1.Col = 4
    Me.MousePointer = vbDefault
    '-------------------------------
End Sub

Private Sub CambiarMes()
    TabOne1.CurrTab = 0
    xMes = SeleccionaMes(xCon)
    pCargarGrid
End Sub


Private Sub pMuestraDetalle()
    '--default= false: carga los datos desde la base de datos(inf. registrada previamente)
    '--         true:  carga los datos de base de datos, actualiza la informacion relacionada a vacaciones,permiso,licencia, dia festivo
                     ' adicionalemente agregar mas marcaciones que faltan registrar
    Dim nSQL As String
    fg1.Rows = 1
    fg2.Rows = 1
    Fg3.Rows = 1
    Set RstMarca = Nothing
    Dim RstTmp As New ADODB.Recordset
    If CargarDefault = False Then
        nSQL = "SELECT 0 as IdGrid,pla_empleados.id AS idemp, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS nombres, pla_marcaciondet.hingreso AS hini, pla_marcaciondet.hsalida AS hfin, pla_marcaciondet.idori, pla_origenes.descripcion AS origen, pla_marcaciondet.tiporegistro AS tipreg " _
            + vbCr + " FROM pla_empleados INNER JOIN (pla_marcacion INNER JOIN (pla_marcaciondet INNER JOIN pla_origenes ON pla_marcaciondet.idori = pla_origenes.id) ON pla_marcacion.id = pla_marcaciondet.idmarca) ON pla_empleados.id = pla_marcaciondet.idemp " _
            + vbCr + " WHERE (((pla_marcacion.id)= " & RstFrm.Fields("id") & ")) " _
            + vbCr + " ORDER BY pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom, pla_marcaciondet.hingreso;"
            
        RST_Busq RstTmp, nSQL, xCon
        DEFINIR_RST_TMP RstMarca, RstTmp
        CARGAR_RST_TMP RstMarca, RstTmp
        Set RstTmp = Nothing
        
    Else
        Me.MousePointer = vbHourglass
        pMacacionDia CDate(txtfecha.Valor), e_Asist_DomingoDiaFestivos
        If xCon.State = 0 Then Exit Sub
        
        Set RstMarca = fMarcacionDefault(CDate(txtfecha.Valor))
        Me.MousePointer = vbDefault
    End If
    '**********************************************************************************************************
    Agregando = True
    If RstMarca.RecordCount <> 0 Then
        RstMarca.MoveFirst
    End If
    Do While Not RstMarca.EOF
        fg1.Rows = fg1.Rows + 1
        '------
        fg1.TextMatrix(fg1.Rows - 1, 1) = fg1.Rows - 1
        RstMarca.Fields("idgrid") = NulosN(fg1.Rows - 1)
        RstMarca.Update
       
        fg1.TextMatrix(fg1.Rows - 1, 2) = NulosN(RstMarca("idemp"))
        fg1.TextMatrix(fg1.Rows - 1, 3) = NulosC(RstMarca("idori"))
        fg1.TextMatrix(fg1.Rows - 1, 4) = NulosC(RstMarca("nombres"))
        fg1.TextMatrix(fg1.Rows - 1, 5) = NulosC(RstMarca("origen"))
        fg1.TextMatrix(fg1.Rows - 1, 6) = NulosC(RstMarca("hini"))
        fg1.TextMatrix(fg1.Rows - 1, 7) = NulosC(RstMarca("hfin"))
'        fg1.TextMatrix(fg1.Rows - 1, 5) = NulosC(RstMarca("tipo"))
        RstMarca.MoveNext
    Loop
    '--poner los colores en el grid
    '--agrupar -----------------------
    Me.MousePointer = vbHourglass
    GRID_AGRUPAR fg1, 2
    Me.MousePointer = vbDefault
    '-------------------------------
    '--
    
    If fg1.Rows > 1 Then
        fg1.Row = 1
        fg1.Col = 3
'        fg1.SetFocus
    End If
    Agregando = False
   
    Me.MousePointer = vbDefault
End Sub

Private Sub pMuestraTipoHoras(IdEmp&)
    If Agregando = True Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim mTotalSeg&
    
    fg2.Rows = 1
    Fg3.Rows = 1
    
    lblpersonal.Caption = StrConv(fg1.TextMatrix(fg1.Row, 4), 3)
    '---------------
    Set RstTmp = pCalculoHoras(RstMarca, CDate(txtfecha.Valor), IdEmp, IIf(CargarDefault = True, -1, RstFrm.Fields("id")))
    '----------------
    
    If RstTmp.State = 0 Then Exit Sub
    
    Do While Not RstTmp.EOF
        fg2.Rows = fg2.Rows + 1
        fg2.TextMatrix(fg2.Rows - 1, 1) = NulosN(RstTmp("idhora"))
        fg2.TextMatrix(fg2.Rows - 1, 2) = NulosC(RstTmp("descripcion"))
        fg2.TextMatrix(fg2.Rows - 1, 3) = NulosC(RstTmp("tothor"))
        fg2.TextMatrix(fg2.Rows - 1, 4) = NulosC(RstTmp("totseg"))
        
        RstTmp.MoveNext
    Loop
    
    Set RstTmp = Nothing
    mTotalSeg = GRID_SUMAR_COL(fg2, 4)
    If mTotalSeg <> 0 Then
        fg2.Rows = fg2.Rows + 2
        fg2.TextMatrix(fg2.Rows - 1, 2) = "Total Horas >>"
        fg2.TextMatrix(fg2.Rows - 1, 3) = ConvertHora(mTotalSeg)
        '--fila
        GRID_COLOR_FONDO fg2, fg2.Rows - 1, 1, fg2.Rows - 1, fg2.Cols - 1, &HE0FEFE
        '--columna
        GRID_COLOR_FONDO fg2, 1, 3, fg2.Rows - 1, 3, &HE0FEFE
    End If
    
    fg2.Row = fg2.Rows - 1
    
    pMuestraHorario IdEmp
    
End Sub

Private Sub pMuestraHorario(IdEmp&)
    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    
    nSQL = "SELECT mae_horariohora.idhora, mae_tipohora.descripcion, mae_horariohora.hingreso, mae_horariohora.hsalida " _
        + vbCr + " FROM (mae_tipohora INNER JOIN mae_horariohora ON mae_tipohora.id = mae_horariohora.idhora) INNER JOIN mae_horarioemp ON mae_horariohora.idhor = mae_horarioemp.idhor " _
        + vbCr + " WHERE (((mae_horarioemp.IdEmp) = " & IdEmp & ")) AND ((mae_horarioemp.vigencia)=-1) " _
        + vbCr + " ORDER BY mae_tipohora.prioridad; "
    
    RST_Busq RstTmp, nSQL, xCon
        
    RstMarca.Filter = "idemp=" & fg1.TextMatrix(fg1.Row, 2)
    
    If RstTmp.State = 0 Then Exit Sub
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
    Else
        Fg3.Rows = Fg3.Rows + 2
        UNIR_CELDAS Fg3, 1, 1, 1, 3, "No tiene Horario", flexAlignCenterCenter, True
        UNIR_CELDAS Fg3, 2, 1, 2, 3, "Configure el Horario", flexAlignCenterCenter, True
        GRID_COLOR_FONDO Fg3, 1, 1, 2, 1, &H8282FF
        Set RstTmp = Nothing
        Exit Sub
    End If
    
    Do While Not RstTmp.EOF
        Fg3.Rows = Fg3.Rows + 1
        Fg3.TextMatrix(Fg3.Rows - 1, 1) = NulosC(RstTmp("descripcion"))
        Fg3.TextMatrix(Fg3.Rows - 1, 2) = NulosC(RstTmp("hingreso"))
        Fg3.TextMatrix(Fg3.Rows - 1, 3) = NulosC(RstTmp("hsalida"))
        RstTmp.MoveNext
    Loop
    Set RstTmp = Nothing
    
    GRID_COLOR_FONDO Fg3, 1, 2, Fg3.Rows - 1, 2, &HC4C4FF
    GRID_COLOR_FONDO Fg3, 1, 3, Fg3.Rows - 1, 3, &HB3B3FF
    
End Sub

