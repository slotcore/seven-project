VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmOrdenTrabajo 
   Caption         =   "Mantenimiento - Orden de Trabajo"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenTrabajo.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7320
      Left            =   0
      TabIndex        =   4
      Top             =   375
      Width           =   11880
      _cx             =   20955
      _cy             =   12912
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
      Appearance      =   1
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
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6900
         Left            =   -12435
         TabIndex        =   16
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6435
            Left            =   30
            TabIndex        =   17
            Top             =   360
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11351
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Orden"
            Columns(0).DataField=   "numdoc"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fch. Emi."
            Columns(1).DataField=   "fchemi"
            Columns(1).NumberFormat=   "dd/mm/yy"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Tipo Trabajo"
            Columns(2).DataField=   "destip"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Solicitante"
            Columns(3).DataField=   "apenom"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Area"
            Columns(4).DataField=   "desare"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Equipo / Local"
            Columns(5).DataField=   "descripcion"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Fch. Ini."
            Columns(6).DataField=   "fchini"
            Columns(6).NumberFormat=   "dd/mm/yy"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Fch. Fin."
            Columns(7).DataField=   "fchfin"
            Columns(7).NumberFormat=   "dd/mm/yy"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Importe"
            Columns(8).DataField=   "imp"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Estado"
            Columns(9).DataField=   "estado"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1879"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1799"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1561"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1482"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2090"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2011"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2990"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2910"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1905"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1826"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=3254"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=3175"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=1429"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1349"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=1508"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1429"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1667"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1588"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(55)=   "Column(9).Width=1429"
            Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1349"
            Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=516"
            Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
            _StyleDefs(76)  =   "Named:id=33:Normal"
            _StyleDefs(77)  =   ":id=33,.parent=0"
            _StyleDefs(78)  =   "Named:id=34:Heading"
            _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   ":id=34,.wraptext=-1"
            _StyleDefs(81)  =   "Named:id=35:Footing"
            _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(83)  =   "Named:id=36:Selected"
            _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=37:Caption"
            _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(87)  =   "Named:id=38:HighlightRow"
            _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=39:EvenRow"
            _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(91)  =   "Named:id=40:OddRow"
            _StyleDefs(92)  =   ":id=40,.parent=33"
            _StyleDefs(93)  =   "Named:id=41:RecordSelector"
            _StyleDefs(94)  =   ":id=41,.parent=34"
            _StyleDefs(95)  =   "Named:id=42:FilterBar"
            _StyleDefs(96)  =   ":id=42,.parent=33"
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Ordenes de Trabajo"
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
            Index           =   0
            Left            =   105
            TabIndex        =   19
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
            Caption         =   "LblPeriodo"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   9810
            TabIndex        =   18
            Top             =   0
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6900
         Left            =   45
         TabIndex        =   5
         Top             =   375
         Width           =   11790
         Begin VB.TextBox TxtObsItem 
            Height          =   2115
            Left            =   9630
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Text            =   "FrmOrdenTrabajo.frx":277E
            Top             =   4770
            Width           =   2115
         End
         Begin VB.Frame Frame6 
            Caption         =   "[ Estado de Tarea ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   1065
            Left            =   9840
            TabIndex        =   39
            Top             =   3435
            Width           =   1905
            Begin VB.CommandButton CmdTerminar 
               Caption         =   "&Terminar"
               Height          =   270
               Left            =   255
               TabIndex        =   40
               Top             =   300
               Width           =   1380
            End
            Begin VB.Label LblEstado2 
               Alignment       =   2  'Center
               Caption         =   "LblEstado2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   300
               Left            =   120
               TabIndex        =   41
               Top             =   660
               Width           =   1650
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Opciones de Administrador ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   615
            Left            =   8625
            TabIndex        =   36
            Top             =   330
            Width           =   3135
            Begin VB.CommandButton CmdCancel 
               Caption         =   "&Rechazar"
               Height          =   270
               Left            =   1575
               TabIndex        =   38
               Top             =   270
               Width           =   1260
            End
            Begin VB.CommandButton CmdAcepta 
               Caption         =   "&Aceptar"
               Height          =   270
               Left            =   270
               TabIndex        =   37
               Top             =   270
               Width           =   1260
            End
         End
         Begin VB.TextBox TxtObserva 
            Height          =   825
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Text            =   "FrmOrdenTrabajo.frx":2789
            Top             =   3675
            Width           =   9765
         End
         Begin VB.TextBox TxtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   8430
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "TxtTotal"
            Top             =   6510
            Width           =   1035
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   1725
            Left            =   45
            TabIndex        =   6
            Top             =   4770
            Width           =   9555
            _cx             =   16854
            _cy             =   3043
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
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmOrdenTrabajo.frx":2794
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
         Begin VB.Frame Frame5 
            Caption         =   "[ Estado de la Orden de Trabajo ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   705
            Left            =   8625
            TabIndex        =   23
            Top             =   945
            Width           =   3135
            Begin VB.Label LblEstado 
               Alignment       =   2  'Center
               Caption         =   "LblEstado"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   300
               Left            =   135
               TabIndex        =   24
               Top             =   300
               Width           =   2835
            End
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmOrdenTrabajo.frx":28E7
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1065
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipoCompra 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmOrdenTrabajo.frx":2A19
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   750
            Width           =   240
         End
         Begin VB.TextBox TxtNumOrd 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   13
            TabIndex        =   0
            Text            =   "TxtNumOrd"
            Top             =   405
            Width           =   1770
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   7275
            TabIndex        =   2
            Top             =   720
            Width           =   1260
            _ExtentX        =   2223
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
            Valor           =   "23/08/2007"
         End
         Begin VB.TextBox TxtTipo 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "TxtTipo"
            Top             =   720
            Width           =   915
         End
         Begin VB.TextBox TxtIdSol 
            Height          =   300
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "TxtIdSol"
            Top             =   1035
            Width           =   915
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   1470
            Left            =   45
            TabIndex        =   26
            Top             =   1935
            Width           =   11700
            _cx             =   20637
            _cy             =   2593
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
            Rows            =   7
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmOrdenTrabajo.frx":2B4B
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
         Begin VB.Frame Frame3 
            Height          =   495
            Left            =   60
            TabIndex        =   28
            Top             =   6405
            Width           =   7305
            Begin VB.CommandButton CmdSeleccionar 
               Caption         =   "Seleccionar Item"
               Enabled         =   0   'False
               Height          =   285
               Left            =   4305
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   150
               Width           =   1350
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "Agregar Item"
               Enabled         =   0   'False
               Height          =   285
               Left            =   1455
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   150
               Width           =   1350
            End
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "Eliminar Item"
               Enabled         =   0   'False
               Height          =   285
               Left            =   2880
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   150
               Width           =   1350
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Detalle del Item"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   9645
            TabIndex        =   43
            Top             =   4560
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Detalle de las tareas a realizar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   105
            TabIndex        =   35
            Top             =   3465
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total ==>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   7500
            TabIndex        =   33
            Top             =   6555
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Equipos a Trabajar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   105
            TabIndex        =   27
            Top             =   1710
            Width           =   1620
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Items de la Orden"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   105
            TabIndex        =   25
            Top             =   4560
            Width           =   1515
         End
         Begin VB.Label LblAutoriza 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblAutoriza"
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
            Left            =   1800
            TabIndex        =   22
            Top             =   1350
            Width           =   6720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Autorizado Por"
            Height          =   195
            Index           =   4
            Left            =   105
            TabIndex        =   21
            Top             =   1410
            Width           =   1035
         End
         Begin VB.Label LblNomSol 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomSol"
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
            Left            =   2760
            TabIndex        =   15
            Top             =   1035
            Width           =   5760
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante"
            Height          =   195
            Index           =   1
            Left            =   105
            TabIndex        =   14
            Top             =   1095
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Orden de Trabajo"
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
            TabIndex        =   13
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emision"
            Height          =   195
            Index           =   2
            Left            =   5745
            TabIndex        =   12
            Top             =   765
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo "
            Height          =   195
            Index           =   6
            Left            =   105
            TabIndex        =   11
            Top             =   765
            Width           =   360
         End
         Begin VB.Label LblTipo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipo"
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
            Left            =   2760
            TabIndex        =   10
            Top             =   720
            Width           =   2715
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Orden"
            Height          =   195
            Index           =   10
            Left            =   105
            TabIndex        =   9
            Top             =   450
            Width           =   660
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   20
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
         NumButtons      =   14
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
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Guia"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmOrdenTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstOrd As New ADODB.Recordset
Dim Quehace As Integer
Dim SeEjecuto As Boolean
Dim RstTmp As New ADODB.Recordset
Dim Agregando As Boolean

Dim Termina As Boolean  'para controlar si se muestran el boton de terminar una orden de trabajo
Dim Autoriza As Boolean 'para controlar si se muestran el boton de autorizar una orden de trabajo
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim mIdRegistro& '--identificador del registro

Function Grabar() As Boolean


    If NulosC(TxtNumOrd.Text) = "" Then
        MsgBox "No ha especificado el Nº de Orden de Trabajo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumOrd.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtTipo.Text) = "" Then
        MsgBox "No ha especificado el tipo de trabajo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipo.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchEmi.Valor) = "" Then
        MsgBox "No ha especificado la fecha de emision", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtIdSol.Text) = "" Then
        MsgBox "No ha especificado el nombre del solicitante", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdSol.SetFocus
        Exit Function
    End If
    
    If Fg2.Rows = 1 Then
        MsgBox "No ha especificado los equipos a trabajar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If
    
    If MsgBox("Seguro desea " + IIf(Quehace = 1, "grabar", "Modificar") + " el Registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDetItem As New ADODB.Recordset
    Dim A, B, xId As Integer
    
    If Quehace = 1 Then
'        If NulosC(TxtNumOrd.Text) = "" Then
            xId = HallaCodigoTabla("man_ordentrab", xCon, "id")
'        Else
'            xId = NulosN(TxtNumOrd.Text)
'        End If
                
        RST_Busq RstCab, "SELECT * FROM man_ordentrab", xCon
        RST_Busq RstDet, "SELECT * FROM man_ordentrabdet", xCon
        RST_Busq RstDetItem, "SELECT * FROM man_ordentrabdetitem", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstOrd("id")
        RST_Busq RstCab, "SELECT * FROM man_ordentrab WHERE id = " & RstOrd("id") & "", xCon
        xCon.Execute "DELETE * FROM man_ordentrabdet WHERE idord = " & RstOrd("id") & ""
        xCon.Execute "DELETE * FROM man_ordentrabdetitem WHERE id = " & RstOrd("id") & ""
    
        RST_Busq RstDet, "SELECT * FROM man_ordentrabdet", xCon
        RST_Busq RstDetItem, "SELECT * FROM man_ordentrabdetitem", xCon
    End If
    
    mIdRegistro = xId
    
    RstCab("numdoc") = NulosC(TxtNumOrd.Text)
    RstCab("idtipo") = NulosN(TxtTipo.Text)
    RstCab("fchemi") = TxtFchEmi.Valor
    RstCab("idsol") = NulosN(TxtIdSol.Text)
    RstCab("idaut") = 0
    RstCab("imp") = 0
    
    If Quehace = 1 Then RstCab("idestord") = 1
    
    
    RstCab.Update
    
    For A = 1 To Fg2.Rows - 1
        RstDet.AddNew
        RstDet("idord") = xId
        RstDet("idequi") = NulosN(Fg2.TextMatrix(A, 11))
        RstDet("idare") = NulosN(Fg2.TextMatrix(A, 9))
        RstDet("idclaequ") = NulosN(Fg2.TextMatrix(A, 10))
        If IsDate(Fg2.TextMatrix(A, 4)) = True Then RstDet("fchini") = CDate(Fg2.TextMatrix(A, 4))
        If IsDate(Fg2.TextMatrix(A, 5)) = True Then RstDet("fchfin") = CDate(Fg2.TextMatrix(A, 5))
        RstDet("imp") = 0
        RstDet("observa") = Fg2.TextMatrix(A, 12)
        If Quehace = 1 Then
            RstDet("idestado") = 1
        Else
            RstDet("idestado") = Fg2.TextMatrix(A, 14)
        End If
        RstDet.Update
        
        FiltrarTMP Val(Fg2.TextMatrix(A, 11)), 0
        If RstTmp.RecordCount <> 0 Then
            RstTmp.MoveFirst
            For B = 1 To RstTmp.RecordCount
                RstDetItem.AddNew
                RstDetItem("id") = xId
                RstDetItem("tipite") = RstTmp("idtipitem")
                RstDetItem("idequi") = NulosN(Fg2.TextMatrix(A, 11))
                RstDetItem("iditem") = NulosN(RstTmp("iditem"))
                RstDetItem("idunimed") = NulosN(RstTmp("idunimed"))
                RstDetItem("can") = NulosN(RstTmp("cantidad"))
                RstDetItem("preuni") = NulosN(RstTmp("preuni"))
                RstDetItem("total") = NulosN(RstTmp("total"))
                RstDetItem("observa") = NulosC(RstTmp("observa"))
                RstDetItem.Update
                RstTmp.MoveNext
                If RstTmp.EOF = True Then
                    Exit For
                End If
            Next B
        End If
    Next A
    
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDetItem = Nothing
    MsgBox "La orden de trabajo se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Grabar = True
    Exit Function
    
End Function

Sub Bloquea()
    TxtNumOrd.Locked = Not TxtNumOrd.Locked
    TxtTipo.Locked = Not TxtTipo.Locked
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    TxtIdSol.Locked = Not TxtIdSol.Locked
    
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
    CmdSeleccionar.Enabled = Not CmdSeleccionar.Enabled
End Sub

Sub Blanquea()
    TxtNumOrd.Text = ""
    TxtTipo.Text = ""
    TxtFchEmi.Valor = ""
    TxtIdSol.Text = ""
    TxtObserva.Text = ""
    TxtObsItem.Text = ""
    
    LblTipo.Caption = ""
    LblNomSol.Caption = ""
    LblAutoriza.Caption = ""
    TxtTotal.Text = ""
    Fg1.Rows = 1
    Fg2.Rows = 1
End Sub

Private Sub CmdAcepta_Click()
    If RstOrd.State = 0 Then Exit Sub
    If RstOrd.EOF = True Or RstOrd.BOF = True Or RstOrd.RecordCount = 0 Then Exit Sub
    xCon.Execute "UPDATE man_ordentrab SET man_ordentrab.idestord = 2 WHERE man_ordentrab.id = " & RstOrd("id") & ""
    RstOrd.Requery
    LblEstado.Caption = "Aprobada"
    LblEstado.ForeColor = &HC00000
End Sub

Private Sub CmdAddItem_Click()
    If Fg2.Row < 1 Then Exit Sub
    If NulosC(Fg2.TextMatrix(Fg2.Row, 1)) = "" Then
        MsgBox "No ha especificado el area para la tarea", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If NulosN(Fg2.TextMatrix(Fg2.Row, 11)) = 0 Then
        MsgBox "No ha especificado el equipo para la tarea", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg2.SetFocus
        Exit Sub
    End If
    Fg1.Rows = Fg1.Rows + 1
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = 1
    Fg1.SetFocus
End Sub

Private Sub CmdBusTipDoc_Click()
    If Quehace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Apellidos y nombres":   xCampos(0, 1) = "apenom":        xCampos(0, 2) = "6000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Cargo":                 xCampos(1, 1) = "descar":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Tipo":                  xCampos(2, 1) = "tipoper":       xCampos(2, 2) = "1500":         xCampos(2, 3) = "C"
    
    xForm.SQLCad = "SELECT man_personal.id, UCase([pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat])+', '+[pla_empleados]![nom] AS apenom, pla_cargos.descripcion AS descar, " _
        & " IIf([tipo]=1,'Solicitante','Autorizante') AS tipoper FROM man_personal LEFT JOIN (pla_cargos RIGHT JOIN pla_empleados " _
        & " ON pla_cargos.id = pla_empleados.idcargo) ON man_personal.idemp = pla_empleados.id ORDER BY UCase([pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat])+', '+[pla_empleados]![nom]"

    xForm.Titulo = "Buscando Personal"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "apenom"
    xForm.CampoBusca = "apenom"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdSol.Text = xRs("id")
            LblNomSol.Caption = xRs("apenom")
            Fg2.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipoCompra_Click()
    If Quehace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xForm.SQLCad = "SELECT man_tipo.* FROM man_tipo"
    
    xForm.Titulo = "Buscando Tipo Trabajo"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipo.Text = xRs("id")
            LblTipo.Caption = xRs("descripcion")
            TxtFchEmi.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Sub Nuevo()
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Quehace = 1
    Label5.Caption = "Agregando Orden de Trabajo"
    Blanquea
    Bloquea
    PreparaRST
    Fg2.ColComboList(1) = "|..."
    Fg2.ColComboList(2) = "|..."
    Fg2.ColComboList(3) = "|..."
    Fg2.ColComboList(13) = "Alta|Media|Baja"
    
    Fg1.ColComboList(1) = "|..."
    Fg1.ColComboList(2) = "|..."
    
    Fg1.SelectionMode = flexSelectionFree
    Fg2.SelectionMode = flexSelectionFree
    
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    
    Frame4.Visible = False
    CmdTerminar.Visible = False
    
    Fg1.Rows = 1
    Fg2.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    Fg2.Rows = Fg2.Rows + 1
    
    LblEstado.Caption = "Pendiente"
    LblEstado2.Caption = "Pendiente"
    LblEstado.ForeColor = &H8000&
    LblEstado2.ForeColor = &H8000&
    TabOne1.CurrTab = 1
    TxtNumOrd.SetFocus
End Sub

Sub Modificar()
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Quehace = 2
    Label5.Caption = "Modificando Orden de Trabajo"
    'Blanquea
    Bloquea
    'PreparaRST
    Fg2.ColComboList(1) = "|..."
    Fg2.ColComboList(2) = "|..."
    Fg2.ColComboList(3) = "|..."
    Fg2.ColComboList(13) = "Alta|Media|Baja"
    
    Fg1.ColComboList(1) = "|..."
    Fg1.ColComboList(2) = "|..."
    
    Fg1.SelectionMode = flexSelectionFree
    Fg2.SelectionMode = flexSelectionFree
    
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    'Fg1.Rows = 1
    'Fg2.Rows = 1
    'Fg1.Rows = Fg1.Rows + 1
    'Fg2.Rows = Fg2.Rows + 1
    TxtNumOrd.SetFocus
End Sub

Sub Eliminar()
    If RstOrd.EOF = True Or RstOrd.BOF = True Or RstOrd.RecordCount = 0 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    Dim xId&
    xId = NulosN(RstOrd.Fields("id"))
    TabOne1.CurrTab = 0
    If MsgBox("¿Esta seguro de eliminar el Número de Orden: " & NulosC(RstOrd("numdoc")) & " ?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
    
        xCon.Execute "DELETe * FROM man_ordentrabdetitem WHERE id = " & xId & ""
        xCon.Execute "DELETe * FROM man_ordentrabdet WHERE idord = " & xId & ""
        xCon.Execute "DELETe * FROM man_ordentrab WHERE id = " & xId & ""
        
        
        MsgBox "El Número de Orden : " & NulosC(RstOrd("numdoc")) & " Fue eliminado con éxito", vbInformation + vbOKOnly, xTitulo
        RstOrd.Requery
        Dg1.Refresh
        If RstOrd.RecordCount = 0 Then
            If MsgBox("No hay registrado ningún Orden de Requerimiento, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
                Nuevo
            End If
        End If
    End If
End Sub

Private Sub CmdCancel_Click()
    xCon.Execute "UPDATE man_ordentrab SET man_ordentrab.idestord = 3 WHERE man_ordentrab.id = " & RstOrd("id") & ""
    RstOrd.Requery
    LblEstado.Caption = "Rechazada"
    LblEstado.ForeColor = &HC0&
End Sub

Private Sub CmdDelItem_Click()
    If Fg1.Rows = 1 Then Exit Sub
    
    Fg1.RemoveItem Fg1.Row
End Sub

Private Sub CmdTerminar_Click()
    xCon.Execute "UPDATE man_ordentrabdet SET man_ordentrabdet.idestado = 2 WHERE (((man_ordentrabdet.idord)=" & RstOrd("id") & ") " _
        & " AND ((man_ordentrabdet.idequi)=" & NulosN(Fg2.TextMatrix(Fg2.Row, 11)) & "))"
        
    LblEstado2.Caption = "Terminado"
    LblEstado2.ForeColor = &HC00000
End Sub

Private Sub Dg1_DblClick()
    'TabOne1_Switch 0, 1, 0
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstOrd.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    If Col = 1 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
        
        xForm.SQLCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
        
        xForm.Titulo = "Buscando Tipo de Producto"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 1) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 9) = xRs("id")
            End If
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 2 Then
        Dim xCampos2(4, 4) As String

        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Unidad":        xCampos(1, 1) = "abrev":          xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":        xCampos(2, 1) = "codpro":         xCampos(2, 2) = "1500":         xCampos(2, 3) = "C"

        xForm.SQLCad = "SELECT alm_inventario.id, alm_inventario.idunimed, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.tippro " _
            & "  FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = 8 " _
            & " OR (alm_inventario.tippro) = " & NulosN(Fg1.TextMatrix(Fg1.Row, 9)) & ")) ORDER BY alm_inventario.descripcion"
    
        xForm.Titulo = "Buscando Items"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Agregando = True
                Fg1.TextMatrix(Fg1.Row, 2) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 3) = xRs("abrev")
                Fg1.TextMatrix(Fg1.Row, 7) = xRs("id")
                Fg1.TextMatrix(Fg1.Row, 8) = xRs("idunimed")
                
                FiltrarTMP NulosN(Fg2.TextMatrix(Fg2.Row, 11)), NulosN(Fg1.TextMatrix(Fg1.Row, 7))
                
                If RstTmp.RecordCount = 0 Then
                    RstTmp.AddNew
                End If
                RstTmp("iditem") = xRs("id")
                RstTmp("idunimed") = xRs("idunimed")
                RstTmp("descripcion") = xRs("descripcion")
                RstTmp("desuni") = xRs("abrev")
                RstTmp("idequipo") = NulosN(Fg2.TextMatrix(Fg2.Row, 11))
                
                RstTmp("idtipitem") = NulosN(Fg1.TextMatrix(Fg1.Row, 9))
                RstTmp("destipitem") = (Fg1.TextMatrix(Fg1.Row, 1))
                Agregando = False
            End If
        End If
    End If
    
    If Col = 3 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
        
        xForm.SQLCad = "SELECT mae_unidades.* FROM mae_unidades"
        
        xForm.Titulo = "Buscando Unidades"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 3) = xRs("abrev")
                Fg1.TextMatrix(Fg1.Row, 8) = xRs("id")
                
                FiltrarTMP NulosN(Fg2.TextMatrix(Fg2.Row, 11)), NulosN(Fg1.TextMatrix(Fg1.Row, 7))
                If RstTmp.RecordCount <> 0 Then
                    RstTmp("descripcion") = xRs("abrev")
                    RstTmp("desuni") = xRs("abrev")
                End If
            End If
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)

    If Agregando = True Then Exit Sub
    If Col = 4 Or Col = 5 Then
        Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) * NulosN(Fg1.TextMatrix(Fg1.Row, 5))
        Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), "0.00")
    End If
    FiltrarTMP NulosN(Fg2.TextMatrix(Fg2.Row, 11)), NulosN(Fg1.TextMatrix(Fg1.Row, 7))
    If RstTmp.RecordCount <> 0 Then
        If Fg1.Col = 4 Then
            RstTmp("cantidad") = NulosN(Fg1.TextMatrix(Fg1.Row, 4))
        End If
        If Fg1.Col = 5 Then
            RstTmp("preuni") = NulosN(Fg1.TextMatrix(Fg1.Row, 5))
        End If
        RstTmp("total") = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) * NulosN(Fg1.TextMatrix(Fg1.Row, 5))
    End If
    SumarTotales
    
End Sub

Sub SumarTotales()
    Dim A As Integer
    Dim xTotal As Double
    
    For A = 1 To Fg1.Rows - 1
        xTotal = xTotal + NulosN(Fg1.TextMatrix(A, 6))
    Next A
    TxtTotal.Text = Format(xTotal, "0.00")
End Sub

Private Sub Fg1_EnterCell()
    If Fg1.Col = 7 Then
        Fg1.Editable = flexEDNone
    Else
        Fg1.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        Fg1.Rows = Fg1.Rows + 1
    End If
    
    If KeyCode = 46 Then
        If Fg1.Rows = 1 Then Exit Sub
        Fg1.RemoveItem Fg1.Row
    End If
End Sub

Private Sub Fg1_RowColChange()
    If Fg1.Rows = 1 Then Exit Sub
    If Fg2.Rows <= 1 Then Exit Sub
    FiltrarTMP NulosN(Fg2.TextMatrix(Fg2.Row, 11)), NulosN(Fg1.TextMatrix(Fg1.Row, 7))
    If RstTmp.RecordCount <> 0 Then
        TxtObsItem.Text = RstTmp("observa")
    End If
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    If Col = 1 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
        
        xForm.SQLCad = "SELECT pla_area.* FROM pla_area"
        
        xForm.Titulo = "Buscando Area"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg2.TextMatrix(Fg2.Row, 1) = xRs("descripcion")
                Fg2.TextMatrix(Fg2.Row, 9) = xRs("id")
            End If
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 2 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
        
        xForm.SQLCad = "SELECT man_equipoclase.* FROM man_equipoclase"
        
        xForm.Titulo = "Buscando Equipo"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg2.TextMatrix(Fg2.Row, 2) = xRs("descripcion")
                Fg2.TextMatrix(Fg2.Row, 10) = xRs("id")
            End If
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If

    If Col = 3 Then
        xCampos(0, 0) = "Equipo":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Area":     xCampos(1, 1) = "desare":         xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
        
        xForm.SQLCad = "SELECT man_equipo.*, pla_area.descripcion AS desare  FROM man_equipo LEFT JOIN pla_area ON man_equipo.idarea = pla_area.id " _
            & " WHERE (((man_equipo.idclaequ)=" & NulosN(Fg2.TextMatrix(Fg2.Row, 10)) & "))"
        
        xForm.Titulo = "Buscando Equipos"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg2.TextMatrix(Fg2.Row, 3) = xRs("descripcion")
                Fg2.TextMatrix(Fg2.Row, 11) = xRs("id")
            End If
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
End Sub

Sub FiltrarTMP(Idequipo As Integer, IdItem As Integer)
    RstTmp.Filter = adFilterNone
    If IdItem <> 0 Then
        RstTmp.Filter = "idequipo = " & Idequipo & " AND iditem = " & IdItem & ""
    Else
        RstTmp.Filter = "idequipo = " & Idequipo & " "
    End If
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        Fg2.Rows = Fg2.Rows + 1
    End If
    If KeyCode = 46 Then
        If Fg2.Rows = 1 Then Exit Sub
        If Fg2.Row < 1 Then Exit Sub
        Fg2.RemoveItem Fg2.Row
        Fg2_RowColChange
        
        If Fg2.Rows = 1 Then
            Fg2_RowColChange
            Fg2.Rows = 2
            Fg2.Row = 1
            Fg2.Col = 2
            Fg2.SetFocus
        End If
        
    End If
End Sub

Private Sub Fg2_RowColChange()
    TxtObserva.Text = ""
    Fg1.Rows = 1
    
    If Fg2.Rows = 0 Then Exit Sub
    If Fg2.Row < 1 Then Exit Sub
    
    TxtObserva.Text = NulosC(Fg2.TextMatrix(Fg2.Row, 12))
    
    If NulosN(Fg2.TextMatrix(Fg2.Row, 14)) = 1 Then
        If Termina = True Then CmdTerminar.Visible = True
        LblEstado2.Caption = "Pendiente"
        LblEstado2.ForeColor = &H8000&
    Else
        CmdTerminar.Visible = False
        LblEstado2.Caption = "Terminado"
        LblEstado2.ForeColor = &HC00000
    End If
    
    FiltrarTMP NulosN(Fg2.TextMatrix(Fg2.Row, 11)), 0
    If RstTmp.RecordCount <> 0 Then
        MuestraItemTMP
    End If
End Sub

Sub MuestraItemTMP()
    Dim A As Integer
    RstTmp.MoveFirst
    
    If RstTmp.RecordCount <> 0 Then
        Fg1.Rows = 1
        Agregando = True
        RstTmp.MoveFirst
        For A = 1 To RstTmp.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            
            Fg1.TextMatrix(A, 1) = RstTmp("destipitem")
            Fg1.TextMatrix(A, 2) = RstTmp("descripcion")
            Fg1.TextMatrix(A, 3) = RstTmp("desuni")
            Fg1.TextMatrix(A, 4) = Format(RstTmp("cantidad"), "0.00")
            Fg1.TextMatrix(A, 5) = Format(RstTmp("preuni"), "0.00")
            Fg1.TextMatrix(A, 6) = Format(RstTmp("total"), "0.00")
            Fg1.TextMatrix(A, 7) = RstTmp("iditem")
            Fg1.TextMatrix(A, 8) = RstTmp("idunimed")
            Fg1.TextMatrix(A, 9) = RstTmp("idtipitem")
            
            RstTmp.MoveNext
            If RstTmp.EOF = True Then Exit For
        Next A
        Agregando = False
    End If
    SumarTotales
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        
        RST_Busq RstOrd, "SELECT man_ordentrab.id,man_ordentrab.numdoc, man_ordentrab.fchemi, man_tipo.descripcion AS destip, UCase(pla_empleados.apepat & ' ' & pla_empleados.apemat)& ', '& pla_empleados!nom AS apenom, " _
            & " pla_area.descripcion AS desare, man_equipo.descripcion, man_ordentrabdet.fchini, man_ordentrabdet.fchfin, man_ordentrabdet.[imp], man_ordentrab.idaut," _
            & " IIf([idestado]=1,'Pendiente','Culminado') AS estado, man_ordentrab.idtipo, man_ordentrab.idsol, idestord FROM ((man_tipo RIGHT JOIN (man_ordentrab " _
            & " LEFT JOIN (man_equipo RIGHT JOIN (man_ordentrabdet LEFT JOIN pla_area ON man_ordentrabdet.idare = pla_area.id) ON man_equipo.id = man_ordentrabdet.idequi) " _
            & " ON man_ordentrab.id = man_ordentrabdet.idord) ON man_tipo.id = man_ordentrab.idtipo) LEFT JOIN man_personal ON man_ordentrab.idsol = man_personal.id) " _
            & " LEFT JOIN pla_empleados ON man_personal.idemp = pla_empleados.id", xCon

        Set Dg1.DataSource = RstOrd
        Dg1.Refresh
        
        If RstOrd.RecordCount = 0 Then
            Dim Rpta As Integer
            Rpta = MsgBox("No se han registrado ordenes de trabajo ¿Desea agregar uno ahora?", vbInformation + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Unload Me
                Set RstOrd = Nothing
            End If
        End If
    End If
End Sub

Sub MuestraSegundoTab()
    Dim RstFg1 As New ADODB.Recordset
    Dim A As Integer
    Blanquea
    If RstOrd.State = 0 Then Exit Sub
    If RstOrd.EOF = True Or RstOrd.BOF = True Or RstOrd.RecordCount = 0 Then Exit Sub

    TxtNumOrd.Text = NulosC(RstOrd("numdoc"))
    TxtTipo.Text = NulosC(RstOrd("idtipo"))
    LblTipo.Caption = NulosC(RstOrd("destip"))
    TxtIdSol.Text = NulosC(RstOrd("idsol"))
    LblNomSol.Caption = NulosC(RstOrd("apenom"))
    
    If IsDate(RstOrd("fchemi")) = True Then TxtFchEmi.Valor = CDate(RstOrd("fchemi"))
    
    Set RstFg1 = BuscaConCriterio("SELECT UCase([pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat])+', '+[pla_empleados]![nom] AS apenom, man_personal.id" _
        & " FROM man_personal LEFT JOIN pla_empleados ON man_personal.idemp = pla_empleados.id WHERE (((man_personal.id)=" & RstOrd("idaut") & "))", xCon)

    If RstFg1.RecordCount <> 0 Then
        LblAutoriza.Caption = RstFg1("apenom")
    Else
        LblAutoriza.Caption = ""
    End If
    
    PreparaRST
    
    RST_Busq RstFg1, "SELECT man_ordentrabdet.*, man_equipo.descripcion AS desequi, pla_area.descripcion AS desarea, man_equipoclase.descripcion AS desequicla," _
        & "  IIf([man_ordentrabdet]![idestado]=1,'Pendiente','Terminado') AS desestado " _
        & " FROM (man_equipoclase RIGHT JOIN (man_equipo RIGHT JOIN man_ordentrabdet ON man_equipo.id = man_ordentrabdet.idequi) " _
        & " ON man_equipoclase.id = man_ordentrabdet.idclaequ) LEFT JOIN pla_area ON man_ordentrabdet.idare = pla_area.id WHERE " _
        & " (((man_ordentrabdet.idord)=" & RstOrd("id") & "))", xCon
    
    If RstFg1.RecordCount <> 0 Then
        RstFg1.MoveFirst
        For A = 1 To RstFg1.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = RstFg1("desarea")
            Fg2.TextMatrix(A, 2) = RstFg1("desequicla")
            Fg2.TextMatrix(A, 3) = RstFg1("desequi")
            Fg2.TextMatrix(A, 4) = Format(RstFg1("fchini"), "dd/mm/yy")
            Fg2.TextMatrix(A, 5) = Format(RstFg1("fchfin"), "dd/mm/yy")
            Fg2.TextMatrix(A, 6) = Format(RstFg1("fchter"), "dd/mm/yy")
            Fg2.TextMatrix(A, 7) = RstFg1("diadif")
            Fg2.TextMatrix(A, 8) = Format(RstFg1("imp"), "0.00")
            
            Fg2.TextMatrix(A, 9) = RstFg1("idare")
            Fg2.TextMatrix(A, 10) = RstFg1("idclaequ")
            Fg2.TextMatrix(A, 11) = RstFg1("idequi")
            Fg2.TextMatrix(A, 12) = RstFg1("observa")
            Fg2.TextMatrix(A, 14) = RstFg1("idestado")
            RstFg1.MoveNext
            
            If RstFg1.EOF = True Then Exit For
        Next A
    End If
    
    Set RstFg1 = Nothing

    RST_Busq RstFg1, "SELECT man_ordentrabdetitem.*, mae_unidades.abrev AS desunimed, alm_inventario.descripcion AS descitem, mae_tipoproducto.descripcion AS destippro " _
        & " FROM ((man_ordentrabdetitem LEFT JOIN alm_inventario ON man_ordentrabdetitem.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON " _
        & " man_ordentrabdetitem.idunimed = mae_unidades.id) LEFT JOIN mae_tipoproducto ON man_ordentrabdetitem.tipite = mae_tipoproducto.id " _
        & " Where (((man_ordentrabdetitem.id) = " & RstOrd("id") & ")) ORDER BY alm_inventario.descripcion", xCon

    If RstFg1.RecordCount <> 0 Then
        RstFg1.MoveFirst
        For A = 1 To RstFg1.RecordCount
            RstTmp.AddNew
            RstTmp("iditem") = RstFg1("iditem")
            RstTmp("idunimed") = RstFg1("idunimed")
            RstTmp("descripcion") = RstFg1("descitem")
            RstTmp("desuni") = RstFg1("desunimed")
            RstTmp("cantidad") = RstFg1("can")
            RstTmp("preuni") = RstFg1("preuni")
            RstTmp("total") = RstFg1("total")
            RstTmp("idequipo") = RstFg1("idequi")
            RstTmp("idtipitem") = RstFg1("tipite")
            RstTmp("destipitem") = NulosC(RstFg1("destippro"))
            RstTmp("observa") = NulosC(RstFg1("observa"))
            RstFg1.MoveNext
            If RstFg1.EOF = True Then Exit For
            
        Next A
    End If
    
    FiltrarTMP Val(Fg2.TextMatrix(Fg2.Row, 11)), 0
    TxtObserva.Text = NulosC(Fg2.TextMatrix(Fg2.Row, 12))
    
    If RstTmp.RecordCount <> 0 Then
        MuestraItemTMP
    End If
    
    If RstOrd("idestord") = 1 Or NulosN(RstOrd("idestord")) = 0 Then
        LblEstado.Caption = "Pendiente"
        LblEstado.ForeColor = &H8000&
    End If
    If RstOrd("idestord") = 2 Then
        LblEstado.Caption = "Aprobada"
        LblEstado.ForeColor = &HC00000
    End If
    If RstOrd("idestord") = 3 Then
        LblEstado.Caption = "Rechazada"
        LblEstado.ForeColor = &HC0&
    End If

    If RstOrd("idestord") = 2 Then
        CmdAcepta.Enabled = False
    Else
        CmdAcepta.Enabled = True
    End If

    'mostramos el estado del trabajo
    If NulosN(Fg2.TextMatrix(Fg2.Row, 14)) = 1 Then
        If Termina = True Then CmdTerminar.Visible = True
        LblEstado2.Caption = "Pendiente"
        LblEstado2.ForeColor = &H8000&
    Else
        CmdTerminar.Visible = False
        LblEstado2.Caption = "Terminado"
        LblEstado2.ForeColor = &HC00000
    End If

End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Form_Load()
    TabOne1.CurrTab = 0
    Quehace = 3
    SeEjecuto = False
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    Fg2.ColWidth(9) = 0
    Fg2.ColWidth(10) = 0
    Fg2.ColWidth(11) = 0
    Fg2.ColWidth(12) = 0
    Fg2.ColWidth(14) = 0
    
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow

    Dim rst As New ADODB.Recordset
    Set rst = BuscaConCriterio("SELECT * FROM mae_usuarios WHERE id =" & xIdUsuario & " ", xCon)
    
    If rst.RecordCount <> 0 Then
        If rst("autoriza") = -1 Then
            Autoriza = True
            Frame4.Visible = True
        Else
            Autoriza = False
            Frame4.Visible = False
        End If
        
        If rst("termina") = -1 Then
            Termina = True
            CmdTerminar.Visible = True
        Else
            Termina = False
            CmdTerminar.Visible = False
        End If
    Else
        Frame4.Visible = False
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If Quehace = 3 Then MuestraSegundoTab
    End If
End Sub

Sub Cancelar()
    Quehace = 3
    Frame4.Visible = True
    ActivaTool
    Label5.Caption = "Detalle Orden de Trabajo"
    
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.Refresh
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstOrd.Requery
            Cancelar
            If RstOrd.RecordCount <> 0 Then
                RstOrd.MoveFirst
                RstOrd.Find "id=" & mIdRegistro
                If RstOrd.EOF = True Then RstOrd.MoveFirst
            End If
            Dg1.Refresh
            Dg1.SetFocus
            
            
        End If
    End If
    If Button.Index = 6 Then
        Cancelar
    End If
    
    If Button.Index = 14 Then
        Set RstOrd = Nothing
        Unload Me
    End If
End Sub

Private Sub TxtIdSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdSol_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc_Click
    End If
End Sub

Private Sub TxtNumOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumOrd_Validate(Cancel As Boolean)
    '----------
    
    If Quehace = 3 Then Exit Sub
    If NulosC(TxtNumOrd.Text) <> "" Then
    
        TxtNumOrd.Text = Format(TxtNumOrd.Text, "0000000000")
        
        Dim rst As New ADODB.Recordset
        Dim nSQL As String
        '--ver si existe el numero de doc
        If Quehace <> 1 Then nSQL = " and man_ordentrab.id <> " & NulosN(RstOrd("id"))
        
        RST_Busq rst, "SELECT man_ordentrab.id, man_ordentrab.numdoc,man_ordentrab.fchemi, man_tipo.descripcion AS destip, UCase(pla_empleados.apepat & ' ' & pla_empleados.apemat) & ', ' & pla_empleados!nom AS apenom " _
                & " FROM man_tipo RIGHT JOIN ((man_ordentrab LEFT JOIN man_personal ON man_ordentrab.idsol = man_personal.id) LEFT JOIN pla_empleados ON man_personal.idemp = pla_empleados.id) ON man_tipo.id = man_ordentrab.idtipo " _
                & " WHERE (((man_ordentrab.numdoc)='" & NulosC(TxtNumOrd.Text) & "')) " & nSQL, xCon
                
        If rst.RecordCount <> 0 Then
            '--poner el nuevo numero doc
            
            MsgBox "La Orden de Trabajo ya existe " & vbCr & "Nº Orden:         " & NulosC(rst("numdoc")) & vbCr & "Fecha Emi.       " & NulosC(rst("fchemi")) & vbCr & "Solicitado Por:  " & NulosC(rst("apenom")), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtNumOrd.Text = ""
            TxtNumOrd.SetFocus
        End If
        Set rst = Nothing
        
    End If
    
End Sub

Private Sub TxtObserva_Validate(Cancel As Boolean)
    Fg2.TextMatrix(Fg2.Row, 12) = NulosC(TxtObserva.Text)
End Sub

Private Sub TxtObsItem_Validate(Cancel As Boolean)
    If Quehace = 3 Then Exit Sub
    On Error Resume Next
    If Fg1.Row < 1 Then
        MsgBox "Ingrese primero el Item de la Orden", vbExclamation, xTitulo
        TxtObsItem.Text = ""
        CmdAddItem.SetFocus
        Exit Sub
    End If
    FiltrarTMP Val(Fg2.TextMatrix(Fg2.Row, 11)), Val(Fg1.TextMatrix(Fg1.Row, 7))
    If RstTmp.RecordCount <> 0 Then
        RstTmp("observa") = TxtObsItem.Text
    End If
    Err.Clear
End Sub

Private Sub TxtTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipoCompra_Click
    End If
End Sub

Private Sub TxtTipo_Validate(Cancel As Boolean)
    If TxtTipo.Text = "" Then Exit Sub
    LblTipo.Caption = Busca_Codigo(Val(TxtTipo.Text), "id", "descripcion", "man_tipo", "N", xCon)
    If LblTipo.Caption = "" Then
        LblTipo.Caption = ""
        TxtTipo.Text = ""
    End If
End Sub

Private Sub VSFlexGrid1_Click()

End Sub

Sub PreparaRST()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(11, 3) As String

    xCampos(0, 0) = "iditem":        xCampos(0, 1) = "N":      xCampos(0, 2) = "2"
    xCampos(1, 0) = "idunimed":      xCampos(1, 1) = "N":      xCampos(1, 2) = "2"
    xCampos(2, 0) = "descripcion":   xCampos(2, 1) = "C":      xCampos(2, 2) = "200"
    xCampos(3, 0) = "desuni":        xCampos(3, 1) = "C":      xCampos(3, 2) = "5"
    xCampos(4, 0) = "cantidad":      xCampos(4, 1) = "D":      xCampos(4, 2) = "2"
    xCampos(5, 0) = "preuni":        xCampos(5, 1) = "D":      xCampos(5, 2) = "2"
    xCampos(6, 0) = "total":         xCampos(6, 1) = "D":      xCampos(6, 2) = "2"
    xCampos(7, 0) = "idequipo":      xCampos(7, 1) = "N":      xCampos(7, 2) = "2"
    xCampos(8, 0) = "idtipitem":     xCampos(8, 1) = "N":      xCampos(8, 2) = "2"
    xCampos(9, 0) = "destipitem":    xCampos(9, 1) = "C":      xCampos(9, 2) = "50"
    xCampos(10, 0) = "observa":      xCampos(10, 1) = "C":     xCampos(10, 2) = "200"
    
    Set RstTmp = xFun.CrearRstTMP(xCampos)

    RstTmp.Open
End Sub

