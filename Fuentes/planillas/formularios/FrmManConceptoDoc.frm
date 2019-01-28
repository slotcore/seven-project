VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManConceptoDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Concepto"
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoDoc.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoDoc.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoDoc.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoDoc.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoDoc.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoDoc.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoDoc.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoDoc.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoDoc.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoDoc.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoDoc.frx":1EA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   10
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
      TabIndex        =   11
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
         TabIndex        =   14
         Top             =   375
         Width           =   11745
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6405
            Left            =   45
            TabIndex        =   15
            Top             =   390
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   11298
            _LayoutType     =   4
            _RowHeight      =   14
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "CodSunat"
            Columns(0).DataField=   "codsun"
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
            Columns(3).Caption=   "Tipo"
            Columns(3).DataField=   "tiponombre"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Categoría"
            Columns(4).DataField=   "catnombre"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=6773"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6694"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=3201"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3122"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=4789"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=4710"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
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
            _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
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
            TabIndex        =   16
            Top             =   60
            Width           =   11400
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6795
         Left            =   45
         TabIndex        =   12
         Top             =   375
         Width           =   11745
         Begin VB.TextBox txt 
            BackColor       =   &H0000FFFF&
            Height          =   285
            Index           =   0
            Left            =   10665
            TabIndex        =   19
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   -15
            Visible         =   0   'False
            Width           =   945
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6165
            Left            =   -75
            TabIndex        =   17
            Top             =   420
            Width           =   11700
            _cx             =   20637
            _cy             =   10874
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
            Caption         =   " Datos Concepto | Datos Contables "
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   0
            Position        =   2
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
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Height          =   6075
               Left            =   12660
               TabIndex        =   21
               Top             =   45
               Width           =   11295
            End
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   6075
               Left            =   360
               TabIndex        =   18
               Top             =   45
               Width           =   11295
               Begin VB.TextBox txt 
                  Height          =   330
                  Index           =   4
                  Left            =   6060
                  MaxLength       =   4
                  TabIndex        =   5
                  Tag             =   "null"
                  Text            =   "txt(4)"
                  Top             =   1830
                  Width           =   1350
               End
               Begin VB.TextBox txt 
                  Height          =   330
                  Index           =   3
                  Left            =   1350
                  MaxLength       =   30
                  TabIndex        =   4
                  Tag             =   "null"
                  Text            =   "txt(3)"
                  Top             =   1815
                  Width           =   3585
               End
               Begin VB.TextBox txt 
                  BackColor       =   &H8000000F&
                  Height          =   330
                  Index           =   2
                  Left            =   1350
                  MaxLength       =   50
                  TabIndex        =   3
                  Text            =   "txt(2)"
                  Top             =   1440
                  Width           =   4245
               End
               Begin VB.TextBox txt 
                  Height          =   330
                  Index           =   1
                  Left            =   1350
                  MaxLength       =   200
                  TabIndex        =   2
                  Text            =   "txt(1)"
                  Top             =   1065
                  Width           =   9510
               End
               Begin VB.Frame Frame5 
                  Caption         =   "Seleccionar"
                  Height          =   645
                  Left            =   165
                  TabIndex        =   35
                  Top             =   2235
                  Width           =   5760
                  Begin VB.OptionButton opt_planilla 
                     Caption         =   "No Considerar en Planilla"
                     Height          =   225
                     Index           =   1
                     Left            =   3180
                     TabIndex        =   36
                     Top             =   300
                     Width           =   2100
                  End
                  Begin VB.OptionButton opt_planilla 
                     Caption         =   "Considerar en Planilla"
                     Height          =   225
                     Index           =   0
                     Left            =   420
                     TabIndex        =   6
                     Top             =   300
                     Value           =   -1  'True
                     Width           =   1905
                  End
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   1
                  Left            =   1875
                  Picture         =   "FrmManConceptoDoc.frx":23EC
                  Style           =   1  'Graphical
                  TabIndex        =   31
                  ToolTipText     =   "Seleccione el Sexo"
                  Top             =   750
                  Width           =   210
               End
               Begin VB.CommandButton cb 
                  Height          =   225
                  Index           =   0
                  Left            =   1875
                  Picture         =   "FrmManConceptoDoc.frx":251E
                  Style           =   1  'Graphical
                  TabIndex        =   27
                  ToolTipText     =   "Seleccione el Tipo de Documento"
                  Top             =   420
                  Width           =   210
               End
               Begin VB.TextBox txt 
                  Height          =   870
                  Index           =   5
                  Left            =   165
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   9
                  Tag             =   "null"
                  Text            =   "FrmManConceptoDoc.frx":2650
                  Top             =   4920
                  Width           =   10800
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
                  Height          =   1575
                  Left            =   165
                  TabIndex        =   22
                  Top             =   2910
                  Width           =   10815
                  Begin VB.CheckBox chk_formula 
                     Caption         =   "Fórmula"
                     Enabled         =   0   'False
                     Height          =   225
                     Left            =   150
                     TabIndex        =   7
                     Top             =   45
                     Width           =   885
                  End
                  Begin VB.CommandButton cmd_formula 
                     Caption         =   "Editar Formula"
                     Enabled         =   0   'False
                     Height          =   645
                     Left            =   90
                     Picture         =   "FrmManConceptoDoc.frx":2659
                     Style           =   1  'Graphical
                     TabIndex        =   8
                     Top             =   735
                     Width           =   1275
                  End
                  Begin VB.TextBox txt_formula 
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
                     TabIndex        =   23
                     Text            =   "FrmManConceptoDoc.frx":275B
                     Top             =   270
                     Width           =   9045
                  End
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   0
                  Left            =   1350
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   0
                  Text            =   "txt_cb(0)"
                  Top             =   390
                  Width           =   765
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   1
                  Left            =   1350
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   1
                  Text            =   "txt_cb(1)"
                  ToolTipText     =   "Ingrese el Sexo (1:Masculino, 2:Femenino)"
                  Top             =   720
                  Width           =   765
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cod.Sunat"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   4
                  Left            =   5160
                  TabIndex        =   38
                  ToolTipText     =   "Es nombre abreviado del concepto, con posibilidad de que aparesca en el reporte"
                  Top             =   1965
                  Width           =   750
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nombre Corto"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   3
                  Left            =   165
                  TabIndex        =   37
                  ToolTipText     =   "Es nombre abreviado del concepto, con posibilidad de que aparesca en el reporte"
                  Top             =   1950
                  Width           =   975
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo"
                  Height          =   195
                  Index           =   1
                  Left            =   165
                  TabIndex        =   34
                  Top             =   810
                  Width           =   315
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
                  Left            =   3870
                  TabIndex        =   32
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   975
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
                  Left            =   3870
                  TabIndex        =   29
                  Top             =   390
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.Label lbl_capt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Categoría"
                  Height          =   195
                  Index           =   0
                  Left            =   165
                  TabIndex        =   28
                  Top             =   495
                  Width           =   705
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Comentario :"
                  Height          =   210
                  Index           =   5
                  Left            =   165
                  TabIndex        =   26
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
                  Left            =   165
                  TabIndex        =   25
                  Top             =   1575
                  Width           =   570
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nombre"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Index           =   1
                  Left            =   165
                  TabIndex        =   24
                  Top             =   1200
                  Width           =   555
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
                  Left            =   2145
                  TabIndex        =   33
                  Top             =   720
                  Width           =   4905
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
                  Left            =   2145
                  TabIndex        =   30
                  Top             =   390
                  Width           =   4395
               End
            End
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   10125
            TabIndex        =   20
            Top             =   75
            Visible         =   0   'False
            Width           =   495
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
            TabIndex        =   13
            Top             =   60
            Width           =   11400
         End
      End
   End
End
Attribute VB_Name = "FrmManConceptoDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstFrm As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim Agregando As Boolean

Dim fOrdenLista As Boolean ''--especifica el orden de la lista de la consulta

Sub Cancelar()
    ActivaTool
    Label5.Caption = "Detalle del Personal"
    QueHace = 3
    Bloquea False
    TabOne1.TabEnabled(0) = True
    Dim K&
    For K = 1 To TabOne2.NumTabs - 1
        TabOne2.TabEnabled(K) = True
    Next K
    TabOne1.CurrTab = 0
End Sub

Private Sub chk_formula_Click()
    If chk_formula.Value = 1 Then
        cmd_formula.Enabled = True
        txt_formula.Text = txt_formula.Tag
    Else
        cmd_formula.Enabled = False
        txt_formula.Text = ""
    End If
End Sub

Private Sub cmd_formula_Click()
    FrmManConceptoFormula.Show
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

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
   
    SeEjecuto = False
    pCargarGrid
    
    SeEjecuto = True
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount = 0 Then
        If MsgBox("No se ha registrado Canje de Documento, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
            nuevo
        End If
    End If
    
    
End Sub


Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
    
End Sub


Sub Blanquea()
    txt_formula.Text = ""
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cb
    LimpiaText lbl_cod

End Sub

Sub Bloquea(band As Boolean)

    habilitar_Locked txt, Not band
    habilitar_Locked txt_cb, Not band
    
    chk_formula.Enabled = band
    
    If (QueHace = 1) Or (QueHace = 2 And NulosC(txt(2).Text) = "") Then
        txt(2).Enabled = True
        txt(2).BackColor = vbWhite
    Else
        txt(2).Enabled = False
        txt(2).BackColor = &H8000000F
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

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Function Grabar() As Boolean

    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " al Concepto ", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId&
    On Error GoTo LaCague

    xCon.BeginTrans

    If QueHace = 1 Then

        xId = HallaCodigoTabla("pla_concepto", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_concepto", xCon

        RstCab.AddNew
        RstCab("id") = xId
    Else
        RST_Busq RstCab, "SELECT * FROM pla_concepto WHERE id = " & RstFrm("id") & "", xCon
        xId = RstCab("id")
    End If

    RstCab("idtipo") = NulosN(txt_cb(1).Text)
    RstCab("codsun") = NulosC(txt(4).Text)
    RstCab("descripcion") = NulosC(txt(1).Text)
    RstCab("variable") = NulosC(txt(2).Text)
    RstCab("formula") = NulosC(txt_formula.Text)
    RstCab("aplanilla") = IIf(opt_planilla(0).Value = True, -1, 0)
    RstCab("nomcorto") = NulosC(txt(3).Text)
    RstCab("observacion") = NulosC(txt(5).Text)
    
    RstCab.Update
    
    MsgBox "Los datos del Concepto" + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Grabar = True
    Exit Function
LaCague:
    Set RstCab = Nothing
    Set RstDet = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar al Personal por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Sub nuevo()
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea True
    Blanquea
    Label5.Caption = "Agregando Conceptos"
    TabOne2.CurrTab = 0
    txt_cb(0).SetFocus
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
    txt_cb(0).SetFocus

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
    Dim nSQL As String
    Rpta = MsgBox("Esta seguro de eliminar al Concepto seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Dim RstBus As New ADODB.Recordset
        xId = RstFrm.Fields("id")
        nSQL = "SELECT pla_concepto.id, pla_concepto.descripcion, pla_concepto.variable, pla_concepto.formula " _
            + vbCr + " FROM pla_concepto " _
            + vbCr + " WHERE (((pla_concepto.id)<>" & xId & ") AND ((pla_concepto.formula) Like '*" & RstFrm.Fields("formula") & "*'));"
            
        RST_Busq RstBus, nSQL, xCon
        If RstFrm.RecordCount <> 0 Then
            MsgBox "No se puede eliminar el Concepto, figura en fórmulas de otros conceptos" + vbCr + "Ej. Concepto: " & RstBus.Fields("descripcion"), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set RstBus = Nothing
            Exit Sub
        End If
        Set RstBus = Nothing
        
        '--falta validar que el concepto no este ya en planilla
        
        MsgBox "El Concepto se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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

    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then
        RstFrm.Filter = adFilterNone
    End If

    If Button.Index = 10 Then Buscar

    If Button.Index = 14 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Buscar()
    Dim xRs As New ADODB.Recordset
    Dim nSQL  As String
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "CodSunat":    xCampos(0, 1) = "codsun":      xCampos(0, 2) = "800":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripción": xCampos(1, 1) = "descripcion": xCampos(1, 2) = "3200":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Tipo":        xCampos(2, 1) = "tiponombre":  xCampos(2, 2) = "2200":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Categoría":   xCampos(3, 1) = "catnombre":   xCampos(3, 2) = "2000":    xCampos(3, 3) = "C"
    
    nSQL = "SELECT pla_conceptotipo.idcat, pla_conceptocat.descripcion AS catnombre, pla_conceptotipo.descripcion AS tiponombre, pla_concepto.*, pla_conceptotipo.codsun AS tiposun " _
        + vbCr + " FROM (pla_conceptocat RIGHT JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) RIGHT JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
        + vbCr + " ORDER BY pla_conceptocat.descripcion desc,pla_concepto.codsun asc, pla_conceptotipo.descripcion, pla_concepto.descripcion; "

    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Conceptos", "descripcion", "descripcion", Principio
    
    If xRs.State = 1 Then
        RstFrm.MoveFirst
        RstFrm.Find "id = " & xRs("id") & ""
    End If
    Set xRs = Nothing
    
End Sub

Sub Filtrar()
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 3) As String

    xCampos(0, 0) = "CodSunat":    xCampos(0, 1) = "codsun":      xCampos(0, 2) = "C":     xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripción": xCampos(1, 1) = "descripcion": xCampos(1, 2) = "c":    xCampos(1, 3) = "3200"
    xCampos(2, 0) = "Variable":    xCampos(2, 1) = "variable":    xCampos(2, 2) = "c":    xCampos(2, 3) = "3200"
    xCampos(3, 0) = "Tipo":        xCampos(3, 1) = "tiponombre":  xCampos(3, 2) = "c":    xCampos(3, 3) = "2200"
    xCampos(4, 0) = "Categoría":   xCampos(4, 1) = "catnombre":   xCampos(4, 2) = "c":    xCampos(4, 3) = "2000"

    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1

End Sub



'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--categoria de concepto
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            nTitulo = "Categoría de Concepto"
            nSQL = "SELECT pla_conceptocat.id, pla_conceptocat.descripcion AS nombre, pla_conceptocat.id AS cod " _
                + vbCr + " FROM pla_conceptocat;"
        
        Case 1 '--tipo de concepto
            If NulosN(txt_cb(0).Text) = 0 Then
                MsgBox "Seleccione la Categoría de Concepto", vbExclamation, xTitulo
                Exit Sub
            End If
        
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
            
            nTitulo = "Buscando Tipo Concepto"

            nSQL = "SELECT pla_conceptotipo.id, pla_conceptotipo.descripcion AS nombre, pla_conceptotipo.id AS cod " _
                + vbCr + " FROM pla_conceptotipo " _
                + vbCr + " WHERE (((pla_conceptotipo.idcat)=" & NulosN(txt_cb(0).Text) & "));"
                
    End Select

    Dim xRs As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO

    If Trim(lbl_cod(Index).Tag) <> Trim(lbl_cod(Index).Caption) Then
        Select Case Index
            Case 0 '--categoria
                txt_cb(1).Text = ""
        End Select
    End If
    Select Case Index
        Case 0 '--categoria
            txt(1).SetFocus
        Case 1 '--tipo
            txt(1).SetFocus
    End Select

salir:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
        If Index = 3 Then '--departamento
            txt_cb(4).Text = ""
            txt_cb(5).Text = ""
        ElseIf Index = 4 Then '--provincia
            txt_cb(5).Text = ""
        End If
    End If

End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
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
'    On Error GoTo error
    Select Case Index

        Case 0 '--categoria de concepto
            nSQL = "SELECT pla_conceptocat.id, pla_conceptocat.descripcion AS nombre, pla_conceptocat.id AS cod " _
                + vbCr + " FROM pla_conceptocat " _
                + vbCr + " WHERE pla_conceptocat.id = " & NulosN(txt_cb(Index).Text) & ";"
        
        Case 1 '--tipo de concepto
            If NulosN(txt_cb(0).Text) = 0 Then
                MsgBox "Seleccione la Categoría de Concepto", vbExclamation, xTitulo
                Exit Sub
            End If
            nSQL = "SELECT pla_conceptotipo.id, pla_conceptotipo.descripcion AS nombre, pla_conceptotipo.id AS cod " _
                + vbCr + " FROM pla_conceptotipo " _
                + vbCr + " WHERE (((pla_conceptotipo.idcat)=" & NulosN(txt_cb(0).Text) & ")) and pla_conceptotipo.id = " & NulosN(txt_cb(Index).Text) & ";"
                
    End Select

    If xCon.State = 0 Then GoTo salir
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    '--------------
    If Trim(lbl_cod(Index).Tag) <> Trim(lbl_cod(Index).Caption) Then
        Select Case Index
            Case 0 '--categoria de concepto
                txt_cb(1).Text = ""
        End Select
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
    QueHaceTmp = QueHace
    QueHace = -1 '--comodin para entrar a [txt_cb_Validate]
    txt(0).Text = RstFrm("id")
    '--datos de concepto
    If NulosN(RstFrm("idcat")) <> 0 Then
        txt_cb(0).Text = NulosN(RstFrm("idcat"))
        txt_cb_Validate 0, False
    End If
    If NulosN(RstFrm("idtipo")) <> 0 Then
        txt_cb(1).Text = NulosN(RstFrm("idtipo"))
        txt_cb_Validate 1, False
    End If
    
    txt(1).Text = NulosC(RstFrm("descripcion"))
    txt(2).Text = NulosC(RstFrm("variable"))
    txt(3).Text = NulosC(RstFrm("nomcorto"))
    txt(4).Text = NulosC(RstFrm("codsun"))
    
    If NulosC(RstFrm("formula")) <> "" Then
        chk_formula.Value = 1
        txt_formula.Text = RstFrm("formula")
    Else
        chk_formula.Value = 0
    End If
    txt(5).Text = NulosC(RstFrm("observacion"))
    txt_formula.Text = NulosC(RstFrm("formula"))
    chk_formula.Tag = txt_formula.Text
    '--datos de contables
    
    QueHace = QueHaceTmp

End Sub


Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL As String
    
    nSQL = "SELECT pla_conceptotipo.idcat, pla_conceptocat.descripcion AS catnombre, pla_conceptotipo.descripcion AS tiponombre, pla_concepto.*, pla_conceptotipo.codsun AS tiposun " _
        + vbCr + " FROM (pla_conceptocat RIGHT JOIN pla_conceptotipo ON pla_conceptocat.id = pla_conceptotipo.idcat) RIGHT JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo " _
        + vbCr + " ORDER BY pla_conceptocat.descripcion desc,pla_concepto.codsun asc, pla_conceptotipo.descripcion, pla_concepto.descripcion; "

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
    TabOne2.CurrTab = 0
    
    band = Validar(txt_cb)
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl_capt(band).Caption, vbInformation, xTitulo
       txt_cb(band).SetFocus
       Exit Function
    End If

    
    band = Validar(txt)
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl(band).Caption, vbInformation, xTitulo
       txt(band).SetFocus
       Exit Function
    End If
    
    If opt_planilla(0).Value = True And NulosC(txt(4).Text) = "" Then
        MsgBox "Ingrese el Campo Codigo Sunat", vbExclamation, xTitulo
        txt(4).SetFocus
        Exit Function
    End If
    
    If chk_formula.Value = 1 And NulosC(txt_formula.Text) = "" Then
        MsgBox "Ingrese la fórmula", vbExclamation, xTitulo
        cmd_formula.SetFocus
        Exit Function
    End If
    
    '--
    fValidarDatos = True
    
End Function

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If Index <> 4 Then Exit Sub
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
    
End Sub
